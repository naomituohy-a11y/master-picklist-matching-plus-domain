import os
import re
import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import gradio as gr
import gradio.themes as gthemes

# ============================================================
# 🧩 Setup – Normalization helpers and constants
# ============================================================

SUFFIXES = {
    "ltd", "limited", "co", "company", "corp", "corporation", "inc", "incorporated",
    "plc", "public", "llc", "lp", "llp", "ulc", "pc", "pllc", "sa", "ag", "nv",
    "se", "bv", "oy", "ab", "aps", "as", "kft", "zrt", "rt", "sarl", "sas", "spa",
    "gmbh", "ug", "bvba", "cvba", "nvsa", "pte", "pty", "bhd", "sdn", "kabushiki",
    "kaisha", "kk", "godō", "dmcc", "pjsc", "psc", "jsc", "ltda", "srl", "s.r.l",
    "group", "holdings", "limitedpartnership"
}

COUNTRY_EQUIVALENTS = {
    "uk": "united kingdom", "u.k.": "united kingdom", "england": "united kingdom",
    "great britain": "united kingdom", "britain": "united kingdom",
    "usa": "united states", "u.s.a.": "united states", "us": "united states",
    "america": "united states", "united states of america": "united states",
    "uae": "united arab emirates", "u.a.e.": "united arab emirates",
    "south korea": "republic of korea", "korea": "republic of korea",
    "north korea": "democratic people's republic of korea",
    "russia": "russian federation", "czechia": "czech republic",
    "côte d’ivoire": "ivory coast", "cote d'ivoire": "ivory coast",
    "iran": "islamic republic of iran", "venezuela": "bolivarian republic of venezuela",
    "taiwan": "republic of china", "hong kong sar": "hong kong", "macao sar": "macau", "prc": "china"
}

# ============================================================
# 🧹 Text cleaning helpers
# ============================================================

def _normalize_tokens(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    parts = [w for w in text.split() if w not in SUFFIXES]
    return " ".join(parts).strip()

def _clean_domain(domain: str) -> str:
    if not isinstance(domain, str):
        return ""
    domain = domain.lower()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    domain = re.sub(r"^www\.", "", domain)
    parts = domain.split(".")
    return parts[-2] if len(parts) >= 2 else domain

def _extract_domain_from_email(email: str) -> str:
    if not isinstance(email, str) or "@" not in email:
        return ""
    domain = email.split("@")[-1].lower().strip()
    domain = re.sub(r"^www\.", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    return domain

# ============================================================
# 🌐 Company ↔ Domain Comparison
# ============================================================

def compare_company_domain(company: str, domain: str):
    if not isinstance(company, str) or not isinstance(domain, str):
        return "Unsure – Please Check", 0, "missing input"

    c = _normalize_tokens(company)
    d_raw = domain.lower().strip()
    d = _clean_domain(d_raw)

    # Alias equivalence table (expandable)
    aliases = {
        "johnlewis": "john lewis group",
        "directlinegroup": "direct line group",
        "dlg": "direct line group",
        "matalan": "matalan",
        "ticketmaster": "ticketmaster",
        "deliveroo": "deliveroo",
        "motorway": "motorway",
        "monsoon": "monsoon accessorize",
        "uktv": "uktv",
        "mg": "mg motor",
        "thg": "the hut group",
        "ihg": "intercontinental hotels group",
        "imperialbrands": "imperial brands"
    }

    if d in aliases:
        d = aliases[d]

    if d.replace(" ", "") in c.replace(" ", "") or c.replace(" ", "") in d.replace(" ", ""):
        return "Likely Match", 100, "direct containment"

    if len(d) <= 3 and d in c:
        return "Likely Match", 95, f"short alias match ({d})"

    score = fuzz.token_set_ratio(c, d)
    if score >= 85:
        return "Likely Match", score, "strong fuzzy"
    elif score >= 70:
        return "Unsure – Please Check", score, "weak fuzzy"
    else:
        return "Likely NOT Match", score, "low similarity"

# ============================================================
# 🧮 Main Matching Function
# ============================================================

def run_matching(master_file, picklist_file, highlight_changes=True, progress=gr.Progress(track_tqdm=True)):
    try:
        progress(0, desc="📂 Reading uploaded files...")
        df_master = pd.read_excel(master_file.name)
        df_picklist = pd.read_excel(picklist_file.name)

        progress(0.2, desc="🔧 Preparing data...")
        df_out = df_master.copy()
        corrected_cells = set()

        # Country normalization & value matching
        for col in df_master.columns:
            if "country" in col.lower():
                df_out[col] = df_master[col].astype(str).apply(lambda x: COUNTRY_EQUIVALENTS.get(x.strip().lower(), x))

        # ---- Domain vs Company (from email only) ----
        progress(0.6, desc="🌐 Validating company ↔ email domain...")
        company_cols = [c for c in df_master.columns if c.strip().lower() in ["company", "companyname", "company name"]]
        email_cols = [c for c in df_master.columns if "email" in c.lower()]

        if company_cols and email_cols:
            company_col = company_cols[0]
            email_col = email_cols[0]

            statuses, scores, reasons = [], [], []
            for i in range(len(df_master)):
                comp = df_master.at[i, company_col]
                dom = _extract_domain_from_email(df_master.at[i, email_col]) if pd.notna(df_master.at[i, email_col]) else ""
                status, score, reason = compare_company_domain(comp, dom)
                statuses.append(status)
                scores.append(score)
                reasons.append(reason)

            df_out["Domain_Check_Status"] = statuses
            df_out["Domain_Check_Score"] = scores
            df_out["Domain_Check_Reason"] = reasons
        else:
            df_out["Domain_Check_Status"] = "No company/email columns found"
            df_out["Domain_Check_Score"] = None
            df_out["Domain_Check_Reason"] = None

        # ---- Save + Formatting ----
        progress(0.9, desc="💾 Saving results...")
        out_file = f"{os.path.splitext(master_file.name)[0]} - Full_Check_Results.xlsx"
        df_out.to_excel(out_file, index=False)

        wb = load_workbook(out_file)
        ws = wb.active
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        for col_idx, col in enumerate(df_out.columns, start=1):
            if "Domain_Check_Status" in col:
                for row in range(2, ws.max_row + 1):
                    val = str(ws.cell(row=row, column=col_idx).value).strip().lower()
                    if "likely match" in val:
                        ws.cell(row=row, column=col_idx).fill = green
                    elif "not match" in val:
                        ws.cell(row=row, column=col_idx).fill = red
                    else:
                        ws.cell(row=row, column=col_idx).fill = yellow

        wb.save(out_file)
        progress(1.0, desc="✅ Done! File ready for download.")
        return out_file

    except Exception as e:
        return f"❌ Error: {str(e)}"

# ============================================================
# 🎨 Fancy UI Theme (fixed for Gradio 4.44)
# ============================================================

fancy_theme = gthemes.Soft(
    primary_hue="blue",
    secondary_hue="indigo",
    neutral_hue="slate",
    text_size="md",
    radius_size="lg",
).set(
    body_background_fill="#f8fafc",
    block_background_fill="#ffffff",
    border_color_primary="#d1d5db",
    button_primary_background_fill="#2563eb",
    button_primary_text_color="white",
)

custom_css = """
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap');
body, .gradio-container {
    font-family: 'Poppins', sans-serif !important;
    background: linear-gradient(180deg, #f9fafb 0%, #eef2ff 100%) !important;
}
h1, h2, h3, .title {
    color: #1e293b !important;
    font-weight: 600 !important;
}
.gr-button {
    background: linear-gradient(90deg, #2563eb, #4f46e5) !important;
    color: white !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    transition: all 0.2s ease-in-out !important;
}
"""

# ============================================================
# 🎛️ Gradio Interface
# ============================================================

demo = gr.Interface(
    fn=run_matching,
    inputs=[
        gr.File(label="Upload MASTER Excel file (.xlsx)"),
        gr.File(label="Upload PICKLIST Excel file (.xlsx)"),
        gr.Checkbox(label="Highlight changed values (blue)", value=True)
    ],
    outputs=gr.File(label="Download Processed File"),
    title="📊 Master–Picklist + Domain Matching Tool",
    description="Upload MASTER & PICKLIST Excel files to auto-match, validate domains, and optionally highlight changed values.",
    theme=fancy_theme,
    css=custom_css
)

# ============================================================
# 🚀 Launch (Railway compatible)
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.queue(concurrency_count=1, max_size=10).launch(
        server_name="0.0.0.0",
        server_port=port,
        share=False,
        show_api=False,
        favicon_path=None,
        quiet=True
    )
