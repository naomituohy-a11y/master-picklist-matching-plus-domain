import os
import re
import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import gradio as gr

# ============================================================
# Domain ↔ Company matching helpers
# ============================================================

SUFFIXES = {
    # --- legal / corporate ---
    "ltd", "limited", "co", "company", "corp", "corporation", "inc", "incorporated",
    "plc", "public", "llc", "lp", "llp", "ulc", "pc", "pllc", "sa", "ag", "nv",
    "se", "bv", "oy", "ab", "aps", "as", "kft", "zrt", "rt", "sarl", "sas", "spa",
    "gmbh", "ug", "bvba", "cvba", "nvsa", "pte", "pty", "bhd", "sdn", "kabushiki",
    "kaisha", "kk", "godo", "dmcc", "pjsc", "psc", "jsc", "ltda", "srl",
    "s.r.l", "group", "holdings", "limitedpartnership"
}

STOPWORDS = {"net", "pro", "it", "web", "data", "info", "biz"}
THRESHOLD = 70  # fuzzy cutoff


def _normalize_tokens(text: str) -> str:
    """Clean and simplify text for comparison."""
    if not isinstance(text, str):
        return ""
    text = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    parts = [w for w in text.split() if w not in SUFFIXES]
    return " ".join(parts).strip()


def _clean_domain(domain: str) -> str:
    """Extract the core domain name."""
    if not isinstance(domain, str):
        return ""
    domain = domain.lower()
    domain = re.sub(r"^https?://", "", domain)
    domain = re.sub(r"/.*$", "", domain)
    domain = re.sub(r"^www\.", "", domain)
    parts = domain.split(".")
    if len(parts) >= 2:
        return parts[-2]
    return domain


def compare_company_domain(company: str, domain: str):
    """Return (match_status, score, reason)."""
    c = _normalize_tokens(company)
    d = _clean_domain(domain)
    if not c or not d:
        return "Unsure – Please Check", 0, "missing input"

    # direct containment
    if d in c.replace(" ", "") or c.replace(" ", "") in d:
        return "Likely Match", 100, "direct containment"

    # fuzzy ratio
    score = fuzz.token_set_ratio(c, d)
    if score >= 85:
        return "Likely Match", score, "strong fuzzy"
    elif score >= THRESHOLD:
        return "Unsure – Please Check", score, "weak fuzzy"
    else:
        return "Likely NOT Match", score, "low similarity"


# ============================================================
# Main matching logic (Master ↔ Picklist)
# ============================================================

def run_matching(master_file, picklist_file, progress=gr.Progress(track_tqdm=True)):
    try:
        progress(0, desc="📂 Reading uploaded files...")
        df_master = pd.read_excel(master_file.name)
        df_picklist = pd.read_excel(picklist_file.name)

        progress(0.2, desc="🔧 Preparing data...")

        # ---- Step 1: Define exact match columns ----
        EXACT_PAIRS = [
            ("c_industry", "c_industry"),
            ("asset_title", "asset_title"),
            ("lead_country", "lead_country"),
            ("departments", "departments"),
            ("c_state", "c_state"),
        ]

        df_out = df_master.copy()
        corrected_cells = set()

        # ---- Step 2: Normalize and compare columns ----
        progress(0.4, desc="🔍 Running master–picklist matching...")
        for master_col, picklist_col in EXACT_PAIRS:
            out_col = f"Match_{master_col}"
            if master_col in df_master.columns and picklist_col in df_picklist.columns:
                pick_map = {v.strip().lower(): v.strip()
                            for v in df_picklist[picklist_col].dropna().astype(str)}
                matches, new_vals = [], []
                for i, val in enumerate(df_master[master_col].fillna("").astype(str)):
                    val_norm = val.strip().lower()
                    if val_norm in pick_map:
                        matches.append("Yes")
                        new_val = pick_map[val_norm]
                        new_vals.append(new_val)
                        if new_val != val:
                            corrected_cells.add((master_col, i + 2))
                    else:
                        matches.append("No")
                        new_vals.append(val)
                df_out[out_col] = matches
                df_out[master_col] = new_vals
            else:
                df_out[out_col] = "Column Missing"

        # ---- Step 3: Seniority parsing ----
        def parse_seniority(title):
            if not isinstance(title, str):
                return "Entry", "no title"
            t = title.lower().strip()
            if re.search(r"\bchief\b|\bcio\b|\bcto\b|\bceo\b|\bcfo\b|\bciso\b|\bcpo\b|\bcso\b|\bcoo\b|\bchro\b|\bpresident\b", t):
                return "C Suite", "c-level"
            if re.search(r"\bvice president\b|\bvp\b", t):
                return "VP", "vp"
            if re.search(r"\bhead\b", t):
                return "Head", "head"
            if re.search(r"\bdirector\b", t):
                return "Director", "director"
            if re.search(r"\bmanager\b|\bmgr\b", t):
                return "Manager", "manager"
            if re.search(r"\bsenior\b|\bsr\b|\blead\b|\bprincipal\b", t):
                return "Senior", "senior"
            if re.search(r"\bintern\b|\btrainee\b|\bassistant\b|\bgraduate\b", t):
                return "Entry", "entry"
            return "Entry", "none"

        if "jobtitle" in df_master.columns:
            parsed = df_master["jobtitle"].apply(parse_seniority)
            df_out["Parsed_Seniority"] = parsed.apply(lambda x: x[0])
            df_out["Seniority_Logic"] = parsed.apply(lambda x: x[1])
        else:
            df_out["Parsed_Seniority"] = None
            df_out["Seniority_Logic"] = "jobtitle column not found"

        # ---- Step 4: Domain vs Company check ----
        progress(0.6, desc="🌐 Checking company ↔ domain connections...")

        company_cols = [c for c in df_master.columns if c.strip().lower() in
                        ["companyname", "company", "company name", "company_name"]]
        domain_cols = [c for c in df_master.columns if c.strip().lower() in
                       ["website", "domain", "email domain", "email_domain"]]

        if company_cols and domain_cols:
            company_col = company_cols[0]
            # prefer "website" if available
            if "website" in [c.lower() for c in domain_cols]:
                domain_col = [c for c in domain_cols if c.lower() == "website"][0]
            else:
                domain_col = domain_cols[0]

            statuses, scores, reasons = [], [], []
            for i in range(len(df_master)):
                comp = df_master.at[i, company_col]
                dom = df_master.at[i, domain_col]
                status, score, reason = compare_company_domain(comp, dom)
                statuses.append(status)
                scores.append(score)
                reasons.append(reason)

            df_out["Domain_Check_Status"] = statuses
            df_out["Domain_Check_Score"] = scores
            df_out["Domain_Check_Reason"] = reasons
        else:
            df_out["Domain_Check_Status"] = "No company/domain columns found"
            df_out["Domain_Check_Score"] = None
            df_out["Domain_Check_Reason"] = None

        # ---- Step 5: Save Excel output ----
        progress(0.9, desc="💾 Writing Excel output...")

        out_file = f"{os.path.splitext(master_file.name)[0]} - Full_Check_Results.xlsx"
        df_out.to_excel(out_file, index=False)

        # ---- Step 6: Reapply formatting ----
        wb = load_workbook(out_file)
        ws = wb.active
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        for col_idx, col in enumerate(df_out.columns, start=1):
            if col.startswith("Match_"):
                for row in range(2, ws.max_row + 1):
                    val = str(ws.cell(row=row, column=col_idx).value).strip().lower()
                    if val == "yes":
                        ws.cell(row=row, column=col_idx).fill = green
                    elif val == "no":
                        ws.cell(row=row, column=col_idx).fill = red
                    else:
                        ws.cell(row=row, column=col_idx).fill = yellow

        wb.save(out_file)
        progress(1.0, desc="✅ Done! File ready for download.")
        return out_file

    except Exception as e:
        return f"❌ Error: {str(e)}"


# ============================================================
# Gradio Blocks Interface (v5 compatible)
# ============================================================

with gr.Blocks(title="📊 Master–Picklist + Domain Matching Tool") as demo:
    gr.Markdown("## 📊 Master–Picklist + Domain Matching Tool")
    gr.Markdown(
        "Upload your MASTER and PICKLIST Excel files below to perform automated matching, seniority parsing, and domain/company validation."
    )

    master_file = gr.File(label="Upload MASTER Excel file (.xlsx)")
    picklist_file = gr.File(label="Upload PICKLIST Excel file (.xlsx)")
    run_btn = gr.Button("🚀 Run Matching Process")
    output_file = gr.File(label="Download Processed File")

    run_btn.click(fn=run_matching, inputs=[master_file, picklist_file], outputs=output_file)


# ============================================================
# Launch for local + Railway deployment
# ============================================================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
