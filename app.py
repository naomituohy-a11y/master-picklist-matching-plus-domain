import os
import re
import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import gradio as gr

# ==============================================================
# Domain ↔ Company helpers (no AI, no web calls)
# ==============================================================

SUFFIXES = {
    # --- legal / corporate ---
    "ltd","limited","co","company","corp","corporation","inc","incorporated","plc","public",
    "llc","lp","llp","ulc","pc","pllc","sa","ag","nv","se","bv","oy","ab","aps","as","kft",
    "zrt","rt","sarl","sas","spa","gmbh","ug","kg","bvba","cvba","nv","pte","pty","bhd","sdn",
    "kabushiki","kaisha","kk","godo","gk","dmcc","pjsc","psc","jsc","ltda","srl","s.r.l.",
    # --- descriptive / brand (safe to ignore for matching) ---
    "group","holdings","holding","intl","international","global","ventures","solutions","systems",
    "technology","technologies","services","enterprise","enterprises","industries","industry",
    "consulting","partners","management","capital","investments","investment","resources","media",
    "marketing","communications","digital","creative","agency","engineering","finance","financial",
    "energy","power","renewables","medical","healthcare","hospital","labs","pharma","realestate",
    "developers","construction","travel","tourism","transport","shipping","mining","metals",
    "automotive","education","training","academy","university","institute"
}
STOPWORDS = {"net","pro","it","web","data","info","biz"}
THRESHOLD = 70  # lower bound for "Unsure – Please Check"

def _normalize_tokens(text: str) -> str:
    s = re.sub(r"[^\w\s]", "", str(text).lower())
    toks = [t for t in s.split() if t not in SUFFIXES]
    return " ".join(toks)

def _clean_company_series(s: pd.Series) -> pd.Series:
    return s.astype(str).apply(_normalize_tokens)

def _clean_domain_series(s: pd.Series) -> pd.Series:
    tmp = (
        s.astype(str)
        .str.lower()
        .str.replace(r"^https?://", "", regex=True)
        .str.replace(r"/.*$", "", regex=True)
        .str.replace(r"^www\.", "", regex=True)
    )
    # second-level label (before .com / .co.uk etc.)
    core = tmp.str.split(".").str[-2]
    # strip common brandy suffixes on the domain core
    core = core.str.replace(
        r"(andco|group|intl|international|global|solutions|systems|holdings|ltd|inc)$",
        "",
        regex=True
    )
    return core

def _make_acronym(name: str) -> str:
    toks = re.findall(r"\b\w", str(name))
    return "".join(toks).lower()

def score_domain_company(company_raw: str, domain_raw: str):
    """
    Returns: (is_match: bool, score: int, method: str, status: str)
    Status ∈ {"Likely Match","Unsure – Please Check","Likely Not Match"}
    """
    if not company_raw or not domain_raw:
        return False, 0, "none", "Likely Not Match"

    comp = _normalize_tokens(company_raw)
    dom  = re.sub(r"[^\w]", "", str(domain_raw).lower())

    if not comp or not dom:
        return False, 0, "none", "Likely Not Match"

    if len(dom) < 4 or dom in STOPWORDS:
        return False, 0, "generic", "Unsure – Please Check"

    # 1) containment on de-spaced company core
    if dom in comp.replace(" ", ""):
        return True, 100, "contains", "Likely Match"

    # 2) acronym match (>=3 chars)
    ac = _make_acronym(company_raw)
    if len(ac) >= 3 and (dom == ac or dom.startswith(ac)):
        return True, 95, "acronym", "Likely Match"

    # 3) fuzzy combo
    m1 = fuzz.partial_ratio(dom, comp)
    m2 = fuzz.token_set_ratio(dom, comp)
    fused = max(m1, m2)
    if fused >= 90:
        return True, fused, "fuzzy", "Likely Match"
    if fused >= THRESHOLD:
        return True, fused, "fuzzy", "Unsure – Please Check"
    return False, fused, "fuzzy", "Likely Not Match"


# ==============================================================
# Master ↔ Picklist matching (original logic, unchanged)
# ==============================================================

def run_matching(master_file, picklist_file, progress=gr.Progress(track_tqdm=True)):
    try:
        progress(0, desc="Reading uploaded files...")
        df_master = pd.read_excel(master_file.name)
        df_picklist_raw = pd.read_excel(picklist_file.name)

        progress(0.2, desc="Preparing data...")

        # EXACT pairs to check (same as original)
        EXACT_PAIRS = [
            ("c_industry", "c_industry"),
            ("asset_title", "asset_title"),
            ("lead_country", "lead_country"),
            ("departments", "departments"),
            ("c_state", "c_state"),
        ]

        keep_cols = ["c_industry", "asset_title", "lead_country", "departments", "c_state", "seniority"]
        df_picklist = df_picklist_raw[[c for c in keep_cols if c in df_picklist_raw.columns]].dropna(how="all").reset_index(drop=True)

        def normalize(s):
            return s.fillna("").astype(str).str.strip().str.lower()

        df_out = df_master.copy()
        corrected_cells = set()  # (col_name, excel_row_index) for blue highlights

        progress(0.4, desc="Running matching logic...")

        for master_col, picklist_col in EXACT_PAIRS:
            out_col = f"Match_{master_col}"
            if master_col in df_master.columns and picklist_col in df_picklist.columns:
                pick_map = {v.strip().lower(): v.strip() for v in df_picklist[picklist_col].dropna().astype(str)}
                matches, new_vals = [], []
                for i, val in enumerate(df_master[master_col].fillna("").astype(str)):
                    val_norm = val.strip().lower()
                    if val_norm in pick_map:
                        matches.append("Yes")
                        new_val = pick_map[val_norm]
                        new_vals.append(new_val)
                        if new_val != val:
                            corrected_cells.add((master_col, i + 2))  # +2 because excel rows start at 1 w/ header
                    else:
                        matches.append("No")
                        new_vals.append(val)
                df_out[out_col] = matches
                df_out[master_col] = new_vals
            else:
                df_out[out_col] = "Column Missing"

        # -------------------- Seniority parsing (original) --------------------
        def parse_seniority(title):
            if not isinstance(title, str):
                return "Entry", "default: no seniority term found"
            t = title.lower().strip()
            if re.search(r"\bchief\b|\bcio\b|\bcto\b|\bceo\b|\bcfo\b|\bciso\b|\bcpo\b|\bcso\b|\bcoo\b|\bchro\b|\bpresident\b", t):
                return "C Suite", "keyword: c-level"
            if re.search(r"\bvice president\b|\bvp\b", t):
                return "VP", "keyword: vp"
            if re.search(r"\bhead\b", t):
                return "Head", "keyword: head"
            if re.search(r"\bdirector\b", t):
                return "Director", "keyword: director"
            if re.search(r"\bmanager\b|\bmgr\b", t):
                return "Manager", "keyword: manager/mgr"
            if re.search(r"\bsenior\b|\bsr\b|\blead\b|\bprincipal\b", t):
                return "Senior", "keyword: senior/lead"
            if re.search(r"\bintern\b|\btrainee\b|\bassistant\b|\bgraduate\b", t):
                return "Entry", "keyword: entry-level term"
            if re.search(r"\bengineer\b|\barchitect\b|\banalyst\b|\bdeveloper\b|\bconsultant\b|\bscientist\b|\btechnician\b|\bdesigner\b|\bassociate\b|\bcoordinator\b", t):
                return "Entry", "default: technical role"
            return "Entry", "default: none found"

        if "jobtitle" in df_master.columns:
            parsed = df_master["jobtitle"].apply(parse_seniority)
            df_out["Parsed_Seniority"] = parsed.apply(lambda x: x[0])
            df_out["Seniority_Logic"]   = parsed.apply(lambda x: x[1])
        else:
            df_out["Parsed_Seniority"] = None
            df_out["Seniority_Logic"]  = "jobtitle column not found"

        # -------------------- Seniority match vs picklist --------------------
        if "seniority" in df_picklist.columns:
            sen_set = set(normalize(df_picklist["seniority"]))
            df_out["Seniority_Match"] = df_out["Parsed_Seniority"].apply(
                lambda x: "Yes" if isinstance(x, str) and x.strip().lower() in sen_set else "No"
            )
        else:
            df_out["Seniority_Match"] = "Picklist Missing"

        # ==============================================================
        # NEW: Domain ↔ Company validation appended to df_out
        # ==============================================================

        # Column auto-detection per your rules
        cols_lower = {c.lower(): c for c in df_master.columns}
        company_candidates = ["companyname", "company", "company name", "company_name"]
        domain_candidates  = ["website", "domain", "email domain", "email_domain"]

        company_col = None
        for key in company_candidates:
            if key in cols_lower:
                company_col = cols_lower[key]
                break

        domain_col = None
        for key in domain_candidates:
            if key in cols_lower:
                domain_col = cols_lower[key]
                break

        if company_col and domain_col:
            comp_series = df_master[company_col]
            dom_series  = df_master[domain_col]

            # Cleaned for reference (not strictly required to save)
            comp_clean = _clean_company_series(comp_series)
            dom_clean  = _clean_domain_series(dom_series)

            scores, methods, statuses, matches = [], [], [], []
            for comp_raw, dom_raw in zip(comp_series, dom_series):
                ok, sc, mth, stat = score_domain_company(comp_raw, dom_raw)
                matches.append("Yes" if ok else "No")
                scores.append(sc)
                methods.append(mth)
                statuses.append(stat)

            # Attach to output
            df_out["Domain_Match"]              = matches
            df_out["Domain_Score"]              = scores
            df_out["Domain_Method"]             = methods
            df_out["Domain_Connection_Status"]  = statuses
        else:
            # If columns missing, still append clear columns so the UI is predictable
            df_out["Domain_Match"]              = "Columns Missing"
            df_out["Domain_Score"]              = None
            df_out["Domain_Method"]             = "Columns Missing"
            df_out["Domain_Connection_Status"]  = "Columns Missing (companyname / website preferred)"

        # ==============================================================
        # Save while PRESERVING original formatting
        #   - Keep fonts/fills/widths for existing cells
        #   - Overwrite values where needed
        #   - Append NEW columns to the right with default styling
        # ==============================================================

        base_name = os.path.splitext(os.path.basename(master_file.name))[0]
        out_file  = f"{base_name} - Full_Check_Results.xlsx"

        # Load original workbook to preserve styles
        wb = load_workbook(master_file.name)
        ws = wb.active

        # Build header → column index map from row 1
        header_map = {}
        max_col = ws.max_column
        for col_idx in range(1, max_col + 1):
            header_val = ws.cell(row=1, column=col_idx).value
            if header_val is not None:
                header_map[str(header_val)] = col_idx

        # 1) Update existing columns' values (preserves existing styles)
        n_rows = len(df_out)
        for col_name in df_out.columns:
            if col_name in header_map:
                col_idx = header_map[col_name]
                col_values = df_out[col_name].tolist()
                for i, v in enumerate(col_values, start=2):  # data starts row 2
                    ws.cell(row=i, column=col_idx).value = v

        # 2) Append NEW columns (not in original sheet) to the right
        #    and color "Yes/No" + Domain status as required.
        yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        green  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        blue   = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
        amber  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # Unsure

        orig_cols_set = set(header_map.keys())
        new_cols = [c for c in df_out.columns if c not in orig_cols_set]

        # Write each new column at the end
        write_col_start = ws.max_column + 1
        for idx, col_name in enumerate(new_cols):
            col_idx = write_col_start + idx
            ws.cell(row=1, column=col_idx).value = col_name  # header
            # Fill header light yellow to indicate "added"
            ws.cell(row=1, column=col_idx).fill = yellow
            # Data
            for row_i, val in enumerate(df_out[col_name].tolist(), start=2):
                ws.cell(row=row_i, column=col_idx).value = val

                # Color coding for generic Yes/No match columns (green/red)
                if isinstance(val, str) and val.strip().lower() in {"yes", "no"}:
                    ws.cell(row=row_i, column=col_idx).fill = green if val.strip().lower() == "yes" else red

                # Specific coloring for Domain_Connection_Status
                if col_name == "Domain_Connection_Status" and isinstance(val, str):
                    v = val.lower()
                    if "likely match" in v:
                        ws.cell(row=row_i, column=col_idx).fill = green
                    elif "unsure" in v:
                        ws.cell(row=row_i, column=col_idx).fill = amber
                    elif "not match" in v:
                        ws.cell(row=row_i, column=col_idx).fill = red

        # 3) Blue highlight for cells corrected to picklist casing (existing columns)
        for col_name, row in corrected_cells:
            if col_name in header_map:
                ws.cell(row=row, column=header_map[col_name]).fill = blue

        wb.save(out_file)

        progress(1.0, desc="✅ Done! File ready for download.")
        return out_file

    except Exception as e:
        return f"❌ Error: {str(e)}"


# ==============================================================
# Gradio UI (same UX as original)
# ==============================================================

demo = gr.Interface(
    fn=run_matching,
    inputs=[
        gr.File(label="Upload MASTER Excel file (.xlsx)"),
        gr.File(label="Upload PICKLIST Excel file (.xlsx)")
    ],
    outputs=gr.File(label="Download Processed File"),
    title="📊 Master–Picklist + Domain Checker",
    description=(
        "Upload your MASTER and PICKLIST Excel files. "
        "The tool performs exact picklist matching, seniority parsing, and verifies "
        "whether each domain likely belongs to the listed company. "
        "Output preserves your original workbook formatting and adds color-coded result columns."
    )
)

if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=int(os.environ.get("PORT", 7860))
    )
