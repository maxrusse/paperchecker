# End-to-end pipeline: PDF -> Driver(JSON) -> Verifier(JSON, all critical decisions) -> Validator -> Excel + Word (human review log)
#
# Dependencies (pip):
#   pip install -U openai google-genai pymupdf python-docx openpyxl jsonschema
#
# Notes:
# - All critical decisions are forced through a verifier review with evidence + explanation.
# - Output always writes Excel + Word as long as at least one verifier pass exists.
# - Any DISAGREE/UNSURE or missing review becomes a CRITICAL validation issue and is clearly logged in Word.

import os, json, re, copy
from datetime import datetime, UTC
import fitz  # PyMuPDF
import openpyxl
from docx import Document

from openai import OpenAI
from google import genai
from google.genai import types


# -------------------------
# CONFIG
# -------------------------
PDF_PATHS = [
    # Example:
    # "/mnt/data/paper1.pdf",
]

TEMPLATE_XLSX = "/mnt/data/Prevention of MRONJ_Extraction Sheet (Oli).xlsx"
OUT_XLSX = f"/mnt/data/mronj_prevention_filled_{datetime.now(UTC).strftime('%Y%m%d_%H%M%S')}.xlsx"
OUT_DOCX = f"/mnt/data/mronj_prevention_human_review_log_{datetime.now(UTC).strftime('%Y%m%d_%H%M%S')}.docx"

OPENAI_DRIVER_MODEL = "gpt-5.2"
GEMINI_DRIVER_MODEL = "gemini-3-pro-preview"
OPENAI_VERIFIER_MODEL = "gpt-5.2"
GEMINI_VERIFIER_MODEL = "gemini-3-pro-preview"
SUPERVISOR_MODEL = "gpt-5.2"

REASONING_EFFORT_OPENAI = "medium"   # none|low|medium|high|xhigh
THINKING_LEVEL_GEMINI = "low"        # minimal|low|high

MAX_VIEW_CHARS = 60000
VERIFIER_CHUNK_SIZE = 24
NUMERIC_TOL_ABS = 0.01
NUMERIC_TOL_REL = 0.01


# -------------------------
# EXCEL MAP (template-specific)
# -------------------------
EXCEL_MAP = {
    "sheet_key_to_name": {
        "included_articles": "Included Articles",
        "level_of_evidence": "Level of Evidence",
        "rct_appraisal": "Critical Appraisal of RCTS",
        "cohort_appraisal": "Critical Appraisal of Cohort",
        "case_series_appraisal": "Critical Appraisal of Case Seri",
        "case_control_appraisal": "Critical Appraisal of Case Cont",
        "systematic_appraisal": "Critical Appraisal of Systemati",
    },
    "sheets": {
        "included_articles": {
            "header_rows": 3,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "n_pts": "E",
                "age_mean_years": "F",
                "gender_male_n": "G",
                "gender_female_n": "H",
                "site_maxilla": "I",
                "site_mandible": "J",
                "site_both": "K",
                "primary_cause_breast_cancer": "L",
                "primary_cause_prostate_cancer": "M",
                "primary_cause_mm": "N",
                "primary_cause_osteoporosis": "O",
                "primary_cause_other": "P",
                "primary_cause_other_details": "Q",
                "ards_bisphosphonates_alendronate": "R",
                "ards_bisphosphonates_zoledronate": "S",
                "ards_bisphosphonates_risedronate": "T",
                "ards_bisphosphonates_neridronate": "U",
                "ards_bisphosphonates_pamidronate": "V",
                "ards_bisphosphonates_others": "W",
                "ards_bisphosphonates_others_details": "X",
                "ards_denosumab": "Z",
                "ards_both": "AA",
                "ards_other_drug": "Y",
                "ards_other_drug_details": "AD",
                "route_iv": "AB",
                "route_oral": "AC",
                "route_im": "AE",
                "route_subcutaneous": "AF",
                "route_both": "AG",
                "route_not_reported": "AA",  # kept as mapped in template (if present in your file; adjust if needed)
                "mronj_stage_at_risk": "AH",
                "mronj_stage_0": "AI",
                "prevention_technique": "AJ",
                "group_intervention": "AK",
                "group_control": "AL",
                "follow_up_mean_months": "AM",
                "follow_up_range": "AN",
                "outcome_variable": "AO",
                "mronj_development": "AP",
                "mronj_development_details": "AQ",
            },
        },
        "level_of_evidence": {
            "header_rows": 1,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "level_of_evidence": "E",
                "grade_of_recommendation": "F",
            },
        },
        "rct_appraisal": {
            "header_rows": 3,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "q1_randomized": "E",
                "q2_randomization_method": "F",
                "q3_double_blind": "G",
                "q4_blinding_method": "H",
                "q5_withdrawals": "I",
                "total_score": "J",
            },
        },
        "cohort_appraisal": {
            "header_rows": 2,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "q1_clear_question": "E",
                "q2_cohort_recruited": "F",
                "q3_exposure_measured": "G",
                "q4_outcome_measured": "H",
                "q5_confounders": "I",
                "q6_followup_complete": "J",
                "total_score": "K",
            },
        },
        "case_series_appraisal": {
            "header_rows": 2,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "q1_clear_aim": "E",
                "q2_inclusion_criteria": "F",
                "q3_consecutive_cases": "G",
                "q4_outcomes_defined": "H",
                "q5_followup_sufficient": "I",
                "q6_statistical_analysis": "J",
                "total_score": "K",
            },
        },
        "case_control_appraisal": {
            "header_rows": 2,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "q1_clear_question": "E",
                "q2_cases_representative": "F",
                "q3_controls_selected": "G",
                "q4_exposure_measured": "H",
                "q5_confounders": "I",
                "q6_results_precise": "J",
                "total_score": "K",
            },
        },
        "systematic_appraisal": {
            "header_rows": 3,
            "key": {"field": "pmid", "col": "A"},
            "columns": {
                "pmid": "A",
                "author": "B",
                "year": "C",
                "study_design": "D",
                "q1_focus_question": "E",
                "q2_inclusion_criteria": "F",
                "q3_comprehensive_search": "G",
                "q4_6_search_and_duplication": "H",
                "q7_quality_assessed": "I",
                "q8_combining_appropriate": "J",
                "q9_conclusions_supported": "K",
                "total_score": "L",
            },
        },
    },
}


# -------------------------
# JSON SCHEMAS (strict top-level)
# -------------------------
def build_sheet_schema(columns):
    field_types = ["string", "number", "integer", "boolean", "null"]
    return {
        "type": "object",
        "additionalProperties": False,
        "required": list(columns),
        "properties": {key: {"type": field_types} for key in columns},
    }

SHEET_SCHEMAS = {
    sheet_key: build_sheet_schema((cfg.get("columns") or {}).keys())
    for sheet_key, cfg in (EXCEL_MAP.get("sheets") or {}).items()
}
SCALAR_TYPES = ["string", "number", "integer", "boolean", "null"]
DRIVER_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["paper_id", "study_type", "record", "critical_decisions", "confidence", "notes"],
    "properties": {
        "paper_id": {
            "type": "object",
            "additionalProperties": False,
            "required": ["pmid", "doi", "title"],
            "properties": {
                "pmid": {"type": ["integer", "null"]},
                "doi": {"type": ["string", "null"]},
                "title": {"type": ["string", "null"]},
            },
        },
        "study_type": {
            "type": "string",
            "enum": ["rct", "cohort", "case_series", "case_control", "systematic_review", "other", "unclear"],
        },
        "record": {
            "type": "object",
            "additionalProperties": False,
            "required": ["sheets"],
            "properties": {
                "sheets": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": [
                        "included_articles",
                        "level_of_evidence",
                        "rct_appraisal",
                        "cohort_appraisal",
                        "case_series_appraisal",
                        "case_control_appraisal",
                        "systematic_appraisal",
                    ],
                    "properties": {
                        "included_articles": {"anyOf": [SHEET_SCHEMAS["included_articles"], {"type": "null"}]},
                        "level_of_evidence": {"anyOf": [SHEET_SCHEMAS["level_of_evidence"], {"type": "null"}]},
                        "rct_appraisal": {"anyOf": [SHEET_SCHEMAS["rct_appraisal"], {"type": "null"}]},
                        "cohort_appraisal": {"anyOf": [SHEET_SCHEMAS["cohort_appraisal"], {"type": "null"}]},
                        "case_series_appraisal": {"anyOf": [SHEET_SCHEMAS["case_series_appraisal"], {"type": "null"}]},
                        "case_control_appraisal": {"anyOf": [SHEET_SCHEMAS["case_control_appraisal"], {"type": "null"}]},
                        "systematic_appraisal": {"anyOf": [SHEET_SCHEMAS["systematic_appraisal"], {"type": "null"}]},
                    },
                }
            },
        },
        "critical_decisions": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["path", "value", "evidence", "is_critical"],
                "properties": {
                    "path": {"type": "string"},
                    "value": {"type": SCALAR_TYPES},
                    "evidence": {"type": "string"},
                    "is_critical": {"type": "boolean"},
                },
            },
        },
        "confidence": {"type": "number", "minimum": 0.0, "maximum": 1.0},
        "notes": {"type": "string"},
    },
}

VERIFIER_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["verdict", "critical_errors", "decision_reviews", "suggested_patch", "rationale", "confidence"],
    "properties": {
        "verdict": {"type": "string", "enum": ["AGREE", "DISAGREE", "UNSURE"]},
        "critical_errors": {"type": "array", "items": {"type": "string"}},
        "decision_reviews": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["path", "is_critical", "status", "driver_value", "proposed_value", "explanation", "evidence"],
                "properties": {
                    "path": {"type": "string"},
                    "is_critical": {"type": "boolean"},
                    "status": {"type": "string", "enum": ["AGREE", "DISAGREE", "UNSURE"]},
                    "driver_value": {"type": SCALAR_TYPES},
                    "proposed_value": {"type": SCALAR_TYPES},
                    "explanation": {"type": "string"},
                    "evidence": {"type": "string"},
                },
            },
        },
        "suggested_patch": {"type": ["object", "null"]},
        "rationale": {"type": "string"},
        "confidence": {"type": "number", "minimum": 0.0, "maximum": 1.0},
    },
}

SUPERVISOR_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["status", "value", "evidence", "confidence"],
    "properties": {
        "status": {"type": "string", "enum": ["PICK_A", "PICK_B", "UNSURE"]},
        "value": {"type": SCALAR_TYPES},
        "evidence": {"type": "string"},
        "confidence": {"type": "number", "minimum": 0.0, "maximum": 1.0},
    },
}


# -------------------------
# PIPELINE CORE (validator + excel + word)
# -------------------------
def json_pointer_get(obj, pointer):
    if pointer == "" or pointer == "/":
        return obj
    parts = pointer.lstrip("/").split("/")
    cur = obj
    for p in parts:
        p = p.replace("~1", "/").replace("~0", "~")
        if isinstance(cur, list):
            cur = cur[int(p)]
        else:
            cur = cur.get(p)
    return cur

def json_pointer_set(obj, pointer, value):
    parts = pointer.lstrip("/").split("/")
    cur = obj
    for i, p in enumerate(parts):
        p = p.replace("~1", "/").replace("~0", "~")
        last = (i == len(parts) - 1)
        if last:
            if isinstance(cur, list):
                cur[int(p)] = value
            else:
                cur[p] = value
        else:
            if isinstance(cur, list):
                cur = cur[int(p)]
            else:
                if p not in cur or not isinstance(cur[p], (dict, list)):
                    cur[p] = {}
                cur = cur[p]

def deep_merge(a, b):
    if not isinstance(a, dict) or not isinstance(b, dict):
        return copy.deepcopy(b)
    out = copy.deepcopy(a)
    for k, v in b.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out

def sanitize_for_model_input(obj):
    if not isinstance(obj, dict):
        return obj
    scrub = copy.deepcopy(obj)
    for k in list(scrub.keys()):
        if k in ("verification", "validation", "verifier_model", "model_meta", "model"):
            scrub.pop(k, None)
    return scrub

def _normalize_string(v):
    if not isinstance(v, str):
        return v
    s = v.strip()
    if s == "":
        return None
    sl = s.lower()
    canonical = {
        "yes": "Yes",
        "no": "No",
        "unclear": "Unclear",
        "nr": "NR",
        "n/a": "NR",
        "na": "NR",
        "not reported": "NR",
    }
    if sl in canonical:
        return canonical[sl]
    return s

def _coerce_number(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except ValueError:
            return None
    return None

def _values_match(a, b, abs_tol=NUMERIC_TOL_ABS, rel_tol=NUMERIC_TOL_REL):
    if a is None and b is None:
        return True
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    na = _normalize_string(a)
    nb = _normalize_string(b)
    if na == nb:
        return True
    ca = _coerce_number(a)
    cb = _coerce_number(b)
    if ca is not None and cb is not None:
        if abs(ca - cb) <= abs_tol:
            return True
        if abs(cb) > 0 and abs(ca - cb) / abs(cb) <= rel_tol:
            return True
    return False

def extract_report_payload(final_obj):
    return {
        "paper_id": final_obj.get("paper_id") or {"pmid": None, "doi": None, "title": None},
        "study_type": final_obj.get("study_type", "unclear"),
        "record": final_obj.get("record") or {"sheets": {}},
    }

def report_leaf_paths(payload):
    base = {"study_type": payload.get("study_type"), "record": payload.get("record")}
    paths = infer_all_leaf_paths(base)
    for key in ("pmid", "doi", "title"):
        paths.append(f"/paper_id/{key}")
    out = []
    seen = set()
    for p in paths:
        if p not in seen:
            out.append(p)
            seen.add(p)
    return out

def compare_reports_with_tolerance(payload_a, payload_b, abs_tol=NUMERIC_TOL_ABS, rel_tol=NUMERIC_TOL_REL):
    mismatches = []
    paths = report_leaf_paths(payload_a) + report_leaf_paths(payload_b)
    paths = list(dict.fromkeys(paths))
    for path in paths:
        va = json_pointer_get(payload_a, path)
        vb = json_pointer_get(payload_b, path)
        if _values_match(va, vb, abs_tol=abs_tol, rel_tol=rel_tol):
            continue
        mismatches.append({"path": path, "a": va, "b": vb})
    return mismatches

def _normalize_excel_value(v):
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, str):
        s = v.strip()
        sl = s.lower()
        if sl in ("true", "yes", "y", "1"):
            return 1
        if sl in ("false", "no", "n", "0"):
            return 0
        return s
    return v

def column_index_from_string(col):
    col = col.upper().strip()
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c) - ord("A") + 1)
    return idx

def _find_or_create_row(ws, key_col_letter, key_value, header_rows):
    key_col_idx = column_index_from_string(key_col_letter)
    start_row = header_rows + 1
    max_row = max(ws.max_row, start_row)

    if key_value not in (None, ""):
        for r in range(start_row, max_row + 1):
            if ws.cell(r, key_col_idx).value == key_value:
                return r

    for r in range(start_row, max_row + 1):
        if ws.cell(r, key_col_idx).value in (None, ""):
            return r

    return max_row + 1

def apply_to_workbook(final_obj, template_xlsx, out_xlsx, excel_map):
    wb = openpyxl.load_workbook(template_xlsx)
    sheets_data = ((final_obj.get("record") or {}).get("sheets")) or {}
    pmid = (final_obj.get("paper_id") or {}).get("pmid")

    for sheet_key, payload in sheets_data.items():
        if not isinstance(payload, dict):
            continue
        sheet_name = (excel_map.get("sheet_key_to_name") or {}).get(sheet_key)
        if not sheet_name or sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]
        sheet_cfg = (excel_map.get("sheets") or {}).get(sheet_key) or {}
        header_rows = int(sheet_cfg.get("header_rows") or 1)
        key_cfg = sheet_cfg.get("key") or {"field": "pmid", "col": "A"}
        key_col = key_cfg.get("col") or "A"
        row_idx = _find_or_create_row(ws, key_col, pmid, header_rows)

        cols = sheet_cfg.get("columns") or {}
        for field, col_letter in cols.items():
            if field == "pmid":
                ws[f"{col_letter}{row_idx}"].value = pmid
                continue
            if field in payload:
                ws[f"{col_letter}{row_idx}"].value = _normalize_excel_value(payload.get(field))

        inc = sheets_data.get("included_articles") or {}
        if isinstance(inc, dict):
            for f in ("author", "year", "study_design"):
                if f in cols and ws[f"{cols[f]}{row_idx}"].value in (None, ""):
                    if f in inc and inc.get(f) not in (None, ""):
                        ws[f"{cols[f]}{row_idx}"].value = _normalize_excel_value(inc.get(f))

    wb.save(out_xlsx)

def compute_scores_inplace(driver_out):
    sheets = (driver_out.get("record") or {}).get("sheets") or {}

    rct = sheets.get("rct_appraisal")
    if isinstance(rct, dict):
        score = 0
        for k in ("q1_randomized", "q2_randomization_method", "q3_double_blind", "q4_blinding_method", "q5_withdrawals"):
            v = rct.get(k)
            if v in (1, True, "1", "true", "True", "YES", "Yes"):
                score += 1
        rct["total_score"] = score

    for key in ("cohort_appraisal", "case_series_appraisal", "case_control_appraisal", "systematic_appraisal"):
        sd = sheets.get(key)
        if isinstance(sd, dict):
            score = 0
            for k, v in sd.items():
                if str(k).startswith("q") and v in (1, True, "1", "true", "True", "YES", "Yes"):
                    score += 1
            sd["total_score"] = score

def infer_all_leaf_paths(driver_out):
    paths = ["/study_type"]

    def walk(base, obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                p = base + "/" + str(k).replace("~", "~0").replace("/", "~1")
                if isinstance(v, dict):
                    walk(p, v)
                elif isinstance(v, list):
                    for i, it in enumerate(v):
                        walk(p + f"/{i}", it)
                else:
                    paths.append(p)
        elif isinstance(obj, list):
            for i, it in enumerate(obj):
                walk(base + f"/{i}", it)

    record = driver_out.get("record") or {}
    sheets = (record.get("sheets") or {})
    for sheet_key, payload in sheets.items():
        if isinstance(payload, dict):
            walk(f"/record/sheets/{sheet_key}", payload)

    out = []
    seen = set()
    for p in paths:
        if p not in seen:
            out.append(p)
            seen.add(p)
    return out

def rule_validation(merged_driver):
    issues = []
    sheets = (merged_driver.get("record") or {}).get("sheets") or {}
    inc = sheets.get("included_articles") or {}

    def _count_true(keys):
        c = 0
        for k in keys:
            v = inc.get(k)
            if v in (True, 1, "1", "true", "True", "YES", "Yes"):
                c += 1
        return c

    site_keys = ["site_maxilla", "site_mandible", "site_both"]
    if _count_true(site_keys) == 0:
        issues.append({"severity": "WARN", "code": "SITE_EMPTY", "message": "No site marked (maxilla/mandible/both).", "path": "/record/sheets/included_articles"})
    if _count_true(site_keys) > 1 and not (inc.get("site_both") in (True, 1, "1", "true", "True")):
        issues.append({"severity": "WARN", "code": "SITE_INCONSISTENT", "message": "Multiple site flags set but site_both not set.", "path": "/record/sheets/included_articles"})

    route_keys = ["route_iv", "route_oral", "route_im", "route_subcutaneous", "route_both", "route_not_reported"]
    if _count_true(route_keys) == 0:
        issues.append({"severity": "WARN", "code": "ROUTE_EMPTY", "message": "No route marked.", "path": "/record/sheets/included_articles"})
    if inc.get("route_both") in (True, 1, "1", "true", "True") and _count_true(["route_iv", "route_oral", "route_im", "route_subcutaneous"]) == 0:
        issues.append({"severity": "WARN", "code": "ROUTE_BOTH_NO_DETAILS", "message": "route_both is set but no specific route marked.", "path": "/record/sheets/included_articles"})
    if inc.get("route_not_reported") in (True, 1, "1", "true", "True") and _count_true(["route_iv", "route_oral", "route_im", "route_subcutaneous", "route_both"]) > 0:
        issues.append({"severity": "WARN", "code": "ROUTE_NR_CONFLICT", "message": "route_not_reported is set but other route flags are also set.", "path": "/record/sheets/included_articles"})

    mdev = inc.get("mronj_development")
    if isinstance(mdev, str) and mdev.strip().lower() not in ("yes", "no", "unclear", "n/a", "na", "nr", "not reported"):
        issues.append({"severity": "WARN", "code": "MRONJ_DEV_UNEXPECTED", "message": "mronj_development is not a standard token (Yes/No/Unclear).", "path": "/record/sheets/included_articles/mronj_development"})

    return issues

def compile_critical_decision_report(passes, critical_paths, final_driver):
    issues = []
    latest_review_by_path = {}
    for p in passes:
        for dr in (p.get("decision_reviews") or []):
            path = dr.get("path")
            if path:
                latest_review_by_path[path] = dr

    critical_report = []
    for path in critical_paths:
        review = latest_review_by_path.get(path)
        if review is None:
            critical_report.append({
                "path": path,
                "final_value": json_pointer_get(final_driver, path),
                "status": "MISSING",
                "explanation": "Missing verifier review for critical decision.",
                "evidence": "",
            })
            issues.append({
                "severity": "CRITICAL",
                "code": "MISSING_VERIFIER_REVIEW",
                "message": f"Critical decision not reviewed by verifier: {path}",
                "path": path,
            })
            continue

        status = review.get("status", "UNSURE")
        final_val = review.get("proposed_value", review.get("driver_value"))

        critical_report.append({
            "path": path,
            "final_value": final_val,
            "status": status,
            "explanation": review.get("explanation", ""),
            "evidence": review.get("evidence", ""),
        })

        if status in ("DISAGREE", "UNSURE"):
            issues.append({
                "severity": "CRITICAL",
                "code": f"VERIFIER_{status}",
                "message": f"Verifier status {status} for critical decision: {path}",
                "path": path,
            })

    return critical_report, issues

def write_human_review_docx(final_obj, docx_path, append=True):
    if append and os.path.exists(docx_path):
        doc = Document(docx_path)
        doc.add_page_break()
    else:
        doc = Document()
        doc.add_heading("MRONJ prevention extraction - human review log", level=0)

    paper_id = final_obj.get("paper_id") or {}
    pmid = paper_id.get("pmid")
    doi = paper_id.get("doi")
    title = paper_id.get("title")

    doc.add_heading(f"PMID: {pmid if pmid is not None else 'null'}", level=1)
    if title:
        doc.add_paragraph("Title: " + str(title))
    if doi:
        doc.add_paragraph("DOI: " + str(doi))

    doc.add_paragraph("Study type: " + str(final_obj.get("study_type")))

    needs = ((final_obj.get("validation") or {}).get("needs_human_review"))
    doc.add_paragraph("Needs human review: " + ("YES" if needs else "NO"))

    doc.add_heading("Critical decisions (verifier)", level=2)
    t = doc.add_table(rows=1, cols=4)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    hdr[0].text = "path"
    hdr[1].text = "status"
    hdr[2].text = "explanation"
    hdr[3].text = "evidence"

    for cd in (final_obj.get("verification") or {}).get("critical_decisions") or []:
        r = t.add_row().cells
        r[0].text = str(cd.get("path", ""))
        r[1].text = str(cd.get("status", ""))
        r[2].text = str(cd.get("explanation", ""))
        r[3].text = str(cd.get("evidence", ""))

    doc.add_heading("Validation issues", level=2)
    issues = (final_obj.get("validation") or {}).get("issues") or []
    if not issues:
        doc.add_paragraph("None.")
    else:
        for it in issues:
            doc.add_paragraph(f"[{it.get('severity')}] {it.get('code')}: {it.get('message')} (path={it.get('path')})")

    doc.add_heading("Verifier passes summary", level=2)
    passes = (final_obj.get("verification") or {}).get("passes") or []
    for i, p in enumerate(passes, 1):
        doc.add_paragraph(f"pass {i}: verdict={p.get('verdict')} confidence={p.get('confidence')} errors={'; '.join(p.get('critical_errors') or [])}")

    doc.save(docx_path)

def build_final_object(driver_out, verifier_passes, verifier_model=None, version="2.1"):
    merged = copy.deepcopy(driver_out)
    for p in verifier_passes:
        patch = p.get("suggested_patch")
        if isinstance(patch, dict) and patch:
            merged = deep_merge(merged, patch)

    compute_scores_inplace(merged)

    critical_paths = infer_all_leaf_paths(merged)

    critical_report, issues = compile_critical_decision_report(verifier_passes, critical_paths, merged)
    issues.extend(rule_validation(merged))

    final_obj = {
        "version": version,
        "paper_id": merged.get("paper_id") or {"pmid": None, "doi": None, "title": None},
        "study_type": merged.get("study_type", "unclear"),
        "record": merged.get("record") or {"sheets": {}},
        "verification": {
            "verifier_model": verifier_model,
            "passes": verifier_passes,
            "critical_decisions": critical_report,
        },
        "validation": {
            "needs_human_review": any(i.get("severity") == "CRITICAL" for i in issues),
            "issues": issues,
        },
    }
    return final_obj

def apply_final_to_outputs(final_obj, template_xlsx, out_xlsx, excel_map, review_docx_path):
    passes = ((final_obj.get("verification") or {}).get("passes")) or []
    if not passes:
        raise RuntimeError("No verifier passes provided. Refusing to write outputs.")
    apply_to_workbook(final_obj, template_xlsx, out_xlsx, excel_map)
    write_human_review_docx(final_obj, review_docx_path, append=True)


# -------------------------
# PDF TEXT + VIEW
# -------------------------
def extract_pdf_text(pdf_path):
    doc = fitz.open(pdf_path)
    parts = []
    for i in range(doc.page_count):
        parts.append(doc.load_page(i).get_text("text"))
    doc.close()
    return "\n".join(parts)

def make_view(full_text, max_chars=MAX_VIEW_CHARS):
    t = re.sub(r"[ \t]+\n", "\n", full_text)
    t = re.sub(r"\n{3,}", "\n\n", t)
    tl = t.lower()

    def win_at(needle, span=12000):
        idx = tl.find(needle)
        if idx == -1:
            return ""
        start = max(0, idx - 1500)
        end = min(len(t), idx + span)
        return t[start:end]

    chunks = []
    chunks.append(t[:7000])
    for key in [
        "abstract",
        "introduction",
        "methods",
        "materials and methods",
        "results",
        "discussion",
        "conclusion",
        "table",
        "supplement",
    ]:
        c = win_at(key)
        if c:
            chunks.append("\n\n===== " + key.upper() + " (WINDOW) =====\n" + c)

    combined = "\n".join(chunks)
    return combined[:max_chars]


# -------------------------
# PROMPTS
# -------------------------
INCLUDED_KEYS = [
    "pmid","author","year","study_design",
    "n_pts","age_mean_years","gender_male_n","gender_female_n",
    "site_maxilla","site_mandible","site_both",
    "primary_cause_breast_cancer","primary_cause_prostate_cancer","primary_cause_mm","primary_cause_osteoporosis","primary_cause_other","primary_cause_other_details",
    "ards_bisphosphonates_alendronate","ards_bisphosphonates_zoledronate","ards_bisphosphonates_risedronate","ards_bisphosphonates_neridronate","ards_bisphosphonates_pamidronate",
    "ards_bisphosphonates_others","ards_bisphosphonates_others_details",
    "ards_denosumab","ards_both","ards_other_drug","ards_other_drug_details",
    "route_iv","route_oral","route_im","route_subcutaneous","route_both","route_not_reported",
    "mronj_stage_at_risk","mronj_stage_0",
    "prevention_technique","group_intervention","group_control",
    "follow_up_mean_months","follow_up_range","outcome_variable",
    "mronj_development","mronj_development_details",
]

INCLUDED_KEY_GROUPS = {
    "Identification": ["pmid", "author", "year", "study_design"],
    "Participants & demographics": ["n_pts", "age_mean_years", "gender_male_n", "gender_female_n"],
    "Lesion site": ["site_maxilla", "site_mandible", "site_both"],
    "Primary cause": [
        "primary_cause_breast_cancer",
        "primary_cause_prostate_cancer",
        "primary_cause_mm",
        "primary_cause_osteoporosis",
        "primary_cause_other",
        "primary_cause_other_details",
    ],
    "Anti-resorptive drugs (ARDs)": [
        "ards_bisphosphonates_alendronate",
        "ards_bisphosphonates_zoledronate",
        "ards_bisphosphonates_risedronate",
        "ards_bisphosphonates_neridronate",
        "ards_bisphosphonates_pamidronate",
        "ards_bisphosphonates_others",
        "ards_bisphosphonates_others_details",
        "ards_denosumab",
        "ards_both",
        "ards_other_drug",
        "ards_other_drug_details",
    ],
    "Administration route": [
        "route_iv",
        "route_oral",
        "route_im",
        "route_subcutaneous",
        "route_both",
        "route_not_reported",
    ],
    "MRONJ staging": ["mronj_stage_at_risk", "mronj_stage_0"],
    "Interventions & groups": ["prevention_technique", "group_intervention", "group_control"],
    "Follow-up & outcomes": [
        "follow_up_mean_months",
        "follow_up_range",
        "outcome_variable",
        "mronj_development",
        "mronj_development_details",
    ],
}

APPRAISAL_KEY_GROUPS = {
    "rct_appraisal": ["q1_randomized", "q2_randomization_method", "q3_double_blind", "q4_blinding_method", "q5_withdrawals", "total_score"],
    "cohort_appraisal": ["q1_clear_question", "q2_cohort_recruited", "q3_exposure_measured", "q4_outcome_measured", "q5_confounders", "q6_followup_complete", "total_score"],
    "case_series_appraisal": ["q1_clear_aim", "q2_inclusion_criteria", "q3_consecutive_cases", "q4_outcomes_defined", "q5_followup_sufficient", "q6_statistical_analysis", "total_score"],
    "case_control_appraisal": ["q1_clear_question", "q2_cases_representative", "q3_controls_selected", "q4_exposure_measured", "q5_confounders", "q6_results_precise", "total_score"],
    "systematic_appraisal": ["q1_focus_question", "q2_inclusion_criteria", "q3_comprehensive_search", "q4_6_search_and_duplication", "q7_quality_assessed", "q8_combining_appropriate", "q9_conclusions_supported", "total_score"],
}

def build_grouped_keys_text(groups):
    blocks = []
    for title, keys in groups.items():
        blocks.append(f"- {title}: {', '.join(keys)}")
    return "\n".join(blocks)

DRIVER_SYSTEM = (
    "You are an evidence extraction agent for MRONJ prevention literature.\n"
    "Quality > one-shot completeness: accuracy is more important than filling every field.\n"
    "Use ONLY the provided paper text. Do not guess.\n"
    "If uncertain, use null and lower confidence.\n"
    "Evidence must be short (1 sentence), no long quotes.\n"
    "You MUST return strict JSON that matches the provided schema.\n"
)

DRIVER_USER_TEMPLATE = (
    "TASK:\n"
    "A) Identify paper_id (pmid/doi/title) if present.\n"
    "B) Classify study_type as one of: rct|cohort|case_series|case_control|systematic_review|other|unclear.\n"
    "C) Fill record.sheets.included_articles with the keys listed below (use null if not reported).\n"
    "D) Fill record.sheets.level_of_evidence if the paper explicitly states it; else null.\n"
    "E) Fill exactly ONE appraisal sheet based on study_type, others must be null:\n"
    "   - rct -> rct_appraisal\n"
    "   - cohort -> cohort_appraisal\n"
    "   - case_series -> case_series_appraisal\n"
    "   - case_control -> case_control_appraisal\n"
    "   - systematic_review -> systematic_appraisal\n"
    "   - other/unclear -> all appraisal sheets null\n"
    "F) Appraisal questions: set 1 for Yes, 0 for No, null for unclear/not stated.\n"
    "G) critical_decisions: MUST contain an entry for study_type AND for EVERY non-null key you set anywhere in record.sheets.*.\n"
    "   Each entry MUST include:\n"
    "     - path (JSON pointer)\n"
    "     - value (the exact value you set)\n"
    "     - evidence (1 sentence)\n"
    "     - is_critical=true\n"
    "\n"
    "Normalization rules (important):\n"
    "- mronj_development must be one of: Yes|No|Unclear|NR\n"
    "- Site flags: set maxilla/mandible/both as applicable (null if NR).\n"
    "- Route flags: set the most specific route(s); if truly not reported set route_not_reported=1.\n"
    "- Drug flags: set specific bisphosphonate subtype(s) if stated; denosumab if stated; ards_both if both.\n"
    "\n"
    "Included Articles keys to fill (grouped):\n"
    f"{build_grouped_keys_text(INCLUDED_KEY_GROUPS)}\n"
    "\n"
    "Work group-by-group. If you cannot find evidence for a group, leave those fields null.\n"
    "Do not force completion; accuracy and verifiable evidence are required.\n"
    "\n"
    "Appraisal sheet keys by study_type (fill only the matching sheet):\n"
    f"{build_grouped_keys_text(APPRAISAL_KEY_GROUPS)}\n"
    "\n"
    "PAPER_TEXT (VIEW):\n"
    "{VIEW}\n"
)

VERIFIER_SYSTEM = (
    "You are an independent verifier.\n"
    "Check whether each listed decision is supported by the provided paper text.\n"
    "For each decision: return AGREE, DISAGREE (with proposed_value), or UNSURE.\n"
    "Evidence must be short (1 sentence), no long quotes.\n"
    "If DISAGREE, propose the minimal corrected value.\n"
    "Also provide suggested_patch as a minimal JSON object patch (only the corrected fields).\n"
    "Return strict JSON that matches the provided schema.\n"
)

VERIFIER_USER_TEMPLATE = (
    "PAPER_TEXT (VIEW):\n"
    "{VIEW}\n\n"
    "DRIVER_JSON (context):\n"
    "{DRIVER_JSON}\n\n"
    "DECISIONS_TO_REVIEW (only review these):\n"
    "{DECISIONS_TO_REVIEW}\n"
)

SUPERVISOR_SYSTEM = (
    "You are an independent adjudicator.\n"
    "Given paper text and two candidate values for a single field, choose the better supported value.\n"
    "Return PICK_A, PICK_B, or UNSURE.\n"
    "Evidence must be short (1 sentence), no long quotes.\n"
    "Return strict JSON that matches the provided schema.\n"
)

SUPERVISOR_USER_TEMPLATE = (
    "PAPER_TEXT (VIEW):\n"
    "{VIEW}\n\n"
    "FIELD_PATH:\n"
    "{PATH}\n\n"
    "CANDIDATE_A:\n"
    "{VAL_A}\n\n"
    "CANDIDATE_B:\n"
    "{VAL_B}\n"
)


# -------------------------
# LLM CALLS
# -------------------------
def openai_driver_extract(oai_client, view_text):
    driver_user = DRIVER_USER_TEMPLATE.replace("{VIEW}", view_text)
    resp = oai_client.responses.create(
        model=OPENAI_DRIVER_MODEL,
        reasoning={"effort": REASONING_EFFORT_OPENAI},
        input=[
            {"role": "system", "content": DRIVER_SYSTEM},
            {"role": "user", "content": driver_user},
        ],
        text={"format": {"type": "json_schema", "name": "mronj_prevention_driver", "schema": DRIVER_SCHEMA, "strict": True}},
    )
    return json.loads(resp.output_text)

def gemini_driver_extract(gclient, view_text):
    driver_user = DRIVER_USER_TEMPLATE.replace("{VIEW}", view_text)
    resp = gclient.models.generate_content(
        model=GEMINI_DRIVER_MODEL,
        contents=driver_user,
        config=types.GenerateContentConfig(
            system_instruction=DRIVER_SYSTEM,
            response_mime_type="application/json",
            response_json_schema=DRIVER_SCHEMA,
            thinking_config=types.ThinkingConfig(thinking_level=THINKING_LEVEL_GEMINI),
            temperature=0.0,
        ),
    )
    return json.loads(resp.text)

def gemini_verify_chunk(gclient, view_text, driver_json, decisions_to_review):
    verifier_user = VERIFIER_USER_TEMPLATE.format(
        VIEW=view_text,
        DRIVER_JSON=json.dumps(driver_json, ensure_ascii=True),
        DECISIONS_TO_REVIEW=json.dumps(decisions_to_review, ensure_ascii=True),
    )
    resp = gclient.models.generate_content(
        model=GEMINI_VERIFIER_MODEL,
        contents=verifier_user,
        config=types.GenerateContentConfig(
            system_instruction=VERIFIER_SYSTEM,
            response_mime_type="application/json",
            response_json_schema=VERIFIER_SCHEMA,
            thinking_config=types.ThinkingConfig(thinking_level=THINKING_LEVEL_GEMINI),
            temperature=0.0,
        ),
    )
    return json.loads(resp.text)

def openai_verify_chunk(oai_client, view_text, driver_json, decisions_to_review):
    verifier_user = VERIFIER_USER_TEMPLATE.format(
        VIEW=view_text,
        DRIVER_JSON=json.dumps(driver_json, ensure_ascii=True),
        DECISIONS_TO_REVIEW=json.dumps(decisions_to_review, ensure_ascii=True),
    )
    resp = oai_client.responses.create(
        model=OPENAI_VERIFIER_MODEL,
        reasoning={"effort": REASONING_EFFORT_OPENAI},
        input=[
            {"role": "system", "content": VERIFIER_SYSTEM},
            {"role": "user", "content": verifier_user},
        ],
        text={"format": {"type": "json_schema", "name": "mronj_prevention_verifier", "schema": VERIFIER_SCHEMA, "strict": True}},
    )
    return json.loads(resp.output_text)

def openai_supervise_decision(oai_client, view_text, path, val_a, val_b):
    supervisor_user = SUPERVISOR_USER_TEMPLATE.format(
        VIEW=view_text,
        PATH=path,
        VAL_A=json.dumps(val_a, ensure_ascii=True),
        VAL_B=json.dumps(val_b, ensure_ascii=True),
    )
    resp = oai_client.responses.create(
        model=SUPERVISOR_MODEL,
        reasoning={"effort": REASONING_EFFORT_OPENAI},
        input=[
            {"role": "system", "content": SUPERVISOR_SYSTEM},
            {"role": "user", "content": supervisor_user},
        ],
        text={"format": {"type": "json_schema", "name": "mronj_supervisor", "schema": SUPERVISOR_SCHEMA, "strict": True}},
    )
    return json.loads(resp.output_text)


# -------------------------
# DECISION LIST + CHUNKING
# -------------------------
def build_decisions_from_driver(driver_out):
    # Start with driver-provided decisions
    out = []
    seen = set()

    for cd in (driver_out.get("critical_decisions") or []):
        path = cd.get("path")
        if not path or path in seen:
            continue
        out.append({
            "path": path,
            "value": cd.get("value"),
            "evidence": cd.get("evidence", ""),
            "is_critical": True,
        })
        seen.add(path)

    # Ensure /study_type is present
    if "/study_type" not in seen:
        out.append({
            "path": "/study_type",
            "value": driver_out.get("study_type"),
            "evidence": "Driver classification; verify against methods/abstract.",
            "is_critical": True,
        })
        seen.add("/study_type")

    # Ensure every leaf in record.sheets.* has a decision entry
    leaf_paths = infer_all_leaf_paths(driver_out)
    for p in leaf_paths:
        if p in seen:
            continue
        v = json_pointer_get(driver_out, p)
        # include null leaves too (still a decision), but keep evidence empty
        out.append({
            "path": p,
            "value": v,
            "evidence": "",
            "is_critical": True,
        })
        seen.add(p)

    return out

def chunk_list(xs, n):
    return [xs[i:i+n] for i in range(0, len(xs), n)]


# -------------------------
# RUNNER (Colab-ready)
# -------------------------
def _progress(progress_fn, message):
    if progress_fn:
        ts = datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S")
        progress_fn(f"[{ts} UTC] {message}")


def run_pipeline_for_pdf(
    pdf_path,
    oai_client,
    gclient,
    template_xlsx,
    out_xlsx,
    out_docx,
    progress_fn=print,
):
    _progress(progress_fn, f"Starting PDF: {pdf_path}")
    _progress(progress_fn, "Extracting text from PDF...")
    full_text = extract_pdf_text(pdf_path)
    _progress(progress_fn, f"PDF text extracted (chars={len(full_text)}). Building view...")
    view = make_view(full_text)
    _progress(progress_fn, f"View built (chars={len(view)}). Calling driver model...")

    driver_out = openai_driver_extract(oai_client, view)
    _progress(progress_fn, "Driver model completed. Building decision list...")

    # Build full decision list, then verify in chunks
    decisions = build_decisions_from_driver(driver_out)
    decision_chunks = chunk_list(decisions, VERIFIER_CHUNK_SIZE)
    _progress(progress_fn, f"Verifier round 1: {len(decision_chunks)} chunk(s).")

    verifier_passes = []
    working_driver = copy.deepcopy(driver_out)

    # Round 1: verify all decisions
    for idx, ch in enumerate(decision_chunks, 1):
        _progress(progress_fn, f"Verifier round 1: chunk {idx}/{len(decision_chunks)}...")
        vpass = gemini_verify_chunk(gclient, view, sanitize_for_model_input(working_driver), ch)
        verifier_passes.append(vpass)

        patch = vpass.get("suggested_patch")
        if isinstance(patch, dict) and patch:
            working_driver = deep_merge(working_driver, patch)

    # Round 2: re-verify only DISAGREE/UNSURE paths after patching
    flagged_paths = []
    for p in verifier_passes:
        for dr in (p.get("decision_reviews") or []):
            if dr.get("status") in ("DISAGREE", "UNSURE"):
                flagged_paths.append(dr.get("path"))

    flagged_paths = [p for p in flagged_paths if p]
    flagged_paths = list(dict.fromkeys(flagged_paths))  # de-dup preserving order
    if flagged_paths:
        _progress(progress_fn, f"Verifier round 2: {len(flagged_paths)} flagged decision(s).")
        flagged_decisions = []
        for p in flagged_paths:
            flagged_decisions.append({
                "path": p,
                "value": json_pointer_get(working_driver, p),
                "evidence": "",
                "is_critical": True,
            })
        flagged_chunks = chunk_list(flagged_decisions, VERIFIER_CHUNK_SIZE)
        for idx, ch in enumerate(flagged_chunks, 1):
            _progress(progress_fn, f"Verifier round 2: chunk {idx}/{len(flagged_chunks)}...")
            vpass2 = gemini_verify_chunk(gclient, view, sanitize_for_model_input(working_driver), ch)
            verifier_passes.append(vpass2)
            patch2 = vpass2.get("suggested_patch")
            if isinstance(patch2, dict) and patch2:
                working_driver = deep_merge(working_driver, patch2)
    else:
        _progress(progress_fn, "Verifier round 2 skipped (no flagged decisions).")

    _progress(progress_fn, "Building final object + writing outputs...")
    final_obj = build_final_object(working_driver, verifier_passes, verifier_model=GEMINI_VERIFIER_MODEL, version="2.1")

    # Persist: for first PDF, template_xlsx is the original template.
    # For subsequent PDFs in the same run, call with template_xlsx=out_xlsx to accumulate rows.
    apply_final_to_outputs(final_obj, template_xlsx, out_xlsx, EXCEL_MAP, out_docx)

    # Also dump audit JSON alongside (optional)
    audit_path = out_xlsx.replace(".xlsx", f".audit_{(final_obj.get('paper_id') or {}).get('pmid')}.json")
    with open(audit_path, "w", encoding="utf-8") as f:
        json.dump(final_obj, f, ensure_ascii=False, indent=2)

    _progress(progress_fn, f"Completed PDF: {pdf_path}")
    return final_obj


def run_driver_verifier_pair(
    view_text,
    driver_fn,
    verify_fn,
    verifier_model,
    progress_fn=print,
):
    driver_out = driver_fn(view_text)
    decisions = build_decisions_from_driver(driver_out)
    decision_chunks = chunk_list(decisions, VERIFIER_CHUNK_SIZE)

    verifier_passes = []
    working_driver = copy.deepcopy(driver_out)

    for idx, ch in enumerate(decision_chunks, 1):
        _progress(progress_fn, f"Verifier round 1: chunk {idx}/{len(decision_chunks)}...")
        vpass = verify_fn(view_text, sanitize_for_model_input(working_driver), ch)
        verifier_passes.append(vpass)
        patch = vpass.get("suggested_patch")
        if isinstance(patch, dict) and patch:
            working_driver = deep_merge(working_driver, patch)

    flagged_paths = []
    for p in verifier_passes:
        for dr in (p.get("decision_reviews") or []):
            if dr.get("status") in ("DISAGREE", "UNSURE"):
                flagged_paths.append(dr.get("path"))

    flagged_paths = [p for p in flagged_paths if p]
    flagged_paths = list(dict.fromkeys(flagged_paths))
    if flagged_paths:
        _progress(progress_fn, f"Verifier round 2: {len(flagged_paths)} flagged decision(s).")
        flagged_decisions = []
        for p in flagged_paths:
            flagged_decisions.append({
                "path": p,
                "value": json_pointer_get(working_driver, p),
                "evidence": "",
                "is_critical": True,
            })
        flagged_chunks = chunk_list(flagged_decisions, VERIFIER_CHUNK_SIZE)
        for idx, ch in enumerate(flagged_chunks, 1):
            _progress(progress_fn, f"Verifier round 2: chunk {idx}/{len(flagged_chunks)}...")
            vpass2 = verify_fn(view_text, sanitize_for_model_input(working_driver), ch)
            verifier_passes.append(vpass2)
            patch2 = vpass2.get("suggested_patch")
            if isinstance(patch2, dict) and patch2:
                working_driver = deep_merge(working_driver, patch2)
    else:
        _progress(progress_fn, "Verifier round 2 skipped (no flagged decisions).")

    final_obj = build_final_object(working_driver, verifier_passes, verifier_model=verifier_model, version="2.1")
    return final_obj


def run_pipeline_for_pdf_abba(
    pdf_path,
    oai_client,
    gclient,
    template_xlsx,
    out_xlsx,
    out_docx,
    progress_fn=print,
    supervisor_enabled=True,
):
    _progress(progress_fn, f"Starting PDF (ABBA): {pdf_path}")
    _progress(progress_fn, "Extracting text from PDF...")
    full_text = extract_pdf_text(pdf_path)
    _progress(progress_fn, f"PDF text extracted (chars={len(full_text)}). Building view...")
    view = make_view(full_text)

    _progress(progress_fn, "Pass A: OpenAI driver -> Gemini verifier.")
    final_a = run_driver_verifier_pair(
        view_text=view,
        driver_fn=lambda t: openai_driver_extract(oai_client, t),
        verify_fn=lambda t, d, c: gemini_verify_chunk(gclient, t, d, c),
        verifier_model=GEMINI_VERIFIER_MODEL,
        progress_fn=progress_fn,
    )

    _progress(progress_fn, "Pass B: Gemini driver -> OpenAI verifier.")
    final_b = run_driver_verifier_pair(
        view_text=view,
        driver_fn=lambda t: gemini_driver_extract(gclient, t),
        verify_fn=lambda t, d, c: openai_verify_chunk(oai_client, t, d, c),
        verifier_model=OPENAI_VERIFIER_MODEL,
        progress_fn=progress_fn,
    )

    payload_a = extract_report_payload(final_a)
    payload_b = extract_report_payload(final_b)
    mismatches = compare_reports_with_tolerance(payload_a, payload_b)

    if mismatches:
        _progress(progress_fn, f"ABBA compare: {len(mismatches)} mismatch(es).")
    else:
        _progress(progress_fn, "ABBA compare: reports match within tolerance.")

    resolved_payload = copy.deepcopy(payload_a)
    unresolved = []
    supervision_notes = []

    if mismatches and supervisor_enabled:
        _progress(progress_fn, "Supervisor adjudication starting...")
        for mm in mismatches:
            decision = openai_supervise_decision(
                oai_client=oai_client,
                view_text=view,
                path=mm["path"],
                val_a=mm["a"],
                val_b=mm["b"],
            )
            status = decision.get("status")
            if status == "PICK_B":
                json_pointer_set(resolved_payload, mm["path"], mm["b"])
            elif status == "PICK_A":
                json_pointer_set(resolved_payload, mm["path"], mm["a"])
            else:
                unresolved.append(mm)
            supervision_notes.append({
                "path": mm["path"],
                "status": status,
                "evidence": decision.get("evidence", ""),
                "confidence": decision.get("confidence", 0.0),
            })

    if mismatches and (unresolved or not supervisor_enabled):
        raise RuntimeError(f"ABBA mismatch unresolved ({len(unresolved)}). Refusing to write outputs.")

    final_obj = copy.deepcopy(final_a)
    final_obj["record"] = resolved_payload.get("record")
    final_obj["study_type"] = resolved_payload.get("study_type")
    final_obj["paper_id"] = resolved_payload.get("paper_id")
    compute_scores_inplace(final_obj)
    extra_issues = rule_validation(final_obj)
    base_validation = final_obj.get("validation") or {}
    base_issues = base_validation.get("issues") or []
    final_obj["validation"] = {
        "needs_human_review": any(i.get("severity") == "CRITICAL" for i in (base_issues + extra_issues)),
        "issues": base_issues + extra_issues,
    }
    final_obj["verification"]["abba_compare"] = {
        "mismatch_count": len(mismatches),
        "supervised": bool(mismatches and supervisor_enabled),
        "supervision_notes": supervision_notes,
    }

    _progress(progress_fn, "Building final object + writing outputs...")
    apply_final_to_outputs(final_obj, template_xlsx, out_xlsx, EXCEL_MAP, out_docx)

    audit_path = out_xlsx.replace(".xlsx", f".audit_{(final_obj.get('paper_id') or {}).get('pmid')}.json")
    with open(audit_path, "w", encoding="utf-8") as f:
        json.dump(final_obj, f, ensure_ascii=False, indent=2)

    _progress(progress_fn, f"Completed PDF (ABBA): {pdf_path}")
    return final_obj

def run_pipeline(
    pdf_paths=None,
    template_xlsx=TEMPLATE_XLSX,
    out_xlsx=OUT_XLSX,
    out_docx=OUT_DOCX,
    openai_api_key=None,
    google_api_key=None,
    progress_fn=print,
    use_abba=True,
    supervisor_enabled=True,
):
    if not pdf_paths:
        raise RuntimeError("pdf_paths is empty. Provide at least one PDF path.")
    if not os.path.exists(template_xlsx):
        raise FileNotFoundError(template_xlsx)

    openai_key = openai_api_key or os.getenv("OPENAI_API_KEY")
    google_key = google_api_key or os.getenv("GOOGLE_API_KEY")
    if not openai_key:
        raise RuntimeError("Missing OPENAI_API_KEY (env var or openai_api_key arg).")
    if not google_key:
        raise RuntimeError("Missing GOOGLE_API_KEY (env var or google_api_key arg).")

    oai_client = OpenAI(api_key=openai_key)
    gclient = genai.Client(api_key=google_key)

    current_template = template_xlsx
    finals = []

    for pdf in pdf_paths:
        if not os.path.exists(pdf):
            raise FileNotFoundError(pdf)

        if use_abba:
            final_obj = run_pipeline_for_pdf_abba(
                pdf_path=pdf,
                oai_client=oai_client,
                gclient=gclient,
                template_xlsx=current_template,
                out_xlsx=out_xlsx,
                out_docx=out_docx,
                progress_fn=progress_fn,
                supervisor_enabled=supervisor_enabled,
            )
        else:
            final_obj = run_pipeline_for_pdf(
                pdf_path=pdf,
                oai_client=oai_client,
                gclient=gclient,
                template_xlsx=current_template,
                out_xlsx=out_xlsx,
                out_docx=out_docx,
                progress_fn=progress_fn,
            )
        current_template = out_xlsx
        finals.append(final_obj)

        pid = final_obj.get("paper_id") or {}
        _progress(
            progress_fn,
            "DONE pdf="
            + str(pdf)
            + " pmid="
            + str(pid.get("pmid"))
            + " study_type="
            + str(final_obj.get("study_type"))
            + " needs_human_review="
            + str((final_obj.get("validation") or {}).get("needs_human_review")),
        )

    _progress(progress_fn, f"WROTE_XLSX: {out_xlsx}")
    _progress(progress_fn, f"WROTE_DOCX: {out_docx}")
    return finals
