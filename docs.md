# PaperChecker Documentation

This pipeline extracts structured data from medical research PDFs into an Excel template format.

The Excel template structure is auto-generated based on the `EXCEL_MAP` configuration in `script.py`. The pipeline produces:

1) a filled Excel workbook (same layout as template), and  
2) an optional Word "review log" that explains what the LLM claimed and what the verifier agreed/disagreed with.

---

## 1. What this pipeline covers (and what it does not)

### Covered (fits systematic review / guideline evidence tables)
- Per-paper data extraction into **Included Articles** (population, indication, drug exposure, intervention, follow-up, outcome).
- Per-paper critical appraisal into the appropriate checklist sheet:
  - RCTs
  - Cohorts
  - Case series
  - Case-control
  - Systematic reviews
- A verifier step that challenges extracted values against the paper text and flags uncertain/unsupported items.

### Not covered (required for a full S3 guideline workflow)
The Excel template is an *input artifact* to an S3 guideline workflow, but it is not the entire workflow. A classic S3 guideline (AWMF style) additionally needs, at minimum:

- Scope + key questions (PICO) and outcome prioritization.
- Systematic search strategy + documentation (databases, dates, query strings).
- Study selection process + PRISMA-style reporting (screening, exclusions).
- Evidence synthesis:
  - effect sizes (per outcome, per comparison),
  - meta-analysis (if appropriate),
  - heterogeneity exploration,
  - publication bias assessment.
- Certainty of evidence / evidence profiles (often GRADE) and summary-of-findings tables.
- Recommendation development:
  - evidence-to-decision reasoning,
  - structured consensus process,
  - strength of recommendation (and consensus strength).
- External review + update plan.

**Implication:** the Excel output is necessary, but not sufficient, for an S3 guideline. If you need S3-compliant outputs, you should add a second layer that turns extracted study-level data into outcome-level effect estimates and evidence profiles (can be stored in JSON/Word without changing the Excel).

---

## 2. Excel template structure

Sheets (1 row per paper, keyed by PMID):

- **Included Articles** (3 header rows, 43 columns)
- **Level of Evidence** (6 columns)
- **Critical Appraisal of RCTS**
- **Critical Appraisal of Cohort**
- **Critical Appraisal of Case Seri**
- **Critical Appraisal of Case Cont**
- **Critical Appraisal of Systemati**

### Multi-line headers
The **Included Articles** sheet uses **3 header rows** (grouped bands + subfields). Do not treat this as 3 separate rows of questions; it is 1 row per paper.

Some appraisal sheets also have extra header rows with scoring instructions:
- RCTS: 3 header rows (row 3 contains scoring guidance)
- Cohort / Case Series / Case Control: 2 header rows (row 2 contains Yes/No guidance)

---

## 3. Field definitions (Included Articles)

The pipeline maps Excel columns to stable internal keys (used by the JSON outputs and by `script.py`).

### Identification
- `pmid` (PMID)
- `author` (citation text in the template is acceptable; do not overwrite if already filled)
- `year` (publication year)
- `study_design` (must match Excel validation tokens when possible)

Study design tokens used in the template:
- RCT
- Retrospective Cohort
- Prospective Cohort
- Case-Control
- Retrospective Case-Series
- Prospective Case Series
- Systematic Review
- Metaanalysis

### Participants
- `n_pts` (integer, total participants/patients)
- `age_mean_years` (mean age)
- `gender_male_n`, `gender_female_n` (counts)

### Site
- `site_maxilla`, `site_mandible`, `site_both` (use **1** if explicitly stated; otherwise blank)

### Primary cause (indication)
- `primary_cause_breast_cancer`
- `primary_cause_prostate_cancer`
- `primary_cause_mm` (multiple myeloma)
- `primary_cause_osteoporosis`
- `primary_cause_other`

These are **non-exclusive** flags: a study can include multiple indications.

### Anti-resorptive drugs (ARDs)
Bisphosphonates:
- `ards_bisphosphonates_zoledronate`
- `ards_bisphosphonates_pamidronate`
- `ards_bisphosphonates_risedronate`
- `ards_bisphosphonates_alendronate`
- `ards_bisphosphonates_ibandronate`
- `ards_bisphosphonates_combination`
- `ards_bisphosphonates_etidronate`
- `ards_bisphosphonates_clodronate`
- `ards_bisphosphonates_unknown_other`

Other:
- `ards_denosumab`
- `ards_both` (bisphosphonates + denosumab in the study population)

### Route of administration
- `route_iv`, `route_oral`, `route_im`, `route_subcutaneous`
- `route_both` (multiple routes are used)
- `route_not_reported` (only set if the route truly is not reported)

### MRONJ stage at baseline
- `mronj_stage_at_risk`
- `mronj_stage_0`

### Intervention, follow-up, outcomes
- `prevention_technique`
- `group_intervention`, `group_control`
- `follow_up_mean_months`, `follow_up_range`
- `outcome_variable`
- `mronj_development` (Excel token: **Yes** or **No**; blank if unclear)
- `mronj_development_details` (free text; counts and definitions belong here if not otherwise captured)

---

## 4. Critical appraisal sheets (how to fill)

### General rule
For Cohort / Case Series / Case Control / Systematic Review appraisals, the Excel expects strings from:

- **Yes**
- **No**
- **Unclear**
- **Not Applicable**

This matches the drop-down validations in the workbook (where present).

### RCT appraisal (RCTS)
The RCTS sheet is scored with values:
- `q1_randomized`, `q3_double_blind`, `q5_withdrawals_dropouts`: **"0"** or **"1"**
- `q2_randomization_method`, `q4_blinding_method`: **"-1"**, **"0"**, or **"+1"**

A `total_score` (integer) can be computed and written.

---

## 5. LLM task design (why v2 splits tasks)

A single monolithic prompt that tries to fill every sheet tends to:
- miss fields (especially flags),
- mix up concepts (indication vs drug vs route),
- hallucinate negatives (writing 0 where the paper is simply silent),
- create huge verification workloads.

`script.py` splits extraction into 5 smaller tasks:

1) **Meta + design**
2) **Population**
3) **Indication + drugs + route + site**
4) **Intervention + outcomes**
5) **Critical appraisal** (only the relevant checklist, based on study type)

Each task:
- receives the full-text view,
- outputs only a small set of fields,
- attaches short evidence for every non-null value.

The verifier only checks **non-null** decisions, so the review log stays readable.

---

## 6. Notes on "Level of Evidence" and "Grade of Recommendation"

These cannot be filled reliably without defining the framework. In guideline workflows:

- **Level of evidence** may be assigned by design + risk-of-bias (e.g., Oxford/SIGN) or via GRADE certainty.
- **Grade of recommendation** depends on more than design: effect size, consistency, applicability, benefit/harm, patient values, resources, etc.

Recommendation:
- Decide the exact scheme (Oxford vs SIGN vs GRADE vs AWMF-specific) and encode it as rules (or a dedicated LLM task with definitions).
- Keep the Excel columns as-is; store the reasoning in the Word log and/or audit JSON.

---

## 7. Quality gates and common failure modes

### Preferred conventions
- Use **null/blank** for "not reported" (instead of writing 0).
- Use 1 for flags only when explicitly present.
- Do not overwrite pre-filled author/citation cells in the template.

### Checks (implemented lightly in v2)
- `mronj_development` should be Yes/No/blank (template data validation).
- `route_not_reported` must not be set together with a specific route.

### Typical human-review triggers
- Mixed populations where counts for male/female or indications are incomplete.
- Outcomes reported only in tables/figures with poor PDF text extraction.
- Studies using multiple ARDs and/or switching routes over time.

---

## 8. Outputs

- Filled Excel workbook: `mronj_prevention_filled_YYYYMMDD_HHMMSS.xlsx`
- Review log (Word): `mronj_prevention_review_log_YYYYMMDD_HHMMSS.docx`
- Per-paper audit JSON: `...audit_<PMID>.json`
