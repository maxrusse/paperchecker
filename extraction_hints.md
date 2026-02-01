# Extraction hints by feature/item

This document summarizes what the LLM is told to extract (per field) and the expectations for evidence and output. The pipeline now uses full-text views (no keyword-based snippets), so the emphasis is on accurate extraction and concise evidence.

## Global extraction rules (applies to all tasks)
- Use ONLY the provided text; do not guess.
- If not reported, return null.
- For 1/0 flag fields: use 1 when explicitly present, otherwise null.
- Evidence must be short (1 sentence), no long quotes.
- If a page number can be inferred from a snippet header (e.g., “--- PAGE 3”), include that page.
- Return strict JSON matching the schema.

## Task-level extraction expectations
Each task receives the full-text view of the paper and should limit outputs to the fields listed below.

### Task 1: meta/design
**Fields covered and expectations**:
- paper_id: doi, title (from paper header/abstract); leave pmid null (resolved via PubMed)
- study_type: MUST be one of the allowed enum values
- included_articles.author: first author surname (e.g., “Smith”)
- included_articles.year: publication year as integer
- included_articles.study_design: brief description (e.g., “Retrospective cohort”)
- level_of_evidence.level_of_evidence: e.g. “1a”, “2b”, “III” — only if explicitly stated
- level_of_evidence.grade_of_recommendation: e.g. “A”, “B”, “C” — only if explicitly stated

### Task 2: population
**Fields covered and expectations**:
- n_pts: total number of patients/participants as integer
- age_mean_years: mean age in years as number (e.g., 65.4)
- gender_male_n: number of male participants as integer
- gender_female_n: number of female participants as integer
- NOTE: Leave paper_id and study_type as null (handled by another task)

### Task 3: indication/drugs/route/site
**Fields covered and expectations** (flags: 1 if present, null if not mentioned):
- SITE (where MRONJ occurred): site_maxilla, site_mandible, site_both
- PRIMARY CAUSE/INDICATION: primary_cause_breast_cancer, primary_cause_prostate_cancer, primary_cause_mm (multiple myeloma), primary_cause_osteoporosis, primary_cause_other
- DRUGS — Bisphosphonates: ards_bisphosphonates_zoledronate (Zometa/Reclast), ards_bisphosphonates_pamidronate (Aredia), ards_bisphosphonates_alendronate (Fosamax), ards_bisphosphonates_risedronate (Actonel), ards_bisphosphonates_ibandronate (Boniva), ards_bisphosphonates_etidronate, ards_bisphosphonates_clodronate, ards_bisphosphonates_combination (multiple BPs), ards_bisphosphonates_unknown_other
- DRUGS — Other: ards_denosumab (Prolia/Xgeva), ards_both (BP + denosumab)
- ROUTE of administration: route_iv, route_oral, route_im, route_subcutaneous, route_both, route_not_reported
- NOTE: Leave paper_id and study_type as null (handled by another task)

### Task 4: intervention/outcomes
**Fields covered and expectations**:
- STAGING: mronj_stage_at_risk (number of patients at risk stage, integer), mronj_stage_0 (number of patients at stage 0, integer)
- INTERVENTION: prevention_technique (description of prevention method), group_intervention (description), group_control (description)
- FOLLOW-UP: follow_up_mean_months (mean follow-up duration in months, number), follow_up_range (range as string, e.g., “6-24 months”)
- OUTCOMES: outcome_variable (primary outcome), mronj_development (“Yes” or “No”), mronj_development_details (details about MRONJ cases)
- NOTE: Leave paper_id and study_type as null (handled by another task)

### Task 5: critical appraisal (study-type specific)
**Fields covered and expectations**:
- Fill only the appraisal sheet for the detected study_type (schema varies by study_type).

## Verifier context
The verifier receives the full-text view and the decision list. It does not rely on keyword snippets, so evidence should be concise and clearly anchored to the paper text.
