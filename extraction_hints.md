# Extraction hints by feature/item

This document summarizes what the LLM is told to extract (per field) and the keyword hints used to retrieve and verify evidence for each extracted feature/item in the pipeline. The intent is to show both **what drives the LLM answers** and **what it is expected to output**.

## Global extraction rules (applies to all tasks)
- Use ONLY the provided text; do not guess.
- If not reported, return null.
- For 1/0 flag fields: use 1 when explicitly present, otherwise null.
- Evidence must be short (1 sentence), no long quotes.
- If a page number can be inferred from a snippet header (e.g., “--- PAGE 3”), include that page.
- Return strict JSON matching the schema.

## Task-level retrieval hints (keyword views)
These are the keywords passed to `make_task_view(...)` to build the focused text snippet for each extraction task.

### Task 1: meta/design
**Fields covered and expectations**:
- paper_id: doi, title (from paper header/abstract); leave pmid null (resolved via PubMed)
- study_type: MUST be one of the allowed enum values
- included_articles.author: first author surname (e.g., “Smith”)
- included_articles.year: publication year as integer
- included_articles.study_design: brief description (e.g., “Retrospective cohort”)
- level_of_evidence.level_of_evidence: e.g. “1a”, “2b”, “III” — only if explicitly stated
- level_of_evidence.grade_of_recommendation: e.g. “A”, “B”, “C” — only if explicitly stated

**Hint keywords**:
- pmid
- doi
- random
- cohort
- case
- systematic review
- methods
- abstract
- level of evidence
- grade
- recommendation
- oxford
- sign

### Task 2: population
**Fields covered and expectations**:
- n_pts: total number of patients/participants as integer
- age_mean_years: mean age in years as number (e.g., 65.4)
- gender_male_n: number of male participants as integer
- gender_female_n: number of female participants as integer
- NOTE: Leave paper_id and study_type as null (handled by another task)

**Hint keywords**:
- participants
- patients
- sample
- n=
- mean age
- male
- female
- table 1

### Task 3: indication/drugs/route/site
**Fields covered and expectations** (flags: 1 if present, null if not mentioned):
- SITE (where MRONJ occurred): site_maxilla, site_mandible, site_both
- PRIMARY CAUSE/INDICATION: primary_cause_breast_cancer, primary_cause_prostate_cancer, primary_cause_mm (multiple myeloma), primary_cause_osteoporosis, primary_cause_other
- DRUGS — Bisphosphonates: ards_bisphosphonates_zoledronate (Zometa/Reclast), ards_bisphosphonates_pamidronate (Aredia), ards_bisphosphonates_alendronate (Fosamax), ards_bisphosphonates_risedronate (Actonel), ards_bisphosphonates_ibandronate (Boniva), ards_bisphosphonates_etidronate, ards_bisphosphonates_clodronate, ards_bisphosphonates_combination (multiple BPs), ards_bisphosphonates_unknown_other
- DRUGS — Other: ards_denosumab (Prolia/Xgeva), ards_both (BP + denosumab)
- ROUTE of administration: route_iv, route_oral, route_im, route_subcutaneous, route_both, route_not_reported
- NOTE: Leave paper_id and study_type as null (handled by another task)

**Hint keywords**:
- breast
- prostate
- myeloma
- osteoporosis
- zoled
- pamid
- alend
- rised
- iband
- etid
- clodron
- denos
- intraven
- oral
- subcut
- mandible
- maxilla

### Task 4: intervention/outcomes
**Fields covered and expectations**:
- STAGING: mronj_stage_at_risk (number of patients at risk stage, integer), mronj_stage_0 (number of patients at stage 0, integer)
- INTERVENTION: prevention_technique (description of prevention method), group_intervention (description), group_control (description)
- FOLLOW-UP: follow_up_mean_months (mean follow-up duration in months, number), follow_up_range (range as string, e.g., “6-24 months”)
- OUTCOMES: outcome_variable (primary outcome), mronj_development (“Yes” or “No”), mronj_development_details (details about MRONJ cases)
- NOTE: Leave paper_id and study_type as null (handled by another task)

**Hint keywords**:
- prevention
- dental
- extraction
- antibiotic
- photodynamic
- chlorhexidine
- follow-up
- months
- outcome
- mronj
- osteonecrosis

### Task 5: critical appraisal (study-type specific)
**Fields covered and expectations**:
- Fill only the appraisal sheet for the detected study_type (schema varies by study_type).

**Hint keywords**:
- methods
- random
- blind
- withdraw
- confound
- follow up
- loss to follow up
- search strategy
- protocol
- meta-analysis
- risk of bias

## Verifier hints by field (decision keywords)
These keywords are used to build the verification view (`build_verifier_view(...)`) for each decision path. When the field is part of a `q*` appraisal item, the generic appraisal keyword list is used.

### Bibliographic + study descriptors
- pmid: pmid
- doi: doi
- title: title
- author: author
- year: year
- study_design: study design, randomized, cohort, case, systematic review

### Population
- n_pts: participants, patients, sample, n=
- age_mean_years: mean age, age
- gender_male_n: male, men
- gender_female_n: female, women

### Intervention + outcomes
- prevention_technique: prevention, technique
- group_intervention: intervention, treatment
- group_control: control, comparison
- follow_up_mean_months: follow-up, months
- follow_up_range: follow-up, range
- outcome_variable: outcome, endpoint
- mronj_development: mronj, osteonecrosis
- mronj_development_details: mronj, osteonecrosis

### Evidence grading
- level_of_evidence: level of evidence
- grade_of_recommendation: grade of recommendation

### Critical appraisal items
- Any appraisal field with a path leaf starting with `q`: methods, random, blind, confound, follow-up, risk of bias
