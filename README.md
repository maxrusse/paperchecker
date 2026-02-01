# PaperChecker

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/paperchecker/paperchecker/blob/main/paperchecker_colab.ipynb)

## High-level data flow

1. **Input assets**
   - One or more MRONJ-prevention PDFs.
   - The fixed Excel template: `Prevention of MRONJ_Extraction Sheet (Oli).xlsx`.
2. **PDF text extraction**
   - Each PDF is parsed into per-page text.
   - A global view and targeted keyword windows are created to keep LLM prompts focused and small. 
3. **LLM extraction (round-based)**
   - The pipeline runs focused extraction rounds (see below), each returning only a small field subset plus evidence snippets.
   - Outputs are merged into a single working JSON object representing the Excel rows.
4. **Verifier pass**
   - Only non-null decisions are reviewed by a verifier model, chunked for manageable context.
   - Disagreements produce minimal corrective patches.
5. **Post-processing**
   - Appraisal scores are computed (where applicable).
   - Lightweight validation rules flag conflicts (e.g., route flags).
6. **Outputs**
   - Filled Excel workbook (same structure as the template).
   - Optional Word review log summarizing verifier decisions.
   - Per-paper audit JSON for traceability.

## Extraction rounds (tasks per paper)

Each PDF is processed through focused LLM rounds to reduce hallucinations and make verification tractable:

1. **Round 1: Metadata + study design**
   - PMID/DOI/title, author/year, study design, and coarse study-type classification.
2. **Round 2: Population**
   - Sample size, mean age, and male/female counts.
3. **Round 3: Indications + drugs + route + site**
   - Indication flags, ARD drug flags, administration route flags, and maxilla/mandible site flags.
4. **Round 4: Intervention + outcomes**
   - Prevention techniques, intervention/control groups, follow-up, outcomes, and MRONJ development status.
5. **Round 5: Critical appraisal (conditional)**
   - If the study type is RCT/cohort/case series/case-control/systematic review, the relevant appraisal checklist is filled.

## Colab notebook

Use the Colab notebook for an end-to-end run (PDF → JSON → validation → Excel + Word outputs).

1. Open `paperchecker_colab.ipynb` in Colab using the badge above.
2. Install dependencies and set API keys.
3. Upload your PDFs and Excel template.
4. Run the pipeline.
