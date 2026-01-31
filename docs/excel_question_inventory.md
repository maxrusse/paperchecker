# Excel question inventory (default template)

Source: `Prevention of MRONJ_Extraction Sheet (Oli).xlsx`.

## Per-sheet column counts (per paper row)

| Sheet | Columns (per paper) | Notes |
| --- | ---: | --- |
| Included Articles | 43 | Includes grouped header bands like **Gender**, **Site**, **Primary cause**, **Type of ARDs**, **Route of admin**, **Stage of MRONJ**, and **Groups** that expand into multiple columns. |
| Level of Evidence | 6 | Basic identification + evidence/grade fields. |
| Critical Appraisal of RCTS | 10 | 5 appraisal questions + total score. |
| Critical Appraisal of Cohort | 15 | 11 appraisal questions (no explicit total column in template). |
| Critical Appraisal of Case Seri | 15 | 10 appraisal questions + total score. |
| Critical Appraisal of Case Cont | 14 | 10 appraisal questions (no explicit total column in template). |
| Critical Appraisal of Systemati | 19 | 16 appraisal questions + total score. |

## Column headers (verbatim)

### Included Articles (43 columns)

Note: this sheet uses **three header rows**. Row 1 provides grouped bands (e.g., Gender/Site/Primary cause/Type of ARDs), row 2 breaks those into specific fields (e.g., Male/Female, Maxilla/Mandible/Both), and row 3 further splits the ARD subtypes and MRONJ stage. The per-column list below reflects the **combined per-column prompts** from rows 1–3 (so every blank header cell becomes a distinct subquestion).

#### Grouped subquestions with expected inputs

- **Cols 1–6: Identification & demographics**
  - 1. PMID — expected: numeric ID (integer).
  - 2. Author — expected: string (first author).
  - 3. Year — expected: integer year.
  - 4. Study Design — expected: string (e.g., RCT, cohort, case series).
  - 5. Number of pts — expected: integer.
  - 6. Age (Mean in Years) — expected: number (mean age).

- **Cols 7–8: Gender**
  - 7. Gender / Male — expected: integer count (male n).
  - 8. Female — expected: integer count (female n).

- **Cols 9–11: Site**
  - 9. Site / Maxilla — expected: flag (1/0) or count if reported.
  - 10. Mandible — expected: flag (1/0) or count if reported.
  - 11. Both — expected: flag (1/0) or count if reported.

- **Cols 12–16: Primary cause**
  - 12. Primary cause / Brst cancer — expected: flag (1/0) or count.
  - 13. Prst Cancer — expected: flag (1/0) or count.
  - 14. MM — expected: flag (1/0) or count.
  - 15. Osteoporosis — expected: flag (1/0) or count.
  - 16. Other — expected: flag (1/0) or count (details should go in the “Other details” field in the template if available elsewhere).

- **Cols 17–27: Type of ARDs**
  - 17. Type of ARDs / Bisphosphonates / Zoledronate — expected: flag (1/0) or count.
  - 18. Pamidronate — expected: flag (1/0) or count.
  - 19. Residronate — expected: flag (1/0) or count.
  - 20. Alendronate — expected: flag (1/0) or count.
  - 21. Ibandronate — expected: flag (1/0) or count.
  - 22. Combination — expected: flag (1/0) or count.
  - 23. Etidronate — expected: flag (1/0) or count.
  - 24. Clodronate — expected: flag (1/0) or count.
  - 25. Unknown/Other — expected: flag (1/0) or count (details if present).
  - 26. Denosumab — expected: flag (1/0) or count.
  - 27. Both — expected: flag (1/0) or count (both ARD classes).

- **Cols 28–33: Route of admin**
  - 28. Route of admin / Intra-V — expected: flag (1/0) or count.
  - 29. Oral — expected: flag (1/0) or count.
  - 30. IM — expected: flag (1/0) or count.
  - 31. Subcutaneous — expected: flag (1/0) or count.
  - 32. Both — expected: flag (1/0) or count.
  - 33. Not reported(N/R) — expected: flag (1/0).

- **Cols 34–35: Stage of MRONJ**
  - 34. Stage of MRONJ / At Risk — expected: flag (1/0) or count.
  - 35. Stage 0 — expected: flag (1/0) or count.

- **Cols 36–38: Interventions & groups**
  - 36. Prevention Technique — expected: short string description.
  - 37. Groups / Intervention — expected: short string description.
  - 38. Control — expected: short string description.

- **Cols 39–43: Follow-up & outcomes**
  - 39. Follow-up (Mean in Months) — expected: number.
  - 40. Follow-up Range — expected: string or number range.
  - 41. Outcome variable — expected: short string description.
  - 42. MRONJ Development — expected: categorical (Yes/No/Unclear/NR).
  - 43. If Yes, Details — expected: short string detail.

#### Per-column prompts (combined headers, verbatim)

1. PMID
2. Author
3. Year
4. Study Design
5. Number of pts
6. Age (Mean in Years)
7. Gender / Male
8. Female
9. Site / Maxilla
10. Mandible
11. Both
12. Primary cause / Brst cancer
13. Prst Cancer
14. MM
15. Osteoporosis
16. Other
17. Type of ARDs / Bisphosphonates / Zoledronate
18. Pamidronate
19. Residronate
20. Alendronate
21. Ibandronate
22. Combination
23. Etidronate
24. Clodronate
25. Unknown/Other
26. Denosumab
27. Both
28. Route of admin / Intra-V
29. Oral
30. IM
31. Subcutaneous
32. Both
33. Not reported(N/R)
34. Stage of MRONJ / At Risk
35. Stage 0
36. Prevention Technique
37. Groups / Intervention
38. Control
39. Follow-up (Mean in Months)
40. Follow-up Range
41. Outcome variable
42. MRONJ Development
43. If Yes, Details

### Level of Evidence (6 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. Level of Evidence
6. Grade of Recommendation

### Critical Appraisal of RCTS (10 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. 1. Was the study described as randomized (this includes words such as randomly, random, and randomization)?
6. 2. Was the method used to generate the sequence of randomization described and appropriate (table of random numbers, computer-generated, etc)?
7. 3. Was the study described as double blind?
8. 4. Was the method of double blinding described and appropriate (identical placebo, active placebo, dummy, etc)?
9. 5. Was there a description of withdrawals and dropouts?
10. Total Score (out of 5)

### Critical Appraisal of Cohort (15 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. 1.     Were the two groups similar and recruited from the same population?
6. 2.     Were the exposures measured similarly to assign people to both exposed and unexposed groups?
7. 3.     Was the exposure measured in a valid and reliable way?
8. 4.     Were confounding factors identified?
9. 5.     Were strategies to deal with confounding factors stated?
10. 6.     Were the groups/participants free of the outcome at the start of the study (or at the moment of exposure)?
11. 7.     Were the outcomes measured in a valid and reliable way?
12. 8.     Was the follow up time reported and sufficient to be long enough for outcomes to occur?
13. 9.     Was follow up complete, and if not, were the reasons to loss to follow up described and explored?
14. 10.  Were strategies to address incomplete follow up utilized?
15. 11.  Was appropriate statistical analysis used?

### Critical Appraisal of Case Seri (15 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. Were there clear criteria for inclusion in the case series? 
6. Was the condition measured in a standard, reliable way for all participants included in the case series?
7. Were valid methods used for identification of the condition for all participants included in the case series?
8. Did the case series have consecutive inclusion of participants?
9. Did the case series have complete inclusion of participants?
10. Was there clear reporting of the demographics of the participants in the study?
11. Was there clear reporting of clinical information of the participants?
12. Were the outcomes or follow up results of cases clearly reported?
13. Was there clear reporting of the presenting site(s)/clinic(s) demographic information?
14. Was statistical analysis appropriate?
15. Total Score

### Critical Appraisal of Case Cont (14 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. 1.    Were the groups comparable other than the presence of disease in cases or the absence of disease in controls?
6. 2.    Were cases and controls matched appropriately?
7. 3.    Were the same criteria used for identification of cases and controls?
8. 4.    Was exposure measured in a standard, valid and reliable way?
9. 5.    Was exposure measured in the same way for cases and controls?
10. 6.    Were confounding factors identified? 
11. 7.    Were strategies to deal with confounding factors stated?
12. 8.    Were outcomes assessed in a standard, valid and reliable way for cases and controls?
13. 9.    Was the exposure period of interest long enough to be meaningful?
14. 10.  Was appropriate statistical analysis used?

### Critical Appraisal of Systemati (19 columns)

1. PMID
2. Author
3. Year
4. Study Design
5. 1. Did the research questions and inclusion criteria for the review include the components of PICO?
6. 2. Did the report of the review contain an explicit statement that the review methods were established prior to the conduct of the review and did the report justify any significant deviations from the protocol?
7. 3. Did the review authors explain their selection of the study designs for inclusion in the review?
8. 4. Did the review authors use a comprehensive literature search strategy? 5. Did the review authors perform study selection in duplicate? 6. Did the review authors perform data extraction in duplicate?
9. 7. Did the review authors provide a list of excluded studies and justify the exclusions?
10. 8. Did the review authors describe the included studies in adequate detail?
11. 9. Did the review authors use a satisfactory technique for assessing the risk of bias (RoB) in individual studies that were included in the review?
12. 10. Did the review authors report on the sources of funding for the studies included in the review?
13. 11. If meta-analysis was performed did the review authors use appropriate methods for statistical combination of results?
14. 12. If meta-analysis was performed, did the review authors assess the potential impact of RoB in individual studies on the results of the metaanalysis or other evidence synthesis?
15. 13. Did the review authors account for RoB in individual studies when interpreting/discussing the results of the review?
16. 14. Did the review authors provide a satisfactory explanation for, and discussion of, any heterogeneity observed in the results of the review?
17. 15. If they performed quantitative synthesis did the review authors carry out an adequate investigation of publication bias (small study bias) and discuss its likely impact on the results of the review?
18. 16. Did the review authors report any potential sources of conflict of interest, including any funding they received for conducting the review?
19. Total
