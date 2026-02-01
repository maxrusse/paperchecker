# PaperChecker

LLM-powered pipeline for extracting structured data from medical research PDFs.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/maxrusse/paperchecker/blob/main/paperchecker_colab.ipynb)

## Overview

PaperChecker automates the extraction of structured data from medical research papers (specifically MRONJ prevention studies) into Excel templates. It uses multi-round LLM extraction with verification to ensure accuracy.

**Key Features:**
- Extracts data from PDF research papers into a structured Excel format
- Multi-round extraction reduces hallucinations (metadata, population, drugs, interventions, appraisal)
- Verification pass reviews all non-null decisions
- Supports OpenAI and Google Gemini models
- Generates audit trails and review logs
- Optional PubMed lookup for missing PMIDs

## Quick Start (Google Colab)

1. Click the **Open in Colab** badge above
2. Add your API keys to Colab Secrets:
   - `OPENAI_API_KEY`
   - `GOOGLE_API_KEY`
3. Upload your PDF files
4. Run the pipeline and download results

## Local Installation

```bash
# Clone the repository
git clone https://github.com/maxrusse/paperchecker.git
cd paperchecker

# Install dependencies
pip install -r requirements.txt

# Set API keys
export OPENAI_API_KEY="your-key"
export GOOGLE_API_KEY="your-key"

# Optional PubMed lookup
export PUBMED_API_KEY="your-key"
export PUBMED_EMAIL="you@example.com"
```

## Usage

```python
from script import run_pipeline

results = run_pipeline(
    pdf_paths=["paper1.pdf", "paper2.pdf"],
    out_xlsx="output/mronj_extraction.xlsx",
    out_docx="output/mronj_review_log.docx",
)
```

## Pipeline Overview

```
PDF files
    |
    v
[Text Extraction] -- PyMuPDF
    |
    v
[LLM Extraction] -- 5 focused rounds
    |  1. Metadata + study design
    |  2. Population (n, age, gender)
    |  3. Indications + drugs + route
    |  4. Interventions + outcomes
    |  5. Critical appraisal
    |
    v
[Verification Pass] -- Review non-null decisions
    |
    v
[Outputs]
    ├── Excel workbook (filled template)
    ├── Word review log
    └── JSON audit files
```

## Configuration

Edit `script.py` to customize:

| Setting | Default | Description |
|---------|---------|-------------|
| `OPENAI_MODEL` | `gpt-5.2` | OpenAI model for extraction |
| `GEMINI_MODEL` | `gemini-3-pro-preview` | Google model for verification |
| `ENABLE_PUBMED_LOOKUP` | `True` | Auto-fetch missing PMIDs |

## Outputs

- **Excel workbook**: `output/mronj_extraction_YYYYMMDD_HHMMSS.xlsx` (filled template with all study data)
- **Word review log**: `output/mronj_review_log_YYYYMMDD_HHMMSS.docx` (verifier decisions and conflicts)
- **Audit JSON files**: `output/mronj_extraction_YYYYMMDD_HHMMSS.audit_<PMID>.json` (per-paper evidence)

## Documentation

See [docs.md](docs.md) for detailed technical documentation including:
- Excel template structure
- Field definitions
- Critical appraisal checklist specifications
- LLM task design rationale

See [extraction_hints.md](extraction_hints.md) for the keyword hints and field-level extraction rules.

## Requirements

- Python 3.9+
- OpenAI API key (for extraction)
- Google API key (for verification)
- Optional PubMed API key + email (for PMID lookup)

## License

MIT License - see [LICENSE](LICENSE) for details.
