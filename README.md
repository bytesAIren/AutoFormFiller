# 🏗️ Tender Form Filler CLI

Automatically fills Italian procurement tender forms using your company profile data.  
Supports **.docx** (Word) and **.pdf** formats.

## Quick Start

```bash
# 1. Install dependencies
pip install python-docx pymupdf lxml

# 2. Edit your company profile (one-time setup)
#    Open company_profile.json and replace with YOUR company data
nano company_profile.json

# 3. Fill a form!
python tender_filler.py --form "path/to/your/form.docx"
python tender_filler.py --form "path/to/your/form.pdf"
```

## How It Works

The tool uses **5 strategies** to fill Word documents:

1. **Underscore blanks** — `Il sottoscritto ___________` → `Il sottoscritto Marco Bianchi`
2. **FORMTEXT fields** — Interactive Word form fields (fldChar pattern)
3. **SDT checkboxes** — Structured Document Tags (☐ → ☒)
4. **Table cells** — Label in one cell, value in another
5. **Context runs** — `Ragione sociale: ...` → `Ragione sociale: Idrotech Servizi S.r.l.`

For PDFs:
- **Interactive forms** (AcroForm) — fills form fields directly
- **Flat PDFs** — overlays text at precise coordinates after label text

## Semantic Field Matching

The tool doesn't need exact field names. It uses pattern matching to understand that:
- "Il sottoscritto" = "Nome e Cognome" = "Legale Rappresentante" → all map to your representative's name
- "C.F." = "Codice Fiscale" = "CF persona" → all map to the fiscal code
- "Sede legale" = "con sede in" = "Via" → all map to your registered address

## Company Profile

Edit `company_profile.json` with your data. The structure is:

```json
{
  "azienda": {
    "ragione_sociale": "Your Company Name",
    "cf_piva": "01234567890",
    ...
  },
  "legale_rappresentante": {
    "nome_completo": "First Last",
    "codice_fiscale": "AAABBB00A00A000A",
    ...
  },
  "soci": [...]
}
```

## Options

```
--form, -f     Path to the form file (.docx or .pdf)  [required]
--profile, -p  Path to company profile JSON             [default: company_profile.json]
--output, -o   Output file path                         [default: adds _COMPILATO suffix]
```

## Examples

```bash
# Basic usage
python tender_filler.py -f "Allegato_1.docx"

# Specify profile and output
python tender_filler.py -f "Istanza.docx" -p my_company.json -o "Istanza_filled.docx"

# Fill a PDF
python tender_filler.py -f "Domanda.pdf"
```

## Requirements

- Python 3.8+
- python-docx
- pymupdf (fitz)
- lxml

## 100% Local

All processing happens on YOUR machine. No data is sent anywhere. Your company data stays in the JSON file on your computer.
