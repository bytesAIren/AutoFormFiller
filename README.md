# 🏗️ Tender Form Filler CLI

Automatically fills Italian procurement tender forms using your company profile data.
Supports **.docx** (Word) and **.pdf** formats with intelligent field recognition and highlighting of unfilled fields.

## Quick Start

```bash
# 1. Install dependencies
pip install python-docx pymupdf lxml

# 2. Edit your company profile (one-time setup)
#    Open company_profile.json or company_profile.csv and replace with YOUR company data
nano company_profile.json
# or
nano company_profile.csv

# 3. Fill forms automatically!
python tender_filler.py --auto
# This processes all .docx/.pdf in EMPTY_FORM/ and saves filled copies to FILLED_FORM/

# Or fill a single form manually:
python tender_filler.py --form "path/to/your/form.docx"
python tender_filler.py --form "path/to/your/form.pdf"

# Debug mode to see field matching details:
python tender_filler.py --auto --verbose
```

## How It Works

The tool uses **6 strategies** to fill Word documents:

1. **Underscore blanks** — `Il sottoscritto ___________` → `Il sottoscritto Marco Bianchi`
2. **FORMTEXT fields** — Interactive Word form fields (fldChar pattern)
3. **SDT checkboxes** — Structured Document Tags (☐ → ☒)
4. **Table cells** — Label in one cell, value in another
5. **Context runs** — `Ragione sociale: ...` → `Ragione sociale: Idrotech Servizi S.r.l.`
6. **Empty field highlighting** — Unfilled fields are highlighted in **yellow** for easy identification

For PDFs:
- **Interactive forms** (AcroForm) — fills form fields directly
- **Flat PDFs** — overlays text at precise coordinates after label text

## Semantic Field Matching

The tool doesn't need exact field names. It uses pattern matching to understand that:
- "Il sottoscritto" = "Nome e Cognome" = "Legale Rappresentante" → all map to your representative's name
- "C.F." = "Codice Fiscale" = "CF persona" → all map to the fiscal code
- "Sede legale" = "con sede in" = "Via" → all map to your registered address

## Company Profile

You can use either JSON or CSV format for your company profile. CSV is easier to edit in spreadsheets.

**Note**: The example files contain Italian sample data, but you can enter company information in any language. The system recognizes Italian form labels but accepts company data in any language.

### JSON Format (company_profile.json)
```json
{
  "azienda": {
    "ragione_sociale": "Your Company Name",
    "cf_piva": "01234567890",
    "sede_legale": "123 Main Street",
    "telefono": "+39 123 456789",
    "email": "info@company.com",
    "dipendenti_totale": "25",
    "ateco_descrizione": "Manufacturing and services"
  },
  "legale_rappresentante": {
    "nome_completo": "John Doe",
    "codice_fiscale": "AAABBB00A00A000A"
  }
}
```

### CSV Format (company_profile.csv)
Columns: section,key,value,nome,quota,ruolo

For simple fields:
```
section,key,value
azienda,ragione_sociale,Your Company Name
azienda,cf_piva,01234567890
azienda,sede_legale,123 Main Street
azienda,telefono,+39 123 456789
azienda,email,info@company.com
azienda,dipendenti_totale,25
azienda,ateco_descrizione,Manufacturing and services
legale_rappresentante,nome_completo,John Doe
legale_rappresentante,codice_fiscale,AAABBB00A00A000A
```

For shareholders (soci):
```
section,,nome,,quota,ruolo
soci,,Shareholder Name,,50%,Partner
```

## Command Line Options

```
--form, -f     Path to the form file (.docx or .pdf)
--profile, -p  Path to company profile JSON/CSV     [default: company_profile.json]
--output, -o   Output file path                     [default: adds _COMPILATO suffix]
--auto         Process all forms in EMPTY_FORM/ and save to FILLED_FORM/
--verbose, -v  Enable detailed debug logging for field matching
```

## Project Structure

```
AutoFormFiller/
├── EMPTY_FORM/          # Place your empty forms here
├── FILLED_FORM/         # Filled forms are saved here automatically
├── tender_filler.py     # Main script
├── company_profile.json # Company data (JSON format)
├── company_profile.csv  # Company data (CSV format)
└── README.md           # This file
```

## Features

- ✅ **Automatic processing** of all forms in a directory
- ✅ **Dual format support** (JSON and CSV for company profiles)
- ✅ **Intelligent field recognition** using semantic patterns
- ✅ **Empty field highlighting** in yellow for manual review
- ✅ **Debug mode** to troubleshoot field matching issues
- ✅ **Clean output** with proper formatting and no leftover underscores
- ✅ **Multi-language data support** (company info can be in any language)

## Requirements

- Python 3.6+
- python-docx
- pymupdf (PyMuPDF)
- lxml

## License

This project is open source. Feel free to modify and distribute.
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
