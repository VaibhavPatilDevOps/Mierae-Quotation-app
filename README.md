# Mierae Invoice/Quotation Manager (Streamlit)

A simple Streamlit app to create, manage, and search quotations/invoices using a DOCX template. Only the yellow-highlighted fields in the template are replaced.

## Requirements
- Windows (for docx2pdf to use Microsoft Word; PDF generation will be skipped if Word is unavailable)
- Python 3.9+

## Setup
```bash
python -m pip install -r requirements.txt
```

## Run
```bash
python -m streamlit run "streamli app.py"
```

The app will create an `invoices.db` SQLite database and an `output/` folder with generated files:
- `output/docx/<Customer+Quotation>.docx`
- `output/pdf/<Customer+Quotation>.pdf` (if conversion succeeds)

## Notes
- Quotation No auto-increments, starting from `RRSS/AP/APEPDCL/VSP/0020`.
- Validity auto-fills as 30 days after the selected Date of Quotation.
- The "Search Invoice" tab supports filtering, downloading, editing, and deleting.
- Editing regenerates DOCX/PDF and updates the record while keeping the original Quotation No.

## Template
Place `Mierae Quotation Template.docx` in the project root (it is already present). Only yellow-highlighted runs are replaced, in this order (based on the sample image):
1. Customer Name
2. Location
3. City
4. State
5. Pincode
6. Phone/Mobile
7. Product & Service
8. Quotation No
9. Date of Quotation
10. Validity of Quotation

If your template's highlighted field order differs, adjust the `values` list in `generate_docx()` callers within `app.py`.
