import os
import io
import sqlite3
from datetime import datetime, timedelta
from typing import Dict, Optional, List, Tuple

import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.text.run import Run
from docx.oxml.shared import OxmlElement, qn
from docx.document import Document as DocxDocument
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx2pdf import convert as docx2pdf_convert

# ---------------------------
# Constants and configuration
# ---------------------------
APP_TITLE = "Mierae Invoice/Quotation Manager"
DB_PATH = os.path.join(os.getcwd(), "invoices.db")
TEMPLATE_PATH = os.path.join(os.getcwd(), "Mierae Quotation Template new.docx")
OUTPUT_DIR = os.path.join(os.getcwd(), "output")
DOCX_DIR = os.path.join(OUTPUT_DIR, "docx")
PDF_DIR = os.path.join(OUTPUT_DIR, "pdf")

PRODUCT_OPTIONS = [
    "3.3 kW Residential Rooftop Solar System",
    "5.5 kW Residential Rooftop Solar System",
]

QUOTATION_PREFIX = "RRSS/AP/APEPDCL/VSP/"
QUOTATION_START_NUMBER = 20  # corresponds to 0020

# ---------------------------
# Utility: database
# ---------------------------

def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quotation_no TEXT UNIQUE NOT NULL,
            product TEXT NOT NULL,
            customer_name TEXT NOT NULL,
            mobile TEXT,
            location TEXT,
            city TEXT,
            state TEXT,
            pincode TEXT,
            staff_name TEXT,
            date_of_quotation TEXT,
            validity_date TEXT,
            docx_path TEXT,
            pdf_path TEXT,
            created_at TEXT,
            updated_at TEXT
        )
        """
    )
    return conn

# ---------------------------
# Helpers for quotation number
# ---------------------------

def next_quotation_no(conn: sqlite3.Connection) -> str:
    cur = conn.cursor()
    cur.execute(
        "SELECT quotation_no FROM invoices WHERE quotation_no LIKE ? ORDER BY id DESC LIMIT 1",
        (f"{QUOTATION_PREFIX}%",),
    )
    row = cur.fetchone()
    if row and isinstance(row[0], str):
        try:
            suffix = int(row[0].split("/")[-1])
            nxt = suffix + 1
        except Exception:
            nxt = QUOTATION_START_NUMBER
    else:
        nxt = QUOTATION_START_NUMBER
    return f"{QUOTATION_PREFIX}{nxt:04d}"

# ---------------------------
# File system helpers
# ---------------------------

def ensure_dirs():
    os.makedirs(DOCX_DIR, exist_ok=True)
    os.makedirs(PDF_DIR, exist_ok=True)


def safe_filename(name: str) -> str:
    """Return a filesystem-safe filename fragment (no path separators or illegal chars)."""
    illegal = ['\\\\', '/', ':', '*', '?', '"', '<', '>', '|']
    safe = name
    for ch in illegal:
        safe = safe.replace(ch, '-')
    # Collapse spaces and trim
    safe = " ".join(safe.split())
    return safe

# ---------------------------
# DOCX processing - replace only yellow highlighted runs
# ---------------------------

def iter_paragraphs_and_cells(doc: DocxDocument) -> List[Paragraph]:
    items: List[Paragraph] = []
    # Paragraphs at document level
    items.extend(doc.paragraphs)
    # Paragraphs inside tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                items.extend(cell.paragraphs)
                # Also nested tables inside a cell
                for tbl in cell.tables:
                    for r in tbl.rows:
                        for c in r.cells:
                            items.extend(c.paragraphs)
    return items


def get_yellow_runs(doc: DocxDocument) -> List[Run]:
    yellow_runs: List[Run] = []
    for p in iter_paragraphs_and_cells(doc):
        for run in p.runs:
            try:
                if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                    yellow_runs.append(run)
            except Exception:
                # Some runs might not have highlight attribute accessible
                pass
    return yellow_runs


def replace_yellow_fields(doc: DocxDocument, values_in_order: List[str]) -> None:
    """Backward-compatible: keep for potential future use."""
    runs = get_yellow_runs(doc)
    for i, run in enumerate(runs):
        if i < len(values_in_order):
            run.text = str(values_in_order[i])


def replace_by_labels(doc: DocxDocument, data: Dict[str, str]) -> None:
    """Replace highlighted values based on paragraph labels to avoid misaligned fields.
    Labels handled:
    - Customer Name:
    - Location:
    - City:
    - State:
    - Phone:
    - Product & Service:
    - Quotation No:
    - Date of Quotation:
    - Validity of Quotation:
    """
    def fmt_date(val: str) -> str:
        try:
            # Accept YYYY-MM-DD and return DD/MM/YYYY
            dt = datetime.strptime(val, "%Y-%m-%d")
            return dt.strftime("%d/%m/%Y")
        except Exception:
            return val or ""

    targets = {
        "customer name": data.get("customer_name", ""),
        "location": data.get("location", "").replace("\n", " ").strip(),
        "city": data.get("city", ""),
        "state": data.get("state", ""),
        # Explicit pincode labels (handle both single word and spaced variant)
        "pincode": data.get("pincode", ""),
        "pin code": data.get("pincode", ""),
        # Support multiple possible labels used in templates for phone/mobile/customer no
        "phone": data.get("mobile", ""),
        "customer no": data.get("mobile", ""),
        "mobile no": data.get("mobile", ""),
        "mobile number": data.get("mobile", ""),
        "product & service": data.get("product", ""),
        "date of quotation": fmt_date(data.get("date_of_quotation", "")),
        "validity of quotation": fmt_date(data.get("validity_date", "")),
        "quotation no": str(data.get("quotation_no", "")),
    }

    def replace_in_paragraph(p: Paragraph, label: str, value: str, all_labels: List[str]):
        text = p.text
        text_lower = text.lower()
        # detect label with ':' or '-'
        candidates = [f"{label}:", f"{label}-"]
        start_idx = -1
        for cand in candidates:
            start_idx = text_lower.find(cand)
            if start_idx != -1:
                start_idx += len(cand)
                break
        if start_idx == -1:
            return  # label not in this paragraph

        # find the next other label occurrence to bound our clearing range
        next_idx = len(text)
        for other in all_labels:
            if other == label:
                continue
            for sep in (":", "-"):
                i = text_lower.find(f"{other}{sep}", start_idx)
                if i != -1:
                    next_idx = min(next_idx, i)
        # iterate runs and find first yellow run whose run range begins after start_idx and before next_idx
        pos = 0
        replaced = False
        for r in p.runs:
            rt = r.text
            begin = pos
            end = pos + len(rt)
            pos = end
            if end <= start_idx:
                continue
            if begin >= next_idx:
                break
            try:
                if r.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                    if not replaced:
                        r.text = "" if value is None else str(value)
                        replaced = True
                    else:
                        # clear leftover highlighted placeholders in this label's region
                        r.text = ""
            except Exception:
                pass

    search_paras = iter_paragraphs_and_cells(doc)
    label_list = list(targets.keys())
    for p in search_paras:
        # Combined State and Pincode in same paragraph special-case still supported
        tl = p.text.lower()
        if ("state:" in tl or "state-" in tl) and ("pincode" in tl or "pin code" in tl):
            values = [data.get("state", ""), data.get("pincode", "")]
            idx = 0
            for r in p.runs:
                try:
                    if r.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                        if idx < len(values):
                            r.text = str(values[idx])
                            idx += 1
                except Exception:
                    pass
            # If we only managed to set State (one yellow run) but not Pincode, try to place pincode by label position
            if idx < 2 and values[1]:
                # Determine run range for text after the word 'pincode'/'pin code'
                txt = p.text
                low = txt.lower()
                pin_label_pos = -1
                for key in ("pincode", "pin code"):
                    pin_label_pos = low.find(key)
                    if pin_label_pos != -1:
                        # move to after possible ':' or '-' and space
                        j = pin_label_pos + len(key)
                        while j < len(txt) and txt[j] in [':', '-', ' ', '\u00A0']:
                            j += 1
                        pin_label_pos = j
                        break
                # Walk runs and set the first run starting at/after pincode label position if it's empty/placeholder
                if pin_label_pos != -1:
                    pos = 0
                    candidate_after = None
                    candidate_after_highlight = None
                    for r in p.runs:
                        begin = pos
                        end = pos + len(r.text)
                        pos = end
                        if begin < pin_label_pos:
                            continue
                        # first run after the label
                        if candidate_after is None:
                            candidate_after = r
                        try:
                            if r.font.highlight_color == WD_COLOR_INDEX.YELLOW and candidate_after_highlight is None:
                                candidate_after_highlight = r
                        except Exception:
                            pass
                    target_run = candidate_after_highlight or candidate_after
                    if target_run is not None:
                        target_run.text = str(values[1])
                    else:
                        # As a last resort, append a run
                        try:
                            p.add_run(str(values[1]))
                        except Exception:
                            pass
            # continue with other labels too (in case paragraph also contains others)
        # Do scoped replacement for each label
        for label, value in targets.items():
            replace_in_paragraph(p, label, value, label_list)

    # Final cleanup: remove any leftover demo placeholders like 'replace ... here'
    for p in search_paras:
        for r in p.runs:
            try:
                if r.font.highlight_color == WD_COLOR_INDEX.YELLOW and r.text.strip().lower().startswith("replace"):
                    r.text = ""
            except Exception:
                pass

# ---------------------------
# Core required functions
# ---------------------------

def create_invoice(form_data: Dict[str, str]) -> Tuple[str, Optional[str]]:
    """Create invoice: generate DOCX, convert to PDF, and save DB record.
    Returns (docx_path, pdf_path or None if conversion failed)
    """
    ensure_dirs()
    conn = get_conn()
    try:
        qno = next_quotation_no(conn)
        form_data = dict(form_data)
        form_data["quotation_no"] = qno

        # Prepare values order for highlighted replacements
        # Order assumption based on the template sample provided:
        # [customer_name, location, city, state, pincode, mobile, product, quotation_no, date_of_quotation, validity_date]
        values = [
            form_data.get("customer_name", ""),
            form_data.get("location", ""),
            form_data.get("city", ""),
            form_data.get("state", ""),
            form_data.get("pincode", ""),
            form_data.get("mobile", ""),
            form_data.get("product", ""),
            form_data.get("quotation_no", ""),
            form_data.get("date_of_quotation", ""),
            form_data.get("validity_date", ""),
        ]

        docx_bytes = generate_docx(values, form_data)
        # Save DOCX
        safe_customer = safe_filename(form_data["customer_name"]) if form_data.get("customer_name") else "customer"
        safe_qno = safe_filename(form_data["quotation_no"]) if form_data.get("quotation_no") else "qno"
        base_name = f"{safe_customer.strip()}+{safe_qno}"
        docx_path = os.path.join(DOCX_DIR, f"{base_name}.docx")
        # Safe overwrite if file exists
        try:
            if os.path.exists(docx_path):
                os.remove(docx_path)
        except Exception:
            pass
        with open(docx_path, "wb") as f:
            f.write(docx_bytes.getvalue())

        # Convert to PDF
        target_pdf = os.path.join(PDF_DIR, f"{base_name}.pdf")
        try:
            if os.path.exists(target_pdf):
                os.remove(target_pdf)
        except Exception:
            pass
        pdf_path = convert_to_pdf(docx_path, target_pdf)

        # Save DB
        save_to_db(conn, {
            **form_data,
            "docx_path": docx_path,
            "pdf_path": pdf_path,
        })
        return docx_path, pdf_path
    finally:
        conn.close()


def generate_docx(values_in_order: List[str], form_data: Dict[str, str]) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found at {TEMPLATE_PATH}")
    doc = Document(TEMPLATE_PATH)

    # Replace values by labels for accuracy
    replace_by_labels(doc, form_data)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def convert_to_pdf(docx_path: str, target_pdf_path: str) -> Optional[str]:
    try:
        # docx2pdf requires Word on Windows. We'll attempt, but tolerate failures.
        docx2pdf_convert(docx_path, target_pdf_path)
        return target_pdf_path if os.path.exists(target_pdf_path) else None
    except Exception:
        return None


def save_to_db(conn: sqlite3.Connection, record: Dict[str, str]) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """
        INSERT INTO invoices (
            quotation_no, product, customer_name, mobile, location, city, state, pincode,
            staff_name, date_of_quotation, validity_date, docx_path, pdf_path, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record.get("quotation_no"),
            record.get("product"),
            record.get("customer_name"),
            record.get("mobile"),
            record.get("location"),
            record.get("city"),
            record.get("state"),
            record.get("pincode"),
            record.get("staff_name"),
            record.get("date_of_quotation"),
            record.get("validity_date"),
            record.get("docx_path"),
            record.get("pdf_path"),
            now,
            now,
        ),
    )
    conn.commit()


def load_invoices() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query(
            "SELECT id, customer_name, product, date_of_quotation, quotation_no, docx_path, pdf_path FROM invoices ORDER BY id DESC",
            conn,
        )
    finally:
        conn.close()
    return df


def delete_invoice(inv_id: int) -> None:
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT docx_path, pdf_path FROM invoices WHERE id = ?", (inv_id,))
        row = cur.fetchone()
        if row:
            docx_path, pdf_path = row
            try:
                if docx_path and os.path.exists(docx_path):
                    os.remove(docx_path)
            except Exception:
                pass
            try:
                if pdf_path and os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except Exception:
                pass
        conn.execute("DELETE FROM invoices WHERE id = ?", (inv_id,))
        conn.commit()
    finally:
        conn.close()


def edit_invoice(inv_id: int, form_data: Dict[str, str]) -> Tuple[str, Optional[str]]:
    """Update record and regenerate files. Returns (docx_path, pdf_path)."""
    ensure_dirs()
    conn = get_conn()
    try:
        # Fetch existing quotation_no to keep it stable
        cur = conn.cursor()
        cur.execute("SELECT quotation_no FROM invoices WHERE id = ?", (inv_id,))
        row = cur.fetchone()
        if not row:
            raise ValueError("Invoice not found")
        quotation_no = row[0]
        form_data = dict(form_data)
        form_data["quotation_no"] = quotation_no

        values = [
            form_data.get("customer_name", ""),
            form_data.get("location", ""),
            form_data.get("city", ""),
            form_data.get("state", ""),
            form_data.get("pincode", ""),
            form_data.get("mobile", ""),
            form_data.get("product", ""),
            form_data.get("quotation_no", ""),
            form_data.get("date_of_quotation", ""),
            form_data.get("validity_date", ""),
        ]
        docx_bytes = generate_docx(values, form_data)

        safe_customer = safe_filename(form_data["customer_name"]) if form_data.get("customer_name") else "customer"
        safe_qno = safe_filename(form_data["quotation_no"]) if form_data.get("quotation_no") else "qno"
        base_name = f"{safe_customer.strip()}+{safe_qno}"
        docx_path = os.path.join(DOCX_DIR, f"{base_name}.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes.getvalue())

        pdf_path = convert_to_pdf(docx_path, os.path.join(PDF_DIR, f"{base_name}.pdf"))

        now = datetime.now().isoformat(timespec="seconds")
        conn.execute(
            """
            UPDATE invoices SET
                product=?, customer_name=?, mobile=?, location=?, city=?, state=?, pincode=?, staff_name=?,
                date_of_quotation=?, validity_date=?, docx_path=?, pdf_path=?, updated_at=?
            WHERE id=?
            """,
            (
                form_data.get("product"),
                form_data.get("customer_name"),
                form_data.get("mobile"),
                form_data.get("location"),
                form_data.get("city"),
                form_data.get("state"),
                form_data.get("pincode"),
                form_data.get("staff_name"),
                form_data.get("date_of_quotation"),
                form_data.get("validity_date"),
                docx_path,
                pdf_path,
                now,
                inv_id,
            ),
        )
        conn.commit()
        return docx_path, pdf_path
    finally:
        conn.close()

# ---------------------------
# Streamlit UI
# ---------------------------

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    ensure_dirs()
    _ = get_conn()  # ensure DB exists

    tabs = st.tabs(["Dashboard", "Create New Invoice", "Search Invoice"])

    with tabs[0]:
        st.info("Dashboard is coming soon.")

    with tabs[1]:
        render_create_form()

    with tabs[2]:
        render_search_tab()


# ---------------------------
# UI helpers
# ---------------------------

def render_create_form(prefill: Optional[Dict[str, str]] = None, edit_id: Optional[int] = None):
    conn = get_conn()
    try:
        qno_preview = next_quotation_no(conn) if edit_id is None else None
    finally:
        conn.close()

    st.subheader("Create New Invoice" if edit_id is None else "Edit Invoice")

    # Date of Quotation & Validity outside the form so validity updates immediately when date changes
    date_default = datetime.now().date()
    if prefill and prefill.get("date_of_quotation"):
        try:
            date_default = datetime.strptime(prefill["date_of_quotation"], "%Y-%m-%d").date()
        except Exception:
            pass
    date_of_quotation = st.date_input(
        "Date of Quotation",
        value=date_default,
        key=f"doq_{edit_id or 'new'}",
    )
    validity_date = date_of_quotation + timedelta(days=30)

    with st.form(key=f"invoice_form_{edit_id or 'new'}"):
        product = st.selectbox("Product & Service", options=PRODUCT_OPTIONS, index=(PRODUCT_OPTIONS.index(prefill.get("product")) if prefill and prefill.get("product") in PRODUCT_OPTIONS else 0))
        customer_name = st.text_input("Customer Name", value=(prefill.get("customer_name") if prefill else ""))
        mobile = st.text_input("Mobile Number", value=(prefill.get("mobile") if prefill else ""))
        location = st.text_input("Location", value=(prefill.get("location") if prefill else ""))
        city = st.text_input("City", value=(prefill.get("city") if prefill else ""))
        state = st.text_input("State", value=(prefill.get("state") if prefill else ""))
        pincode = st.text_input("Pincode", value=(prefill.get("pincode") if prefill else ""))
        staff_name = st.text_input("Staff Name (kept only in DB)", value=(prefill.get("staff_name") if prefill else ""))
        # Show computed values inside the form (read-only)
        st.text_input("Date of Quotation", value=date_of_quotation.isoformat(), disabled=True)
        st.text_input("Validity (auto)", value=validity_date.isoformat(), disabled=True)

        if edit_id is None and qno_preview:
            st.text_input("Quotation No (auto)", value=qno_preview, disabled=True)
        elif edit_id is not None and prefill and prefill.get("quotation_no"):
            st.text_input("Quotation No", value=prefill.get("quotation_no"), disabled=True)

        submit_label = "Update Invoice" if edit_id is not None else "Create Invoice"
        submitted = st.form_submit_button(submit_label)

    if submitted:
        data = {
            "product": product,
            "customer_name": customer_name.strip(),
            "mobile": str(mobile).strip(),
            "location": location.strip(),
            "city": city.strip(),
            "state": state.strip(),
            "pincode": str(pincode).strip(),
            "staff_name": staff_name.strip(),
            "date_of_quotation": date_of_quotation.isoformat(),
            "validity_date": validity_date.isoformat(),
        }
        try:
            if edit_id is None:
                docx_path, pdf_path = create_invoice(data)
                st.success("Invoice created successfully.")
                st.toast(f"Saved invoice {qno_preview}", icon="✅")
            else:
                docx_path, pdf_path = edit_invoice(edit_id, data)
                st.success("Invoice updated successfully.")
                st.toast("Invoice updated", icon="✏️")
            if pdf_path is None:
                st.warning("PDF conversion failed or Word is not available. DOCX was generated.")
            with open(docx_path, "rb") as f:
                st.download_button("Download DOCX", data=f.read(), file_name=os.path.basename(docx_path), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    st.download_button("Download PDF", data=f.read(), file_name=os.path.basename(pdf_path), mime="application/pdf")
        except Exception as e:
            st.error(f"Failed to process invoice: {e}")


def render_search_tab():
    st.subheader("Search Invoice")

    df = load_invoices()
    if df.empty:
        st.info("No invoices yet.")
        return

    # Filters
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        f_name = st.text_input("Filter by Customer Name")
    with c2:
        f_qno = st.text_input("Filter by Quotation No")
    with c3:
        f_product = st.selectbox("Filter by Product", options=["All"] + PRODUCT_OPTIONS, index=0)

    mask = pd.Series([True] * len(df))
    if f_name:
        mask &= df["customer_name"].str.contains(f_name, case=False, na=False)
    if f_qno:
        mask &= df["quotation_no"].str.contains(f_qno, case=False, na=False)
    if f_product != "All":
        mask &= df["product"] == f_product

    filtered = df[mask].reset_index(drop=True)

    st.dataframe(filtered[["customer_name", "product", "date_of_quotation", "quotation_no"]], use_container_width=True)

    st.markdown("### Actions")
    for _, row in filtered.iterrows():
        with st.container():
            c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 1, 1, 1])
            with c1:
                st.write(f"Customer: **{row['customer_name']}**")
                st.caption(f"Quotation: {row['quotation_no']}")
            with c2:
                st.write(row["product"])
            with c3:
                st.write(row["date_of_quotation"]) 
            with c4:
                if row["docx_path"] and os.path.exists(row["docx_path"]):
                    with open(row["docx_path"], "rb") as f:
                        st.download_button("DOCX", f.read(), file_name=os.path.basename(row["docx_path"]), key=f"docx_{row['id']}")
                else:
                    st.button("DOCX", disabled=True, key=f"docx_{row['id']}")
            with c5:
                if row["pdf_path"] and os.path.exists(row["pdf_path"]):
                    with open(row["pdf_path"], "rb") as f:
                        st.download_button("PDF", f.read(), file_name=os.path.basename(row["pdf_path"]), key=f"pdf_{row['id']}")
                else:
                    st.button("PDF", disabled=True, key=f"pdf_{row['id']}")
            with c6:
                st.write("")

            c7, c8 = st.columns([1, 1])
            with c7:
                if st.button("Edit", key=f"edit_{row['id']}"):
                    prefill = fetch_full_record(row["id"]) or {}
                    render_create_form(prefill=prefill, edit_id=int(row["id"]))
                    st.stop()
            with c8:
                if st.button("Delete", key=f"del_{row['id']}"):
                    delete_invoice(int(row["id"]))
                    st.success("Deleted.")
                    st.experimental_rerun()


def fetch_full_record(inv_id: int) -> Optional[Dict[str, str]]:
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT id, quotation_no, product, customer_name, mobile, location, city, state, pincode, staff_name, date_of_quotation, validity_date FROM invoices WHERE id=?",
            (inv_id,),
        )
        r = cur.fetchone()
        if not r:
            return None
        return {
            "id": r[0],
            "quotation_no": r[1],
            "product": r[2],
            "customer_name": r[3],
            "mobile": r[4],
            "location": r[5],
            "city": r[6],
            "state": r[7],
            "pincode": r[8],
            "staff_name": r[9],
            "date_of_quotation": r[10],
            "validity_date": r[11],
        }
    finally:
        conn.close()


if __name__ == "__main__":
    main()
