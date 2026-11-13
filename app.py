"""InvoiceFlow Pro v0.4: auditable, conservative extraction of text PDF invoices.

No OCR, no invented totals, only USD amounts confidently identified as USD.
"""
from __future__ import annotations

from datetime import datetime
from io import BytesIO
import re

import pandas as pd
import pymupdf

OUTPUT_COLUMNS = [
    "file", "invoice_number", "invoice_date", "supplier",
    "customer", "total_usd", "status", "review_reason",
]
REQUIRED_FIELDS = ["invoice_number", "invoice_date", "supplier", "customer", "total_usd"]


def _field(lines: list[str], pattern: str) -> str | None:
    """Find an English labelled field, including label and value on adjacent lines."""
    for index, line in enumerate(lines):
        match = re.match(r"^\s*(?:" + pattern + r")\s*[:#-]\s*(.*?)\s*$", line, re.I)
        if not match:
            continue
        value = match.group(1).strip()
        if value:
            return value
        if index + 1 < len(lines):
            candidate = lines[index + 1].strip()
            # Never mistake a different label for a field value.
            if candidate and not re.match(r"^\w[\w ]{1,30}\s*:\s*", candidate):
                return candidate
    return None


def _normalized_date(value: str | None) -> tuple[str | None, str | None]:
    if not value:
        return None, "Missing invoice date"
    value = value.strip()
    # Explicit ISO dates, clear day.month.year or unambiguous English month names.
    formats = ["%Y-%m-%d", "%d.%m.%Y", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y"]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date().isoformat(), None
        except ValueError:
            pass
    if re.search(r"\d[/]\d", value):
        return None, "Ambiguous date format (use YYYY-MM-DD)"
    return None, "Unrecognized invoice date"


def parse_invoice_text(filename: str, text: str) -> dict:
    """Conservatively extract known fields. Flag uncertainty for review."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    invoice_number_text = _field(lines, r"invoice\s*(?:number|no\.?|#)|inv\s*(?:number|no\.?|#)")
    invoice_number = None
    if invoice_number_text:
        match = re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._/-]*", invoice_number_text)
        invoice_number = invoice_number_text if match else None

    date_text = _field(lines, r"invoice\s*date|date\s*of\s*invoice")
    invoice_date, date_error = _normalized_date(date_text)
    supplier = _field(lines, r"supplier|seller|vendor|from")
    if not supplier:
        for i, line in enumerate(lines[:-1]):
            if line.upper() == "INVOICE" and lines[i + 1].strip():
                supplier = lines[i + 1].strip()
                break
    customer = _field(lines, r"bill\s*to|customer|billed\s*to")

    total_text = _field(lines, r"grand\s*total|invoice\s*total|total(?:\s*(?:amount|due))?|amount\s*due|balance\s*due")
    errors = []
    amount = None
    if total_text:
        # Only allow an explicitly marked USD amount; never label EUR or unspecified money USD.
        tokens = re.fullmatch(
            r"(?:(USD|EUR|GBP|CAD)\s*)?(\$)?\s*([\d,]+(?:\.\d{1,2})?)\s*(USD|EUR|GBP|CAD)?",
            total_text.strip(), re.I,
        )
        if tokens:
            pre, dollar, digits, post = tokens.groups()
            currencies = {c.upper() for c in (pre, post) if c}
            if dollar:
                currencies.add("USD")
            currency_field = _field(lines, r"currency")
            if currency_field:
                currencies.add(currency_field.strip().upper())
            if currencies == {"USD"}:
                amount = float(digits.replace(",", ""))
            elif len(currencies) == 0:
                errors.append("Currency not confirmed (USD only)")
            else:
                errors.append("Unsupported or conflicting currency (USD only)")
        else:
            errors.append("Unrecognized total amount")
    else:
        errors.append("Missing total amount")

    if not invoice_number:
        errors.append("Missing or invalid invoice number")
    if date_error:
        errors.append(date_error)
    if not supplier:
        errors.append("Missing supplier")
    if not customer:
        errors.append("Missing customer")

    record = {
        "file": filename,
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "supplier": supplier,
        "customer": customer,
        "total_usd": amount,
        "status": "Needs review" if errors else "Complete",
        "review_reason": "; ".join(errors),
    }
    return record


def pdf_to_text(content: bytes) -> str:
    with pymupdf.open(stream=content, filetype="pdf") as pdf:
        if pdf.page_count == 0:
            raise ValueError("PDF contains no pages")
        return "\n".join(page.get_text("text") for page in pdf)


def _excel_safe(value):
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def make_excel(df: pd.DataFrame) -> bytes:
    """Three well-formatted XLSX sheets. Review records excluded from verified total."""
    from openpyxl.styles import Alignment, Font, PatternFill

    output = BytesIO()
    for_export = df.reindex(columns=OUTPUT_COLUMNS).copy()
    for column in for_export.select_dtypes(include=["object", "string"]).columns:
        for_export[column] = for_export[column].map(_excel_safe)
    review = for_export[for_export["status"] == "Needs review"].copy()
    verified = pd.to_numeric(df.loc[df["status"] == "Complete", "total_usd"], errors="coerce")
    summary = pd.DataFrame({
        "Metric": ["Files processed", "Complete", "Needs review", "Verified total (USD)"],
        "Value": [len(df), int((df["status"] == "Complete").sum()),
                  int((df["status"] == "Needs review").sum()), float(verified.sum())],
    })
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for_export.to_excel(writer, sheet_name="Invoices", index=False)
        review.to_excel(writer, sheet_name="Review Needed", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="17324C")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(vertical="center")
            sheet.row_dimensions[1].height = 28
            for column in sheet.columns:
                width = max(len(str(cell.value or "")) for cell in column)
                sheet.column_dimensions[column[0].column_letter].width = min(max(width + 3, 14), 48)
            if sheet.title in ("Invoices", "Review Needed"):
                for row in sheet.iter_rows(min_row=2):
                    row[5].number_format = '#,##0.00'
                    if row[6].value == "Needs review":
                        for cell in row:
                            cell.fill = PatternFill("solid", fgColor="FFF1D6")
            if sheet.title == "Summary":
                sheet["B5"].number_format = '#,##0.00'
    return output.getvalue()


def detect_potential_duplicates(records: list[dict]) -> list[dict]:
    """Conservative batch-level validation: same supplier + invoice number.

    Mark *all* members of a group as needing review: there is not enough
    evidence to determine which copy (if any) is authoritative.
    """
    from collections import Counter

    def normalized(value) -> str:
        return " ".join(str(value or "").casefold().split())

    keys = [
        (normalized(rec.get("supplier")), normalized(rec.get("invoice_number")))
        for rec in records
    ]
    counts = Counter(key for key in keys if all(key))
    for record, key in zip(records, keys):
        count = counts.get(key, 0)
        if count > 1:
            reason = f"Potential duplicate: same supplier and invoice number ({count} files)"
            existing = record.get("review_reason") or ""
            record["review_reason"] = "; ".join(filter(None, (existing, reason)))
            record["status"] = "Needs review"
    return records


def screen_dataframe(frame: pd.DataFrame) -> pd.DataFrame:
    displayed = frame.copy()
    displayed["total_usd"] = displayed["total_usd"].map(
        lambda x: "Missing value" if pd.isna(x) else f"{x:,.2f}"
    )
    return displayed


def main() -> None:
    import streamlit as st

    st.set_page_config(
        page_title="InvoiceFlow Pro | PDF to Excel",
        page_icon="📑", layout="wide", initial_sidebar_state="expanded",
    )
    st.markdown("""
    <style>
    .block-container {padding-top: 1.35rem; padding-bottom: 2rem;}
    .hero {background:linear-gradient(115deg,#102b46,#214f74); color:white;
      border-radius:15px;padding:24px 28px;margin-bottom:19px;}
    .hero h1 {color:white;margin:0;font-size:2rem;}
    .hero p {color:#d6e6f5;margin:7px 0 0;}
    .small-note {font-size:.88rem;color:#64748b;}
    [data-testid="stMetric"] {background:#f5f8fc;border:1px solid #e0e8f2;
      padding:14px 16px;border-radius:12px;}
    </style>
    """, unsafe_allow_html=True)
    st.markdown("""<div class="hero"><h1>InvoiceFlow Pro</h1>
    <p>Turn text-based PDF invoices into organized Excel reports — with transparent exception checks.</p>
    </div>""", unsafe_allow_html=True)

    with st.sidebar:
        st.header("Workflow")
        st.write("1. Upload invoices\n\n2. Check extracted fields\n\n3. Review exceptions\n\n4. Export to Excel")
        st.divider()
        st.caption("Version 4 · Demonstration")
        st.warning("Text-based English PDFs only. USD totals only; scanned PDFs need OCR.")
        st.caption("No extracted number is independently authenticated. Always review amounts before financial use.")

    files = st.file_uploader(
        "Upload PDF invoices", type=["pdf"], accept_multiple_files=True,
        help="Use synthetic sample files for testing. Do not upload confidential files to a public deployment.",
    )
    if not files:
        st.info("Try the included artificial PDF invoices in the samples folder.")
        st.markdown("""**What this app demonstrates:** structured extraction, different label styles,
        currency checks, missing-field detection, duplicate detection and formatted Excel exports.""")
        return

    records, raw_texts = [], []
    for uploaded in files:
        try:
            raw = pdf_to_text(uploaded.getvalue())
            if raw.strip():
                record = parse_invoice_text(uploaded.name, raw)
            else:
                record = dict.fromkeys(OUTPUT_COLUMNS, None)
                record.update(
                    file=uploaded.name, status="Needs review",
                    review_reason="No selectable text (OCR required)",
                )
            raw_texts.append((uploaded.name, raw or "No selectable text detected."))
        except Exception:
            record = dict.fromkeys(OUTPUT_COLUMNS, None)
            record.update(
                file=uploaded.name, status="Needs review", review_reason="Unreadable PDF",
            )
            raw_texts.append((uploaded.name, "Could not extract text from this PDF."))
        records.append(record)

    records = detect_potential_duplicates(records)
    df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    completed = int((df["status"] == "Complete").sum())
    needs_review = int((df["status"] == "Needs review").sum())
    parsed_total = pd.to_numeric(
        df.loc[df["status"] == "Complete", "total_usd"], errors="coerce"
    ).sum()
    duplicates = int(df["review_reason"].fillna("").str.contains("Potential duplicate").sum())

    st.success(f"Processed {len(df)} PDF file(s).")
    a, b, c, d = st.columns(4)
    a.metric("PDFs", len(df))
    b.metric("Complete", completed)
    c.metric("Needs review", needs_review)
    d.metric("Validated-format USD total", f"${parsed_total:,.2f}")
    st.caption("Validated-format means automated checks passed; it does not mean independent invoice verification.")

    if needs_review:
        st.warning(f"{needs_review} record(s) require review. {duplicates} record(s) were flagged as potential duplicates.")

    tab_data, tab_review, tab_chart, tab_text = st.tabs(
        ["All invoices", "Review needed", "Overview", "Original PDF text"]
    )
    with tab_data:
        statuses = ["All", "Complete", "Needs review"]
        selection = st.selectbox("Filter by status", statuses, index=0)
        search = st.text_input("Search file, supplier, customer or invoice number", "")
        filtered = df if selection == "All" else df[df["status"] == selection]
        if search.strip():
            fields = ["file", "invoice_number", "supplier", "customer"]
            contains = filtered[fields].fillna("").astype(str).apply(
                lambda col: col.str.contains(search.strip(), case=False, regex=False)
            ).any(axis=1)
            filtered = filtered.loc[contains]
        st.dataframe(screen_dataframe(filtered), use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(filtered)} of {len(df)} records.")
    with tab_review:
        review = df.loc[df["status"] == "Needs review"].copy()
        st.dataframe(screen_dataframe(review), use_container_width=True, hide_index=True)
        if review.empty:
            st.success("No records currently require review.")
        else:
            st.info("Records are flagged for manual checking; no missing amounts are guessed.")
            safe_csv = review.copy()
            for column in safe_csv.select_dtypes(include=["object", "string"]).columns:
                safe_csv[column] = safe_csv[column].map(_excel_safe)
            st.download_button(
                "Download review list (CSV)",
                data=safe_csv.to_csv(index=False).encode("utf-8-sig"),
                file_name="InvoiceFlow_V4_Review_Needed.csv", mime="text/csv",
            )
    with tab_chart:
        st.subheader("Processing overview")
        status_df = pd.DataFrame({
            "Result": ["Complete", "Needs review"], "Invoices": [completed, needs_review],
        }).set_index("Result")
        st.bar_chart(status_df)
        st.caption("Only 'Complete' USD invoices contribute to the automatically checked total.")
    with tab_text:
        st.caption("Reference text from the original documents for cross-checking.")
        for filename, raw in raw_texts:
            with st.expander(filename):
                st.code(raw[:12000], language=None)

    st.divider()
    st.download_button(
        "⬇ Download Excel Report", data=make_excel(df),
        file_name="InvoiceFlow_Pro_V4_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption("Portfolio demo. Review extracted data before financial use; never upload confidential invoices to public demos.")


if __name__ == "__main__":
    main()
