"""
PO PDF extractor for Utopia Industries system-generated purchase orders.

Extracts:
  - PO header (PO number, dates, department, delivery location, supplier/buyer)
  - Line items (variable 1-N SKUs with delivery date, qty, unit cost, totals)

Uses pdfplumber for text extraction from computer-generated PDFs.
"""
import re
from datetime import datetime


MONTH_REGEX = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
LINE_DATE_REGEX = r'(?:' + MONTH_REGEX + r'\s+\d{1,2},\s+\d{4}|\d{4}-\d{2}-\d{2})'
_SKU_TOKEN_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9._/-]{3,}$')


def _looks_like_date_token(value):
    token = str(value or '').strip()
    if not token:
        return False

    # 2026-03-12, 12/03/2026, 12-03-2026
    if re.match(r'^\d{4}-\d{2}-\d{2}$', token):
        return True
    if re.match(r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$', token):
        return True

    # Mar 12, 2026
    if re.match(r'^(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}$', token, re.IGNORECASE):
        return True

    return False


def _parse_date(text):
    """Parse 'Apr 10, 2026' or 'Apr 10, 2026 01:37 AM' -> ISO date string."""
    if not text:
        return None
    text = text.strip()
    for fmt in ('%b %d, %Y %I:%M %p', '%b %d, %Y'):
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    # Try direct ISO
    try:
        return datetime.strptime(text[:10], '%Y-%m-%d').strftime('%Y-%m-%d')
    except ValueError:
        return None


def _clean_amount(text):
    """'Rs 2,400.00' or '2,400.00' -> float."""
    if not text:
        return None
    cleaned = re.sub(r'[Rs,\s]', '', text)
    try:
        return float(cleaned)
    except ValueError:
        return None


def _clean_location(text):
    if not text:
        return None
    cleaned = str(text).strip().strip(':').strip()
    # Remove accidental trailing labels captured from dense PDF text blocks.
    cleaned = re.split(r'\s+(?:Buyer\s+Details|SUPPLIER\s+DETAILS|BUYER\s+DETAILS)\b', cleaned, maxsplit=1)[0].strip()
    return cleaned or None


def _looks_like_sku_token(value):
    if not value:
        return False
    token = str(value).strip()
    if _looks_like_date_token(token):
        return False
    if not _SKU_TOKEN_RE.match(token):
        return False
    upper = token.upper()
    if upper.startswith(('PO-', 'PR-', 'SITE-', 'NTN', 'STRN', 'RS')):
        return False
    if upper in {'PIECE', 'PCS', 'UNIT', 'SET', 'BOX', 'ROLL', 'PACK', 'KG', 'METER', 'YARD'}:
        return False
    return True


def _build_sku_jobname_map(text, table_blobs=None):
    """Best-effort map for layouts where job name is on the line before SKU."""
    mapping = {}

    lines = [ln.strip() for ln in str(text).splitlines() if ln and ln.strip()]
    for idx, line in enumerate(lines):
        if not _looks_like_sku_token(line):
            continue
        if idx == 0:
            continue
        candidate_name = lines[idx - 1].strip()
        if not candidate_name:
            continue
        # Skip obvious headers and structural rows.
        if re.search(r'(PURCHASE ORDER|DELIVERY DATE|GRAND TOTAL|SUPPLIER DETAILS|BUYER DETAILS|DESCRIPTION)', candidate_name, re.IGNORECASE):
            continue
        if _looks_like_sku_token(candidate_name):
            continue
        mapping[line.upper()] = candidate_name

    if table_blobs:
        blobs = [str(b).strip() for b in table_blobs if str(b).strip()]
        for idx, blob in enumerate(blobs):
            # If a row starts with SKU token and previous row is human-readable name, map it.
            first_token = blob.split()[0] if blob.split() else ''
            if not _looks_like_sku_token(first_token):
                continue
            if idx == 0:
                continue
            prev_blob = blobs[idx - 1]
            if re.search(r'(DELIVERY DATE|QUANTITY|UNIT COST|SUBTOTAL|GST|NET TOTAL)', prev_blob, re.IGNORECASE):
                continue
            if _looks_like_sku_token(prev_blob):
                continue
            mapping[first_token.upper()] = prev_blob

    return mapping


def extract_po_from_pdf(file_obj):
    """
    Extract PO data from a file-like object (Django uploaded file).
    Returns dict with keys: po_number, po_date, approval_date, department,
    delivery_location, supplier_name, buyer_name, grand_total, items[].
    Raises ValueError with a message if extraction fails.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ValueError("pdfplumber is not installed. Run: pip install pdfplumber")

    # Read all pages as combined text and table rows
    full_text = ''
    table_blobs = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text(x_tolerance=3, y_tolerance=3) or ''
            full_text += page_text + '\n'

            # Keep table rows as flattened strings for fallback item parsing.
            for table in (page.extract_tables() or []):
                for row in table or []:
                    parts = [str(col).strip() for col in (row or []) if str(col).strip()]
                    if parts:
                        table_blobs.append(' '.join(parts))

    if not full_text.strip():
        raise ValueError("Could not extract any text from the PDF. Please check the file.")

    result = _parse_po_text(full_text, table_blobs)
    return result


def _parse_po_text(text, table_blobs=None):
    """Parse raw extracted text into structured PO dict."""

    # ── PO Number ──────────────────────────────────────────────────────────────
    po_number = None
    m = re.search(r'PURCHASE ORDER\s+(PO-[\w-]+)', text)
    if m:
        po_number = m.group(1).strip()

    if not po_number:
        # Fallback: look for standalone PO-XX-XXXX-XXXXXX pattern
        m = re.search(r'\b(PO-\d{2}-\d{4}-\d+)\b', text)
        if m:
            po_number = m.group(1)

    # ── Dates ──────────────────────────────────────────────────────────────────
    po_date = None
    m = re.search(r'Dated\s+((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}(?:\s+\d{1,2}:\d{2}\s+[AP]M)?)', text)
    if m:
        po_date = _parse_date(m.group(1))

    approval_date = None
    m = re.search(r'Approval Date\s+((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4}(?:\s+\d{1,2}:\d{2}\s+[AP]M)?)', text)
    if m:
        approval_date = _parse_date(m.group(1))

    # ── Department ─────────────────────────────────────────────────────────────
    department = None
    delivery_location = None
    m = re.search(r'Department\s*/\s*Broker\s+(.+?)(?=\n|Delivery Location)', text)
    if m:
        raw = m.group(1).strip()
        # Format: "DEPT / null" or "DEPT / BROKER"
        parts = raw.split('/')
        department = parts[0].strip()

    # Try more targeted parse when previous pattern is noisy
    m = re.search(r'Department\s*/\s*Broker\s+([\w &]+(?:\s*/\s*[\w &]+)?)', text)
    if m:
        raw = m.group(1).strip()
        parts = re.split(r'\s*/\s*', raw, maxsplit=1)
        department = parts[0].strip()

    # Delivery Location can appear as:
    # 1) "Delivery Location SITE-2"
    # 2) "Delivery Location: SITE-2"
    # 3) "Delivery Location" on one line and value on next line
    delivery_patterns = [
        r'Delivery\s*Location\s*[:\-]?\s*([^\n]+)',
        r'Delivery\s*Location\s*[:\-]?\s*\n\s*([^\n]+)',
        r'Location\s*[:\-]\s*([^\n]+)',
    ]
    for pattern in delivery_patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if not m:
            continue
        candidate = _clean_location(m.group(1))
        if candidate:
            delivery_location = candidate
            break

    # ── Supplier / Buyer ───────────────────────────────────────────────────────
    supplier_name = None
    buyer_name = None
    # Supplier block: "Name UTOPIA PRINTING & PACKAGING"  (appears before NTN/STRN)
    m = re.search(r'SUPPLIER DETAILS.*?Name\s+([^\n]+)', text, re.DOTALL)
    if m:
        supplier_name = m.group(1).strip()
    m = re.search(r'BUYER DETAILS.*?Name\s+([^\n]+)', text, re.DOTALL)
    if m:
        buyer_name = m.group(1).strip()

    # ── Grand Total ────────────────────────────────────────────────────────────
    grand_total = None
    m = re.search(r'GRAND TOTAL\s+Rs\s+([\d,]+\.?\d*)', text)
    if m:
        grand_total = _clean_amount(m.group(1))

    # ── Line Items ─────────────────────────────────────────────────────────────
    # Try strict layout first, then flexible regex, then extracted table rows.
    sku_jobname_map = _build_sku_jobname_map(text, table_blobs)

    items = _extract_items_strict(text, sku_jobname_map)
    if not items:
        items = _extract_items_flexible(text, sku_jobname_map)
    if not items and table_blobs:
        items = _extract_items_from_table_blobs(table_blobs, sku_jobname_map)
    if not items:
        items = _extract_items_from_text_windows(text, sku_jobname_map)

    if not items:
        raise ValueError(
            f"Could not detect any line items in the PO PDF. "
            f"Extracted text preview: {text[:400]!r}"
        )

    return {
        'po_number': po_number,
        'po_date': po_date,
        'approval_date': approval_date,
        'department': department,
        'delivery_location': delivery_location,
        'supplier_name': supplier_name,
        'buyer_name': buyer_name,
        'grand_total': grand_total,
        'items': items,
    }


def _append_item(items, sku, delivery_date_raw, qty_raw, unit_raw, unit_cost_raw, subtotal_raw, gst_raw, net_total_raw, sku_jobname_map=None):
    sku_value = (sku or '').strip()
    if not sku_value:
        return
    if _looks_like_date_token(sku_value):
        return
    if sku_value.upper().startswith(('SUBTOTAL', 'GRAND', 'TOTAL', 'DESCRIPTION')):
        return

    try:
        quantity = float(str(qty_raw).replace(',', '').strip())
    except ValueError:
        quantity = None

    item = {
        'line_no': len(items) + 1,
        'sku': sku_value,
        'job_name': (sku_jobname_map or {}).get(sku_value.upper(), sku_value),
        'delivery_date': _parse_date(delivery_date_raw),
        'quantity': quantity,
        'unit': (unit_raw or '').upper().rstrip('.'),
        'unit_cost': _clean_amount(unit_cost_raw),
        'subtotal': _clean_amount(subtotal_raw),
        'gst': _clean_amount(gst_raw),
        'net_total': _clean_amount(net_total_raw),
    }

    # Must have at least sku + date + qty to be considered a valid line item.
    if not item['delivery_date'] or item['quantity'] is None:
        return

    items.append(item)


def _extract_items_strict(text, sku_jobname_map=None):
    amount = r'(?:Rs\s*)?[\d,]+\.?\d*'
    item_pattern = re.compile(
        r'^([A-Za-z0-9][\w\-./]+)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,]+\.?\d*)\s+(PIECE\.?|PCS\.?|UNIT\.?|SET\.?|BOX\.?|ROLL\.?|PACK\.?|KG\.?|METER\.?|YARD\.?)\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')',
        re.MULTILINE | re.IGNORECASE,
    )

    items = []
    for m in item_pattern.finditer(text):
        _append_item(
            items,
            m.group(1),
            m.group(2),
            m.group(3),
            m.group(4),
            m.group(5),
            m.group(6),
            m.group(7),
            m.group(8),
            sku_jobname_map,
        )
    return items


def _extract_items_flexible(text, sku_jobname_map=None):
    # Handles wrapped or slightly shifted lines and optional currency prefixes.
    normalized = '\n'.join(' '.join(line.split()) for line in text.splitlines() if line.strip())
    amount = r'(?:Rs\s*)?[\d,]+\.?\d*'
    pattern = re.compile(
        r'(.+?)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,]+\.?\d*)\s+'
        r'(PIECE\.?|PCS\.?|UNIT\.?|SET\.?|BOX\.?|ROLL\.?|PACK\.?|KG\.?|METER\.?|YARD\.?)\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')',
        re.IGNORECASE,
    )

    items = []
    for m in pattern.finditer(normalized):
        sku = (m.group(1) or '').strip()
        # Trim any accidental leading labels before actual SKU token.
        sku = re.sub(r'^(ITEM\s*CODE|ITEM|DESCRIPTION)\s*[:\-]?\s*', '', sku, flags=re.IGNORECASE)
        _append_item(
            items,
            sku,
            m.group(2),
            m.group(3),
            m.group(4),
            m.group(5),
            m.group(6),
            m.group(7),
            m.group(8),
            sku_jobname_map,
        )
    return items


def _extract_items_from_table_blobs(table_blobs, sku_jobname_map=None):
    items = []
    amount = r'(?:Rs\s*)?[\d,]+\.?\d*'
    pattern = re.compile(
        r'(.+?)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,]+\.?\d*)\s+'
        r'(PIECE\.?|PCS\.?|UNIT\.?|SET\.?|BOX\.?|ROLL\.?|PACK\.?|KG\.?|METER\.?|YARD\.?)\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')\s+'
        r'(' + amount + r')',
        re.IGNORECASE,
    )

    for blob in table_blobs:
        compact = ' '.join(str(blob).split())
        m = pattern.search(compact)
        if not m:
            continue
        _append_item(
            items,
            m.group(1),
            m.group(2),
            m.group(3),
            m.group(4),
            m.group(5),
            m.group(6),
            m.group(7),
            m.group(8),
            sku_jobname_map,
        )

    return items


def _extract_items_from_text_windows(text, sku_jobname_map=None):
    """
    Last-resort parser for PDFs where item rows are split across nearby lines.
    It scans around date-bearing lines and assembles SKU/qty/unit/amount values.
    """
    lines = [' '.join(line.split()) for line in str(text).splitlines() if line and line.strip()]
    if not lines:
        return []

    date_re = re.compile(r'(' + LINE_DATE_REGEX + r')', re.IGNORECASE)
    unit_re = re.compile(r'\b(PIECE\.?|PCS\.?|UNIT\.?|SET\.?|BOX\.?|ROLL\.?|PACK\.?|KG\.?|METER\.?|YARD\.?)\b', re.IGNORECASE)
    amount_re = re.compile(r'(?:Rs\s*)?([\d,]+\.?\d*)', re.IGNORECASE)
    qty_unit_re = re.compile(
        r'([\d,]+\.?\d*)\s*(PIECE\.?|PCS\.?|UNIT\.?|SET\.?|BOX\.?|ROLL\.?|PACK\.?|KG\.?|METER\.?|YARD\.?)',
        re.IGNORECASE,
    )

    blocked_headers = re.compile(r'\b(PURCHASE ORDER|Dated|Quotation Date|Approval Date|Department|Delivery Location|SUPPLIER DETAILS|BUYER DETAILS|GRAND TOTAL|SUBTOTAL|GST)\b', re.IGNORECASE)

    items = []
    seen = set()
    for idx, line in enumerate(lines):
        m_date = date_re.search(line)
        if not m_date:
            continue

        if blocked_headers.search(line) and not unit_re.search(line):
            continue

        delivery_date = m_date.group(1)
        before = line[: m_date.start()].strip()
        after = line[m_date.end() :].strip()

        window_text = line
        if idx + 1 < len(lines):
            window_text = f"{window_text} {lines[idx + 1]}"

        sku = None
        if before:
            for token in reversed(before.split()):
                if _looks_like_sku_token(token):
                    sku = token
                    break

        if not sku and idx > 0:
            prev = lines[idx - 1].strip()
            if _looks_like_sku_token(prev):
                sku = prev
            else:
                for token in reversed(prev.split()):
                    if _looks_like_sku_token(token):
                        sku = token
                        break

        if not sku:
            continue

        qty_raw = None
        unit_raw = ''
        qty_unit = qty_unit_re.search(after) or qty_unit_re.search(window_text)
        if qty_unit:
            qty_raw = qty_unit.group(1)
            unit_raw = qty_unit.group(2)
        else:
            qty_match = re.search(r'\b([\d,]+\.?\d*)\b', after) or re.search(r'\b([\d,]+\.?\d*)\b', window_text)
            if qty_match:
                qty_raw = qty_match.group(1)

        if qty_raw is None:
            continue

        amount_source = f"{after} {lines[idx + 1]}" if idx + 1 < len(lines) else after
        amount_matches = amount_re.findall(amount_source)
        if len(amount_matches) < 4:
            amount_matches = amount_re.findall(window_text)

        unit_cost_raw = amount_matches[0] if len(amount_matches) > 0 else None
        subtotal_raw = amount_matches[1] if len(amount_matches) > 1 else None
        gst_raw = amount_matches[2] if len(amount_matches) > 2 else None
        net_total_raw = amount_matches[3] if len(amount_matches) > 3 else None

        key = (sku.upper(), delivery_date, str(qty_raw))
        if key in seen:
            continue

        before_count = len(items)
        _append_item(
            items,
            sku,
            delivery_date,
            qty_raw,
            unit_raw,
            unit_cost_raw,
            subtotal_raw,
            gst_raw,
            net_total_raw,
            sku_jobname_map,
        )
        if len(items) > before_count:
            seen.add(key)

    return items
