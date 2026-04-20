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
# Unit pattern — extend here to support new units across all parsers
UNIT_PATTERN = r'(?:PIECE|PCS|UNIT|SET|BOX|ROLL|PACK|KG|METER|YARD|NOS|EA|EACH|RL|MTR)\.?'
_SKU_TOKEN_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9._/-]{2,}$')
_SKU_BLOCK_WORDS = {
    'DATED',
    'GENERATED',
    'QUOTATION',
    'APPROVAL',
    'PURCHASE',
    'ORDER',
    'DEPARTMENT',
    'BROKER',
    'DELIVERY',
    'LOCATION',
    'SUPPLIER',
    'BUYER',
    'DETAILS',
    'REFERENCE',
    'INCOTERM',
    'NAME',
    'THIS',
}


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
    """Parse 'Apr 10, 2026' or 'Apr 10, 2026 01:37 AM' or '2026-04-10' -> ISO date string."""
    if not text:
        return None
    text = str(text).strip()
    # Strip watermark chars like "D\n" prefix from PDF page-border text
    text = re.sub(r'^[A-Z]\n', '', text).strip()
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
    """Convert amount strings to float. Handles Rs/€ and both European and US formats."""
    if not text:
        return None
    # Strip currency symbols and whitespace
    cleaned = re.sub(r'[Rs€\$\s]', '', str(text)).strip()
    if not cleaned:
        return None
    # European format: dot as thousands separator, comma as decimal  e.g. "3.000,00"
    if re.search(r',\d{2}$', cleaned):
        cleaned = cleaned.replace('.', '').replace(',', '.')
    else:
        # Standard format: comma as thousands separator  e.g. "3,000.00"
        cleaned = cleaned.replace(',', '')
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


def _clean_party_name(text):
    if not text:
        return None
    cleaned = str(text).strip().strip(':').strip()
    # Stop when structured labels start bleeding into the captured name.
    cleaned = re.split(
        r'\b(?:NTN|STRN|Address|Phone|Contact|Email|SUPPLIER\s+DETAILS|BUYER\s+DETAILS|Delivery\s+Date|Description|GRAND\s+TOTAL)\b',
        cleaned,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()
    return cleaned or None


def _extract_party_name_from_section(text, section_label, end_markers):
    start_match = re.search(section_label, text, flags=re.IGNORECASE)
    if not start_match:
        return None

    start_idx = start_match.end()
    end_idx = len(text)
    tail = text[start_idx:]
    for marker in end_markers:
        marker_match = re.search(marker, tail, flags=re.IGNORECASE)
        if marker_match:
            end_idx = min(end_idx, start_idx + marker_match.start())

    block = text[start_idx:end_idx]
    if not block.strip():
        return None

    name_match = re.search(r'\bName\b\s*[:\-]?\s*([^\n]+)', block, flags=re.IGNORECASE)
    if name_match:
        return _clean_party_name(name_match.group(1))

    # Fallback: use the first meaningful line in the section.
    for line in (ln.strip() for ln in block.splitlines() if ln and ln.strip()):
        candidate = _clean_party_name(line)
        if candidate:
            return candidate
    return None


def _looks_like_sku_token(value):
    if not value:
        return False
    token = str(value).strip()
    if _looks_like_date_token(token):
        return False
    # Allow alphabet-only operational SKUs (some Utopia SKUs are pure letters).
    # Keep short words blocked through length/block-word checks below.
    if token.isalpha() and len(token) < 6:
        return False
    if not _SKU_TOKEN_RE.match(token):
        return False
    upper = token.upper()
    if upper in _SKU_BLOCK_WORDS:
        return False
    if upper.startswith(('PO-', 'PR-', 'SITE-', 'NTN', 'STRN', 'RS')):
        return False
    if upper in {'PIECE', 'PCS', 'UNIT', 'SET', 'BOX', 'ROLL', 'PACK', 'KG', 'METER', 'YARD'}:
        return False
    return True


def _extract_best_sku_token(raw_value):
    """Extract a single token that looks like a real SKU from noisy text."""
    text = str(raw_value or '').strip()
    if not text:
        return None

    # If already a clean token, use directly.
    if _looks_like_sku_token(text):
        return text

    tokens = re.findall(r'[A-Za-z0-9._/-]+', text)
    candidates = [tok for tok in tokens if _looks_like_sku_token(tok)]
    if not candidates:
        return None

    # Avoid dimension fragments like 95x45 when a real SKU token exists.
    non_dimension = [tok for tok in candidates if not re.fullmatch(r'\d{2,4}x\d{2,4}', tok, re.IGNORECASE)]
    pool = non_dimension or candidates

    # Prefer richer tokens (longer, non-numeric-leading) over short numeric fragments.
    pool = sorted(
        pool,
        key=lambda tok: (
            tok[0].isalpha(),
            len(tok),
            any(ch.isdigit() for ch in tok),
        ),
        reverse=True,
    )
    return pool[0]


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
    table_rows = []
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
                        table_rows.append(parts)

    if not full_text.strip():
        raise ValueError("Could not extract any text from the PDF. Please check the file.")

    result = _parse_po_text(full_text, table_blobs, table_rows)
    return result


def _parse_po_text(text, table_blobs=None, table_rows=None):
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

    # Side-by-side layout in flattened PDF text:
    # "SUPPLIER DETAILS BUYER DETAILS Name <supplier> Name <buyer> NTN..."
    paired_name_match = re.search(
        r'SUPPLIER\s+DETAILS\s+BUYER\s+DETAILS\s+Name\s+(.+?)\s+Name\s+(.+?)\s+(?:NTN|Contact\s+Person|Address|#\s*SKU)',
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if paired_name_match:
        supplier_name = _clean_party_name(paired_name_match.group(1))
        buyer_name = _clean_party_name(paired_name_match.group(2))

    if not supplier_name:
        supplier_name = _extract_party_name_from_section(
            text,
            r'SUPPLIER\s+DETAILS',
            [r'BUYER\s+DETAILS', r'GRAND\s+TOTAL', r'Delivery\s+Date', r'\n\s*#\s*\n'],
        )
    if not buyer_name:
        buyer_name = _extract_party_name_from_section(
            text,
            r'BUYER\s+DETAILS',
            [r'Delivery\s+Date', r'GRAND\s+TOTAL', r'\n\s*#\s*\n'],
        )

    # Last-resort narrow fallback if section parsing fails.
    if not supplier_name:
        m = re.search(r'SUPPLIER\s+DETAILS\s+Name\s+([^\n]+)', text, flags=re.IGNORECASE)
        if m:
            supplier_name = _clean_party_name(m.group(1))
    if not buyer_name:
        m = re.search(r'BUYER\s+DETAILS\s+Name\s+([^\n]+)', text, flags=re.IGNORECASE)
        if m:
            buyer_name = _clean_party_name(m.group(1))

    # ── Grand Total ────────────────────────────────────────────────────────────
    grand_total = None
    m = re.search(r'GRAND TOTAL\s+Rs\s+([\d,]+\.?\d*)', text)
    if m:
        grand_total = _clean_amount(m.group(1))

    # ── Line Items ─────────────────────────────────────────────────────────────
    # Prioritize table_rows extraction (most reliable for Utopia 2-row layout),
    # then text-based fallbacks.
    sku_jobname_map = _build_sku_jobname_map(text, table_blobs)

    expected_line_count = _detect_expected_line_count(text, table_rows)

    def _best_of(current, candidate):
        """Pick the better candidate, preferring count closest to expected # lines."""
        if expected_line_count:
            current_gap = abs(len(current) - expected_line_count)
            candidate_gap = abs(len(candidate) - expected_line_count)
            if candidate_gap < current_gap:
                return candidate
            if candidate_gap > current_gap:
                return current
        return candidate if len(candidate) > len(current) else current

    def _needs_fallback(current_items):
        if not current_items:
            return True
        if expected_line_count:
            return len(current_items) != expected_line_count
        return False

    # Try table_rows first (most reliable for structured PDFs)
    items = []
    if table_rows:
        items = _extract_items_from_table_rows(table_rows, sku_jobname_map)
    
    # Fallback to text-based parsers if table extraction insufficient
    if _needs_fallback(items):
        items = _best_of(items, _extract_items_strict(text, sku_jobname_map))
    if _needs_fallback(items):
        items = _best_of(items, _extract_items_flexible(text, sku_jobname_map))
    if _needs_fallback(items):
        if table_blobs:
            items = _best_of(items, _extract_items_from_table_blobs(table_blobs, sku_jobname_map))
    if _needs_fallback(items):
        items = _best_of(items, _extract_items_from_text_windows(text, sku_jobname_map))

    if expected_line_count and len(items) > expected_line_count:
        # Keep deterministic top rows only when parser over-detects noisy lines.
        items = items[:expected_line_count]

    if not items:
        raise ValueError(
            f"Could not detect any line items in the PO PDF. "
            f"Extracted text preview: {text[:400]!r}"
        )

    extraction_warning = None
    if expected_line_count and len(items) < expected_line_count:
        # Dump raw parse data to a temp file for debugging.
        import json as _json, tempfile, os as _os
        _dump = {
            'expected': expected_line_count,
            'items_found': len(items),
            'table_rows': table_rows,
            'full_text_lines': str(text).splitlines(),
        }
        _dump_path = _os.path.join(tempfile.gettempdir(), 'po_extractor_debug.json')
        with open(_dump_path, 'w', encoding='utf-8') as _f:
            _json.dump(_dump, _f, indent=2, default=str)
        extraction_warning = (
            f"PDF has {expected_line_count} line item(s) in # column, "
            f"but only {len(items)} could be parsed. "
            f"Please review the items below and add any missing ones manually."
        )
    elif expected_line_count and len(items) > expected_line_count:
        extraction_warning = (
            f"PDF has {expected_line_count} line item(s) in # column. "
            f"Extra noisy row(s) were ignored and only {expected_line_count} item(s) were kept."
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
        'extraction_warning': extraction_warning,
        'expected_line_count': expected_line_count,
    }


def _detect_expected_line_count(text, table_rows=None):
    serials = set()

    for row in table_rows or []:
        if not row:
            continue
        first = str(row[0]).strip()
        if re.fullmatch(r'\d{1,3}', first):
            serials.add(int(first))

    if serials:
        return len(serials)

    # Text fallback: lines that start with serial number and contain a date.
    for line in str(text).splitlines():
        compact = ' '.join(line.split())
        m = re.match(r'^(\d{1,3})\s+', compact)
        if not m:
            continue
        if re.search(LINE_DATE_REGEX, compact, re.IGNORECASE):
            serials.add(int(m.group(1)))

    return len(serials)


def _append_item(items, sku, delivery_date_raw, qty_raw, unit_raw, unit_cost_raw, subtotal_raw, gst_raw, net_total_raw, sku_jobname_map=None):
    sku_value = _extract_best_sku_token(sku)
    if not sku_value:
        return
    if _looks_like_date_token(sku_value):
        return
    if sku_value.upper() in _SKU_BLOCK_WORDS:
        return
    if sku_value.upper().startswith(('SUBTOTAL', 'GRAND', 'TOTAL', 'DESCRIPTION')):
        return

    try:
        # Handle European qty format: "20,0" -> 20.0 (comma as decimal)
        qty_str = str(qty_raw).replace(',', '.').strip()
        # If multiple dots remain (e.g. "3.000.00"), take first numeric part
        m = re.match(r'^([\d.]+)', qty_str)
        quantity = float(m.group(1)) if m else None
    except (ValueError, AttributeError):
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
    amount = r'(?:Rs\s*|€\s*)?[\d,.]+'
    item_pattern = re.compile(
        r'^([A-Za-z0-9][\w\-./]+)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,.]+)\s+(' + UNIT_PATTERN + r')\s+'
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
    amount = r'(?:Rs\s*|€\s*)?[\d,.]+'
    pattern = re.compile(
        r'(.+?)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,.]+)\s+'
        r'(' + UNIT_PATTERN + r')\s+'
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
    amount = r'(?:Rs\s*|€\s*)?[\d,.]+'
    pattern = re.compile(
        r'(.+?)\s+'
        r'(' + LINE_DATE_REGEX + r')\s+'
        r'([\d,.]+)\s+'
        r'(' + UNIT_PATTERN + r')\s+'
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


def _extract_items_from_table_rows(table_rows, sku_jobname_map=None):
    """
    Handles both single-row and two-row-per-item layouts.

    Two-row layout (as seen in Utopia PO system):
      Row A: ["1", "Job Name / Description", None, None, ...]
      Row B: ["None", "SKUTOKEN", "May 12, 2026", "20.0 PIECE", "40,00 €", ...]
    """
    rows = list(table_rows or [])
    items = []
    seen = set()
    i = 0

    def _clean_cell(c):
        """Strip 'None' sentinel and leading watermark chars like 'D\n'."""
        s = str(c).strip()
        if s.lower() == 'none':
            return ''
        # Remove single-letter watermark prefix on its own line  e.g. "D\nMay 12"
        s = re.sub(r'^[A-Z]\n', '', s).strip()
        return s

    while i < len(rows):
        row = rows[i]
        if not row:
            i += 1
            continue

        first = _clean_cell(row[0])

        # ── Two-row layout: serial number row followed by data row ──────────
        if re.fullmatch(r'\d{1,3}', first) and i + 1 < len(rows):
            serial = first
            job_name_raw = _clean_cell(row[1]) if len(row) > 1 else ''
            # Strip watermark prefix from job name
            job_name_raw = re.sub(r'^[A-Z]\n', '', job_name_raw).strip()
            # Take only the part before " / " if present (description / specs)
            job_name_raw = job_name_raw.split(' / ')[0].strip()

            next_row = rows[i + 1]
            next_first = _clean_cell(next_row[0]) if next_row else 'x'

            if next_first == '' or next_first.lower() == 'none':
                # Data row
                data_cells = [_clean_cell(c) for c in next_row]
                data_cells = [c for c in data_cells if c]  # drop empty/None

                if len(data_cells) >= 2:
                    sku_raw = data_cells[0]

                    # Find delivery date cell
                    date_idx = None
                    for idx, cell in enumerate(data_cells):
                        if _parse_date(cell):
                            date_idx = idx
                            break

                    if date_idx is not None:
                        delivery_date_raw = data_cells[date_idx]
                        after = data_cells[date_idx + 1:]

                        qty_raw = None
                        unit_raw = ''
                        qty_cell_idx = None
                        for cell_idx, cell in enumerate(after):
                            m = re.search(
                                r'([\d,.]+)\s*(' + UNIT_PATTERN + r')',
                                cell, re.IGNORECASE,
                            )
                            if m:
                                qty_raw = m.group(1)
                                unit_raw = m.group(2)
                                qty_cell_idx = cell_idx
                                break

                        if qty_raw is None:
                            for cell_idx, cell in enumerate(after):
                                m = re.match(r'^([\d,.]+)$', cell)
                                if m:
                                    qty_raw = m.group(1)
                                    qty_cell_idx = cell_idx
                                    break

                        if qty_raw is not None:
                            # Amount fields — skip the qty+unit cell by index
                            amount_cells = [
                                c for idx, c in enumerate(after)
                                if idx != qty_cell_idx
                                and re.search(r'[\d,.]', c)
                                and not re.fullmatch(UNIT_PATTERN, c, re.IGNORECASE)
                            ]
                            unit_cost_raw = amount_cells[0] if len(amount_cells) > 0 else None
                            subtotal_raw = amount_cells[1] if len(amount_cells) > 1 else None
                            gst_raw = amount_cells[2] if len(amount_cells) > 2 else None
                            net_total_raw = amount_cells[3] if len(amount_cells) > 3 else None

                            key = (serial, sku_raw.upper(), delivery_date_raw, str(qty_raw))
                            if key not in seen:
                                before = len(items)
                                # Use job_name from serial row if sku_jobname_map doesn't have it
                                effective_map = dict(sku_jobname_map or {})
                                if job_name_raw and sku_raw.upper() not in effective_map:
                                    effective_map[sku_raw.upper()] = job_name_raw
                                _append_item(
                                    items, sku_raw, delivery_date_raw, qty_raw, unit_raw,
                                    unit_cost_raw, subtotal_raw, gst_raw, net_total_raw,
                                    effective_map,
                                )
                                if len(items) > before:
                                    seen.add(key)
                i += 2
                continue

        # ── Single-row layout fallback ──────────────────────────────────────
        if re.fullmatch(r'\d{1,3}', first):
            cells = [_clean_cell(c) for c in row]
            cells = [c for c in cells if c]
            if len(cells) >= 3:
                date_idx = None
                for idx, cell in enumerate(cells):
                    if _parse_date(cell):
                        date_idx = idx
                        break
                if date_idx is not None:
                    sku_candidate = ' '.join(cells[1:date_idx]).strip() if date_idx > 1 else (cells[1] if len(cells) > 1 else '')
                    delivery_date_raw = cells[date_idx]
                    after = cells[date_idx + 1:]
                    qty_raw = None
                    unit_raw = ''
                    qty_cell_idx = None
                    for cell_idx, cell in enumerate(after):
                        m = re.search(r'([\d,.]+)\s*(' + UNIT_PATTERN + r')', cell, re.IGNORECASE)
                        if m:
                            qty_raw, unit_raw = m.group(1), m.group(2)
                            qty_cell_idx = cell_idx
                            break
                    if qty_raw is None:
                        for cell_idx, cell in enumerate(after):
                            m = re.match(r'^([\d,.]+)$', cell)
                            if m:
                                qty_raw = m.group(1)
                                qty_cell_idx = cell_idx
                                break
                    if qty_raw is not None:
                        amount_cells = [
                            c for idx, c in enumerate(after)
                            if idx != qty_cell_idx and re.search(r'[\d,.]', c)
                        ]
                        _append_item(
                            items, sku_candidate, delivery_date_raw, qty_raw, unit_raw,
                            amount_cells[0] if amount_cells else None,
                            amount_cells[1] if len(amount_cells) > 1 else None,
                            amount_cells[2] if len(amount_cells) > 2 else None,
                            amount_cells[3] if len(amount_cells) > 3 else None,
                            sku_jobname_map,
                        )
        i += 1

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
        r'([\d,.]+)\s*(' + UNIT_PATTERN + r')',
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
