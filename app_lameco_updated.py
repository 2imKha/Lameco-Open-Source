import io, os, re, shutil, tempfile, zipfile, subprocess, glob
from datetime import datetime

from flask import (
    Flask, request, redirect, url_for,
    send_file, render_template_string, flash
)
from lxml import etree
import pandas as pd
import xml.etree.ElementTree as ET
import hashlib

# For docx template comparison
from collections import Counter

import json

# Embedded Lameco logo (data URI) for a zero-config UI
LOGO_DATA_URI = 'data:image/png;base64,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'

# ----------------------------
# Persistence state for checkboxes
# ----------------------------
# File to persist checkbox states across restarts. It stores a mapping
# from file_hash (MD5 of the uploaded file) to a list of row indices that were marked as done.
STATE_FILE = os.path.join(os.path.dirname(__file__), 'checkbox_state.json')
try:
    with open(STATE_FILE, 'r') as _sf:
        STATE = json.load(_sf)
except Exception:
    STATE = {}

def save_state():
    """Persist the STATE dictionary to disk."""
    try:
        with open(STATE_FILE, 'w') as _sf:
            json.dump(STATE, _sf)
    except Exception:
        pass


# ----------------------------
# Persistence state for UI settings
# ----------------------------
# File to persist user-configurable settings across restarts.
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'ui_settings.json')

DEFAULT_SETTINGS = {
    "pptx": {
        # Slide size in inches (PowerPoint uses EMU internally; 914400 EMU = 1 inch)
        "slide_preset": "Widescreen (16:9)",
        "slide_width_in": 13.333,
        "slide_height_in": 7.5,

        # Font rules
        "global_font": "Montserrat",
        "title_font": "Montserrat",
        "body_font": "Montserrat",

        # Size rules (pt)
        "title_size_pt": 60,
        "body_size_pt": 28,
        "allowed_sizes_pt": [16, 18, 20, 24, 28, 32, 40, 42, 44, 48, 49, 54, 60, 72],
        "allowed_title_sizes_pt": [54, 60, 72],

        # Optional slide-specific rules (can be disabled)
        "enable_slide_specific_rules": True,
        "agenda_label_text": "Agenda",
        "agenda_label_size_pt": 49,
        "agenda_index_size_pt": 44,
        "agenda_item_size_pt": 28,
        "closing_tagline_text": "/ digitaal denken/ digitaal doen",
        "closing_tagline_size_pt": 60,
    },
    "docx": {
        "page_size": "A4",
        "margin_top_mm": 25,
        "margin_right_mm": 25,
        "margin_bottom_mm": 25,
        "margin_left_mm": 25,

        # Backwards-compatible aliases (old keys)
        "default_font": "Montserrat", # Updated from app_updated.py logic
        "font_size_pt": 12,

        # Editable rules
        "body_font": "Montserrat", # Updated from app_updated.py logic
        "body_size_pt": 12,        # Updated from app_updated.py logic
        "heading_font": "Montserrat",
        "heading_size_pt": 32,     # Used for Title in logic default
        "line_spacing": 1.15,

        # If True, also compare against the reference template document baseline
        "use_template_baseline": False,
    }
}

def _deep_merge_defaults(dst: dict, defaults: dict):
    """Recursively fill missing keys in dst using defaults (in-place)."""
    if not isinstance(dst, dict) or not isinstance(defaults, dict):
        return
    for k, v in defaults.items():
        if isinstance(v, dict):
            dst.setdefault(k, {})
            _deep_merge_defaults(dst[k], v)
        else:
            dst.setdefault(k, v)

try:
    with open(SETTINGS_FILE, 'r') as _sf:
        SETTINGS = json.load(_sf)
except Exception:
    SETTINGS = DEFAULT_SETTINGS.copy()

_deep_merge_defaults(SETTINGS, DEFAULT_SETTINGS)

def save_settings():
    """Persist SETTINGS to disk."""
    try:
        with open(SETTINGS_FILE, 'w') as _sf:
            json.dump(SETTINGS, _sf, indent=2)
    except Exception:
        pass

# ----------------------------
# Persistence state for History
# ----------------------------
# File to keep a short history of processed documents across restarts.
HISTORY_FILE = os.path.join(os.path.dirname(__file__), 'history.json')
try:
    with open(HISTORY_FILE, 'r') as _hf:
        HISTORY = json.load(_hf)
        if not isinstance(HISTORY, list):
            HISTORY = []
except Exception:
    HISTORY = []

def save_history():
    """Persist HISTORY list to disk."""
    try:
        with open(HISTORY_FILE, 'w') as _hf:
            json.dump(HISTORY[-50:], _hf, indent=2)
    except Exception:
        pass

def add_history_entry(entry: dict):
    """Add/update a history entry (deduplicated) and persist it (keeps last 50).

    Deduplication rule:
      - Prefer file_hash (stable across filenames).
      - Fallback to kind+filename.
    If an entry already exists, we update its stats and bump it to newest instead
    of creating duplicates.
    """
    try:
        entry = dict(entry or {})
        now = datetime.now().isoformat(timespec="seconds")
        entry.setdefault("created_at", now)
        entry["last_opened_at"] = now

        kind = entry.get("kind") or ""
        filename = entry.get("filename") or ""
        file_hash = entry.get("file_hash") or ""
        key = file_hash if file_hash else f"{kind}::{filename}"
        entry["key"] = key

        # Find existing entry
        existing_idx = None
        for i, it in enumerate(HISTORY):
            if (it or {}).get("key") == key:
                existing_idx = i
                break

        if existing_idx is not None:
            old = HISTORY.pop(existing_idx) or {}
            # keep original created_at if present
            entry["created_at"] = old.get("created_at", entry["created_at"])

        HISTORY.append(entry)
        save_history()
    except Exception:
        pass

# ----------------------------
# Config
# ----------------------------
ALLOWED_EXTS = {'.pptx', '.zip', '.docx'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100 MB

app = Flask(__name__)
app.secret_key = "dev-only-key"
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# ----------------------------
# Namespaces
# ----------------------------
ns_ppt = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

# wordprocessingML namespaces
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ns_docx = {'w': NS_W}

# ----------------------------------------------------------------------------
# DOCX template baseline and checking
#
# To validate Word documents against a reference template, we compute a baseline
# of formatting properties (font family, size, bold, italic, underline) for
# each paragraph style in the provided template document.  The baseline is
# computed lazily on first use from the file ``Zoekwoordenonderzoek juist.docx``
# located next to this script.  During document analysis, each run is
# compared against the baseline for its paragraph style and any mismatches
# generate a template error.  See ``check_docx_template`` for details.
# ----------------------------------------------------------------------------

# Global cache for the baseline.
DOCX_TEMPLATE_BASELINE = None
_WORD_TEMPLATE_RULES_CACHE = None

def _compute_docx_baseline():
    """Compute the baseline formatting for each paragraph style in the template."""
    global DOCX_TEMPLATE_BASELINE
    if DOCX_TEMPLATE_BASELINE is not None:
        return DOCX_TEMPLATE_BASELINE
    baseline = {}
    template_path = os.path.join(os.path.dirname(__file__), 'Zoekwoordenonderzoek juist.docx')
    try:
        # Extract features from the template document using the existing
        # extractor.
        recs = extract_docx_features(template_path)
    except Exception:
        DOCX_TEMPLATE_BASELINE = {}
        return DOCX_TEMPLATE_BASELINE
    
    from collections import defaultdict
    style_groups = defaultdict(list)
    for rec in recs:
        style = rec.get('paragraph_style')
        style_groups[style].append(rec)
    
    def most_common(seq):
        seq = [x for x in seq if x is not None]
        if not seq:
            return None
        return Counter(seq).most_common(1)[0][0]
        
    for style, records in style_groups.items():
        fonts = [r.get('font_name') for r in records]
        sizes = [r.get('font_size_pt') for r in records if r.get('font_size_pt') is not None]
        bolds = [r.get('bold') for r in records]
        italics = [r.get('italic') for r in records]
        underlines = [r.get('underline') for r in records]
        baseline[style] = {
            'font_name': most_common(fonts),
            'font_size_pt': most_common([int(s) if s is not None else None for s in sizes]) if sizes else None,
            'bold': most_common(bolds),
            'italic': most_common(italics),
            'underline': most_common(underlines),
        }
    DOCX_TEMPLATE_BASELINE = baseline
    return DOCX_TEMPLATE_BASELINE

def get_word_template_rules():
    """Return a mapping of paragraph styles to their expected formatting.

    The returned dict maps lower‑cased paragraph style names to a
    dictionary describing the expected font family (``font_name``), font
    size in points (``font_size_pt``), and boolean flags for
    ``bold``, ``italic`` and ``underline``.  The values are derived
    from the reference Word document found next to this script.  If
    the reference file cannot be read, the cache will be empty and
    validation will effectively be disabled until a valid baseline can
    be computed.

    Additionally, if the reference document does not contain a
    ``Footer`` style, a default footer rule (Calibri 10pt, no
    bold/italic/underline) is added so that footers can still be
    validated.
    """
    global _WORD_TEMPLATE_RULES_CACHE
    if _WORD_TEMPLATE_RULES_CACHE is not None:
        return _WORD_TEMPLATE_RULES_CACHE
    
    # Check if we should use baseline from settings (default True to match app_updated behavior)
    use_baseline = True
    if 'SETTINGS' in globals() and 'docx' in SETTINGS:
        # If user explicitly sets to False, we respect it.
        # But we default to True if key is missing to keep robust legacy behavior.
        use_baseline = SETTINGS['docx'].get('use_template_baseline', True)
        
    baseline = {}
    if use_baseline:
        baseline = _compute_docx_baseline()

    rules = {}
    # Lower‑case the style names for case‑insensitive lookup
    for style_name, attrs in baseline.items():
        key = (style_name or '').lower()
        if not key:
            continue
        rules[key] = {
            'font_name': attrs.get('font_name'),
            'font_size_pt': attrs.get('font_size_pt'),
            'bold': attrs.get('bold'),
            'italic': attrs.get('italic'),
            'underline': attrs.get('underline'),
        }
    # If the template is missing a footer definition, fall back to a
    # sensible default.  The footer in the Lameco documents uses
    # Calibri 10pt with no additional styling.
    if 'footer' not in rules:
        rules['footer'] = {
            'font_name': 'Calibri',
            'font_size_pt': 10,
            'bold': False,
            'italic': False,
            'underline': False,
        }
    # When the baseline cannot be computed (e.g. the reference document
    # is missing), ``rules`` may only contain the fallback footer entry.
    # Provide a static default mapping in that case so that the style
    # checker has sensible values to work with.  These defaults were
    # extracted from the official template and should be updated when
    # the template changes.
    if not rules or (len(rules) == 1 and 'footer' in rules):
        if 'SETTINGS' in globals() and 'docx' in SETTINGS:
             d = SETTINGS['docx']
             body_font = d.get('body_font', 'Montserrat')
             heading_font = d.get('heading_font', 'Montserrat')
             try: body_size = int(round(float(d.get('body_size_pt', 12))))
             except: body_size = 12
             try: heading_size = int(round(float(d.get('heading_size_pt', 32))))
             except: heading_size = 32
             try: title_size = int(round(float(d.get('title_size_pt', heading_size))))
             except: title_size = heading_size
             title_font = d.get('title_font', heading_font)

             rules = {
                'title': {
                    'font_name': title_font,
                    'font_size_pt': title_size,
                    'bold': True,
                    'italic': False,
                    'underline': False,
                },
                'subtitle': {
                    'font_name': heading_font,
                    'font_size_pt': 20,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                'normal': {
                    'font_name': body_font,
                    'font_size_pt': body_size,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                'list paragraph': {
                    'font_name': body_font,
                    'font_size_pt': body_size,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                'footer': {
                    'font_name': 'Calibri',
                    'font_size_pt': 10,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
             }
        else:
            rules = {
                'title': {
                    'font_name': 'Montserrat',
                    'font_size_pt': 32,
                    'bold': True,
                    'italic': False,
                    'underline': False,
                },
                'subtitle': {
                    'font_name': 'Montserrat',
                    'font_size_pt': 20,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                # The body paragraphs in the template use 12pt Montserrat. 10pt
                # is reserved for table cell contents.
                'normal': {
                    'font_name': 'Montserrat',
                    'font_size_pt': 12,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                # List paragraphs follow the normal body size unless within a
                # table, in which case 10pt is expected.
                'list paragraph': {
                    'font_name': 'Montserrat',
                    'font_size_pt': 12,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
                'footer': {
                    'font_name': 'Calibri',
                    'font_size_pt': 10,
                    'bold': False,
                    'italic': False,
                    'underline': False,
                },
            }
    _WORD_TEMPLATE_RULES_CACHE = rules
    return rules

def check_docx_template(records):
    """
    Validate Word document formatting against a reference template.

    ``records`` should be a list of dictionaries returned by
    ``extract_docx_features``.  Each record contains information about a run
    of text including its paragraph style, font name, size and formatting.
    The function compares each run to the baseline for its paragraph style
    and yields errors for any mismatches.

    Returns:
      - errors: list of dicts with keys {code, page, row_index, message, text}
      - record_error_map: mapping from row index to list of error codes
    """
    # Compute the expected formatting rules from the reference template.
    # This call populates and returns a cache so subsequent calls are cheap.
    rules = get_word_template_rules()
    errors = []
    record_error_map = defaultdict(list)
    
    # Load settings for dynamic checks
    doc_cfg = (SETTINGS.get("docx") if isinstance(SETTINGS, dict) else {}) or {}
    try: conf_body_size = int(doc_cfg.get('body_size_pt', 12))
    except: conf_body_size = 12
    
    conf_body_font = str(doc_cfg.get('body_font', 'Montserrat')).strip()
    if not conf_body_font:
        conf_body_font = str(doc_cfg.get('default_font', 'Montserrat')).strip()
    
    # Configure dynamic Titles if user overrides them in settings
    # Default to 32 only if both title_size and heading_size are missing/invalid
    try: 
        ts = doc_cfg.get('title_size_pt')
        if ts is not None:
             conf_title_size = int(round(float(ts)))
        else:
             hs = doc_cfg.get('heading_size_pt')
             if hs is not None:
                 conf_title_size = int(round(float(hs)))
             else:
                 conf_title_size = 32
    except: 
        conf_title_size = 32
    conf_title_font = str(doc_cfg.get('title_font', doc_cfg.get('heading_font', 'Montserrat'))).strip()

    if not records:
        return [], {}
    for idx, rec in enumerate(records):
        # Determine effective style key.  If the run originates from the
        # document footer, always treat it as 'footer' regardless of the
        # declared paragraph style.  This avoids false positives when
        # footers reuse the normal style but use a different font.
        style = rec.get('paragraph_style') or ''
        location = rec.get('location') or ''
        page = rec.get('page')
        text_val = rec.get('text', '')
        # Normalise style for lookup: lower‑case; override to 'footer' if
        # location is footer.
        if location == 'footer':
            style_key = 'footer'
            style = 'Footer'
        else:
            style_key = style.lower()

        rule = rules.get(style_key)
        if not rule:
            # Unknown style
            msg = (
                f"Paragraph style '{style}' is not part of the template. "
                f"Use one of: {', '.join(sorted(rules.keys()))}."
            )
            errors.append({
                'code': 'STYLE',
                'page': page,
                'row_index': idx,
                'message': msg,
                'text': text_val,
            })
            record_error_map[idx].append('STYLE')
            continue

        # Font family check
        expected_font = rule.get('font_name')
        if style_key == 'title':
            expected_font = conf_title_font
            
        font_name = rec.get('font_name')
        
        # Check if the font matches the rule
        if expected_font and font_name and font_name != expected_font:
            # If the font matches the user-configured body font, we allow it for body styles
            is_body = style_key in ('normal', 'list paragraph', 'body text')
            if is_body and font_name == conf_body_font:
                pass
            else:
                msg = (
                    f"Paragraph style '{style}' should use font {expected_font} "
                    f"(found '{font_name}')."
                )
                errors.append({
                    'code': 'FONT_FAMILY',
                    'page': page,
                    'row_index': idx,
                    'message': msg,
                    'text': text_val,
                })
                record_error_map[idx].append('FONT_FAMILY')

        # Determine expected size.  Body text in tables is 10 pt; otherwise
        # body styles ('normal' and 'list paragraph') use the configured Size.  For
        # other styles we rely on the baseline.
        expected_size = rule.get('font_size_pt')
        if style_key in ('normal', 'list paragraph'):
            # override expected size depending on location
            if location == 'table':
                expected_size = 10
            else:
                expected_size = conf_body_size
        elif style_key == 'title':
            expected_size = conf_title_size

        sz = rec.get('font_size_pt')
        if expected_size is not None and sz is not None:
            try:
                sz_int = int(round(float(sz)))
            except Exception:
                sz_int = None
            if sz_int is not None and sz_int != expected_size:
                msg = (
                    f"Paragraph style '{style}' should have font size {expected_size} pt "
                    f"(found {sz} pt)."
                )
                errors.append({
                    'code': 'FONT_SIZE',
                    'page': page,
                    'row_index': idx,
                    'message': msg,
                    'text': text_val,
                })
                record_error_map[idx].append('FONT_SIZE')
        # Bold enforcement: only apply for titles and subtitles.  For other
        # styles we do not enforce bold to allow table headers and other
        # emphasised text.
        if style_key in ('title', 'subtitle'):
            expected_bold = rule.get('bold')
            bold_val = bool(rec.get('bold'))
            if expected_bold is not None and bold_val != bool(expected_bold):
                msg = (
                    f"Paragraph style '{style}' bold should be {expected_bold} "
                    f"(found {bold_val})."
                )
                errors.append({
                    'code': 'BOLD',
                    'page': page,
                    'row_index': idx,
                    'message': msg,
                    'text': text_val,
                })
                record_error_map[idx].append('BOLD')
        # General font check: disallow exotic fonts for non‑title/subtitle/footer styles.
        if style_key not in ('title', 'subtitle', 'footer'):
            fn = rec.get('font_name')
            allowed_fonts = {conf_body_font, 'Montserrat', 'Calibri'}
            if fn and fn not in allowed_fonts:
                msg = (
                    f"Font '{fn}' is not allowed in style '{style}'. "
                    f"Use {conf_body_font} for body text or Calibri for footers."
                )
                errors.append({
                    'code': 'FONT_FAMILY',
                    'page': page,
                    'row_index': idx,
                    'message': msg,
                    'text': text_val,
                })
                record_error_map[idx].append('FONT_FAMILY')
        # Underline enforcement: for list paragraph, require underline only
        # when the run ends with a colon (e.g. "Categorie:").  Do not
        # enforce underline on the following run.
        if style_key == 'list paragraph' and text_val:
            # Strip whitespace and inspect trailing colon
            stripped = text_val.strip()
            if ':' in stripped and stripped.endswith(':'):
                underline_val = bool(rec.get('underline'))
                if not underline_val:
                    msg = (
                        f"List labels (text before a colon) should be underlined. "
                        f"Style '{style}', text '{text_val}' is missing underline."
                    )
                    errors.append({
                        'code': 'UNDERLINE',
                        'page': page,
                        'row_index': idx,
                        'message': msg,
                        'text': text_val,
                    })
                    record_error_map[idx].append('UNDERLINE')
    return errors, dict(record_error_map)
def render_ppt_thumbnails(src_path, out_dir):
    """Render PPTX slides to PNG images.

    Primary path:
      1) Use LibreOffice (soffice) to convert the PPTX to a PDF.
      2) Use `pdftoppm` (Poppler) to convert each PDF page to a PNG.

    This reliably produces one image per slide on most platforms.
    If either step fails (no soffice / no pdftoppm / non‑zero exit), we
    fall back to LibreOffice's direct PNG export which at least gives the
    first slide.

    Returns a list of dicts: [{"page": n, "filename": "slide_n.png"}, ...]
    If everything fails, returns an empty list.
    """
    os.makedirs(out_dir, exist_ok=True)
    slides = []
    pngs = []

    # --- Preferred: PPTX -> PDF -> PNG (one image per page) ---
    try:
        pdf_cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            src_path,
        ]
        subprocess.run(pdf_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_candidates = sorted(glob.glob(os.path.join(out_dir, "*.pdf")))
        if not pdf_candidates:
            raise RuntimeError("No PDF produced from PPTX")
        pdf_path = pdf_candidates[0]

        ppm_prefix = os.path.join(out_dir, "slide")
        img_cmd = [
            "pdftoppm",
            "-png",
            pdf_path,
            ppm_prefix,
        ]
        subprocess.run(img_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # Typical pdftoppm output names: slide-1.png, slide-01.png, etc.
        pngs = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
        if not pngs:
            pngs = sorted(glob.glob(os.path.join(out_dir, "slide*.png")))
    except Exception:
        # --- Fallback: LibreOffice direct PNG export (often only first slide) ---
        try:
            cmd = [
                "soffice",
                "--headless",
                "--convert-to", "png",
                "--outdir", out_dir,
                src_path,
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            pngs = sorted(glob.glob(os.path.join(out_dir, "*.png")))
        except Exception:
            return []

    if not pngs:
        return []

    for idx, png_path in enumerate(sorted(pngs), start=1):
        fname = f"slide_{idx}.png"
        dst = os.path.join(out_dir, fname)
        if os.path.abspath(png_path) != os.path.abspath(dst):
            shutil.move(png_path, dst)
        slides.append({"page": idx, "filename": fname})

    return slides



def render_docx_thumbnails(src_path, out_dir):
    """Render DOCX pages to PNG images.

    Preferred path:
      1) LibreOffice (soffice) converts DOCX -> PDF.
      2) pdftoppm converts PDF -> PNG (one image per page).

    Returns a list of dicts: [{"page": n, "filename": "page_n.png"}, ...]
    If everything fails, returns an empty list.
    """
    os.makedirs(out_dir, exist_ok=True)
    pages = []
    pngs = []

    # --- Preferred: DOCX -> PDF -> PNG ---
    try:
        pdf_cmd = [
            "soffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            src_path,
        ]
        subprocess.run(pdf_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_candidates = sorted(glob.glob(os.path.join(out_dir, "*.pdf")))
        if not pdf_candidates:
            raise RuntimeError("No PDF produced from DOCX")
        pdf_path = pdf_candidates[0]

        prefix = os.path.join(out_dir, "page")
        img_cmd = [
            "pdftoppm",
            "-png",
            "-r", "150",
            pdf_path,
            prefix,
        ]
        subprocess.run(img_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pngs = sorted(glob.glob(os.path.join(out_dir, "page-*.png")))
        if not pngs:
            pngs = sorted(glob.glob(os.path.join(out_dir, "page*.png")))
    except Exception:
        # --- Fallback: LibreOffice direct PNG export ---
        try:
            cmd = [
                "soffice",
                "--headless",
                "--convert-to", "png",
                "--outdir", out_dir,
                src_path,
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            pngs = sorted(glob.glob(os.path.join(out_dir, "*.png")))
        except Exception:
            return []

    if not pngs:
        return []

    # Sort in numeric page order when possible (pdftoppm uses suffix numbers)
    def _num_key(p):
        b = os.path.basename(p)
        m = re.search(r"-(\d+)\.png$", b)
        if not m:
            m = re.search(r"(\d+)\.png$", b)
        return int(m.group(1)) if m else 10**9

    pngs_sorted = sorted(pngs, key=_num_key)

    for idx, png_path in enumerate(pngs_sorted, start=1):
        fname = f"page_{idx}.png"
        dst = os.path.join(out_dir, fname)
        if os.path.abspath(png_path) != os.path.abspath(dst):
            try:
                shutil.move(png_path, dst)
            except Exception:
                # In rare cases, move can fail across filesystems; fall back to copy.
                shutil.copyfile(png_path, dst)
        pages.append({"page": idx, "filename": fname})

    return pages


def first_nonempty(seq):
    if not seq:
        return None
    for x in seq:
        if x is None:
            continue
        s = str(x).strip()
        if s != "":
            return s
    return None


# =============================================================================
# PowerPoint feature extractor (adapted from your existing app) :contentReference[oaicite:0]{index=0}
# =============================================================================
_WEIGHT_TOKENS = (
    " thin"," extralight"," ultralight"," light"," semilight"," regular"," book",
    " medium"," semibold"," demibold"," bold"," extrabold"," ultrabold"," black"," heavy"
)

DO_NOT_STRIP = {"Arial Black","Archivo Black","Avenir Next Condensed","DIN Condensed","Impact"}

def as_pt(sz):
    try:
        return int(sz) / 100.0
    except Exception:
        return None

def as_bool01(val):
    if val is None:
        return None
    v = str(val).strip().lower()
    if v in ("1", "true", "on"):
        return True
    if v in ("0", "false", "off"):
        return False
    return None

def ui_family(face):
    if not face:
        return None
    if face in DO_NOT_STRIP:
        return face
    s = face.strip()
    low = s.lower()
    for tok in _WEIGHT_TOKENS:
        if low.endswith(tok):
            return s[:len(s)-len(tok)].rstrip()
    return s

def slide_num_from_path(path):
    m = re.search(r'slide(\d+)\.xml$', os.path.basename(path))
    return int(m.group(1)) if m else 0

def get_linked_file(xml_path, link_type, ppt_root):
    if not xml_path:
        return None

    if str(xml_path).endswith(".rels"):
        rels_path = xml_path
        rels_dir  = os.path.dirname(rels_path)
        part_dir  = os.path.dirname(rels_dir)
        part_name = os.path.basename(rels_path).replace(".rels", "")
        part_path = os.path.join(part_dir, part_name)
    else:
        part_path = xml_path
        part_dir  = os.path.dirname(part_path)
        rels_path = os.path.join(part_dir, "_rels", os.path.basename(part_path) + ".rels")

    if not os.path.exists(rels_path):
        return None

    rel_tree = etree.parse(rels_path)
    targets = rel_tree.xpath(
        f'//rel:Relationship[contains(@Type, "{link_type}")]/@Target', namespaces=ns_ppt
    )
    if not targets:
        return None

    target = targets[0]
    abs_path = os.path.normpath(os.path.join(part_dir, target))
    if os.path.exists(abs_path):
        return abs_path

    candidate = os.path.normpath(os.path.join(ppt_root, target.lstrip("/")))
    if os.path.exists(candidate):
        return candidate

    return None

def get_theme_path(ppt_root):
    pres_rels = os.path.join(ppt_root, "_rels", "presentation.xml.rels")
    if os.path.exists(pres_rels):
        rel_tree = etree.parse(pres_rels)
        targets = rel_tree.xpath(
            '//rel:Relationship[contains(@Type,"/theme")]/@Target', namespaces=ns_ppt
        )
        if targets:
            path = os.path.normpath(os.path.join(ppt_root, targets[0]))
            if os.path.exists(path):
                return path
    default = os.path.join(ppt_root, "theme", "theme1.xml")
    return default if os.path.exists(default) else None

def resolve_theme_font(ppt_root, theme_font_ref):
    theme_path = get_theme_path(ppt_root)
    if not theme_path:
        return theme_font_ref
    theme_tree = etree.parse(theme_path)
    mapping = {
        '+mn-lt': '//a:minorFont/a:latin/@typeface',
        '+mj-lt': '//a:majorFont/a:latin/@typeface',
        '+mn-ea': '//a:minorFont/a:ea/@typeface',
        '+mj-ea': '//a:majorFont/a:ea/@typeface',
        '+mn-cs': '//a:minorFont/a:cs/@typeface',
        '+mj-cs': '//a:majorFont/a:cs/@typeface',
    }
    xp = mapping.get(theme_font_ref)
    if not xp:
        return theme_font_ref
    vals = theme_tree.xpath(xp, namespaces=ns_ppt)
    return vals[0] if vals else theme_font_ref

def extract_from_rPr(rPr):
    font = size = bold = italic = underline = None
    latin = rPr.xpath('./a:latin/@typeface', namespaces=ns_ppt)
    if latin:
        font = latin[0]
    else:
        buFont = rPr.xpath('./a:buFont/@typeface', namespaces=ns_ppt)
        if buFont:
            font = buFont[0]
    size = rPr.get('sz')
    bold = rPr.get('b')
    italic = rPr.get('i')
    underline = rPr.get('u')
    return font, size, bold, italic, underline

def set_if_missing(cur_list, new_val):
    return cur_list or ([new_val] if new_val is not None else [])

def set_bool_if_missing(cur_list, new_val):
    return cur_list if cur_list else ([new_val] if new_val is not None else [])

def find_placeholder_style(tree, ph_type, ph_idx, para_level):
    candidates = []
    xps = []
    if ph_type and ph_idx:
        xps.append(f'//p:sp[.//p:ph[@type="{ph_type}" and @idx="{ph_idx}"]]')
    if ph_type:
        xps.append(f'//p:sp[.//p:ph[@type="{ph_type}"]]')
    if ph_idx:
        xps.append(f'//p:sp[.//p:ph[@idx="{ph_idx}"]]')
    xps.extend([
        '//p:sp[.//p:ph[@type="body"]]',
        '//p:sp[.//p:ph[@type="ctrTitle"]]',
        '//p:sp[.//p:ph[@type="title"]]',
        '//p:sp[.//p:ph]'
    ])
    for xp in xps:
        nodes = tree.xpath(xp, namespaces=ns_ppt)
        if nodes:
            candidates = nodes
            break

    if not candidates:
        return find_text_styles_in_master(tree, ph_type, para_level)

    lvl_xpath = f'.//a:lstStyle/a:lvl{para_level + 1}pPr/a:defRPr'
    for sp in candidates:
        rPr = sp.xpath(lvl_xpath, namespaces=ns_ppt)
        if rPr:
            vals = extract_from_rPr(rPr[0])
            if any(vals):
                return vals
        rPr = sp.xpath('.//a:defRPr', namespaces=ns_ppt)
        if rPr:
            vals = extract_from_rPr(rPr[0])
            if any(vals):
                return vals
        rPr = sp.xpath('.//a:pPr/a:defRPr', namespaces=ns_ppt)
        if rPr:
            vals = extract_from_rPr(rPr[0])
            if any(vals):
                return vals

    return find_text_styles_in_master(tree, ph_type, para_level)

def find_text_styles_in_master(tree, ph_type, para_level):
    style_map = {
        'title': 'titleStyle', 'ctrTitle': 'titleStyle',
        'subTitle': 'bodyStyle', 'body': 'bodyStyle', None: 'bodyStyle'
    }
    style_name = style_map.get(ph_type, 'bodyStyle')

    r = tree.xpath(f'//p:txStyles/p:{style_name}/a:lvl{para_level + 1}pPr/a:defRPr', namespaces=ns_ppt)
    if r:
        vals = extract_from_rPr(r[0])
        if any(vals):
            return vals

    r = tree.xpath(f'//p:txStyles/p:{style_name}//a:defRPr', namespaces=ns_ppt)
    if r:
        vals = extract_from_rPr(r[0])
        if any(vals):
            return vals

    r = tree.xpath('//p:txStyles//a:defRPr', namespaces=ns_ppt)
    if r:
        vals = extract_from_rPr(r[0])
        if any(vals):
            return vals

    return None, None, None, None, None

def trace_to_layout_master(slide_path, ppt_root, ph_type, ph_idx, para_level):
    font = size = bold = italic = underline = None

    layout_path = get_linked_file(slide_path, "slideLayout", ppt_root)
    if layout_path and os.path.exists(layout_path):
        layout_tree = etree.parse(layout_path)
        l_font, l_size, l_bold, l_italic, l_underline = find_placeholder_style(layout_tree, ph_type, ph_idx, para_level)
        font      = font or l_font
        size      = size or l_size
        bold      = bold if bold is not None else l_bold
        italic    = italic if italic is not None else l_italic
        underline = underline or l_underline

    master_path = get_linked_file(layout_path, "slideMaster", ppt_root) if layout_path else None
    if not master_path:
        master_path = get_linked_file(slide_path, "slideMaster", ppt_root)
    if master_path and os.path.exists(master_path):
        master_tree = etree.parse(master_path)
        m_font, m_size, m_bold, m_italic, m_underline = find_placeholder_style(master_tree, ph_type, ph_idx, para_level)
        font      = font or m_font
        size      = size or m_size
        bold      = bold if bold is not None else m_bold
        italic    = italic if italic is not None else m_italic
        underline = underline or m_underline

        if not font:
            t_font, *_ = find_text_styles_in_master(master_tree, ph_type, para_level)
            font = t_font or font

    return font, size, bold, italic, underline

def get_autofit_scale_shape(shape):
    fs = shape.xpath('.//p:txBody/a:bodyPr/a:normAutofit/@fontScale', namespaces=ns_ppt)
    if fs:
        try:
            return int(fs[0]) / 100000.0
        except Exception:
            return None
    if shape.xpath('.//p:txBody/a:bodyPr/a:spAutoFit', namespaces=ns_ppt):
        return 0.0
    return None

def find_placeholder_node(tree, ph_type, ph_idx):
    xps = []
    if ph_type and ph_idx:
        xps.append(f'//p:sp[.//p:ph[@type="{ph_type}" and @idx="{ph_idx}"]]')
    if ph_type:
        xps.append(f'//p:sp[.//p:ph[@type="{ph_type}"]]')
    if ph_idx:
        xps.append(f'//p:sp[.//p:ph[@idx="{ph_idx}"]]')
    xps.append('//p:sp[.//p:ph]')
    for xp in xps:
        nodes = tree.xpath(xp, namespaces=ns_ppt)
        if nodes:
            return nodes[0]
    return None

def effective_size_points(sz_pt, shape, slide_path, ppt_root, ph_type, ph_idx):
    if sz_pt is None:
        return None, None

    scale = get_autofit_scale_shape(shape)
    if scale is not None:
        return (round(sz_pt * scale) if scale > 0 else None), "shape"

    layout_path = get_linked_file(slide_path, "slideLayout", ppt_root)
    if layout_path and os.path.exists(layout_path):
        layout_tree = etree.parse(layout_path)
        if ph_type or ph_idx:
            sp = find_placeholder_node(layout_tree, ph_type, ph_idx)
            if sp is not None:
                scale = get_autofit_scale_shape(sp)
                if scale is not None:
                    return (round(sz_pt * scale) if scale > 0 else None), "layout"

    master_path = get_linked_file(layout_path, "slideMaster", ppt_root) if layout_path else None
    if not master_path:
        master_path = get_linked_file(slide_path, "slideMaster", ppt_root)
    if master_path and os.path.exists(master_path):
        master_tree = etree.parse(master_path)
        if (ph_type or ph_idx):
            sp = find_placeholder_node(master_tree, ph_type, ph_idx)
            if sp is not None:
                scale = get_autofit_scale_shape(sp)
                if scale is not None:
                    return (round(sz_pt * scale) if scale > 0 else None), "master"

    return None, None

def process_ppt_folder_to_records(ppt_root):
    slides_dir = os.path.join(ppt_root, "slides")
    if not os.path.isdir(slides_dir):
        raise RuntimeError(f"slides/ not found under {ppt_root}")

    slide_files = sorted(
        [os.path.join(slides_dir, f) for f in os.listdir(slides_dir) if re.match(r'slide\d+\.xml$', f)],
        key=slide_num_from_path
    )
    if not slide_files:
        raise RuntimeError(f"No slide*.xml under {slides_dir}")

    records = []
    for slide_path in slide_files:
        slide_tree = etree.parse(slide_path)
        slide_no = slide_num_from_path(slide_path)

        for shape_idx, shape in enumerate(slide_tree.xpath('//p:sp', namespaces=ns_ppt), start=1):
            ph_type = (shape.xpath('.//p:ph/@type', namespaces=ns_ppt) or [None])[0]
            ph_idx  = (shape.xpath('.//p:ph/@idx',  namespaces=ns_ppt) or [None])[0]

            paragraphs = shape.xpath('.//a:p', namespaces=ns_ppt)
            for para_idx, p in enumerate(paragraphs, start=1):
                text_content = ''.join(p.xpath('.//a:t/text()', namespaces=ns_ppt)).strip()
                if not text_content:
                    continue

                lvl_attr = p.xpath('./a:pPr/@lvl', namespaces=ns_ppt)
                para_level = int(lvl_attr[0]) if lvl_attr else 0

                font = []; size = []; bold = []; italic = []; underline = []

                f = p.xpath('.//a:r/a:rPr/a:latin/@typeface', namespaces=ns_ppt)
                s = p.xpath('.//a:r/a:rPr/@sz',               namespaces=ns_ppt)
                b = p.xpath('.//a:r/a:rPr/@b',                namespaces=ns_ppt)
                i = p.xpath('.//a:r/a:rPr/@i',                namespaces=ns_ppt)
                u = p.xpath('.//a:r/a:rPr/@u',                namespaces=ns_ppt)
                font      = set_if_missing(font,      first_nonempty(f))
                size      = set_if_missing(size,      first_nonempty(s))
                bold      = set_bool_if_missing(bold, first_nonempty(b))
                italic    = set_bool_if_missing(italic,first_nonempty(i))
                underline = set_if_missing(underline, first_nonempty(u))

                r = p.xpath('./a:pPr/a:defRPr', namespaces=ns_ppt)
                if r:
                    f,s,b,i,u = extract_from_rPr(r[0])
                    font      = set_if_missing(font,      f)
                    size      = set_if_missing(size,      s)
                    bold      = set_bool_if_missing(bold, b)
                    italic    = set_bool_if_missing(italic,i)
                    underline = set_if_missing(underline, u)

                r = shape.xpath(f'.//a:lstStyle/a:lvl{para_level + 1}pPr/a:defRPr', namespaces=ns_ppt)
                if r:
                    f,s,b,i,u = extract_from_rPr(r[0])
                    font      = set_if_missing(font,      f)
                    size      = set_if_missing(size,      s)
                    bold      = set_bool_if_missing(bold, b)
                    italic    = set_bool_if_missing(italic,i)
                    underline = set_if_missing(underline, u)

                if (ph_type or ph_idx):
                    f,s,b,i,u = trace_to_layout_master(slide_path, ppt_root, ph_type, ph_idx, para_level)
                    font      = set_if_missing(font,      f)
                    size      = set_if_missing(size,      s)
                    bold      = set_bool_if_missing(bold, b)
                    italic    = set_bool_if_missing(italic,i)
                    underline = set_if_missing(underline, u)

                if not (font and size and bold and italic and underline):
                    layout_path = get_linked_file(slide_path, "slideLayout", ppt_root)
                    master_path = get_linked_file(layout_path, "slideMaster", ppt_root) if layout_path else None
                    if not master_path:
                        master_path = get_linked_file(slide_path, "slideMaster", ppt_root)
                    if master_path and os.path.exists(master_path):
                        master_tree = etree.parse(master_path)
                        f,s,b,i,u = find_text_styles_in_master(master_tree, ph_type, para_level)
                        font      = set_if_missing(font,      f)
                        size      = set_if_missing(size,      s)
                        bold      = set_bool_if_missing(bold, b)
                        italic    = set_bool_if_missing(italic,i)
                        underline = set_if_missing(underline, u)

                ft = first_nonempty(font)
                if ft and isinstance(ft, str) and ft.startswith('+'):
                    ft = resolve_theme_font(ppt_root, ft)
                ft_family = ui_family(ft) if ft else None

                sz_raw = first_nonempty(size)
                sz_pt  = as_pt(sz_raw) if sz_raw is not None else None
                eff_pt, eff_src = effective_size_points(sz_pt, shape, slide_path, ppt_root, ph_type, ph_idx)

                bd = as_bool01(first_nonempty(bold))
                it = as_bool01(first_nonempty(italic))
                ul = first_nonempty(underline)

                records.append({
                    "page": slide_no,  # normalised name: page for both ppt and docx
                    "kind": "pptx",
                    "shape_index": shape_idx,
                    "paragraph_index": para_idx,
                    "text": text_content,
                    "level": para_level,
                    "placeholder_type": ph_type or "",
                    "placeholder_idx": ph_idx or "",
                    "font_face": ft or "",
                    "font_family_ui": ft_family or "",
                    "size_pt_nominal": sz_pt if sz_pt is not None else "",
                    "size_pt_effective": round(eff_pt, 1) if eff_pt is not None else "",
                    "size_effective_source": eff_src or "",
                    "bold": "" if bd is None else bd,
                    "italic": "" if it is None else it,
                    "underline": ul or ""
                })
    return records



# =============================================================================
# PowerPoint template checker based on Lameco presentation template
# =============================================================================
from collections import defaultdict

def _parse_int_list(val, fallback=None):
    if isinstance(val, (list, tuple)):
        out = []
        for x in val:
            try:
                out.append(int(round(float(x))))
            except Exception:
                pass
        return set(out) if out else (set(fallback) if fallback else set())
    if isinstance(val, str):
        parts = re.split(r"[;,\s]+", val.strip())
        out = []
        for p in parts:
            if not p:
                continue
            try:
                out.append(int(round(float(p))))
            except Exception:
                pass
        return set(out) if out else (set(fallback) if fallback else set())
    return set(fallback) if fallback else set()


def _ppt_slide_size_in(ppt_root):
    """Return (width_in, height_in) from ppt/presentation.xml, or (None, None)."""
    try:
        pres = os.path.join(ppt_root, "presentation.xml")
        tree = etree.parse(pres)
        node = tree.xpath("//p:sldSz", namespaces=ns_ppt)
        if not node:
            return None, None
        node = node[0]
        cx = node.get("cx")
        cy = node.get("cy")
        if not cx or not cy:
            return None, None
        # EMU -> inches
        return (int(cx) / 914400.0), (int(cy) / 914400.0)
    except Exception:
        return None, None


def check_ppt_template(records, ppt_root=None):
    """
    records: list of dicts from process_ppt_folder_to_records()
    
    Robust check logic imported from app_multi_preview_scroll_big_update.py
    but using SETTINGS for checking values.
    """
    df = pd.DataFrame(records)
    errors = []
    record_error_map = defaultdict(list)

    if df.empty:
        return [], {}

    # numeric pages if possible
    df["page"] = pd.to_numeric(df.get("page"), errors="coerce")
    
    # --- Load Settings with safe defaults ---
    ppt_cfg = (SETTINGS.get("pptx") if isinstance(SETTINGS, dict) else {}) or {}
    
    # 1. Fonts
    expected_font = str(ppt_cfg.get("global_font") or "Montserrat").strip()
    
    # 2. Sizes
    # fallback default allowed sizes from old app
    default_allowed = [16, 18, 20, 24, 28, 32, 40, 42, 44, 48, 49, 54, 60, 72]
    
    def pars_sizes(val, fallback):
        out = set()
        if isinstance(val, (list, tuple)):
            for x in val:
                try:
                    out.add(int(round(float(x))))
                except Exception:
                    pass
        return out if out else set(fallback)

    allowed_sizes = pars_sizes(ppt_cfg.get("allowed_sizes_pt"), default_allowed)
    
    # 3. Slide specific
    enable_slide_specific = bool(ppt_cfg.get("enable_slide_specific_rules", True))
    
    def to_int(val, default):
        try:
            return int(round(float(val)))
        except Exception:
            return default

    agenda_label_text = str(ppt_cfg.get("agenda_label_text", "Agenda")).strip()
    agenda_label_size = to_int(ppt_cfg.get("agenda_label_size_pt"), 49)
    
    # Agenda Item
    agenda_item_size = to_int(ppt_cfg.get("agenda_item_size_pt"), 28)
    
    # Agenda Index (numbers)
    agenda_index_size = to_int(ppt_cfg.get("agenda_index_size_pt"), 44)
    
    # Closing
    closing_text = str(ppt_cfg.get("closing_tagline_text", "/ digitaal denken/ digitaal doen")).strip()
    closing_size = to_int(ppt_cfg.get("closing_tagline_size_pt"), 60)
    
    # Titles
    allowed_title_sizes_list = ppt_cfg.get("allowed_title_sizes_pt")
    if not allowed_title_sizes_list:
        allowed_title_sizes = {54, 60, 72}
    else:
        allowed_title_sizes = pars_sizes(allowed_title_sizes_list, [54, 60, 72])
        
    # [FIX] Ensure UI-configured strictly specific sizes are treated as valid
    conf_title_sz = to_int(ppt_cfg.get("title_size_pt"), None)
    if conf_title_sz:
        allowed_title_sizes.add(conf_title_sz)
        allowed_sizes.add(conf_title_sz)
        
    # If the user has defined "allowed_title_sizes", we must ensure these are considered valid "sizes" globally
    # to prevent "Font size X is not part of the template scale" errors on titles.
    if allowed_title_sizes:
        allowed_sizes.update(allowed_title_sizes)

    conf_body_sz = to_int(ppt_cfg.get("body_size_pt"), None)
    if conf_body_sz:
        allowed_sizes.add(conf_body_sz)
        
    # Also help prevent false positives for other configured structural elements
    for sz in (agenda_label_size, agenda_item_size, agenda_index_size, closing_size):
        if sz: allowed_sizes.add(sz)
        
    try: cover_title_size = max(allowed_title_sizes)
    except: cover_title_size = 72

    def add_error(code, message, row_idx, page):
        text_val = ""
        if "text" in df.columns and row_idx is not None and row_idx in df.index:
            text_val = str(df.loc[row_idx, "text"])
        err = {
            "code": code,
            "page": int(page) if pd.notna(page) else None,
            "row_index": int(row_idx) if row_idx is not None else None,
            "message": message,
            "text": text_val,
        }
        errors.append(err)
        if row_idx is not None:
            record_error_map[int(row_idx)].append(code)

    # ---------- 1) Font family ----------
    if "font_family_ui" in df.columns:
        # Check if font starts with expected font (case insensitive)
        bad = ~df["font_family_ui"].fillna("").astype(str).str.strip().str.lower().str.startswith(expected_font.lower())
        for i in df.index[bad]:
            page = df.loc[i, "page"]
            face = df.loc[i, "font_family_ui"]
            add_error(
                "FONT_FAMILY",
                f"Use {expected_font} for all text (found '{face}'). Change this text box to {expected_font}.",
                i,
                page,
            )

    # ---------- 2) Allowed font sizes ----------
    if "size_pt_nominal" in df.columns:
        allowed_sorted = ", ".join(str(x) for x in sorted(allowed_sizes))
        for i in df.index:
            size = df.loc[i, "size_pt_nominal"]
            if pd.notna(size):
                try:
                    s = int(round(float(size)))
                except Exception:
                    continue
                if s not in allowed_sizes:
                    page = df.loc[i, "page"]
                    add_error(
                        "FONT_SIZE",
                                                                                             f"Font size {size} pt is not part of the template scale. Change it to one of: {allowed_sorted} pt.",
                        i,
                        page,
                    )

    # ---------- 3) Agenda slide (page 2) ----------
    if enable_slide_specific and 2 in df["page"].dropna().unique():
        df2 = df[df["page"] == 2]
        # Normalize text for comparison
        tnorm = df2["text"].astype(str).str.strip()

        # "Agenda" label check
        # We look for exact match case-insensitive
        agenda_label = df2[tnorm.str.lower() == agenda_label_text.lower()]
        if agenda_label.empty:
            any_idx = df2.index[0] if not df2.empty else None
            add_error(
                "AGENDA_LABEL_MISSING",
                f"Slide 2 must contain the label '{agenda_label_text}' in the middle column.",
                any_idx,
                2,
            )
        else:
            for i in agenda_label.index:
                size = df2.loc[i, "size_pt_nominal"]
                bold_val = df2.loc[i, "bold"]
                
                try: s_int = int(round(float(size))) 
                except: s_int = 0
                
                ok = (s_int == agenda_label_size) and bool(bold_val)
                if not ok:
                    add_error(
                        "AGENDA_LABEL_STYLE",
                        f"The '{agenda_label_text}' label must be {expected_font} {agenda_label_size} pt, bold. Adjust the style.",
                        i,
                        2,
                    )

        # Left-column agenda items = non-empty, non-number, not 'Agenda'
        for i, row in df2.iterrows():
            t = str(row["text"]).strip()
            if not t or t.lower() == agenda_label_text.lower() or t.isdigit():
                continue
            size = row["size_pt_nominal"]
            try: s_int = int(round(float(size))) 
            except: s_int = 0
            
            if s_int != agenda_item_size:
                add_error(
                    "AGENDA_ITEM_SIZE",
                    f"Agenda item '{t}' should be {expected_font} {agenda_item_size} pt SemiBold. Fix the size.",
                    i,
                    2,
                )

        # Right-column numbers 1..N = agenda_index_size pt bold
        df_nums = df2[tnorm.str.fullmatch(r"\d+")]
        for i in df_nums.index:
            size = df2.loc[i, "size_pt_nominal"]
            bold_val = df2.loc[i, "bold"]
            
            try: s_int = int(round(float(size))) 
            except: s_int = 0
            
            ok = (s_int == agenda_index_size) and bool(bold_val)
            if not ok:
                add_error(
                    "AGENDA_INDEX_STYLE",
                    f"Agenda index numbers must be {expected_font} {agenda_index_size} pt, bold. Update their style.",
                    i,
                    2,
                )

    # ---------- 4) Title placeholders ----------
    if "placeholder_type" in df.columns:
        df_title = df[df["placeholder_type"].isin(["title", "ctrTitle"])]
        if not df_title.empty:
            first_title_page = df_title["page"].min()
            for i, row in df_title.iterrows():
                size = row["size_pt_nominal"]
                bold_val = row["bold"]
                page = row["page"]
                try: s = int(round(float(size))) 
                except: s = None

                if s is not None and s not in allowed_title_sizes:
                    sorted_titles = ", ".join(str(x) for x in sorted(allowed_title_sizes))
                    add_error(
                        "TITLE_SIZE",
                        f"Titles must be {sorted_titles} pt. Change this title from {size} pt.",
                        i,
                        page,
                    )

                # bold: everything except cover (first title with cover_title_size)
                if not bool(bold_val):
                    if not (page == first_title_page and s == cover_title_size):
                        add_error(
                            "TITLE_BOLD",
                            f"Titles must be bold (except the {cover_title_size} pt cover title). Make this title bold.",
                            i,
                            page,
                        )

    # ---------- 5) Closing slide tagline ----------
    if enable_slide_specific:
        last_page = df["page"].max()
        df_last = df[df["page"] == last_page]
        if not df_last.empty:
            tnorm_last = (
                df_last["text"]
                .astype(str)
                .str.replace("\u00a0", " ")
                .str.replace("\n", " ")
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
                .str.lower()
            )
            target = closing_text.lower()
            tag_idx = df_last[tnorm_last == target].index
    
            if len(tag_idx) == 0:
                any_idx = df_last.index[0]
                add_error(
                    "CLOSING_TAGLINE_MISSING",
                    f"Last slide must contain '{closing_text}' as the closing tagline.",
                    any_idx,
                    last_page,
                )
            else:
                for i in tag_idx:
                    size = df_last.loc[i, "size_pt_nominal"]
                    bold_val = df_last.loc[i, "bold"]
                    
                    try: s_int = int(round(float(size))) 
                    except: s_int = 0
                    
                    ok = (s_int == closing_size) and bool(bold_val)
                    if not ok:
                        add_error(
                            "CLOSING_TAGLINE_STYLE",
                            f"Closing tagline must be {expected_font} {closing_size} pt, bold. Fix the style on the last slide.",
                            i,
                            last_page,
                        )
                        
    # Slide size check removed per user request

    return errors, dict(record_error_map)
def _load_theme_fonts(docx_zip):
    font_map = {}
    try:
        theme_bytes = docx_zip.read('word/theme/theme1.xml')
    except KeyError:
        return font_map
    tree = ET.fromstring(theme_bytes)
    ns_a = {'a': NS_A}
    minor_latin = tree.find('.//a:fontScheme/a:minorFont/a:latin', ns_a)
    major_latin = tree.find('.//a:fontScheme/a:majorFont/a:latin', ns_a)
    if minor_latin is not None:
        font_map['minorHAnsi'] = minor_latin.get('typeface')
        font_map['minorAscii'] = minor_latin.get('typeface')
    if major_latin is not None:
        font_map['majorHAnsi'] = major_latin.get('typeface')
        font_map['majorAscii'] = major_latin.get('typeface')
    return font_map

def _parse_rpr_docx(rPr, theme_fonts):
    props = {}
    if rPr is None:
        return props
    rFonts = rPr.find('w:rFonts', ns_docx)
    if rFonts is not None:
        font_name = rFonts.get(f'{{{NS_W}}}ascii') or rFonts.get(f'{{{NS_W}}}hAnsi')
        if font_name:
            props['FontName'] = font_name
        else:
            theme_key = rFonts.get(f'{{{NS_W}}}asciiTheme') or rFonts.get(f'{{{NS_W}}}hAnsiTheme')
            if theme_key and theme_key in theme_fonts:
                props['FontName'] = theme_fonts[theme_key]
    sz = rPr.find('w:sz', ns_docx)
    if sz is not None:
        val = sz.get(f'{{{NS_W}}}val')
        if val and val.isdigit():
            props['FontSizePt'] = int(val) / 2
    if rPr.find('w:b', ns_docx) is not None:
        props['Bold'] = True
    if rPr.find('w:i', ns_docx) is not None:
        props['Italic'] = True
    if rPr.find('w:u', ns_docx) is not None:
        props['Underline'] = True
    color = rPr.find('w:color', ns_docx)
    if color is not None:
        cval = color.get(f'{{{NS_W}}}val')
        if cval:
            props['FontColor'] = cval
    return props

def extract_docx_features(docx_path):
    rows = []
    with zipfile.ZipFile(docx_path) as z:
        theme_fonts = _load_theme_fonts(z)
        styles_xml = z.read('word/styles.xml')
        styles_tree = ET.fromstring(styles_xml)
        style_name_map = {}
        style_run_props = {}
        style_para_props = {}
        default_p_style_id = None
        default_r_style_id = None
        for style in styles_tree.findall('w:style', ns_docx):
            sid = style.get(f'{{{NS_W}}}styleId')
            stype = style.get(f'{{{NS_W}}}type')
            is_default = style.get(f'{{{NS_W}}}default')
            name_el = style.find('w:name', ns_docx)
            if sid:
                style_name_map[sid] = name_el.get(f'{{{NS_W}}}val') if name_el is not None else None
            rPr = style.find('w:rPr', ns_docx)
            if rPr is not None:
                style_run_props[sid] = _parse_rpr_docx(rPr, theme_fonts)
            pPr = style.find('w:pPr', ns_docx)
            if pPr is not None:
                rPr_p = pPr.find('w:rPr', ns_docx)
                if rPr_p is not None:
                    style_para_props[sid] = _parse_rpr_docx(rPr_p, theme_fonts)
            if is_default in ['1', 'on', 'true']:
                if stype == 'paragraph':
                    default_p_style_id = sid
                elif stype == 'character':
                    default_r_style_id = sid

        r_default = {}
        docDefaults = styles_tree.find('w:docDefaults', ns_docx)
        if docDefaults is not None:
            rPrDefault = docDefaults.find('w:rPrDefault', ns_docx)
            if rPrDefault is not None:
                default_rPr = rPrDefault.find('w:rPr', ns_docx)
                if default_rPr is not None:
                    r_default = _parse_rpr_docx(default_rPr, theme_fonts)

        def merge_props(rPr, r_sid, p_sid):
            props = {}
            props.update(_parse_rpr_docx(rPr, theme_fonts))
            if r_sid and r_sid in style_run_props:
                for k, v in style_run_props[r_sid].items():
                    props.setdefault(k, v)
            if p_sid and p_sid in style_run_props:
                for k, v in style_run_props[p_sid].items():
                    props.setdefault(k, v)
            if p_sid and p_sid in style_para_props:
                for k, v in style_para_props[p_sid].items():
                    props.setdefault(k, v)
            if default_r_style_id and default_r_style_id in style_run_props:
                for k, v in style_run_props[default_r_style_id].items():
                    props.setdefault(k, v)
            if default_p_style_id and default_p_style_id in style_run_props:
                for k, v in style_run_props[default_p_style_id].items():
                    props.setdefault(k, v)
            for k, v in r_default.items():
                props.setdefault(k, v)
            return props

        # --- Pagination helpers ---
        def _norm_for_search(s):
            if not s:
                return ""
            s = str(s)
            s = s.replace("\u00ad", "")  # soft hyphen
            s = re.sub(r"\s+", " ", s).strip().lower()
            s = re.sub(r"[^\w\s]", "", s)
            s = re.sub(r"\s+", " ", s).strip()
            return s

        def _iter_run_text_segments(run, page_num_holder):
            """
            Yield (page_number, text_segment) pairs for a run, respecting:
              - explicit page breaks: <w:br w:type="page"/>
              - calculated page breaks from Word pagination: <w:lastRenderedPageBreak/>
            We keep ordering inside the run, so a page-break occurring between
            <w:t> nodes flips the page for subsequent text.
            """
            seg = []
            cur_page = int(page_num_holder[0])
            # iterate in document order
            for node in list(run):
                tag = node.tag.split('}')[-1]
                if tag == "rPr":
                    continue
                if tag == "t":
                    seg.append(node.text or "")
                    continue
                if tag == "tab":
                    seg.append("\t")
                    continue
                if tag == "br":
                    br_type = node.get(f'{{{NS_W}}}type')
                    if br_type == "page":
                        if seg:
                            yield cur_page, "".join(seg)
                            seg = []
                        page_num_holder[0] = int(page_num_holder[0]) + 1
                        cur_page = int(page_num_holder[0])
                        continue
                    # treat non-page breaks as plain whitespace to keep text searchable
                    seg.append("\n")
                    continue
                if tag == "lastRenderedPageBreak":
                    if seg:
                        yield cur_page, "".join(seg)
                        seg = []
                    page_num_holder[0] = int(page_num_holder[0]) + 1
                    cur_page = int(page_num_holder[0])
                    continue
                # ignore other nodes (drawing, footnoteRef, etc.)
            if seg:
                yield cur_page, "".join(seg)

        def _extract_pdf_page_texts(pdf_path):
            """
            Return list of page texts (1 item per page).
            Tries pypdf/PyPDF2 first, then PyMuPDF, then pdftotext if available.
            """
            # 1) pypdf / PyPDF2
            try:
                try:
                    from pypdf import PdfReader
                except Exception:
                    from PyPDF2 import PdfReader
                reader = PdfReader(pdf_path)
                out = []
                for p in reader.pages:
                    try:
                        out.append(p.extract_text() or "")
                    except Exception:
                        out.append("")
                return out
            except Exception:
                pass

            # 2) PyMuPDF
            try:
                import fitz  # PyMuPDF
                doc = fitz.open(pdf_path)
                out = []
                for i in range(doc.page_count):
                    out.append(doc.load_page(i).get_text("text") or "")
                doc.close()
                return out
            except Exception:
                pass

            # 3) pdftotext (Poppler)
            try:
                tmp_txt = tempfile.NamedTemporaryFile(delete=False, suffix=".txt")
                tmp_txt.close()
                subprocess.run(
                    ["pdftotext", pdf_path, tmp_txt.name],
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                )
                data = Path(tmp_txt.name).read_text(encoding="utf-8", errors="ignore")
                os.unlink(tmp_txt.name)
                # Poppler separates pages with form-feed
                return data.split("\f")
            except Exception:
                try:
                    if "tmp_txt" in locals() and os.path.exists(tmp_txt.name):
                        os.unlink(tmp_txt.name)
                except Exception:
                    pass
                return []

        def _find_para_page(para_norm, pdf_pages_norm, start_page_idx):
            if not para_norm:
                return start_page_idx
            words = para_norm.split()
            # phrase match first (in order of decreasing strictness)
            for n in (14, 10, 7, 5):
                phrase = " ".join(words[:n]).strip()
                if len(phrase) < 10:
                    continue
                for j in range(start_page_idx, len(pdf_pages_norm)):
                    if phrase in pdf_pages_norm[j]:
                        return j
            # token match fallback (helps with line-wrapping/hyphenation)
            tokens = [w for w in words if len(w) >= 4][:10]
            if tokens:
                needed = max(3, min(6, (len(tokens) + 1) // 2))
                for j in range(start_page_idx, len(pdf_pages_norm)):
                    page_txt = pdf_pages_norm[j]
                    hit = 0
                    for t in tokens:
                        if t in page_txt:
                            hit += 1
                            if hit >= needed:
                                return j
            return start_page_idx

        # We collect paragraph spans so we can do a PDF-based page mapping if
        # the DOCX doesn't contain soft page-break markers.
        para_spans = []  # [(start_idx, end_idx, para_norm, location)]

        def process_container(elem, location, page_num_holder):
            for child in elem:
                tag = child.tag.split('}')[-1]
                if tag == 'p':
                    p = child
                    p_style = None
                    pPr = p.find('w:pPr', ns_docx)
                    if pPr is not None:
                        pStyle = pPr.find('w:pStyle', ns_docx)
                        if pStyle is not None:
                            p_style = pStyle.get(f'{{{NS_W}}}val')
                    if not p_style:
                        p_style = default_p_style_id

                    para_row_start = len(rows)
                    para_text_acc = []

                    for run in p.findall('w:r', ns_docx):
                        rPr = run.find('w:rPr', ns_docx)
                        r_style = None
                        if rPr is not None:
                            rStyle = rPr.find('w:rStyle', ns_docx)
                            if rStyle is not None:
                                r_style = rStyle.get(f'{{{NS_W}}}val')
                        if not r_style:
                            r_style = default_r_style_id

                        # Yield 1+ segments if this run contains page breaks.
                        for seg_page, seg_text in _iter_run_text_segments(run, page_num_holder):
                            if seg_text:
                                props = merge_props(rPr, r_style, p_style)
                                rows.append({
                                    "page": int(seg_page),
                                    "kind": "docx",
                                    "location": location,
                                    "paragraph_style": style_name_map.get(p_style, p_style),
                                    "run_style": style_name_map.get(r_style, r_style),
                                    "text": seg_text,
                                    "font_name": props.get('FontName'),
                                    "font_size_pt": props.get('FontSizePt'),
                                    "bold": props.get('Bold', False),
                                    "italic": props.get('Italic', False),
                                    "underline": props.get('Underline', False),
                                    "font_color": props.get('FontColor'),
                                })
                                if location in ("body", "table"):
                                    para_text_acc.append(seg_text)

                    para_row_end = len(rows)
                    if location in ("body", "table") and para_row_end > para_row_start:
                        para_text = "".join(para_text_acc).strip()
                        para_norm = _norm_for_search(para_text)
                        if para_norm:
                            para_spans.append((para_row_start, para_row_end, para_norm, location))
                elif tag == 'tbl':
                    process_container(child, 'table', page_num_holder)
                else:
                    process_container(child, location, page_num_holder)

        # --- Parse main document ---
        document_xml = z.read('word/document.xml')
        doc_tree = ET.fromstring(document_xml)
        body = doc_tree.find('w:body', ns_docx)
        page_num_holder = [1]
        process_container(body, 'body', page_num_holder)

        # If we still have only page=1 for all body/table content, try PDF-based mapping.
        try:
            pages_seen = {int(r.get("page", 1)) for r in rows if r.get("kind") == "docx" and r.get("location") in ("body", "table")}
        except Exception:
            pages_seen = {1}

        if pages_seen == {1} and para_spans:
            tmp_dir = tempfile.mkdtemp(prefix="docx_pag_")
            pdf_path = None
            try:
                pdf_cmd = [
                    "soffice",
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", tmp_dir,
                    docx_path,
                ]
                subprocess.run(pdf_cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                pdf_candidates = sorted(glob.glob(os.path.join(tmp_dir, "*.pdf")))
                pdf_path = pdf_candidates[0] if pdf_candidates else None

                if pdf_path and os.path.exists(pdf_path):
                    pdf_page_texts = _extract_pdf_page_texts(pdf_path)
                    pdf_pages_norm = [_norm_for_search(t) for t in (pdf_page_texts or [])]
                    if pdf_pages_norm:
                        cur_page_idx = 0
                        for (s, e, pnorm, loc) in para_spans:
                            found_idx = _find_para_page(pnorm, pdf_pages_norm, cur_page_idx)
                            # clamp to valid range
                            if found_idx < 0:
                                found_idx = 0
                            if found_idx >= len(pdf_pages_norm):
                                found_idx = len(pdf_pages_norm) - 1
                            page_no = int(found_idx) + 1
                            for i in range(s, e):
                                # only override body/table rows (leave footer alone)
                                if rows[i].get("location") in ("body", "table"):
                                    rows[i]["page"] = page_no
                            if found_idx > cur_page_idx:
                                cur_page_idx = found_idx
            except Exception:
                pass
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)

        # --- Parse footer(s) (keep footer pages stable at 1) ---
        for name in z.namelist():
            if name.startswith('word/footer') and name.endswith('.xml'):
                f_xml = z.read(name)
                f_tree = ET.fromstring(f_xml)
                process_container(f_tree, 'footer', [1])

    return rows



# =============================================================================
# Serialization helpers
# =============================================================================
def records_to_csv_bytes(records):
    df = pd.DataFrame(records)
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")

def records_to_excel_bytes(records):
    df = pd.DataFrame(records)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="features")
    buf.seek(0)
    return buf


# =============================================================================
# YOLO hook placeholder
# =============================================================================
def run_logo_detector(file_path, kind):
    """
    Hook for your YOLO model.

    - file_path: path to uploaded PPTX/DOCX file.
    - kind: 'pptx' or 'docx'.

    Return value is currently unused; adapt it later when you
    integrate real logo detection.
    """
    # TODO: replace this stub with an actual YOLO call.
    # Example of a future return:
    # [
    #   {"page": 1, "label": "logo_missing", "bbox": [x1, y1, x2, y2]},
    # ]
    return []


# =============================================================================
# Flask views
# =============================================================================
INDEX_HTML = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Lameco document Checker – feature extractor</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { font-family: ui-sans-serif,system-ui,Segoe UI,Roboto,Helvetica,Arial; margin: 0; background:#faf5ff; }
    .shell { max-width: 1040px; margin: 2rem auto; padding: 1.5rem 2rem; background:#fff; border-radius: 18px;
             box-shadow: 0 10px 30px rgba(15,23,42,.12); border:1px solid #e5e7eb; }
    h1 { margin: 0 0 .25rem; font-size: 1.6rem; }
    h2 { margin-top: 1.5rem; font-size: 1.1rem; }
    .muted { color:#6b7280; font-size:.9rem; }
    .upload-box { margin-top:1.25rem; padding:1.5rem; border-radius:16px; border:2px dashed #e5e7eb;
                  display:flex; flex-direction:column; align-items:center; justify-content:center; background:#f9fafb; }
    .btn { display:inline-flex; align-items:center; justify-content:center; border-radius:999px;
           padding:.55rem 1.2rem; font-size:.9rem; border:none; cursor:pointer; text-decoration:none; }
    .btn-primary { background:#b91c1c; color:#fff; }
    .btn-secondary { background:#fee2e2; color:#991b1b; }
    .btn-ghost { background:transparent; color:#4b5563; }
    .btn[disabled] { opacity:.5; cursor:default; }
    .segmented { display:inline-flex; border-radius:999px; background:#f3f4f6; padding:3px; }
    .segmented button { border:none; background:transparent; padding:.3rem .9rem; border-radius:999px;
                        font-size:.85rem; cursor:pointer; color:#4b5563; }
    .segmented button.active { background:#b91c1c; color:#fff; }
    .row { margin-top:1rem; }
    .flash-ok { color:#065f46; font-size:.9rem; }
    .flash-err { color:#991b1b; font-size:.9rem; }
    table { width:100%; border-collapse:collapse; font-size:.8rem; margin-top:.75rem; }
    th, td { padding:.35rem .4rem; border-bottom:1px solid #e5e7eb; text-align:left; }
    th { background:#f9fafb; font-weight:600; }
    .pill { display:inline-block; padding:.15rem .55rem; border-radius:999px; font-size:.7rem; background:#fee2e2; color:#b91c1c; }
    .downloads { margin-top:1rem; display:flex; gap:.5rem; }
  
    .error-summary {
      margin-top: 1.25rem;
      padding: 1rem 1.25rem;
      background:#fef2f2;
      border-radius: 14px;
      border:1px solid #fecaca;
    }
    .error-summary h3 { margin:0 0 .5rem; font-size:1.05rem; }
    .error-list { list-style:none; margin:0; padding:0; }
    .error-list li { margin-bottom:.4rem; cursor:pointer; }
    .error-pill {
      display:inline-block;
      font-size:.75rem;
      padding:0 .5rem;
      border-radius:999px;
      background:#b91c1c;
      color:#fff;
      margin-right:.4rem;
    }
    tr.has-error { background:#fef2f2; }
    tr.row-highlight { outline:2px solid #b91c1c; }

    .slide-tabs {
      margin: 1rem 0 .5rem;
      display: flex;
      flex-wrap: wrap;
      gap: .5rem;
    }
    .slide-tab {
      padding: .25rem .75rem;
      border-radius: 999px;
      border: 1px solid #e5e7eb;
      background: #f9fafb;
      cursor: pointer;
      font-size: .85rem;
    }
    .slide-tab.active {
      background: #111827;
      color: #fff;
      border-color: #111827;
    }

  </style>
</head>
<body>
  <div class="shell">
    <h1>Lameco document Checker</h1>
    <p class="muted">Step 1 · Upload a PowerPoint or Word file and extract all text + formatting features. Issues & YOLO logo checks come in the next step.</p>

    <form method="post" enctype="multipart/form-data" action="{{ url_for('upload') }}">
      <div class="row">
        <div class="segmented">
          <button type="button" class="active" id="btn-word">WORD</button>
          <button type="button" id="btn-ppt">POWERPOINT</button>
        </div>
      </div>
      <div class="upload-box">
        <p><strong>Upload your document</strong></p>
        <p class="muted">Drag &amp; drop here or click to browse.</p>
        <input type="file" name="file" id="file-input" style="margin-top:.75rem;" required>
        <input type="hidden" name="doc_kind_hint" id="doc-kind-hint" value="docx">
      </div>
      <div class="row">
        <button class="btn btn-primary" type="submit">Extract features</button>
      </div>
    </form>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% if messages %}
        <div class="row">
          {% for category, msg in messages %}
            <div class="{{ 'flash-ok' if category=='success' else 'flash-err' }}">{{ msg|safe }}</div>
          {% endfor %}
        </div>
      {% endif %}
    {% endwith %}

    {% if job %}
      <hr class="row">
      <h2>Step 2 · Review extracted content</h2>
      <p class="muted">
        Parsed {{ job['count'] }} text items from <code>{{ job['source'] }}</code>.
        <span class="pill">{{ job['kind']|upper }}</span>
      </p>

      <div class="downloads">
        <a class="btn btn-secondary" href="{{ url_for('download_result', fmt='csv', token=job['token']) }}">Download CSV</a>
        <a class="btn btn-secondary" href="{{ url_for('download_result', fmt='xlsx', token=job['token']) }}">Download Excel</a>
        <a class="btn btn-ghost" href="{{ url_for('index') }}">New upload</a>
      </div>

      <p class="muted" style="margin-top:1rem;">
        Preview below is a simple table. Later you can replace this with the page preview + checkbox UI
        from your Figma prototype. Each row is one text run / paragraph your rule engine can attach issues to.
      </p>

      {% if job['kind'] == 'pptx' %}
        {% if job.get('template_errors') %}
          <div class="error-summary">
            <h3>Template issues detected</h3>
            <ul class="error-list">
              {% for e in job['template_errors'] %}
                <li data-row-index="{{ e['row_index'] }}" data-page="{{ e['page'] }}">
                  <span class="error-pill">{{ e['code'] }}</span>
                  <strong>Slide {{ e['page'] }}</strong> – {{ e['message'] }}
                  {% if e['text'] %}
                    <br><span class="muted">Text: "{{ e['text'] }}"</span>
                  {% endif %}
                </li>
              {% endfor %}
            </ul>
            <p class="muted" style="margin-top:.5rem;">
              Tip: click an issue to jump to the corresponding row in the table.
            </p>
          </div>
        {% else %}
          <p class="muted" style="margin:.75rem 0 0;">
            No template issues detected for this presentation (based on the current rules).
          </p>
        {% endif %}
      {% endif %}

      <div id="slide-tabs" class="slide-tabs"></div>

      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>Page / Slide</th>
            <th>Text</th>
            <th>Font</th>
            <th>Size</th>
            <th>B/I/U</th>
          </tr>
        </thead>
        <tbody>
        {% for rec in job['preview_rows'] %}
          <tr class="{% if rec.get('_errors') %}has-error{% endif %}"
              data-row-index="{{ rec.get('_idx') }}"
              data-page="{{ rec.get('page','') }}">
            <td>
              {{ loop.index }}
              {% if rec.get('_errors') %}
                <span class="error-pill">!</span>
              {% endif %}
            </td>
            <td>{{ rec.get('page','?') }}</td>
            <td>{{ rec.get('text','') }}</td>
            <td>{{ rec.get('font_family_ui', rec.get('font_name','')) }}</td>
            <td>{{ rec.get('size_pt_nominal', rec.get('font_size_pt','')) }}</td>
            <td>
              {{ 'B' if rec.get('bold') else '' }}
              {{ 'I' if rec.get('italic') else '' }}
              {{ 'U' if rec.get('underline') else '' }}
            </td>
          </tr>
        {% endfor %}
        </tbody>
      </table>

      <p class="muted" style="margin-top:.75rem;">
        YOLO hook placeholder: once you wire your model into <code>run_logo_detector</code>, you can show
        logo-related issues in a side panel here and connect them with checkboxes + page jumps.
      </p>
    {% endif %}
  </div>

  <script>
    // purely visual toggle – backend detects file type from extension anyway
    const btnWord = document.getElementById('btn-word');
    const btnPpt  = document.getElementById('btn-ppt');
    const kindHint = document.getElementById('doc-kind-hint');

    function setKind(kind) {
      if (kind === 'docx') {
        btnWord.classList.add('active');
        btnPpt.classList.remove('active');
      } else {
        btnPpt.classList.add('active');
        btnWord.classList.remove('active');
      }
      kindHint.value = kind;
    }
    btnWord.addEventListener('click', () => setKind('docx'));
    btnPpt.addEventListener('click',  () => setKind('pptx'));
  

    // Build slide tabs from table rows and enable per-slide filtering
    (function() {
      const rows = Array.from(document.querySelectorAll('tr[data-page]'));
      if (!rows.length) return;
      const pages = Array.from(new Set(
        rows
          .map(r => r.getAttribute('data-page'))
          .filter(p => p && p.trim() !== '')
      )).sort((a, b) => parseInt(a) - parseInt(b));
      const container = document.getElementById('slide-tabs');
      if (!container || !pages.length) return;

      let current = 'ALL';

      function applyFilter() {
        rows.forEach(r => {
          const p = r.getAttribute('data-page');
          if (current === 'ALL' || p === current) {
            r.style.display = '';
          } else {
            r.style.display = 'none';
          }
        });

        document.querySelectorAll('.error-summary li[data-row-index]').forEach(li => {
          const p = li.getAttribute('data-page');
          if (!p || current === 'ALL' || p === current) {
            li.style.display = '';
          } else {
            li.style.display = 'none';
          }
        });
      }

      function makeTab(label, value) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = label;
        btn.className = 'slide-tab';
        if (value === current) btn.classList.add('active');
        btn.addEventListener('click', () => {
          current = value;
          document.querySelectorAll('.slide-tab').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          applyFilter();
        });
        return btn;
      }

      container.appendChild(makeTab('All slides', 'ALL'));
      pages.forEach(p => {
        container.appendChild(makeTab('Slide ' + p, p));
      });

      applyFilter();
    })();

    // Clicking an error in the summary scrolls to the matching table row
    document.querySelectorAll('.error-summary li[data-row-index]').forEach(li => {
      li.addEventListener('click', () => {
        const idx = li.getAttribute('data-row-index');
        if (!idx) return;
        const row = document.querySelector('tr[data-row-index="' + idx + '"]');
        if (!row) return;
        row.scrollIntoView({ behavior: 'smooth', block: 'center' });
        row.classList.add('row-highlight');
        setTimeout(() => row.classList.remove('row-highlight'), 1500);
      });
    });
</script>
</body>
</html>
"""

JOBS = {}

@app.get("/")
def index():
    # When no job is active, render the template with an empty checked_rows list
    return render_template_string(INDEX_HTML, job=None, checked_rows=[], logo_data_uri=LOGO_DATA_URI, settings=SETTINGS)

@app.post("/upload")
def upload():
    f = request.files.get("file")
    if not f or f.filename == "":
        flash("No file provided.", "error")
        return redirect(url_for("index"))

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXTS:
        flash("Unsupported file type. Upload a .pptx, .docx or a .zip containing a PPTX 'ppt/' folder.", "error")
        return redirect(url_for("index"))

    tmp_root = tempfile.mkdtemp(prefix="docjob_")
    src_path = os.path.join(tmp_root, "upload" + ext)
    f.save(src_path)

    kind = None
    records = []
    ppt_root = None

    try:
        if ext in (".pptx", ".zip"):
            # unzip pptx/zip into a workdir and process the ppt/ folder
            workdir = os.path.join(tmp_root, "unzipped")
            os.makedirs(workdir, exist_ok=True)
            with zipfile.ZipFile(src_path) as zf:
                zf.extractall(workdir)
            ppt_root = os.path.join(workdir, "ppt")
            if not os.path.isdir(ppt_root):
                candidate = None
                for root, dirs, files in os.walk(workdir):
                    if os.path.basename(root) == "ppt":
                        candidate = root
                        break
                ppt_root = candidate if candidate else ppt_root
            if not os.path.isdir(ppt_root):
                raise RuntimeError("Could not find a 'ppt/' folder inside the archive.")
            records = process_ppt_folder_to_records(ppt_root)
            kind = "pptx"
        elif ext == ".docx":
            records = extract_docx_features(src_path)
            kind = "docx"
        else:
            raise RuntimeError("Unexpected extension handling branch.")
    except Exception as e:
        shutil.rmtree(tmp_root, ignore_errors=True)
        flash(f"Extraction failed: {e}", "error")
        return redirect(url_for("index"))

        # Optional YOLO call – currently a no-op, but this is where it belongs.
    try:
        _logo_issues = run_logo_detector(src_path, kind)
    except Exception:
        _logo_issues = []

    # Template checks (PPTX or DOCX)
    template_errors = []
    record_error_map = {}
    errors_per_page = {}
    if kind == "pptx":
        # Run PPTX-specific template checks
        try:
            template_errors, record_error_map = check_ppt_template(records, ppt_root=ppt_root)
        except Exception:
            template_errors, record_error_map = [], {}
        # count errors per slide
        for e in template_errors:
            try:
                p = e.get("page")
            except AttributeError:
                p = e.get("page") if isinstance(e, dict) else None
            if p is None:
                continue
            key = str(int(p))
            errors_per_page[key] = errors_per_page.get(key, 0) + 1
    elif kind == "docx":
        # Run DOCX-specific template checks based on the provided template document
        try:
            template_errors, record_error_map = check_docx_template(records)
        except Exception:
            template_errors, record_error_map = [], {}
        # count errors per page (always page 1 at the moment)
        for e in template_errors:
            p = e.get("page")
            if p is None:
                continue
            key = str(int(p))
            errors_per_page[key] = errors_per_page.get(key, 0) + 1

    # Generate slide thumbnails for PPTX (if LibreOffice is available)
    # and always create a slide entry for every page we saw in the records.
    slides = []
    if kind == "pptx":
        # Collect all distinct page numbers from the extracted records
        pages = sorted({r.get("page") for r in records if isinstance(r.get("page"), (int, float))})
        try:
            slides_dir = os.path.join(tmp_root, "slides")
            thumbs = render_ppt_thumbnails(src_path, slides_dir)
        except Exception:
            thumbs = []

        # Map thumbnails to pages by order; if there are fewer thumbnails than pages,
        # we still create slide entries without a filename so the UI can show them.
        for idx, page in enumerate(pages):
            filename = None
            if idx < len(thumbs):
                filename = thumbs[idx].get("filename")
            try:
                page_int = int(page)
            except Exception:
                continue
            slides.append({"page": page_int, "filename": filename})

    elif kind == "docx":
        # Collect all distinct page numbers from the extracted records
        pages = sorted({r.get("page") for r in records if isinstance(r.get("page"), (int, float))})
        try:
            slides_dir = os.path.join(tmp_root, "slides")
            thumbs = render_docx_thumbnails(src_path, slides_dir)
        except Exception:
            thumbs = []

        # If we managed to render pages but the extractor didn't find page numbers,
        # fall back to a simple 1..N page list.
        if thumbs and not pages:
            pages = list(range(1, len(thumbs) + 1))

        # If rendering produced more pages than we saw in records (e.g. blank pages),
        # extend the page list so the preview can still show them.
        if thumbs and pages:
            try:
                max_p = max(int(p) for p in pages)
            except Exception:
                max_p = 0
            if len(thumbs) > max_p:
                for p in range(max_p + 1, len(thumbs) + 1):
                    pages.append(p)

        pages = sorted(set(int(p) for p in pages if isinstance(p, (int, float))))

        # Map thumbnails to pages by order; if there are fewer thumbnails than pages,
        # we still create entries without a filename so the UI can show the page list.
        for idx, page in enumerate(pages):
            filename = None
            if idx < len(thumbs):
                filename = thumbs[idx].get("filename")
            slides.append({"page": int(page), "filename": filename})

    # Attach row index + error codes to preview rows (limit for UI)
    preview_rows = []
    for idx, rec in enumerate(records[:200]):
        r = dict(rec)
        r["_idx"] = idx
        r["_errors"] = record_error_map.get(idx, [])
        preview_rows.append(r)

    # compute a stable hash of the uploaded file for persistent checkbox state
    try:
        with open(src_path, 'rb') as _fh:
            file_hash = hashlib.md5(_fh.read()).hexdigest()
    except Exception:
        file_hash = None

    checked_rows = STATE.get(file_hash, []) if file_hash else []

    token = os.path.basename(tmp_root)
    csv_bytes = records_to_csv_bytes(records)
    xlsx_bytes = records_to_excel_bytes(records).read()

    results_dir = os.path.join(tmp_root, "results")
    os.makedirs(results_dir, exist_ok=True)
    csv_path = os.path.join(results_dir, "features.csv")
    xlsx_path = os.path.join(results_dir, "features.xlsx")
    with open(csv_path, "wb") as cf:
        cf.write(csv_bytes)
    with open(xlsx_path, "wb") as ef:
        ef.write(xlsx_bytes)

    JOBS[token] = {
        "tmp_root": tmp_root,
        "csv": csv_path,
        "xlsx": xlsx_path,
        "count": len(records),
        "source": f.filename,
        "token": token,
        "kind": kind,
        "preview_rows": preview_rows,
        "logo_issues": _logo_issues,
        "template_errors": template_errors,
        "slides": slides,
        "errors_per_page": errors_per_page,
        "file_hash": file_hash,
    }

    # Add to History (persisted across restarts)
    total_issues = len(template_errors) if template_errors else 0
    # On upload, we might already have checked rows in STATE if file hash matches
    already_checked = len(STATE.get(file_hash, []))
    remaining_issues = max(0, total_issues - already_checked)

    add_history_entry({
        "token": token,
        "file_hash": file_hash,
        "filename": f.filename,
        "kind": kind,
        "record_count": len(records),
        "total_issues": total_issues,
        "issue_count": remaining_issues,
    })

    flash("Extraction complete.", "success")
    # pass the checked_rows list to the template so rows can be pre-crossed
    return render_template_string(INDEX_HTML, job=JOBS[token], checked_rows=checked_rows, logo_data_uri=LOGO_DATA_URI, settings=SETTINGS)


@app.post("/analyze")
def analyze():
    # Backward-compat alias used by some UI revisions.
    return upload()

@app.get("/slide/<token>/<int:page>.png")
def slide_image(token, page):
    job = JOBS.get(token)
    if not job:
        flash("Job not found or expired.", "error")
        return redirect(url_for("index"))
    slides = job.get("slides") or []
    for s in slides:
        try:
            p = int(s.get("page", 0))
        except Exception:
            continue
        if p == int(page):
            path = os.path.join(job["tmp_root"], "slides", s["filename"])
            if os.path.exists(path):
                return send_file(path, mimetype="image/png")
    flash("Slide image not found (is LibreOffice installed for PPT rendering?).", "error")
    return redirect(url_for("index"))


@app.get("/download/<fmt>")
def download_result(fmt):
    token = request.args.get("token", "")
    job = JOBS.get(token)
    if not job:
        flash("Job not found or expired.", "error")
        return redirect(url_for("index"))

    if fmt == "csv":
        path = job["csv"]
        return send_file(
            path,
            mimetype="text/csv",
            as_attachment=True,
            download_name="features.csv",
        )
    elif fmt == "xlsx":
        path = job["xlsx"]
        return send_file(
            path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="features.xlsx",
        )
    else:
        flash("Unknown format.", "error")
        return redirect(url_for("index"))


# -----------------------------------------------------------------------------
# Route to toggle checkbox state and persist done rows
# -----------------------------------------------------------------------------
@app.post("/toggle")
def toggle_checkbox():
    """
    Persist the checked/unchecked state for a given row index of a file.
    Updates HISTORY issue count immediately.
    """
    data = request.get_json() or {}
    file_hash = data.get('file_hash')
    idx = data.get('index') if data.get('index') is not None else data.get('row_index')
    checked = data.get('checked')

    if not file_hash or idx is None:
        return '', 400
    
    try:
        idx_int = int(idx)
    except Exception:
        return '', 400

    # 1. Update Checkbox State
    rows = STATE.setdefault(file_hash, [])
    if checked:
        if idx_int not in rows:
            rows.append(idx_int)
    else:
        if idx_int in rows:
            rows.remove(idx_int)
    save_state()

    # 2. Update History Count
    # We iterate global HISTORY to find the matching item and update its issue_count.
    found = False
    for item in HISTORY:
        if item.get("file_hash") == file_hash:
            # Baseline total
            if "total_issues" not in item:
                # Legacy fallback: assume stored issue_count was the total
                try:
                    item["total_issues"] = int(item.get("issue_count", 0))
                except:
                    item["total_issues"] = 0
            
            # Recalculate remaining
            try:
                total = int(item["total_issues"])
                current_checked_count = len(rows)
                new_remaining = max(0, total - current_checked_count)
                item["issue_count"] = new_remaining
                found = True
            except Exception:
                pass
            break
    
    if found:
        save_history()

    return '', 204


# -----------------------------------------------------------------------------
# Settings + History API routes (JSON)
# -----------------------------------------------------------------------------
@app.get("/api/settings")
def api_get_settings():
    return SETTINGS

@app.post("/api/settings")
def api_set_settings():
    data = request.get_json(silent=True) or {}

    def _parse_int_list(val):
        if isinstance(val, list):
            out = []
            for x in val:
                try:
                    out.append(int(round(float(x))))
                except Exception:
                    pass
            return out
        if isinstance(val, str):
            parts = re.split(r"[,\s]+", val.strip())
            out = []
            for p in parts:
                if not p:
                    continue
                try:
                    out.append(int(round(float(p))))
                except Exception:
                    pass
            return out
        return []

    def _to_float(val, default=None):
        try:
            return float(val)
        except Exception:
            return default

    def _to_int(val, default=None):
        try:
            return int(round(float(val)))
        except Exception:
            return default

    try:
        # Shallow-merge to keep unknown fields from breaking
        for k in ("pptx", "docx"):
            if isinstance(data.get(k), dict):
                SETTINGS.setdefault(k, {}).update(data[k])

        # Normalize types (critical for rule checks)
        ppt = SETTINGS.setdefault("pptx", {})
        if isinstance(ppt.get("allowed_sizes_pt"), str):
            lst = _parse_int_list(ppt.get("allowed_sizes_pt"))
            if lst:
                ppt["allowed_sizes_pt"] = lst
        if isinstance(ppt.get("allowed_title_sizes_pt"), str):
            lst = _parse_int_list(ppt.get("allowed_title_sizes_pt"))
            if lst:
                ppt["allowed_title_sizes_pt"] = lst

        for k in ("slide_width_in", "slide_height_in"):
            if k in ppt:
                v = _to_float(ppt.get(k), None)
                if v is not None:
                    ppt[k] = v
        for k in ("title_size_pt", "body_size_pt", "agenda_label_size_pt", "agenda_index_size_pt", "agenda_item_size_pt", "closing_tagline_size_pt"):
            if k in ppt:
                v = _to_int(ppt.get(k), None)
                if v is not None:
                    ppt[k] = v

        # Keep global_font in sync if user only edits title/body fonts
        # Always update global_font if body_font is provided, as it is the primary UI field
        if ppt.get("body_font"):
             ppt["global_font"] = ppt.get("body_font")
        elif ppt.get("title_font") and not ppt.get("global_font"):
             ppt["global_font"] = ppt.get("title_font")

        doc = SETTINGS.setdefault("docx", {})
        for k in ("body_size_pt", "heading_size_pt", "title_size_pt", "font_size_pt"):
            if k in doc:
                v = _to_int(doc.get(k), None)
                if v is not None:
                    doc[k] = v
        for k in ("line_spacing",):
            if k in doc:
                v = _to_float(doc.get(k), None)
                if v is not None:
                    doc[k] = v

        try:
            global _WORD_TEMPLATE_RULES_CACHE
            _WORD_TEMPLATE_RULES_CACHE = None
        except:
            pass

        save_settings()
        return {"ok": True, "settings": SETTINGS}
    except Exception as e:
        return {"ok": False, "error": str(e)}, 400
@app.get("/api/history")
def api_get_history():
    # Return newest first
    return list(reversed(HISTORY[-50:]))

@app.post("/api/history/clear")
def api_clear_history():
    try:
        HISTORY.clear()
        save_history()
    except Exception:
        pass
    return {"ok": True}

@app.get("/job/<token>")
def open_job(token):
    job = JOBS.get(token)
    if not job:
        flash("That job is no longer available in memory. Please re-upload the document.", "error")
        return redirect(url_for("index"))
    checked_rows = STATE.get(job.get("file_hash") or "", [])
    return render_template_string(INDEX_HTML, job=job, checked_rows=checked_rows, logo_data_uri=LOGO_DATA_URI, settings=SETTINGS)


# -----------------------------------------------------------------------------
# Prototype-based UI template (overrides previous INDEX_HTML)
# -----------------------------------------------------------------------------
\
INDEX_HTML = '''<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Lameco | Document Format Checker</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root{
      --red:#FF343F;
      --red-700:#C81F2B;
      --ink:#141415;
      --muted:#6B7280;
      --line:#E5E7EB;
      --bg:#F7F8FA;
      --card:#FFFFFF;
      --soft:#FFF5F6;
      --shadow: 0 14px 35px rgba(2, 6, 23, .10);
      --radius: 18px;
    }
    *{box-sizing:border-box}
    body{
      margin:0;
      font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      color:var(--ink);
      background: linear-gradient(180deg, #FFFFFF, var(--bg));
    }
    a{color:inherit}
    .nav{
      height:64px;
      border-bottom:1px solid var(--line);
      background: rgba(255,255,255,.85);
      backdrop-filter: blur(10px);
      position: sticky; top:0; z-index: 20;
    }
    .nav .inner{
      max-width: 1200px;
      margin:0 auto;
      height:100%;
      padding: 0 18px;
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
    }
    .brand{
      display:flex; align-items:center; gap:10px;
      font-weight: 700;
      letter-spacing: .2px;
    }
    .brand img{height:36px}
    .pill{
      border:1px solid #FFD2D6;
      background: var(--soft);
      color: var(--red-700);
      border-radius: 999px;
      padding: 6px 10px;
      font-size: 12px;
      font-weight: 650;
      display:flex;
      align-items:center;
      gap:8px;
    }
    .wrap{
      max-width: 1200px;
      margin:0 auto;
      padding: 22px 18px 30px;
    }
    .grid{
      display:grid;
      grid-template-columns: 1fr 380px;
      gap: 18px;
      align-items:start;
    }
    @media (max-width: 980px){
      .grid{grid-template-columns:1fr}
    }
    .h1{
      font-size: 34px;
      line-height: 1.1;
      margin: 10px 0 6px;
      letter-spacing: -.6px;
    }
    .sub{
      color: var(--muted);
      margin: 0 0 14px;
      font-size: 14px;
      max-width: 70ch;
    }
    .card{
      background: var(--card);
      border:1px solid var(--line);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
    }
    .card .hd{
      padding: 14px 14px 10px;
      border-bottom: 1px solid var(--line);
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:10px;
    }
    .card .hd h3{
      margin:0;
      font-size: 13px;
      font-weight: 750;
      letter-spacing: .25px;
      text-transform: uppercase;
      color: #111827;
      display:flex;
      align-items:center;
      gap:8px;
    }
    .card .bd{padding: 14px;}
    .drop{
      border: 2px dashed #E5E7EB;
      border-radius: 18px;
      padding: 28px 18px;
      text-align:center;
      background: linear-gradient(180deg, #FFFFFF, #FAFAFB);
      position:relative;
      cursor:pointer;
    }
    .drop:hover{border-color:#FFD2D6; background: #FFF7F8;}
    .drop .icon{
      width: 44px; height:44px;
      border-radius: 14px;
      margin:0 auto 10px;
      background: var(--soft);
      display:grid; place-items:center;
      color: var(--red-700);
      font-weight: 900;
    }
    .drop h4{margin:0; font-size: 14px;}
    .drop p{margin:6px 0 0; color:var(--muted); font-size: 12px;}
    input[type=file]{display:none}
    .row{
      display:flex; gap:10px; align-items:center; flex-wrap:wrap;
    }
    .btn{
      border:1px solid var(--line);
      background: #FFF;
      padding: 10px 14px;
      border-radius: 12px;
      cursor:pointer;
      font-weight: 650;
      font-size: 13px;
      display:inline-flex; align-items:center; gap:8px;
    }
    .btn:hover{border-color:#D1D5DB}
    .btn.primary{
      background: var(--red);
      border-color: var(--red);
      color:#fff;
    }
    .btn.primary:hover{background: var(--red-700); border-color: var(--red-700);}
    .btn.ghost{background: #FFF;}
    .btn.small{padding: 8px 10px; font-size: 12px; border-radius: 10px;}
    .muted{color:var(--muted)}
    .filechip{
      border:1px solid #D1FAE5;
      background:#ECFDF5;
      border-radius: 14px;
      padding: 12px;
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
    }
    .fileleft{display:flex; align-items:center; gap:10px;}
    .fileicon{
      width:36px; height:36px; border-radius: 12px;
      background:#D1FAE5;
      display:grid; place-items:center;
      font-weight: 900;
      color:#065F46;
    }
    .filemeta{display:flex; flex-direction:column; gap:3px;}
    .badges{display:flex; gap:10px; font-size:12px;}
    .badge{display:inline-flex; gap:6px; align-items:center}
    .badge.err{color:#DC2626}
    .badge.warn{color:#D97706}
    table{
      width:100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th,td{
      padding: 10px 10px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      text-align:left;
    }
    th{
      font-size: 12px;
      color: var(--muted);
      font-weight: 750;
      text-transform: uppercase;
      letter-spacing: .25px;
      background: #FBFBFC;
    }
    tr.crossed td{
      color:#9CA3AF;
      text-decoration: line-through;
    }
    .type{
      display:inline-flex;
      align-items:center;
      padding: 2px 8px;
      border-radius:999px;
      font-size: 12px;
      font-weight: 700;
      border:1px solid var(--line);
      background:#fff;
    }
    .type.error{border-color:#FECACA; background:#FEF2F2; color:#B91C1C}
    .type.warn{border-color:#FED7AA; background:#FFF7ED; color:#9A3412}
    .type.info{border-color:#BFDBFE; background:#EFF6FF; color:#1D4ED8}
    .settings-grid{
      display:grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }
    .field label{
      display:block;
      font-size:12px;
      color: var(--muted);
      margin: 0 0 6px;
    }
    .field input, .field select{
      width:100%;
      border:1px solid var(--line);
      border-radius: 12px;
      padding: 10px 10px;
      font-size: 13px;
      outline:none;
      background:#fff;
    }
    .field input:focus, .field select:focus{border-color:#FCA5A5; box-shadow: 0 0 0 4px rgba(255,52,63,.12);}
    details{
      border:1px solid var(--line);
      border-radius: 14px;
      background:#fff;
      overflow:hidden;
    }
    summary{
      list-style:none;
      padding: 12px 12px;
      cursor:pointer;
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap: 10px;
      font-weight: 750;
      font-size: 13px;
    }
    summary::-webkit-details-marker{display:none}
    .chev{color: var(--muted)}
    details[open] summary{border-bottom: 1px solid var(--line); background:#FBFBFC;}
    .section{padding: 12px;}
    .hist{
      display:flex;
      flex-direction:column;
      gap: 10px;
    }
    .hist-item{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap: 10px;
      padding: 10px 10px;
      border:1px solid var(--line);
      border-radius: 14px;
      cursor:pointer;
      background:#fff;
    }
    .hist-item:hover{border-color:#D1D5DB}
    .hist-left{display:flex; gap:10px; align-items:center;}
    .hist-meta{display:flex; flex-direction:column; gap:2px;}
    .hist-meta .fn{font-weight: 700; font-size: 13px;}
    .hist-meta .sm{color: var(--muted); font-size: 12px;}
    .arrow{color: var(--muted); font-weight: 900;}
    .toast{
      position: fixed;
      right: 18px;
      bottom: 18px;
      background:#111827;
      color:#fff;
      padding: 10px 12px;
      border-radius: 12px;
      font-size: 13px;
      box-shadow: var(--shadow);
      opacity:0;
      transform: translateY(8px);
      transition: .18s ease;
      z-index: 50;
      max-width: 320px;
    }
    .toast.show{opacity:1; transform: translateY(0);}
    .preview{
      display:grid;
      grid-template-columns: 140px 1fr;
      gap: 12px;
      align-items:start;
    }
    .thumbs{
      display:flex;
      flex-direction:column;
      gap: 8px;
      max-height: 340px;
      overflow:auto;
      padding-right: 4px;
    }
    .thumb{
      border:1px solid var(--line);
      border-radius: 12px;
      padding: 8px;
      cursor:pointer;
      background:#fff;
      display:flex;
      align-items:center;
      gap: 8px;
    }
    .thumb.active{border-color:#FCA5A5; box-shadow: 0 0 0 4px rgba(255,52,63,.10);}
    .thumb .p{font-weight:800; font-size: 12px;}
    .thumb .c{font-size: 12px; color:var(--muted)}
    .canvas{
      border:1px solid var(--line);
      border-radius: 16px;
      overflow:hidden;
      background:#fff;
      min-height: 340px;
      display:flex;
      align-items:center;
      justify-content:center;
    }
    .canvas img{max-width:100%; height:auto; display:block;}
    .canvas .ph{color:var(--muted); font-size: 13px; padding: 18px;}
    .zoombar{
      display:flex; align-items:center; gap: 8px;
    }
  
    .processing{ display:none; } /* legacy (unused) */

/* Full-page overlay shown during processing */
body.is-processing { overflow: hidden; }

#processing-overlay{
  position: fixed;
  inset: 0;
  width: 100%;
  height: 100%;
  display: none;
  align-items: center;
  justify-content: center;
  background: rgba(255,255,255,.65);
  backdrop-filter: blur(6px);
  -webkit-backdrop-filter: blur(6px);
  z-index: 9999;
}
#processing-overlay .processing-box{
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 16px 18px;
  display: flex;
  align-items: center;
  gap: 10px;
  color: var(--red);
  font-weight: 850;
  box-shadow: 0 20px 60px rgba(0,0,0,.12);
}
    .spin{
      width:12px;height:12px;
      border:2px solid rgba(230,57,70,.25);
      border-top-color: var(--red);
      border-radius:999px;
      animation: spin 0.9s linear infinite;
    }
    @keyframes spin{to{transform:rotate(360deg);}}
</style>
</head>
<body>
  <div class="nav">
    <div class="inner">
      <div class="brand">
        <img src="{{ logo_data_uri }}" alt="Lameco"/>
      </div>
      <div class="pill">Document Format Checker</div>
    </div>
  </div>

  <div class="wrap">
    <div class="grid">

      <!-- LEFT: checker -->
      <div>
        <div class="h1">Document Format Checker</div>
        <p class="sub">Ensure your documents match the correct formatting standards. Upload your files and let us check them for you.</p>

        <div class="card">
          <div class="bd">
            <form id="upload-form" method="post" action="/upload" enctype="multipart/form-data">
              <input id="file" name="file" type="file" accept=".pptx,.docx,.zip" required/>
              <div class="drop" id="drop">
                <div class="icon">⤒</div>
                <h4>Drop your files here</h4>
                <p>or click to browse • .docx • .pptx • .zip</p>
              </div>

              <!-- Client-side selected file (before submitting) -->
              <div id="pendingWrap" style="display:none; margin-top:14px;">
                <div class="row" style="justify-content:space-between;">
                  <div style="font-weight:750;">Uploaded Files (1)</div>
                  <div class="muted" style="font-size:12px;" id="pendingKind"></div>
                </div>
                <div style="height:10px"></div>
                <div class="filecard">
                  <div class="fileleft">
                    <div class="fileicon">📄</div>
                    <div class="filemeta">
                      <div style="font-weight:750;" id="pendingName">—</div>
                      <div class="badges"><span class="badge warn">● Ready</span></div>
                    </div>
                  </div>
                  <button type="button" class="iconbtn" id="pendingClear" title="Remove">✕</button>
                </div>
              </div>

              {% if job %}
                <div style="height:14px"></div>
                <div class="row" style="justify-content:space-between;">
                  <div style="font-weight:750;">Uploaded Files (1)</div>
                  <div class="muted" style="font-size:12px;">{{ job.kind|upper }}</div>
                </div>
                <div style="height:10px"></div>

                <div class="filechip">
                  <div class="fileleft">
                    <div class="fileicon">📄</div>
                    <div class="filemeta">
                      <div style="font-weight:750;">{{ job.source }}</div>
                      <div class="badges">
                        <span class="badge err" id="issue-count-badge">● {{ job.template_errors|length }} issues</span>
                        {% if job.logo_issues %}<span class="badge warn">● {{ job.logo_issues|length }} logo warnings</span>{% endif %}
                      </div>
                    </div>
                  </div>
                  <div style="color:#16A34A; font-weight:900;">✓</div>
                </div>

                <div style="height:12px"></div>
                <div class="row">
                  <button class="btn primary" type="submit">▶ Check Format</button>
                  <a class="btn" href="/download/xlsx?token={{ job.token }}">⇩ Export Report</a>
                </div>
              {% else %}
                <div style="height:12px"></div>
                <div class="row">
                  <button class="btn primary" type="submit">▶ Check Format</button>
                  <span class="muted" style="font-size:12px;">Upload a file first.</span>
                </div>
              {% endif %}
              <div id="processing-overlay" aria-live="polite" aria-busy="true">
                <div class="processing-box"><span class="spin"></span><span>Processing...</span></div>
              </div>
</form>
          </div>
        </div>

        <div style="height:14px"></div>

        <div class="card">
          <div class="hd">
            <h3>Format Issues</h3>
            {% if job %}
              <div class="muted" style="font-size:12px;" id="issue-summary-text">
                {{ job.template_errors|length }} issues found • {{ checked_rows|length if checked_rows else 0 }} resolved
              </div>
            {% else %}
              <div class="muted" style="font-size:12px;">No analysis yet</div>
            {% endif %}
          </div>
          <div class="bd" style="padding-top:0;">
            {% if job and job.template_errors %}
              <table id="issues-table">
                <thead>
                  <tr>
                    <th style="width:62px;">Done</th>
                    <th style="width:92px;">Type</th>
                    <th style="width:120px;">Location</th>
                    <th>Issue</th>
                  </tr>
                </thead>
                <tbody>
                  {% for e in job.template_errors %}
                    {% set idx = e.row_index %}
                    {% set is_done = (idx in checked_rows) if checked_rows else false %}
                    <tr data-row="{{ idx }}" class="{% if is_done %}crossed{% endif %}">
                      <td>
                        <input type="checkbox"
                               {% if is_done %}checked{% endif %}
                               onclick="toggleCrossFromIssue(event, this)"
                               data-row="{{ idx }}"
                               data-filehash="{{ job.file_hash or '' }}">
                      </td>
                      <td>
                        {% if e.code and 'WARN' in e.code %}
                          <span class="type warn">Warning</span>
                        {% elif e.code and 'INFO' in e.code %}
                          <span class="type info">Info</span>
                        {% else %}
                          <span class="type error">Error</span>
                        {% endif %}
                      </td>
                      <td class="muted">
                        {% if e.page %}Page {{ e.page }}{% else %}—{% endif %}
                      </td>
                      <td>{{ e.message }}</td>
                    </tr>
                  {% endfor %}
                </tbody>
              </table>
            {% else %}
              <div class="muted" style="font-size:13px;">No issues to show yet. Upload a file and run “Check Format”.</div>
            {% endif %}
          </div>
        </div>

        <div style="height:14px"></div>

        <div class="card">
          <div class="hd"><h3>Document Preview</h3>
            <div class="zoombar">
              <button class="btn small" type="button" id="zoom-out">−</button>
              <span class="muted" id="zoom-label" style="font-size:12px;">100%</span>
              <button class="btn small" type="button" id="zoom-in">＋</button>
            </div>
          </div>
          <div class="bd">
            {% if job and job.kind in ('pptx', 'docx') %}
              <div class="preview">
                <div class="thumbs" id="thumbs">
                  {% for s in job.slides %}
                    <div class="thumb {% if loop.first %}active{% endif %}" data-page="{{ s.page }}">
                      <div class="p">Page {{ s.page }}</div>
                      <div class="c">{{ job.errors_per_page.get(s.page|string, 0) }} issue(s)</div>
                    </div>
                  {% endfor %}
                </div>
                <div class="canvas" id="canvas">
                  {% if job.slides and job.slides[0].filename %}
                    <img id="slide-img" src="/slide/{{ job.token }}/{{ job.slides[0].page }}.png" alt="slide"/>
                  {% else %}
                    <div class="ph">No rendered slide image available. Install LibreOffice + poppler-utils to enable document preview.</div>
                  {% endif %}
                </div>
              </div>
            {% else %}
              <div class="muted" style="font-size:13px;">Upload a document to see a preview.</div>
            {% endif %}
          </div>
        </div>

      </div>

      <!-- RIGHT: settings + history -->
      <div>

        <div class="card">
          <div class="hd"><h3>Format Settings</h3></div>
          <div class="bd">
            <details open>
              <summary>Document Settings <span class="chev">▾</span></summary>
              <div class="section settings-grid">
                <div class="field">
                  <label>Body Font</label>
                  <input id="doc_body_font" value="{{ settings.docx.body_font if settings and settings.docx and settings.docx.body_font is defined else settings.docx.default_font }}"/>
                </div>
                <div class="field">
                  <label>Body Size</label>
                  <select id="doc_body_size">
                    {% for s in [9,10,11,12,13,14,16,18] %}
                      <option value="{{ s }}" {% if (settings.docx.body_size_pt|int if settings and settings.docx and settings.docx.body_size_pt is defined else settings.docx.font_size_pt|int) == s %}selected{% endif %}>{{ s }}pt</option>
                    {% endfor %}
                  </select>
                </div>

                <div class="field">
                  <label>Title Font</label>
                  <input id="doc_title_font" value="{{ settings.docx.title_font if settings and settings.docx and settings.docx.title_font is defined else (settings.docx.heading_font if settings and settings.docx and settings.docx.heading_font is defined else settings.docx.default_font) }}"/>
                </div>
                <div class="field">
                  <label>Title Size</label>
                  <select id="doc_title_size">
                    {% for s in [12,14,16,18,20,24,26,28,30,32,36,40,44,48,54,60,72] %}
                      <option value="{{ s }}" {% if (settings.docx.title_size_pt|int if settings and settings.docx and settings.docx.title_size_pt is defined else (settings.docx.heading_size_pt|int if settings.docx.heading_size_pt is defined else 32)) == s %}selected{% endif %}>{{ s }}pt</option>
                    {% endfor %}
                  </select>
                </div>

                <div class="field">
                  <label>Heading Font</label>
                  <input id="doc_heading_font" value="{{ settings.docx.heading_font if settings and settings.docx and settings.docx.heading_font is defined else (settings.docx.body_font if settings.docx.body_font is defined else settings.docx.default_font) }}"/>
                </div>
                <div class="field">
                  <label>Heading Size</label>
                  <select id="doc_heading_size">
                    {% for s in [12,13,14,16,18,20,24] %}
                      <option value="{{ s }}" {% if (settings.docx.heading_size_pt|int if settings and settings.docx and settings.docx.heading_size_pt is defined else 14) == s %}selected{% endif %}>{{ s }}pt</option>
                    {% endfor %}
                  </select>
                </div>

                <div class="field">
                  <label>Line Spacing</label>
                  <select id="doc_line_spacing">
                    {% for v in [1.0,1.15,1.2,1.5,2.0] %}
                      <option value="{{ v }}" {% if (settings.docx.line_spacing|float if settings and settings.docx and settings.docx.line_spacing is defined else 1.15) == v %}selected{% endif %}>{{ v }}</option>
                    {% endfor %}
                  </select>
                </div>
                <div class="field">
                  <label>Margins</label>
                  <select id="doc_margins">
                    <option value="Normal" selected>Normal</option>
                  </select>
                </div>
              </div>
            </details>

            <div style="height:10px"></div>

            <details open>
              <summary>PowerPoint Settings <span class="chev">▾</span></summary>
              <div class="section settings-grid">
                <div class="field">
                  <label>Title Font</label>
                  <input id="ppt_title_font" value="{{ settings.pptx.title_font }}"/>
                </div>
                <div class="field">
                  <label>Title Size (Allowed sizes)</label>
                   <input id="ppt_allowed_title_sizes" value="{{ (settings.pptx.allowed_title_sizes_pt|join(',') if settings and settings.pptx and settings.pptx.allowed_title_sizes_pt is defined else '54,60,72') }}"/>
                </div>

                <div class="field">
                  <label>Body Font</label>
                  <input id="ppt_body_font" value="{{ settings.pptx.body_font }}"/>
                </div>
                <div class="field">
                  <label>Body Size (Allowed sizes)</label>
                  <input id="ppt_allowed_sizes" value="{{ (settings.pptx.allowed_sizes_pt|join(',') if settings and settings.pptx and settings.pptx.allowed_sizes_pt is defined else '16,18,20,24,28,32,40,42,44,48,49,54,60,72') }}"/>
                </div>


                <!-- Slide size fields removed -->

              </div>
            </details>

            <div style="height:12px"></div>
            <button class="btn primary" type="button" id="save-settings" style="width:100%; justify-content:center;">
              Save Settings
            </button>
          </div>
        </div>

        <div style="height:14px"></div>

        <div class="card">
          <div class="hd"><h3>Recent History</h3></div>
          <div class="bd">
            <div class="hist" id="history">
              <div class="muted" style="font-size:13px;">Loading…</div>
            </div>
          </div>
        </div>

      </div>

    </div>
  </div>

  <div class="toast" id="toast"></div>

  <script>
    // UI loading indicator

    const drop = document.getElementById('drop');
    const fileInput = document.getElementById('file');
    const form = document.getElementById('upload-form');

    const processing = document.getElementById('processing-overlay');
    function showProcessing(){
      if(processing){
        processing.style.display = 'flex';
        document.body.classList.add('is-processing');
      }
    }
    function hideProcessing(){
      if(processing){
        processing.style.display = 'none';
        document.body.classList.remove('is-processing');
      }
    }
    // Show as soon as the form is submitted (drag-drop or button click)
    form.addEventListener('submit', () => {
      showProcessing();
    });
    const pendingWrap = document.getElementById('pendingWrap');
    const pendingName = document.getElementById('pendingName');
    const pendingKind = document.getElementById('pendingKind');
    const pendingClear = document.getElementById('pendingClear');

    function extKind(name){
      const n = (name||'').toLowerCase();
      if(n.endsWith('.pptx') || n.endsWith('.zip')) return 'PPTX';
      if(n.endsWith('.docx')) return 'DOCX';
      return '';
    }
    function showPending(file){
      if(!file){ pendingWrap.style.display='none'; return; }
      pendingWrap.style.display='block';
      pendingWrap.scrollIntoView({ behavior: 'smooth' });
      pendingName.textContent = file.name;
      pendingKind.textContent = extKind(file.name);
    }
    pendingClear?.addEventListener('click', ()=>{
      fileInput.value = '';
      showPending(null);
    });


    function toast(msg){
      const t = document.getElementById('toast');
      t.textContent = msg;
      t.classList.add('show');
      setTimeout(()=>t.classList.remove('show'), 2200);
    }

    drop?.addEventListener('click', ()=> fileInput?.click());
    drop?.addEventListener('dragover', (e)=>{e.preventDefault(); drop.style.borderColor = '#FFD2D6';});
    drop?.addEventListener('dragleave', ()=>{drop.style.borderColor = '#E5E7EB';});
    drop?.addEventListener('drop', (e)=>{
      e.preventDefault();
      drop.style.borderColor = '#E5E7EB';
      if(!e.dataTransfer.files?.length) return;
      fileInput.files = e.dataTransfer.files;
      showPending(fileInput.files[0]);
    });
    fileInput?.addEventListener('change', ()=> {
      showPending(fileInput.files[0]);
    });

    // --- settings save ---
    async function saveSettings(){
      const body = {
        docx: {
          body_font: document.getElementById('doc_body_font').value.trim(),
          body_size_pt: parseInt(document.getElementById('doc_body_size').value, 10),
          heading_font: document.getElementById('doc_heading_font').value.trim(),
          heading_size_pt: parseInt(document.getElementById('doc_heading_size').value, 10),
          title_font: document.getElementById('doc_title_font').value.trim(),
          title_size_pt: parseInt(document.getElementById('doc_title_size').value, 10),
          line_spacing: parseFloat(document.getElementById('doc_line_spacing').value),
          // keep compatibility keys in sync
          default_font: document.getElementById('doc_body_font').value.trim(),
          font_size_pt: parseInt(document.getElementById('doc_body_size').value, 10),
               },
        pptx: {
          title_font: document.getElementById('ppt_title_font').value.trim(),
          body_font: document.getElementById('ppt_body_font').value.trim(),
          // For compatibility, we can just grab the first values from the allowed lists
          // or just send 0 if backend now relies on the lists exclusively.
          // Or we parse the first item of the list to keep old keys valid.
          title_size_pt: 0, 
          body_size_pt: 0,
          allowed_sizes_pt: document.getElementById('ppt_allowed_sizes').value.trim(),
          allowed_title_sizes_pt: document.getElementById('ppt_allowed_title_sizes').value.trim()
        }
      };

      // Slide size logic removed


      const res = await fetch('/api/settings', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify(body)
      });
      const data = await res.json();
      if(data.ok){
        toast('Settings saved. Next check will use them.');
      }else{
        toast('Failed to save settings');
      }
    }
    document.getElementById('save-settings')?.addEventListener('click', saveSettings);

    // --- History list ---
    async function loadHistory(){
      const el = document.getElementById('history');
      if (!el) return;
      try{
        // Add timestamp to prevent caching
        const res = await fetch('/api/history?t=' + Date.now());
        const items = await res.json();
        if(!Array.isArray(items) || items.length === 0){
          el.innerHTML = '<div class="muted" style="font-size:13px;">No history yet.</div>';
          return;
        }
        el.innerHTML = '';
        for(const it of items){
          const div = document.createElement('div');
          div.className = 'hist-item';
          const when = it.last_opened_at || it.created_at || '';
          const issues = (it.issue_count ?? 0);
          const totalInfo = (it.total_issues !== undefined && it.total_issues > issues) 
             ? `<span style="color:#6B7280; font-weight:400; font-size:11px;">(of ${it.total_issues})</span>` 
             : '';

          // Highlight issues in red if > 0, otherwise green or muted
          // "decrease by the amount of checked boxes" implies we show the remaining count.
          const issueHtml = issues > 0 
            ? `<span style="color:#DC2626;font-weight:700;">${issues} issues</span> ${totalInfo}` 
            : `<span style="color:#16A34A;font-weight:700;">All resolved</span>`;
          
          div.innerHTML = `
            <div class="hist-left">
              <div class="fileicon">📄</div>
              <div class="hist-meta">
                <div class="fn">${(it.filename || 'Unknown')}</div>
                <div class="sm">${when} • ${issueHtml}</div>
              </div>
            </div>
            <div class="arrow">›</div>
          `;
          div.addEventListener('click', ()=>{
            if(it.token){ window.location.href = '/job/' + it.token; }
            else { toast('Cannot open this item. Re-upload the file.'); }
          });
          el.appendChild(div);
        }
      }catch(e){
        console.error(e);
      }
    }
    loadHistory();

    // --- Slide preview click + zoom ---
    let zoom = 1.0;
    const zoomLabel = document.getElementById('zoom-label');
    const slideImg = document.getElementById('slide-img');

    function applyZoom(){
      if(!slideImg) return;
      slideImg.style.transformOrigin = 'center top';
      slideImg.style.transform = `scale(${zoom})`;
      zoomLabel.textContent = Math.round(zoom*100) + '%';
    }
    document.getElementById('zoom-in')?.addEventListener('click', ()=>{ zoom = Math.min(2.5, zoom + 0.25); applyZoom(); });
    document.getElementById('zoom-out')?.addEventListener('click', ()=>{ zoom = Math.max(0.5, zoom - 0.25); applyZoom(); });
    applyZoom();

    document.querySelectorAll('.thumb')?.forEach(t=>{
      t.addEventListener('click', ()=>{
        document.querySelectorAll('.thumb').forEach(x=>x.classList.remove('active'));
        t.classList.add('active');
        const page = t.dataset.page;
        const token = "{{ job.token if job else '' }}";
        if(!page || !token || !slideImg) return;
        slideImg.src = `/slide/${token}/${page}.png`;
        zoom = 1.0;
        applyZoom();
      });
    });

    // --- Checkbox + row crossing sync ---
    async function setDone(fileHash, rowIndex, checked){
      const res = await fetch('/toggle', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body: JSON.stringify({file_hash: fileHash, row_index: rowIndex, checked})
      });
      return res.ok;
    }

    function updateIssueCounts() {
      const rows = document.querySelectorAll('#issues-table tbody tr[data-row]');
      if (!rows.length) return;
      
      const total = rows.length;
      let checked = 0;
      rows.forEach(r => {
        const cb = r.querySelector('input[type="checkbox"]');
        if (cb && cb.checked) checked++;
      });
      
      const badge = document.getElementById('issue-count-badge');
      if (badge) {
        badge.textContent = `● ${total - checked} issues`;
      }
      
      const summary = document.getElementById('issue-summary-text');
      if (summary) {
        summary.textContent = `${total} issues found • ${checked} resolved`;
      }
    }

    function toggleCrossFromIssue(ev, cb){
      ev.stopPropagation();
      const rowIndex = parseInt(cb.dataset.row, 10);
      const fileHash = cb.dataset.filehash || '';
      const checked = cb.checked;
      const tr = document.querySelector(`tr[data-row="${rowIndex}"]`);
      if(tr) tr.classList.toggle('crossed', checked);
      
      // Update counts immediately
      updateIssueCounts();
      
      setDone(fileHash, rowIndex, checked).then(() => {
          // Reload history to reflect the updated issue count
          loadHistory();
      }).catch(()=>{});
    }
    
    // Initialize counts on load
    updateIssueCounts();
    
    window.toggleCrossFromIssue = toggleCrossFromIssue;
  </script>
</body>
</html>'''


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)