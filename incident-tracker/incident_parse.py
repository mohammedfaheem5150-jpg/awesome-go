"""Parser for AGL 'INCIDENT & SAFETY HAZARD REPORT' PDFs (AUH).

Reports are typed into the PDF text layer (no OCR needed).  Page 1 is the
form; page 2+ is the photo annex (before/after images, sometimes the
fitting reference as text).  Signatures are embedded IMAGES placed near
the two 'Signature:' labels; names are typed.

Two template variants seen in the wild differ only in geometry, so all
positional matching is relative (column x-positions are read from the
header row of the actual document, never hard-coded).
"""

import datetime
import os
import re

import pdfplumber

# Fitting asset refs seen: TEC102-01/015, LICATC-36/024, TEC102-03/120
FITTING_REF_RE = re.compile(r"\b([A-Z][A-Z0-9]{2,14}-\d{1,4}/\d{1,4})\b")
# Numbered fitting lines in the description:
#   "1.TEC102-01/015 - WO number# 43233489"
FITTING_WO_RE = re.compile(
    r"\d+\s*\.\s*([A-Z][A-Z0-9]{2,14}-\d{1,4}/\d{1,4})\s*[-–—:]*\s*"
    r"WO\s*(?:number)?\s*#?\s*(\d{6,10})",
    re.IGNORECASE,
)
REPORT_ID_RE = re.compile(
    r"Report\s*ID\s*[:#]?\s*(\d{8})\s*-?\s*(\d{1,3})?\s*WO\s*[:#]?\s*(\d{6,10})",
    re.IGNORECASE,
)
TIME_RE = re.compile(r"(\d{1,2}:\d{2})\s*(?:LT|HRS?)?", re.IGNORECASE)
DATE_TOKEN_RE = re.compile(r"\d{1,2}[-/][A-Za-z]{3,9}[-/]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}")
# Filename convention: 20251107_01_TEC_Fittings_damaged_..._WO_43233460_02.pdf
FILENAME_RE = re.compile(r"(\d{8})(?:_(\d{1,3}))?.*?WO_+(\d{6,10})", re.IGNORECASE)

SECTION_HEADINGS = [
    "DETAILS OF INCIDENT",
    "OBSERVATION",
    "ACTION",
    "ROOT CAUSE",
    "RESOLUTION",
]

# Checkbox labels on the form.  Section 1 boxes sit BEFORE their label,
# Section 2 boxes sit AFTER their label.
SECTION1_LABELS = ["Terminal Building", "Landside", "Airside", "Other Area"]
SECTION2_LABELS = [
    "Work Related Illness",   # longer labels first so they win overlaps
    "Work Related",
    "Property Damage",
    "Public Involved",
    "Vehicle Incident",
    "Private Property Involved",
    "Plant Damage",
    "Equipment Damage",
]
CATEGORY_LABELS = SECTION1_LABELS + SECTION2_LABELS

# Wingdings glyphs used for a checked box (0xF0FE = ballot box with X).
CHECKED_GLYPHS = {"\uf0fe", "\uf052", "\u2713", "\u2714", "\u2717", "\u2718", "X"}


def _is_reddish(color):
    if isinstance(color, (tuple, list)) and len(color) == 3:
        r, g, b = color
        return r > 0.6 and g < 0.45 and b < 0.45
    return False


def parse_date(text):
    """Parse the date formats the team uses: 7-Nov-25, 07/11/2025, ..."""
    if not text:
        return None
    text = text.strip().rstrip(".,;")
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%d-%B-%y", "%d-%B-%Y",
                "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_time(text):
    if not text:
        return None
    m = TIME_RE.search(text)
    if not m:
        return None
    try:
        return datetime.datetime.strptime(m.group(1), "%H:%M").time()
    except ValueError:
        return None


def _find_label_boxes(words):
    """Locate CATEGORY_LABELS on the page as (label, x0, x1, top) boxes."""
    boxes = []
    used = set()
    for label in CATEGORY_LABELS:
        parts = label.split()
        for i, w in enumerate(words):
            if i in used or w["text"] != parts[0]:
                continue
            seq = [w]
            j = i + 1
            for part in parts[1:]:
                while j < len(words) and abs(words[j]["top"] - w["top"]) > 3:
                    j += 1
                if j < len(words) and words[j]["text"].rstrip(":") == part:
                    seq.append(words[j])
                    j += 1
                else:
                    seq = None
                    break
            if seq:
                used.update(range(i, i + len(seq)))
                boxes.append({
                    "label": label,
                    "x0": seq[0]["x0"],
                    "x1": seq[-1]["x1"],
                    "top": seq[0]["top"],
                })
                break
    return boxes


def _detect_categories(page):
    """Checked boxes = red-marked rects OR Wingdings 'checked' glyphs.

    A mark belongs to the label on the same row whose box (label extent
    widened by the checkbox offset on both sides) contains the mark.
    """
    words = page.extract_words()
    boxes = _find_label_boxes(words)
    marks = []
    for r in page.rects:
        if (r["x1"] - r["x0"]) < 16 and (r["bottom"] - r["top"]) < 16:
            if _is_reddish(r.get("non_stroking_color")) or _is_reddish(r.get("stroking_color")):
                marks.append({"x": (r["x0"] + r["x1"]) / 2, "top": r["top"],
                              "strong": bool(r.get("fill"))})
    for c in page.chars:
        if c["text"] in CHECKED_GLYPHS and "Wingding" in (c.get("fontname") or ""):
            marks.append({"x": (c["x0"] + c["x1"]) / 2, "top": c["top"], "strong": True})

    checked = []
    for m in marks:
        best = None
        best_dx = None
        for b in boxes:
            if abs(b["top"] - m["top"]) > 8:
                continue
            if b["label"] in SECTION1_LABELS:
                # box PRECEDES the label; red outlines bleeding from an
                # adjacent hand-drawn highlight are not 'strong', skip them
                if not m["strong"]:
                    continue
                dx = b["x0"] - m["x"]
                ok = -5 <= dx <= 40
            else:
                # box FOLLOWS the label, but before the next label starts
                dx = m["x"] - b["x1"]
                nxt = min((o["x0"] for o in boxes
                           if abs(o["top"] - b["top"]) <= 8 and o["x0"] > b["x1"]),
                          default=page.width)
                ok = -5 <= dx <= 120 and m["x"] < nxt
            if ok and (best_dx is None or abs(dx) < best_dx):
                best, best_dx = b, abs(dx)
        if best and best["label"] not in checked:
            checked.append(best["label"])
    return checked


def _column_values(words, header_first_words, n_cols):
    """Generic 'header row -> value row' column matcher.

    header_first_words: list of (first_word, offset_words) tuples marking
    each column's first header word occurrence order on one row.
    Returns list of value strings per column.
    """
    # locate the header row: the row containing all first words in order
    for w in words:
        if w["text"] != header_first_words[0]:
            continue
        y = w["top"]
        row = [x for x in words if abs(x["top"] - y) < 3]
        texts = [x["text"] for x in row]
        if all(h in texts for h in header_first_words):
            # column anchors: x0 of each header first word, in encounter order
            anchors = []
            seen = 0
            expect = list(header_first_words)
            for x in row:
                if expect and x["text"] == expect[0]:
                    anchors.append(x["x0"])
                    expect.pop(0)
                    seen += 1
            if len(anchors) != n_cols:
                continue
            values = [[] for _ in range(n_cols)]
            for x in words:
                if not (y + 4 < x["top"] < y + 26):
                    continue
                col = 0
                for k, a in enumerate(anchors):
                    if x["x0"] >= a - 8:
                        col = k
                values[col].append(x["text"])
            return [" ".join(v).strip() for v in values]
    return [""] * n_cols


def _split_sections(text):
    """Split the description body into DETAILS/OBSERVATION/ACTION/... blocks."""
    out = {}
    positions = []
    for h in SECTION_HEADINGS:
        m = re.search(re.escape(h) + r"\s*:?", text, re.IGNORECASE)
        if m:
            positions.append((m.start(), m.end(), h))
    positions.sort()
    end_m = re.search(r"Report\s+Prepared\s+by\s*:", text, re.IGNORECASE)
    end_of_body = end_m.start() if end_m else len(text)
    for i, (start, hend, h) in enumerate(positions):
        nxt = positions[i + 1][0] if i + 1 < len(positions) else end_of_body
        out[h.lower().replace(" ", "_")] = text[hend:nxt].strip()
    return out


def _detect_signatures(page):
    """Signature slots: images vertically near the 'Signature:' labels.

    Returns (prepared_signed, submitted_signed).  The left label is the
    preparer (AGL team), the right one the receiving ADA engineer.
    """
    words = page.extract_words()
    labels = [w for w in words if w["text"].startswith("Signature")]
    if not labels:
        return None, None
    labels.sort(key=lambda w: w["x0"])
    left = labels[0]
    right = labels[-1] if len(labels) > 1 else None
    mid = right["x0"] if right else page.width / 2
    prepared = submitted = False
    for img in page.images:
        # near the signature row: image band overlapping label row +/- 45pt
        if img["bottom"] < left["top"] - 45 or img["top"] > left["bottom"] + 45:
            continue
        if img["x0"] < mid - 10:
            prepared = True
        else:
            submitted = True
    return prepared, submitted


def parse_incident_pdf(path):
    """Parse one incident report PDF into a flat record (dict).

    Never raises on missing fields - absent values come back as None/empty
    and are surfaced by review_record() instead, so one malformed report
    cannot stall the run.
    """
    rec = {
        "file": os.path.basename(path),
        "report_id": None, "report_serial": None, "wo_number": None,
        "report_date": None, "incident_date": None,
        "reporting_person": None, "reporting_manager": None,
        "location": None, "area_of_incident": None,
        "time_occurred": None, "time_attended": None, "time_rectified": None,
        "fittings": [],            # [{'ref':..., 'wo':...}]
        "categories": [],
        "details": None, "observation": None, "action": None,
        "root_cause": None, "resolution": None,
        "prepared_by": None, "prepared_designation": None,
        "submitted_to": None, "submitted_designation": None,
        "submission_date": None, "submitted_date": None,
        "sig_prepared": None, "sig_submitted": None,
        "annex_photos": 0, "annex_refs": [],
        "parse_errors": [],
    }

    fn = FILENAME_RE.search(os.path.basename(path))
    if fn:
        rec["report_id"] = fn.group(1)
        rec["report_serial"] = fn.group(2)
        rec["wo_number"] = fn.group(3)

    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        words = page.extract_words()

        m = REPORT_ID_RE.search(text)
        if m:
            rec["report_id"] = m.group(1)
            # serial is usually only in the filename; doc value wins if typed
            if m.group(2):
                rec["report_serial"] = m.group(2)
            doc_wo = m.group(3)
            if rec["wo_number"] and rec["wo_number"] != doc_wo:
                rec["parse_errors"].append(
                    "filename WO %s != document WO %s" % (rec["wo_number"], doc_wo))
            rec["wo_number"] = doc_wo
        elif "INCIDENT" not in text.upper():
            rec["parse_errors"].append("page 1 does not look like an incident report")

        # dates + names row (4 columns, matched by header x-positions)
        vals = _column_values(words, ["Date", "Date", "Name", "Name"], 4)
        rec["report_date"] = parse_date(vals[0])
        rec["incident_date"] = parse_date(vals[1])
        rec["reporting_person"] = vals[2] or None
        rec["reporting_manager"] = vals[3] or None

        m = re.search(r"Location\s*:\s*(.+)", text)
        if m:
            rec["location"] = m.group(1).strip()

        # area + three times row: the line after the 'Area of Incident ...
        # Time of Rectification' header that carries >= 2 clock times
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if "Area of Incident" in line and "Time" in line:
                for cand in lines[i + 1:i + 3]:
                    times = TIME_RE.findall(cand)
                    if len(times) >= 2:
                        rec["area_of_incident"] = TIME_RE.split(cand)[0].strip() or None
                        for key, t in zip(
                                ("time_occurred", "time_attended", "time_rectified"),
                                times):
                            rec[key] = t
                        break
                break

        for ref, wo in FITTING_WO_RE.findall(text):
            rec["fittings"].append({"ref": ref, "wo": wo})
        if not rec["fittings"]:
            # fall back: bare refs in the body, paired with the header WO
            body_refs = FITTING_REF_RE.findall(text)
            for ref in dict.fromkeys(body_refs):
                rec["fittings"].append({"ref": ref, "wo": rec["wo_number"]})
            if not rec["fittings"]:
                rec["parse_errors"].append("no fitting reference found")

        rec.update(_split_sections(text))
        rec["categories"] = _detect_categories(page)

        m = re.search(r"Report\s+Prepared\s+by\s*:\s*(.*?)\s*Submitted\s+to\s*:\s*(.*)",
                      text, re.IGNORECASE)
        if m:
            rec["prepared_by"] = m.group(1).strip() or None
            rec["submitted_to"] = m.group(2).strip() or None
        m = re.search(r"Designation\s*:\s*(.*?)\s*Designation\s*:\s*(.*)", text)
        if m:
            rec["prepared_designation"] = m.group(1).strip() or None
            rec["submitted_designation"] = m.group(2).strip() or None
        m = re.search(r"Submission\s+Date\s*:\s*(\S+)\s+Date\s*:\s*(\S+)", text)
        if m:
            rec["submission_date"] = parse_date(m.group(1))
            rec["submitted_date"] = parse_date(m.group(2))

        rec["sig_prepared"], rec["sig_submitted"] = _detect_signatures(page)

        for annex in pdf.pages[1:]:
            rec["annex_photos"] += len(annex.images or [])
            atext = annex.extract_text() or ""
            for ref in FITTING_REF_RE.findall(atext):
                if ref not in rec["annex_refs"]:
                    rec["annex_refs"].append(ref)

    return rec


# --------------------------------------------------------------------------
# QA review (config `incident: review:` section)

DEFAULT_REVIEW = {
    "require_preparer_signature": True,
    "require_submitted_signature": False,   # ADA usually signs later
    "flag_location_mismatch": True,
    "require_annex_photos": True,
    "max_report_date_drift_days": 1,
    "stand_series": [100, 800],
    "expected_prepared_by": "",             # blank = don't check
}

_STAND_RE = re.compile(r"\bST(?:AND)?\.?\s*0*(\d{3})\s*[LRA-Z]?\b", re.IGNORECASE)


def _norm_loc(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower().replace("between", "bw")
                  .replace("b/w", "bw"))


def review_record(rec, review_cfg=None):
    """Return a list of QA flag strings for one parsed record."""
    cfg = dict(DEFAULT_REVIEW)
    cfg.update(review_cfg or {})
    flags = list(rec.get("parse_errors") or [])

    if not rec.get("wo_number"):
        flags.append("missing WO number")
    if not rec.get("report_date"):
        flags.append("missing/unparseable report date")
    if rec.get("report_id") and rec.get("report_date"):
        if rec["report_date"].strftime("%Y%m%d") != rec["report_id"]:
            flags.append("Report ID %s does not match report date %s"
                         % (rec["report_id"], rec["report_date"]))
    if rec.get("incident_date") and rec.get("report_date"):
        drift = (rec["report_date"] - rec["incident_date"]).days
        if drift < 0:
            flags.append("report date before incident date")
        elif drift > cfg["max_report_date_drift_days"]:
            flags.append("report filed %d days after incident" % drift)

    # header WO should appear among the fitting WOs
    fit_wos = {f["wo"] for f in rec.get("fittings", []) if f.get("wo")}
    if rec.get("wo_number") and fit_wos and rec["wo_number"] not in fit_wos:
        flags.append("header WO %s not among fitting WOs %s"
                     % (rec["wo_number"], sorted(fit_wos)))

    if cfg["flag_location_mismatch"] and rec.get("location") and rec.get("area_of_incident"):
        if _norm_loc(rec["location"]) != _norm_loc(rec["area_of_incident"]):
            flags.append("Location '%s' != Area of Incident '%s'"
                         % (rec["location"], rec["area_of_incident"]))

    lo, hi = cfg["stand_series"]
    for src in (rec.get("location"), rec.get("area_of_incident")):
        m = _STAND_RE.search(src or "")
        if m and not lo <= int(m.group(1)) <= hi:
            flags.append("stand %s outside %d-%d series (%s)"
                         % (m.group(1), lo, hi, src))

    t_occ, t_att, t_rec = (parse_time(rec.get(k)) for k in
                           ("time_occurred", "time_attended", "time_rectified"))
    if t_occ and t_att and t_att < t_occ:
        flags.append("attended before occurrence time")
    if t_att and t_rec and t_rec < t_att:
        flags.append("rectified before attending time")
    if not (t_occ and t_att and t_rec):
        flags.append("one or more times missing/unparseable")

    if cfg["require_preparer_signature"] and rec.get("sig_prepared") is False:
        flags.append("preparer signature image missing")
    if cfg["require_submitted_signature"] and rec.get("sig_submitted") is False:
        flags.append("receiver (ADA) signature image missing")
    if cfg["expected_prepared_by"]:
        if (rec.get("prepared_by") or "").strip().lower() != cfg["expected_prepared_by"].strip().lower():
            flags.append("prepared by '%s' (expected '%s')"
                         % (rec.get("prepared_by"), cfg["expected_prepared_by"]))

    if cfg["require_annex_photos"] and not rec.get("annex_photos"):
        flags.append("no annex photos found")
    if rec.get("annex_refs"):
        page1 = {f["ref"] for f in rec.get("fittings", [])}
        for ref in rec["annex_refs"]:
            if page1 and ref not in page1:
                flags.append("annex fitting ref %s not on page 1" % ref)

    return flags


def fitting_type(ref):
    """Leading letters of the asset ref: TEC102-01/015 -> TEC, LICATC-36/024 -> LICATC."""
    m = re.match(r"([A-Z]+)", ref or "")
    return m.group(1) if m else ""
