"""Incident report tracker - entry point.

Collects incident report PDFs (Outlook attachments or files dropped in
``incident_inbox/``), parses them, QA-reviews them and updates the
damaged-fitting tracker workbook, including the monthly IR annex sheets.

Usage:
    python run_incident.py               # Outlook fetch + process
    python run_incident.py --manual      # process incident_inbox/ only
    python run_incident.py --reprocess   # ignore ledger, re-read everything
    python run_incident.py --dry-run     # parse + report, no workbook write

Idempotent: re-runs and re-sent mails never duplicate rows
(incident_processed.json ledger keyed by file content hash).
"""

import argparse
import datetime
import hashlib
import json
import logging
import os
import sys

try:
    import yaml
except ImportError:
    yaml = None

from incident_parse import parse_incident_pdf, review_record
from incident_tracker import (open_tracker, upsert_fittings, rebuild_annex_sheets,
                              rebuild_qa_sheet, rebuild_summary_sheet, save_tracker)

HERE = os.path.dirname(os.path.abspath(__file__))

DEFAULTS = {
    "inbox_folder": "incident_inbox",
    "tracker_file": "DAMAGED_FITTING_TRACKER.xlsx",
    "ledger_file": "incident_processed.json",
    "log_file": os.path.join("logs", "incident.log"),
    "outlook": {
        "enabled": True,
        "mailbox": "",                 # default account when blank
        "folder": "Inbox",
        "subject_contains": ["incident"],   # any-of, case-insensitive
        "lookback_days": 30,
        "attachment_ext": ".pdf",
    },
    "review": {},                      # see incident_parse.DEFAULT_REVIEW
}


def load_config(path=None):
    """Read the `incident:` section of the suite's config.yaml (or a
    standalone yaml); missing file/section just means defaults."""
    cfg = json.loads(json.dumps(DEFAULTS))  # deep copy
    candidates = [path] if path else [
        os.path.join(HERE, "config.yaml"),
        os.path.join(HERE, "config_incident.yaml"),
    ]
    for cand in candidates:
        if cand and os.path.exists(cand) and yaml:
            with open(cand, encoding="utf-8") as fh:
                data = yaml.safe_load(fh) or {}
            section = data.get("incident", data if path else {})
            for key, val in (section or {}).items():
                if isinstance(val, dict) and isinstance(cfg.get(key), dict):
                    cfg[key].update(val)
                else:
                    cfg[key] = val
            break
    return cfg


def setup_logging(log_file):
    os.makedirs(os.path.dirname(os.path.join(HERE, log_file)) or ".", exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(os.path.join(HERE, log_file), encoding="utf-8"),
                  logging.StreamHandler(sys.stdout)])
    return logging.getLogger("incident")


def load_ledger(path):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    return {}


def save_ledger(ledger, path):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(ledger, fh, indent=1, default=str)
    os.replace(tmp, path)


def file_sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def fetch_from_outlook(cfg, inbox_dir, log):
    """Save matching PDF attachments from Outlook into the inbox folder.
    Uses the same no-IT-permissions COM approach as the other trackers."""
    try:
        import win32com.client  # noqa: WPS433 (windows only)
    except ImportError:
        log.warning("pywin32 not available - skipping Outlook fetch "
                    "(use --manual and drop PDFs into %s)", inbox_dir)
        return []
    ocfg = cfg["outlook"]
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    if ocfg.get("mailbox"):
        folder = outlook.Folders[ocfg["mailbox"]].Folders[ocfg.get("folder", "Inbox")]
    else:
        folder = outlook.GetDefaultFolder(6)  # olFolderInbox
    since = datetime.datetime.now() - datetime.timedelta(days=int(ocfg.get("lookback_days", 30)))
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    items = items.Restrict("[ReceivedTime] >= '%s'" % since.strftime("%m/%d/%Y %H:%M %p"))
    needles = [s.lower() for s in ocfg.get("subject_contains", [])]
    saved = []
    for item in items:
        try:
            subject = (item.Subject or "").lower()
        except Exception:
            continue
        if needles and not any(n in subject for n in needles):
            continue
        try:
            attachments = item.Attachments
        except Exception:
            continue
        for i in range(1, attachments.Count + 1):
            att = attachments.Item(i)
            name = str(att.FileName or "")
            if not name.lower().endswith(ocfg.get("attachment_ext", ".pdf")):
                continue
            dest = os.path.join(inbox_dir, name)
            if not os.path.exists(dest):
                att.SaveAsFile(dest)
                saved.append(dest)
                log.info("saved attachment %s (mail: %s)", name, item.Subject)
    return saved


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--config", help="path to config yaml (default: config.yaml, section 'incident')")
    ap.add_argument("--manual", action="store_true",
                    help="skip Outlook; process files already in the inbox folder")
    ap.add_argument("--reprocess", action="store_true",
                    help="ignore the ledger and re-read every PDF in the inbox folder")
    ap.add_argument("--dry-run", action="store_true",
                    help="parse and report only; do not touch the workbook/ledger")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    log = setup_logging(cfg["log_file"])
    inbox_dir = os.path.join(HERE, cfg["inbox_folder"])
    os.makedirs(inbox_dir, exist_ok=True)
    tracker_path = os.path.join(HERE, cfg["tracker_file"])
    ledger_path = os.path.join(HERE, cfg["ledger_file"])

    if not args.manual and cfg.get("outlook", {}).get("enabled", True):
        fetch_from_outlook(cfg, inbox_dir, log)

    ledger = load_ledger(ledger_path)
    pdfs = sorted(f for f in os.listdir(inbox_dir) if f.lower().endswith(".pdf"))
    new_records = []
    for name in pdfs:
        path = os.path.join(inbox_dir, name)
        sha = file_sha1(path)
        if sha in ledger and not args.reprocess:
            continue
        try:
            rec = parse_incident_pdf(path)
        except Exception as exc:  # keep the run alive on one bad PDF
            log.error("FAILED to parse %s: %s", name, exc)
            ledger.setdefault(sha, {"file": name, "error": str(exc),
                                    "processed": str(datetime.datetime.now())})
            continue
        flags = review_record(rec, cfg.get("review"))
        new_records.append((sha, rec, flags))
        log.info("parsed %s: WO %s, %d fitting(s)%s", name, rec.get("wo_number"),
                 len(rec.get("fittings", [])),
                 " - FLAGS: " + "; ".join(flags) if flags else "")

    if args.dry_run:
        log.info("dry run: %d new report(s), workbook untouched", len(new_records))
        return 0

    if not new_records and not args.reprocess:
        log.info("nothing new to process")
        return 0

    wb = open_tracker(tracker_path)
    for sha, rec, flags in new_records:
        added, updated, warns = upsert_fittings(wb, rec, flags)
        for w in warns:
            log.warning("%s: %s", rec["file"], w)
        ledger[sha] = {
            "file": rec["file"], "wo": rec.get("wo_number"),
            "fittings": [f["ref"] for f in rec.get("fittings", [])],
            "report_date": str(rec.get("report_date")),
            "flags": flags, "rows_added": added, "rows_updated": updated,
            "record": {k: str(v) for k, v in rec.items()
                       if k not in ("details_of_incident", "observation", "action",
                                    "root_cause", "resolution")},
            "processed": str(datetime.datetime.now()),
        }

    # QA sheet covers everything ever processed (ledger) incl. this run
    reviews = []
    seen = set()
    for sha, rec, flags in new_records:
        reviews.append((rec, flags))
        seen.add(rec["file"])
    for sha, entry in ledger.items():
        if entry.get("file") in seen or "record" not in entry:
            continue
        stub = dict(entry["record"])
        stub["annex_photos"] = int(stub.get("annex_photos") or 0)
        stub["sig_prepared"] = stub.get("sig_prepared") == "True"
        stub["sig_submitted"] = stub.get("sig_submitted") == "True"
        stub["fittings"] = [{"ref": r} for r in entry.get("fittings", [])]
        rd = entry.get("report_date")
        stub["report_date"] = (datetime.date.fromisoformat(rd)
                               if rd and rd != "None" else None)
        reviews.append((stub, entry.get("flags", [])))

    rebuild_annex_sheets(wb)
    rebuild_qa_sheet(wb, reviews)
    rebuild_summary_sheet(wb)
    save_tracker(wb, tracker_path)
    save_ledger(ledger, ledger_path)
    log.info("tracker updated: %s (%d new report(s))", tracker_path, len(new_records))
    return 0


if __name__ == "__main__":
    sys.exit(main())
