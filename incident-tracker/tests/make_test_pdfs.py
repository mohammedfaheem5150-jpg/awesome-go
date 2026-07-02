"""Fabricate incident-report PDFs that mimic the real AUH form layout.

All names/WOs are fictional.  Used by the test suite; can also be run
directly to eyeball the output:  python tests/make_test_pdfs.py outdir
"""

import io
import os
import struct
import sys
import zlib

from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


def _tiny_png(width=60, height=30, rgb=(20, 20, 120)):
    """Minimal in-memory PNG (solid colour) - a stand-in signature image."""
    raw = b"".join(b"\x00" + bytes(rgb) * width for _ in range(height))

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def make_report(path, *, report_date="14-Jan-26", incident_date="14-Jan-26",
                wo="43900001", fittings=(("TEC102-05/001", "43900001"),),
                location="Twy C b/w C1 and C2", area=None,
                times=("08:00", "08:10", "09:00"),
                person="Testperson", manager="Testmanager Lead",
                prepared_by="Test Preparer", submitted_to="Test Receiver",
                signed=True, annex=True, airside_mark=True):
    """Draw a page-1 form + optional annex page matching the real geometry."""
    area = area if area is not None else location
    c = canvas.Canvas(path, pagesize=letter)
    w, h = letter

    def text(x, y_top, s, size=9, bold=False):
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x, h - y_top, s)

    text(200, 40, "INCIDENT & SAFETY HAZARD REPORT", 12, True)
    text(200, 60, "Report ID %s- WO: %s" % (report_date and
         __import__("datetime").datetime.strptime(report_date, "%d-%b-%y").strftime("%Y%m%d"), wo), 10)
    text(52, 90, "SECTION 1: (To be completed by person reporting Hazard/Incident)", 8)

    # four-column header + values row (same relative x layout as the form;
    # small font so neighbouring labels never overlap/merge)
    text(95, 120, "Date of Report:", 6)
    text(186, 120, "Date of Incident or Safety Hazard:", 6)
    text(305, 120, "Name of Reporting Person", 6)
    text(420, 120, "Name of Reporting Manager", 6)
    text(104, 135, report_date, 9)
    text(220, 135, incident_date, 9)
    text(310, 135, person, 9)
    text(425, 135, manager, 9)

    text(52, 165, "Location: %s" % location, 9)
    # section-1 checkbox row: grey boxes before each label; Airside marked red
    labels = [("Terminal Building", 93), ("Landside", 220), ("Airside", 319),
              ("Other Area", 402)]
    for label, x in labels:
        c.setStrokeColorRGB(0.8, 0.8, 0.8)
        c.rect(x - 25, h - 200, 10, 10, stroke=1, fill=0)
        text(x, 198, label, 8)
    if airside_mark:
        c.setFillColorRGB(1, 0, 0)
        c.rect(319 - 25, h - 200, 10, 10, stroke=0, fill=1)
        c.setFillColorRGB(0, 0, 0)

    text(52, 225, "Area of Incident or Safety Hazard Time of Occurrence of fault "
                  "Time of Attending the Fault Time of Rectification of Fault", 7)
    text(52, 240, "%s %sLT %sLT %sLT" % (area, *times), 9)

    text(52, 270, "Description of Incident or Safety Hazard:", 9, True)
    y = 290
    text(52, y, "DETAILS OF INCIDENT:", 9, True); y += 14
    text(52, y, "During inspection AGL Team observed fittings damaged at %s"
         % location, 9); y += 14
    text(52, y, "AGL HD recorded the incident in Maximo, under the WO number given below:", 9); y += 14
    for i, (ref, fwo) in enumerate(fittings, start=1):
        text(52, y, "%d.%s - WO number# %s" % (i, ref, fwo), 9); y += 14
    text(52, y, "OBSERVATION:", 9, True); y += 14
    text(52, y, "01- Fittings damaged by unknown Vehicle / Equipment.", 9); y += 14
    text(52, y, "ACTION:", 9, True); y += 14
    text(52, y, "AGL Team replaced the damaged fittings.", 9); y += 14
    text(52, y, "ROOT CAUSE:", 9, True); y += 14
    text(52, y, "Damage caused by unknown vehicle/equipment", 9); y += 14
    text(52, y, "RESOLUTION:", 9, True); y += 14
    text(52, y, "Damaged fittings replaced.", 9); y += 24

    text(52, y, "Report Prepared by: %s Submitted to: %s" % (prepared_by, submitted_to), 9)
    y += 14
    text(52, y, "Designation: AGL Team Leader Designation: ADA AGL Engineer", 9)
    y += 14
    text(52, y, "Submission Date: %s Date: %s" % (report_date, report_date), 9)
    y += 18
    sig_y = y
    text(52, sig_y, "Signature:", 9)
    text(304, sig_y, "Signature:", 9)
    if signed:
        img = ImageReader(io.BytesIO(_tiny_png()))
        c.drawImage(img, 130, h - sig_y - 20, width=55, height=25)
    c.showPage()

    if annex:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(270, h - 60, "ANNEX")
        c.setFont("Helvetica", 9)
        c.drawString(72, h - 80, fittings[0][0])
        img = ImageReader(io.BytesIO(_tiny_png(200, 150, (90, 90, 90))))
        c.drawImage(img, 72, h - 300, width=200, height=150)   # before
        c.drawImage(img, 300, h - 300, width=200, height=150)  # after
        c.setFont("Helvetica", 9)
        c.drawString(150, h - 320, "Before")
        c.drawString(380, h - 320, "After")
        c.showPage()
    c.save()
    return path


def make_all(outdir):
    os.makedirs(outdir, exist_ok=True)
    made = []
    made.append(make_report(
        os.path.join(outdir, "20260114_01_TEC_Fittings_damaged_on_Twy_C_WO_43900001.pdf")))
    # multi-fitting report
    made.append(make_report(
        os.path.join(outdir, "20260114_02_TEC_Fittings_damaged_on_Twy_D_WO_43900002.pdf"),
        wo="43900002", location="Twy D b/w D1 and D2",
        fittings=(("TEC102-07/010", "43900002"), ("TEC102-07/011", "43900003"))))
    # bad report: unsigned, area mismatch, stand out of series, no annex
    made.append(make_report(
        os.path.join(outdir, "20260220_01_LIC_Fitting_damaged_on_stand_902_WO_43900010.pdf"),
        report_date="22-Feb-26", incident_date="20-Feb-26", wo="43900010",
        location="Stand 902", area="Stand 611",
        fittings=(("LICATC-40/001", "43900010"),),
        signed=False, annex=False))
    return made


if __name__ == "__main__":
    for p in make_all(sys.argv[1] if len(sys.argv) > 1 else "test_pdfs"):
        print(p)
