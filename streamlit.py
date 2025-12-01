import os
import pandas as pd
import traceback
from collections import defaultdict
import streamlit as st
from io import BytesIO
from zipfile import ZipFile

# ===== PDF Libraries =====
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ===== Logging =====
import logging
from logging.handlers import RotatingFileHandler

LOG_FILENAME = "app.log"
logger = logging.getLogger("ExamSeatingLogger")
logger.setLevel(logging.DEBUG)
handler = RotatingFileHandler(LOG_FILENAME, maxBytes=500000, backupCount=3)
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
if not logger.hasHandlers():
    logger.addHandler(handler)

# ============================================================
# ABSOLUTE PATH FOR no_image_available.png
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
NO_IMAGE_PATH = os.path.join(BASE_DIR, "no_image_available.png")

if not os.path.exists(NO_IMAGE_PATH):
    raise FileNotFoundError(
        f"\n❌ 'no_image_available.png' missing!\nPlace it here:\n{NO_IMAGE_PATH}"
    )

# ============================================================
# Utilities
# ============================================================

def safe_strip(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def check_clashes(subject_rolls):
    clashes = []
    subs = list(subject_rolls.keys())
    for i in range(len(subs)):
        for j in range(i + 1, len(subs)):
            inter = subject_rolls[subs[i]].intersection(subject_rolls[subs[j]])
            for r in inter:
                clashes.append((subs[i], subs[j], r))
    return clashes

# ============================================================
# Allocation
# ============================================================

def allocate_for_slot(date, day, slot, subjects, students_by_subject, rooms_df, buffer, density, roll_name_map):

    subject_rolls = {}
    subject_counts = {}

    for s in subjects:
        rolls = set(safe_strip(r) for r in students_by_subject.get(s, []) if safe_strip(r))
        subject_rolls[s] = rolls
        subject_counts[s] = len(rolls)

    rooms_info = {}
    rooms_by_block = defaultdict(list)

    for _, row in rooms_df.iterrows():
        room = safe_strip(row["Room No."])
        block = safe_strip(row["Block"]) or "Block"
        cap = int(row["Exam Capacity"])
        rooms_info[room] = {"block": block, "capacity": cap, "remaining": 0}
        rooms_by_block[block].append(room)

    def eff_cap_room(r):
        return max(0, rooms_info[r]["capacity"] - buffer)

    def eff_cap_per_sub(r):
        base = eff_cap_room(r)
        return base // 2 if density == "sparse" else base

    for r in rooms_info:
        rooms_info[r]["remaining"] = eff_cap_per_sub(r)

    clashes = check_clashes(subject_rolls)

    assignments = {s: [] for s in subjects}

    for s in sorted(subjects, key=lambda x: -subject_counts[x]):
        needed = subject_counts[s]
        assigned = 0

        for block, room_list in rooms_by_block.items():
            for r in room_list:
                if assigned >= needed:
                    break
                take = min(rooms_info[r]["remaining"], needed - assigned)
                if take > 0:
                    rolls = list(subject_rolls[s])[assigned:assigned + take]
                    assignments[s].append({"room": r, "rolls": rolls})
                    rooms_info[r]["remaining"] -= take
                    assigned += take

            if assigned >= needed:
                break

    overall_rows = []
    seats_left_rows = []

    for s in subjects:
        for item in assignments[s]:
            overall_rows.append({
                "Date": date,
                "Day": day,
                "course_code": s,
                "Room": item["room"],
                "Allocated_students_count": len(item["rolls"]),
                "Roll_list(semicolon separated)": ";".join(item["rolls"])
            })

    for r, info in rooms_info.items():
        used = info["capacity"] - info["remaining"]
        seats_left_rows.append({
            "Room No.": r,
            "Exam Capacity": info["capacity"],
            "Block": info["block"],
            "Alloted": used,
            "Vacant": info["remaining"]
        })

    return assignments, overall_rows, seats_left_rows, clashes

# ============================================================
# PDF GENERATOR (with *relative* column widths)
# ============================================================

styles = getSampleStyleSheet()

student_style = ParagraphStyle(
    "student_style",
    parent=styles["Normal"],
    fontSize=10,
    leading=12,
    alignment=0   # left
)

PAGE_WIDTH = A4[0] - 20   # minus margins
CARD_WIDTH = PAGE_WIDTH / 3

def make_card(name, roll):

    img = Image(NO_IMAGE_PATH, width=32, height=32)

    text = Paragraph(
        f"<b>{name}</b><br/>Roll: {roll}<br/>Sign: _______________________",
        student_style
    )

    card = Table(
        [[img, text]],
        colWidths=[CARD_WIDTH * 0.25, CARD_WIDTH * 0.75]
    )

    card.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 1, colors.black),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))

    return card

def build_attendance_pdf(pdf_buffer, date, day, slot, room, subject, rolls, roll_name_map):

    pdf = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        leftMargin=10,
        rightMargin=10,
        topMargin=10
    )

    elements = []
    elements.append(Paragraph("<b>IITP Attendance System</b>", styles["Title"]))
    elements.append(Spacer(1, 8))

    header = Table([
        [f"Date: {date} ({day}) | Shift: {slot} | Room: {room} | Students: {len(rolls)}"],
        [f"Subject: {subject} | Present: | Absent:"]
    ], colWidths=[PAGE_WIDTH])

    header.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))

    elements.append(header)
    elements.append(Spacer(1, 12))

    # ---- Cards in rows of 3 ----
    rows = []
    row = []

    for roll in rolls:
        name = roll_name_map.get(roll, "(name not found)")
        row.append(make_card(name, roll))

        if len(row) == 3:
            rows.append(row)
            row = []

    if row:
        while len(row) < 3:
            row.append("")
        rows.append(row)

    table = Table(rows, colWidths=[CARD_WIDTH, CARD_WIDTH, CARD_WIDTH])
    table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))

    elements.append(table)
    elements.append(Spacer(1, 20))

    # Invigilator
    elements.append(Paragraph("<b>Invigilator Name & Signature</b>", styles["Heading4"]))

    inv = [["Sl No.", "Name", "Signature"]] + [["", "", ""] for _ in range(8)]
    inv_table = Table(inv, colWidths=[60, PAGE_WIDTH - 260, 200])
    inv_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 1, colors.black)
    ]))

    elements.append(Spacer(1, 6))
    elements.append(inv_table)

    pdf.build(elements)

# ============================================================
# STREAMLIT APP
# ============================================================

st.title("🪑 Exam Seating Arrangement Generator (Auto-Aligned PDF)")

uploaded_file = st.file_uploader("Upload input_data.xlsx", type=["xlsx"])
buffer = st.number_input("Buffer", 0, 50, value=5)
density = st.radio("Density", ["dense", "sparse"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        timetable = pd.read_excel(xls, "in_timetable", dtype=str)
        students_df = pd.read_excel(xls, "in_course_roll_mapping", dtype=str)
        rooms_df = pd.read_excel(xls, "in_room_capacity", dtype=str)
        mapping_df = pd.read_excel(xls, "in_roll_name_mapping", dtype=str)

        students_by_subject = defaultdict(list)
        for _, r in students_df.iterrows():
            s = safe_strip(r["course_code"])
            roll = safe_strip(r["rollno"])
            if s and roll:
                students_by_subject[s].append(roll)

        roll_name_map = {
            safe_strip(r["Roll"]): safe_strip(r["Name"])
            for _, r in mapping_df.iterrows()
            if safe_strip(r["Roll"])
        }

        zip_buffer = BytesIO()

        with ZipFile(zip_buffer, "w") as zipf:

            for _, row in timetable.iterrows():

                date_val = safe_strip(row["Date"])
                if isinstance(date_val, pd.Timestamp):
                    date = date_val.strftime("%Y-%m-%d")
                else:
                    date = date_val.split(" ")[0]

                day = safe_strip(row["Day"])

                for slot in ["Morning", "Evening"]:
                    subjects = [
                        safe_strip(s) for s in safe_strip(row.get(slot, "")).split(";")
                        if safe_strip(s)
                    ]
                    if not subjects:
                        continue

                    assignments, _, _, _ = allocate_for_slot(
                        date, day, slot, subjects,
                        students_by_subject, rooms_df, buffer, density, roll_name_map
                    )

                    for course, allocs in assignments.items():
                        for alloc in allocs:

                            pdf_buf = BytesIO()
                            build_attendance_pdf(
                                pdf_buf,
                                date=date,
                                day=day,
                                slot=slot,
                                room=alloc["room"],
                                subject=course,
                                rolls=alloc["rolls"],
                                roll_name_map=roll_name_map
                            )
                            pdf_buf.seek(0)

                            filename = f"{date}/{slot}/{date}_{course}_{alloc['room']}_{slot}.pdf"
                            zipf.writestr(filename, pdf_buf.read())

        zip_buffer.seek(0)
        st.success("🎉 Generated Successfully!")
        st.download_button(
            "📥 Download ZIP",
            data=zip_buffer,
            file_name="seating_arrangement.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(str(e))
        st.text(traceback.format_exc())
        logger.error(str(e), exc_info=True)
