import os
import pandas as pd
import traceback
from collections import defaultdict
import streamlit as st
from io import BytesIO
from zipfile import ZipFile

# PDF libs
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import mm

# Logging
import logging
from logging.handlers import RotatingFileHandler

# ========== Logger Setup ==========
LOG_FILENAME = "app.log"
logger = logging.getLogger("ExamSeatingLogger")
logger.setLevel(logging.DEBUG)

handler = RotatingFileHandler(LOG_FILENAME, maxBytes=1_000_000, backupCount=5)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)

if not logger.hasHandlers():
    logger.addHandler(handler)


# ========== Utility Functions ==========
def safe_strip(x):
    try:
        if pd.isna(x):
            return ''
        return str(x).strip()
    except:
        return ''


def check_clashes(subject_rolls):
    clashes = []
    try:
        subs = list(subject_rolls.keys())
        for i in range(len(subs)):
            for j in range(i + 1, len(subs)):
                s1, s2 = subs[i], subs[j]
                inter = subject_rolls[s1].intersection(subject_rolls[s2])
                if inter:
                    for r in sorted(inter):
                        clashes.append((s1, s2, r))
        return clashes
    except Exception as e:
        logger.error(f"Error in check_clashes: {e}", exc_info=True)
        return []


def find_building_allocation(subject, count, rooms_by_building, rooms_info, eff_cap_func):
    try:
        allocation = []
        for b, room_list in rooms_by_building.items():
            total = sum(eff_cap_func(r) for r in room_list if rooms_info[r]["remaining"] > 0)
            if total >= count:
                for r in room_list:
                    if count <= 0:
                        break
                    rem = rooms_info[r]["remaining"]
                    if rem <= 0:
                        continue
                    take = min(rem, count)
                    allocation.append((r, take))
                    count -= take
                if count == 0:
                    return allocation

        # fallback
        for b, room_list in rooms_by_building.items():
            for r in room_list:
                if count <= 0:
                    break
                rem = rooms_info[r]["remaining"]
                if rem <= 0:
                    continue
                take = min(rem, count)
                allocation.append((r, take))
                count -= take
            if count <= 0:
                return allocation

        return allocation
    except Exception as e:
        logger.error(f"Error in find_building_allocation: {e}", exc_info=True)
        return []


# ========== Core Allocation Logic ==========
def allocate_for_slot(date, day, slot, subjects, students_by_subject, rooms_df, buffer, density, roll_name_map):

    try:
        subject_rolls, subject_counts = {}, {}
        for s in subjects:
            rolls = set(
                safe_strip(r)
                for r in students_by_subject.get(s, [])
                if safe_strip(r)
            )
            subject_rolls[s] = rolls
            subject_counts[s] = len(rolls)

        rooms_info, rooms_by_building = {}, defaultdict(list)
        for _, row in rooms_df.iterrows():
            room = safe_strip(row["Room No."])
            building = safe_strip(row["Block"]) or "UnknownBlock"
            cap = int(row["Exam Capacity"])
            rooms_info[room] = {"building": building, "capacity": cap, "remaining": 0}
            rooms_by_building[building].append(room)

        def eff_cap_room(r):
            return max(0, rooms_info[r]["capacity"] - int(buffer))

        def eff_cap_per_sub(r):
            ec = eff_cap_room(r)
            return ec // 2 if density == "sparse" else ec

        for r in rooms_info:
            rooms_info[r]["remaining"] = eff_cap_per_sub(r)

        clashes = check_clashes(subject_rolls)

        subj_assignments = {s: [] for s in subjects}

        for s in sorted(subjects, key=lambda x: -subject_counts[x]):
            count = subject_counts[s]
            alloc = find_building_allocation(
                s, count, rooms_by_building, rooms_info, eff_cap_per_sub
            )
            assigned = 0
            for r, take in alloc:
                selected = list(subject_rolls[s])[assigned : assigned + take]
                subj_assignments[s].append({"room": r, "rolls": selected})
                assigned += len(selected)
                rooms_info[r]["remaining"] -= len(selected)

        overall_rows = []
        seats_left_rows = []

        for s in subjects:
            for p in subj_assignments[s]:
                overall_rows.append({
                    "Date": date,
                    "Day": day,
                    "course_code": s,
                    "Room": p["room"],
                    "Allocated_students_count": len(p["rolls"]),
                    "Roll_list(semicolon separated)": ";".join(p["rolls"])
                })

        for r, info in rooms_info.items():
            allocated = info["capacity"] - info["remaining"]
            seats_left_rows.append({
                "Room No.": r,
                "Exam Capacity": info["capacity"],
                "Block": info["building"],
                "Alloted": allocated,
                "Vacant": info["remaining"]
            })

        return subj_assignments, overall_rows, seats_left_rows, clashes

    except Exception as e:
        logger.error(f"Error in allocate_for_slot: {e}", exc_info=True)
        raise


# ========== PDF HELPER FUNCTIONS ==========

styles = getSampleStyleSheet()
label = styles["Normal"]
label.fontSize = 10
label.leading = 12


def make_card(name, roll):
    placeholder = Image("no_image_available.png", width=40, height=40)

    text = Paragraph(
        f"<b>{name}</b><br/>Roll: {roll}<br/>Sign: ____________________________",
        label
    )

    box = Table(
        [[placeholder, text]],
        colWidths=[45*mm, 90*mm],
        rowHeights=[30*mm]
    )

    box.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('VALIGN', (0,0), (1,0), 'TOP'),
        ('LEFTPADDING', (1,0), (1,0), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    return box


def build_attendance_pdf(pdf_buffer, date, day, slot, room, subject, rolls, roll_name_map):
    pdf = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        leftMargin=10,
        rightMargin=10,
        topMargin=10,
        bottomMargin=10
    )

    elements = []

    # HEADER (exact IITP style)
    title = Paragraph("<b>IITP Attendance System</b>", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 6))

    student_count = len(rolls)

    header_table = Table([
        [f"Date: {date} ({day}) | Shift: {slot} | Room No: {room} | Student count: {student_count}"],
        [f"Subject: {subject} | Stud Present: | Stud Absent:"]
    ], colWidths=[540])

    header_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
    ]))

    elements.append(header_table)
    elements.append(Spacer(1, 12))

    # STUDENT CARDS — GRID (3 per row)
    grid = []
    row_cards = []

    for i, roll in enumerate(rolls):
        name = roll_name_map.get(roll, "(name not found)")
        row_cards.append(make_card(name, roll))

        if len(row_cards) == 3:
            grid.append(row_cards)
            row_cards = []

    # last incomplete row
    if row_cards:
        while len(row_cards) < 3:
            row_cards.append("")  # empty cell
        grid.append(row_cards)

    grid_table = Table(grid, colWidths=[180, 180, 180])
    grid_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))

    elements.append(grid_table)
    elements.append(Spacer(1, 20))

    # INVIGILATOR TABLE
    inv_title = Paragraph("<b>Invigilator Name & Signature</b>", styles["Heading4"])
    elements.append(inv_title)
    elements.append(Spacer(1, 6))

    inv_data = [["Sl No.", "Name", "Signature"]] + [["", "", ""] for _ in range(8)]
    inv_table = Table(inv_data, colWidths=[60, 260, 200])

    inv_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.whitesmoke),
        ('FONTSIZE', (0,0), (-1,-1), 10),
    ]))

    elements.append(inv_table)

    pdf.build(elements)



# ========== STREAMLIT UI ==========
st.title("🪑 Exam Seating Arrangement Generator (PDF, IITP Layout)")

uploaded_file = st.file_uploader("Upload input_data.xlsx", type=["xlsx"])
buffer = st.number_input("Enter Buffer", min_value=0, value=5)
density = st.radio("Select Arrangement Type", ["dense", "sparse"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)

        timetable = pd.read_excel(xls, 'in_timetable', dtype=str)
        students_df = pd.read_excel(xls, 'in_course_roll_mapping', dtype=str)
        rooms_df = pd.read_excel(xls, 'in_room_capacity', dtype=str)
        mapping_df = pd.read_excel(xls, 'in_roll_name_mapping', dtype=str)

        students_by_subject = defaultdict(list)
        for _, row in students_df.iterrows():
            s, r = safe_strip(row.get("course_code", "")), safe_strip(row.get("rollno", ""))
            if s and r:
                students_by_subject[s].append(r)

        roll_name_map = {
            safe_strip(r["Roll"]): safe_strip(r["Name"]) or ""
            for _, r in mapping_df.iterrows()
            if safe_strip(r["Roll"])
        }

        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, "w") as zipf:
            for _, trow in timetable.iterrows():
                date_val = trow["Date"]
                if isinstance(date_val, pd.Timestamp):
                    date = date_val.strftime("%Y-%m-%d")
                else:
                    date = safe_strip(str(date_val).split(" ")[0])

                day = safe_strip(trow["Day"])

                for slot_col in ["Morning", "Evening"]:
                    subjects = [
                        safe_strip(s)
                        for s in safe_strip(trow.get(slot_col, "")).split(";")
                        if safe_strip(s)
                    ]
                    if not subjects:
                        continue

                    subj_assignments, overall_rows, seats_left_rows, clashes = allocate_for_slot(
                        date, day, slot_col, subjects, students_by_subject, rooms_df, buffer, density, roll_name_map
                    )

                    day_folder = f"{date}/{slot_col}"

                    for course, allocations in subj_assignments.items():
                        for alloc in allocations:
                            room = alloc["room"]
                            rolls = alloc["rolls"]
                            filename = f"{day_folder}/{date}_{course}_{room}_{slot_col}.pdf"

                            pdf_buf = BytesIO()
                            build_attendance_pdf(
                                pdf_buffer=pdf_buf,
                                date=date,
                                day=day,
                                slot=slot_col,
                                room=room,
                                subject=course,
                                rolls=rolls,
                                roll_name_map=roll_name_map
                            )
                            pdf_buf.seek(0)
                            zipf.writestr(filename, pdf_buf.read())

        zip_buffer.seek(0)

        st.success("✅ Seating arrangement PDFs generated successfully!")
        st.download_button(
            "📥 Download ZIP of PDFs",
            data=zip_buffer,
            file_name="seating_arrangement_structured.zip",
            mime="application/zip"
        )

    except Exception as e:
        logger.error(f"Error: {e}", exc_info=True)
        st.error(f"Error: {e}")
        st.text(traceback.format_exc())
