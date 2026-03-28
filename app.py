"""
Seating Arrangement Generator - 1st Summative Evaluation 2026
Streamlit app to auto-generate exam seating from Excel student data.
"""

import streamlit as st
import pandas as pd
import io
import math
from collections import defaultdict

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, PageBreak, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF

# ─── Constants & Helpers ─────────────────────────────────────────────────────

CLASS_ORDER = ["V", "VI", "VII", "VIII", "IX", "X"]

CLASS_COLORS = {
    "V":    "#4CAF50",
    "VI":   "#2196F3",
    "VII":  "#9C27B0",
    "VIII": "#FF5722",
    "IX":   "#00BCD4",
    "X":    "#FF9800",
}

SCHOOL_NAME = "1st Summative Evaluation 2026"

def sort_classes(class_iterable):
    """Sorts an iterable of class names strictly according to the Roman Numeral CLASS_ORDER."""
    valid_classes = [c for c in class_iterable if pd.notna(c)]
    return sorted(valid_classes, key=lambda c: CLASS_ORDER.index(c) if c in CLASS_ORDER else 999)

# ─── Data Ingestion ──────────────────────────────────────────────────────────

def _normalize_class(raw):
    s = str(raw).strip()
    for prefix, cls in [
        ("V  : ", "V"), ("VI : ", "VI"), ("VII : ", "VII"),
        ("VIII : ", "VIII"), ("IX", "IX"), ("X", "X"),
    ]:
        if s.startswith(prefix):
            return cls
    return None

def read_students(file) -> pd.DataFrame:
    """Read Sheet1, extract class / roll / name / gender columns."""
    df = pd.read_excel(file, sheet_name="Sheet1", header=None)
    rows = []
    for _, r in df.iloc[2:].iterrows():   # row 0=title, 1=header, 2+=data
        cls = _normalize_class(r[0])
        if cls is None:
            continue
        try:
            roll = int(r[1])
        except (ValueError, TypeError):
            continue
        gender = str(r[10]).strip().upper()   # col 10 = Gender
        if gender not in ("BOYS", "GIRLS"):
            continue
        name = str(r[4]).strip() if pd.notna(r[4]) else ""  # col 4 = Name
        rows.append({"class": cls, "roll": roll, "name": name, "gender": gender})

    df_out = pd.DataFrame(rows)
    if not df_out.empty:
        # Enforce strict sorting at the dataset level
        df_out['class_rank'] = df_out['class'].apply(lambda x: CLASS_ORDER.index(x) if x in CLASS_ORDER else 99)
        df_out['gender_rank'] = df_out['gender'].apply(lambda x: 0 if x == 'BOYS' else 1)
        df_out = df_out.sort_values(['class_rank', 'gender_rank', 'roll']).drop(columns=['class_rank', 'gender_rank']).reset_index(drop=True)
        return df_out
    return df_out

# ─── Dynamic Room Distribution ───────────────────────────────────────────────

def distribute_to_rooms(df: pd.DataFrame, rooms_config: list, separate_genders: bool) -> tuple[dict, list]:
    """Distributes students into available rooms based on configured capacity and rules."""
    allocated_rooms = {r["name"]: [] for r in rooms_config}
    unassigned_students = []
    
    room_capacities = {r["name"]: sum(r["cols"]) * 3 for r in rooms_config}
    room_gender_locks = {r["name"]: None for r in rooms_config}
    
    def pop_mixed_student(student_dict):
        available_classes = [c for c in student_dict.keys() if len(student_dict[c]) > 0]
        if not available_classes:
            return None
        available_classes.sort(key=lambda c: len(student_dict[c]), reverse=True)
        return student_dict[available_classes[0]].pop(0)

    queues = {"BOYS": defaultdict(list), "GIRLS": defaultdict(list)}
    for s in df.to_dict("records"):
        queues[s["gender"]][s["class"]].append(s)

    genders_to_process = ["BOYS", "GIRLS"] if separate_genders else ["MIXED"]
    
    if not separate_genders:
        mixed_queues = defaultdict(list)
        for g in ["BOYS", "GIRLS"]:
            for c, students in queues[g].items():
                mixed_queues[c].extend(students)
        queues = {"MIXED": mixed_queues}

    for target_gender in genders_to_process:
        active_queue = queues[target_gender]
        
        while any(active_queue.values()):
            student = pop_mixed_student(active_queue)
            if not student:
                break
            
            placed = False
            for room in rooms_config:
                r_name = room["name"]
                if len(allocated_rooms[r_name]) >= room_capacities[r_name]:
                    continue
                if separate_genders:
                    current_lock = room_gender_locks[r_name]
                    if current_lock is None:
                        room_gender_locks[r_name] = target_gender
                    elif current_lock != target_gender:
                        continue 
                
                allocated_rooms[r_name].append(student)
                placed = True
                break
            
            if not placed:
                unassigned_students.append(student)

    return allocated_rooms, unassigned_students

# ─── Bench Seating Algorithm ─────────────────────────────────────────────────

def create_bench_layout(students: list[dict]) -> list[list]:
    """Arrange students into bench rows of 3."""
    groups = defaultdict(list)
    for s in students:
        groups[s["class"]].append(s)

    # Use the custom Roman numeral sort
    class_order = sort_classes(groups.keys())
    queues = {c: list(groups[c]) for c in class_order}
    benches = []

    while any(queues.values()):
        available = [(c, queues[c]) for c in class_order if queues[c]]
        if not available: break

        if len(available) == 1:
            cls, q = available[0]
            while q:
                triple = [q.pop(0), q.pop(0) if q else None, q.pop(0) if q else None]
                benches.append(triple)
            break

        available.sort(key=lambda x: len(x[1]), reverse=True)
        cls_a, q_a = available[0]
        cls_b, q_b = available[1]

        left   = q_a.pop(0)
        middle = q_b.pop(0)
        right  = q_a.pop(0) if q_a else None
        benches.append([left, middle, right])

    return benches

# ─── PDF & Excel Generation ──────────────────────────────────────────────────

def _style(name, **kwargs):
    base = dict(fontName="Helvetica", fontSize=9, alignment=TA_CENTER)
    base.update(kwargs)
    return ParagraphStyle(name, **base)

def _seat_cell(student):
    if student is None: return "—"
    g = "Boy" if student["gender"] == "BOYS" else "Girl"
    return f"Roll: {student['roll']}\nClass {student['class']}  [{g}]\n{student['name']}"

def _room_diagram(benches: list[list], room_config: dict) -> Drawing:
    col_heights = room_config["cols"]
    B_W, B_H = 50, 34                    
    GAP_X, GAP_Y = 8, 8
    SEAT_R = 6

    COLS = len(col_heights)
    max_rows = max(col_heights) if col_heights else 1
    
    dw = COLS * (B_W + GAP_X) + GAP_X
    dh = max_rows * (B_H + GAP_Y) + GAP_Y + 22  

    d = Drawing(dw, dh)

    board_w = min(dw * 0.6, 200)
    bx = (dw - board_w) / 2
    d.add(Rect(bx, dh - 20, board_w, 14, fillColor=colors.HexColor("#2e7d32"), strokeColor=colors.HexColor("#1b5e20"), strokeWidth=1))
    d.add(String(dw / 2, dh - 13, "BLACKBOARD", fontName="Helvetica-Bold", fontSize=7, fillColor=colors.white, textAnchor="middle"))

    bench_idx = 0
    for col_idx, rows_in_col in enumerate(col_heights):
        for row_idx in range(rows_in_col):
            if bench_idx >= len(benches): break
            bench = benches[bench_idx]
            x = GAP_X + col_idx * (B_W + GAP_X)
            y = dh - 22 - (row_idx + 1) * (B_H + GAP_Y)

            cls = next((s["class"] for s in bench if s), "V")
            fill_hex = CLASS_COLORS.get(cls, "#90caf9")
            fill = colors.HexColor(fill_hex)

            d.add(Rect(x, y, B_W, B_H, fillColor=colors.HexColor("#e3f2fd"), strokeColor=colors.HexColor("#90caf9"), strokeWidth=0.8))
            d.add(String(x + B_W / 2, y + B_H - 9, f"B{bench_idx+1}", fontName="Helvetica-Bold", fontSize=6, fillColor=colors.HexColor("#1a237e"), textAnchor="middle"))

            for sp_x, sp_y in [(x + 10, y + 10), (x + B_W / 2, y + 10), (x + B_W - 10, y + 10)]:
                d.add(Rect(sp_x - SEAT_R, sp_y - SEAT_R, SEAT_R * 2, SEAT_R * 2, fillColor=fill, strokeColor=colors.HexColor("#37474f"), strokeWidth=0.6))
            bench_idx += 1
    return d

def generate_pdf(allocated_rooms: dict, rooms_config: list) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=13 * mm, rightMargin=13 * mm, topMargin=13 * mm, bottomMargin=13 * mm)
    st_exam    = _style("exam", fontName="Helvetica-Bold", fontSize=15, spaceAfter=1*mm, textColor=colors.HexColor("#0d1b2a"))
    st_room    = _style("room", fontName="Helvetica-Bold", fontSize=22, spaceAfter=3*mm, textColor=colors.HexColor("#0f3460"))
    st_section = _style("section", fontName="Helvetica-Bold", fontSize=9, spaceAfter=2*mm, textColor=colors.HexColor("#37474f"))
    st_footer  = _style("footer", fontSize=7, textColor=colors.grey, fontName="Helvetica-Oblique")

    story = []
    for idx, config in enumerate(rooms_config):
        r_name = config["name"]
        students = allocated_rooms[r_name]
        if not students: continue
        if idx > 0: story.append(PageBreak())

        benches = create_bench_layout(students)

        story.append(Paragraph(SCHOOL_NAME, st_exam))
        story.append(Paragraph(str(r_name).upper(), st_room))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#0f3460"), spaceAfter=3*mm))

        counts = defaultdict(lambda: {"BOYS": 0, "GIRLS": 0})
        for s in students: counts[s["class"]][s["gender"]] += 1

        hdr = [["Class", "Boys", "Girls", "Total"]]
        body, tb, tg = [], 0, 0
        
        # Enforce strict Class order in the PDF table
        for cls in sort_classes(counts.keys()):
            b, g = counts[cls]["BOYS"], counts[cls]["GIRLS"]
            body.append([f"Class {cls}", str(b) if b else "–", str(g) if g else "–", str(b + g)])
            tb += b; tg += g
            
        body.append(["TOTAL", str(tb), str(tg), str(tb + tg)])

        summary_tbl = Table(hdr + body, colWidths=[35*mm, 24*mm, 24*mm, 24*mm], hAlign="CENTER")
        summary_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f3460")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#dde8f0")),
            ("ALIGN", (0, 0), (-1,-1), "CENTER"),
            ("VALIGN", (0, 0), (-1,-1), "MIDDLE"),
            ("GRID", (0, 0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
        ]))

        story.append(Paragraph("CLASS-WISE STUDENT COUNT", st_section))
        story.append(summary_tbl)
        story.append(Spacer(1, 4*mm))

        story.append(Paragraph(f"CLASSROOM LAYOUT  (Total Benches: {sum(config['cols'])})", st_section))
        story.append(_room_diagram(benches, config))
        story.append(Spacer(1, 4*mm))

        story.append(Paragraph("BENCH-WISE SEATING ARRANGEMENT", st_section))
        bench_hdr = [["Bench\nNo.", "LEFT SEAT\n(Roll | Class | Gender | Name)", "MIDDLE SEAT\n(Roll | Class | Gender | Name)", "RIGHT SEAT\n(Roll | Class | Gender | Name)"]]
        bench_rows = [[str(i + 1)] + [_seat_cell(s) for s in bench] for i, bench in enumerate(benches)]

        bench_tbl = Table(bench_hdr + bench_rows, colWidths=[12*mm, 53*mm, 53*mm, 53*mm], repeatRows=1)
        ts = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#16213e")),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#e8eaf6")),
            ("FONTSIZE",   (0, 0), (-1, -1), 7.5),
            ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#c0c0c0")),
        ]
        bench_tbl.setStyle(TableStyle(ts))
        story.append(bench_tbl)
        story.append(Spacer(1, 3*mm))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#aaaaaa")))
        story.append(Paragraph(f"{r_name}  ·  Total Students: {len(students)}  ·  {SCHOOL_NAME}", st_footer))

    doc.build(story)
    buf.seek(0)
    return buf

def generate_student_list_excel(student_df, classes_to_print):
    """Generates an Excel workbook with one sheet per class."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for cls in classes_to_print:
            cls_data = student_df[student_df['class'] == cls].copy()
            if cls_data.empty: continue
            
            boys = cls_data[cls_data['gender'] == 'BOYS'].sort_values('roll')
            girls = cls_data[cls_data['gender'] == 'GIRLS'].sort_values('roll')
            
            # Independent serial numbers starting at 1
            boys['SL.'] = range(1, len(boys) + 1)
            girls['SL.'] = range(1, len(girls) + 1)
            
            combined = pd.concat([boys, girls])
            
            export_df = combined[['gender', 'SL.', 'roll', 'name']].rename(columns={
                'gender': 'Gender',
                'roll': 'Roll Number',
                'name': 'Student Name'
            })
            
            export_df.to_excel(writer, sheet_name=f"Class {cls}", index=False)
            
    buf.seek(0)
    return buf

# ─── Streamlit UI ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Seating Arrangement Generator", page_icon="🏫", layout="wide")

    st.markdown("""
    <style>
        .main-title   { font-size:2.2rem; font-weight:800; color:#0f3460; margin-bottom:0; }
        .sub-title    { font-size:1rem;   color:#555;      margin-bottom:1.5rem; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<p class="main-title">🏫 Seating Arrangement Generator</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Fully Configurable Automated PDF & Excel Generation</p>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("⚙️ Configuration")
        class_mode = st.radio("Class Selection", ["All Classes", "Custom Classes"])
        st.markdown("---")
        st.subheader("Gender Rules")
        separate_genders = st.checkbox("🚫 Separate Boys & Girls into different rooms", value=False)
        
    uploaded = st.file_uploader("📂 Upload Student Excel File (.xlsx)", type=["xlsx"])

    if not uploaded:
        st.info("👆 Please upload the Excel file to begin.")
        return

    with st.spinner("Reading Excel..."):
        try:
            raw_df = read_students(uploaded)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return
            
    if raw_df.empty:
        st.warning("No valid students found in the file. Check formatting.")
        return

    # Use custom Roman sorting for Class Selectors
    available_classes = sort_classes(raw_df["class"].unique())
    if class_mode == "Custom Classes":
        selected_classes = st.sidebar.multiselect("Select Classes to Process", available_classes, default=available_classes)
        df = raw_df[raw_df["class"].isin(selected_classes)]
    else:
        df = raw_df

    st.success(f"✅ Loaded **{len(df):,}** students to process.")

    # =========================================================================
    # 4. CLASS SUMMARY & STUDENT LISTS (WITH STRICT SORTING & EXCEL EXPORT)
    # =========================================================================
    st.markdown("---")
    st.header("📋 Class Summary & Student Lists")

    if not df.empty:
        st.subheader("Class Summary")
        summary_data = []
        
        # Strictly sort classes for Summary Table
        for cls in sort_classes(df['class'].unique()):
            cls_df = df[df['class'] == cls]
            boys_count = len(cls_df[cls_df['gender'] == 'BOYS'])
            girls_count = len(cls_df[cls_df['gender'] == 'GIRLS'])
            summary_data.append({
                "Class": cls,
                "Total Students": boys_count + girls_count,
                "Boys": boys_count,
                "Girls": girls_count
            })
        
        summary_df = pd.DataFrame(summary_data)
        st.dataframe(summary_df, use_container_width=True)

        def generate_summary_pdf(sum_df):
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
            elements = []
            title_style = ParagraphStyle(name="Title", fontSize=16, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=10*mm)
            elements.append(Paragraph(f"{SCHOOL_NAME} - Class Summary", title_style))
            
            data = [["Class", "Total Students", "Boys", "Girls"]]
            for _, row in sum_df.iterrows():
                data.append([str(row['Class']), str(row['Total Students']), str(row['Boys']), str(row['Girls'])])
                
            t = Table(data, colWidths=[40*mm, 40*mm, 40*mm, 40*mm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f3460")),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,0), 8),
                ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#f5f8ff")),
                ('GRID', (0,0), (-1,-1), 1, colors.black),
            ]))
            elements.append(t)
            doc.build(elements)
            buf.seek(0)
            return buf

        st.download_button(
            label="📥 Download Class Summary PDF",
            data=generate_summary_pdf(summary_df),
            file_name="Class_Summary.pdf",
            mime="application/pdf"
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.subheader("Generate Class-wise Student List")
        
        col1, col2 = st.columns(2)
        with col1:
            list_option = st.radio("Select Generation Mode:", ["All Classes", "Selected Class Only"])
        
        selected_cls_list = []
        with col2:
            if list_option == "Selected Class Only":
                selected_cls = st.selectbox("Choose Class", sort_classes(df['class'].unique()))
                selected_cls_list = [selected_cls]
            else:
                selected_cls_list = sort_classes(df['class'].unique())

        def generate_student_list_pdf(student_df, classes_to_print):
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
            elements = []
            
            title_style = ParagraphStyle(name="Title", fontSize=16, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=2*mm)
            subtitle_style = ParagraphStyle(name="Sub", fontSize=14, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=6*mm)
            gender_style = ParagraphStyle(name="Gender", fontSize=12, alignment=TA_CENTER, fontName="Helvetica-Bold", spaceAfter=3*mm, spaceBefore=6*mm)
            
            for idx, cls in enumerate(classes_to_print):
                if idx > 0: elements.append(PageBreak())
                cls_data = student_df[student_df['class'] == cls]
                
                elements.append(Paragraph(f"{SCHOOL_NAME}", title_style))
                elements.append(Paragraph(f"Class: {cls} - Student List", subtitle_style))
                
                tbl_style = TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#e0e0e0")),
                    ('ALIGN', (0,0), (1,-1), 'CENTER'),
                    ('ALIGN', (2,0), (2,-1), 'LEFT'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,-1), 10),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('BOTTOMPADDING', (0,0), (-1,0), 6),
                ])
                
                boys = cls_data[cls_data['gender'] == 'BOYS'].sort_values('roll')
                if not boys.empty:
                    elements.append(Paragraph("BOYS", gender_style))
                    b_data = [["SL.", "Roll Number", "Student Name"]]
                    for i, (_, row) in enumerate(boys.iterrows(), 1):
                        b_data.append([str(i), str(row['roll']), str(row['name'])])
                    b_table = Table(b_data, colWidths=[20*mm, 40*mm, 110*mm], repeatRows=1)
                    b_table.setStyle(tbl_style)
                    elements.append(b_table)
                
                girls = cls_data[cls_data['gender'] == 'GIRLS'].sort_values('roll')
                if not girls.empty:
                    elements.append(Paragraph("GIRLS", gender_style))
                    g_data = [["SL.", "Roll Number", "Student Name"]]
                    for i, (_, row) in enumerate(girls.iterrows(), 1):
                        g_data.append([str(i), str(row['roll']), str(row['name'])])
                    g_table = Table(g_data, colWidths=[20*mm, 40*mm, 110*mm], repeatRows=1)
                    g_table.setStyle(tbl_style)
                    elements.append(g_table)

            doc.build(elements)
            buf.seek(0)
            return buf

        # Download Buttons Container
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="📄 Download Student List (PDF)",
                data=generate_student_list_pdf(df, selected_cls_list),
                file_name=f"Student_List_{list_option.replace(' ', '_')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        with col_dl2:
            st.download_button(
                label="📊 Download Student List (EXCEL)",
                data=generate_student_list_excel(df, selected_cls_list),
                file_name=f"Student_List_{list_option.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.markdown("---")

    # =========================================================================

    st.subheader("🚪 Room Configuration")
    st.markdown("Add rooms and define their seating layout. In the Layout column, enter the number of benches per column separated by commas (e.g., `6,9` means two columns of 6 and 9 benches).")
    
    default_rooms = pd.DataFrame([
        {"Room Name": "Room 1", "Layout (comma separated)": "11,11"},
        {"Room Name": "Room 2", "Layout (comma separated)": "6,6"},
        {"Room Name": "Room 3", "Layout (comma separated)": "7,6"},
        {"Room Name": "Room 4", "Layout (comma separated)": "6,6"},
        {"Room Name": "Room 5", "Layout (comma separated)": "9,9"}
    ])
    
    edited_room_df = st.data_editor(default_rooms, num_rows="dynamic", use_container_width=True)
    
    rooms_config = []
    total_system_capacity = 0
    for _, row in edited_room_df.iterrows():
        name = str(row["Room Name"]).strip()
        layout_str = str(row["Layout (comma separated)"]).strip()
        if not name or not layout_str: continue
        try:
            layout_str = layout_str.replace(":", ",")
            cols = [int(c.strip()) for c in layout_str.split(",") if c.strip().isdigit()]
            if cols:
                room_cap = sum(cols) * 3
                total_system_capacity += room_cap
                rooms_config.append({"name": name, "cols": cols, "capacity": room_cap})
        except ValueError:
            st.error(f"Invalid layout format in {name}. Please use numbers separated by commas.")
            return

    st.info(f"🪑 **Total System Capacity:** {total_system_capacity} Seats | **Total Students:** {len(df)}")
    if len(df) > total_system_capacity:
        st.error(f"⚠️ Warning: Not enough seats! You are short by {len(df) - total_system_capacity} seats.")

    allocated_rooms, unassigned = distribute_to_rooms(df, rooms_config, separate_genders)

    if unassigned:
        st.error(f"⚠️ {len(unassigned)} students could not be seated due to lack of space or strict gender isolation rules.")
        with st.expander("View Unassigned Students"): st.dataframe(unassigned)

    st.subheader("📊 Allocation Preview")
    preview_data = []
    for config in rooms_config:
        r_name = config["name"]
        students = allocated_rooms[r_name]
        preview_data.append({
            "Room Name": r_name,
            "Assigned Boys": sum(1 for s in students if s["gender"] == "BOYS"),
            "Assigned Girls": sum(1 for s in students if s["gender"] == "GIRLS"),
            "Total Occupied": f"{len(students)} / {config['capacity']}"
        })
    st.dataframe(pd.DataFrame(preview_data).set_index("Room Name"), use_container_width=True)

    if st.button("🖨️ Generate Seating Arrangement PDF", type="primary", use_container_width=True):
        with st.spinner("Calculating matrices and rendering PDF..."):
            try:
                pdf_buf = generate_pdf(allocated_rooms, rooms_config)
                st.balloons()
                st.download_button(
                    label="📥 Download Final Seating PDF",
                    data=pdf_buf,
                    file_name="Custom_Seating_Arrangement.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"PDF generation failed: {e}")
                st.exception(e)

if __name__ == "__main__":
    main()
