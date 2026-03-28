"""
Seating Arrangement Generator - 1st Summative Evaluation 2026
Streamlit app to auto-generate exam seating from Excel student data.
"""

import streamlit as st
import pandas as pd
import io
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
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics import renderPDF

# ─── Constants ───────────────────────────────────────────────────────────────

CLASS_ORDER = ["V", "VI", "VII", "VIII", "IX", "X"]

CLASS_COLORS = {
    "V":    "#4CAF50",
    "VI":   "#2196F3",
    "VII":  "#9C27B0",
    "VIII": "#FF5722",
    "IX":   "#00BCD4",
    "X":    "#FF9800",
}

# IX & X Boys → rooms 2,3,4 | IX & X Girls → rooms 1,5
# Classes V–VIII → all 5 rooms
ROOM_TARGETS = {
    "IX": {"BOYS": [2, 3, 4], "GIRLS": [1, 5]},
    "X":  {"BOYS": [2, 3, 4], "GIRLS": [1, 5]},
    **{c: {"BOYS": [1,2,3,4,5], "GIRLS": [1,2,3,4,5]} for c in ["V","VI","VII","VIII"]},
}

# Bench column counts per room (used for visual diagram)
ROOM_BENCH_COLS = {1: 11, 2: 6, 3: 6, 4: 6, 5: 9}

SCHOOL_NAME = "1st Summative Evaluation 2026"


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
    return df_out.sort_values(["class", "gender", "roll"]).reset_index(drop=True)


# ─── Room Distribution ───────────────────────────────────────────────────────

def distribute_to_rooms(df: pd.DataFrame) -> dict[int, list[dict]]:
    rooms: dict[int, list] = {i: [] for i in range(1, 6)}

    for cls in CLASS_ORDER:
        for gender in ["BOYS", "GIRLS"]:
            students = (
                df[(df["class"] == cls) & (df["gender"] == gender)]
                .to_dict("records")
            )
            if not students:
                continue
            targets = ROOM_TARGETS[cls][gender]
            n, nr = len(students), len(targets)
            idx = 0
            for i, room_id in enumerate(targets):
                chunk = n // nr + (1 if i < n % nr else 0)
                rooms[room_id].extend(students[idx : idx + chunk])
                idx += chunk

    return rooms


# ─── Bench Seating Algorithm ─────────────────────────────────────────────────

def create_bench_layout(students: list[dict]) -> list[list]:
    """
    Arrange students into bench rows of 3.
    Rule: LEFT & RIGHT seats = same class  |  MIDDLE seat = different class
    Returns list of [left, middle, right] where any slot can be None.
    """
    groups = defaultdict(list)
    for s in students:
        groups[s["class"]].append(s)

    queues = {c: list(groups[c]) for c in CLASS_ORDER if c in groups}
    benches = []

    while any(queues.values()):
        available = [(c, queues[c]) for c in CLASS_ORDER
                     if c in queues and queues[c]]

        if not available:
            break

        if len(available) == 1:
            cls, q = available[0]
            while q:
                triple = [q.pop(0), q.pop(0) if q else None,
                          q.pop(0) if q else None]
                benches.append(triple)
            break

        # Largest queue → outer seats; second largest → middle seat
        available.sort(key=lambda x: len(x[1]), reverse=True)
        cls_a, q_a = available[0]
        cls_b, q_b = available[1]

        left   = q_a.pop(0)
        middle = q_b.pop(0)
        right  = q_a.pop(0) if q_a else None
        benches.append([left, middle, right])

    return benches


# ─── PDF Generation ──────────────────────────────────────────────────────────

def _style(name, **kwargs):
    base = dict(fontName="Helvetica", fontSize=9, alignment=TA_CENTER)
    base.update(kwargs)
    return ParagraphStyle(name, **base)


def _seat_cell(student):
    if student is None:
        return "—"
    g = "Boy" if student["gender"] == "BOYS" else "Girl"
    return f"Roll: {student['roll']}\nClass {student['class']}  [{g}]\n{student['name']}"


def _room_diagram(benches: list[list], room_id: int) -> Drawing:
    """
    Draw a top-down classroom layout.
    Each bench = a small labeled rectangle with 3 seat circles.
    At most 10 benches per row in the diagram; rows wrap.
    """
    COLS = ROOM_BENCH_COLS[room_id]      # visual columns
    B_W, B_H = 50, 34                   # bench box width / height (pt)
    GAP_X, GAP_Y = 8, 8
    SEAT_R = 6

    total_benches = len(benches)
    rows = (total_benches + COLS - 1) // COLS
    dw = COLS * (B_W + GAP_X) + GAP_X
    dh = rows * (B_H + GAP_Y) + GAP_Y + 22  # 22 = top label
    d = Drawing(dw, dh)

    # blackboard
    board_w = min(dw * 0.6, 200)
    bx = (dw - board_w) / 2
    d.add(Rect(bx, dh - 20, board_w, 14,
               fillColor=colors.HexColor("#2e7d32"),
               strokeColor=colors.HexColor("#1b5e20"), strokeWidth=1))
    d.add(String(dw / 2, dh - 13, "BLACKBOARD",
                 fontName="Helvetica-Bold", fontSize=7,
                 fillColor=colors.white, textAnchor="middle"))

    for idx, bench in enumerate(benches):
        col = idx % COLS
        row = idx // COLS
        x = GAP_X + col * (B_W + GAP_X)
        y = dh - 22 - (row + 1) * (B_H + GAP_Y)

        # Determine fill color from class of first occupied seat
        cls = next((s["class"] for s in bench if s), "V")
        fill_hex = CLASS_COLORS.get(cls, "#90caf9")
        fill = colors.HexColor(fill_hex)
        light = colors.HexColor(fill_hex + "55")  # approximate lighter

        d.add(Rect(x, y, B_W, B_H,
                   fillColor=colors.HexColor("#e3f2fd"),
                   strokeColor=colors.HexColor("#90caf9"), strokeWidth=0.8))

        # Bench label
        d.add(String(x + B_W / 2, y + B_H - 9,
                     f"B{idx+1}",
                     fontName="Helvetica-Bold", fontSize=6,
                     fillColor=colors.HexColor("#1a237e"), textAnchor="middle"))

        # 3 seat circles: left, middle, right
        seat_positions = [
            (x + 10, y + 10),
            (x + B_W / 2, y + 10),
            (x + B_W - 10, y + 10),
        ]
        seat_labels = ["L", "M", "R"]
        for sp_x, sp_y in seat_positions:
            d.add(Rect(sp_x - SEAT_R, sp_y - SEAT_R,
                       SEAT_R * 2, SEAT_R * 2,
                       fillColor=fill,
                       strokeColor=colors.HexColor("#37474f"), strokeWidth=0.6))

    return d


def generate_pdf(rooms: dict[int, list]) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=13 * mm, rightMargin=13 * mm,
        topMargin=13 * mm, bottomMargin=13 * mm,
    )

    st_exam    = _style("exam",    fontName="Helvetica-Bold", fontSize=15, spaceAfter=1*mm,
                         textColor=colors.HexColor("#0d1b2a"))
    st_room    = _style("room",    fontName="Helvetica-Bold", fontSize=22, spaceAfter=3*mm,
                         textColor=colors.HexColor("#0f3460"))
    st_section = _style("section", fontName="Helvetica-Bold", fontSize=9,  spaceAfter=2*mm,
                         textColor=colors.HexColor("#37474f"))
    st_footer  = _style("footer",  fontSize=7, textColor=colors.grey,
                         fontName="Helvetica-Oblique")

    GREY_STRIPE = colors.HexColor("#f5f8ff")
    HDR_DARK    = colors.HexColor("#0f3460")
    HDR_MID     = colors.HexColor("#16213e")

    LEFT_BG  = colors.HexColor("#e8f5e9")
    LEFT_BG2 = colors.HexColor("#c8e6c9")
    MID_BG   = colors.HexColor("#fff3e0")
    MID_BG2  = colors.HexColor("#ffe0b2")

    story = []

    for room_id in range(1, 6):
        if room_id > 1:
            story.append(PageBreak())

        students = rooms[room_id]
        benches  = create_bench_layout(students)

        # ── Header ──────────────────────────────────────────────────────────
        story.append(Paragraph(SCHOOL_NAME, st_exam))
        story.append(Paragraph(f"ROOM  NO. {room_id}", st_room))
        story.append(HRFlowable(width="100%", thickness=2,
                                 color=HDR_DARK, spaceAfter=3*mm))

        # ── Class Summary Table ──────────────────────────────────────────────
        counts = defaultdict(lambda: {"BOYS": 0, "GIRLS": 0})
        for s in students:
            counts[s["class"]][s["gender"]] += 1

        hdr = [["Class", "Boys", "Girls", "Total"]]
        body, tb, tg = [], 0, 0
        for cls in CLASS_ORDER:
            if cls in counts:
                b = counts[cls]["BOYS"]
                g = counts[cls]["GIRLS"]
                body.append([f"Class {cls}",
                              str(b) if b else "–",
                              str(g) if g else "–",
                              str(b + g)])
                tb += b; tg += g
        body.append(["TOTAL", str(tb), str(tg), str(tb + tg)])

        summary_tbl = Table(hdr + body, colWidths=[35*mm, 24*mm, 24*mm, 24*mm],
                            hAlign="CENTER")
        summary_tbl.setStyle(TableStyle([
            ("BACKGROUND",  (0, 0), (-1, 0), HDR_DARK),
            ("TEXTCOLOR",   (0, 0), (-1, 0), colors.white),
            ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME",    (0,-1), (-1,-1), "Helvetica-Bold"),
            ("BACKGROUND",  (0,-1), (-1,-1), colors.HexColor("#dde8f0")),
            ("FONTSIZE",    (0, 0), (-1,-1), 9),
            ("ALIGN",       (0, 0), (-1,-1), "CENTER"),
            ("VALIGN",      (0, 0), (-1,-1), "MIDDLE"),
            ("ROWHEIGHT",   (0, 0), (-1,-1), 7*mm),
            ("GRID",        (0, 0), (-1,-1), 0.5, colors.HexColor("#aaaaaa")),
            ("ROWBACKGROUNDS", (0, 1), (-1,-2), [colors.white, GREY_STRIPE]),
        ]))

        story.append(Paragraph("CLASS-WISE STUDENT COUNT", st_section))
        story.append(summary_tbl)
        story.append(Spacer(1, 4*mm))

        # ── Room Diagram ─────────────────────────────────────────────────────
        story.append(Paragraph("CLASSROOM LAYOUT  (each box = 1 bench)", st_section))
        diagram = _room_diagram(benches, room_id)
        story.append(diagram)
        story.append(Spacer(1, 4*mm))

        # ── Legend ───────────────────────────────────────────────────────────
        legend_data = [["Seat Position", "Rule", "Color in table"]]
        legend_data += [
            ["LEFT  &  RIGHT", "Same class as each other", "Green tint"],
            ["MIDDLE",         "Different class from L/R",  "Orange tint"],
        ]
        legend_tbl = Table(legend_data, colWidths=[40*mm, 65*mm, 40*mm], hAlign="CENTER")
        legend_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), HDR_MID),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE",   (0, 0), (-1,-1), 8),
            ("ALIGN",      (0, 0), (-1,-1), "CENTER"),
            ("VALIGN",     (0, 0), (-1,-1), "MIDDLE"),
            ("ROWHEIGHT",  (0, 0), (-1,-1), 6*mm),
            ("GRID",       (0, 0), (-1,-1), 0.5, colors.grey),
            ("BACKGROUND", (0, 1), (0, 1), LEFT_BG),
            ("BACKGROUND", (0, 2), (0, 2), MID_BG),
        ]))
        story.append(legend_tbl)
        story.append(Spacer(1, 4*mm))

        # ── Bench-by-Bench Seating Table ─────────────────────────────────────
        story.append(Paragraph("BENCH-WISE SEATING ARRANGEMENT", st_section))

        bench_hdr = [["Bench\nNo.",
                      "LEFT SEAT\n(Roll | Class | Gender | Name)",
                      "MIDDLE SEAT\n(Roll | Class | Gender | Name)",
                      "RIGHT SEAT\n(Roll | Class | Gender | Name)"]]

        bench_rows = []
        for i, bench in enumerate(benches):
            bench_rows.append([str(i + 1)] + [_seat_cell(s) for s in bench])

        bench_tbl = Table(
            bench_hdr + bench_rows,
            colWidths=[12*mm, 53*mm, 53*mm, 53*mm],
            repeatRows=1,
        )

        ts = [
            ("BACKGROUND", (0, 0), (-1, 0), HDR_MID),
            ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
            ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME",   (0, 1), (0, -1), "Helvetica-Bold"),
            ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#e8eaf6")),
            ("FONTSIZE",   (0, 0), (-1, -1), 7.5),
            ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("ROWHEIGHT",  (0, 0), (0,  0), 14*mm),
            ("ROWHEIGHT",  (0, 1), (-1, -1), 13*mm),
            ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#c0c0c0")),
        ]
        for r in range(1, len(bench_rows) + 1):
            even = r % 2 == 0
            ts += [
                ("BACKGROUND", (1, r), (1, r), LEFT_BG2 if even else LEFT_BG),
                ("BACKGROUND", (2, r), (2, r), MID_BG2  if even else MID_BG),
                ("BACKGROUND", (3, r), (3, r), LEFT_BG2 if even else LEFT_BG),
            ]
        bench_tbl.setStyle(TableStyle(ts))
        story.append(bench_tbl)

        # ── Footer ───────────────────────────────────────────────────────────
        story.append(Spacer(1, 3*mm))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#aaaaaa")))
        story.append(Paragraph(
            f"Room {room_id}  ·  Total Students: {len(students)}  ·  "
            f"Total Benches: {len(benches)}  ·  {SCHOOL_NAME}",
            st_footer,
        ))

    doc.build(story)
    buf.seek(0)
    return buf


# ─── Streamlit UI ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Seating Arrangement Generator",
        page_icon="🏫",
        layout="wide",
    )

    # ── Custom CSS ────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
        .main-title   { font-size:2.2rem; font-weight:800; color:#0f3460; margin-bottom:0; }
        .sub-title    { font-size:1rem;   color:#555;      margin-bottom:1.5rem; }
        .metric-card  { background:#f0f4ff; border-radius:10px; padding:16px 20px;
                        border-left:4px solid #0f3460; margin-bottom:8px; }
        .metric-card h3 { margin:0; font-size:2rem; color:#0f3460; }
        .metric-card p  { margin:0; color:#666; font-size:0.85rem; }
        .room-badge { background:#0f3460; color:white; padding:2px 10px;
                      border-radius:12px; font-size:0.8rem; font-weight:600; }
        .rule-box { background:#fffde7; border:1px solid #f9a825;
                    border-radius:8px; padding:12px 16px; margin:8px 0; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<p class="main-title">🏫 Seating Arrangement Generator</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">1st Summative Evaluation 2026 — Automated PDF Generation</p>',
                unsafe_allow_html=True)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("ℹ️ How It Works")
        st.markdown("""
**Step 1** — Upload the Excel file  
**Step 2** — Review the auto-detected student counts  
**Step 3** — Click **Generate PDF**  
**Step 4** — Download the seating PDF

---
**Room Rules:**
- Class IX & X **Boys** → Rooms 2, 3, 4  
- Class IX & X **Girls** → Rooms 1, 5  
- Classes V–VIII → All 5 rooms  

**Bench Rule:**
- Left & Right seats = **same class**  
- Middle seat = **different class**
        """)
        st.markdown("---")
        st.caption("Template columns used: Class · Roll No · Name · Gender")

    # ── File Upload ───────────────────────────────────────────────────────────
    uploaded = st.file_uploader(
        "📂  Upload Student Excel File (.xlsx)",
        type=["xlsx"],
        help="Use the standard school Excel template — columns are auto-detected.",
    )

    if not uploaded:
        st.info("👆 Please upload the Excel file to begin.")
        return

    # ── Read Data ─────────────────────────────────────────────────────────────
    with st.spinner("Reading Excel..."):
        try:
            df = read_students(uploaded)
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

    st.success(f"✅  Loaded **{len(df):,}** students from '{uploaded.name}'")

    # ── Stats Row ─────────────────────────────────────────────────────────────
    st.subheader("📊 Student Summary")
    cols = st.columns(6)
    for i, cls in enumerate(CLASS_ORDER):
        sub = df[df["class"] == cls]
        with cols[i]:
            st.markdown(f"""
            <div class="metric-card">
              <h3>{len(sub)}</h3>
              <p>Class {cls}<br>
                 <span style="color:#4caf50">&#9646; {len(sub[sub.gender=='BOYS'])}B</span>&nbsp;
                 <span style="color:#e91e63">&#9646; {len(sub[sub.gender=='GIRLS'])}G</span>
              </p>
            </div>""", unsafe_allow_html=True)

    # ── Distribute to rooms ───────────────────────────────────────────────────
    rooms = distribute_to_rooms(df)

    # ── Room Preview Table ────────────────────────────────────────────────────
    st.subheader("🚪 Room Distribution Preview")

    room_rows = []
    for room_id in range(1, 6):
        counts = defaultdict(lambda: {"BOYS": 0, "GIRLS": 0})
        for s in rooms[room_id]:
            counts[s["class"]][s["gender"]] += 1
        row = {"Room": f"Room {room_id}"}
        for cls in CLASS_ORDER:
            b = counts[cls]["BOYS"]
            g = counts[cls]["GIRLS"]
            row[f"Cl.{cls} B"] = b if b else ""
            row[f"Cl.{cls} G"] = g if g else ""
        row["Total"] = len(rooms[room_id])
        room_rows.append(row)

    st.dataframe(pd.DataFrame(room_rows).set_index("Room"), use_container_width=True)

    st.markdown("""
    <div class="rule-box">
    🪑 <b>Seating rule applied:</b> &nbsp;
    Left &amp; Right seats of each bench belong to the <b>same class</b>.
    The Middle seat belongs to a <b>different class</b>.
    </div>
    """, unsafe_allow_html=True)

    # ── Generate PDF ──────────────────────────────────────────────────────────
    st.subheader("📄 Generate PDF")

    if st.button("🖨️  Generate Seating Arrangement PDF", type="primary",
                  use_container_width=True):
        with st.spinner("Generating PDF — please wait..."):
            try:
                pdf_buf = generate_pdf(rooms)
                st.success("✅  PDF generated successfully!")
                st.download_button(
                    label="📥  Download Seating Arrangement PDF",
                    data=pdf_buf,
                    file_name="seating_arrangement_2026.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"PDF generation failed: {e}")
                st.exception(e)


if __name__ == "__main__":
    main()
