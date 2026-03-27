import streamlit as st
import pandas as pd
from fpdf import FPDF
from collections import defaultdict

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Seating Arrangement Generator", layout="wide")

def clean_class_name(class_str):
    if pd.isna(class_str):
        return None
    return str(class_str).split(':')[0].strip()

def get_seating_plan(df):
    room_config = {
        "Room 1": [11, 11],
        "Room 2": [6, 6],
        "Room 3": [7, 6],
        "Room 4": [6, 6],
        "Room 5": [9, 9]
    }
    
    df['Roll No.'] = pd.to_numeric(df['Roll No.'], errors='coerce')
    df = df.sort_values(by=['Present Class', 'Roll No.'])
    
    class_queues = {}
    for cls in df['Present Class'].unique():
        class_students = df[df['Present Class'] == cls].to_dict('records')
        class_queues[cls] = class_students
    
    available_classes = list(class_queues.keys())
    
    seating_results = []
    raw_assignments = [] # To keep track of who is in which room for the PDF summary

    for room_name, columns in room_config.items():
        bench_number = 1
        for col_idx, benches_in_col in enumerate(columns):
            for _ in range(benches_in_col):
                available_classes.sort(key=lambda c: len(class_queues[c]), reverse=True)
                
                class_a, class_b = None, None
                
                if len(available_classes) > 0 and len(class_queues[available_classes[0]]) >= 2:
                    class_a = available_classes[0]
                elif len(available_classes) > 0 and len(class_queues[available_classes[0]]) > 0:
                    class_a = available_classes[0]  
                
                if len(available_classes) > 1 and len(class_queues[available_classes[1]]) > 0:
                    class_b = available_classes[1]
                elif len(available_classes) > 0 and len(class_queues[available_classes[0]]) > 0:
                     class_b = available_classes[0] 
                
                seat_left, seat_middle, seat_right = None, None, None
                
                if class_a and len(class_queues[class_a]) > 0:
                    seat_left = class_queues[class_a].pop(0)
                    raw_assignments.append({"Room": room_name, "Class": seat_left['Present Class'], "Gender": seat_left['Gender'], "Roll": seat_left['Roll No.']})
                if class_a and len(class_queues[class_a]) > 0:
                    seat_right = class_queues[class_a].pop(0)
                    raw_assignments.append({"Room": room_name, "Class": seat_right['Present Class'], "Gender": seat_right['Gender'], "Roll": seat_right['Roll No.']})
                    
                if class_b and len(class_queues[class_b]) > 0:
                    seat_middle = class_queues[class_b].pop(0)
                    raw_assignments.append({"Room": room_name, "Class": seat_middle['Present Class'], "Gender": seat_middle['Gender'], "Roll": seat_middle['Roll No.']})
                
                seating_results.append({
                    "Room": room_name,
                    "Column": f"Column {col_idx + 1}",
                    "Bench No": bench_number,
                    "Left Seat": f"{seat_left['Present Class']} - {int(seat_left['Roll No.'])}" if seat_left else "Empty",
                    "Middle Seat": f"{seat_middle['Present Class']} - {int(seat_middle['Roll No.'])}" if seat_middle else "Empty",
                    "Right Seat": f"{seat_right['Present Class']} - {int(seat_right['Roll No.'])}" if seat_right else "Empty",
                })
                
                bench_number += 1
                
                available_classes = [c for c in available_classes if len(class_queues[c]) > 0]
                if not available_classes: break
            if not available_classes: break
        if not available_classes: break

    unassigned = {cls: len(q) for cls, q in class_queues.items() if len(q) > 0}
    return pd.DataFrame(seating_results), raw_assignments, unassigned

def create_pdf(seating_df, raw_assignments):
    """Generates a PDF with one room per page, summary lists, and bordered tables."""
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    
    rooms = seating_df['Room'].unique()
    
    for room in rooms:
        pdf.add_page()
        
        # --- TITLE ---
        pdf.set_font("helvetica", "B", 16)
        pdf.cell(0, 8, "1st Summative Evaluation 2026", align="C", ln=1)
        pdf.set_font("helvetica", "B", 14)
        pdf.cell(0, 8, f"{room}", align="C", ln=1)
        pdf.ln(5)
        
        # --- SUMMARY LIST (Roll Numbers by Class & Gender) ---
        pdf.set_font("helvetica", "", 12)
        room_students = [s for s in raw_assignments if s['Room'] == room]
        
        # Group data
        summary = defaultdict(lambda: defaultdict(list))
        for s in room_students:
            summary[s['Class']][s['Gender']].append(int(s['Roll']))
            
        for cls in sorted(summary.keys()):
            for gender in sorted(summary[cls].keys()):
                rolls = sorted(summary[cls][gender])
                roll_str = ", ".join(map(str, rolls))
                count = len(rolls)
                
                # Format: CLASS - V: BOYS - 1, 2, 3 = 03
                line = f"CLASS - {cls}: {gender} - {roll_str} = {count:02d}"
                pdf.multi_cell(0, 8, line)
                
        pdf.ln(10)
        
        # --- SEATING TABLE WITH IMAGE OUTLINE (BORDERS) ---
        room_df = seating_df[seating_df['Room'] == room]
        
        for col_name in room_df['Column'].unique():
            # Column Title
            pdf.set_font("helvetica", "B", 12)
            pdf.cell(0, 8, f"--- {col_name} ---", align="C", ln=1)
            
            col_df = room_df[room_df['Column'] == col_name]
            
            # Table Headers
            pdf.set_font("helvetica", "B", 10)
            col_widths = [20, 55, 55, 55] # Total = 185mm (fits A4 nicely)
            headers = ["Bench", "Left Seat", "Middle Seat", "Right Seat"]
            
            for i, header in enumerate(headers):
                pdf.cell(col_widths[i], 10, header, border=1, align="C")
            pdf.ln()
            
            # Table Rows
            pdf.set_font("helvetica", "", 10)
            for _, row in col_df.iterrows():
                pdf.cell(col_widths[0], 10, str(row['Bench No']), border=1, align="C")
                pdf.cell(col_widths[1], 10, str(row['Left Seat']), border=1, align="C")
                pdf.cell(col_widths[2], 10, str(row['Middle Seat']), border=1, align="C")
                pdf.cell(col_widths[3], 10, str(row['Right Seat']), border=1, align="C")
                pdf.ln()
                
            pdf.ln(5) # Space between columns
            
    # Output to byte string for Streamlit download
    return bytes(pdf.output())

# --- UI START ---
st.title("🏫 Examination Seating Arrangement Automation")
st.markdown("Upload your Student Excel (`.xlsx`) file to generate seating arrays and formatted PDF prints.")

uploaded_file = st.file_uploader("Upload Student Data", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=1)
        else:
            df = pd.read_excel(uploaded_file, header=1) 
        
        df = df.dropna(subset=['Present Class', 'Roll No.'])
        df['Present Class'] = df['Present Class'].apply(clean_class_name)
        
        st.success("File successfully loaded!")
        
        if st.button("Generate Seating Arrangement"):
            with st.spinner("Calculating optimal seating arrangement & rendering PDF..."):
                seating_df, raw_assignments, unassigned = get_seating_plan(df)
                
                # Generate the PDF byte data
                pdf_bytes = create_pdf(seating_df, raw_assignments)
            
            if unassigned:
                st.warning(f"Not enough seats available! Unassigned students: {unassigned}")
            else:
                st.success("All students successfully assigned!")
            
            # --- DOWNLOAD BUTTONS ---
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📄 Download Formatted PDF (Ready to Print)",
                    data=pdf_bytes,
                    file_name='Exam_Seating_Arrangement.pdf',
                    mime='application/pdf',
                )
                
            with col2:
                csv = seating_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Raw Data (CSV)",
                    data=csv,
                    file_name='seating_data.csv',
                    mime='text/csv',
                )
                
            # Display preview in Streamlit
            st.header("Preview Data")
            st.dataframe(seating_df, use_container_width=True)

    except Exception as e:
        st.error(f"An error occurred: {e}")
