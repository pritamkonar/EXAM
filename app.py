import streamlit as st
import pandas as pd
import math

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Seating Arrangement Generator", layout="wide")

def clean_class_name(class_str):
    """ Extract the Roman numeral class name (e.g., 'V : 2026' -> 'V') """
    if pd.isna(class_str):
        return None
    return str(class_str).split(':')[0].strip()

def get_seating_plan(df):
    """ Core logic to assign students to rooms based on bench capacities """
    
    # Define Room Configuration (Benches per column)
    room_config = {
        "Room 1": [11, 11],
        "Room 2": [6, 6],
        "Room 3": [7, 6],
        "Room 4": [6, 6],
        "Room 5": [9, 9]
    }
    
    # Group students by Class and Gender
    # Ensure they are sorted by Roll No.
    df['Roll No.'] = pd.to_numeric(df['Roll No.'], errors='coerce')
    df = df.sort_values(by=['Present Class', 'Roll No.'])
    
    # Store students in a queue-like structure by Class
    class_queues = {}
    for cls in df['Present Class'].unique():
        class_students = df[df['Present Class'] == cls].to_dict('records')
        class_queues[cls] = class_students
    
    available_classes = list(class_queues.keys())
    
    seating_results = []

    # Assign students room by room
    for room_name, columns in room_config.items():
        total_benches = sum(columns)
        
        bench_number = 1
        for col_idx, benches_in_col in enumerate(columns):
            for _ in range(benches_in_col):
                
                # We need 2 students from Class A (Ends) and 1 from Class B (Middle)
                # Find the class with the most unassigned students for the Ends
                available_classes.sort(key=lambda c: len(class_queues[c]), reverse=True)
                
                class_a = None
                class_b = None
                
                if len(available_classes) > 0 and len(class_queues[available_classes[0]]) >= 2:
                    class_a = available_classes[0]
                elif len(available_classes) > 0 and len(class_queues[available_classes[0]]) > 0:
                    class_a = available_classes[0]  # Take whatever is left
                
                if len(available_classes) > 1 and len(class_queues[available_classes[1]]) > 0:
                    class_b = available_classes[1]
                elif len(available_classes) > 0 and len(class_queues[available_classes[0]]) > 0:
                     class_b = available_classes[0] # Fallback
                
                seat_left, seat_middle, seat_right = None, None, None
                
                # Pop students for Left and Right seats (Class A)
                if class_a and len(class_queues[class_a]) > 0:
                    seat_left = class_queues[class_a].pop(0)
                if class_a and len(class_queues[class_a]) > 0:
                    seat_right = class_queues[class_a].pop(0)
                    
                # Pop student for Middle seat (Class B)
                if class_b and len(class_queues[class_b]) > 0:
                    seat_middle = class_queues[class_b].pop(0)
                
                # Save the bench assignment
                seating_results.append({
                    "Room": room_name,
                    "Column": f"Column {col_idx + 1}",
                    "Bench No": bench_number,
                    "Left Seat (Class/Roll/Gender)": f"{seat_left['Present Class']} - {seat_left['Roll No.']} ({seat_left['Gender']})" if seat_left else "Empty",
                    "Left Seat Name": seat_left['Name of Student'] if seat_left else "",
                    
                    "Middle Seat (Class/Roll/Gender)": f"{seat_middle['Present Class']} - {seat_middle['Roll No.']} ({seat_middle['Gender']})" if seat_middle else "Empty",
                    "Middle Seat Name": seat_middle['Name of Student'] if seat_middle else "",
                    
                    "Right Seat (Class/Roll/Gender)": f"{seat_right['Present Class']} - {seat_right['Roll No.']} ({seat_right['Gender']})" if seat_right else "Empty",
                    "Right Seat Name": seat_right['Name of Student'] if seat_right else "",
                })
                
                bench_number += 1
                
                # Cleanup empty classes
                available_classes = [c for c in available_classes if len(class_queues[c]) > 0]
                if not available_classes:
                    break
            if not available_classes:
                break
        if not available_classes:
            break

    # Check for unassigned students
    unassigned = {cls: len(q) for cls, q in class_queues.items() if len(q) > 0}
    
    return pd.DataFrame(seating_results), unassigned

# --- UI START ---
st.title("🏫 Examination Seating Arrangement Automation")
st.markdown("Upload your Student Excel (`.xlsx`) file. The app will fetch the students, categorize them by class/gender, and generate the seating layout.")

uploaded_file = st.file_uploader("Upload Student Data (Excel Format)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=1) # Accounting for the double header in your files
        else:
            df = pd.read_excel(uploaded_file, header=1) 
        
        # Keep only essential columns and drop purely empty rows
        required_cols = ['Present Class', 'Roll No.', 'Gender', 'Name of Student']
        
        # Clean the dataset
        df = df.dropna(subset=['Present Class', 'Roll No.'])
        df['Present Class'] = df['Present Class'].apply(clean_class_name)
        
        st.success("File successfully loaded and processed!")
        
        # --- SUMMARY STATISTICS ---
        st.header("📊 Student Summary")
        summary_df = df.groupby(['Present Class', 'Gender']).size().reset_index(name='Total Students')
        
        # Pivot the table for better UI
        pivot_summary = summary_df.pivot(index='Present Class', columns='Gender', values='Total Students').fillna(0).astype(int)
        pivot_summary['Total'] = pivot_summary.sum(axis=1)
        st.table(pivot_summary)
        
        # --- SEATING ARRANGEMENT GENERATION ---
        st.header("🪑 Generate Seating Plan")
        if st.button("Generate Seating Arrangement"):
            with st.spinner("Calculating optimal seating arrangement..."):
                seating_df, unassigned = get_seating_plan(df)
            
            # Display Unassigned if any
            if unassigned:
                st.warning(f"Not enough seats available! Unassigned students: {unassigned}")
            else:
                st.success("All students successfully assigned!")
            
            # Display Room by Room
            rooms = seating_df['Room'].unique()
            for room in rooms:
                st.subheader(room)
                room_df = seating_df[seating_df['Room'] == room]
                st.dataframe(room_df.drop(columns=['Room']), use_container_width=True)
                
            # Download button
            csv = seating_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Complete Seating Plan (CSV)",
                data=csv,
                file_name='seating_arrangement.csv',
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"An error occurred while reading the file: {e}")
        st.write("Please ensure the uploaded Excel file follows the standard template with headers like 'Present Class', 'Roll No.', 'Gender', and 'Name of Student'.")