import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import random
import pandas as pd
from datetime import datetime
import os

# ================== Subject-specific percentage weights ==================
subject_weights = {
    "Math": {"Written": 40, "Performance": 40, "Assessment": 20},
    "English": {"Written": 30, "Performance": 50, "Assessment": 20},
    "Science": {"Written": 35, "Performance": 45, "Assessment": 20},
    "Filipino": {"Written": 25, "Performance": 50, "Assessment": 25},
    "Araling Panlipunan": {"Written": 30, "Performance": 50, "Assessment": 20},
}

# ================== Grade Logic ==================
def calculate_component_grade(scores, perfect_scores):
    return (sum(scores) / sum(perfect_scores)) * 100 if sum(perfect_scores) > 0 else 0

def find_combination(w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight, target_grade, tolerance=0.01, max_attempts=100000):
    for _ in range(max_attempts):
        w_scores = [random.randint(0, p) for p in w_perfect]
        p_scores = [random.randint(0, p) for p in p_perfect]
        a_scores = [random.randint(0, p) for p in a_perfect]

        w_grade = calculate_component_grade(w_scores, w_perfect)
        p_grade = calculate_component_grade(p_scores, p_perfect)
        a_grade = calculate_component_grade(a_scores, a_perfect)

        final = (w_grade * w_weight + p_grade * p_weight + a_grade * a_weight) / 100

        if abs(final - target_grade) <= tolerance:
            return {
                "Written Works": (w_scores, w_grade),
                "Performance Task": (p_scores, p_grade),
                "Quarterly Assessment": (a_scores, a_grade),
                "Final Grade": round(final, 2)
            }
    return None

def create_excel_file(results, subject, w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight):
    """Create an Excel file with the grades in the specified format"""
    try:
        # Ask user where to save the file
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Excel File",
            initialname=f"{subject}_Grades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not filename:
            return False
            
        # Create the Excel data structure
        data = []
        
        # Header row 1 - Component names and weights
        header1 = [''] + [f'WRITTEN WORKS ({w_weight}%)'] + [''] * (len(w_perfect) + 1) + \
                 [f'PERFORMANCE TASKS ({p_weight}%)'] + [''] * (len(p_perfect) + 1) + \
                 [f'QUARTERLY ASSESSMENT ({a_weight}%)', '', 'Initial']
        
        # Header row 2 - Activity numbers and totals
        header2 = [''] + [str(i+1) for i in range(len(w_perfect))] + ['Total', 'PS', 'WS'] + \
                 [str(i+1) for i in range(len(p_perfect))] + ['Total', 'PS', 'WS'] + \
                 ['1', 'PS', 'WS', 'Grade']
        
        # Perfect scores row
        perfect_row = [''] + w_perfect + [sum(w_perfect), 100.00, f'{w_weight}%'] + \
                     p_perfect + [sum(p_perfect), 100.00, f'{p_weight}%'] + \
                     [sum(a_perfect), 100.00, f'{a_weight}%', '']
        
        # Add headers and perfect scores
        data.append(header1)
        data.append(header2)
        data.append(perfect_row)
        
        # Add student data
        for i, result in enumerate(results, 1):
            if "Error" not in result:
                # Written Works data
                w_scores = result["Written Works"][0]
                w_grade = result["Written Works"][1]
                w_total = sum(w_scores)
                
                # Performance Tasks data
                p_scores = result["Performance Task"][0]
                p_grade = result["Performance Task"][1]
                p_total = sum(p_scores)
                
                # Assessment data
                a_scores = result["Quarterly Assessment"][0]
                a_grade = result["Quarterly Assessment"][1]
                a_total = sum(a_scores)
                
                # Final grade
                final_grade = result["Final Grade"]
                
                # Create student row
                student_row = [str(i)] + w_scores + [w_total, f'{w_grade:.2f}', f'{(w_grade * w_weight / 100):.2f}'] + \
                             p_scores + [p_total, f'{p_grade:.2f}', f'{(p_grade * p_weight / 100):.2f}'] + \
                             a_scores + [f'{a_grade:.2f}', f'{(a_grade * a_weight / 100):.2f}', f'{final_grade:.2f}']
                
                data.append(student_row)
            else:
                # Handle error case
                error_row = [str(i)] + ['ERROR'] * (len(header2) - 1)
                data.append(error_row)
        
        # Create DataFrame
        df = pd.DataFrame(data)
        
        # Save to Excel with formatting
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=f'{subject} Grades', index=False, header=False)
            
            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets[f'{subject} Grades']
            
            # Apply some basic formatting
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
            
            # Header formatting
            header_font = Font(bold=True, size=10)
            header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            center_align = Alignment(horizontal="center", vertical="center")
            
            # Apply formatting to headers
            for row in range(1, 4):  # First 3 rows are headers
                for col in range(1, df.shape[1] + 1):
                    cell = worksheet.cell(row=row, column=col)
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = center_align
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 15)  # Max width of 15
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        messagebox.showinfo("Success", f"Excel file created successfully:\n{filename}")
        return True
        
    except ImportError:
        messagebox.showerror("Error", "Required libraries not found. Please install:\npip install pandas openpyxl")
        return False
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create Excel file:\n{str(e)}")
        return False

# ================== GUI ==================
root = tk.Tk()
root.title("Multi-Student Grade Generator")
root.geometry("800x600")

# Globals to retain across screens
entries = {}
target_entries = []
subject_var = tk.StringVar()
generated_results = []  # Store results for Excel export

def build_initial_form():
    for widget in root.winfo_children():
        widget.destroy()

    tk.Label(root, text="Select Subject:", font=("Arial", 14)).grid(row=0, column=0, sticky="e", padx=10, pady=10)
    dropdown = ttk.Combobox(root, textvariable=subject_var, state="readonly", font=("Arial", 14))
    dropdown["values"] = list(subject_weights.keys())
    dropdown.grid(row=0, column=1, padx=10, pady=10)
    dropdown.bind("<<ComboboxSelected>>", lambda e: build_form())

def build_form():
    for widget in root.winfo_children()[2:]:  # Keep dropdown and label
        widget.destroy()

    form_labels = [
        ("Written Count", 1), ("Written Perfects", 2),
        ("Performance Count", 3), ("Performance Perfects", 4),
        ("Assessment Count", 5), ("Assessment Perfects", 6),
        ("Number of Students", 7)
    ]

    for label, row in form_labels:
        tk.Label(root, text=label + ":", font=("Arial", 14)).grid(row=row, column=0, sticky="e", padx=10, pady=5)
        entry = tk.Entry(root, font=("Arial", 14), width=40)
        entry.grid(row=row, column=1, pady=5, padx=5)
        entries[label] = entry

    next_btn = tk.Button(root, text="Next", font=("Arial", 14, "bold"), bg="#add8e6", command=build_targets)
    next_btn.grid(row=8, column=0, columnspan=2, pady=20)

def build_targets():
    global w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight

    try:
        subject = subject_var.get()
        weights = subject_weights[subject]
        w_weight = weights["Written"]
        p_weight = weights["Performance"]
        a_weight = weights["Assessment"]

        w_count = int(entries["Written Count"].get())
        w_perfect = list(map(int, entries["Written Perfects"].get().split(',')))

        p_count = int(entries["Performance Count"].get())
        p_perfect = list(map(int, entries["Performance Perfects"].get().split(',')))

        a_count = int(entries["Assessment Count"].get())
        a_perfect = list(map(int, entries["Assessment Perfects"].get().split(',')))

        num_students = int(entries["Number of Students"].get())
        if len(w_perfect) != w_count or len(p_perfect) != p_count or len(a_perfect) != a_count:
            raise ValueError("Mismatch in activity count and scores.")

    except Exception as e:
        messagebox.showerror("Input Error", str(e))
        return

    for widget in root.winfo_children():
        widget.destroy()

    tk.Label(root, text=f"Enter Target Grades for {num_students} Student(s):", font=("Arial", 16, "bold")).pack(pady=20)
    frame = tk.Frame(root)
    frame.pack()

    target_entries.clear()
    for i in range(num_students):
        tk.Label(frame, text=f"Student {i+1} Target Grade:", font=("Arial", 14)).grid(row=i, column=0, padx=10, pady=5)
        entry = tk.Entry(frame, font=("Arial", 14))
        entry.grid(row=i, column=1, padx=10, pady=5)
        target_entries.append(entry)

    button_frame = tk.Frame(root)
    button_frame.pack(pady=20)
    
    generate_btn = tk.Button(button_frame, text="Generate Grades", font=("Arial", 14, "bold"), bg="#add8e6", command=generate_grades)
    generate_btn.pack(side=tk.LEFT, padx=10)

def generate_grades():
    global generated_results
    
    try:
        results = []
        for entry in target_entries:
            grade = float(entry.get())
            result = find_combination(w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight, grade)
            if result:
                results.append(result)
            else:
                results.append({"Final Grade": grade, "Error": "No matching combination found."})

        generated_results = results  # Store for Excel export
        
        msg = ""
        for i, res in enumerate(results):
            msg += f"Student {i+1} Target Grade: {res['Final Grade']}%\n"
            if "Error" in res:
                msg += f"  {res['Error']}\n"
            else:
                for comp in ["Written Works", "Performance Task", "Quarterly Assessment"]:
                    msg += f"  {comp}: Scores: {res[comp][0]}, Grade: {round(res[comp][1], 2)}%\n"
                msg += "\n"

        # Show results and add Excel export button
        result_window = tk.Toplevel(root)
        result_window.title("Generated Grades")
        result_window.geometry("600x400")
        
        text_widget = tk.Text(result_window, wrap=tk.WORD, font=("Arial", 10))
        scrollbar = tk.Scrollbar(result_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        text_widget.insert("1.0", msg)
        text_widget.config(state="disabled")
        
        # Add Excel export button
        export_btn = tk.Button(result_window, text="Export to Excel", font=("Arial", 12, "bold"), 
                              bg="#90EE90", command=lambda: export_to_excel())
        export_btn.pack(pady=10)
        
    except Exception as e:
        messagebox.showerror("Error", str(e))

def export_to_excel():
    """Export the generated results to Excel"""
    if not generated_results:
        messagebox.showwarning("Warning", "No grades generated yet!")
        return
    
    subject = subject_var.get()
    success = create_excel_file(generated_results, subject, w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight)
    
    if success:
        # Ask if user wants to open the file location
        response = messagebox.askyesno("Success", "Excel file created successfully!\nWould you like to open the file location?")
        if response:
            try:
                import subprocess
                import os
                # Open the directory containing the file (works on Windows, Mac, Linux)
                if os.name == 'nt':  # Windows
                    subprocess.run(['explorer', os.path.dirname(filename)], check=True)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['open' if 'darwin' in os.sys.platform else 'xdg-open', 
                                   os.path.dirname(filename)], check=True)
            except:
                pass  # Silently fail if can't open directory

build_initial_form()
root.mainloop()