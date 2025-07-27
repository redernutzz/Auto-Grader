import tkinter as tk
from tkinter import messagebox, ttk
import random

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

# ================== GUI Setup ==================
root = tk.Tk()
root.title("Subject Grade Generator")

# ---------------- Subject Selection ----------------
tk.Label(root, text="Select Subject:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
subject_var = tk.StringVar()
subject_dropdown = ttk.Combobox(root, textvariable=subject_var, state="readonly")
subject_dropdown["values"] = list(subject_weights.keys())
subject_dropdown.grid(row=0, column=1, padx=5, pady=5)

# ---------------- Dynamic Form ----------------
form_frame = tk.Frame(root)
form_frame.grid(row=1, column=0, columnspan=2, pady=10)

fields = [
    "Written Count", "Written Perfects",
    "Performance Count", "Performance Perfects",
    "Assessment Count", "Assessment Perfects",
    "Target Grade"
]

entries = {}

def build_form():
    for widget in form_frame.winfo_children():
        widget.destroy()

    selected_subject = subject_var.get()
    if not selected_subject:
        return

    tk.Label(form_frame, text=f"Grading breakdown for {selected_subject}:",
             font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 10))

    weights = subject_weights[selected_subject]
    tk.Label(form_frame, text=f"Written Works: {weights['Written']}%", fg="blue").grid(row=1, column=0, sticky="w")
    tk.Label(form_frame, text=f"Performance Task: {weights['Performance']}%", fg="green").grid(row=1, column=1, sticky="w")
    tk.Label(form_frame, text=f"Quarterly Assessment: {weights['Assessment']}%", fg="purple").grid(row=1, column=2, sticky="w")

    for idx, field in enumerate(fields, start=2):
        label = tk.Label(form_frame, text=field + ":")
        label.grid(row=idx, column=0, sticky="e", padx=5, pady=3)
        entry = tk.Entry(form_frame, width=40)
        entry.grid(row=idx, column=1, columnspan=2, pady=3, padx=5, sticky="w")
        entries[field] = entry

    submit_btn = tk.Button(form_frame, text="Generate Grade", command=submit_form)
    submit_btn.grid(row=len(fields) + 2, column=0, columnspan=3, pady=10)

subject_dropdown.bind("<<ComboboxSelected>>", lambda e: build_form())

# ---------------- Form Submission ----------------
def submit_form():
    try:
        subject = subject_var.get()
        if not subject:
            raise ValueError("Please select a subject.")

        w_weight = subject_weights[subject]["Written"]
        p_weight = subject_weights[subject]["Performance"]
        a_weight = subject_weights[subject]["Assessment"]

        w_count = int(entries["Written Count"].get())
        w_perfect = list(map(int, entries["Written Perfects"].get().split(',')))
        if len(w_perfect) != w_count:
            raise ValueError("Mismatch in Written Work count and scores")

        p_count = int(entries["Performance Count"].get())
        p_perfect = list(map(int, entries["Performance Perfects"].get().split(',')))
        if len(p_perfect) != p_count:
            raise ValueError("Mismatch in Performance Task count and scores")

        a_count = int(entries["Assessment Count"].get())
        a_perfect = list(map(int, entries["Assessment Perfects"].get().split(',')))
        if len(a_perfect) != a_count:
            raise ValueError("Mismatch in Assessment count and scores")

        target_grade = float(entries["Target Grade"].get())

        result = find_combination(w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight, target_grade)

        if result:
            msg = f"🎯 Target Grade: {result['Final Grade']}%\n\n"
            for comp in ["Written Works", "Performance Task", "Quarterly Assessment"]:
                msg += f"{comp}:\n  Scores: {result[comp][0]}\n  Grade: {round(result[comp][1], 2)}%\n\n"
            messagebox.showinfo("Result", msg)
        else:
            messagebox.showwarning("No Match", "Couldn't find a matching score combination.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

root.mainloop()
