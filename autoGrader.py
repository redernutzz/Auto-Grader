import tkinter as tk
from tkinter import simpledialog, messagebox
import random

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

def run_program():
    try:
        w_weight = int(simpledialog.askstring("Input", "Written Works Weight (%):"))
        p_weight = int(simpledialog.askstring("Input", "Performance Task Weight (%):"))
        a_weight = int(simpledialog.askstring("Input", "Quarterly Assessment Weight (%):"))

        w_count = int(simpledialog.askstring("Input", "Number of Written Works Activities:"))
        w_perfect = list(map(int, simpledialog.askstring("Input", "Perfect Scores for Written Works (comma-separated):").split(',')))
        if len(w_perfect) != w_count:
            raise ValueError("Mismatch in written work activity count and scores.")

        p_count = int(simpledialog.askstring("Input", "Number of Performance Task Activities:"))
        p_perfect = list(map(int, simpledialog.askstring("Input", "Perfect Scores for Performance Tasks (comma-separated):").split(',')))
        if len(p_perfect) != p_count:
            raise ValueError("Mismatch in performance task activity count and scores.")

        a_count = int(simpledialog.askstring("Input", "Number of Quarterly Assessments:"))
        a_perfect = list(map(int, simpledialog.askstring("Input", "Perfect Scores for Quarterly Assessments (comma-separated):").split(',')))
        if len(a_perfect) != a_count:
            raise ValueError("Mismatch in assessment activity count and scores.")

        target_grade = float(simpledialog.askstring("Input", "Target Final Grade (%):"))

        result = find_combination(w_perfect, p_perfect, a_perfect, w_weight, p_weight, a_weight, target_grade)

        if result:
            message = f"Target Grade: {result['Final Grade']}%\n\n"
            for key in ["Written Works", "Performance Task", "Quarterly Assessment"]:
                message += f"{key}:\n  Scores: {result[key][0]}\n  Grade: {round(result[key][1], 2)}%\n\n"
            messagebox.showinfo("Grade Combination Found", message)
        else:
            messagebox.showwarning("No Match", "No suitable score combination found within the attempt limit.")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Run GUI
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    run_program()
