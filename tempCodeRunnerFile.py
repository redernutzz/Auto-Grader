 w_weight = int(simpledialog.askstring("Input", "Written Works Weight (%):"))
        p_weight = int(simpledialog.askstring("Input", "Performance Task Weight (%):"))
        a_weight = int(simpledialog.askstring("Input", "Quarterly Assessment Weight (%):"))

        w_count = int(simpledialog.askstring("Input", "Number of Written Works Activities:"))
        w_perfect = list(map(int, simpledialog.askstring("Input", "Perfect Scores for Written Works (comma-separated):").split(',')))
        if len(w_perfect) != w_count:
            raise ValueError("Mismatch in written work activity count and scores.")

        p_count = int(simpledialog.askstring("Input", "Number of Performance Task Activities:"))
        p_perfect = list(map(int, simpledialog.askstring("Input", "Perfect Scores for Performance Tasks (comma-separated):").split(',')))