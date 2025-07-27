import itertools

# Hardcoded input example
written_weight = 30
performance_weight = 50
assessment_weight = 20

written_scores = [16, 12, 18, 18, 16]
performance_scores = [10, 10, 10, 10, 10]
assessment_scores = [6]

target_grade = 70.37

# Generate all possible score combinations for a component
def generate_score_combinations(perfect_scores):
    ranges = [range(p + 1) for p in perfect_scores]
    return list(itertools.product(*ranges))

# Calculate the grade percentage of a component
def calculate_component_grade(scores, perfect_scores):
    total_score = sum(scores)
    total_perfect = sum(perfect_scores)
    return (total_score / total_perfect) * 100

# Generate all combinations (brute-force, may take time for large inputs)
written_combos = generate_score_combinations(written_scores)
performance_combos = generate_score_combinations(performance_scores)
assessment_combos = generate_score_combinations(assessment_scores)

# Try all combinations to find a match
found = False
for w in written_combos:
    for p in performance_combos:
        for a in assessment_combos:
            w_grade = calculate_component_grade(w, written_scores)
            p_grade = calculate_component_grade(p, performance_scores)
            a_grade = calculate_component_grade(a, assessment_scores)
            final = (w_grade * written_weight + p_grade * performance_weight + a_grade * assessment_weight) / 100

            if round(final, 2) == round(target_grade, 2):
                print("\nSOLUTION FOUND:")
                print("Written scores:     ", w)
                print("Performance scores: ", p)
                print("Assessment scores:  ", a)
                print("Final Grade:        ", round(final, 2))
                found = True
                break
        if found:
            break
    if found:
        break

if not found:
    print("No combination found that matches the target grade.")
