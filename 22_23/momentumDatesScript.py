import pandas as pd
import random

def compute_empirical_score_at_indices(sequence, indices, score_map):
    if not indices:
        return None, 0
    scores = [score_map.get(sequence[i], 0.0) for i in indices]
    return sum(scores) / len(scores), len(indices)

def simulate_bias(p, n, k, target, simulations=5000, score_map=None):
    empirical_scores = []
    choices = [target, "D" if target == "W" else "W", "L" if target != "L" else "D"]
    for _ in range(simulations):
        seq = [random.choices(choices, weights=[p, (1 - p)/2, (1 - p)/2])[0] for _ in range(n)]
        # Find matching indices
        indices = [i for i in range(k, len(seq)) if all(seq[i-j-1] == target for j in range(k))]
        score, _ = compute_empirical_score_at_indices(seq, indices, score_map)
        if score is not None:
            empirical_scores.append(score)
    return sum(empirical_scores) / len(empirical_scores) if empirical_scores else 0

def correct_for_bias(empirical, p, n, k, target, score_map):
    estimated_bias = simulate_bias(p, n, k, target, score_map=score_map)
    correction = estimated_bias - p
    corrected = empirical + correction
    return corrected, estimated_bias

def analyze_tight_vs_loose(file_path):
    xls = pd.ExcelFile(file_path)
    records = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if "W/D/L" not in df.columns or "Tight Schedule" not in df.columns:
            continue

        results = df["W/D/L"].tolist()
        tight_flags = df["Tight Schedule"].tolist()
        n = len(results)

        for momentum_type, target, score_map in [
            ("Positive (WW)", "W", {"W": 1.0, "D": 0.5, "L": 0.0}),
            ("Negative (LL)", "L", {"L": 1.0, "D": 0.5, "W": 0.0})
        ]:
            for tight_value in [True, False]:
                indices = []
                for i in range(2, len(results)):
                    if tight_flags[i] != tight_value:
                        continue
                    if results[i-1] == target and results[i-2] == target:
                        indices.append(i)

                if not indices:
                    continue

                p = results.count(target) / n if n > 0 else 0.5
                empirical_score, count = compute_empirical_score_at_indices(results, indices, score_map)
                corrected_score, bias = correct_for_bias(empirical_score, p, len(results), 2, target, score_map)

                records.append({
                    "Team": sheet_name,
                    "Momentum": momentum_type,
                    "Tight Schedule": tight_value,
                    "Empirical Score": round(empirical_score, 3),
                    "Corrected Score": round(corrected_score, 3),
                    "Bias": round(bias - p, 3),
                    "Count": count
                })

    df_results = pd.DataFrame(records)
    if df_results.empty:
        print("No momentum events found.")
        return

    # Group and print summary
    summary = df_results.groupby(["Momentum", "Tight Schedule"]).agg(
        Avg_Empirical=("Empirical Score", "mean"),
        Avg_Corrected=("Corrected Score", "mean"),
        Total_Count=("Count", "sum")
    ).reset_index()

    print("\n📊 Momentum Analysis (Corrected) by Tight Schedule:\n")
    for _, row in summary.iterrows():
        print(f"{row['Momentum']} | Tight={row['Tight Schedule']} → "
              f"Empirical: {row['Avg_Empirical'] * 100:.1f}%, "
              f"Corrected: {row['Avg_Corrected'] * 100:.1f}% (n={row['Total_Count']})")
# Example usage:
file_path = "PLteamsData22:23_with_tight_schedule.xlsx"
analyze_tight_vs_loose(file_path)
