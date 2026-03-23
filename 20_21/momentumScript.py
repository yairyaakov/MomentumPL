import pandas as pd
import random

def compute_empirical_score(sequence, k, target, score_map=None):
    indices = []
    for i in range(k, len(sequence)):
        if all(sequence[i - j - 1] == target for j in range(k)):
            indices.append(i)
    if not indices:
        return None, 0
    scores = [score_map.get(sequence[i], 0.0) for i in indices]
    return sum(scores) / len(scores), len(indices)

def simulate_bias(p, n, k, target, simulations=5000, score_map=None):
    empirical_scores = []
    choices = [target, "D" if target == "W" else "W", "L" if target != "L" else "D"]
    for _ in range(simulations):
        seq = [random.choices(choices, weights=[p, (1 - p) / 2, (1 - p) / 2])[0] for _ in range(n)]
        score, _ = compute_empirical_score(seq, k, target=target,score_map=score_map)
        if score is not None:
            empirical_scores.append(score)
    return sum(empirical_scores) / len(empirical_scores) if empirical_scores else 0

def correct_for_bias(empirical, p, n, k, target,score_map=None):
    estimated_bias = simulate_bias(p, n, k, target, score_map=score_map)
    correction = estimated_bias - p
    corrected = empirical + correction
    return corrected, estimated_bias

def analyze_success_vs_momentum(file_path):
    xls = pd.ExcelFile(file_path)
    results_summary = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if "W/D/L" not in df.columns:
            continue

        results = df["W/D/L"].dropna().tolist()
        n = len(results)
        if n < 5:
            continue

        score_map = {"W": 1.0, "D": 0.5, "L": 0.0}
        season_scores = [score_map[r] for r in results if r in score_map]
        season_avg = sum(season_scores) / len(season_scores) if season_scores else 0

        # WW Momentum
        emp_ww_score, num_ww = compute_empirical_score(results, k=2, target="W", score_map=score_map)
        p_w = results.count("W") / n if n > 0 else 0.5
        if emp_ww_score is not None:
            corrected_ww_prob, estimated_ww_bias = correct_for_bias(emp_ww_score, p_w, n, k=2, target="W", score_map=score_map)
            ww_bias = round(p_w - estimated_ww_bias, 3)
        else:
            corrected_ww_prob = None
            ww_bias = None

        score_map = {"L": 1.0, "D": 0.5, "W": 0.0}
        season_scores = [score_map[r] for r in results if r in score_map]
        season_lose_avg = sum(season_scores) / len(season_scores) if season_scores else 0
        # LL Momentum
        emp_ll_score, num_ll = compute_empirical_score(results, k=2, target="L",score_map=score_map)
        p_l = results.count("L") / n if n > 0 else 0.5
        if emp_ll_score is not None:
            corrected_ll_prob, estimated_ll_bias = correct_for_bias(emp_ll_score, p_l, n, k=2, target="L",score_map=score_map)
            ll_bias = round(p_l - estimated_ll_bias, 3)
        else:
            corrected_ll_prob = None
            ll_bias = None

        # Append to summary
        results_summary.append({
            "Team": sheet_name,
            "Season Avg %": round(season_avg * 100, 1),
            "WW Momentum Win % (corrected)": round(corrected_ww_prob * 100, 1) if corrected_ww_prob is not None else None,
            "WW Diff": round((corrected_ww_prob - season_avg) * 100, 1) if corrected_ww_prob is not None else None,
            "WW Bias": ww_bias,
            "Season Lose Avg %": round(season_lose_avg * 100, 1),
            "LL Momentum Loss % (corrected)": round(corrected_ll_prob * 100, 1) if corrected_ll_prob is not None else None,
            "LL Diff": round((corrected_ll_prob - season_lose_avg) * 100, 1) if corrected_ll_prob is not None else None,
            "LL Bias": ll_bias
        })
    df_summary = pd.DataFrame(results_summary)
    print("\n📊 Corrected Momentum vs. Season Success Rate (with Bias):\n")
    print(df_summary.to_string(index=False))
    # Calculate and print overall averages
    avg_season_win = df_summary["Season Avg %"].mean()
    avg_ww_momentum = df_summary["WW Momentum Win % (corrected)"].mean()
    avg_season_lose = df_summary["Season Lose Avg %"].mean()
    avg_ll_momentum = df_summary["LL Momentum Loss % (corrected)"].mean()

    print("\n📈 Overall Averages:")
    print(f"Average Season Win %: {avg_season_win:.1f}")
    print(f"Average WW Momentum Win % (corrected): {avg_ww_momentum:.1f}")
    print(f"Average Season Lose %: {avg_season_lose:.1f}")
    print(f"Average LL Momentum Loss % (corrected): {avg_ll_momentum:.1f}")
# Example usage
file_path = "PLteamsData20:21_with_tight_schedule.xlsx"
analyze_success_vs_momentum(file_path)
