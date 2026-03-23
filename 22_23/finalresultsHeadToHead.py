import pandas as pd

# Load the merged head-to-head data
file_path = "Merged_Head_to_Head.xlsx"
df = pd.read_excel(file_path)

# Create a list to store filtered games
filtered_games = []

# Iterate through each row to process both First Leg and Second Leg
for _, row in df.iterrows():
    first_leg_momentum_team = row["First Leg Momentum (Team)"]
    first_leg_momentum_opponent = row["First Leg Momentum (Opponent)"]
    second_leg_momentum_team = row["Second Leg Momentum (Team)"]
    second_leg_momentum_opponent = row["Second Leg Momentum (Opponent)"]

    # Check if the first leg has different momentum
    if first_leg_momentum_team != first_leg_momentum_opponent:
        filtered_games.append({
            "Team": row["Team"],
            "Opponent": row["Opponent"],
            "Leg": "First Leg",
            "Momentum Winner": row["First Leg Winner"],
            "Momentum (Team)": first_leg_momentum_team,
            "Momentum (Opponent)": first_leg_momentum_opponent
        })

    # Check if the second leg has different momentum
    if second_leg_momentum_team != second_leg_momentum_opponent:
        filtered_games.append({
            "Team": row["Team"],
            "Opponent": row["Opponent"],
            "Leg": "Second Leg",
            "Momentum Winner": row["Second Leg Winner"],
            "Momentum (Team)": second_leg_momentum_team,
            "Momentum (Opponent)": second_leg_momentum_opponent
        })

# Convert the filtered games into a DataFrame
df_filtered = pd.DataFrame(filtered_games)

# Total Matches Analyzed (where at least one game had a momentum difference)
total_matches_analyzed = len(df_filtered)

# Count how many times the team with the better momentum won
momentum_winner_victories = df_filtered[
    ((df_filtered["Momentum (Team)"] > df_filtered["Momentum (Opponent)"]) & (df_filtered["Momentum Winner"] == df_filtered["Team"])) |
    ((df_filtered["Momentum (Team)"] < df_filtered["Momentum (Opponent)"]) & (df_filtered["Momentum Winner"] == df_filtered["Opponent"]))
].shape[0]



# Count cases where momentum switched (First Leg to Second Leg)
momentum_switch_cases = df[
    ((df["First Leg Momentum (Team)"] > df["First Leg Momentum (Opponent)"]) &
     (df["Second Leg Momentum (Team)"] < df["Second Leg Momentum (Opponent)"])) |
    ((df["First Leg Momentum (Team)"] < df["First Leg Momentum (Opponent)"]) &
     (df["Second Leg Momentum (Team)"] > df["Second Leg Momentum (Opponent)"]))
].shape[0]

# Count cases where momentum switched AND the winner also changed,
# while ensuring the team with higher momentum won each match
momentum_switch_different_winner = df[
    (((df["First Leg Momentum (Team)"] > df["First Leg Momentum (Opponent)"]) &
      (df["Second Leg Momentum (Team)"] < df["Second Leg Momentum (Opponent)"]) &
      (df["First Leg Winner"] == df["Team"]) &
      (df["Second Leg Winner"] == df["Opponent"])) |

     ((df["First Leg Momentum (Team)"] < df["First Leg Momentum (Opponent)"]) &
      (df["Second Leg Momentum (Team)"] > df["Second Leg Momentum (Opponent)"]) &
      (df["First Leg Winner"] == df["Opponent"]) &
      (df["Second Leg Winner"] == df["Team"])))
].shape[0]

momentum_switch_negative_winner = df[
    (((df["First Leg Momentum (Team)"] > df["First Leg Momentum (Opponent)"]) &
      (df["Second Leg Momentum (Team)"] < df["Second Leg Momentum (Opponent)"]) &
      (df["First Leg Winner"] == df["Opponent"]) &
      (df["Second Leg Winner"] == df["Team"])) |

     ((df["First Leg Momentum (Team)"] < df["First Leg Momentum (Opponent)"]) &
      (df["Second Leg Momentum (Team)"] > df["Second Leg Momentum (Opponent)"]) &
      (df["First Leg Winner"] == df["Team"]) &
      (df["Second Leg Winner"] == df["Opponent"])))
].shape[0]

# Print results
print("Total Matches Analyzed:", total_matches_analyzed)
print("Momentum Winner Victories:", momentum_winner_victories)
print("Total Cases Where Momentum Switched:", momentum_switch_cases)
print("Momentum Switched & Different Winners:", momentum_switch_different_winner)
print("Momentum Switched & Negative Winner:", momentum_switch_negative_winner)
