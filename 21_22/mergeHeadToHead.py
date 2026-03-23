import pandas as pd

# Load the head-to-head analysis file
file_path = "Head_to_Head_analysis.xlsx"
df = pd.read_excel(file_path)

# Ensure data is sorted properly
df = df.sort_values(by=["Team", "Opponent", "Match Number"])

# Dictionary to track processed matches
processed_pairs = set()
merged_matches = []

# Iterate through each row to find and merge home/away matches
for index, row in df.iterrows():
    team, opponent = row["Team"], row["Opponent"]

    # Skip if this pair has already been processed
    if (team, opponent) in processed_pairs or (opponent, team) in processed_pairs:
        continue

    # Filter only relevant matches for the current team
    team_matches = df[(df["Team"] == team) | (df["Team"] == opponent)]

    # Find both matches between these teams
    matches = team_matches[((team_matches["Team"] == team) & (team_matches["Opponent"] == opponent)) |
                           ((team_matches["Team"] == opponent) & (team_matches["Opponent"] == team))]


    # Ensure there are exactly two matches
    if len(matches) == 4:
        first_leg = matches.iloc[0]
        second_leg = matches.iloc[1]

        # Create a merged row
        merged_row = {
            "Team": team,
            "Opponent": opponent,
            "First Leg Match Number": first_leg["Match Number"],
            "First Leg Momentum (Team)": first_leg["Team Momentum Before Match"],
            "First Leg Momentum (Opponent)": first_leg["Opponent Momentum Before Match"],
            "First Leg Winner": first_leg["Winner"],
            "Second Leg Match Number": second_leg["Match Number"],
            "Second Leg Momentum (Team)": second_leg["Team Momentum Before Match"],
            "Second Leg Momentum (Opponent)": second_leg["Opponent Momentum Before Match"],
            "Second Leg Winner": second_leg["Winner"]
        }

        merged_matches.append(merged_row)

        # Mark this pair as processed
        processed_pairs.add((team, opponent))
        processed_pairs.add((opponent, team))

# Convert merged data to DataFrame
merged_df = pd.DataFrame(merged_matches)

# Save the cleaned file
output_file_path = "Merged_Head_to_Head.xlsx"
merged_df.to_excel(output_file_path, index=False)

