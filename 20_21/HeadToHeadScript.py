import pandas as pd

def calculate_momentum(last_three_results):
    score = 0
    for result in last_three_results:
        if result == "W":
            score += 3
        elif result == "D":
            score += 1
    return score


def custom_date_comparator(date1, date2):
    if not isinstance(date1, str):
        date1 = str(date1)
    if not isinstance(date2, str):
        date2 = str(date2)
    day1, month1 = map(int, date1.split("."))
    day2, month2 = map(int, date2.split("."))
    # Adjust month order: months 8-12 should be before months 1-7
    adjusted_month1 = month1 if month1 >= 8 else month1 + 12
    adjusted_month2 = month2 if month2 >= 8 else month2 + 12

    # Compare by adjusted month, then by day
    if adjusted_month1 < adjusted_month2:
        return True
    elif adjusted_month1 > adjusted_month2:
        return False
    else:
        return day1 < day2
def analyze_match_results(file_path):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    team_histories = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)  # Force 'Date' as str
        if "Date" in df.columns and "Opponent" in df.columns and "W/D/L" in df.columns:
            df = df.dropna(subset=["Opponent", "W/D/L"])
            df["Date"] = df["Date"].astype(str).apply(lambda x: f"{int(x.split('.')[0]):02}.{int(x.split('.')[1]):02}" if '.' in x else x)
            team_histories[sheet_name] = df

    # Dictionary to store match analysis
    match_analysis = []
    # Iterate over each sheet (team) in the Excel file
    for sheet_name, df in team_histories.items():
        team_history = []

        # Iterate through each row to analyze the first and second meetings
        for index, row in df.iterrows():
            opponent = row["Opponent"]
            result = row["W/D/L"]

            # Get last 3 matches of the team
            if len(team_history) >= 3:
                last_three_team = team_history[-3:]
                team_momentum = calculate_momentum(last_three_team)
            else:
                team_momentum = 0

            # Store match result for team
            team_history.append(result)

            # Get opponent's last 3 matches using `i`
            if opponent in team_histories:
                opponent_df = team_histories[opponent]
                opponent_df["Date"] = opponent_df["Date"].astype(str)  # Ensure opponent date is str
                opponent_past_games = opponent_df.loc[opponent_df["Date"].apply(lambda x: custom_date_comparator(x, row["Date"])), "W/D/L"].tolist()
                if len(opponent_past_games) >= 3:
                    last_three_opponent = opponent_past_games[-3:]
                    opponent_momentum = calculate_momentum(last_three_opponent)
                else:
                    opponent_momentum = 0
            else:
                opponent_momentum = 0

            # Determine winner of this match
            winner = None
            if result == "W":
                winner = sheet_name
            elif result == "L":
                winner = opponent
            else:
                winner = "Draw"

            # Store analysis
            match_analysis.append({
                "Team": sheet_name,
                "Opponent": opponent,
                "Match Number": index+1,
                "Team Momentum Before Match": team_momentum,
                "Opponent Momentum Before Match": opponent_momentum,
                "Winner": winner
            })

    # Convert to DataFrame and save as an Excel file
    if match_analysis:
        analysis_df = pd.DataFrame(match_analysis)
        output_file_path = "Head_to_Head_analysis.xlsx"
        analysis_df.to_excel(output_file_path, index=False)
        return output_file_path
    else:
        return None

# Example usage:
file_path = "PLteamsData20:21.xlsx"
analyze_match_results(file_path)

