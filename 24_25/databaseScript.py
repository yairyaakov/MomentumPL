import pandas as pd

def process_team_data(writer, file_path, columns, sheet_name):

    data = pd.read_excel(file_path, sheet_name="Fixtures_by_Clubs")

    # Extract relevant columns based on input
    relevant_data = data[columns]

    # Rename columns for clarity (custom names)
    relevant_data.columns = ['Date', 'Opponent', 'Home/Away', 'Result', 'Column7', 'Column8', 'W/D/L']
    filtered_data = relevant_data.copy()  # Create a copy to ensure we don't overwrite original data

    # Map months to numbers
    month_map = {
        "January": 1,
        "February": 2,
        "March": 3,
        "April": 4,
        "May": 5,
        "June": 6,
        "July": 7,
        "August": 8,
        "September": 9,
        "October": 10,
        "November": 11,
        "December": 12
    }

    # Convert 'Column8' to int and then to string
    filtered_data['Column8'] = pd.to_numeric(filtered_data['Column8'], errors='coerce').fillna(0).astype(int).astype(str)

    # Merge columns 'Result', 'Column7', 'Column8' into a single column
    filtered_data['Result'] = (
        filtered_data[['Result', 'Column7', 'Column8']]
        .fillna('')  # Replace NaN with an empty string
        .astype(str)  # Convert all values to strings
        .agg(' '.join, axis=1)  # Join the values with a space
        .str.strip()  # Remove leading/trailing spaces
    )

    # Drop the merged columns (Column7, Column8)
    filtered_data = filtered_data.drop(columns=['Column7', 'Column8'])

    # Update the Date column to the format day.month
    current_month = None
    formatted_dates = []

    # Iterate through the 'Date' column and format the date accordingly
    for index, date in enumerate(filtered_data['Date']):
        if isinstance(date, str) and date in month_map:  # Check if the row contains a month name
            current_month = month_map[date]
        elif isinstance(date, (int)):  # Check if it's a valid day number
            if current_month is not None:
                formatted_dates.append(f"{int(date)}.{current_month}")
            else:
                formatted_dates.append(date)

    # Ensure no missing data in other columns (optional)
    filtered_data['Opponent'] = filtered_data['Opponent'].fillna(method='bfill')
    filtered_data['Home/Away'] = filtered_data['Home/Away'].fillna(method='bfill')
    filtered_data['Result'] = filtered_data['Result'].fillna(method='bfill')
    filtered_data['W/D/L'] = filtered_data['W/D/L'].fillna(method='bfill')

    # Remove duplicates in the 'Opponent' column, keeping only the first occurrence
    filtered_data['Opponent'] = filtered_data['Opponent'].where(
        filtered_data['Opponent'] != filtered_data['Opponent'].shift(-1)
    )

    # Drop rows where 'Opponent' is NaN after the removal of duplicates
    filtered_data = filtered_data.dropna(subset=['Opponent'])

    # Ensure the length of formatted_dates matches the length of the DataFrame
    # If formatted_dates is shorter, pad with NaN or original values
    while len(formatted_dates) < len(filtered_data):
        formatted_dates.append(None)  # This will add None for missing dates, you can use NaN if needed

    # Assign the new Date column
    filtered_data['Date'] = formatted_dates

    # Remove rows below the 39th row (index 38)
    filtered_data = filtered_data.iloc[:38]

    # Export to Excel with each team's data in a separate sheet
    filtered_data.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"Data exported successfully for {sheet_name}.")


def main():
    file_path = "PL24.25.xlsx"

    # Load teams from the Dashboard sheet
    df_dashboard = pd.read_excel(file_path, sheet_name="Dashboard")
    teams = df_dashboard.iloc[:, 1].dropna().tolist()  # Assuming teams are in the second column

    start_index = 0
    step = 12

    with pd.ExcelWriter("PLteamsData24:25.xlsx", engine="xlsxwriter") as writer:
        for i, team in enumerate(teams):
            # Check if "Unnamed: 0" exists; otherwise, use `team`
            df_fixtures = pd.read_excel(file_path, sheet_name="Fixtures_by_Clubs", nrows=1)
            first_column = f"Unnamed: {start_index}" if f"Unnamed: {start_index}" in df_fixtures.columns else team

            columns = [first_column] + [f"Unnamed: {start_index + j}" for j in [4, 5, 6, 7, 8, 9]]

            process_team_data(writer, file_path, columns, team)
            start_index += step  # Move to the next team's set of columns


if __name__ == "__main__":
    main()