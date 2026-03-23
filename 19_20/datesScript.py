import pandas as pd

def process_team_schedule(sheet_df, default_year=2023, max_gap=4):
    sheet_df = sheet_df.copy()
    sheet_df['Date'] = sheet_df['Date'].astype(str)
    sheet_df['Parsed Date'] = pd.to_datetime(sheet_df['Date'] + f'.{default_year}', format='%d.%m.%Y', errors='coerce')
    sheet_df = sheet_df.dropna(subset=['Parsed Date'])
    sheet_df = sheet_df.sort_values(by='Parsed Date').reset_index(drop=True)
    sheet_df['Days Since Last Match'] = sheet_df['Parsed Date'].diff().dt.days
    sheet_df['Tight Schedule'] = sheet_df['Days Since Last Match'] <= max_gap
    return sheet_df

# === Main Execution ===
file_path = "PLteamsData19:20_updated.xlsx"
output_file_path = "PLteamsData19:20_with_tight_schedule.xlsx"
excel_data = pd.ExcelFile(file_path)

# Process and write all teams to a new file
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for team in excel_data.sheet_names:
        df = excel_data.parse(team)
        processed_df = process_team_schedule(df)
        processed_df.to_excel(writer, sheet_name=team, index=False)

print(f"✅ All teams saved to: {output_file_path}")
