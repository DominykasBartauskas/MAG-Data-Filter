import pandas as pd
import re
from collections import Counter

def process_file(file_path):
    df = pd.read_excel(file_path)
    print("Column names in the DataFrame:", df.columns)

    kills_pattern = re.compile(r'Kills:\s*(\d+)')
    username_counter = Counter()
    lost_rows = []

    for index, row in df.iterrows():
        if isinstance(row['Mentions'], str) and isinstance(row['Content'], str):
            usernames = re.findall(r'[\w\.#]+#\d+', row['Mentions'])
            kills_match = kills_pattern.search(row['Content'])
            if kills_match:
                kills = int(kills_match.group(1))
                for username in usernames:
                    username_counter[username] += kills
            else:
                lost_rows.append(row)
        else:
            lost_rows.append(row)

    usernames_df = pd.DataFrame(username_counter.items(), columns=['Username', 'Kills'])
    lost_df = pd.DataFrame(lost_rows)

    output_path = file_path.replace('.xlsx', '_processed_defense.xlsx')
    with pd.ExcelWriter(output_path) as writer:
        usernames_df.to_excel(writer, sheet_name='Usernames Kills', index=False)
        lost_df.to_excel(writer, sheet_name='Lost Rows', index=False)

    print("Defense processing completed.")
