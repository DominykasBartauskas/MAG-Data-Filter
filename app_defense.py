import pandas as pd
import re
from collections import Counter

# Load the Excel file
df = pd.read_excel('data_defense.xlsx')

# Print column names to verify the correct column names
print("Column names in the DataFrame:", df.columns)

# Regex pattern for finding the word "kills:" and the next number
kills_pattern = re.compile(r'Kills:\s*(\d+)')

# Initialize a counter for usernames
username_counter = Counter()

# List to store the data that doesn't fit the filter
lost_rows = []

# Process each row
for index, row in df.iterrows():
    # Check if the row has usernames in Column E and content in Column D
    if isinstance(row['Mentions'], str) and isinstance(row['Content'], str):
        # Extract usernames
        usernames = re.findall(r'[\w\.#]+#\d+', row['Mentions'])
        
        # Extract the number of kills from the 'Content' column
        kills_match = kills_pattern.search(row['Content'])
        if kills_match:
            kills = int(kills_match.group(1))
            
            # Add kills count to each username
            for username in usernames:
                username_counter[username] += kills
        else:
            # If no kills match is found, copy the row to lost_rows
            lost_rows.append(row)
    else:
        # If 'Mentions' or 'Content' columns are not strings, copy the row to lost_rows
        lost_rows.append(row)

# Create a DataFrame from the username counter
usernames_df = pd.DataFrame(username_counter.items(), columns=['Username', 'Kills'])

# Create a DataFrame for lost rows
lost_df = pd.DataFrame(lost_rows)

# Save the results to a new Excel file
with pd.ExcelWriter('processed_defense_data.xlsx') as writer:
    usernames_df.to_excel(writer, sheet_name='Usernames Kills', index=False)
    lost_df.to_excel(writer, sheet_name='Lost Rows', index=False)

print("Completed.")
