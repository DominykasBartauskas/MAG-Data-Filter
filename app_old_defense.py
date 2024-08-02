import pandas as pd
import re
from collections import Counter

# Load the Excel file
df = pd.read_excel('data_defense.xlsx')

# Print column names to verify the correct column names
print("Column names in the DataFrame:", df.columns)

# Regex pattern for result format
result_pattern = r'(\d+)\s*[vV][sS]?\s*(\d+)'

# Initialize a counter for usernames
username_counter = Counter()

# List to store the data that doesn't fit the filter
lost_rows = []

# Process each row
for index, row in df.iterrows():
    # Check if the row has an image link in Column F and usernames in Column E
    if isinstance(row['link'], str) and isinstance(row['Mentions'], str):
        # Extract usernames
        usernames = re.findall(r'[\w\.#]+#\d+', row['Mentions'])
        
        # Ensure the 'Content' column is a string before processing
        if isinstance(row['Content'], str):
            result_text = row['Content']
            result_numbers = re.findall(result_pattern, result_text)
            if result_numbers:
                defenders, attackers = map(int, result_numbers[0])
            else:
                defenders, attackers = 0, 0
            
            # Check the result text for 'win' or 'lost'
            result = result_text.lower()
            if 'win' in result:
                # Add attackers count to each username
                for username in usernames:
                    username_counter[username] += attackers
            elif 'lost' in result or 'lose' in result or not result:
                # Copy the row to lost_rows
                lost_rows.append(row)
            else:
                # Copy the row to lost_rows
                lost_rows.append(row)
        else:
            # Copy the row to lost_rows
            lost_rows.append(row)
    else:
        # Copy the row to lost_rows
        lost_rows.append(row)

# Create a DataFrame from the username counter
usernames_df = pd.DataFrame(username_counter.items(), columns=['Username', 'Count'])

# Create a DataFrame for lost rows
lost_df = pd.DataFrame(lost_rows)

# Save the results to a new Excel file
with pd.ExcelWriter('processed_defense_data.xlsx') as writer:
    usernames_df.to_excel(writer, sheet_name='Usernames Count', index=False)
    lost_df.to_excel(writer, sheet_name='Lost Rows', index=False)

print("Completed.")
