import pandas as pd
import re
from collections import Counter

# Load the Excel file
df = pd.read_excel('data_delivery.xlsx')

# Print column names to verify the correct column names
print("Column names in the DataFrame:", df.columns)

# Initialize a counter for usernames
username_counter = Counter()

# List to store the data that doesn't fit the filter
lost_rows = []

# Process each row
for index, row in df.iterrows():
    # Check if the row has a username in Column B and data in Column D
    if isinstance(row['Username'], str) and isinstance(row['Content'], str):
        # Extract username
        username = row['Username']
        
        # Ensure the 'Data' column is a string before processing
        if isinstance(row['Content'], str):
            data_text = row['Content']
            # Check if the data text contains the phrase 'delive' (case insensitive)
            if re.search(r'delive', data_text, re.IGNORECASE):
                # Increment the count for the username
                username_counter[username] += 1
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
with pd.ExcelWriter('processed_delivery_data.xlsx') as writer:
    usernames_df.to_excel(writer, sheet_name='Usernames Count', index=False)
    lost_df.to_excel(writer, sheet_name='Lost Rows', index=False)

print("Completed.")
