import pandas as pd
import re
from collections import Counter

def process_file(file_path):
    df = pd.read_excel(file_path)
    print("Column names in the DataFrame:", df.columns)

    username_counter = Counter()
    lost_rows = []

    for index, row in df.iterrows():
        if isinstance(row['Username'], str) and isinstance(row['Content'], str):
            username = row['Username']
            if isinstance(row['Content'], str):
                data_text = row['Content']
                if re.search(r'delivered', data_text, re.IGNORECASE):
                    username_counter[username] += 2
                else:
                    lost_rows.append(row)
            else:
                lost_rows.append(row)
        else:
            lost_rows.append(row)

    usernames_df = pd.DataFrame(username_counter.items(), columns=['Username', 'Count'])
    lost_df = pd.DataFrame(lost_rows)

    output_path = file_path.replace('.xlsx', '_processed_delivery.xlsx')
    with pd.ExcelWriter(output_path) as writer:
        usernames_df.to_excel(writer, sheet_name='Usernames Count', index=False)
        lost_df.to_excel(writer, sheet_name='Lost Rows', index=False)

    print("Delivery processing completed.")
    return output_path
