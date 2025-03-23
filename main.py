import pandas as pd
import os
from datetime import datetime

# Define the Excel file name based on the year
current_year = datetime.now().year
file_name = f"{current_year}_SSA.xlsx"

# Check if file exists, if not create it
if not os.path.exists(file_name):
    df = pd.DataFrame(columns=["Date", "Clock-in", "Interval Start", "Interval End", "Clock-out", "Status", "Work Hours Needed", "Total Worked Hours", "Hours Bank"])
    df.to_excel(file_name, index=False)

# Load the existing Excel file
df = pd.read_excel(file_name)

def register_time():
    global df  # Ensure df is accessible
    df = pd.read_excel(file_name)  # Load the existing Excel file
    today = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")
    
    # Find today's entry
    entry_index = df[df["Date"] == today].index
    
    if entry_index.empty:
        # New entry for today
        new_data = {
            "Date": today,
            "Clock-in": current_time,
            "Interval Start": "",
            "Interval End": "",
            "Clock-out": "",
            "Status": "Working",
            "Work Hours Needed": 8,
            "Total Worked Hours": "",
            "Hours Bank": 0
        }
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
    else:
        idx = entry_index[0]
        if pd.isna(df.at[idx, "Clock-in"]):
            df.at[idx, "Clock-in"] = current_time
            df.at[idx, "Status"] = "Working"
        elif pd.isna(df.at[idx, "Interval Start"]):
            df.at[idx, "Interval Start"] = current_time
            df.at[idx, "Status"] = "On Break"
        elif pd.isna(df.at[idx, "Interval End"]):
            df.at[idx, "Interval End"] = current_time
            df.at[idx, "Status"] = "Working Again"
        elif pd.isna(df.at[idx, "Clock-out"]):
            df.at[idx, "Clock-out"] = current_time
            df.at[idx, "Status"] = "Out of Work"
            
            # Calculate total worked hours
            fmt = "%H:%M:%S"
            start_time = datetime.strptime(df.at[idx, "Clock-in"], fmt)
            end_time = datetime.strptime(df.at[idx, "Clock-out"], fmt)
            interval_start = datetime.strptime(df.at[idx, "Interval Start"], fmt) if pd.notna(df.at[idx, "Interval Start"]) else start_time
            interval_end = datetime.strptime(df.at[idx, "Interval End"], fmt) if pd.notna(df.at[idx, "Interval End"]) else start_time
            
            work_duration = (end_time - start_time) - (interval_end - interval_start)
            df.at[idx, "Total Worked Hours"] = str(work_duration)
        if pd.notna(df.at[idx, "Clock-in"]) and pd.notna(df.at[idx, "Interval Start"]) and not pd.notna(df.at[idx, "Interval End"]):

            fmt = "%H:%M:%S"
            start_time = datetime.strptime(df.at[idx, "Clock-in"], fmt)
            interval_start = datetime.strptime(df.at[idx, "Interval Start"], fmt) if pd.notna(df.at[idx, "Interval Start"]) else start_time
            work_duration = (interval_start - start_time)
            df.at[idx, "Total Worked Hours"] = str(work_duration)
            

    df.to_excel(file_name, index=False)
    print("Time registered successfully!")

if __name__ == "__main__":
    register_time()
