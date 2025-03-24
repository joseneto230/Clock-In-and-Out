import pandas as pd
import os
from datetime import datetime, timedelta

# Define the Excel file name based on the year
current_year = datetime.now().year
file_name = f"{current_year}_SSA.xlsx"

# Check if file exists, if not create it
if not os.path.exists(file_name):
    df = pd.DataFrame(columns=["Date", "Clock-in", "Interval Start", "Interval End", "Clock-out", "Status", "Work Hours Needed", "Total Worked Hours", "Hours Bank"])
    df.to_excel(file_name, index=False)

# Load the existing Excel file
df = pd.read_excel(file_name)

# Ensure 'Total Worked Hours' and 'Hours Bank' are object dtype
df['Total Worked Hours'] = df['Total Worked Hours'].astype(object)
df['Hours Bank'] = df['Hours Bank'].astype(object) #added this line

def parse_hours(hours_str):
    """Parses 'HH:MM:SS' or 'X hours' string into timedelta, handling errors."""
    try:
        # Try parsing 'HH:MM:SS' format
        time_obj = datetime.strptime(hours_str, '%H:%M:%S').time()
        return timedelta(hours=time_obj.hour, minutes=time_obj.minute, seconds=time_obj.second)
    except ValueError:
        # If 'HH:MM:SS' parsing fails, try 'X hours' format
        try:
            hours = float(hours_str.split()[0])
            return timedelta(hours=hours)
        except (ValueError, IndexError):
            print(f"Error: Invalid hours string '{hours_str}'. Returning 0 hours.")
            return timedelta(0)

def bank(work_duration, idx):
    hours_need_str = df.at[idx, "Work Hours Needed"]
    hours_need = parse_hours(hours_need_str)

    print(f"Hours Needed: {hours_need}")
    print(f"Work Duration: {work_duration}")

    # Ensure work_duration and hours_need are timedelta objects
    if not isinstance(work_duration, timedelta):
        work_duration = parse_hours(work_duration)
    
    if not isinstance(hours_need, timedelta):
        hours_need = parse_hours(hours_need)

    # Calculate bank hours
    bank_hours = work_duration - hours_need  

    # Normalize negative timedelta format
    if bank_hours.total_seconds() < 0:
        total_seconds = abs(bank_hours.total_seconds())  # Get absolute value
        negative_time = str(timedelta(seconds=total_seconds))  # Convert to HH:MM:SS
        negative_time = "-" + negative_time  # Manually add the negative sign
        bank_hours = negative_time
    else:
        bank_hours = str(bank_hours)

    print(f"Bank Hours: {bank_hours}")
    df.at[idx, "Hours Bank"] = bank_hours  # Store properly formatted value

def register_time():
    global df  # Ensure df is accessible
    df = pd.read_excel(file_name)  # Load the existing Excel file
    # Ensure 'Total Worked Hours' and 'Hours Bank' are object dtype after reloading.
    df['Total Worked Hours'] = df['Total Worked Hours'].astype(object)
    df['Hours Bank'] = df['Hours Bank'].astype(object)
    df['Clock-in'] = df['Clock-in'].astype(object)
    df['Interval Start'] = df['Interval Start'].astype(object)
    df['Interval End'] = df['Interval End'].astype(object)
    df['Clock-out'] = df['Clock-out'].astype(object)

    today = datetime.now().strftime("%Y-%m-%d")
    current_time = datetime.now().strftime("%H:%M:%S")

    # Find today's entry
    entry_index = df[df["Date"] == today].index

    if entry_index.empty:
        # New entry for today
        new_data = {
            "Date": today,
            "Clock-in": current_time,
            "Interval Start": pd.NA,
            "Interval End": pd.NA,
            "Clock-out": pd.NA,
            "Status": "Working",
            "Work Hours Needed": "08:00:00",  # Store as string for parsing
            "Total Worked Hours": pd.NA,
            "Hours Bank": "0:00:00" #store as a timedelta string.
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
            bank(work_duration, idx)

        if pd.notna(df.at[idx, "Clock-in"]) and pd.notna(df.at[idx, "Interval Start"]) and pd.isna(df.at[idx, "Interval End"]):
            fmt = "%H:%M:%S"
            start_time = datetime.strptime(df.at[idx, "Clock-in"], fmt)
            interval_start = datetime.strptime(df.at[idx, "Interval Start"], fmt) if pd.notna(df.at[idx, "Interval Start"]) else start_time
            work_duration = (interval_start - start_time)
            df.at[idx, "Total Worked Hours"] = str(work_duration)
            bank(work_duration, idx)

    df.to_excel(file_name, index=False)
    print("Time registered successfully!")

if __name__ == "__main__":
    register_time()