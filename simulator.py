import pandas as pd
import random

from datetime import datetime, timedelta

# Define the Excel file name
current_year = datetime.now().year
file_name = f"{current_year}_SSA.xlsx"

# Function to parse hours string to timedelta
def parse_hours(hours_str):
    h, m, s = map(int, hours_str.split(":"))
    return timedelta(hours=h, minutes=m, seconds=s)

def negative_hours(hours):
    if hours.total_seconds() < 0:
        total_seconds = abs(hours.total_seconds())  # Get absolute value
        negative_time = str(timedelta(seconds=total_seconds))  # Convert to HH:MM:SS
        negative_time = "-" + negative_time  # Manually add the negative sign
        hours = negative_time
    else:
        hours = str(hours)
    return hours

# Generate 10 days of test data
test_data = []
start_date = datetime.now()
total_bank_hours = timedelta(0)  # To track total bank hours

for i in range(10):
    work_date = (start_date - timedelta(days=i+1)).strftime("%d-%m")

    # Generate Clock-in (between 07:00 - 10:00)
    clock_in = datetime.strptime(f"{random.randint(7, 10)}:{random.randint(0, 59)}:00", "%H:%M:%S")

    # Generate Interval Start (2-5 hours after clock-in)
    interval_start = clock_in + timedelta(hours=random.randint(2, 5), minutes=random.randint(0, 30))

    # Generate Interval End (~1 hour after Interval Start, ±5 min)
    interval_end = interval_start + timedelta(minutes=60 + random.randint(-5, 5))

    # Generate Clock-out (after Interval End, between 16:00 - 20:00)
    clock_out = interval_end + timedelta(hours=random.randint(4, 6), minutes=random.randint(0, 30))

    # Convert times to strings
    clock_in_str = clock_in.strftime("%H:%M:%S")
    interval_start_str = interval_start.strftime("%H:%M:%S")
    interval_end_str = interval_end.strftime("%H:%M:%S")
    clock_out_str = clock_out.strftime("%H:%M:%S")

    # Calculate total worked hours: (Clock-out - Clock-in) - (Interval End - Interval Start)
    work_duration = (clock_out - clock_in) - (interval_end - interval_start)
    total_worked_hours_str = str(work_duration)

    # Define required work hours (8 hours)
    work_hours_needed = parse_hours("08:00:00")

    # Calculate bank hours: Total Worked Hours - Work Hours Needed
    bank_hours = work_duration - work_hours_needed
    bank_hours = negative_hours(bank_hours)
    # total_bank_hours += bank_hours  # Accumulate for final sum


    # Convert bank hours to string
    bank_hours_str = str(bank_hours)

    # Add row to test data
    test_data.append({
        "Date": work_date,
        "Clock-in": clock_in_str,
        "Interval Start": interval_start_str,
        "Interval End": interval_end_str,
        "Clock-out": clock_out_str,
        "Status": "Out of Work",
        "Work Hours Needed": "08:00:00",
        "Total Worked Hours": total_worked_hours_str,
        "Hours Bank": bank_hours_str
    })
# Add a final row with the total sum of Hours Bank

test_data.append({
    "Date": "TOTAL",
    "Clock-in": "",
    "Interval Start": "",
    "Interval End": "",
    "Clock-out": "",
    "Status": "Summary",
    "Work Hours Needed": "",
    "Total Worked Hours": "",
    "Hours Bank": str(total_bank_hours)
})

# Convert to DataFrame and save to Excel

df_test = pd.DataFrame(test_data)
df_test.to_excel(file_name, index=False)

print(f"Test data for 10 days saved in {file_name}, with Hours Bank total: {total_bank_hours}")