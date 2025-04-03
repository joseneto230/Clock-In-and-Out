import pandas as pd
import os
from datetime import datetime, timedelta

# Define the Excel file name based on the year
current_year = datetime.now().year  # cria uma variavel com o valor do ano atual
curren_month = datetime.now().strftime("%B")# cria uma variavel com o valor do mês atual
file_name = f"{current_year}_SSA.csv"  #cria uma variavel com o nome e o tipo do arquivo a ser criado nesse caso excel com o ano e o nome da empresa

# Check if file exists, if not create it
def initialize_file():
    if not os.path.exists(file_name):   #faz uma busca no caminho do arquivo e procura a existência do mesmo
        df = pd.DataFrame(columns=[
            "Date", "Clock-in", "Interval Start", "Interval End", "Clock-out", "Status", 
            "Work Hours Needed", "Total Worked Hours", "Hours Bank"
        ])  # Formatação da primeira linha sendo cada objeto uma coluna
        df.to_csv(file_name, index=False) # cria o arquivo 

def load_excel():
    df = pd.read_csv(file_name)
    df["Total Worked Hours"] = df["Total Worked Hours"].astype(object)
    df["Hours Bank"] = df["Hours Bank"].astype(object)
    df["Clock-in"] = df["Clock-in"].astype(object)
    df["Interval Start"] = df["Interval Start"].astype(object)
    df["Interval End"] = df["Interval End"].astype(object)
    df["Clock-out"] = df["Clock-out"].astype(object)
    return df

def save_excel(df):
    df.to_csv(file_name, index=False)

def parse_hours(hours_str):
    try:
        time_obj = datetime.strptime(hours_str, '%H:%M:%S').time()
        return timedelta(hours=time_obj.hour, minutes=time_obj.minute, seconds=time_obj.second)
    except ValueError:
        try:
            hours = float(hours_str.split()[0])
            return timedelta(hours=hours)
        except (ValueError, IndexError):
            return timedelta(0)

def parse_hours_sum(hours_str):
    negative = hours_str.startswith("-")
    if negative:
        hours_str = hours_str[1:]
    parsed_time = parse_hours(hours_str)
    return -parsed_time if negative else parsed_time

def negative_hours(hours):
    return "-" + str(timedelta(seconds=abs(hours.total_seconds()))) if hours.total_seconds() < 0 else str(hours)

def sum_bank_hours():
    df = load_excel()
    total_bank = timedelta(0)
    
    for idx, value in df["Hours Bank"].items():
        if isinstance(value, str) and value.strip():
            try:
                bank_time = parse_hours_sum(value)
                total_bank += bank_time
            except Exception as e:
                print(f"Error parsing bank hour at index {idx}: {value} -> {e}")
    
    total_bank = negative_hours(total_bank)
    total_index = df[df["Date"] == "TOTAL"].index
    if not total_index.empty:
        df.loc[total_index[0], "Hours Bank"] = total_bank
    else:
        df = pd.concat([df, pd.DataFrame([{
            "Date": "TOTAL", "Clock-in": "", "Interval Start": "", "Interval End": "", 
            "Clock-out": "", "Status": "Summary", "Work Hours Needed": "", 
            "Total Worked Hours": "", "Hours Bank": total_bank
        }])], ignore_index=True)
    
    save_excel(df)
    print(f"Total Bank Hours: {total_bank}")

def bank(work_duration, idx, df):
    hours_need = parse_hours(df.at[idx, "Work Hours Needed"])
    bank_hours = negative_hours(work_duration - hours_need)
    df.at[idx, "Hours Bank"] = bank_hours

def register_time():
    df = load_excel()
    today = datetime.now().strftime("%d-%m")
    current_time = datetime.now().strftime("%H:%M:%S")
    total_index = df[df["Date"] == "TOTAL"].index
    entry_index = df[df["Date"] == today].index

    if entry_index.empty:
        new_data = pd.DataFrame([{
            "Date": today, "Clock-in": current_time, "Interval Start": pd.NA, "Interval End": pd.NA, 
            "Clock-out": pd.NA, "Status": "Working", "Work Hours Needed": "08:00:00", 
            "Total Worked Hours": pd.NA, "Hours Bank": "0:00:00"
        }])
        if not total_index.empty:
            total_position = total_index[0]
            df = pd.concat([df.iloc[:total_position], new_data, df.iloc[total_position:]], ignore_index=True)
        else:
            df = pd.concat([df, new_data], ignore_index=True)
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
            fmt = "%H:%M:%S"
            start_time = datetime.strptime(df.at[idx, "Clock-in"], fmt)
            end_time = datetime.strptime(df.at[idx, "Clock-out"], fmt)
            interval_start = datetime.strptime(df.at[idx, "Interval Start"], fmt) if pd.notna(df.at[idx, "Interval Start"]) else start_time
            interval_end = datetime.strptime(df.at[idx, "Interval End"], fmt) if pd.notna(df.at[idx, "Interval End"]) else start_time
            work_duration = (end_time - start_time) - (interval_end - interval_start)
            df.at[idx, "Total Worked Hours"] = str(work_duration)
            bank(work_duration, idx, df)
    
    save_excel(df)
    print("Time registered successfully!")

if __name__ == "__main__":
    initialize_file()
    register_time()
    sum_bank_hours()
