import pandas as pd
import numpy as np
import datetime as dt
from calendar import monthrange

# Notes: need to install from Settings/Interpreter packages
# Need  xlrd  and openpyxl
# --


def check_valid_month_fcast(start, end, mon, amt):
    days_in_month = monthrange(2020, mon)[1]
    range_start = dt.datetime(2020, mon, 1)
    range_end = dt.datetime(2020, mon, days_in_month)
    # print("start={}\tend={}\trs={}\tre={}".format(start, end, range_start, range_end))
    if start <= range_end and end >= range_start:
        # Person is supposed to be here this month
        if amt == 0.0:
            return "UNDER"
        else:
            # Could probably be smarter here
            return "Good"
    else:
        # Person is NOT supposed to be here this month
        if amt != 0.0:
            return "OVER"
    return "Good"


# Load the input spreadsheets
df_namelist = pd.read_excel(r'ss_namelist.xlsx')
df_planning = pd.read_excel(r'ss_planning.xlsx', dtype={'Feb':np.float64, 'May':np.float64})

# Both need to share the "Owner" column, put in ZEROs where NaN/NaT
df_merged = pd.merge(df_namelist.fillna(0), df_planning.fillna(0), how='outer', on='Owner').fillna(0)

df_merged.to_excel('ss_tocheck.xlsx', sheet_name='merged', index=False)

for row in df_merged.iterrows():
    name = row[1][0]
    start_date = row[1][1]
    end_date = row[1][2]
    if start_date == 0 or end_date == 0:
        print("Name: {} does not exist in namelistinput, or exists but has invalid start/end date(s)".format(name))
        continue    # jump to next iterable row
    print(name, "\t{} \t{}".format(start_date.strftime("%c"), end_date.strftime("%c")), end='')
    col_offset = 2
    for month in range(12):
        # print("\t {}".format(month+1), end='')
        fcast = row[1][month+col_offset+1]
        fcast_status = check_valid_month_fcast(start_date, end_date, month+1, fcast)
        print("\t{}".format(fcast_status), end='')
    print("")
