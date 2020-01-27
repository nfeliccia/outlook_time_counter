import pandas as pd


def load_calendar(in_cal_name):
    """ Takes a file name of a .csv input and lods it
    :param in_cal_name
    return dataframe
    """
    cal_df_lc = pd.read_csv(in_cal_name, encoding='cp1252')
    cal_df_lc.drop(cal_df_lc[cal_df_lc['All day event'] == True].index, inplace=True)
    cal_df_lc.drop(labels=['All day event'], axis='columns', inplace=True)
    return cal_df_lc.copy()


def calculate_durations(in_cal_df):
    """
    :param dataframe representing calender
    :returns dataframe with the dates converted to date time objects.
    """
    whack_columns_cd = ['Start Date', 'Start Time', 'End Date', 'End Time', 'Reminder Date', 'Reminder Time']
    cal_df_cd = in_cal_df.copy()
    cal_df_cd['start_dt_object'] = pd.to_datetime(cal_df_cd['Start Date'] + ' ' + cal_df_cd['Start Time'],
                                                  format=outlook_date_format)
    cal_df_cd['end_dt_object'] = pd.to_datetime(cal_df_cd['End Date'] + ' ' + cal_df_cd['End Time'],
                                                format=outlook_date_format)
    cal_df_cd['event_duration'] = ((cal_df_cd['end_dt_object'] - cal_df_cd['start_dt_object']).astype(
        'timedelta64[m]')) / 60
    cal_df_cd.drop(labels=whack_columns_cd, axis='columns', inplace=True)
    return cal_df_cd


#  Constant declarations
calendar_file_name = 'test_cal.csv'

whack_columns = ['Meeting Organizer', 'Required Attendees', 'Optional Attendees', 'Meeting Resources',
                 'Billing Information', 'Sensitivity', 'Private', 'Show time as', 'Location', 'Mileage', 'Priority',
                 'Reminder on/off']
weekDays = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
outlook_date_format = '%m/%d/%Y %H:%M:%S'
full_list_sheet_name = 'All_Data'
total_sum_time_sheet_name = 'Total_Sum_Time'
year_month_sheet_name = 'Year_Month'
year_week_sheet_name = 'Year_Week'
year_month_pivot_sheet_name = 'Year_Month_Pivot'

# functions
# load the calendar
cal_df = load_calendar(calendar_file_name).copy()
cal_df = calculate_durations(cal_df).copy()
# Retrieve the day of week and convert it to a day
cal_df['weekday'] = cal_df['start_dt_object'].dt.day_name()
cal_df.sort_values('start_dt_object', ascending=True, inplace=True)
cal_df.drop(labels=whack_columns, axis='columns', inplace=True)
# Add year and month grouping
cal_df['year'] = cal_df['start_dt_object'].dt.year
cal_df['month'] = cal_df['start_dt_object'].dt.month
cal_df['week'] = cal_df['start_dt_object'].dt.week

# Super high level summary
cal_summary_high = cal_df.groupby(['Categories'])[['event_duration']].agg('sum')
cal_summary_high['event_duration'] = round(cal_summary_high['event_duration'], 2)
cal_summary_high.sort_values('event_duration', ascending=False, inplace=True)

# Year and month sommary
cal_summary_ym = pd.crosstab([cal_df.year, cal_df.month], cal_df.Categories, values=cal_df.event_duration,
                             aggfunc='sum').round(2)
cal_summary_ym.sort_values(['year', 'month'], ascending=[True, True], inplace=True)

# Year and week sommary
cal_summary_yw = pd.crosstab([cal_df.year, cal_df.month, cal_df.week], cal_df.Categories, values=cal_df.event_duration,
                             aggfunc='sum').round(2)
cal_summary_yw.sort_values(['year', 'month', 'week'], ascending=[True, True, True], inplace=True)

# Make the file name
# Find biggest date to use as a file name.
most_recent_date = str(cal_df['start_dt_object'].max())[0:10]
excel_file_name = 'calendar_analysis_' + most_recent_date + '.xlsx'

with pd.ExcelWriter(excel_file_name) as writing_jawn:
    cal_df.to_excel(writing_jawn, sheet_name=full_list_sheet_name, index=False)
    all_data_sheet_object = writing_jawn.sheets[full_list_sheet_name]
    all_data_sheet_object.set_column(0,0, 18)
    all_data_sheet_object.set_column(3,3, 18)
    all_data_sheet_object.set_column(4,4, 18)
    all_data_sheet_object.set_column(6,6, 10)
    all_data_sheet_object.set_column(7,7, 5)
    all_data_sheet_object.set_column(8,8, 6)
    all_data_sheet_object.set_column(9,9, 5)
    cal_summary_high.to_excel(writing_jawn, sheet_name=total_sum_time_sheet_name)
    cal_summary_ym.to_excel(writing_jawn, sheet_name=year_month_sheet_name)
    cal_summary_yw.to_excel(writing_jawn, sheet_name=year_week_sheet_name)
    writing_jawn.save()
