import pandas as pd

calendar_file_name = 'test_cal.csv'
duration_file_name = calendar_file_name[:-4] + '_with_durations.csv'
summary_file_name = calendar_file_name[:-4] + '_summary.csv'
whack_columns = ['Start Date', 'Start Time', 'End Date', 'End Time', 'Reminder Date', 'Reminder Time',
                 'Meeting Organizer', 'Required Attendees', 'Optional Attendees', 'Meeting Resources',
                 'Billing Information', 'Sensitivity', 'Private', 'Show time as', 'Location', 'Mileage', 'Priority',
                 'Reminder on/off']
cal_df = pd.read_csv(calendar_file_name, encoding='cp1252')

cal_df['start_dt_raw'] = cal_df['Start Date'] + ' ' + cal_df['Start Time']
cal_df['start_dt_object'] = pd.to_datetime(cal_df['start_dt_raw'], format='%m/%d/%Y %H:%M:%S')
cal_df['end_dt_raw'] = cal_df['End Date'] + ' ' + cal_df['End Time']
cal_df['end_dt_object'] = pd.to_datetime(cal_df['end_dt_raw'], format='%m/%d/%Y %H:%M:%S')
cal_df['event_duration'] = ((cal_df['end_dt_object'] - cal_df['start_dt_object']).astype('timedelta64[m]')) / 60
cal_df.drop(labels=whack_columns, axis='columns', inplace=True)
cal_df.to_csv(duration_file_name, encoding='utf-8', index=False)
cal_summary = cal_df.groupby(['Categories'])[['event_duration']].agg('sum')
cal_summary.to_csv(summary_file_name, encoding='utf-8', index=False)
print(cal_summary)
