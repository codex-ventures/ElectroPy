import numpy as np
import pandas as pd
import os  # to get the directory
import shutil  # to move the file in the end
import glob  # to get all excel files in the folder
from datetime import date
from pathlib import Path

# Get the folder directory
path = os.getcwd()

# Get all excel files in the folder
all_files = glob.glob(path + "/*.xlsx")
number_of_files = len(all_files)

# Turning each file into a dataframe
data = {}
file_number = 0

for file in all_files:
    # Turn each file into a dataframe
    filename = Path(file).stem    
    data[filename] = pd.read_excel(all_files[file_number])
    
    # Get rid of rows with empty values
    data[filename] = data[filename].dropna()
    
    # Getting the number of entries from each file
    entries = data[filename]['Event Start Time (ms)'].count()
    
    # Calculate the frequency for each row
    data[filename]['Period (s)'] = abs(data[filename]['Event Start Time (ms)'].diff(periods=-1))/1000
    data[filename]['Frequency (Hz)'] = 1/(data[filename]['Period (s)'])
    
    # Establish the conditions for a possible burst
    data[filename]['Time to Next Event (s)'] = (abs(data[filename]['Event Start Time (ms)'].diff(periods=-1)) + (data[filename]['Event End Time (ms)'] - data[filename]['Event Start Time (ms)']))/1000
    cond1 = data[filename]['Frequency (Hz)'] >= 2
    cond2 = data[filename]['Time to Next Event (s)'] < 2
    
    # Creating an empty list to append the possible bursts of this file
    possible_burst = []
    
    # Create a new dataframe with only the events that follow the above conditions
    for i in range(len(data[filename]['Frequency (Hz)'])):
        if cond1[i] & cond2[i] == True: 
            possible_burst.append(1)
        else:
            possible_burst.append(0)
    
    # Creating a column in the dataframe for possible bursts
    data[filename]['Possible Burst'] = possible_burst
    
    # To prevent issues from data inconsistencies, I've added two rows to the dataframe, one in the beginning and one in the end
    first_row = data[filename].head(1)
    first_row['Possible Burst'] = first_row['Possible Burst'].replace(1,0)
    last_row = data[filename].tail(1)
    last_row['Possible Burst'] = last_row['Possible Burst'].replace(1,0)
    data[filename] = pd.concat([first_row, data[filename].loc[:], last_row]).reset_index(drop=True)
    
    # Creating a new dataframe to establish the last condition for a burst
    burst_detection = pd.DataFrame()
    start_time = []
    end_time = []
    
    for i in range(0, data[filename].shape[0]-1):
        if data[filename].iloc[i]['Possible Burst'] == 0 and data[filename].iloc[i+1]['Possible Burst'] == 1:
            start_time.append(data[filename].iloc[i+1]['Event Start Time (ms)'])
        elif data[filename].iloc[i]['Possible Burst'] == 1 and data[filename].iloc[i+1]['Possible Burst'] == 0:
            end_time.append(data[filename].iloc[i]['Event End Time (ms)'])
    
    burst_detection['Start Time'] = start_time
    burst_detection['End Time'] = end_time
    burst_detection['Burst Time (s)'] = (burst_detection['End Time'] - burst_detection['Start Time'])/1000
    
    # Last condition to get the bursts
    bursts_times = burst_detection[burst_detection['Burst Time (s)'] > 6]
    
    # Statement to check if there are any bursts and end the program if there are not
    if bursts_times.empty is True:
        no_burst_txt = os.path.join(filename + '_results.txt')
        f = open(no_burst_txt, "w+")
        f.write('There are no bursts in this dataset. ')
        f.close()
        file_number += 1
        continue
    
    # Add the number of bursts to the burst times dataframe
    burst_number = []
    for i in range(0, bursts_times.shape[0]):
        burst_number.append(str(i+1))
    bursts_times['Burst Number'] = burst_number
    
    # Get the index of the rows where the bursts start and end
    event_index = []
    for i in range(0, data[filename].shape[0]):
        for x in range(0, bursts_times.shape[0]):
            if data[filename].iloc[i]['Event Start Time (ms)'] == bursts_times.iloc[x]['Start Time']:
                event_index.append(i)
            elif data[filename].iloc[i]['Event End Time (ms)'] == bursts_times.iloc[x]['End Time']:
                event_index.append(i)
    
    # New dataframe with all data from all bursts
    bursts_full_data = pd.DataFrame()
    for i in range(0, data[filename].shape[0]):
        for x in range(0, bursts_times.shape[0]):
            if data[filename].iloc[i]['Event Start Time (ms)'] >= bursts_times.iloc[x]['Start Time'] and data[filename].iloc[i]['Event End Time (ms)'] <= bursts_times.iloc[x]['End Time'] :
                bursts_full_data = bursts_full_data.append(data[filename].iloc[i])
    
    # Create empty dataframes according to the number of bursts
    burst = {}
    for i in burst_number:
        burst[i] = pd.DataFrame()
    
    # Get the data from each burst individually
    n = 1
    for i in range(0, len(event_index), 2):
        burst[n] = data[filename].iloc[event_index[i]:(event_index[i+1]+1)]
        n += 1
    
    # Calculate results for the bursts full data
    total_num_events = bursts_full_data['Event Start Time (ms)'].count()
    total_num_bursts = bursts_times['Burst Number'].count()
    total_avg_num_events = round(total_num_events/total_num_bursts, 2)
    total_avg_burst_time = round((bursts_times['Burst Time (s)'].sum())/total_num_bursts, 2)
    total_avg_peak_amp = round((bursts_full_data['Peak Amp (mV)'].sum())/total_num_events, 2)
    total_avg_frequency = round((bursts_full_data['Frequency (Hz)'].sum())/total_num_events, 2)
    total_dic = {'Filename': [filename],
            'Number of Events': [total_num_events],
            'Number of Bursts': [total_num_bursts],
            'Average Number of Events': [total_avg_num_events],
            'Average Burst Time [s]': [total_avg_burst_time],
            'Average Peak Amplitude [mV]': [total_avg_peak_amp],
            'Average Frequency [Hz]': [total_avg_frequency]}
    total = pd.DataFrame.from_dict(total_dic)
    
    # Calculate results for individual bursts
    num_events = []
    burst_time = []
    avg_peak_amp = []
    avg_frequency = []
    for i in range(0, bursts_times.shape[0]):
        num_events.append(burst[i+1]['Event Start Time (ms)'].count())
        burst_time.append(np.round(bursts_times.iloc[i]['Burst Time (s)'], 2))
        avg_peak_amp.append(np.round((burst[i+1]['Peak Amp (mV)'].sum())/num_events[i], 2))
        avg_frequency.append(np.round((burst[i+1]['Frequency (Hz)'].sum())/num_events[i], 2))
    
    individual = pd.DataFrame(list(zip(burst_number, num_events, burst_time, avg_peak_amp, avg_frequency)),
                                    columns=['Burst Number', 'Number of Events', 'Burst Time [s]', 'Average Peak Amplitude [mV]', 'Average Frequency [Hz]'])
    
    # Create a new folder
    new_folder = os.path.join(path, filename)
    if not os.path.exists(new_folder):
        os.mkdir(new_folder)
    
    # Get the directory of the new folder
    for dirpath, dirnames, filenames in os.walk(path):
        for i in filenames:
            if i == filename:
                i = os.path.join(dirpath, i)
    new_folder_path = dirpath
    
    # Save an excel file with the bursts full and individual data
    writer = pd.ExcelWriter(new_folder_path + '\\' + filename + '_results.xlsx', engine='xlsxwriter')
    total.to_excel(writer, sheet_name='All Bursts', index=False)
    individual.to_excel(writer, sheet_name='Individual Bursts', index=False)    
    writer.close()
    
    # Move the original file to the new folder (checking if there is already one there)
    if os.path.exists(new_folder_path + '\\' + filename + '.xlsx'):
        new_name = filename + '_new'
        shutil.move(path + '\\' + filename + '.xlsx', new_folder_path + '\\' + new_name + '.xlsx')
    else:
        shutil.move(path + '\\' + filename + '.xlsx', new_folder_path + '\\' + filename + '.xlsx')
    
    file_number += 1