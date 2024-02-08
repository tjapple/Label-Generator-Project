
'''
Excel macro will scrape data from Test Submission and export CSV files to designated location.
The macro will call this python script with the Shell command, and this script will load the CSV files into
dataframes, then clean, reformat, combine, and export finished data to a new location, assigned by output_csv_filepath.
This final CSV file can then be loaded into the DYMO software, mapped to corresponding text boxes, and printed out as a batch.

This script and the helper file need to be saved in the location the macro is pulling from. 
Currently that is "L:\Label_Generator_Project\Python_Scripts".

Make sure the computer running the program has python, pandas, and numpy installed. 

Further improvements:
--if labels not desired for certain delay temperatures or humidities, excluding that data from the dataframe.
'''


import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import re
from helpers import *


# Test names that will trigger a return notice for reliability Add as needed. Must be exact test name.
return_after_discharge_reliability = ['CPD - LKG', 'IPD - LKG', 'ODL']

# Key terms that will trigger a return notice for safety. Add as needed. Can be just a partial string.
return_after_discharge_safety_key_terms = ['PD Cell', ' 25%', ' 50%', ' 75%', ' 100%', 'Over-Discharge', 'Forced Discharge'] #space before percents to distinguish between temp and test name


# Set input file locations. These are set in the macro, so must be the same as where macro sends them.
input_performance = r'L:\Label_Generator_Project\Test_Group_CSVs\performance_labels.csv'
input_DD = r'L:\Label_Generator_Project\Test_Group_CSVs\DD_labels.csv'
input_shelf = r'L:\Label_Generator_Project\Test_Group_CSVs\shelf_labels.csv'
input_leakage = r'L:\Label_Generator_Project\Test_Group_CSVs\leakage_labels.csv'
input_bench = r'L:\Label_Generator_Project\Test_Group_CSVs\bench_labels.csv'
input_safety = r'L:\Label_Generator_Project\Test_Group_CSVs\safety_labels.csv'

# File destination of completed and cleaned CSV
output_csv_filepath = r'L:\Label_Generator_Project\FINISHED_LABELS.csv'

#Load data from CSV files into dataframes
performance_df = pd.read_csv(input_performance).dropna(axis=0, how='all')
DD_df = pd.read_csv(input_DD).dropna(axis=0, how='all')
shelf_df = pd.read_csv(input_shelf).dropna(subset=['Cells'], axis=0, how='all')
leakage_df = pd.read_csv(input_leakage).dropna(axis=0, how='all')
bench_df = pd.read_csv(input_bench).dropna(axis=0, how='all')
safety_df = pd.read_csv(input_safety).dropna(axis=0, how='all')






#### Special treatment before merging ####

# Extract experiment number, engineer, lots, cell size, and date made before possibly getting dropped
exp_number = '-'.join(performance_df.loc[0, 'Experiment'].split('-')[:2])
initials = ''.join([part[0].upper() for part in performance_df.loc[0, 'Engineer'].split()])
unformatted_lots = [x for x in performance_df.loc[0, 'Lots'].split()]
cell_size = performance_df.loc[0, 'Boxes'] #cell size initially stored in 'boxes column'

performance_df['Date Made'] = performance_df.loc[0, 'Date Made']          #copy date made into every test group
performance_df['Date Made'] = pd.to_datetime(performance_df['Date Made']) #convert to datetime object
date_made = performance_df.loc[0, 'Date Made']


# Drop rows without cell numbers 
performance_df.dropna(subset="Cells", inplace=True)

# If any performance Tests:
if performance_df.shape[0] > 0:

    # Reformat delay-group names
    performance_df['Test'] = performance_df['Test'].str.replace('\n', ' ', regex=False)

    # Make due date column for RT tests
    performance_df['Due'] = ""                                                #create empty "Due" column
    for test in performance_df['Test']:
        if "RT" in test:
            if "Month" in test:
                # If less than 6 Months, count a month as 4 weeks. Otherwise, count as proper months.
                delay_value = int(test.split()[0]) 
                if delay_value < 6: 
                    due_date = date_made + pd.to_timedelta((delay_value * 4), unit='W')
                else:
                    due_date = date_made + relativedelta(months=delay_value)
                # Check if test is past due and trigger a warning
                if datetime.date.today() > due_date.date():
                    performance_df.loc[performance_df['Test'] == test, 'Due'] = "OOPSIE!! Due:  " + due_date.strftime("%m-%d-%Y")
                else: 
                    performance_df.loc[performance_df['Test'] == test, 'Due'] = "Due:  " + due_date.strftime("%m-%d-%Y")
        
            elif "Week" in test:
                delay_value = int(test.split()[0])
                due_date = date_made + pd.to_timedelta(delay_value, unit='W')
                if datetime.date.today() > due_date.date(): #If overdue, trigger warning
                    performance_df.loc[performance_df['Test'] == test, 'Due'] = "OOPSIE!! Due: " + due_date.strftime("%m-%d-%Y")
                else:
                    performance_df.loc[performance_df['Test'] == test, 'Due'] = "Due: " + due_date.strftime("%m-%d-%Y")
            
            elif "Fresh" in test:
                performance_df.loc[performance_df['Test'] == test, 'Due'] = "Due: ASAP"

    # Duplicate labels per number of boxes needed
    if cell_size not in ['AA', 'AAA', 'C', 'D']: #default to 1 box per delay group
        performance_df['Boxes'] = 1
    else: 
        performance_df['Boxes'] = get_number_of_boxes(performance_df['Cells'], len(unformatted_lots), cell_size)



# Change deep discharge test names
if DD_df.shape[0] > 0:
    DD_df.dropna(subset=['Cells'], inplace=True)
    new_DD_names = []
    for x in DD_df['Test']:
        if 'JIS' in x:    
            new_DD_names.append('JIS')
        elif 'ODL' in x:
            new_DD_names.append('ODL')
        else:
            new_DD_names.append(x)
    DD_df['Test'] = new_DD_names

    # Duplicate labels per number of boxes needed
    if cell_size not in ['AA', 'AAA', 'C', 'D']: #default to 1 box per delay group
        DD_df['Boxes'] = 1
    else: 
        DD_df['Boxes'] = get_number_of_boxes(DD_df['Cells'], len(unformatted_lots), cell_size)




# Change shelf test names
if shelf_df.shape[0] > 0:
    shelf_df.dropna(subset=['Cells'], inplace=True)
    new_shelf_names = [strip_parens_from_name(test) for test in shelf_df['Test']]
    shelf_df['Test'] = new_shelf_names
    shelf_df['Test'] = shelf_df['Test'].str.replace('Undischarged', 'Shelf')

    # Duplicate labels per number of boxes needed
    if cell_size not in ['AA', 'AAA', 'C', 'D']: #default to 1 box per delay group
        shelf_df['Boxes'] = 1
    else: 
        shelf_df['Boxes'] = get_number_of_boxes(shelf_df['Cells'], len(unformatted_lots), cell_size)




# Change leakage test names
if leakage_df.shape[0] > 0:
    leakage_df.dropna(subset=['Cells'], inplace=True)
    new_LKG_names = []
    for x in leakage_df['Test']:
        if 'Undischarged' in x:
            new_LKG_names.append('UD - LKG')
        elif 'Continuous' in x:
            new_LKG_names.append('CPD - LKG')
        elif 'Intermittent' in x:
            new_LKG_names.append('IPD - LKG')
    leakage_df['Test'] = new_LKG_names

    # Duplicate labels per number of boxes needed
    if cell_size not in ['AA', 'AAA', 'C', 'D']: #default to 1 box per delay group
        leakage_df['Boxes'] = 1
    else: 
        leakage_df['Boxes'] = get_number_of_boxes(leakage_df['Cells'], len(unformatted_lots), cell_size)


# Change safety test names
if safety_df.shape[0] > 0:
    safety_df.dropna(subset=['Cells'], inplace=True)
    new_safety_names = [strip_parens_from_name(test) for test in safety_df['Test']]
    safety_df['Test'] = new_safety_names

    # Duplicate labels per number of boxes needed
    if cell_size not in ['AA', 'AAA', 'C', 'D']: #default to 1 box per delay group
        safety_df['Boxes'] = 1
    else: 
        safety_df['Boxes'] = get_number_of_boxes(safety_df['Cells'], len(unformatted_lots), cell_size)




#########################################################################################
#########################################################################################
##### Merge the dataframes #####
combined_df = pd.concat([performance_df, DD_df, shelf_df, leakage_df, bench_df, safety_df], ignore_index=True) #keep in proper order

# Drop rows that have no cell numbers
combined_df['Cells'].replace(["", " ", "NA", "NaN"], np.nan, inplace=True)
combined_df.dropna(subset=['Cells'], inplace=True) #redundant but good backup

# Make sure "Boxes" is in correct format
combined_df['Boxes'].replace(np.nan, 1, inplace=True)
combined_df['Boxes'] = combined_df['Boxes'].astype(int)



# Add a return notice to qualifying tests
combined_df['Return Notice'] = ""
for test in combined_df['Test']:
    if test in return_after_discharge_reliability:
        combined_df.loc[combined_df['Test'] == test, 'Return Notice'] = "Return to Reliability After Discharge"
    for term in return_after_discharge_safety_key_terms:
        if term in test:
            combined_df.loc[combined_df['Test'] == test, 'Return Notice'] = "Return to Safety After Discharge"



# Extract lots from full experiment numbers
unformatted_lots = [x.split('-')[-1] for x in unformatted_lots]
unformatted_lots = [int(x.lstrip('0')) for x in unformatted_lots]
unformatted_lots.sort()
formatted_lots = format_lots(unformatted_lots)

# Update DataFrame
combined_df['Experiment'] = exp_number
combined_df['Engineer'] = initials
combined_df['Lots'] = formatted_lots

# Create final DataFrame with duplicated rows for more than 1 box
duplicated_rows = [combined_df.loc[i:i].copy() for i in combined_df.index for _ in range(combined_df.at[i, 'Boxes'])]
final_df = pd.concat(duplicated_rows, ignore_index=True)




# Export to CSV. This will overwrite current finished_labels.csv file
final_df.to_csv(output_csv_filepath, index=False, mode='w+')


