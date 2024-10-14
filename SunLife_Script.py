# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 19:15:54 2023

@author: mrigank-saxena
"""

import pandas as pd
import re
import logging
from openpyxl import Workbook
from datetime import datetime


# =============================================================================
#                               Sunlife Carrier 
# =============================================================================


# =============================================================================
#                         Function for converting Report 1
# =============================================================================

def report_1(data):
    Report_1 = pd.read_excel(data, skipfooter= 4)
    Group = Report_1[Report_1.columns[0]].str.contains("Experience Group")
    Group_values = []
    for i in range(len(Group)):
        if Group.iloc[i] == True:
            Group_values.append(Report_1.iloc[i,0])
    
    df = []
    
    for i in range(len(Group_values)):
        indices = Report_1[Report_1[Report_1.columns[0]] == Group_values[i]].index
        indices_next = Report_1[Report_1[Report_1.columns[0]] == Group_values[i+1]].index if i+1 < len(Group_values) else None
        if indices_next is not None:
            start_index = indices[0]
            end_index = indices_next[0]
            small_df = Report_1.iloc[start_index:end_index]
        else:
            start_index = indices[0]
            end_index = Report_1.index.max()
            small_df = Report_1.iloc[start_index:end_index]
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        filterded_rows = small_df[small_df[small_df.columns[0]].str.contains("|".join(months), case= False, na= False)]
        pd.options.mode.chained_assignment = None
        filterded_rows['Group_values'] = Group_values[i]
        df.append(filterded_rows)
    
    combined_df = pd.concat(df)
    x = combined_df.columns[0]
    contract_number = re.search(r'\d+', x).group()
    combined_df['Contract'] = contract_number
    combined_df.iloc[:, 1:7] = combined_df.iloc[: , 1:7].astype(str)
    combined_df.iloc[:, 1:7] = combined_df.iloc[: , 1:7].apply(lambda x: x.str.replace('            ', 'Drop'))
    drop_column = combined_df.columns[combined_df.eq('Drop').any()]
    drop_nan = combined_df.columns[combined_df.eq('nan').any()]
    combined_df = combined_df.drop(drop_column, axis= 1)
    combined_df = combined_df.drop(drop_nan, axis= 1)
    Column_name = ['Months','Number of Lives','Volume','Premium','Paid Claims']
    combined_df = combined_df.rename(columns=dict(zip(combined_df.columns, Column_name)))
    split_values = combined_df.pop('Group_values').str.split('-', n = 1 ,expand = True)
    combined_df[['Benefit', 'Class']] = split_values.iloc[:,:2]
    combined_df['Class'] = combined_df['Class'].str.replace('Experience Group', '').str.strip()
    combined_df.iloc[:, 1:3] = combined_df.iloc[:, 1:3].astype(int)
    print("Report 1 created and saved")
    return combined_df


# =============================================================================
#               Function for converting  Report 2
# =============================================================================

def report_2(data):
    Report_1 = pd.read_excel(data, skipfooter= 4)
    Group = Report_1[Report_1.columns[0]].str.contains("Experience Group")
    Group_values = []
    for i in range(len(Group)):
        if Group.iloc[i] == True:
            Group_values.append(Report_1.iloc[i,0])
    
    df = []
    
    for i in range(len(Group_values)):
        indices = Report_1[Report_1[Report_1.columns[0]] == Group_values[i]].index
        indices_next = Report_1[Report_1[Report_1.columns[0]] == Group_values[i+1]].index if i+1 < len(Group_values) else None
        if indices_next is not None:
            start_index = indices[0]
            end_index = indices_next[0]
            small_df = Report_1.iloc[start_index:end_index]
        else:
            start_index = indices[0]
            end_index = Report_1.index.max()
            small_df = Report_1.iloc[start_index:end_index]
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        filterded_rows = small_df[small_df[small_df.columns[0]].str.contains("|".join(months), case= False, na= False)]
        pd.options.mode.chained_assignment = None
        filterded_rows['Group_values'] = Group_values[i]
        df.append(filterded_rows)
    
    combined_df = pd.concat(df, ignore_index= True)
    x = combined_df.columns[0]
    contract_number = re.search(r'\d+', x).group()
    combined_df['Contract'] = contract_number
    combined_df.iloc[:, 1:10] = combined_df.iloc[: , 1:10].astype(str)
    combined_df.iloc[:, 1:10] = combined_df.iloc[: , 1:10].apply(lambda x: x.str.replace('            ', 'Drop'))
    drop_column = combined_df.columns[combined_df.eq('Drop').any()]
    drop_nan = combined_df.columns[combined_df.eq('nan').any()]
    combined_df = combined_df.drop(drop_column, axis= 1)
    combined_df = combined_df.drop(drop_nan, axis= 1)

    Column_name = ['Months','Single','Family','Total','Premium','Paid Claims','Ratio']
    combined_df = combined_df.rename(columns=dict(zip(combined_df.columns, Column_name)))
    split_values = combined_df.pop('Group_values').str.split('-',n=1 ,expand = True)
    combined_df[['Benefit', 'Class']] = split_values.iloc[:,:2]
    combined_df['Class'] = combined_df['Class'].str.replace('Experience Group', '').str.strip()
    
    combined_df.iloc[:, 1:4] = combined_df.iloc[:, 1:4].astype(int)
    print("Report 2 created and saved")
    return combined_df


# =============================================================================
#           Function for converting report 3 and 4 report
# =============================================================================


def report_3_4(data):
    Report_3 = pd.read_excel(data, skipfooter= 2)
    for i in range(len(Report_3.columns)):
        # print(Report_3.columns[i])
        Prior_year = Report_3[Report_3.columns[i]].str.contains("Prior Period").fillna(False)
        for j in range(len(Prior_year)):
            if Prior_year[j] == True:
                # print(i)
                df_new = Report_3.iloc[:,:i]
                # df_old = Report_3.iloc[:,[0] + list(range(i, Report_3.shape[1]))]
    
    Total_check = df_new[df_new.columns.max()].str.contains('Total').fillna(False)
    s = Total_check.index[Total_check].tolist()
    if len(s) == 0:
        last_column_index = df_new.columns.get_loc(df_new.columns.max())
        df_new = df_new.iloc[:, :last_column_index]
    
    # Adding year or we can add report year
    x = 'Oct 2023'
    # y = '10-31-2023'
    # Parse string and convert to datetime
    dt = datetime.strptime(x, "%b %Y")
    formatted_dt = dt.strftime("%b %Y")
    
    df_new['Month From'] = formatted_dt
    # df_new['Month To'] = pd.to_datetime(y)
    
    # df_old['Month From'] = pd.to_datetime(x) - pd.DateOffset(years = 1)
    # df_old['Month To'] = pd.to_datetime(y) - pd.DateOffset(years = 1)
    
    # Loop through loop and renaming column names
    # dataframes = [df_new,df_old]
    dataframes = [df_new]
    Column_name = ['Service Type','Amount Submitted','Amount Eligible','Amount Paid','%of Total Amount Paid','Month']
    
    for df in dataframes:
        df.columns = Column_name
    
    # =============================================================================
    #     Creating report and loop through column by which report can be made
    # =============================================================================
    
    df_combined = []
    
    for df in dataframes:
        # Finding "Details" and trimming
        Report_3_test = df.copy()
        Report_3_test[Report_3_test.columns[0]] = Report_3_test[Report_3_test.columns[0]].fillna('')
        result = Report_3_test[Report_3_test.columns[0]].str.contains("Details") 
        index_value = result.index[result].to_list()
        
        for i in range(len(index_value)):
            # Finding "Total" and trimming the data
            
            if index_value is not None:
                start_index = index_value[i]
                end_index = Report_3_test.index.max()
                new_df = Report_3_test.loc[start_index:end_index]
                Total_remove = new_df[new_df.columns[0]].str.contains("Total")
                Total_index_value = Total_remove.index[Total_remove].to_list()
                new_df = new_df.drop(Total_index_value)
                new_df = new_df.reset_index(drop = True)
                relation = new_df[new_df.columns[0]].str.contains("Relation")
                relation_value = relation.index[relation].to_list()
                
                for j in range(len(relation_value)):
                    # Trimming Data with "Relation"
                    
                    start_index_relation = relation_value[j]
                    end_index_relation = relation_value[j+1] if j+1 <len(relation_value) else None
                    # print(start_index_relation,end_index_relation)
                    # Trimming Data "Group" wise and adding "Relationship" column
                    
                    if end_index_relation is not None:
                        # Created small dataframe Group wise

                        small_df = new_df.loc[start_index_relation:end_index_relation]
                        small_df['Relation'] = new_df.iloc[start_index_relation,0]
                        small_df = small_df.reset_index(drop= True)
                        Exp_group = small_df[small_df.columns[0]].str.contains('Group')
                        Exp_index_value = Exp_group.index[Exp_group].to_list()
                        for r in range(len(Exp_index_value)):
                            # Created small dataframe Relation wise
                            
                            start_group_value = Exp_index_value[r]
                            end_group_value = Exp_index_value[r+1] if r+1 < len(Exp_index_value) else small_df.index.max()
                            small_df_group = small_df.loc[start_group_value:end_group_value]
                            small_df_group['Class'] = small_df.iloc[start_group_value,0]
                            small_df_group = small_df_group.reset_index(drop= True)
                            small_df_group = small_df_group.drop(axis= 0, index=0)
                            small_df_group = small_df_group.drop(axis= 0, index= small_df_group.index.max())
                            test = small_df_group[small_df_group.columns[1]][small_df_group[small_df_group.columns[1]].isna()].index
                            
                            for k in range(len(test)):
                                # Created small dataframe with "Drug"
                                
                                r = small_df_group.iloc[test[k]-1, 0]
                                start_drug = test[k]
                                end_drug = test[k+1] if k+1 < len(test) else small_df_group.index.max()
                                small_drug = small_df_group[start_drug:end_drug]
                                small_drug['Group'] = r
                                column_value = small_drug[small_drug.columns[0]].str.contains('  ')
                                column_spaces = column_value.index[column_value].to_list()
                                small_drug = small_drug.drop(column_spaces)
                                df_combined.append(small_drug)
                            
                    else:
                        # This loop if for "Child" and other which are not identified
                        end_index_relation = new_df.index.max()
                        small_df = new_df.loc[start_index_relation:end_index_relation]
                        small_df['Relation'] = new_df.iloc[start_index_relation,0]
                        small_df = small_df.reset_index(drop = True)
                        Exp_group = small_df[small_df.columns[0]].str.contains('Group')
                        Exp_index_value = Exp_group.index[Exp_group].to_list()
                        
                        for r in range(len(Exp_index_value)):
                            # Created small dataframe Relation wise
                            
                            start_group_value = Exp_index_value[r]
                            end_group_value = Exp_index_value[r+1] if r+1 < len(Exp_index_value) else small_df.index.max()
                            small_df_group = small_df.loc[start_group_value:end_group_value]
                            small_df_group['Class'] = small_df.iloc[start_group_value,0]
                            small_df_group = small_df_group.reset_index(drop= True)
                            small_df_group = small_df_group.drop(axis= 0, index=0)
                            small_df_group = small_df_group.drop(axis= 0, index= small_df_group.index.max())
                            test = small_df_group[small_df_group.columns[1]][small_df_group[small_df_group.columns[1]].isna()].index
                            
                            for k in range(len(test)):
                                # Created small dataframe with "Drug"
                                
                                r = small_df_group.iloc[test[k]-1, 0]
                                start_drug = test[k]
                                end_drug = test[k+1] if k+1 < len(test) else small_df_group.index.max()
                                small_drug = small_df_group[start_drug:end_drug]
                                small_drug['Group'] = r
                                column_value = small_drug[small_drug.columns[0]].str.contains('  ')
                                column_spaces = column_value.index[column_value].to_list()
                                small_drug = small_drug.drop(column_spaces)
                                df_combined.append(small_drug)
        
            else:
                print("In Report Detail bifurcation is not available kindly review the report manually")
    
    # =============================================================================
    #               List is appended to a single dataframe
    # =============================================================================
    
    combined_df = pd.concat(df_combined, axis= 0)
    combined_df = combined_df.reset_index(drop= True)
    Blank_count = combined_df[combined_df[combined_df.columns[1]].isna()].index
    combined_df = combined_df.drop(Blank_count)
    
    # Removing Extract spacing and extra value in string
    
    x = Report_3.columns[0]
    contract_number = re.search(r'\d+', x).group()
    combined_df['Contract'] = contract_number
    combined_df['Relation'] = combined_df['Relation'].str.replace('Relation','').str.strip()
    combined_df['Class'] = combined_df['Class'].str.replace('Experience Group', '').str.strip()
    combined_df[['Relation','Class']] = combined_df[['Relation','Class']].applymap(lambda x: x.replace(':', '').strip())
    combined_df.iloc[:, 1:4] = combined_df.iloc[: , 1:4].astype(int)
    print("Report 3 & 4 created and saved")
    # Unique ID for SQL use 
    return combined_df
    
# =============================================================================


# =============================================================================
#                           Running all functions
# =============================================================================

# Create a logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Create a file handler to write the log to a file
log_file = "function_log.txt"
file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.INFO)

# Create a formatter to include timestamps in the log
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# Add the file handler to the logger
logger.addHandler(file_handler)

# Create an Excel workbook and sheet to store the log
wb = Workbook()
sheet = wb.active
sheet.append(["Timestamp", "Function", "Status"])

# File Path
path = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\SunLife Reports"
path_save = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\SunLife Reports\Filtered Data"
carrier = "Sunlife"
try:
    # Call function1()
    test_1 = report_1(path+"/report 1.xls")
    test_1.to_csv(f"{path_save}/{carrier}_Report_1_test.csv", index = False)
    sheet.append([datetime.now(), "Report 1", "Success"])
except Exception as e:
    logger.error("An error occurred in function1: %s", str(e))
    sheet.append([datetime.now(), "Report 1", f"Error"])

try:
    # Call function2()
    test_2 = report_2(path+'/report 2.xls')
    test_2.to_csv(f"{path_save}/{carrier}_Report_2_test.csv", index = False)
    sheet.append([datetime.now(), "Report 2", "Success"])
except Exception as e:
    logger.error("An error occurred in function2: %s", str(e))
    sheet.append([datetime.now(), "Report 2", "Error"])

try:
    # Call function3()
    test_3 = report_3_4(path+'/report 3.xls')
    test_3.to_csv(f"{path_save}/{carrier}_Report_3_test.csv", encoding = "utf-8" ,index = False)
    sheet.append([datetime.now(), "Report 3", "Success"])
except Exception as e:
    logger.error("An error occurred in function3: %s", str(e))
    sheet.append([datetime.now(), "Report 3", "Error"])

try:
    # Call function4()
    test_4 = report_3_4(path+'/report 4.xls')
    test_4.to_csv(f"{path_save}/{carrier}_Report_4_test.csv", encoding = "utf-8" ,index = False)
    sheet.append([datetime.now(), "Report 4", "Success"])
except Exception as e:
    logger.error("An error occurred in function4: %s", str(e))
    sheet.append([datetime.now(), "Report 4", "Error"])

# Save the Excel workbook with a timestamp
timestamp = datetime.now().strftime("%Y%m%d%H%M")
log_file_excel = f"function_log_{timestamp}.xlsx"
wb.save(f"{path}/{log_file_excel}")

# Close the file handler
file_handler.close()

# Remove the file handler from the logger
logger.removeHandler(file_handler)


