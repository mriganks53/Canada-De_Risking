# Developed by :- Mrigank Saxena
import pandas as pd
import os
from datetime import datetime
import logging
from openpyxl import Workbook

# =============================================================================
#                         Function for converting Report 1
# =============================================================================

def report_1(file_name):
    # file_name = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\ManuLife Reports\Manulife - Jan23 - Dec23  Insured Experience Report.xlsx"
    # Extracting sheet names
    df = pd.ExcelFile(file_name)
    sheet_name = df.sheet_names

    for i in range(len(sheet_name)):
        u = sheet_name[i]
        
        # Extracting data from 1st sheet
        if u.split()[0] == 'Experience':
            df_exp = pd.read_excel(file_name, sheet_name= sheet_name[i])

            benefit = df_exp[df_exp.columns[0]].str.contains("Benefit")

            for i in range(len(benefit)):
                if benefit[i] == True:
                    df_1 = df_exp.iloc[i:,:]
            
            for i in range(len(df_1.columns)):
                u = df_1.columns[i]
                if u.split()[0] == 'Unnamed:':
                    
                    for x in range(len(df_1.index)):
                        u = str(df_1.iloc[x,i]).lower()
                        
                        # Finding Benefit
                        if u == 'benefit':
                            df_1 = df_1.rename(columns={df_1.columns[i] : df_1.iloc[x+1,i]})
                        if u == 'division':
                            df_1 = df_1.rename(columns={df_1.columns[i] : df_1.iloc[x,i]})
                        if u == 'month':
                            df_1 = df_1.rename(columns={df_1.columns[i] : df_1.iloc[x,i]})
                        if u == 'ehc':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "EHC Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "EHC Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "EHC Ratio"})
                        if u == 'dental':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "Dental Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "Dental Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "Dental Ratio"})
                        if u == 'std':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "STD Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "STD Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "STD Ratio"})
                        if u == 'ltd':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "LTD Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "LTD Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "LTD Ratio"})
                        if u == 'basic life':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "Basic Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "Basic Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "Basic Ratio"})
                        if u == 'dependent life':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "Dependent Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "Dependent Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "Dependent Ratio"})
                        if u == 'opt life':
                            # print(df_1.iloc[x+1,i],df_1.iloc[x+1,i+1],df_1.iloc[x+1,i+2])
                            if df_1.iloc[x+1,i] == "Premiums":
                                df_1 = df_1.rename(columns={df_1.columns[i] : "OPT Premiums"})
                            if df_1.iloc[x+1,i+1] == "Claims":
                                df_1 = df_1.rename(columns={df_1.columns[i+1] : "OPT Claims"})
                            if df_1.iloc[x+1,i+2] == "Ratio":
                                df_1 = df_1.rename(columns={df_1.columns[i+2] : "OPT Ratio"})
            
            # All names are 
            df_1 = df_1.iloc[2:,:]

            # Reset index
            df_1.reset_index(drop= True, inplace= True)
            
            # Filling NA
            for i in range(len(df_1.index)):
                if pd.isna(df_1.iloc[i,0]) == True:
                    df_1.iloc[i,0] = df_1.iloc[i-1,0]
                if pd.isna(df_1.iloc[i,1]) == True:
                    df_1.iloc[i,1] = df_1.iloc[i-1,1]
            
            
                # The function `product_df` filters columns in a DataFrame based on a specified pattern,
                # renames columns, and adds a new column with a specified benefit name.
                
                # :param filter_column: The `filter_column` parameter is used to specify a pattern or
                # regular expression to filter columns in the DataFrame `df_1`
                # :param replace_name: The `replace_name` parameter is used to specify the string that you
                # want to replace in the column names of the DataFrame `df`
                # :param benefit_name: The `benefit_name` parameter is used to specify the value that will
                # be assigned to the 'Benefit' column in the DataFrame `df`

            def product_df(filter_column, replace_name, benefit_name):
                df = df_1.filter(regex=filter_column)
                df = df_1.iloc[:,:3].join(df)
                df.columns = df.columns.str.replace(replace_name,'')
                df['Benefit'] = benefit_name

                return df
            # EHC dataframe
            df_ehc = product_df(filter_column= '^EHC', replace_name= 'EHC ', benefit_name= 'EHC')
            # Dental dataframe
            df_dental = product_df(filter_column= '^Dental', replace_name= 'Dental ', benefit_name= 'Dental')
            # STD dataframe
            df_std = product_df(filter_column='^STD', replace_name= 'STD ', benefit_name= 'STD')
            # LTD dataframe
            df_ltd = product_df(filter_column='^LTD', replace_name= 'LTD ', benefit_name= 'LTD')
            # Basic dataframe
            df_basic = product_df(filter_column='^Basic', replace_name= 'Basic ', benefit_name= 'Basic Life')
            # Dependent dataframe
            df_dep = product_df(filter_column='^Dependent', replace_name= 'Dependent ', benefit_name= 'Dependent')
            # Opt Life dataframe
            df_opt = product_df(filter_column='^OPT', replace_name= 'OPT ', benefit_name= 'Opt Life')

            # Merging all dataframe
            # The line `final_df = pd.concat([df_ehc, df_dental, df_std, df_ltd, df_basic, df_dep,
            # df_opt], axis=0)` is concatenating the DataFrames `df_ehc`, `df_dental`, `df_std`,
            # `df_ltd`, `df_basic`, `df_dep`, and `df_opt` along axis 0, which means it is stacking
            # these DataFrames on top of each other to create a single DataFrame `final_df`.

            final_df = pd.concat([df_ehc,df_dental,df_std,df_ltd,df_basic,df_dep,df_opt], axis=0)
            final_df = final_df[(final_df['Policy'] != 'Total') &(final_df['Month'] != 'Total')]
            final_df["Month"] = pd.to_datetime(final_df["Month"],format= '%b-%Y')
    return final_df

# =============================================================================
#                         Function for converting Report 2
# =============================================================================
        
def report_2(file_name, sheetname):
    # Extracting sheet names
    # file_name = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\ManuLife Reports\Manulife - Jan23 - Dec23  Insured Experience Report.xlsx"
    df = pd.ExcelFile(file_name)
    sheet_name = df.sheet_names
    
    for i in range(len(sheet_name)):
        u = sheet_name[i]
        if u.split()[0] == sheetname:
            df_ehc = pd.read_excel(file_name, sheet_name= sheet_name[i])

            # Finding value
            policy = df_ehc[df_ehc.columns[0]].str.contains("Policy")
            for i in range(len(df_ehc.index)):
                if policy[i] == True:
                    df_ehc = df_ehc.iloc[i:,:]

            # Changing column Name
            df_ehc.columns = df_ehc.iloc[0]   
            df_ehc = df_ehc.drop(df_ehc.index[0])

            df_ehc = pd.melt(df_ehc, id_vars=[df_ehc.columns[0], df_ehc.columns[1], df_ehc.columns[2]], var_name= 'Month', value_name= 'Claims Premium')

            df_ehc = df_ehc[(df_ehc['Month'] != 'Total') & (df_ehc['Policy'] != 'Total')]
            df_ehc["Month"] =  pd.to_datetime(df_ehc["Month"],format= '%b-%Y')

    return df_ehc


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

path = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\ManuLife Reports"
path_save = r"\\Innda11fs01v.mercer.com\scratch$\Canada Project\ManuLife Reports\Filtered Data"
files = os.listdir(path)
filtered_files = [file for file in files if file.endswith('Insured Experience Report.xlsx')]
file_name = filtered_files[0]
carrier = "Manulife"
try:
    # Call function1()
    test_1 = report_1(os.path.join(path,file_name))
    test_1.to_excel(f"{path_save}/{carrier}_Report_1_test.xlsx", index = False)
    sheet.append([datetime.now(), "Report 1", "Success"])
except Exception as e:
    logger.error("An error occurred in function1: %s", str(e))
    sheet.append([datetime.now(), "Report 1", f"Error"])

try:
    # Call function2()
    test_2 = report_2(os.path.join(path,file_name), 'EHC')
    test_2.to_excel(f"{path_save}/{carrier}_Report_2_test.xlsx", index = False)
    sheet.append([datetime.now(), "Report 2", "Success"])
except Exception as e:
    logger.error("An error occurred in function2: %s", str(e))
    sheet.append([datetime.now(), "Report 2", "Error"])

try:
    # Call function3()
    test_3 = report_2(os.path.join(path,file_name), 'Dental')
    test_3.to_excel(f"{path_save}/{carrier}_Report_3_test.xlsx",index = False)
    sheet.append([datetime.now(), "Report 3", "Success"])
except Exception as e:
    logger.error("An error occurred in function3: %s", str(e))
    sheet.append([datetime.now(), "Report 3", "Error"])


# Save the Excel workbook with a timestamp
timestamp = datetime.now().strftime("%Y%m%d%H%M")
log_file_excel = f"function_log_{timestamp}.xlsx"
wb.save(f"{path}/{log_file_excel}")

# Close the file handler
file_handler.close()

# Remove the file handler from the logger
logger.removeHandler(file_handler)


































































































