#!/u/dramsoft/.conda/envs/py37/bin/python3

########################################################################################################
# Author:           ryancaraway / mhassana
# Date updated:     8-20-2020
# Description:      This script pulls OPS ELC yield data, parses it, performs data manipulations to get
#                   tables sorted by fallout for individual registers and groups, and then writes those
#                   tables to an xlsx file.
#
#                   From the perspective of yield reporting, when the pareto is generated "this week"
#                   referres to the complete work week data which ended on the previous thursday and
#                   "WTD" referres to whatever work week is in progress as this program pulls data
#                   resulting in only a partial week's worth of data.
#
# Psuedocode:
# Imports / Global variables
# 1. Extract OPS Yield data and put it into csv files broken out by ELC step
# 2. Parse the csv data files into a list of dataframes
# 3. Perform dataframe manipulations to get sorted registers and groups by fallout and
#    write to xlsx file
# main function: handles arguments and program flow
# 
# Future Plans
# TODO: Create option to mesh fallout across steps for the overall pareto
# TODO: Add regular expression matching for grouping registers
# TODO: Add a breakout report of what registers were bucketed into which groups along with their fallout
########################################################################################################
import sys              # Get program arguments
import os               # Enter system commands
import pandas as pd     # Hold and allow manipulations on yield data
import numpy as np
import csv              # csv file i/o
import xlsxwriter       # Write and format xlsx files
import re               # Regular expression pattern matching

# Global Variables
#num_ww = 3  # number of work weeks to pull OPS data for. Currently only support 3 weeks: "Last week", "This week", and "WTD"
output_workbook_name = 'OPS_Data.xlsx'
grouping_filename = 'define_groups.txt'
comment = 'comments.txt'
comment_2 = 'reg_comments.txt'
group_user_columns = {}
register_user_columns = {}
grouping_definitions = {}
files = []
split_token = '^' # Split token used for comments.txt and reg_comments.txt
elc_steps = ['PGSRT','BURN','HSRT','CFIN']
sheets = []
file_directory = 'csv_files/' # directory for supporting files such as csv data files

###########################
# 1. Extract OPS Yield data
###########################
def extract_ops_data(designID,fabNU,config,package_type,tww):
    if not os.path.exists(file_directory): # Make directory for csv files if it doesn't already exist
        os.makedirs(file_directory)
#    os.system('rm ' + output_workbook_name) # remove old workbook
    os.system('rm ' + file_directory + "*.csv") # remove old csv data files
    
    # Pull yield data from the OPS report database for each step, config and package type
    for step in elc_steps:
        for conf in config:
            for pack in package_type:
                if (conf == 'x8') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|X8/SDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+' >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x16') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|X16/SDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+' >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'combo') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|SDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+'  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x8') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|X8/DDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+' >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x16') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|X16/DDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+' >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'combo') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -stepreg="'+step+'|DDP||ALL" -tww='+tww+' +tww='+str(int(tww)-2)+' -fab='+fabNU+' >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                    os.system('/u/prbsoft/scripts/tpose '+file_directory + step +'_'+conf+'_'+pack+'.csv')

    # Append file and Sheet names to include configs and package types
    for step in elc_steps:
        for conf in config:
            for pack in package_type:
                files.append(step+'_'+conf+'_'+pack+'.csv')
                sheets.append(step + '_' + conf + '_'+ pack)
            
    #Trim characters from the register entries
    for fi in files:
        readCSV = csv.reader(open(file_directory + fi), delimiter = ',')
        lines = list(readCSV)
        for line in lines:
            if ((line[0].find("PGSRT") != -1) | (line[0].find("BURN") != -1) | (line[0].find("HSRT") != -1) | (line[0].find("CFIN") != -1)):
               tokens = line[0].split("|")              # returns a list of 2 tokens, the first being the step and the second being SDP_register_REG
               line[0] = tokens[1]
            if ( (line[0].find("X4") != -1) | (line[0].find("X8") != -1) | (line[0].find("X16") != -1) ):
                tokens = line[0].split("_",1)        # returns a list of 2 tokens, want to get rid of the leading "_X8/X16"                
                line[0] = tokens[1] 
            if (line[0].find("SDP") != -1):          # retreive the second token containing the register
                tokens = line[0].split("_",1)        # returns a list of 2 tokens, want to get rid of the leading "SDP_" 
                line[0] = tokens[1]
            if (line[0].find("REG") != -1):             
                tokens = line[0].rsplit("_",1)       # returns a list of 2 tokens, want to get rid of the trailing "_REG"             
                line[0] = tokens[0]                  # retreive the first token containing the trimmed register
        writeCSV = csv.writer(open(file_directory + fi, 'w'))
        writeCSV.writerows(lines)

#########################################
# 2. Parse the csv data files
#########################################

# Read each csv file into a Pandas DataFrame
# Returns a list of DataFrames where each index corrosponds to an ELC step
def parse_csv_files():    
    data_frames = []    
    for file in files:
        this_step_data_frame = pd.read_csv(file_directory + file)
        data_frames.append(this_step_data_frame)
    return data_frames

#############################################
# Helper functions for generate_xlsx_document
#############################################

# Get the weighted percent for each step for the Overall Pareto
def get_weighted_percent(row, weight_pgsrt, weight_burn, weight_hsrt, weight_cfin):
    cols = list(row)
    if row['STEP'] == "PGSRT":
        return (row[1] * weight_pgsrt)
    if row['STEP'] == "BURN":
        return (row[1] * weight_burn)
    if row['STEP'] == "HSRT":
        return (row[1] * weight_hsrt)
    if row['STEP'] == "CFIN":
        return (row[1] * weight_cfin)

# Get the weighted percent for each step for the Overall Pareto last week
def get_weighted_percent_last_week(row, weight_pgsrt, weight_burn, weight_hsrt, weight_cfin):
    cols = list(row)
    if row['STEP'] == "PGSRT":
        return (row[2] * weight_pgsrt)
    if row['STEP'] == "BURN":
        return (row[2] * weight_burn)
    if row['STEP'] == "HSRT":
        return (row[2] * weight_hsrt)
    if row['STEP'] == "CFIN":
        return (row[2] * weight_cfin)

# Check that value is not nan
def isNaN(string):
    return string != string

# Takes a dataframe and prints it to an xlsx document
def write_to_xlsx(df,user_columns, worksheet, start_col=0, total=-1, total_lastweek=-1):
    # Create a table from the data frame
    df = df.fillna("-1") # Fill null values with negative 1
    df_cols = list(df)
    table = []
    for col in df_cols:
        table.append(df[col])
    # Print the headings to the worksheet
    worksheet_row = 0
    worksheet_col = start_col
    for heading in df_cols:
        if (heading in user_columns) or heading=="WEIGHTED" or heading=="WEIGHTED_LAST_WW" or heading == "GROUP" or heading == "REGISTER" or heading == "STEP":
            worksheet.write(worksheet_row, worksheet_col, heading)
        else:
            worksheet.write(worksheet_row, worksheet_col, "WW" + heading[len(heading)-2:len(heading)]) # reformat work week header: WWXX
        worksheet_col += 1
    # Print the table to the worksheet
    worksheet_row = 1
    worksheet_col = start_col
    for column in table:
        for item in column:
            if ((item == "-1") | (item == "nan")):
                worksheet.write(worksheet_row, worksheet_col, " ")
            else:
                worksheet.write(worksheet_row, worksheet_col, item)
            worksheet_row += 1
        worksheet_col += 1
        worksheet_row = 1  
    if (total != -1) and (total_lastweek != -1):   # Default skip this unless total passed as arg
        if (isNaN(total)): # Check if total fallout of this week is zero
            worksheet.write(0,start_col+len(user_columns)+3, "Accounted Yield")
            worksheet.write(1,start_col+len(user_columns)+3, "NULL")            
        else:
            worksheet.write(0,start_col+len(user_columns)+3, "Accounted Yield")
            worksheet.write(1,start_col+len(user_columns)+3, total)            
        if (isNaN(total_lastweek)): # Check if total fallout of last week is zero
            worksheet.write(0,start_col+len(user_columns)+4, "Accounted Last Yield")
            worksheet.write(1,start_col+len(user_columns)+4, "NULL")
        else:
            worksheet.write(0,start_col+len(user_columns)+4, "Accounted Last Yield")
            worksheet.write(1,start_col+len(user_columns)+4, total_lastweek)


#########################################
# 3. Write to xlsx file
#########################################

# This function takes a list of Pandas DataFrames where each index corrosponds to an ELC step and
# creates register and group dataframes by week/step/config and then prints them to separate sheets in an xlsx file
# Outer loop iterates over each DataFrame in the list to iterate over each step/config/package_type
# Inner loop iterates over the columns of the dataframe to look at the work-weeks individually
def generate_xlsx_document(data_frames,sheets,config,user_columns,package_type, WTD_mode,tww):
    overall_paretos = {}
    for conf in config:
        overall_paretos.update( {conf : pd.DataFrame()} ) # Update Overall Pareto with configs
    yield_dic = {}
    output_workbook_name2 = output_workbook_name.split(".xlsx")[0]+"_"+str(tww)+".xlsx" #rename file name with tww
    workbook = xlsxwriter.Workbook(output_workbook_name2) 
    # Write each data frame to a worksheet for its associated step
    frame_index = -1
    # Iterate over each step (one step per list item in the list of data frames)
    for df in data_frames:
        frame_index += 1 # index to pull the label from "sheets" list above
        tokens = sheets[frame_index].split("_")
        step = tokens [0]
        configs = tokens[1]              
        # Iterate over each work week's data (1 work week per column)
        cols = list(df)
        i = 1 # skip the first column (this one holds the registers)
        if (WTD_mode == True):
            i = 2
        register_df = df.filter([cols[0], cols[i+1], cols[i]], axis=1) # Grab the first column (registers) and the i-th column (this week's data)
        register_df.columns=["REGISTER", cols[i+1], cols[i]]
        if len(register_user_columns) != 0:
            j = -1
            for col in user_columns:
                j=j+1
                col_dict = {} # Dictionary with GROUP as key and value for just this column instead of list of values in group_user_columns.  Used for easy key:value mapping on the DataFrame
                for key in register_user_columns:
                    value = register_user_columns[key]
                    if (key.find("*") != -1): #  Check for Special Step
                        special_step = key.split("*",1)
                        key = special_step[0]
                        this_step = special_step[1]
                    else:
                        this_step = "all"
                    if (key.find("$") != -1): # Check for Special Config
                        special_conf = key.split("$",1)
                        key = special_conf[0]
                        conf = special_conf[1]
                    else:
                        conf = "all"                                                      
                    if ((str(conf) == configs) or (str(conf) == ("all"))) and ((str(this_step) == step) or (str(this_step) == ("all"))):
                        col_dict.update( { key : value[j] } )
                        register_df[col] = register_df['REGISTER'].map(col_dict)
                        register_df[col] = [str(x) for x in register_df[col]]
                    else:
                        pass
        # Group fallout according to user-defined grouping_definitions
        if len(grouping_definitions) != 0: 
            register_df_copy = register_df.copy()   # do not change original register dataframe
            group_tuples = []
            if step in grouping_definitions:
                for group in grouping_definitions[step]:
                    # Get the grouped fallout for each group
                    grouped_fallout = 0
                    grouped_fallout_2 = 0
                    for group_string in grouping_definitions[step][group]:
                        for tup in register_df_copy.itertuples():
                            reg = tup[1]
                            fallout = tup[2]
                            fallout_2 = tup[3]
                            group_exp = re.compile(group_string)
                            if ( group_exp.search(reg) or group_string in reg ) and pd.notna(fallout): # regular expression match or substring match
                                grouped_fallout = grouped_fallout + fallout
                                if pd.notna(fallout_2):
                                    grouped_fallout_2 = grouped_fallout_2 + fallout_2
                                register_df_copy = register_df_copy[register_df_copy.REGISTER != reg] # remove this register from the copy of the DF so it cant be double-counted
                    group_tuples.append((group, grouped_fallout, grouped_fallout_2))
            # Create a separate DataFrame to store this week's data by groups
            group_df = pd.DataFrame.from_records(group_tuples, columns=["GROUP", cols[i+1], cols[i]]) # Dataframe with GROUP, THIS WEEK FALLOUT and LAST WEEK FALLOUT
            if len(group_user_columns) != 0:
                j=-1
                for col in user_columns:
                    j=j+1
                    col_dict = {} # Dictionary with GROUP as key and value for just this column instead of list of values in group_user_columns.  Used for easy key:value mapping on the DataFrame
                    for key in group_user_columns:
                        value = group_user_columns[key]
                        if (key.find("*") != -1): #  Check for Special Step
                            special_step = key.split("*",1)
                            key = special_step[0]
                            this_step = special_step[1]
                        else:
                            this_step = "all"
                        if (key.find("$") != -1): # Check for Special Config
                            special_conf = key.split("$",1)
                            key = special_conf[0]
                            conf = special_conf[1]
                        else:
                            conf = "all"                                                      
                        if ((str(conf) == configs) or (str(conf) == ("all"))) and ((str(this_step) == step) or (str(this_step) == ("all"))):
                            col_dict.update( { key : value[j] } )
                            group_df[col] = group_df['GROUP'].map(col_dict)
                            group_df[col] = [str(x) for x in group_df[col]]
                        else:
                            pass

            # Add Step and Config Columns to Group DataFrame
            group_df["STEP"] = step
            group_df["CONFIG"] = configs
            
            # Append Group Dataframe to the Overall Pareto Dataframe when configs match 
            for conf in config:
                if conf == configs:
                    overall_paretos[conf] = overall_paretos[conf].append(group_df[group_df["CONFIG"] == conf])

            # Creating Yield Dictionary
            if df.empty:
                print("DF is empty")
                yield_dic.update( {sheets[frame_index] : 0} )
            else:
                y = register_df[cols[i+1]]
                if y.shape[0] > 1:
                    yield_dic.update( {sheets[frame_index] : y[1]} )
                else:
                    yield_dic.update( {sheets[frame_index] : 0} )

            group_df.sort_values(by=cols[i+1], inplace=True, ascending=False) # Sort Group Dataframe by This Week's fallout
            idx = df.index[df[cols[0]] == "YIELD"]
            if register_df.shape[1] > 2 and register_df[cols[i+1]].loc[idx].shape[0] > 0:
                accounted_group_yield = register_df[cols[i+1]].loc[idx].iloc[0] + group_df.iloc[:, 1].sum(skipna = True) # Yield value + Total Fallout
                accounted_reg_yield   = register_df[cols[i+1]].loc[idx].iloc[0] + register_df.iloc[2:, 1].sum(skipna = True) # Yield value + Total Fallout of all registers
            if register_df.shape[1] > 1 and register_df[cols[i]].loc[idx].shape[0] > 0:
                accounted_group_yield_lastweek = register_df[cols[i]].loc[idx].iloc[0] + group_df.iloc[:, 2].sum(skipna = True) # Yield Value + Total Fallout of Last Week
                accounted_reg_yield_lastweek   = register_df[cols[i]].loc[idx].iloc[0] + register_df.iloc[2:, 2].sum(skipna = True) # Yield value + Total Fallout of all registers of last week
            filtered_columns = ['GROUP']
            filtered_columns.append(cols[i+1])
            filtered_columns.append(cols[i])
            if len(group_user_columns) != 0:
                for col in user_columns: # Append User Defined Columns to the Group DataFrame
                    filtered_columns.append(col)
            group_df = group_df[filtered_columns]

        # Sort this week's values for Register DataFrame
        register_df.sort_values(by=cols[i+1], inplace=True, ascending=False)
        worksheet = workbook.add_worksheet(sheets[frame_index]) # Adds a new excel worksheet for the pareto associated with this step
        reg_total = register_df.iloc[2:, 1].sum(skipna = True) # Get total Fallout of all registers of this Week
        reg_total_last = register_df.iloc[2:, 2].sum(skipna = True) # Get total Fallout of all registers of last Week
        accounted_group_yield = 0
        accounted_reg_yield = 0
        accounted_group_yield_lastweek = 0
        accounted_reg_yield_lastweek = 0 

        # Apply formatting
        border_format = workbook.add_format({'border':1})
        worksheet.set_column(0,16,20,border_format)

        bold_format = workbook.add_format({'bold':True,'border':'1'})
        worksheet.set_row(0,15,bold_format)
       
        # Print DataFrames to an xlsx document
        if (accounted_reg_yield != 'nan') and (accounted_reg_yield != '') and (accounted_reg_yield != 'NaN') :
            write_to_xlsx(register_df,user_columns, worksheet,0,accounted_reg_yield,accounted_reg_yield_lastweek)
        else:
            write_to_xlsx(register_df,user_columns, worksheet,0)
        if len(grouping_definitions) != 0:
            if accounted_group_yield != 'nan' and accounted_group_yield_lastweek != 'nan': 
                write_to_xlsx(group_df,user_columns, worksheet,len(list(register_df)) + 3, accounted_group_yield, accounted_group_yield_lastweek)
            else:
                write_to_xlsx(group_df,user_columns, worksheet,len(list(register_df)) + 3)
    
    if len(grouping_definitions) != 0:
        overall_pareto_index = -1
        for conf in config:
            for pack in package_type:
                overall_pareto_index += 1
                # Calculate Weighted % of ELC STEPS
                weight_pgsrt = (yield_dic['BURN_' + conf + '_' + pack]/100) * (yield_dic['HSRT_' + conf + '_' + pack]/100) * (yield_dic['CFIN_' + conf + '_' + pack]/100)
                weight_burn = (yield_dic['HSRT_' + conf + '_' + pack]/100) * (yield_dic['CFIN_' + conf + '_' + pack]/100)
                weight_hsrt = (yield_dic['CFIN_' + conf + '_' + pack]/100)
                weight_cfin = 1
                # Update Overall Pareto with Step, GROUP, Weighted % this Week and Weighted % Last Week
                temp = overall_paretos[conf]
                temp['WEIGHTED'] = temp.apply (lambda row : get_weighted_percent(row, weight_pgsrt, weight_burn, weight_hsrt, weight_cfin), axis=1)
                temp['WEIGHTED_LAST_WW'] = temp.apply(lambda row : get_weighted_percent_last_week(row, weight_pgsrt, weight_burn, weight_hsrt, weight_cfin), axis=1)
                temp.sort_values(by='WEIGHTED', inplace=True, ascending=False)
                overall_paretos[conf] = temp
                # Filter and reorganize columns
                filtered_columns = ['STEP' , 'GROUP' , 'WEIGHTED' , 'WEIGHTED_LAST_WW']
                if len(group_user_columns) != 0:
                    for col in user_columns:
                        filtered_columns.append(col)
                overall_paretos[conf] = overall_paretos[conf][filtered_columns]
                worksheet = workbook.add_worksheet("Overall_pareto" + conf) # Adds a new excel worksheet for the pareto associated with this step

                # Apply formatting
                border_format = workbook.add_format({'border':1})
                worksheet.set_column(0,9,20,border_format)

                bold_format = workbook.add_format({'bold':True,'border':'1'})
                worksheet.set_row(0,15,bold_format)


                write_to_xlsx(overall_paretos[conf],user_columns, worksheet,0)  
    workbook.close()

# Strip Helper Function
def strip(string):
    string = string.strip(' ')
    string = string.strip(',')
    string = string.strip('$')
    return string
    
#########################################
# Main function
#########################################

# Handles program options
# Guides program flow
def main():  
    # Passing DesignID and fab_number as arguments
    config = []
    package_type = ['SDP'] # Default SDP if no -package_type is passed
    WTD_mode = False
    designID_passed = False
    fab_passed = False
    configs_passed = False
    tww_passed = False
    grouping_files = ['define_groups.txt']
    if ("help" in str(sys.argv)):
        print(' \n\
This program generates the yield perato for PGSRT, BURN, HSRT, CFIN and then combines into an Overall Pareto based on weighted %.\n\
\n\
Each pareto is printed to its own sheet in an xlsx document with naming based on ELC step and config.\n\
\n\
MANDATORY ARGUMENTS\n\
\n\
-designID   e.g. -designID=y32a             Select design ID to pull data for\n\
\n\
-fab        e.g. -fab=15                    Select fab #\n\
\n\
-config     e.g. -config=x8,x16             Options are: [x4,x8,x16,combo]. IMPORTANT: this is based on the directory your data is stored in in the OPS Database, NOT necessarily the\n\
                                            config of your part.\n\
                                            If your data is not broken out by config in the OPS database then this option should be \'combo\', which means all data in one directory.\n\
                                            For info on how OPS has your data broken out, explore the OPS data directory for your part at: /vol/pye/MTI/OPS/.\n\
\n\
-package    e.g. -package=SDP,DDP           Options are: [SDP,DDP]. Default is SDP.\n\
\n\
-tww        e.g. -tww=202125                Select test workweek to pull data for, currently can only have tww and tww-1 data.\n\
\n\
OPTIONAL ARGUMENTS\n\
\n\
-grouping_files                             e.g. -grouping_files=ate_groupings.txt,burn_groupings.txt\n\
                                            List of text files to pull definitions for grouping registers together. See User-defined groupings in OPTIONAL TEXTFILES for more info.\n\
+WTD                                        Runs the Program in "Week-to-Date" mode. Pulls the latest OPS week reported - even if the work week has not completed\n\
\n\
OPTIONAL TEXTFILES\n\
\n\
Writing Comments                            Program ignores blank lines for all the textfiles used for this program(define_groups.txt, comments.txt and reg_comments.txt) \n\
                                            Program ignores lines starting with (#) so user can add comments to all the textfiles. \n\
\n\
User-defined groupings:                     You can group registers together under a group name in a grouping textfile as follows: \n\
                                            group name, str1, str2, str3...  or  group name, regex... The group file can match using substring matching or regular expression matching\n\
                                            You can convert a com file into a grouping file by removing encapsulating -samereg=//, replace all \',\' tokens with \'\\n\', replace all \'=\'\n\
                                            tokens with \',\' tokens, and adding a comma-separated list of applicable steps to the first line of the file.\n\
                                            All fallout from registers who\'s names contain str1|str2|str3... will be attributed to <group name>\n\
                                            Each register can only have fallout attributed to one group. If multiple registers contain the same string, fallout will be attributed for \n\
                                            the grouping defined on the earlier line.\n\
                                            The first line of a grouping file must be a comma-separated list of steps, which you would like to apply that file\'s groupings to\n\
                                            e.g. PGSRT,HSRT,CFIN or BURN\n\
                                            Grouping files can be passed into the program using the program using the -grouping_files optional argument\n\
\n\
User-defined columns                        You can add comments, owners, ETAs and more by defining them in a file within the same directory as this program called comments.txt\n\
by groups                                   example: group name^comments^owner^ETA \n\
                                            In the above example, comments, owners, and ETAs are mapped using the register name to group pareto items\n\
                                            In order to avoid confliting with commas added in your comments or other fields, this text file is delimited by the carat (\'^\') token.\n\
                                            In the above example, comments, owners, and ETAs are mapped using the group name to group pareto items\n\
                                            The first line of your file must be a list defining your columns, separated by the \'^\' token.\n\
                                            These definitions tell the program what heading to apply to each column as well as how many columns you have defined\n\
\n\
User-defined columns                        You can add comments, owners, ETAs and more by defining them in a file within the same directory as this program called regcomments.txt\n\
by register                                 example: register name^comments^owner^ETA \n\
                                            In the above example, comments, owners, and ETAs are mapped using the register name to group pareto items\n\
                                            In order to avoid confliting with commas added in your comments or other fields, this text file is delimited by the carat (\'^\') token.\n\
                                            In the above example, comments, owners, and ETAs are mapped using the group name to group pareto items\n\
                                            The first line of your file must be a list defining your columns, separated by the \'^\' token.\n\
                                            These definitions tell the program what heading to apply to each column as well as how many columns you have defined\n\
\n\
Special comments for STEP/config            You can write different comments, owners, ETAs for same GROUP/REGSITER according to the config(using $ symbol) and STEP(using * symbol) \n\
by config for GROUP                         example: group name$x8^comments^owner^ETA will only output to x8 Groups\n\
by config for REGSITER                      example: register name$x16^comments^owner^ETA will only output to x16 Registers\n\
by STEP for GROUP                           example: group name*CFIN^comments^owner^ETA will only output to CFIN Groups\n\
by STEP for REGSITER                        example: register name*HSRT^comments^owner^ETA will only output to HSRT Registers\n\
by config and STEP for GROUP                example: group name$x8*PGSRT^comments^owner^ETA will only output to x8 PGSRT Groups\n\
by config and STEP for REGISTER             example: group name$x16*BURN^comments^owner^ETA will only output to x16 BURN Registers.')


        sys.exit()
    else:
        for argument in sys.argv:
            if "designID" in argument: # For Product DesignID
                designID_passed = True
                tokens = argument.split("=")
                designID = tokens[1]
            if "fab" in argument: # For Fab Number
                fab_passed = True
                tokens = argument.split("=")
                fabNU = tokens[1]
            if "config" in argument: # For configs chosen [x4,x8,x16,combo]
                configs_passed = True
                tokens = argument.split("=")
                config_str = tokens[1]
                config = config_str.split(",")
            if "package" in argument: # For Package Type
                tokens = argument.split("=")
                package_str = tokens[1]
                package_type = package_str.split(",")
            if "grouping_files" in argument:
                tokens = argument.split("=")
                grouping_file_str = tokens[1]
                grouping_files = grouping_file_str.split(",")
            if "tww" in argument:
                tww_passed = True
                tokens = argument.split("=")
                tww = tokens[1]
            if "+WTD" in argument: # Run in WTD mode
                WTD_mode = True
    if not designID_passed:
        ID = input("designID: ") # Get DesignID from user if not passed
        designID = str(ID)
    if not fab_passed:
        fabNU = input("fab: ") # Get Fab Number from user if not passed
    if not configs_passed:
        config = []
        config_str = input("config: ") # Get configs from user if not passed
        config = config_str.split(",")
    if not tww_passed:
        tww = input("tww(currently can only have tww and tww-1 data): ") # Get tww from user if not passed

    # Get grouping definitions
    for grouping_file in grouping_files:
        try:
            file = open(grouping_file, 'r')
            step_grouping_definitions = {}
            definition_line = True
            for line in file:
                line = line.strip('\n') # strip newline characters
                line = line.strip(' ') # strip whitespace
                if definition_line:
                    definition_line = False
                    if "," in line:
                        applicable_steps = line.split(",")
                        continue
                    else: 
                        applicable_steps = ['PGSRT', 'BURN', 'HSRT', 'CFIN'] # If this line is missing from the top of the file then these groupings will be applied to all steps
                if (line != '') and (line.startswith("#") == False):
                    tokens = line.split(",",1) # returns 2 tokens: the group name and a comma-separated list of definitions for that group
                    group = tokens[0]
                    definitions = tokens[1]
                    step_grouping_definitions[group] = definitions.split(",") # Adds a dictionary entry for definitions keyed to the group
            for step in applicable_steps:
                grouping_definitions[step] = step_grouping_definitions # Adds dictionary from that file as a value keyed to the steps it was meant to group
            file.close()
        except FileNotFoundError:
            pass

    # Create dictionary for user defined columns for GROUP dataframe
    try:
        file = open(comment, 'r')
        definition_line = True
        for line in file:
            line = line.strip('\n') # strip newline characters
            line = line.strip(' ') # strip whitespace
            if definition_line:
                user_columns = line.split("^")
                definition_line = False
            else:  
                if (line != '') and (line.startswith("#") == False):
                    # Create a dictionary where the key is the GROUP and the value is a list of user-defined column values, eg: {GROUP : [comment, owner, eta, jira]}
                    user_vals = []
                    tokens = line.split(split_token)
                    group = strip(tokens[0])
                    i=0
                    for col in user_columns:
                        i = i+1
                        user_vals.append(strip(tokens[i]))
                        if (line.find("$") != -1):
                            special_conf = line.split("$",1)        # Return the value of the config  
                            conf = special_conf[1]
                        if (line.find("*") != -1):
                            special_step = line.split("*",1)        # Return the value of the config  
                            step = special_step[1]
                    group_user_columns.update( {group : user_vals} )
        file.close()
    except FileNotFoundError:
        pass

    # Create a dictionary for user defined columns for REGISTER dataframe
    try:
        file = open(comment_2, 'r')
        definition_line = True
        for line in file:
            line = line.strip('\n') # strip newline characters
            line = line.strip(' ') # strip whitespace
            if definition_line:
                user_columns = line.split("^")
                definition_line = False
            else:
                if (line != '') and (line.startswith("#") == False):
                    # Create a dictionary where the key is the REGISTER and the value is a list of user-defined column values, eg: {REGISTER : [comment,owner,eta,jira]}
                    user_vals = []
                    tokens = line.split(split_token)
                    group = strip(tokens[0])
                    i = 0
                    for col in user_columns:
                        i = i + 1
                        user_vals.append(strip(tokens[i]))
                        if (line.find("$") != -1):
                            special_conf = line.split("$",1)        # Return the value of the config  
                            conf = special_conf[1]
                        if (line.find("*") != -1):
                            special_step = line.split("*",1)        # Return the value of the config  
                            step = special_step[1]
                    register_user_columns.update( {group : user_vals})
        file.close
    except FileNotFoundError:
        pass
                
    extract_ops_data(designID,fabNU,config,package_type,tww)
    data_frames = parse_csv_files()
    generate_xlsx_document(data_frames,sheets,config,user_columns,package_type, WTD_mode,tww)

# Call the main function
if __name__=="__main__":
    main()
