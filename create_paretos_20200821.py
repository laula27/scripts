########################################################################################################
# Author:           ryancaraway / mhassana
# Date updated:     7-22-2020
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
# Imports
# Global variables
# 1. Extract OPS Yield data and put it into csv files broken out by ELC step
# 2. Parse the csv data files into a list of dataframes
# 3. Perform dataframe manipulations to get sorted registers and groups by fallout and
#    write to xlsx file
# main function: handles arguments and program flow
# 
# Future Plans
# TODO: Get OPS database path from a file in local directory, this might not be the same for all designs
# TODO: Use this script to pull the overall yield information for the Master Fab page
# TODO: Add regular expression matching for grouping registers
# TODO: Add help and optional flags for:
#           - not re-extracting data
#           - design ID
#           - config
#           - grouping optional if filename passed in
# TODO: Pareto creates and manages its own csv file directory
########################################################################################################
import sys              # Get program arguments
import os               # Enter system commands
import pandas as pd     # Hold and allow manipulations on yield data
import csv              # csv file i/o
import xlsxwriter       # Write and format xlsx files
import re               # Regular expression pattern matching

# Global Variables
num_ww = 3  # number of work weeks to pull OPS data for. Currently only support 3 weeks: "Last week", "This week", and "WTD"
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
def extract_ops_data(designID,fabNU,config,package_type):
    if not os.path.exists(file_directory): # Make directory for csv files if it doesn't already exist
        os.makedirs(file_directory)
    os.system('rm ' + output_workbook_name) # remove old workbook
    os.system('rm ' + file_directory + "*.csv") # remove old csv data files
    
    # Pull yield data from the OPS report database for each step, config and package type
    for step in elc_steps:
        for conf in config:
            for pack in package_type:
                if (conf == 'x8') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|X8/SDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x16') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|X16/SDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'combo') and (pack == 'SDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|SDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x8') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|X8/DDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'x16') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|X16/DDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')
                elif (conf == 'combo') and (pack == 'DDP'):
                    os.system('/u/sray/bin/OPSext -dbase='+designID+' -newops -stepreg="'+step+'|DDP||ALL" -'+str(num_ww)+' -fab='+fabNU+' | tpose  >| '+file_directory + step +'_'+conf+'_'+pack+'.csv')

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

###########################################
# Helper function of generate_xlsx_document
###########################################

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
# Outer loop iterates over each DataFrame in the list to iterate over each step/config
# Inner loop iterates over the columns of the dataframe to look at the work-weeks individually
def generate_xlsx_document(data_frames,sheets,config,user_columns,package_type):
    overall_paretos = {}
    for conf in config:
        overall_paretos.update( {conf : pd.DataFrame()} ) # Update Overall Pareto with configs
    yield_dic = {}
    workbook = xlsxwriter.Workbook(output_workbook_name) 
    # Write each data frame to a worksheet for its associated step
    weeks = ['skip_me','','This_Week','WTD']
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
            for group in grouping_definitions:
                # Get the grouped fallout for each group
                grouped_fallout = 0
                grouped_fallout_2 = 0
                for group_string in grouping_definitions[group]:
                    for tup in register_df_copy.itertuples():
                        reg = tup[1]
                        fallout = tup[2]
                        fallout_2 = tup[3]
                        group_exp = re.compile(group_string)
                        if group_exp.match(reg) and pd.notna(fallout):
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
                y = register_df[cols[2]]
                if y.shape[0] > 1:
                    yield_dic.update( {sheets[frame_index] : y[1]} )
                else:
                    yield_dic.update( {sheets[frame_index] : 0} )

            group_df.sort_values(by=cols[i+1], inplace=True, ascending=False) # Sort Group Dataframe by This Week's fallout
            group_total = group_df.iloc[:, 1].sum(skipna = True) # Get total Fallout for this week
            group_total_last = group_df.iloc[:, 2].sum(skipna = True) # Get total Fallout for last week
            y = register_df[cols[2]] # Get this week column
            z = register_df[cols[1]] # Get last week column
            accounted_group_yield = y[1] + group_total # Yield value + Total Fallout
            accounted_group_yield_lastweek = z[1] + group_total_last # Yield Value + Total Fallout of Last Week
            filtered_columns = ['GROUP']
            filtered_columns.append(cols[2])
            filtered_columns.append(cols[1])
            if len(group_user_columns) != 0:
                for col in user_columns: # Append User Defined Columns to the Group DataFrame
                    filtered_columns.append(col)
            group_df = group_df[filtered_columns]


        # Sort this week's values for Register DataFrame
        register_df.sort_values(by=cols[i+1], inplace=True, ascending=False)
        worksheet = workbook.add_worksheet(sheets[frame_index]) # Adds a new excel worksheet for the pareto associated with this step
        reg_total = register_df.iloc[2:, 1].sum(skipna = True) # Get total Fallout of all registers of this Week
        reg_total_last = register_df.iloc[2:, 2].sum(skipna = True) # Get total Fallout of all registers of last Week
        accounted_reg_yield = y[1] + reg_total # Yield value + Total Fallout of all registers
        accounted_reg_yield_lastweek = z[1] + reg_total_last # Yield value + Total Fallout of all registers of last week 

        # Apply formatting
        border_format = workbook.add_format({'border':1})
        worksheet.set_column(0,15,20,border_format)

        bold_format = workbook.add_format({'bold':True,'border':'1'})
        worksheet.set_row(0,15,bold_format)

        # Print DataFrames to an xlsx document
        if (accounted_reg_yield != 'nan') and (accounted_reg_yield != '') and (accounted_reg_yield != 'NaN') :
            write_to_xlsx(register_df,user_columns, worksheet,0,accounted_reg_yield,accounted_reg_yield_lastweek)
        else:
            write_to_xlsx(register_df,user_columns, worksheet,0)
        if len(grouping_definitions) != 0:
            if accounted_group_yield != 'nan': 
                write_to_xlsx(group_df,user_columns, worksheet,len(list(register_df)) + 3, accounted_group_yield, accounted_group_yield_lastweek)
            else:
                write_to_xlsx(group_df,user_columns, worksheet,len(list(register_df)) + 3)
        i += 1
    
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
                worksheet.set_column(0,8,20,border_format)

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
    user_columns = ['COMMENTS', 'OWNER', 'ETA'] # default if no -columns arg is passed
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
-config     e.g. -config=x8,x16             Options are: [x4,x8,x16,combo]. IMPORTANT: this is based on the directory your data is stored in in the OPS Database, NOT necessarily the config of your part.\n\
                                            If your data is not broken out by config in the OPS database then this option should be \'combo\', which means all data in one directory.\n\
                                            For info on how OPS has your data broken out, explore the OPS data directory for your part at: /vol/pye/MTI/OPS/.\n\
\n\
OPTIONAL ARGUMENTS\n\
\n\
-columns    e.g. -columns="comments","jira" You can define columns that you want added and tracked along with the pareto items with things like comments, owners, jira tickets, eta, or anything you would like to define!\n\
                                            If you use this argument, you must define your register/group pareto item to column mappings in the comments.txt or regcomments.txt files.\n\
                                            The number of items you define must match the number of columns you add in the text file\n\
                                            If no -columns is passed but you did include an optional group or register text files, then the default is columns = [\'COMMENTS\',\'OWNERS\', \'ETA\'] \n\
\n\
OPTIONAL TEXTFILES\n\
\n\
User-defined groupings:                     You can group registers together to a group name in a textfile \n\
                                            Define groups in a define_groups.txt as follows: \n\
                                            group name, str1, str2, str3... \n\
                                            All fallout from registers who\'s names contain str1|str2|str3... will be attributed to <group name>\n\
                                            Each register can only have fallout attributed to one group. If multiple registers contain the same string, fallout will be attributed for the grouping defined on the earlier line.\n\
\n\
User-defined columns                        You can add comments, owners, ETAs and more by defining them in a file within the same directory as this program called comments.txt\n\
by groups                                   example: group name^comments^owner^ETA \n\
                                            In the above example, comments, owners, and ETAs are mapped using the group name to group pareto items\n\
                                            You can define your own columns using the -columns argument, though the number of columns you specify must match the number of mappings included in your text file\n\
                                            If the -columns optional argument is not used, the default columns expected in this text file are: name, comments, and eta respectively.\n\
                                            In order to avoid confliting with commas added in your comments or other fields, this text file is delimited by the carat (\'^\') token.\n\
\n\
User-defined columns                        You can add comments, owners, ETAs and more by defining them in a file within the same directory as this program called regcomments.txt\n\
by register                                 example: register name^comments^owner^ETA \n\
                                            In the above example, comments, owners, and ETAs are mapped using the register name to group pareto items\n\
                                            You can define your own columns using the -columns argument, though the number of columns you specify must match the number of mappings included in your text file\n\
                                            In order to avoid confliting with commas added in your comments or other fields, this text file is delimited by the carat (\'^\') token.\n\
\n\
Textfiles                                   Program ignores blank lines for all the textfiles used for this program(define_groups.txt, comments.txt and reg_comments.txt) \n\
Writing Comments                            Program ignores lines starting with (#) so user can add comments to all the textfiles. \n\
\n\
Special comments for STEP/config            You can write different comments, owners, ETAs for same GROUP/REGSITER according to the config(using $ symbol) and STEP(using * symbol) \n\
by config for GROUP                         example: group name$x8^comments^owner^ETA will only output to x8 Groups\n\
by config for REGSITER                      example: register name$x16^comments^owner^ETA will only output to x16 Registers\n\
by STEP for GROUP                           example: group name*CFIN^comments^owner^ETA will only output to CFIN Groups\n\
by STEP for REGSITER                        example: register name*HSRT^comments^owner^ETA will only output to HSRT Registers\n\
by config and STEP for GROUP                example: group name$x8*PGSRT^comments^owner^ETA will only output to x8 PGSRT Groups\n\
by config and STEP for REGISTER             example: group name$x16*BURN^comments^owner^ETA will only output to x16 BURN Registers.')


        sys.exit()
    elif (len(sys.argv) >= 4):
        for argument in sys.argv:
            if "designID" in argument: # For Product DesignID
                tokens = argument.split("=")
                designID = tokens[1]
            if "fab" in argument: # For Fab Number
                tokens = argument.split("=")
                fabNU = tokens[1]
            if "config" in argument: # For configs chosen [x4,x8,x16,combo]
                tokens = argument.split("=")
                config_str = tokens[1]
                config = config_str.split(",")
            if "columns" in argument: # For User Defined Columns
                tokens = argument.split("=")
                col_str = tokens[1]
                user_columns = col_str.split(",")
            if "package" in argument: # For Package Type
                tokens = argument.split("=")
                package_str = tokens[1]
                package_type = package_str.split(",")
    else:
        ID = input("designID: ") # Get DesignID from user if not passed
        designID = str(ID)
        fabNU = input("fab: ") # Get Fab Number from user if not passed
        config = []
        config_str = input("config: ") # Get configs from user if not passed
        config = config_str.split(",")
        package_type = ['SDP'] # default if no -package_type arg is passed
        user_columns = ['COMMENTS', 'OWNER', 'ETA'] # default if no -columns arg is passed        

    # Get grouping definitions
    try:
        file = open(grouping_filename, 'r')
        for line in file:
            line = line.strip('\n') # strip newline characters
            line = line.strip(' ') # strip whitespace
            if (line != '') and (line.startswith("#") == False):
                tokens = line.split(",",1) # returns 2 tokens: the group name and a comma-separated list of definitions for that group
                group = tokens[0]
                definitions = tokens[1]
                grouping_definitions[group] = definitions.split(",") # Adds a dictionary entry for definitions keyed to the group
        file.close()
    except FileNotFoundError:
        pass

    # Create dictionary for user defined columns for GROUP dataframe
    try:
        file = open(comment, 'r')
        for line in file:
            line = line.strip('\n') # strip newline characters
            line = line.strip(' ') # strip whitespace
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
        for line in file:
            line = line.strip('\n') # strip newline characters
            line = line.strip(' ') # strip whitespace
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
                
    extract_ops_data(designID,fabNU,config,package_type)
    data_frames = parse_csv_files()
    generate_xlsx_document(data_frames,sheets,config,user_columns,package_type)

# Call the main function
if __name__=="__main__":
    main()
