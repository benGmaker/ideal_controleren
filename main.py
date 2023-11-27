import pandas as pd
import os
import glob
from pick import pick
import numpy as np

def file_selection(path):
    #returns the selected excel files in the order [website, snelstart, onbekend]
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))  # reading out which excel files are in directory [strings]
    results = []
    titles = ['select website list','select snelstart list','select unkown user list']
    for i in range(0,3):
        option, index = pick(csv_files, titles[i])
        results.append(option)
        csv_files.pop(index)
    return results

WEBSITE_MATCH_NAME = "ID-nummer" #can be any of the column names from the excel file
def clean_website(df):
    #Finds the starting position of the table in the excel file such that it is impossible to paste it incorrectly into excel
    columns = df.columns #finding column names
    #TODO check toevoegen om te kijken of de colommen al kloppen, want anders maakt dit het stuk hieronder
    for i in columns:
        if WEBSITE_MATCH_NAME in df[i].values: #check if the value is in the column
            match = df.loc[df[i] == WEBSITE_MATCH_NAME].index #finding the index of the data
            match = match[0] #getting numerical value out
            continue #index is found
    column_names = df.iloc[match] #reading out the column names
    df = df.set_axis(column_names, axis=1) #setting column names
    df = df.iloc[match+1:] #dropping the useless rows
    df = df.reset_index() #resseting the index to start at 0 again
    df[WEBSITE_MATCH_NAME] = df[WEBSITE_MATCH_NAME].fillna(0) #removing n/a in the id column
    df = df.astype({WEBSITE_MATCH_NAME: 'int'}) #setting id column to integers
    return df

SNELSTART_COLUMN_NAME = 'KLANTRelatiecodeNaamPlaats'
SNELSTART_NEW_COLUMNS = [WEBSITE_MATCH_NAME,'naam','plaats']
def format_snelstart(df):
    #seperates the SNELSTART_COLUMN_NAME into SNELSTART_NEW_COLUMNS
    df_split = df[SNELSTART_COLUMN_NAME].str.split(',', expand=True)
    df_split = df_split.set_axis(SNELSTART_NEW_COLUMNS, axis = 1)
    df = df.join(df_split, how= 'right')
    df[WEBSITE_MATCH_NAME] = df[WEBSITE_MATCH_NAME].fillna(0) #removing n/a in the id column
    df = df.astype({WEBSITE_MATCH_NAME: 'int'}) #setting id column to integers
    #Reformats the snelstart dataframe
    return df

CORRECT_SNELSTART_WEBSITE_COLUMN_NAME = "SNELSTART-WEBSITE_CORRECT"
CORRRECT_NOT_PAID_NAME = "NOT_PAID_CORRECT"
AFGEREKEND = 'Afgerekend' #names of used columns
OMZETAANTAL = 'OmzetAantal'
def correct_website_snelstart(df):
    #
    df[OMZETAANTAL] = df[OMZETAANTAL].fillna(0) #filling empty cells with zero's to make boolean statement easier
    #finding the correct data and writing data
    df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] = np.where( (df[OMZETAANTAL] == 1) &  (df[AFGEREKEND] == 'Ja'),True,False) #finding correct inboekingen
    df[CORRRECT_NOT_PAID_NAME] = np.where( (df[OMZETAANTAL] == 0 ) & (df[AFGEREKEND] == 'Nee'),True,False) #finding correct not paid
    return df

if __name__ == '__main__':
    os_path = (os.getcwd())
    test_path = os_path + "\\testdata"
    #filepaths = file_selection(test_path) #selecting the files to be website, snelstart, unkown

    #TEST
    filepaths = [test_path + '\\web.xlsx', test_path + '\\asml.xlsx', test_path + '\\onbekend.xlsx']  #loading test files

    #reading in the data
    website = pd.read_excel(filepaths[0])
    snelstart = pd.read_excel(filepaths[1])
    onbekend = pd.read_excel(filepaths[2])

    #cleaning up the data
    website = clean_website(website) #removing possible clutter from the website
    snelstart = format_snelstart(snelstart)
    merged = pd.merge(website, snelstart,how="outer", on=['ID-nummer'])
    #todo onbekende klant data invoeg functie toevoegen
    merged.to_excel("RAW DATA.xlsx")

    ## Participant check functionality:
    ## Check for GOOD cases specifacally and assign a true or false accordingly
    ##      -> importantly these are very specific and simple correct cases
    ## If done for all the the true cases, only one case should be true for each inboeking
    ##      -> If all are false something is incorrect
    ##      -> If more than 2 are true something is incorrect -> 2 cases can't be correct at the same time
    snelstart_website_correct = correct_website_snelstart(merged) #checking between website and snelstart correct values
    snelstart_website_correct.to_excel('snelstartfout.xlsx') #writing out resulting data
    #todo onbekende klant controle toevoegen
    #   * onbekend - website match
    #   * onbekend - snelstart match
