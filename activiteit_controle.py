
import os
import pandas as pd
import glob
from pick import pick
import numpy as np

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from datetime import datetime

def _draw_as_table(df, pagesize):
    alternating_colors = [['white'] * len(df.columns), ['lightgray'] * len(df.columns)] * len(df)
    alternating_colors = alternating_colors[:len(df)]
    fig, ax = plt.subplots(figsize=pagesize)
    ax.axis('tight')
    ax.axis('off')
    the_table = ax.table(cellText=df.values,
                        rowLabels=df.index,
                        colLabels=df.columns,
                        rowColours=['lightblue']*len(df),
                        colColours=['lightblue']*len(df.columns),
                        cellColours=alternating_colors,
                        loc='center')
    return fig

def dataframe_to_pdf(df, filename, basename, numpages=(1, 1), pagesize=(11, 8.5)):
    with PdfPages(filename) as pdf:
        nh, nv = numpages
        rows_per_page = len(df) // nh
        cols_per_page = len(df.columns) // nv
        for i in range(0, nh):
            for j in range(0, nv):
                page = df.iloc[(i * rows_per_page):min((i + 1) * rows_per_page, len(df)),
                       (j * cols_per_page):min((j + 1) * cols_per_page, len(df.columns))]
                fig = _draw_as_table(page, pagesize)
                if nh > 1 or nv > 1:
                    # Add a part/page number at bottom-center of page
                    fig.text(0.5, 0.5 / pagesize[0],
                             "Part-{}x{}: Page-{}".format(i + 1, j + 1, i * nv + j + 1) + basename,
                             ha='center', fontsize=8)
                pdf.savefig(fig, bbox_inches='tight')

                plt.close()

KLANTENPER_PAGINA = 30
def create_pdf(df, filename, basename):
    n_pages = int(np.ceil(len(df)/KLANTENPER_PAGINA))
    numpages = (n_pages,1)
    pagesize = (20, 8.5)
    dataframe_to_pdf(df, filename, basename,numpages, pagesize)

NO_UNKOWN_USER_DATA = 'NO UNKOWN USER LIST'
def file_selection(path):
    # returns the selected excel files in the order [website, snelstart, onbekend]
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))  # reading out which excel files are in directory [strings]
    results = []
    titles = ['select website list', 'select snelstart list', 'select unkown user list']
    for i in range(0, 3):
        if i == 2: #add no data option
            csv_files.append(NO_UNKOWN_USER_DATA)
        option, index = pick(csv_files, titles[i])
        results.append(option)
        csv_files.pop(index)
    return results

WEBSITE_MATCH_NAME = "ID-nummer"  # can be any of the column names from the excel file

def clean_website(df):
    # Finds the starting position of the table in the excel file such that it is impossible to paste it incorrectly into excel
    columns = df.columns  # finding column names
    # TODO check toevoegen om te kijken of de colommen al kloppen, want anders maakt dit het stuk hieronder
    match = 0
    for i in columns:
        if WEBSITE_MATCH_NAME in df[i].values:  # check if the value is in the column
            match = df.loc[df[i] == WEBSITE_MATCH_NAME].index  # finding the index of the data
            match = match[0]  # getting numerical value out
            continue  # index is found
    column_names = df.iloc[match]  # reading out the column names
    df = df.set_axis(column_names, axis=1)  # setting column names
    df = df.iloc[match + 1:]  # dropping the useless rows
    df = df.reset_index()  # resseting the index to start at 0 again
    df[WEBSITE_MATCH_NAME] = df[WEBSITE_MATCH_NAME].fillna(0)  # removing n/a in the id column
    df = df.astype({WEBSITE_MATCH_NAME: 'int'})  # setting id column to integers
    return df

SNELSTART_COLUMN_NAME = 'KLANTRelatiecodeNaamPlaats'
SNELSTART_NEW_COLUMNS = [WEBSITE_MATCH_NAME, 'naam', 'plaats']

def format_snelstart(df):
    # seperates the SNELSTART_COLUMN_NAME into SNELSTART_NEW_COLUMNS
    df_split = df[SNELSTART_COLUMN_NAME].str.split(',', expand=True)
    df_split = df_split.set_axis(SNELSTART_NEW_COLUMNS, axis=1)
    df = df.join(df_split, how='right')
    df[WEBSITE_MATCH_NAME] = df[WEBSITE_MATCH_NAME].fillna(0)  # removing n/a in the id column
    df = df.astype({WEBSITE_MATCH_NAME: 'int'})  # setting id column to integers
    # Reformats the snelstart dataframe
    return df

ONBEKEND_NAAM = 'unkown customer name'
ALLIAS_ID_NUMMER = [WEBSITE_MATCH_NAME, 'ID', 'id', 'nummer', 'studenten', 'student-id', 'student', 'studenten nummer']
ALLIAS_NAAM = [ONBEKEND_NAAM, 'naam', 'klant', 'name', 'First name']

def clean_onbekend(df):
    n_id = 0;
    n_name = 0;
    for i in df.columns:
        if (i in ALLIAS_ID_NUMMER):
            df = df.rename(columns={i: WEBSITE_MATCH_NAME})  # replacing name of the collumn with one name
            n_id += 1  # number of found id column is plus one
        if (i in ALLIAS_NAAM):
            df = df.rename(columns={i: ONBEKEND_NAAM})  # replacing name of the collumn with one name
            n_name += 1  # number of found name columns is plus one

    # throwing exceptions if somethings is wrong with the data
    if n_id == 0:
        raise TypeError('No ID number columns could be found, please use one of the following '
                        'alliases in the dataset of unkown customers: \n' + str(ALLIAS_ID_NUMMER))
    elif n_id > 1:
        raise TypeError('Too many ID number columns have been found please be carefull using the following '
                        'alliases in the dataset of unkown customers: \n ' + str(ALLIAS_ID_NUMMER))
    if n_name == 0:
        raise TypeError("No customer names columns could be found, please use one of the following alliases "
                        "in the dataset of unkown customers: \n " + str(ALLIAS_NAAM))
    if n_name > 1:
        raise TypeError("Too customer names columns could be found, please be carefull using the following alliases "
                        "in the dataset of unkown customers: \n " + str(ALLIAS_NAAM))

    return df

CORRECT_SNELSTART_WEBSITE_COLUMN_NAME = "SNELSTART-WEBSITE_CORRECT"
CORRRECT_NOT_PAID_NAME = "NOT_PAID_CORRECT"
AFGEREKEND = 'Afgerekend'  # names of used columns
OMZETAANTAL = 'OmzetAantal'

def correct_website_snelstart(df):
    #
    df[OMZETAANTAL] = df[OMZETAANTAL].fillna(0)  # filling empty cells with zero's to make boolean statement easier
    # finding the correct data and writing data
    df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] = np.where((df[OMZETAANTAL] == 1) & (df[AFGEREKEND] == 'Ja'), 1,
                                                         0)  # finding correct inboekingen
    df[CORRRECT_NOT_PAID_NAME] = np.where((df[OMZETAANTAL] == 0) & (df[AFGEREKEND] == 'Nee'), 1,
                                          0)  # finding correct not paid
    return df

CORRECT_SNELSTART_ONBEKEND = 'ONBEKEND-SNELSTART_MATCH'  #
CORRECT_WEBSITE_ONBEKEND ='ONBEKEND-WEBSITE_MATCH'

def correct_onbekend(df):
    df[CORRECT_SNELSTART_ONBEKEND] = np.where((df[OMZETAANTAL] == 1) & (df[ONBEKEND_NAAM].notnull()), 1, 0)
    df[CORRECT_WEBSITE_ONBEKEND] = np.where((df[AFGEREKEND] == 'Ja') & (df[ONBEKEND_NAAM].notnull()), 1, 0)
    return df

def correct_result(df):
    # Participant check functionality:
    # Check for GOOD cases specifacally and assign a true or false accordingly
    #      -> importantly these are very specific and simple correct cases
    # If done for all the the true cases, only one case should be true for each inboeking
    #      -> If all are false something is incorrect
    #      -> If more than 2 are true something is incorrect -> 2 cases can't be correct at the same time
    df['sum'] = df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] + df[CORRRECT_NOT_PAID_NAME] + df[CORRECT_WEBSITE_ONBEKEND]
    df['result'] = df['sum'] == 1
    return df

WEBSITE_COLUMN_SELECTION = ['ID-nummer','First name','Last name','Afgerekend']
SNELSTART_COLUMN_SELECTION = ['OmzetAantal','OmzetBedragExclusiefBtw','naam']
ONBEKEND_COLUMN_SELECTION = ['unkown customer name']
COLUMN_SELECTION = []
def data_selection(df):
    #data_selection select only the necessary data needed for review

    #creating data selection
    COLUMN_SELECTION = []
    COLUMN_SELECTION.extend(WEBSITE_COLUMN_SELECTION)
    COLUMN_SELECTION.extend(SNELSTART_COLUMN_SELECTION)
    COLUMN_SELECTION.extend(ONBEKEND_COLUMN_SELECTION)
    COLUMN_SELECTION.extend(['result'])

    #selecting data
    df_small = pd.DataFrame() #Creating new dataframe
    df_small[COLUMN_SELECTION] = df[COLUMN_SELECTION] #selecting desired columns

    #sorting data
    df_small = df_small.sort_values(by=['result'],ascending=True) #sorting by validity result

    return df_small


if __name__ == '__main__':
    os_path = (os.getcwd())
    test_path = os_path + "\\testdata"
    filepaths = file_selection(test_path) #selecting the files to be website, snelstart, unkown

    # TEST FILES PATHS
    # filepaths = [test_path + '\\web.xlsx', test_path + '\\asml.xlsx', test_path + '\\onbekend.xlsx']  #loading test files
    #filepaths = [test_path + '\\v2\\website.xlsx', test_path + '\\v2\\snelstart.xlsx',test_path + '\\v2\\onbekend.xlsx']  # loading test files
    #filepaths = [test_path + '\\test3\\website.xlsx', test_path + '\\test3\\snelstart.xlsx',test_path + '\\test3\\onbekend.xlsx']  # loading test files
    # reading in the data
    print('reading data')
    website = pd.read_excel(filepaths[0])
    snelstart = pd.read_excel(filepaths[1])
    if filepaths[2] == NO_UNKOWN_USER_DATA:
        onbekend = pd.DataFrame()
    else:
        onbekend = pd.read_excel(filepaths[2])

    # cleaning up the data
    print('Cleaning data')
    website = clean_website(website)  # removing possible clutter from the website
    snelstart = format_snelstart(snelstart)
    if filepaths[2] != NO_UNKOWN_USER_DATA:
        onbekend = clean_onbekend(onbekend) #only adding unkown user data if it is there
    else:
        onbekend[WEBSITE_MATCH_NAME] = [int(66)] #creating empty dataframe with column name that will be refrenced
        onbekend[ONBEKEND_NAAM] = ['No unkown users given']  # creating empty dataframe with column name that will be refrenced

    #merging data into 1 dataframe
    print('Performing analysis')
    merged = pd.merge(website, snelstart, how="outer", on=['ID-nummer'])
    merged = pd.merge(merged, onbekend, how="outer", on=['ID-nummer'])

    #performing checks on the data
    final_data = correct_website_snelstart(merged)  # checking between website and snelstart correct values
    final_data = correct_onbekend(final_data) #checking matches between onbekend and snelstart + website
    final_data = correct_result(final_data) #doing the final check if all results are correct

    #processing data and printing results
    print('Saving results')

    #creating a file name
    now = datetime.now()#adding date and time information to name
    dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
    date_time = 'made on (' + dt_string + ')'

    #adding activity information
    basis_name = str(round(snelstart['Artikelcode'].iloc[1])) + ' ' + snelstart['Omschrijving'].iloc[1]
    basis_name = basis_name.replace(':','')

    #adding date information and removing ilegal characters
    basis_name = 'made on (' + dt_string + ')' + basis_name

    final_data.to_excel("ALLDATA " + basis_name +".xlsx")  #exporting all data if analysis is desired

    print('creating pdf')
    df_small = data_selection(final_data) #selecting only necesarry information for export

    create_pdf(df_small[df_small['result'] == False], '!ERRORS ' + basis_name + '.pdf', basis_name) #exporting errors
    create_pdf(onbekend, "unkown customers " + basis_name + '.pdf', basis_name) #exporting all unkown users used
    create_pdf(df_small, "result " + basis_name + '.pdf', basis_name)  #exporting all the final data