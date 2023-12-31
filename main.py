import pandas as pd
import os
import glob
from pick import pick
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def _draw_as_table(df, pagesize):
    alternating_colors = [['white'] * len(df.columns), ['lightgray'] * len(df.columns)] * len(df)
    alternating_colors = alternating_colors[:len(df)]
    fig, ax = plt.subplots(figsize=pagesize)
    ax.axis('tight')
    ax.axis('off')
    the_table = ax.table(cellText=df.values,
                         rowLabels=df.index,
                         colLabels=df.columns,
                         rowColours=['lightblue'] * len(df),
                         colColours=['lightblue'] * len(df.columns),
                         cellColours=alternating_colors,
                         loc='center')
    return fig


def dataframe_to_pdf(df, filename, numpages=(1, 1), pagesize=(11, 8.5)):
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
                             "Part-{}x{}: Page-{}".format(i + 1, j + 1, i * nv + j + 1),
                             ha='center', fontsize=8)
                pdf.savefig(fig, bbox_inches='tight')

                plt.close()


def create_pdf(df, name):
    L = len(df.index)
    npages = int(np.ceil(L / 50))
    dataframe_to_pdf(df, name, numpages=(npages, 1), pagesize=(8.5, 20))


def file_selection(path):
    # returns the selected excel files in the order [website, snelstart, onbekend]
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))  # reading out which excel files are in directory [strings]
    results = []
    titles = ['select website list', 'select snelstart list', 'select unkown user list']
    for i in range(0, 3):
        option, index = pick(csv_files, titles[i])
        results.append(option)
        csv_files.pop(index)
    return results


WEBSITE_MATCH_NAME = "ID-nummer"  # can be any of the column names from the excel file


def clean_website(df):
    # Finds the starting position of the table in the excel file such that it is impossible to paste it incorrectly into excel
    columns = df.columns  # finding column names
    # TODO check toevoegen om te kijken of de colommen al kloppen, want anders maakt dit het stuk hieronder
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
    df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] = np.where((df[OMZETAANTAL] == 1) & (df[AFGEREKEND] == 'Ja'), True,
                                                         False)  # finding correct inboekingen
    df[CORRRECT_NOT_PAID_NAME] = np.where((df[OMZETAANTAL] == 0) & (df[AFGEREKEND] == 'Nee'), True,
                                          False)  # finding correct not paid
    return df


CORRECT_SNELSTART_ONBEKEND = 'ONBEKEND-SNELSTART_MATCH'  #
CORRECT_WEBSITE_ONBEKEND = 'ONBEKEND-WEBSITE_MATCH'


def correct_onbekend(df):
    df[CORRECT_SNELSTART_ONBEKEND] = np.where((df[OMZETAANTAL] == 1) & (df[ONBEKEND_NAAM].notnull()), True, False)
    df[CORRECT_WEBSITE_ONBEKEND] = np.where((df[AFGEREKEND] == 'Ja') & (df[ONBEKEND_NAAM].notnull()), True, False)
    return df


if __name__ == '__main__':
    os_path = (os.getcwd())
    test_path = os_path + "\\testdata"
    # filepaths = file_selection(test_path) #selecting the files to be website, snelstart, unkown

    # TEST
    # filepaths = [test_path + '\\web.xlsx', test_path + '\\asml.xlsx', test_path + '\\onbekend.xlsx']  #loading test files
    filepaths = [test_path + '\\v2\\website.xlsx', test_path + '\\v2\\snelstart.xlsx',
                 test_path + '\\v2\\onbekend.xlsx']  # loading test files

    # reading in the data
    website = pd.read_excel(filepaths[0])
    snelstart = pd.read_excel(filepaths[1])
    onbekend = pd.read_excel(filepaths[2])

    # cleaning up the data
    website = clean_website(website)  # removing possible clutter from the website
    snelstart = format_snelstart(snelstart)
    onbekend = clean_onbekend(onbekend)
    partial_merge = pd.merge(website, snelstart, how="outer", on=['ID-nummer'])
    merged = pd.merge(partial_merge, onbekend, how="outer", on=['ID-nummer'])
    # todo onbekende klant data invoeg functie toevoegen
    merged.to_excel("RAW DATA.xlsx")

    ## Participant check functionality:
    ## Check for GOOD cases specifacally and assign a true or false accordingly
    ##      -> importantly these are very specific and simple correct cases
    ## If done for all the the true cases, only one case should be true for each inboeking
    ##      -> If all are false something is incorrect
    ##      -> If more than 2 are true something is incorrect -> 2 cases can't be correct at the same time
    snelstart_website_correct = correct_website_snelstart(
        merged)  # checking between website and snelstart correct values

    df = correct_onbekend(snelstart_website_correct)
    ##TODO dit filter verbeteren want het ziet er ruk uit
    df['result'] = np.where((
                                    (df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] == True) & (df[
                                CORRRECT_NOT_PAID_NAME] == False) & (df[CORRECT_WEBSITE_ONBEKEND] == False) |
                                    (df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] == False) & (df[
                                        CORRRECT_NOT_PAID_NAME] == True) & (df[CORRECT_WEBSITE_ONBEKEND] == False) |
                                     (df[CORRECT_SNELSTART_WEBSITE_COLUMN_NAME] == False) & (df[
                                        CORRRECT_NOT_PAID_NAME] == False) & (df[CORRECT_WEBSITE_ONBEKEND] == True))
                            & (df[CORRECT_SNELSTART_ONBEKEND] == False), True, False)
    # create_pdf(snelstart_website_correct, 'result.pdf') #writing out final results to pdf
    df.to_excel('snelstartfout.xlsx')  # writing out resulting data
    # todo onbekende klant controle toevoegen
    #   * onbekend - website match
    #   * onbekend - snelstart match
