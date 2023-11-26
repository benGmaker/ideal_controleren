import pandas as pd
import os
import glob
from pick import pick

if __name__ == '__main__':
    os_path = (os.getcwd())
    test_path = os_path + "/testdata"
    file_selection(test_path)
    relative_path = os_path + "/testdata/website.xlsx"
    df = pd.read_excel(relative_path)
    print(df)


def file_selection(path):
    #returns the selected excel files in the order [website, snelstart, onbekend]
    csv_files = glob.glob(os.path.join(path, "*.xlsx"))  # reading out which excel files are in directory [strings]
    results = []
    titles = ['select website list','select snelstart list','select unkown user list']
    for i in range(0,2):
        option, index = pick(csv_files, titles(i))
        print(option)
        csv_files.pop(index)
    #def clean_website(df):

