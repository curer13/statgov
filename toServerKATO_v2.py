import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.types import NVARCHAR
from os import listdir
from os.path import isfile, join
import re
import datetime
from tqdm import tqdm
import time


def getShortName(fullname):
    delimiters=[('\"', '\"'), ('\«', '\»'), ('\!', '\!')]

    for delimiter1, delimiter2 in delimiters:
        matches = re.search(f'{delimiter1}(.*){delimiter2}', str(fullname))
        if matches:
            return matches.group(0)[1:-1]

    return fullname


def findLastRequest():
    zayavki = 'zayavki/'
    onlyfiles = {}

    for f in listdir(zayavki):
        filePath = join(zayavki, f)
        if isfile(filePath) and f != '.DS_Store' and not re.search(r'^(~\$)', f):
            onlyfiles[filePath] = None

    print(onlyfiles)

    for file in onlyfiles:
        d_m_y, H_M = getDateFromFile(file)
        req_date = datetime.datetime.strptime(d_m_y + '_' + H_M, '%d-%m-%y_%H-%M')
        onlyfiles[file] = req_date

    return max(onlyfiles, key=lambda req: onlyfiles[req])


def getDateFromFile(file):
    splitlist = file.split('_')
    # print(splitlist)
    d_m_y = splitlist[1]
    H_M = splitlist[2].split('.')[0]
    return d_m_y, H_M


def getDfFromExcel(path):
    # returns dictionary
    df = pd.read_excel(path, sheet_name=None)

    sheetnames = [sheetname for sheetname in df]

    first_sheet = df[sheetnames[0]]

    column_names = first_sheet.iloc[0].to_list()
    # print(column_names)

    column_names = list(map(lambda x: x.strip(), column_names))
    # print(column_names)

    first_sheet = first_sheet.rename(columns=dict(zip(first_sheet.columns, column_names))).iloc[1:]
    # print(first_sheet.columns.to_list())

    dflist = [first_sheet] + [df[sheetname].rename(columns=dict(zip(df[sheetname].columns, column_names))) for sheetname in sheetnames[1:]]

    DF = pd.concat(dflist, ignore_index=True)

    return DF


def chunker(seq, size):
    # from http://stackoverflow.com/a/434328
    if size==0:
        size=1
    return (seq[pos:pos + size] for pos in range(0, len(seq), size))


def insert_with_progress(df, db, table):
    engine = create_engine(
        f'mssql+pyodbc://sa:Abz09jvv1@localhost:1433/{db}?driver=ODBC Driver 17 for SQL Server')
    chunksize = int(len(df) / 10)  # 10%
    with tqdm(total=len(df)) as pbar:
        for i, cdf in enumerate(chunker(df, chunksize)):
            replace = "replace" if i == 0 else "append"
            cdf.to_sql(table, con=engine, if_exists=replace,
                       dtype={col_name: NVARCHAR for col_name in df.columns})
            pbar.update(chunksize)


def insert_without_bar(df, db, table):
    engine = create_engine(
        f'mssql+pyodbc://sa:Abz09jvv1@localhost:1433/{db}?driver=ODBC Driver 17 for SQL Server')
    df.to_sql(table, con=engine, if_exists='replace',
              dtype={col_name: NVARCHAR for col_name in df.columns})


def main():
    # excelFile = findLastRequest()
    excelFile = '/Users/user/Documents/производственная_практика/testbook1.xlsx'
    print(f'Importing {excelFile}')
    df = getDfFromExcel(excelFile)

    df['short_name']=df['Толық атауы'].apply(getShortName)
    # print(len(df))
    # print(df.columns)

    # insert_with_progress(df, 'juridicalFaceDb', 'juridicalInfo')
    insert_with_progress(df, 'employee_task', 'juridicalInfo')
    # insert_without_bar(df, 'juridicalFaceDb', 'juridicalInfo')


if __name__ == '__main__':
    start_time = time.time()
    main()
    # print(getShortName('«how beautiful» this v!iew'))
    # print(str('"QWNG" ЖАУАПКЕРШІЛІГІ ШЕКТЕУЛІ СЕРІКТЕСТІГІ'))
    print("--- %s minutes ---" % ((time.time() - start_time)/60))
