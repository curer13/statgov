from toServerKATO_v2 import getDfFromExcel, chunker, insert_with_progress
from os import listdir
import re

spravochniki = '/Users/user/Documents/производственная_практика/statgov/spravochniki'


def main():
    for f in listdir(spravochniki):
        if f != '.DS_Store' and not re.search(r'^(~\$)', f):
            print(f)
            tablename = 'spravochnik' + f.split('.')[0]
            df = getDfFromExcel(spravochniki + '/' + f)
            if f.split('.')[0] == '12621':
                column_names = df.iloc[1].to_list()
                df = df.iloc[2:]
                df = df.rename(columns=dict(zip(df.columns, column_names)))
            insert_with_progress(df, tablename)


# df = getDfFromExcel(spravochniki + '/' + '2153.xlsx')
# print(df.tail())
# insert_with_progress(df, 'spravochnik'+'2153')
if __name__ == '__main__':
    main()
