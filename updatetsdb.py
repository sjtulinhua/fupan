# update tushare database

import sys
import time
from os import path
sys.path.append('/Users/linhua/PycharmProjects/Fupan')
from private import tushare_token
import tushare as ts
import sqlite3
import pandas as pd

# prepare Tushare data interface
ts.set_token(tushare_token.tushare_token)
pro = ts.pro_api()

#
# Database design:
# table: concept_list
#     | concept_code | concept_name | src |

# table: concept_dashboard
#     | concept_name (1) | concept_name (2) |
#

def update_concept_database(dpath):
    concept_list_df = pro.concept()

    concept_dashboard_code = pd.DataFrame()
    concept_dashboard_name = pd.DataFrame()
    counter = 0

    for idx, row in concept_list_df.iterrows():
        print(str(idx) + '  ' + str(row['name']))
        df = pro.concept_detail(id=row.code, fields='ts_code,name')
        concept_dashboard_code[row['name']] = df['ts_code']
        concept_dashboard_name[row['name']] = df['name']
        concept_dashboard_code = pd.concat([concept_dashboard_code, df['ts_code']], axis=1)
        concept_dashboard_name = pd.concat([concept_dashboard_name, df['name']], axis=1)

        counter = counter + 1
        if (counter > 90):
            counter = 0
            time.sleep(10)

    context = sqlite3.connect(dpath)
    concept_list_df.to_sql('concept_list', context, if_exists='replace')
    concept_dashboard_code.to_sql('concept_dashboard_code', context, if_exists='replace')
    concept_dashboard_name.to_sql('concept_dashboard_name', context, if_exists='replace')

    return concept_list_df, concept_dashboard_code, concept_dashboard_name


if __name__ == '__main__':
    update_concept_database('/Users/linhua/PycharmProjects/Fupan/database/tushare.db')