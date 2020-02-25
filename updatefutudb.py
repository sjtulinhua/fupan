import time
import sqlite3
import pandas as pd
from futu import *


def update_concept_db(dpath, ctx):
    concept_list_df = ctx.get_plate_list(Market.SH, Plate.CONCEPT)[1]

    concept_dashboard_code = pd.DataFrame()
    concept_dashboard_name = pd.DataFrame()
    counter = 0
    for idx, row in concept_list_df.iterrows():
        print(str(idx) + '  ' + str(row['plate_name']))
        plate_df = ctx.get_plate_stock(row['code'])[1]  # 30 秒内请求最多 10 次

        concept_dashboard_code = pd.concat([concept_dashboard_code, plate_df['code']], axis=1)
        concept_dashboard_name = pd.concat([concept_dashboard_name, plate_df['stock_name']], axis=1)
        concept_dashboard_code.rename(columns={'code': row['plate_name']}, inplace=True)
        concept_dashboard_name.rename(columns={'stock_name': row['plate_name']}, inplace=True)

        counter = counter + 1
        if (counter > 9):
            counter = 0
            time.sleep(30)

    sql_ctx = sqlite3.connect(dpath)
    concept_list_df.to_sql('concept_list', sql_ctx, if_exists='replace')
    concept_dashboard_code.to_sql('concept_dashboard_code', sql_ctx, if_exists='replace')
    concept_dashboard_name.to_sql('concept_dashboard_name', sql_ctx, if_exists='replace')


if __name__ == '__main__':
    ctx = OpenQuoteContext(host='127.0.0.1', port=11111)
    update_concept_db('/Users/linhua/PycharmProjects/Fupan/database/futu.db', ctx)
    ctx.close()