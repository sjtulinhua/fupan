import os
import sys
import shutil
import argparse
import datetime
import pandas as pd
import configs as cfg
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, colors
from colordic import Paired_color_map

g_dest_col_num = 1
g_dest_wb = None


def open_dest_xlsx(fpath):
    # open xlsx
    global g_dest_wb
    if g_dest_wb is None:
        g_dest_wb = load_workbook(fpath)
    return g_dest_wb


def backup_before_proc():
    shutil.copyfile(cfg.dest_sheet_tdx, cfg.backup_sheet_tdx)


def get_dest_wb():
    global g_dest_wb
    return g_dest_wb


def get_dest_col():
    return g_dest_col_num


def insert_data_by_col(workbook, sheet_num, col_idx, data):
    ws = workbook.worksheets[sheet_num]
    ws.insert_cols(col_idx)

    for idx, val in enumerate(data, start=0):
        # print(col_idx, idx, val)
        c = ws.cell(column=col_idx, row=idx + 1, value=val)

        if idx == 0:
            c.fill = PatternFill(fill_type='solid', fgColor='d1ffbd')
        elif val in Paired_color_map.keys():
            c.fill = PatternFill(fill_type='solid', fgColor=Paired_color_map[val])
        else:
            c.fill = PatternFill(fill_type='solid', fgColor='fff39a')

        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.font = Font(name='宋体', size=16)
        c.border = Border(left=Side(style='thin', color=colors.BLACK),
                          right=Side(style='thin', color=colors.BLACK),
                          top=Side(style='thin', color=colors.BLACK),
                          bottom=Side(style='thin', color=colors.BLACK))
        # setup row height
        ws.row_dimensions[idx + 1].height = 25


def insert_data(workbook, sheet_num, col, col_data_list):
    """
    :type col_data_list: list
    """
    for data in col_data_list:
        insert_data_by_col(workbook, sheet_num, col, data)


def data_grouping(data_df, date):
    short_date_str = date.strftime('%m%d')

    cnt_df = data_df.groupby('细分行业')['代码'].count().to_frame(name='count')
    cnt_df.sort_values(['count'], ascending=False, inplace=True)

    # add current date to dataframe
    cnt_df.rename(columns={'count': short_date_str}, inplace=True)
    cnt_df.reset_index(inplace=True)

    # prepare data for writing to the xlsx
    col_cnt = pd.Series(cnt_df.columns.array[1])
    col_cnt = col_cnt.append(cnt_df[short_date_str])
    col_industry = pd.Series(cnt_df.columns.array[0])
    col_industry = col_industry.append(cnt_df['细分行业'])

    # percentage
    size = []
    for industry in cnt_df['细分行业']:
        size.append(cfg.board_size[industry])
    cnt_df['board size'] = size
    cnt_df['占比'] = cnt_df[short_date_str] / cnt_df['board size']
    cnt_df['占比'] = cnt_df['占比'].apply(lambda x: format(x, '.1%'))

    col_pct = pd.Series(cnt_df.columns.array[3])
    col_pct = col_pct.append(cnt_df['占比'])
    return col_cnt, col_industry, col_pct


def update_analysis(datafpath, workbook, sheetnum, date):
    if not os.path.exists(datafpath):
        print(f"Skip non-existing file: {datafpath} ...")
        return

    # start zt analysis
    print(f"processing {datafpath}")

    # read data
    df = pd.read_csv(datafpath, sep='\t', encoding="gbk")

    # do analysis
    col_cnt, col_industry, col_pct = data_grouping(df, date)
    insert_data(workbook, sheetnum, get_dest_col(), [col_pct, col_cnt, col_industry])


    # setup col width
    # todo iterate all columns then setup width separately
    # ws.column_dimensions['A'].width = 12
    # ws.column_dimensions['B'].width = 6.5


def lookup_concept(stock_name, plate_dashboard_df):
    # todo new func, not test
    concept_bucket = []
    for label, content in plate_dashboard_df.iteritems():
        if stock_name in content.values:
            concept_bucket.append(label)
    return concept_bucket


def fupan_main(database):
    parser = argparse.ArgumentParser(description='命令行中传入"YYYYMMDD"格式的日期')
    parser.add_argument("sdate", help="input the start date in the format of 'YYYYMMDD'")
    parser.add_argument("--edate", required=False, help="input the end date in the format of 'YYYYMMDD'")
    parser.add_argument("--destcol", type=int, required=False, help="input the num of col (only accept number) "
                                                                    "before which you want to insert new columns")
    args = parser.parse_args()

    if args.destcol is not None:
        print("optional arg: destcol = " + str(args.destcol))
        global g_dest_col_num
        g_dest_col_num = args.destcol


    for db in database:
        if db == 'tdx':
            print(f'using database:  {db}')
            backup_before_proc()
            workbook = open_dest_xlsx(cfg.dest_sheet_tdx)

            sdate_input = datetime.datetime.strptime(args.sdate, '%Y%m%d')
            if args.edate is not None:
                edate_input = datetime.datetime.strptime(args.edate, '%Y%m%d')
                for i in range((edate_input - sdate_input).days + 1):
                    day = sdate_input + datetime.timedelta(days=i)
                    str_day = day.strftime('%Y%m%d')
                    update_analysis(cfg.sheet_zt + str_day + '.txt', workbook, cfg.sheet_zt_num, day)
                    update_analysis(cfg.sheet_lsxg + str_day + '.txt', workbook, cfg.sheet_lsxg_num, day)
                    update_analysis(cfg.sheet_drps + str_day + '.txt', workbook, cfg.sheet_drps_num, day)
                    update_analysis(cfg.sheet_srps + str_day + '.txt', workbook, cfg.sheet_srps_num, day)
            else:
                str_day = args.sdate
                update_analysis(cfg.sheet_zt + str_day + '.txt', workbook, cfg.sheet_zt_num, sdate_input)
                update_analysis(cfg.sheet_lsxg + str_day + '.txt', workbook, cfg.sheet_lsxg_num, sdate_input)
                update_analysis(cfg.sheet_drps + str_day + '.txt', workbook, cfg.sheet_drps_num, sdate_input)
                update_analysis(cfg.sheet_srps + str_day + '.txt', workbook, cfg.sheet_srps_num, sdate_input)
            # save_to_xlsx
            print("Completed, write to xlsx!")
            get_dest_wb().save(cfg.dest_sheet_tdx)

        elif db == 'futu':
            # todo
            print(f'using database:  {db}')


        elif db == 'tushare':
            # todo
            print(f'using database:  {db}')



    # update_analysis(dtp.sheet_zt+args.sdate+'.txt', sdate_input)



if __name__ == '__main__':
    fupan_main(cfg.database)
