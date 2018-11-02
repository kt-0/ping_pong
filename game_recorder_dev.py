import pandas as pd
import os
import datetime



def main():

    #df1 = pd.read_excel("assets/excel/ping_pong_scoresheet.xlsx", "Sheet1")
    df1 = pd.read_excel('assets/excel/ping_pong_scoresheet_dev.xlsx', parse_dates=['Date'], header=0, skipfooter=3).reset_index(drop=True)
    f_score = get_score('Fritz')
    k_score = get_score('Ken')
    side = input('Winner side? (H/A) ').upper()
    while ((side != 'A') and (side != 'H')):
        side = input('Winner side? (H/A) ').upper()

    df1.loc[df1.shape[0]] = [f_score, k_score, pd.datetime.today().date(), side]

    more = input('More scores to enter? (Y/N)').lower()
    while more[:1] == 'y':
        f_score = get_score('Fritz')
        k_score = get_score('Ken')
        side = input('Winner side? H/A').upper()
        df1.loc[df1.shape[0]] = [f_score, k_score, pd.datetime.today().date(), side]
        more = input('More scores to enter? (Y/N)').lower()

    f_scores = df1['Fritz'].tolist()
    k_scores = df1['Ken'].tolist()

    #df1.loc[df1.shape[0]]=[parent_author, subR, author, parent_id, proposed, actual, parent_body]

    wins = {'f':0, 'k':0}
    ppg = {'f':0, 'k':0}

    for x,y in zip(f_scores,k_scores):
        if x>y:
            wins['f'] = wins['f']+1
        else:
            wins['k'] = wins['k']+1

    ppg['f'] = sum(f_scores)/len(f_scores)
    ppg['k'] = sum(k_scores)/len(k_scores)

    total_games = wins['k'] + wins['f']

    df1.loc['Total Wins'] = [wins['f'], wins['k'], pd.to_datetime(pd.NaT), None]
    df1.loc['PPG'] = [ppg['f'], ppg['k'], pd.to_datetime(pd.NaT), None]
    df1.loc['Total Games'] = [None, total_games, pd.to_datetime(pd.NaT), None]
    write_xlsx(df1)

def get_score(who):
    x = ''
    while x.isdigit() == False:
        x = input('Enter {}s score: '.format(who))
    return(int(x))


def write_xlsx(df):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter("assets/excel/ping_pong_scoresheet_v2.xlsx", engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Sheet1")
    game_rows = (df.shape[0] - 3)
    print("game_rows", game_rows)
    tot_wins_row = (df.shape[0] - 2)
    ppg_row = (df.shape[0] - 1)
    tot_games_row = df.shape[0]

    light_blue = '#e8f4ff'
    grey = '#a7a9aa'
    greyish_blue = '#b6d2ed'
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    odd_row_format = workbook.add_format({'bg_color': 'e8f4ff', 'border_color': grey, 'border':1, 'font_name': 'Helvetica Neue', 'font_size':12 })
    odd_date_format = workbook.add_format({'bg_color': 'e8f4ff', 'border_color': grey, 'border':1, 'font_name': 'Helvetica Neue', 'font_size':10, 'italic':True, 'num_format':'mm/dd/yy'})
    main_format = workbook.add_format({'font_name': 'Helvetica Neue', 'font_size':12})
    head_format = workbook.add_format({
        'font_size': 14,
        'font_color': 'white',
        'bg_color': '#1b4ea0',
        'bold': True,
        'bottom': 1,
        'center_across': True})
    date_format = workbook.add_format({
        'border_color': grey,
        'border': 1,
        'font_size': 10,
        'italic': True,
        'font_name': 'Helvetica Neue',
        'num_format': 'mm/dd/yy'})
    game_format = workbook.add_format({'num_format': '#,##0'})
    ppg_format = workbook.add_format({'num_format': '0.000'})
    win_format = workbook.add_format({'bold': True, 'italic': True})
    tot_wins_format = workbook.add_format({
        'top':2,
        'font_size':14,
        'font_name':'Helvetica Neue',
        'bg_color':'#f2db87',
        'num_format': '#,##0'})

    worksheet.set_column("A:C", cell_format=main_format)
    #worksheet.set_column("$B$1:$B$1", None, cell_format=head_format)

    # format the head
    for col_num, value in enumerate(df.columns.values):
        # worksheet.write(row, col, value, format)
        worksheet.write(0, col_num + 1, value, head_format)

    # light blue fill on every other row
    for i in range(1, game_rows+1, 2):

        x = i-1
        f = df.iloc[x, 0]
        k = df.iloc[x, 1]
        d = df.iloc[x, 2]
        s = df.iloc[x, 3]

        # if i%2==0:
        worksheet.write(i, 1, f, odd_row_format)
        worksheet.write(i, 2, k, odd_row_format)
        worksheet.write_datetime(i, 3, d, odd_date_format)
        worksheet.write(i, 4, s, odd_row_format)

    for i in range(2, game_rows+1, 2):
        x = i-1
        worksheet.write_datetime(i, 3, df.iloc[x, 2], date_format)

    for i in range(1, len(df.columns)-1):
        x = i-1
        worksheet.write(tot_wins_row, i, df.iloc[-3, x], tot_wins_format)

    for i in range(2, game_rows+2):

        x = i-2
        f = df.iloc[x, 0]
        k = df.iloc[x, 1]

        worksheet.conditional_format('B{}:B{}'.format(i,i), {'type':'cell','criteria':'>','value':k,'format': win_format})
        worksheet.conditional_format('C{}:C{}'.format(i,i), {'type':'cell','criteria':'>','value':f,'format': win_format})

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()

main()
