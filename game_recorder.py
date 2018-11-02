import pandas as pd
import os
import datetime



def main():

    #df1 = pd.read_excel("assets/excel/ping_pong_scoresheet.xlsx", "Sheet1")
    df1 = pd.read_excel('assets/excel/ping_pong_scoresheet.xlsx', header=0, skipfooter=3).reset_index(drop=True)
    f_score = get_score('Fritz')
    k_score = get_score('Ken')
    side = input('Winner side? H/A').upper()
    while ((side != 'A') and (side != 'H')):
        side = input('Winner side? H/A').upper()

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

    df1.loc['Total Wins'] = [wins['f'], wins['k'], None, None]
    df1.loc['PPG'] = [ppg['f'], ppg['k'], None, None]
    df1.loc['Total Games'] = [None, total_games, None, None]
    write_xlsx(df1)

def get_score(who):
    x = ''
    while x.isdigit() == False:
        x = input('Enter {}s score: '.format(who))
    return(int(x))


def write_xlsx(df):
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter("assets/excel/ping_pong_scoresheet.xlsx", engine="xlsxwriter")
	df.to_excel(writer, sheet_name="Sheet1")

    # merge_format = workbook.add_format({
    #     'border': 1,
    #     'align': 'center',
    #     'valign': 'vcenter'
    #     )

	workbook = writer.book
	worksheet = writer.sheets["Sheet1"]

    # main_format = workbook.add_format({'font_name':'Calibri', 'font_size':12})
    # game_format = workbook.add_format({'num_format': "#,##0"})
	# win_format = workbook.add_format({})
    # ppg_format = workbook.add_format({'num_format': '0.000'})
	# # Add a format. Green fill with dark green text. GREEN = GOOD
	# win_format = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
    #
	# number_rows = len(df1.index)
	# worksheet1.conditional_format("G2:G{}".format(number_rows), {"type": "cell", "criteria":"==", "value": "True", "format": format_true})

	# Format the columns by width

	# # general info columns
	# worksheet1.set_column("A:A", 3)
	# worksheet1.set_column("B:E", 20)
	# worksheet1.set_column("F:F", 15)
    #
	# # Boolean columns
	# worksheet1.set_column("G:H", 8)
	# worksheet1.set_column("I:I", 35)
    #
	# # general info columns
	# worksheet2.set_column("A:A", 3)
	# worksheet2.set_column("B:D", 20)
	# worksheet2.set_column("E:E", 15)
	# worksheet2.set_column("F:F", 8)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()

main()
