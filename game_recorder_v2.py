import pandas as pd
import numpy as np
import os
import datetime, math, random

# TODO: add timestamps, then use them for multi index, then put the timestamps in a hidden column

def main():

	#df1 = pd.read_excel("assets/excel/ping_pong_scoresheet.xlsx", "Sheet1")
	df1 = pd.read_excel('assets/excel/ping_pong_scoresheet_v2.xlsx', header=0, skipfooter=14)
	date = pd.DatetimeIndex(df.index, name='Date')
	df1.index = date
	print(df1)
	f_score = get_score('Fritz')
	k_score = get_score('Ken')
	side = input('Winner side? H/A').upper()
	#start_date = pd.to_datetime('10/11/18', format="%m/%d/%y")
	#df1['Date'].fillna(start_date, inplace=True)
	#df1['Side'].fillna(lambda x: rand_side(['H','A']), inplace=True)

	#df1['Side'] = df1['Side'].apply(lambda x: rand_side(['H','A']))
	while ((side != 'A') and (side != 'H')):
		side = input('Winner side? H/A').upper()

	today = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')

	dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score]}, index=today)
	df1 = df1.append(dfnew, sort=False)
	#df1.loc[pd.datetime.today().date()] = [f_score, k_score, side]

	more = input('More scores to enter? (Y/N)').lower()

	while more[:1] == 'y':
		f_score = get_score('Fritz')
		k_score = get_score('Ken')
		side = input('Winner side? H/A').upper()
		dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Date':pd.datetime.date.today().date()} index='Date')
		df1 = df1.append(dfnew)
		more = input('More scores to enter? (Y/N)').lower()

	f_scores = df1['Fritz'].tolist()
	k_scores = df1['Ken'].tolist()
	sides = df1['Side'].tolist()

	total_games = df1.shape[0]

	stats = {
		'f': {
			'A': {'wins':0,'win_perc':0,'ppg':0},
			'H': {'wins':0,'win_perc':0,'ppg':0},
			'overall': {
				'wins':0,
				'win_perc':0,
				'ppg': nr(sum(f_scores)/total_games)
				}
		},
		'k':{
			'A': {'wins':0,'win_perc':0,'ppg':0},
			'H': {'wins':0,'win_perc':0,'ppg':0},
			'overall': {
				'wins':0,
				'win_perc':0,
				'ppg':nr(sum(k_scores)/total_games)
				}
		}
			}

	opposite = str.maketrans("fkAH", "kfHA")

	for i,x in enumerate(zip(f_scores,k_scores,sides)):
			f = x[0]
			k = x[1]
			side = x[2]
			opp = side.translate(opposite)
			if f>k:
				stats['f'][side]['wins'] = stats['f'][side]['wins']+1
				stats['f'][side]['ppg'] = (stats['f'][side]['ppg']+f)
				stats['k'][opp]['ppg'] = (stats['k'][opp]['ppg']+k)

			else:
				stats['k'][side]['wins'] = stats['k'][side]['wins']+1
				stats['k'][side]['ppg'] = (stats['k'][side]['ppg']+k)
				stats['f'][opp]['ppg'] = (stats['f'][opp]['ppg']+f)


	for x in stats:
		y = x.translate(opposite)
		home,away = 'H','A'
		stats[x][away]['ppg'] = stats[x][away]['ppg']/(stats[x][away]['wins']+stats[y][home]['wins'])
		stats[x][home]['ppg'] = stats[x][home]['ppg']/(stats[x][home]['wins']+stats[y][away]['wins'])
		stats[x][away]['win_perc'] = (stats[x][away]['wins']/(stats[x][away]['wins']+stats[y][home]['wins']))*100
		stats[y][home]['win_perc'] = (stats[y][home]['wins']/(stats[y][home]['wins']+stats[x][away]['wins']))*100

		stats[x]['overall']['wins'] = (stats[x][home]['wins']+stats[x][away]['wins'])
		stats[x]['overall']['win_perc'] = (stats[x]['overall']['wins']/total_games)*100

	multi_ind = [
		np.array(["wins", "wins", "wins", "win%", "win%", "win%", "ppg", "ppg", "ppg", "streak", "streak"]),
		np.array(["H", "A", "overall", "H", "A", "overall", "H", "A", "overall", "Current", "Best"])]

	dfstats = pd.DataFrame(
		{
		'F':[
			stats['f']['H']['wins'], stats['f']['A']['wins'], stats['f']['overall']['wins'],
			stats['f']['H']['win_perc'], stats['f']['A']['win_perc'], stats['f']['overall']['win_perc'],
			stats['f']['H']['ppg'], stats['f']['A']['ppg'], stats['f']['overall']['ppg'], 0, 0],
		'K':[
			stats['k']['H']['wins'], stats['k']['A']['wins'], stats['k']['overall']['wins'],
			stats['k']['H']['win_perc'], stats['k']['A']['win_perc'], stats['k']['overall']['win_perc'],
			stats['k']['H']['ppg'], stats['k']['A']['ppg'], stats['k']['overall']['ppg'], 0, 0]
		}, index=multi_ind)


	write_xlsx(df1, dfstats)

# rounding function (Normal Round)
def nr(n, decimals=2):
	expoN = n * 10 ** decimals
	if abs(expoN) - abs(math.floor(expoN)) < 0.5:
		return math.floor(expoN) / 10 ** decimals
	return math.ceil(expoN) / 10 ** decimals

def get_score(who):
	x = ''
	while x.isdigit() == False:
		x = input('Enter {}s score: '.format(who))
	return(int(x))

def rand_side(x):
	return(random.choice(x))

def percentify(x):
	x = nr(x)
	return("{:.2%}".format(x))

def write_xlsx(df, dfstats):
	# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter("assets/excel/ping_pong_scoresheet_v2.xlsx", engine="xlsxwriter")
	df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0)
	dfstats.to_excel(writer, sheet_name="Sheet1", startrow=(df.shape[0]+2), startcol=0)
	workbook = writer.book
	worksheet = writer.sheets["Sheet1"]

	game_rows = df.shape[0]
	# tot_wins_row = (df.shape[0] - 4)
	# ppg_row = (df.shape[0] - 3)
	#tot_games_row = (df.shape[0] - 1)
	#h_vs_a_row = df.shape[0]

	light_blue = '#e8f4ff'
	grey = '#a7a9aa'
	greyish_blue = '#b6d2ed'
	dark_blue = '#1b4ea0'

	odd_row_format = workbook.add_format({
		'bg_color': light_blue,
		'border_color': grey,
		'border':1,
		'font_name': 'Helvetica Neue',
		'font_size':12 })

	odd_side_format = workbook.add_format({
		'bg_color': light_blue,
		'border_color': grey,
		'border':1,
		'font_name': 'Helvetica Neue',
		'font_size':10,
		'italic':True })

	odd_date_format = workbook.add_format({
		'bg_color': light_blue,
		'border_color': grey,
		'border':1,
		'font_name': 'Helvetica Neue',
		'font_size':10,
		'italic':True,
		'num_format':'mm/dd/yy'})

	main_format = workbook.add_format({
		'font_name': 'Helvetica Neue',
		'font_size':12})

	head_format = workbook.add_format({
		'font_size': 14,
		'font_color': 'white',
		'bg_color': dark_blue,
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
		'top':1,
		'font_size':14,
		'font_name':'Helvetica Neue',
		'bg_color':'#f2db87',
		'num_format': '#,##0'})

	worksheet.set_column("A:E", cell_format=main_format)
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
		#d = df.iloc[x, 2]
		s = df.iloc[x, 2]

		worksheet.write(i, 1, f, odd_row_format)
		worksheet.write(i, 2, k, odd_row_format)

		#worksheet.write_datetime(i, 3, d, odd_date_format)
		worksheet.write(i, 3, s, odd_side_format)

	# for i in range(2, game_rows+1, 2):
	# 	x = i-1
	# 	worksheet.write_datetime(i, 3, df.iloc[x, 2], date_format)

	# for i in range(1, len(df.columns)-1):
	# 	x = i-1
	# 	worksheet.write(tot_wins_row, i, df.iloc[-4, x], tot_wins_format)

	for i in range(2, game_rows+2):

		x = i-2
		f = df.iloc[x, 0]
		k = df.iloc[x, 1]

		worksheet.conditional_format('B{}:B{}'.format(i,i), {'type':'cell','criteria':'>','value':k,'format': win_format})
		worksheet.conditional_format('C{}:C{}'.format(i,i), {'type':'cell','criteria':'>','value':f,'format': win_format})


	#worksheet.set_column('A:A{}'.format(game_rows), None, None, {'hidden': True})

	writer.save()


main()


# old stuff

# stats['k']['H']['win_perc'] = "{:.2%}".format(stats['k']['H']['wins']/(stats['k']['H']['wins']+stats['f']['A']['wins']))
# stats['f']['H']['win_perc'] = "{:.2%}".format(stats['f']['H']['wins']/(stats['f']['H']['wins']+stats['k']['A']['wins']))
# stats['k']['A']['win_perc'] = "{:.2%}".format(stats['k']['A']['wins']/(stats['k']['A']['wins']+stats['f']['H']['wins']))
# stats['f']['A']['win_perc'] = "{:.2%}".format(stats['f']['A']['wins']/(stats['f']['A']['wins']+stats['k']['H']['wins']))

# wins = {'f':{'A':0,'H':0}, 'k':{'A':0,'H':0}}
# win_perc = {'f':{'A':0,'H':0}, 'k':{'A':0,'H':0}}
# ppg = {'f':{'A':0,'H':0}, 'k':{'A':0,'H':0}}

# for i,x in enumerate(zip(f_scores,k_scores)):
# 	if x[0]>x[1]:
# 		wins['f'][sides[i]] = wins['f'][sides[i]]+1
# 	else:
# 		wins['k'][sides[i]] = wins['k'][sides[i]]+1


# ppg['f'] = sum(f_scores)/len(f_scores)
# ppg['k'] = sum(k_scores)/len(k_scores)

# total_games = wins['k'] + wins['f']


# advantage = {'H':((df1['Side'].tolist()).count('H'))/total_games, 'A':((df1['Side'].tolist()).count('A'))/total_games}
# print(advantage)
# ppg['f'] = sum(f_scores)/len(f_scores)
# ppg['k'] = sum(k_scores)/len(k_scores)
# total_games = wins['k'] + wins['f']
# win_perc = {'f':0, 'k':0}
# win_perc['f'] = "{:.2%}".format(wins['f']/total_games)
# win_perc['k'] = "{:.2%}".format(wins['k']/total_games)

# dfstats.loc['Total Wins'] = [stats['f']['overall']['wins'], stats['k']['overall']['wins'], pd.to_datetime(pd.NaT).date(), None]
# dfstats.loc['PPG'] = [stats['f']['overall']['ppg'], stats['k']['overall']['ppg'], pd.to_datetime(pd.NaT).date(), None]
# dfstats.loc['Win % (H)'] = [percentify(stats['f']['H']['win_perc']), percentify(stats['k']['H']['win_perc']), pd.to_datetime(pd.NaT).date(), None]
# dfstats.loc['Win % (A)'] = [percentify(stats['f']['A']['win_perc']), percentify(stats['k']['A']['win_perc']), pd.to_datetime(pd.NaT).date(), None]
# dfstats.loc['Overall Win%'] = [percentify(stats['f']['overall']['win_perc']), percentify(stats['k']['overall']['win_perc']), pd.to_datetime(pd.NaT).date(), None]
