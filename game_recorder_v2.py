import pandas as pd
import numpy as np
import os
import datetime, math, random

# TODO: add timestamps, then use them for multi index, then put the timestamps in a hidden column

class Stats:

	class H:

	H = H()

def main():

	df1 = pd.read_excel('assets/excel/ping_pong_scoresheet_v2.xlsx', header=0, skipfooter=13).reset_index(drop=True)
	dates = pd.DatetimeIndex(df1.Date.dt.date, name='Date')
	start_games = df1.shape[0]
	df1.index = dates

	df1.drop(columns='Date',inplace=True)
	#print(df1.head())

	f_score = get_score('Fritz')
	k_score = get_score('Ken')
	side = input('Winner side? H/A').upper()
	#start_date = pd.to_datetime('10/11/18', format="%m/%d/%y")
	#df1['Date'].fillna(start_date, inplace=True)
	#df1['Side'].fillna(lambda x: rand_side(['H','A']), inplace=True)

	#df1['Side'] = df1['Side'].apply(lambda x: rand_side(['H','A']))
	while ((side != 'A') and (side != 'H')):
		side = input('Winner side? H/A').upper()

	now = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')

	dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=now)
	df1 = df1.append(dfnew, sort=False)
	#df1.loc[pd.datetime.today().date()] = [f_score, k_score, side]

	more = input('More scores to enter? (Y/N)').lower()

	while more[:1] == 'y':
		f_score = get_score('Fritz')
		k_score = get_score('Ken')
		side = input('Winner side? H/A').upper()
		now = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')
		dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=now)
		df1 = df1.append(dfnew, sort=False)
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
				},
			'streak': {
				'current':0,
				'best': 0
				}
			},
		'k':{
			'A': {'wins':0,'win_perc':0,'ppg':0},
			'H': {'wins':0,'win_perc':0,'ppg':0},
			'overall': {
				'wins':0,
				'win_perc':0,
				'ppg':nr(sum(k_scores)/total_games)
				},
			'streak': {
				'current':0,
				'best': 0
				}
			}
		}

	#print(sides)
	opposite = str.maketrans("fkAH", "kfHA")

	for i,x in enumerate(zip(f_scores,k_scores,sides)):
			f = x[0]
			k = x[1]
			side = x[2]
			opp = side.translate(opposite)
			#print(opp)
			if f>k:
				stats['f'][side]['wins'] = stats['f'][side]['wins']+1
				stats['f'][side]['ppg'] = (stats['f'][side]['ppg']+f)
				stats['k'][opp]['ppg'] = (stats['k'][opp]['ppg']+k)
				stats['f']['streak']['current'] = stats['f']['streak']['current']+1
				stats['k']['streak']['current'] = 0
				if (stats['f']['streak']['current'] > stats['f']['streak']['best']):
					stats['f']['streak']['best'] = stats['f']['streak']['current']

			else:
				stats['k'][side]['wins'] = stats['k'][side]['wins']+1
				stats['k'][side]['ppg'] = (stats['k'][side]['ppg']+k)
				stats['f'][opp]['ppg'] = (stats['f'][opp]['ppg']+f)
				stats['k']['streak']['current'] = stats['k']['streak']['current']+1
				stats['f']['streak']['current'] = 0
				if (stats['k']['streak']['current'] > stats['k']['streak']['best']):
					stats['k']['streak']['best'] = stats['k']['streak']['current']


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
			nr(stats['f']['H']['win_perc']), nr(stats['f']['A']['win_perc']),
			nr(stats['f']['overall']['win_perc']), nr(stats['f']['H']['ppg']),
			nr(stats['f']['A']['ppg']), nr(stats['f']['overall']['ppg']),
			nr(stats['f']['streak']['current'],0), nr(stats['f']['streak']['best'],0)],
		'K':[
			stats['k']['H']['wins'], stats['k']['A']['wins'], stats['k']['overall']['wins'],
			nr(stats['k']['H']['win_perc']), nr(stats['k']['A']['win_perc']),
			nr(stats['k']['overall']['win_perc']), nr(stats['k']['H']['ppg']),
			nr(stats['k']['A']['ppg']), nr(stats['k']['overall']['ppg']),
			nr(stats['k']['streak']['current'],0), nr(stats['k']['streak']['best'],0)]
		}, index=multi_ind)


	curr_games = df1.shape[0]

	print("\n====== Games Entered ======\n")
	print(df1.tail(n=curr_games-start_games))
	print("\n======== Stats ========\n")
	print(dfstats.head(n=12))
	print("\nGames played: ", df1.shape[0])

	write_xlsx(df1, dfstats)

############
# modular functions
############
## rounding function (nr = Normal Round)
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

############
# Function for writing/saving the dataframes to excel
############
def write_xlsx(df, dfstats):
	writer = pd.ExcelWriter("assets/excel/ping_pong_scoresheet_v2.xlsx", engine="xlsxwriter", date_format='mm/dd/yy', datetime_format='mm/dd/yy')
	df.to_excel(writer, sheet_name="Sheet1", startrow=0, startcol=0)
	dfstats.to_excel(writer, sheet_name="Sheet1", startrow=(df.shape[0]+2), startcol=0)
	workbook = writer.book
	worksheet = writer.sheets["Sheet1"]

	stats_head_row = df.shape[0]+2

	game_rows = df.shape[0]
	win_index,winper_index,ppg_index = game_rows+3,game_rows+6,game_rows+9

	# tot_wins_row = (df.shape[0] - 4)
	# ppg_row = (df.shape[0] - 3)
	#tot_games_row = (df.shape[0] - 1)
	#h_vs_a_row = df.shape[0]

	light_blue, dark_blue, mid_blue = '#D9E1F1', '#4674C1', '#8FAAD9'
	grey, greyish_blue = '#C9C9C9', '#b6d2ed'
	yellow = '#FED330'
	maroon = '#A9180F'
	odd_row_format = workbook.add_format({
		'bg_color': light_blue,
		'border':1,
		'top_color': mid_blue,
		'bottom_color': mid_blue,
		'left_color': grey,
		'right_color': grey,
		'font_name': 'Helvetica Neue',
		'font_size':12 })

	odd_side_format = workbook.add_format({
		'bg_color': light_blue,
		'border':1,
		'top_color': mid_blue,
		'bottom_color': mid_blue,
		'left_color': grey,
		'right_color': grey,
		'font_name': 'Helvetica Neue',
		'font_size':10,
		'italic':True })

	odd_date_format = workbook.add_format({
		'bg_color': light_blue,
		'border':1,
		'top_color': mid_blue,
		'bottom_color': mid_blue,
		'left_color': grey,
		'right_color': grey,
		'font_name': 'Helvetica Neue',
		'font_size':10,
		'italic':True,
		'num_format':'mm/dd/yy'})

	main_format = workbook.add_format({
		'font_name': 'Helvetica Neue',
		'font_size':12})

	head_format = workbook.add_format({
		'font_size': 14,
		'font_name': 'AppleGothic',
		'font_color': 'white',
		'bg_color': dark_blue,
		'bold': True,
		'bottom': 1,
		'center_across': True })

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

	side_format = workbook.add_format({
		'border_color': grey,
		'border':1,
		'font_name': 'Helvetica Neue',
		'font_size':10,
		'italic':True })

	stat_head_format = workbook.add_format({
		'font_size': 14,
		'font_name': 'AppleGothic',
		'font_color': 'white',
		'bg_color': dark_blue,
		'bold': True,
		'align': 'center',
		'valign': 'vcenter'
	})

	worksheet.set_column("A:E", cell_format=main_format)

	#TODO: implement this TODO TODO
	# worksheet.set_column("A1:A{}".format(game_rows), cell_format=date_format)
	# TODO TODO ^^^^^^^

	#worksheet.set_column("$B$1:$B$1", None, cell_format=head_format)

	# format the head
	for col_num, value in enumerate(df.columns.values):
		# worksheet.write(row, col, value, format)
		#print("col_num: ", col_num, " value: ", value)
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

	for i in range(2, game_rows, 2):
		x = i-1
		s = df.iloc[x, 2]
		worksheet.write(i, 3, s, side_format)
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

	worksheet.write(win_index, 0, dfstats.index[0][0], win_head_format)
	#worksheet.merge_range('B4:D4', 'Merged Range', merge_format)


	# print("index[0]: ", dfstats.index[0])
	# print("index[1]: ", dfstats.index[1])
	# print("index[0][0]: ",dfstats.index[0][0])
	# print("index[0][1]: ",dfstats.index[0][1])
	# print("index[1][1]: ",dfstats.index[1][1])
	writer.save()


main()

class Playerstats:
	def __init__(self, home=Side(), away=Side()):

		self.a = away
		self.h = home
		self._wins = self.a.wins + self.h.wins
		self._win_perc = (self.a.wins+self.h.wins)/(self.a.games+self.h.games)

	@property
	def wins(self):
		self._wins = self.a.wins + self.h.wins
		return(self._wins)

	@property
	def ppg(self):
		self._ppg = self.points/(self.wins+self.losses)
		return(self._ppg)

	@property
	def win_perc(self):
		self._win_perc = (self.a.wins+self.h.wins)/(self.a.games+self.h.games)
		return(self._win_perc)






class Side:
	def __init__(self, w=0, l=0, pts=0):
		self.wins = w
		self.losses = l
		self.points = pts
		self._games = self.wins + self.losses
		self._win_perc = 0
		self._ppg = 0

	@property
	def win_perc(self):
		self._win_perc = self.wins/(self.games)
		return(self._win_perc)

	@property
	def ppg(self):
		self._ppg = self.points/(self.games)
		return(self._ppg)

	@property
	def games(self):
		self._games = self.wins+self.losses
		return(self._games)
