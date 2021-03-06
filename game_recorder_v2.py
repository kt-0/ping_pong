import pandas as pd
import numpy as np
import os
import datetime, math, random

# TODO: add timestamps, then use them for multi index, then put the timestamps in a hidden column


class Stats:
	def __init__(self, w=0, l=0, pts=0):

		self._wins = w
		self._losses = l
		self._win_perc = 0
		self._ppg = 0
		self._points = pts
		self._games = self._wins + self._losses

	@property
	def wins(self):
		return(self._wins)

	@wins.setter
	def wins(self, value):
		self._wins = value
		return(self._wins)

	@property
	def losses(self):
		return(self._losses)

	@losses.setter
	def losses(self, value):
		self._losses = value
		return(self._losses)

	@property
	def points(self):
		return(self._points)

	@points.setter
	def points(self, value):
		self._points = value
		return(self._points)

	@property
	def games(self):
		self._games = self._wins + self._losses
		return(self._games)

	@property
	def ppg(self):
		return(self._ppg)

	@ppg.setter
	def ppg(self, value):
		self._ppg = value
		return(self._ppg)

	# @property
	# def ppg(self):
	# 	self._ppg = self.points/self.games
	# 	return(self._ppg)

	@property
	def win_perc(self):
		self._win_perc = (self._wins/self.games)*100
		return(self._win_perc)

	@property
	def vic_marg(self):
		return(self._vic_margin)

	@vic_marg.setter
	def vic_marg(self, value):
		self._vic_marg = value
		return(self._vic_marg)



class Player(Stats):
	def __init__(self, name):
		Stats.__init__(self)
		self.h = Stats()
		self.a = Stats()
		self.streak = Streak()
		self.name = name

	def __str__(self):
		return(self.name)

	@property
	def games(self):
		return(self.a.games+self.h.games)

	@property
	def wins(self):
		return(self.a.wins+self.h.wins)

	@property
	def losses(self):
		return(self.a.losses+self.h.losses)
	# @property
	# def ppg(self):
	# 	return((self.a.points+self.h.points)/(self.a.games+self.h.games))

	@property
	def ppg(self):
		return(self._ppg)

	@ppg.setter
	def ppg(self, value):
		self._ppg = value
		return(self._ppg)

	@property
	def win_perc(self):
		return(((self.a.wins+self.h.wins)/(self.a.games+self.h.games))*100)

	@property
	def vic_marg(self):
		return(self._vic_marg)

	@vic_marg.setter
	def vic_marg(self, value):
		self._vic_marg = value
		return(self._vic_marg)

	# @property
	# def vic_marg(self):
	# 	return((self.a.wins/(self.wins*self.a.vic_marg)) + (self.h.wins/(self.wins*self.h.vic_marg)))

class Streak:
	def __init__(self, curr=0, b=0):
		self.current = curr
		self.best = b


def main():

	df1 = pd.read_excel('/Users/ktuten/Desktop/ping_pong/assets/excel/ping_pong_scoresheet_v2.xlsx', parse_dates=['Time'], header=0, skipfooter=13).reset_index(drop=True)
	dates = pd.DatetimeIndex(df1.Date.dt.date, name='Date')
	start_games = df1.shape[0]
	df1.index = dates
	df1.drop(columns='Date', inplace=True)

	break_start_time = datetime.datetime.strptime(input('Enter break start time <mm:hh pm> '), '%I:%M %p').time()
	tmp_dt = datetime.datetime.combine(datetime.datetime.today(), break_start_time)
	break_length = datetime.datetime.now() - tmp_dt
	break_length = datetime.timedelta(seconds=round(break_length.total_seconds()))

	f_score = get_score('Fritz')
	k_score = get_score('Ken')
	side = input('Winner side? H/A').upper()

	while ((side != 'A') and (side != 'H')):
		side = input('Winner side? H/A').upper()

	date_index = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')
	dfnew = pd.DataFrame({'Fritz': [f_score], 'Ken': [k_score], 'Side': [side]}, index=date_index)
	more = input('More scores to enter? (Y/N)').lower() or "y"

	while more[:1] == 'y':
		f_score = get_score('Fritz')
		k_score = get_score('Ken')
		side = input('Winner side? H/A').upper()
		while ((side != 'A') and (side != 'H')):
			side = input('Winner side? H/A').upper()
		dfnew = dfnew.append(pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=date_index), sort=False)
		more = input('More scores to enter? (Y/N)').lower() or "y"

	warmup_time = datetime.timedelta(minutes=2.5)
	total_points = sum(dfnew.sum(axis=1))
	time_deltas = []

	for _i,v in enumerate(dfnew.sum(axis=1)):
		time = v/total_points*break_length
		time_deltas.append(datetime.timedelta(seconds=round(time.total_seconds())))

	time_list = [(tmp_dt + warmup_time + sum(time_deltas[:i], datetime.timedelta())) for i,v in enumerate(time_deltas)]

	dfnew['Time'] = time_list
	df1 = df1.append(dfnew, sort=False)

	f_scores = df1['Fritz'].tolist()
	k_scores = df1['Ken'].tolist()
	sides = df1['Side'].tolist()

	ken = Player('Ken')
	fritz = Player('Fritz')

	some_df1 = df1.loc[(((df1.Fritz > df1.Ken) & (df1.Side == 'A')) | ((df1.Ken > df1.Fritz) & (df1.Side == 'H'))), ['Fritz','Ken','Side']]
	ppg1 = some_df1.agg({'Fritz':"mean", 'Ken':"mean", 'Side':"count"})
	some_df2 = df1.loc[(((df1.Fritz > df1.Ken) & (df1.Side == 'H')) | ((df1.Ken > df1.Fritz) & (df1.Side == 'A'))), ['Fritz','Ken','Side']]
	ppg2 = some_df2.agg({'Fritz':"mean", 'Ken':"mean", 'Side':"count"})

	fritz.a.wins,fritz.a.losses = some_df1.groupby('Side')[['Fritz','Ken']].count().Fritz
	ken.h.losses,ken.h.wins = some_df1.groupby('Side')[['Fritz','Ken']].count().Ken

	fritz.h.losses,fritz.h.wins = some_df2.groupby('Side')[['Fritz','Ken']].count().Fritz
	ken.a.wins, ken.a.losses = some_df2.groupby('Side')[['Fritz','Ken']].count().Ken

	fritz.a.ppg, ken.h.ppg, _ = ppg1
	fritz.h.ppg, ken.a.ppg, _ = ppg2

	ken.ppg = df1.Ken.mean()
	fritz.ppg = df1.Fritz.mean()

	dftemp_k = df1.loc[(df1.Fritz < df1.Ken), ['Fritz', 'Ken', 'Side']]
	ken.vic_marg = dftemp_k.diff(axis=1, periods=1).Ken.mean()
	ken.a.vic_marg = dftemp_k.loc[(dftemp_k.Side == 'A'), ['Fritz', 'Ken']].diff(axis=1,periods=1).Ken.mean()
	ken.h.vic_marg = dftemp_k.loc[(dftemp_k.Side == 'H'), ['Fritz', 'Ken']].diff(axis=1,periods=1).Ken.mean()


	dftemp_f = df1.loc[(df1.Fritz > df1.Ken), ['Fritz', 'Ken', 'Side']]
	fritz.vic_marg = dftemp_f.diff(axis=1, periods=-1).Fritz.mean()
	fritz.a.vic_marg = dftemp_f.loc[(dftemp_f.Side == 'A'), ['Fritz', 'Ken']].diff(axis=1,periods=-1).Fritz.mean()
	fritz.h.vic_marg = dftemp_f.loc[(dftemp_f.Side == 'H'), ['Fritz', 'Ken']].diff(axis=1,periods=-1).Fritz.mean()

	for _,x in enumerate(zip(f_scores, k_scores, sides)):
			f = x[0]
			k = x[1]
			side = x[2]
			if f>k:
				ken.streak.current = 0
				fritz.streak.current = fritz.streak.current+1
				if (fritz.streak.current > fritz.streak.best):
					fritz.streak.best = fritz.streak.current
				# if side == 'A':
				# 	fritz.a.wins = fritz.a.wins+1
				# 	fritz.a.points = fritz.a.points+f
				# 	ken.h.losses = ken.h.losses+1
				# 	ken.h.points = ken.h.points+k
				# if side == 'H':
				# 	fritz.h.wins = fritz.h.wins+1
				# 	fritz.h.points = fritz.h.points+f
				# 	ken.a.losses = ken.a.losses+1
				# 	ken.a.points = ken.a.points+k

			if k>f:
				fritz.streak.current = 0
				ken.streak.current = ken.streak.current+1
				if (ken.streak.current > ken.streak.best):
					ken.streak.best = ken.streak.current
				# if side == 'A':
				# 	ken.a.wins = ken.a.wins+1
				# 	ken.a.points = ken.a.points+k
				# 	fritz.h.losses = fritz.h.losses+1
				# 	fritz.h.points = fritz.h.points+f
				# if side == 'H':
				# 	ken.h.wins = ken.h.wins+1
				# 	ken.h.points = ken.h.points+k
				# 	fritz.a.losses = fritz.a.losses+1
				# 	fritz.a.points = fritz.a.points+f


	multi_ind = [
		np.array(["wins", "wins", "wins", "win%", "win%", "win%", "ppg", "ppg", "ppg", "streak", "streak"]),
		np.array(["H", "A", "overall", "H", "A", "overall", "H", "A", "overall", "Current", "Best"])]

	dfstats = pd.DataFrame(
		{
		'F':[
			fritz.h.wins, fritz.a.wins, fritz.wins,
			nr(fritz.h.win_perc), nr(fritz.a.win_perc), nr(fritz.win_perc),
			nr(fritz.h.ppg), nr(fritz.a.ppg), nr(fritz.ppg),
			fritz.streak.current, fritz.streak.best
			],
		'K': [
			ken.h.wins, ken.a.wins, ken.wins,
			nr(ken.h.win_perc), nr(ken.a.win_perc),nr(ken.win_perc),
			nr(ken.h.ppg), nr(ken.a.ppg), nr(ken.ppg),
			ken.streak.current, ken.streak.best
			]
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
# rounding function (nr = Normal Round)
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
	writer = pd.ExcelWriter("assets/excel/ping_pong_scoresheet_v2.xlsx", engine="xlsxwriter")
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
	yellow, maroon = '#FED330', '#A9180F'

	base = {
		'font_name': 'Helvetica Neue',
		'align': 'center',
		'font_size': 10
	}

	odd_base = {**base,
		'bg_color': light_blue,
		'border': 1,
		'top_color': mid_blue,
		'bottom_color': mid_blue,
		'left_color': grey,
		'right_color': grey,
		}

	date = {**base, 'italic': True, 'num_format': 'mm/dd/yy'}
	game = {**base, 'align': 'right', 'num_format': '#,##0', 'font_size': 12}
	side = {**base, 'italic': True}
	stat = {**base, 'num_format': '0.00', 'font_size': 12}
	time = {**base, 'num_format': 'hh:mm AM/PM'}
	f_ = {'right':6, 'right_color': maroon}

	stat_int = {**stat, 'num_format': '#,##0'}
	stat_f = {**stat, **f_}
	stat_int_f = {**stat_int, **f_}
	overall = {**stat,
		'bg_color': yellow,
		'top': 1,
		'top_color': 'black',
		'bottom': 13,
		'bottom_color': grey
		}


	overall_wins = {**overall, 'num_format': '#,##0'}
	overall_wins_f = {**overall_wins, **f_}
	overall_stat = overall #for ppg & win_perc
	overall_stat_f = {**overall, **f_}

	odd_side = {**odd_base, **side}
	odd_date = {**odd_base, **date}
	odd_game = {**odd_base, **game}
	odd_date = {**odd_base, **date}
	odd_time = {**odd_base, **time}

	head = {**base,
		'font_name': 'AppleGothic',
		'font_size': 13,
		'font_color': 'white',
		'bg_color': dark_blue,
		'bold': True,
		'bottom': 1,
		}
	stat_head = {**head, #F, K
		'valign': 'vcenter',
		'bottom':13,
		'bottom_color': maroon
		}
	stat_head_streak = {**stat_head, #streak
		'right': 1,
		'border_color': grey
		}
	stat_head_index1 = {**stat_head_streak, 'bottom': 13} #wins, win%, ppg

	#h, a, current, best
	stat_head_index2 = {**stat_head,
		'font_name': 'Calibri',
		'align': 'center',
		'bg_color': light_blue,
		'border': 1,
		'border_color': grey,
		'font_color': 'black'
		}

	date_format = workbook.add_format(date)
	game_format = workbook.add_format(game)
	head_format = workbook.add_format(head)
	time_format = workbook.add_format(time)
	odd_date_format = workbook.add_format(odd_date)
	odd_game_format = workbook.add_format(odd_game)
	odd_side_format = workbook.add_format(odd_side)
	odd_time_format = workbook.add_format(odd_time)
	overall_stat_format = workbook.add_format(overall_stat)
	overall_stat_f_format = workbook.add_format(overall_stat_f)
	overall_wins_f_format = workbook.add_format(overall_wins_f)
	overall_wins_format = workbook.add_format(overall_wins)
	side_format = workbook.add_format(side)
	stat_f_format = workbook.add_format(stat_f)
	stat_head_format = workbook.add_format(stat_head)
	stat_format = workbook.add_format(stat)
	stat_int_format = workbook.add_format(stat_int)
	stat_int_f_format = workbook.add_format(stat_int_f)
	stat_head_index1_format = workbook.add_format(stat_head_index1)
	stat_head_index2_format = workbook.add_format(stat_head_index2)
	stat_head_streak_format = workbook.add_format(stat_head_streak)

	win_format = workbook.add_format({'bold': True, 'italic': True})

	index_format_dict = {
		'wins': stat_head_index1_format,
		'win%': stat_head_index1_format,
		'ppg': stat_head_index1_format,
		'streak': stat_head_streak_format,
		'H': stat_head_index2_format,
		'A': stat_head_index2_format,
		'overall': overall_stat_format,
		'Current': stat_head_index2_format,
		'Best': stat_head_index2_format
	}

	stat_format_dict = {
		'wins': {
			'H': {
				'F': stat_int_f_format,
				'K': stat_int_format,
				},
			'A': {
				'F': stat_int_f_format,
				'K': stat_int_format,
			},
			'overall': {
				'F': overall_wins_f_format,
				'K': overall_wins_format,
			}
		},
		'win%': {
			'H': {
				'F': stat_f_format,
				'K': stat_format,
				},
			'A': {
				'F': stat_f_format,
				'K': stat_format,
			},
			'overall': {
				'F': overall_stat_f_format,
				'K': overall_stat_format,
			}
		},
		'streak': {
			'Current':{
				'F': stat_int_f_format,
				'K': stat_int_format
			},
			'Best': {
				'F': stat_int_f_format,
				'K': stat_int_format
			}
		}
	}

	stat_format_dict['ppg'] = stat_format_dict['win%']

	# format the head
	for col_num, value in enumerate(df.columns.values):
		# worksheet.write(row, col, value, format)
		#print("col_num: ", col_num, " value: ", value)
		worksheet.write(0, col_num + 1, value, head_format)

	for i,x in enumerate(df.index.names):
		worksheet.write(0, i, x, head_format)

	for i in range(1, game_rows+1):

		x = i-1
		d = df.index[x]
		f, k, s, t = df.iloc[x]

		# light blue fill on every other row
		if i%2 == 0:
			g_format = odd_game_format
			dt_format = odd_date_format
			s_format = odd_side_format
			t_format = odd_time_format

		else:
			g_format = game_format
			dt_format = date_format
			s_format = side_format
			t_format = time_format


		worksheet.write_datetime(i, 0, d, dt_format)
		worksheet.write(i, 1, f, g_format)
		worksheet.write(i, 2, k, g_format)
		worksheet.write(i, 3, s, s_format)
		worksheet.write_datetime(i, 4, t, t_format)

	# set the width of the index (date) column
	worksheet.set_column(0, 0, 15)

	# emboldens/italicizes each win in the game rows using conditional formatting
	for i in range(2, game_rows+2):
		x = i-2
		f = df.iloc[x, 0]
		k = df.iloc[x, 1]
		worksheet.conditional_format('B{}:B{}'.format(i,i), {'type':'cell','criteria':'>','value':k,'format': win_format})
		worksheet.conditional_format('C{}:C{}'.format(i,i), {'type':'cell','criteria':'>','value':f,'format': win_format})

	#formatting stats head
	for i,v in enumerate(dfstats.columns.values):
		x = i+2
		worksheet.write(stats_head_row, x, v, stat_head_format)

	#formatting stats index and 'table''
	for i, v in enumerate(dfstats.index):
		x = win_index+i
		ind1,ind2 = v
		f, k = dfstats.iloc[i]

		worksheet.write(x, 0, ind1, index_format_dict[ind1])
		worksheet.write(x, 1, ind2, index_format_dict[ind2])
		worksheet.write(x, 2, f, stat_format_dict[ind1][ind2]['F'])
		worksheet.write(x, 3, k, stat_format_dict[ind1][ind2]['K'])

	writer.save()


main()
