import pandas as pd
import numpy as np

multi_ind = [
  np.array(["wins", "wins", "wins", "win%", "win%", "win%", "ppg", "ppg", "ppg", "streak", "streak"]),
  np.array(["H", "A", "overall", "H", "A", "overall", "H", "A", "overall", "Current", "Best"])
  ]

dfstats = pd.DataFrame(
	{
	'F':[
		6, 5, 11,
		0, 50, 45,
		13.1, 12.4, 12.85,
		0, 7
		],
	'K': [
		4, 8, 12,
		40, 60, 55,
		12.1, 13.7, 12.9,
		3, 6
		]
	}, index=multi_ind)


12/15/18	9	15	A	11:02 AM

days_of_week = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
for row in dfnew.itertuples(index=False):
	print(row)
	biggest = max(row.Fritz, row.Ken)
	row[row.index(biggest)]
	row._fields[row.index(biggest)]
	row.Side
	days_of_week[row.Time.weekday()]
# BEGIN DRAFT #

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

for i,v in enumerate(dfnew.sum(axis=1)):
    time = v/total_points*break_length
    time_deltas.append(datetime.timedelta(seconds=round(time.total_seconds())))

time_list = [(tmp_dt + warmup_time + sum(time_deltas[:i], datetime.timedelta())) for i,v in enumerate(time_deltas)]

dfnew['Time'] = time_list
df1 = df1.append(dfnew, sort=False)

f_scores = df1['Fritz'].tolist()
k_scores = df1['Ken'].tolist()
sides = df1['Side'].tolist()

ken = Player()
fritz = Player()


# END DRAFT #



time_list = [(tmp_dt + warmup_time + sum(time_deltas[:i], datetime.timedelta())).time() for i,v in enumerate(time_deltas)]

    sum([(x/total_points)*break_length for x in dfnew.sum(axis=1)][:i], datetime.timedelta())

sum([(x/total_points)*break_length for x in dfnew.sum(axis=1)][:i], datetime.timedelta())

time_list = []
total_points = sum(dfnew.sum(axis=1))


	while more[:1] == 'y':
		f_score = get_score('Fritz')
		k_score = get_score('Ken')
		side = input('Winner side? H/A').upper()
		now = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')
		dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=now)
		df1 = df1.append(dfnew, sort=False)
		more = input('More scores to enter? (Y/N)').lower() or "y"

	f_scores = df1['Fritz'].tolist()
	k_scores = df1['Ken'].tolist()
	sides = df1['Side'].tolist()

	ken = Player()
	fritz = Player()

df1 = pd.read_excel('/Users/ktuten/Desktop/ping_pong/assets/excel/ping_pong_scoresheet_v2.xlsx', parse_dates=['Time'], header=0, skipfooter=13).reset_index(drop=True)



time_list.append((tmp_dt + warmup_time).time())

dfnew = pd.DataFrame({'Fritz':[12,15,11], 'Ken': [15,10,15], 'Side': ['A', 'A', 'A']}, index=[datetime.datetime.today().date(), datetime.datetime.today().date(), datetime.datetime.today().date()])

end_time = datetime.datetime.today().time().strftime('%I:%M %p')



break_length = int(input('Break Length (minutes)'))

datetime.timedelta(minutes=break_length)

start_time = (tmp_dt - datetime.timedelta(minutes=break_length)).time()



def main():

	df1 = pd.read_excel('assets/excel/ping_pong_scoresheet_v2.xlsx', header=0, skipfooter=13).reset_index(drop=True)
	dates = pd.DatetimeIndex(df1.Date.dt.date, name='Date')
	start_games = df1.shape[0]
	df1.index = dates
	df1.drop(columns='Date', inplace=True)

	f_score = get_score('Fritz')
	k_score = get_score('Ken')
	side = input('Winner side? H/A').upper()

	while ((side != 'A') and (side != 'H')):
		side = input('Winner side? H/A').upper()

	now = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')

	dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=now)
	df1 = df1.append(dfnew, sort=False)
	more = input('More scores to enter? (Y/N)').lower() or "y"

	while more[:1] == 'y':
		f_score = get_score('Fritz')
		k_score = get_score('Ken')
		side = input('Winner side? H/A').upper()
		now = pd.DatetimeIndex([pd.datetime.today().date()], name='Date')
		dfnew = pd.DataFrame({'Fritz':[f_score], 'Ken': [k_score], 'Side': [side]}, index=now)
		df1 = df1.append(dfnew, sort=False)
		more = input('More scores to enter? (Y/N)').lower() or "y"

	f_scores = df1['Fritz'].tolist()
	k_scores = df1['Ken'].tolist()
	sides = df1['Side'].tolist()

	ken = Player()
	fritz = Player()

	for _,x in enumerate(zip(f_scores, k_scores, sides)):
			f = x[0]
			k = x[1]
			side = x[2]
			if f>k:
				ken.streak.current = 0
				fritz.streak.current = fritz.streak.current+1
				if (fritz.streak.current > fritz.streak.best):
					fritz.streak.best = fritz.streak.current
				if side == 'A':
					fritz.a.wins = fritz.a.wins+1
					fritz.a.points = fritz.a.points+f
					ken.h.losses = ken.h.losses+1
					ken.h.points = ken.h.points+k
				if side == 'H':
					fritz.h.wins = fritz.h.wins+1
					fritz.h.points = fritz.h.points+f
					ken.a.losses = ken.a.losses+1
					ken.a.points = ken.a.points+k

			if k>f:
				fritz.streak.current = 0
				ken.streak.current = ken.streak.current+1
				if (ken.streak.current > ken.streak.best):
					ken.streak.best = ken.streak.current
				if side == 'A':
					ken.a.wins = ken.a.wins+1
					ken.a.points = ken.a.points+k
					fritz.h.losses = fritz.h.losses+1
					fritz.h.points = fritz.h.points+f
				if side == 'H':
					ken.h.wins = ken.h.wins+1
					ken.h.points = ken.h.points+k
					fritz.a.losses = fritz.a.losses+1
					fritz.a.points = fritz.a.points+f


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
	yellow, maroon = '#FED330', '#A9180F'

	base = {
		'font_name': 'AppleGothic',
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
	stat = {**base, 'num_format': '0.000', 'font_size': 12}
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

	head = {**base,
		'font_size': 14,
		'font_color': 'white',
		'bg_color': dark_blue,
		'bold': True,
		'bottom': 1,
		}
	stat_head = {**head, 'valign': 'vcenter'}
	stat_head_column = {**stat_head, 'bottom':13, 'bottom_color': maroon} # F, K
	stat_head_streak = {**stat_head, #streak
		'right': 1,
		'border_color': grey
		}
	stat_head_index1 = {**stat_head_streak, 'bottom': 13} #wins, win%, ppg

	#h, a, current, best
	stat_head_index2 = {**stat_head,
		'align': 'center',
		'bg_color': greyish_blue,
		'border': 1,
		'border_color': grey,
		'font_color': 'black'
		}

	date_format = workbook.add_format(date)
	game_format = workbook.add_format(game)
	head_format = workbook.add_format(head)
	odd_date_format = workbook.add_format(odd_date)
	odd_game_format = workbook.add_format(odd_game)
	odd_side_format = workbook.add_format(odd_side)
	overall_stat_format = workbook.add_format(overall_stat)
    overall_stat_f_format = workbook.add_format(overall_stat_f)
	side_format = workbook.add_format(side)
    stat_f_format = workbook.add_format(stat_f)
	stat_head_format = workbook.add_format(stat_head)
    stat_format = workbook.add_format(stat)
    stat_int_format = workbook.add_format(stat_int)
    stat_int_f_format = workbook.add_format(stat_int_f)
	stat_head_index1_format = workbook.add_format(stat_head_index1)
	stat_head_index2_format = workbook.add_format(stat_head_index2)
    stat_head_streak_format = workbook.add_format(stat_head_streak)

	totals_format = workbook.add_format({
		'top': 1,
		'font_size': 14,
		'font_name': 'AppleGothic',
		'bg_color': '#f2db87',
		'num_format': '#,##0'
		})
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
                'F': overall_stat_f_format,
                'K': overall_stat_format,
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
	worksheet.set_column(0, 0, 20)

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
	for i,v in enumerate(dfstats.index):
		x = win_index+i
		ind_1,ind_2 = dfstats.index[i]
		f = dfstats.iloc[i, 0]
		k = dfstats.iloc[i, 1]


		worksheet.write(x, 1, dfstats.index[i][1], index_format_dict[ind2])
		worksheet.write(x, 2, f, stat_format_dict[ind1][ind2]['F'])
        worksheet.write(x, 3, f, stat_format_dict[ind1][ind2]['K'])
		if i%3==0:
			worksheet.write(x, 0, dfstats.index[i][0], index_format_dict[ind1])


	writer.save()



## NOTE: BEGIN MODELS ###


    class Game(models.Model):
        winner = models.ForeignKey(Player, on_delete=models.CASCADE)
        date =


    class Player(models.Model):

     	NAME_CHOICES = (
    		(FRITZ, 'Fritz'),
    		(KEN, 'Ken'),
    	)

        name = models.CharField(max_length=100)


    class Chart(models.Model):

    	name = models.CharField(max_length=100)
    	chart_type = models.CharField(max_length=25, default='bar')
    	title = models.CharField(max_length=255, default='title')

    	#need same number of Labels, Datas, and BackgroundColors to be created, in the event that they are not present
    	# bord_color_set depends on bg_color_set to determine how many items to create if they dont already exist
    	# labelset and dataset could do same, depending which is called first by the charts template
    	def bg_colors(self):
    		if self.backgroundcolor_set.exists():
    			c = []
    			for color in self.backgroundcolor_set.all():
    				c.append(color.value)
    			return c
    		else:
    			print("No background colors found")

    	def bord_colors(self):
    		# if self.chart_type == 'doughnut':
    		# 	return None

    		if self.bordercolor_set.exists():
    			c = []
    			for color in self.bordercolor_set.all():
    				c.append(color.value)
    			return c
    		else:
    			c = []
    			p = re.compile('^rgba\(\s?(\d{1,3})\,\s?(\d{1,3})\,\s?(\d{1,3}).*\)?$')
    			for color in self.backgroundcolor_set.all():
    				m = p.match(color.value)
    				c.append("rgb( {}, {}, {})".format(m.group(1),m.group(2),m.group(3)))
    			return c
    			#print("No border colors found")

    	def labels(self):
    		if self.label_set.exists():
    			l = []
    			for label in self.label_set.all():
    				l.append(label.name)
    			return l
    		else:
    			print("No labels found")


    	def datas(self):
    		if self.data_set.exists():
    			d = []
    			for data in self.data_set.all():
    				d.append(data.value)
    			return d
    		else:
    			print("No data found")

    	def __str__(self):
    		"""
    		String for representing the Model object.
    		"""
    		return self.name


    # These coincide with the 'labels' seen in a chart. The amount of these attached
    # to a Chart (soon Dataset) should be equal to the amount of Datas (soon Datapoints)
    # attached
    class Label(models.Model):
    	chart = models.ForeignKey(Chart, on_delete=models.CASCADE, default=None)
    	name = models.CharField(max_length=255, default='')

    	def __str__(self):
    		return self.name

    class BackgroundColor(models.Model):

    	#NOTE: (1)
    	#chart = models.ForeignKey(Dataset, on_delete=models.CASCADE, default=None)
    	chart = models.ForeignKey(Chart, on_delete=models.CASCADE)
    	value = models.CharField(max_length=50)
    	def __str__(self):
    		return self.value

    class BorderColor(models.Model):

    	#NOTE: (1)
    	#chart = models.ForeignKey(Dataset, on_delete=models.CASCADE, default=None)
    	chart = models.ForeignKey(Chart, on_delete=models.CASCADE)
    	value = models.CharField(max_length=50)
    	def __str__(self):
    		return self.value

    class Dataset(models.Model):

    	#NOTE: (1)
    	#chart = models.ForeignKey(Chart, on_delete=models.CASCADE)

    	label = models.CharField(max_length=50)
    	chart = models.OneToOneField(Chart, on_delete=models.CASCADE, default=None)
    	def __str__(self):
    		return self.label

    class Data(models.Model):

    	#chart = models.ForeignKey(Dataset, on_delete=models.CASCADE, default=None)
    	chart = models.ForeignKey(Chart, on_delete=models.CASCADE, default=None)
    	#value = models.IntegerField()
    	value = models.FloatField()
    	def __str__(self):
    		return self.value
