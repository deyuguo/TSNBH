import pandas as pd
import functools
Profile = pd.read_excel(r'input\profile.xlsx')
Production_09 = pd.read_table(r'input\2009.txt', skipfooter=15, engine='python')
Production_10 = pd.read_table(r'input\2010.txt', skipfooter=15, engine='python')
Production_11 = pd.read_table(r'input\2011.txt', skipfooter=15, engine='python')
Production_12 = pd.read_table(r'input\2012.txt', skipfooter=15, engine='python')
Production_13 = pd.read_table(r'input\2013.txt', skipfooter=15, engine='python')
Production_14 = pd.read_table(r'input\2014.txt', skipfooter=15, engine='python')
Production_15 = pd.read_table(r'input\2015.txt', skipfooter=15, engine='python')
Production_16 = pd.read_table(r'input\2016.txt', skipfooter=15, engine='python')
Production_17 = pd.read_table(r'input\2017.txt', skipfooter=15, engine='python')
Production_18 = pd.read_table(r'input\2018.txt', skipfooter=27, engine='python')
Production_19 = pd.read_table(r'input\2019.txt', skipfooter=27, engine='python')
Production_20 = pd.read_table(r'input\2020.txt', skipfooter=27, engine='python')
IND_RNs_09 = pd.crosstab(index=Production_09.PROFILE_ID,
                         columns=Production_09.MONTH_SORT,
                         values=Production_09.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_09', margins=True)
IND_RNs_09.rename(columns={1: 'IND_RNs_Jan_09', 2: 'IND_RNs_Feb_09', 3: 'IND_RNs_Mar_09',
                           4: 'IND_RNs_Apr_09', 5: 'IND_RNs_May_09', 6: 'IND_RNs_Jun_09',
                           7: 'IND_RNs_Jul_09', 8: 'IND_RNs_Aug_09', 9: 'IND_RNs_Sep_09',
                           10: 'IND_RNs_Oct_09', 11: 'IND_RNs_Nov_09', 12: 'IND_RNs_Dec_09'
                           }, inplace=True)
IND_Rev_09 = pd.crosstab(index=Production_09.PROFILE_ID,
                         columns=Production_09.MONTH_SORT,
                         values=Production_09.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_09', margins=True)
IND_Rev_09.rename(columns={1: 'IND_Rev_Jan_09', 2: 'IND_Rev_Feb_09', 3: 'IND_Rev_Mar_09',
                           4: 'IND_Rev_Apr_09', 5: 'IND_Rev_May_09', 6: 'IND_Rev_Jun_09',
                           7: 'IND_Rev_Jul_09', 8: 'IND_Rev_Aug_09', 9: 'IND_Rev_Sep_09',
                           10: 'IND_Rev_Oct_09', 11: 'IND_Rev_Nov_09', 12: 'IND_Rev_Dec_09'
                           }, inplace=True)
BLK_RNs_09 = pd.crosstab(index=Production_09.PROFILE_ID,
                         columns=Production_09.MONTH_SORT,
                         values=Production_09.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_09', margins=True)
BLK_RNs_09.rename(columns={1: 'BLK_RNs_Jan_09', 2: 'BLK_RNs_Feb_09', 3: 'BLK_RNs_Mar_09',
                           4: 'BLK_RNs_Apr_09', 5: 'BLK_RNs_May_09', 6: 'BLK_RNs_Jun_09',
                           7: 'BLK_RNs_Jul_09', 8: 'BLK_RNs_Aug_09', 9: 'BLK_RNs_Sep_09',
                           10: 'BLK_RNs_Oct_09', 11: 'BLK_RNs_Nov_09', 12: 'BLK_RNs_Dec_09'
                           }, inplace=True)
BLK_Rev_09 = pd.crosstab(index=Production_09.PROFILE_ID,
                         columns=Production_09.MONTH_SORT,
                         values=Production_09.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_09', margins=True)
BLK_Rev_09.rename(columns={1: 'BLK_Rev_Jan_09', 2: 'BLK_Rev_Feb_09', 3: 'BLK_Rev_Mar_09',
                           4: 'BLK_Rev_Apr_09', 5: 'BLK_Rev_May_09', 6: 'BLK_Rev_Jun_09',
                           7: 'BLK_Rev_Jul_09', 8: 'BLK_Rev_Aug_09', 9: 'BLK_Rev_Sep_09',
                           10: 'BLK_Rev_Oct_09', 11: 'BLK_Rev_Nov_09', 12: 'BLK_Rev_Dec_09'
                           }, inplace=True)
IND_RNs_10 = pd.crosstab(index=Production_10.PROFILE_ID,
                         columns=Production_10.MONTH_SORT,
                         values=Production_10.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_10', margins=True)
IND_RNs_10.rename(columns={1: 'IND_RNs_Jan_10', 2: 'IND_RNs_Feb_10', 3: 'IND_RNs_Mar_10',
                           4: 'IND_RNs_Apr_10', 5: 'IND_RNs_May_10', 6: 'IND_RNs_Jun_10',
                           7: 'IND_RNs_Jul_10', 8: 'IND_RNs_Aug_10', 9: 'IND_RNs_Sep_10',
                           10: 'IND_RNs_Oct_10', 11: 'IND_RNs_Nov_10', 12: 'IND_RNs_Dec_10'
                           }, inplace=True)
IND_Rev_10 = pd.crosstab(index=Production_10.PROFILE_ID,
                         columns=Production_10.MONTH_SORT,
                         values=Production_10.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_10', margins=True)
IND_Rev_10.rename(columns={1: 'IND_Rev_Jan_10', 2: 'IND_Rev_Feb_10', 3: 'IND_Rev_Mar_10',
                           4: 'IND_Rev_Apr_10', 5: 'IND_Rev_May_10', 6: 'IND_Rev_Jun_10',
                           7: 'IND_Rev_Jul_10', 8: 'IND_Rev_Aug_10', 9: 'IND_Rev_Sep_10',
                           10: 'IND_Rev_Oct_10', 11: 'IND_Rev_Nov_10', 12: 'IND_Rev_Dec_10'
                           }, inplace=True)
BLK_RNs_10 = pd.crosstab(index=Production_10.PROFILE_ID,
                         columns=Production_10.MONTH_SORT,
                         values=Production_10.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_10', margins=True)
BLK_RNs_10.rename(columns={1: 'BLK_RNs_Jan_10', 2: 'BLK_RNs_Feb_10', 3: 'BLK_RNs_Mar_10',
                           4: 'BLK_RNs_Apr_10', 5: 'BLK_RNs_May_10', 6: 'BLK_RNs_Jun_10',
                           7: 'BLK_RNs_Jul_10', 8: 'BLK_RNs_Aug_10', 9: 'BLK_RNs_Sep_10',
                           10: 'BLK_RNs_Oct_10', 11: 'BLK_RNs_Nov_10', 12: 'BLK_RNs_Dec_10'
                           }, inplace=True)
BLK_Rev_10 = pd.crosstab(index=Production_10.PROFILE_ID,
                         columns=Production_10.MONTH_SORT,
                         values=Production_10.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_10', margins=True)
BLK_Rev_10.rename(columns={1: 'BLK_Rev_Jan_10', 2: 'BLK_Rev_Feb_10', 3: 'BLK_Rev_Mar_10',
                           4: 'BLK_Rev_Apr_10', 5: 'BLK_Rev_May_10', 6: 'BLK_Rev_Jun_10',
                           7: 'BLK_Rev_Jul_10', 8: 'BLK_Rev_Aug_10', 9: 'BLK_Rev_Sep_10',
                           10: 'BLK_Rev_Oct_10', 11: 'BLK_Rev_Nov_10', 12: 'BLK_Rev_Dec_10'
                           }, inplace=True)
IND_RNs_11 = pd.crosstab(index=Production_11.PROFILE_ID,
                         columns=Production_11.MONTH_SORT,
                         values=Production_11.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_11', margins=True)
IND_RNs_11.rename(columns={1: 'IND_RNs_Jan_11', 2: 'IND_RNs_Feb_11', 3: 'IND_RNs_Mar_11',
                           4: 'IND_RNs_Apr_11', 5: 'IND_RNs_May_11', 6: 'IND_RNs_Jun_11',
                           7: 'IND_RNs_Jul_11', 8: 'IND_RNs_Aug_11', 9: 'IND_RNs_Sep_11',
                           10: 'IND_RNs_Oct_11', 11: 'IND_RNs_Nov_11', 12: 'IND_RNs_Dec_11'
                           }, inplace=True)
IND_Rev_11 = pd.crosstab(index=Production_11.PROFILE_ID,
                         columns=Production_11.MONTH_SORT,
                         values=Production_11.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_11', margins=True)
IND_Rev_11.rename(columns={1: 'IND_Rev_Jan_11', 2: 'IND_Rev_Feb_11', 3: 'IND_Rev_Mar_11',
                           4: 'IND_Rev_Apr_11', 5: 'IND_Rev_May_11', 6: 'IND_Rev_Jun_11',
                           7: 'IND_Rev_Jul_11', 8: 'IND_Rev_Aug_11', 9: 'IND_Rev_Sep_11',
                           10: 'IND_Rev_Oct_11', 11: 'IND_Rev_Nov_11', 12: 'IND_Rev_Dec_11'
                           }, inplace=True)
BLK_RNs_11 = pd.crosstab(index=Production_11.PROFILE_ID,
                         columns=Production_11.MONTH_SORT,
                         values=Production_11.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_11', margins=True)
BLK_RNs_11.rename(columns={1: 'BLK_RNs_Jan_11', 2: 'BLK_RNs_Feb_11', 3: 'BLK_RNs_Mar_11',
                           4: 'BLK_RNs_Apr_11', 5: 'BLK_RNs_May_11', 6: 'BLK_RNs_Jun_11',
                           7: 'BLK_RNs_Jul_11', 8: 'BLK_RNs_Aug_11', 9: 'BLK_RNs_Sep_11',
                           10: 'BLK_RNs_Oct_11', 11: 'BLK_RNs_Nov_11', 12: 'BLK_RNs_Dec_11'
                           }, inplace=True)
BLK_Rev_11 = pd.crosstab(index=Production_11.PROFILE_ID,
                         columns=Production_11.MONTH_SORT,
                         values=Production_11.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_11', margins=True)
BLK_Rev_11.rename(columns={1: 'BLK_Rev_Jan_11', 2: 'BLK_Rev_Feb_11', 3: 'BLK_Rev_Mar_11',
                           4: 'BLK_Rev_Apr_11', 5: 'BLK_Rev_May_11', 6: 'BLK_Rev_Jun_11',
                           7: 'BLK_Rev_Jul_11', 8: 'BLK_Rev_Aug_11', 9: 'BLK_Rev_Sep_11',
                           10: 'BLK_Rev_Oct_11', 11: 'BLK_Rev_Nov_11', 12: 'BLK_Rev_Dec_11'
                           }, inplace=True)
IND_RNs_12 = pd.crosstab(index=Production_12.PROFILE_ID,
                         columns=Production_12.MONTH_SORT,
                         values=Production_12.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_12', margins=True)
IND_RNs_12.rename(columns={1: 'IND_RNs_Jan_12', 2: 'IND_RNs_Feb_12', 3: 'IND_RNs_Mar_12',
                           4: 'IND_RNs_Apr_12', 5: 'IND_RNs_May_12', 6: 'IND_RNs_Jun_12',
                           7: 'IND_RNs_Jul_12', 8: 'IND_RNs_Aug_12', 9: 'IND_RNs_Sep_12',
                           10: 'IND_RNs_Oct_12', 11: 'IND_RNs_Nov_12', 12: 'IND_RNs_Dec_12'
                           }, inplace=True)
IND_Rev_12 = pd.crosstab(index=Production_12.PROFILE_ID,
                         columns=Production_12.MONTH_SORT,
                         values=Production_12.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_12', margins=True)
IND_Rev_12.rename(columns={1: 'IND_Rev_Jan_12', 2: 'IND_Rev_Feb_12', 3: 'IND_Rev_Mar_12',
                           4: 'IND_Rev_Apr_12', 5: 'IND_Rev_May_12', 6: 'IND_Rev_Jun_12',
                           7: 'IND_Rev_Jul_12', 8: 'IND_Rev_Aug_12', 9: 'IND_Rev_Sep_12',
                           10: 'IND_Rev_Oct_12', 11: 'IND_Rev_Nov_12', 12: 'IND_Rev_Dec_12'
                           }, inplace=True)
BLK_RNs_12 = pd.crosstab(index=Production_12.PROFILE_ID,
                         columns=Production_12.MONTH_SORT,
                         values=Production_12.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_12', margins=True)
BLK_RNs_12.rename(columns={1: 'BLK_RNs_Jan_12', 2: 'BLK_RNs_Feb_12', 3: 'BLK_RNs_Mar_12',
                           4: 'BLK_RNs_Apr_12', 5: 'BLK_RNs_May_12', 6: 'BLK_RNs_Jun_12',
                           7: 'BLK_RNs_Jul_12', 8: 'BLK_RNs_Aug_12', 9: 'BLK_RNs_Sep_12',
                           10: 'BLK_RNs_Oct_12', 11: 'BLK_RNs_Nov_12', 12: 'BLK_RNs_Dec_12'
                           }, inplace=True)
BLK_Rev_12 = pd.crosstab(index=Production_12.PROFILE_ID,
                         columns=Production_12.MONTH_SORT,
                         values=Production_12.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_12', margins=True)
BLK_Rev_12.rename(columns={1: 'BLK_Rev_Jan_12', 2: 'BLK_Rev_Feb_12', 3: 'BLK_Rev_Mar_12',
                           4: 'BLK_Rev_Apr_12', 5: 'BLK_Rev_May_12', 6: 'BLK_Rev_Jun_12',
                           7: 'BLK_Rev_Jul_12', 8: 'BLK_Rev_Aug_12', 9: 'BLK_Rev_Sep_12',
                           10: 'BLK_Rev_Oct_12', 11: 'BLK_Rev_Nov_12', 12: 'BLK_Rev_Dec_12'
                           }, inplace=True)
IND_RNs_13 = pd.crosstab(index=Production_13.PROFILE_ID,
                         columns=Production_13.MONTH_SORT,
                         values=Production_13.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_13', margins=True)
IND_RNs_13.rename(columns={1: 'IND_RNs_Jan_13', 2: 'IND_RNs_Feb_13', 3: 'IND_RNs_Mar_13',
                           4: 'IND_RNs_Apr_13', 5: 'IND_RNs_May_13', 6: 'IND_RNs_Jun_13',
                           7: 'IND_RNs_Jul_13', 8: 'IND_RNs_Aug_13', 9: 'IND_RNs_Sep_13',
                           10: 'IND_RNs_Oct_13', 11: 'IND_RNs_Nov_13', 12: 'IND_RNs_Dec_13'
                           }, inplace=True)
IND_Rev_13 = pd.crosstab(index=Production_13.PROFILE_ID,
                         columns=Production_13.MONTH_SORT,
                         values=Production_13.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_13', margins=True)
IND_Rev_13.rename(columns={1: 'IND_Rev_Jan_13', 2: 'IND_Rev_Feb_13', 3: 'IND_Rev_Mar_13',
                           4: 'IND_Rev_Apr_13', 5: 'IND_Rev_May_13', 6: 'IND_Rev_Jun_13',
                           7: 'IND_Rev_Jul_13', 8: 'IND_Rev_Aug_13', 9: 'IND_Rev_Sep_13',
                           10: 'IND_Rev_Oct_13', 11: 'IND_Rev_Nov_13', 12: 'IND_Rev_Dec_13'
                           }, inplace=True)
BLK_RNs_13 = pd.crosstab(index=Production_13.PROFILE_ID,
                         columns=Production_13.MONTH_SORT,
                         values=Production_13.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_13', margins=True)
BLK_RNs_13.rename(columns={1: 'BLK_RNs_Jan_13', 2: 'BLK_RNs_Feb_13', 3: 'BLK_RNs_Mar_13',
                           4: 'BLK_RNs_Apr_13', 5: 'BLK_RNs_May_13', 6: 'BLK_RNs_Jun_13',
                           7: 'BLK_RNs_Jul_13', 8: 'BLK_RNs_Aug_13', 9: 'BLK_RNs_Sep_13',
                           10: 'BLK_RNs_Oct_13', 11: 'BLK_RNs_Nov_13', 12: 'BLK_RNs_Dec_13'
                           }, inplace=True)
BLK_Rev_13 = pd.crosstab(index=Production_13.PROFILE_ID,
                         columns=Production_13.MONTH_SORT,
                         values=Production_13.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_13', margins=True)
BLK_Rev_13.rename(columns={1: 'BLK_Rev_Jan_13', 2: 'BLK_Rev_Feb_13', 3: 'BLK_Rev_Mar_13',
                           4: 'BLK_Rev_Apr_13', 5: 'BLK_Rev_May_13', 6: 'BLK_Rev_Jun_13',
                           7: 'BLK_Rev_Jul_13', 8: 'BLK_Rev_Aug_13', 9: 'BLK_Rev_Sep_13',
                           10: 'BLK_Rev_Oct_13', 11: 'BLK_Rev_Nov_13', 12: 'BLK_Rev_Dec_13'
                           }, inplace=True)
IND_RNs_14 = pd.crosstab(index=Production_14.PROFILE_ID,
                         columns=Production_14.MONTH_SORT,
                         values=Production_14.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_14', margins=True)
IND_RNs_14.rename(columns={1: 'IND_RNs_Jan_14', 2: 'IND_RNs_Feb_14', 3: 'IND_RNs_Mar_14',
                           4: 'IND_RNs_Apr_14', 5: 'IND_RNs_May_14', 6: 'IND_RNs_Jun_14',
                           7: 'IND_RNs_Jul_14', 8: 'IND_RNs_Aug_14', 9: 'IND_RNs_Sep_14',
                           10: 'IND_RNs_Oct_14', 11: 'IND_RNs_Nov_14', 12: 'IND_RNs_Dec_14'
                           }, inplace=True)
IND_Rev_14 = pd.crosstab(index=Production_14.PROFILE_ID,
                         columns=Production_14.MONTH_SORT,
                         values=Production_14.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_14', margins=True)
IND_Rev_14.rename(columns={1: 'IND_Rev_Jan_14', 2: 'IND_Rev_Feb_14', 3: 'IND_Rev_Mar_14',
                           4: 'IND_Rev_Apr_14', 5: 'IND_Rev_May_14', 6: 'IND_Rev_Jun_14',
                           7: 'IND_Rev_Jul_14', 8: 'IND_Rev_Aug_14', 9: 'IND_Rev_Sep_14',
                           10: 'IND_Rev_Oct_14', 11: 'IND_Rev_Nov_14', 12: 'IND_Rev_Dec_14'
                           }, inplace=True)
BLK_RNs_14 = pd.crosstab(index=Production_14.PROFILE_ID,
                         columns=Production_14.MONTH_SORT,
                         values=Production_14.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_14', margins=True)
BLK_RNs_14.rename(columns={1: 'BLK_RNs_Jan_14', 2: 'BLK_RNs_Feb_14', 3: 'BLK_RNs_Mar_14',
                           4: 'BLK_RNs_Apr_14', 5: 'BLK_RNs_May_14', 6: 'BLK_RNs_Jun_14',
                           7: 'BLK_RNs_Jul_14', 8: 'BLK_RNs_Aug_14', 9: 'BLK_RNs_Sep_14',
                           10: 'BLK_RNs_Oct_14', 11: 'BLK_RNs_Nov_14', 12: 'BLK_RNs_Dec_14'
                           }, inplace=True)
BLK_Rev_14 = pd.crosstab(index=Production_14.PROFILE_ID,
                         columns=Production_14.MONTH_SORT,
                         values=Production_14.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_14', margins=True)
BLK_Rev_14.rename(columns={1: 'BLK_Rev_Jan_14', 2: 'BLK_Rev_Feb_14', 3: 'BLK_Rev_Mar_14',
                           4: 'BLK_Rev_Apr_14', 5: 'BLK_Rev_May_14', 6: 'BLK_Rev_Jun_14',
                           7: 'BLK_Rev_Jul_14', 8: 'BLK_Rev_Aug_14', 9: 'BLK_Rev_Sep_14',
                           10: 'BLK_Rev_Oct_14', 11: 'BLK_Rev_Nov_14', 12: 'BLK_Rev_Dec_14'
                           }, inplace=True)
IND_RNs_15 = pd.crosstab(index=Production_15.PROFILE_ID,
                         columns=Production_15.MONTH_SORT,
                         values=Production_15.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_15', margins=True)
IND_RNs_15.rename(columns={1: 'IND_RNs_Jan_15', 2: 'IND_RNs_Feb_15', 3: 'IND_RNs_Mar_15',
                           4: 'IND_RNs_Apr_15', 5: 'IND_RNs_May_15', 6: 'IND_RNs_Jun_15',
                           7: 'IND_RNs_Jul_15', 8: 'IND_RNs_Aug_15', 9: 'IND_RNs_Sep_15',
                           10: 'IND_RNs_Oct_15', 11: 'IND_RNs_Nov_15', 12: 'IND_RNs_Dec_15'
                           }, inplace=True)
IND_Rev_15 = pd.crosstab(index=Production_15.PROFILE_ID,
                         columns=Production_15.MONTH_SORT,
                         values=Production_15.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_15', margins=True)
IND_Rev_15.rename(columns={1: 'IND_Rev_Jan_15', 2: 'IND_Rev_Feb_15', 3: 'IND_Rev_Mar_15',
                           4: 'IND_Rev_Apr_15', 5: 'IND_Rev_May_15', 6: 'IND_Rev_Jun_15',
                           7: 'IND_Rev_Jul_15', 8: 'IND_Rev_Aug_15', 9: 'IND_Rev_Sep_15',
                           10: 'IND_Rev_Oct_15', 11: 'IND_Rev_Nov_15', 12: 'IND_Rev_Dec_15'
                           }, inplace=True)
BLK_RNs_15 = pd.crosstab(index=Production_15.PROFILE_ID,
                         columns=Production_15.MONTH_SORT,
                         values=Production_15.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_15', margins=True)
BLK_RNs_15.rename(columns={1: 'BLK_RNs_Jan_15', 2: 'BLK_RNs_Feb_15', 3: 'BLK_RNs_Mar_15',
                           4: 'BLK_RNs_Apr_15', 5: 'BLK_RNs_May_15', 6: 'BLK_RNs_Jun_15',
                           7: 'BLK_RNs_Jul_15', 8: 'BLK_RNs_Aug_15', 9: 'BLK_RNs_Sep_15',
                           10: 'BLK_RNs_Oct_15', 11: 'BLK_RNs_Nov_15', 12: 'BLK_RNs_Dec_15'
                           }, inplace=True)
BLK_Rev_15 = pd.crosstab(index=Production_15.PROFILE_ID,
                         columns=Production_15.MONTH_SORT,
                         values=Production_15.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_15', margins=True)
BLK_Rev_15.rename(columns={1: 'BLK_Rev_Jan_15', 2: 'BLK_Rev_Feb_15', 3: 'BLK_Rev_Mar_15',
                           4: 'BLK_Rev_Apr_15', 5: 'BLK_Rev_May_15', 6: 'BLK_Rev_Jun_15',
                           7: 'BLK_Rev_Jul_15', 8: 'BLK_Rev_Aug_15', 9: 'BLK_Rev_Sep_15',
                           10: 'BLK_Rev_Oct_15', 11: 'BLK_Rev_Nov_15', 12: 'BLK_Rev_Dec_15'
                           }, inplace=True)
IND_RNs_16 = pd.crosstab(index=Production_16.PROFILE_ID,
                         columns=Production_16.MONTH_SORT,
                         values=Production_16.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_16', margins=True)
IND_RNs_16.rename(columns={1: 'IND_RNs_Jan_16', 2: 'IND_RNs_Feb_16', 3: 'IND_RNs_Mar_16',
                           4: 'IND_RNs_Apr_16', 5: 'IND_RNs_May_16', 6: 'IND_RNs_Jun_16',
                           7: 'IND_RNs_Jul_16', 8: 'IND_RNs_Aug_16', 9: 'IND_RNs_Sep_16',
                           10: 'IND_RNs_Oct_16', 11: 'IND_RNs_Nov_16', 12: 'IND_RNs_Dec_16'
                           }, inplace=True)
IND_Rev_16 = pd.crosstab(index=Production_16.PROFILE_ID,
                         columns=Production_16.MONTH_SORT,
                         values=Production_16.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_16', margins=True)
IND_Rev_16.rename(columns={1: 'IND_Rev_Jan_16', 2: 'IND_Rev_Feb_16', 3: 'IND_Rev_Mar_16',
                           4: 'IND_Rev_Apr_16', 5: 'IND_Rev_May_16', 6: 'IND_Rev_Jun_16',
                           7: 'IND_Rev_Jul_16', 8: 'IND_Rev_Aug_16', 9: 'IND_Rev_Sep_16',
                           10: 'IND_Rev_Oct_16', 11: 'IND_Rev_Nov_16', 12: 'IND_Rev_Dec_16'
                           }, inplace=True)
BLK_RNs_16 = pd.crosstab(index=Production_16.PROFILE_ID,
                         columns=Production_16.MONTH_SORT,
                         values=Production_16.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_16', margins=True)
BLK_RNs_16.rename(columns={1: 'BLK_RNs_Jan_16', 2: 'BLK_RNs_Feb_16', 3: 'BLK_RNs_Mar_16',
                           4: 'BLK_RNs_Apr_16', 5: 'BLK_RNs_May_16', 6: 'BLK_RNs_Jun_16',
                           7: 'BLK_RNs_Jul_16', 8: 'BLK_RNs_Aug_16', 9: 'BLK_RNs_Sep_16',
                           10: 'BLK_RNs_Oct_16', 11: 'BLK_RNs_Nov_16', 12: 'BLK_RNs_Dec_16'
                           }, inplace=True)
BLK_Rev_16 = pd.crosstab(index=Production_16.PROFILE_ID,
                         columns=Production_16.MONTH_SORT,
                         values=Production_16.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_16', margins=True)
BLK_Rev_16.rename(columns={1: 'BLK_Rev_Jan_16', 2: 'BLK_Rev_Feb_16', 3: 'BLK_Rev_Mar_16',
                           4: 'BLK_Rev_Apr_16', 5: 'BLK_Rev_May_16', 6: 'BLK_Rev_Jun_16',
                           7: 'BLK_Rev_Jul_16', 8: 'BLK_Rev_Aug_16', 9: 'BLK_Rev_Sep_16',
                           10: 'BLK_Rev_Oct_16', 11: 'BLK_Rev_Nov_16', 12: 'BLK_Rev_Dec_16'
                           }, inplace=True)
IND_RNs_17 = pd.crosstab(index=Production_17.PROFILE_ID,
                         columns=Production_17.MONTH_SORT,
                         values=Production_17.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_17', margins=True)
IND_RNs_17.rename(columns={1: 'IND_RNs_Jan_17', 2: 'IND_RNs_Feb_17', 3: 'IND_RNs_Mar_17',
                           4: 'IND_RNs_Apr_17', 5: 'IND_RNs_May_17', 6: 'IND_RNs_Jun_17',
                           7: 'IND_RNs_Jul_17', 8: 'IND_RNs_Aug_17', 9: 'IND_RNs_Sep_17',
                           10: 'IND_RNs_Oct_17', 11: 'IND_RNs_Nov_17', 12: 'IND_RNs_Dec_17'
                           }, inplace=True)
IND_Rev_17 = pd.crosstab(index=Production_17.PROFILE_ID,
                         columns=Production_17.MONTH_SORT,
                         values=Production_17.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_17', margins=True)
IND_Rev_17.rename(columns={1: 'IND_Rev_Jan_17', 2: 'IND_Rev_Feb_17', 3: 'IND_Rev_Mar_17',
                           4: 'IND_Rev_Apr_17', 5: 'IND_Rev_May_17', 6: 'IND_Rev_Jun_17',
                           7: 'IND_Rev_Jul_17', 8: 'IND_Rev_Aug_17', 9: 'IND_Rev_Sep_17',
                           10: 'IND_Rev_Oct_17', 11: 'IND_Rev_Nov_17', 12: 'IND_Rev_Dec_17'
                           }, inplace=True)
BLK_RNs_17 = pd.crosstab(index=Production_17.PROFILE_ID,
                         columns=Production_17.MONTH_SORT,
                         values=Production_17.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_17', margins=True)
BLK_RNs_17.rename(columns={1: 'BLK_RNs_Jan_17', 2: 'BLK_RNs_Feb_17', 3: 'BLK_RNs_Mar_17',
                           4: 'BLK_RNs_Apr_17', 5: 'BLK_RNs_May_17', 6: 'BLK_RNs_Jun_17',
                           7: 'BLK_RNs_Jul_17', 8: 'BLK_RNs_Aug_17', 9: 'BLK_RNs_Sep_17',
                           10: 'BLK_RNs_Oct_17', 11: 'BLK_RNs_Nov_17', 12: 'BLK_RNs_Dec_17'
                           }, inplace=True)
BLK_Rev_17 = pd.crosstab(index=Production_17.PROFILE_ID,
                         columns=Production_17.MONTH_SORT,
                         values=Production_17.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_17', margins=True)
BLK_Rev_17.rename(columns={1: 'BLK_Rev_Jan_17', 2: 'BLK_Rev_Feb_17', 3: 'BLK_Rev_Mar_17',
                           4: 'BLK_Rev_Apr_17', 5: 'BLK_Rev_May_17', 6: 'BLK_Rev_Jun_17',
                           7: 'BLK_Rev_Jul_17', 8: 'BLK_Rev_Aug_17', 9: 'BLK_Rev_Sep_17',
                           10: 'BLK_Rev_Oct_17', 11: 'BLK_Rev_Nov_17', 12: 'BLK_Rev_Dec_17'
                           }, inplace=True)
IND_RNs_18 = pd.crosstab(index=Production_18.PROFILE_ID,
                         columns=Production_18.MONTH_SORT,
                         values=Production_18.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_18', margins=True)
IND_RNs_18.rename(columns={1: 'IND_RNs_Jan_18', 2: 'IND_RNs_Feb_18', 3: 'IND_RNs_Mar_18',
                           4: 'IND_RNs_Apr_18', 5: 'IND_RNs_May_18', 6: 'IND_RNs_Jun_18',
                           7: 'IND_RNs_Jul_18', 8: 'IND_RNs_Aug_18', 9: 'IND_RNs_Sep_18',
                           10: 'IND_RNs_Oct_18', 11: 'IND_RNs_Nov_18', 12: 'IND_RNs_Dec_18'
                           }, inplace=True)
IND_Rev_18 = pd.crosstab(index=Production_18.PROFILE_ID,
                         columns=Production_18.MONTH_SORT,
                         values=Production_18.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_18', margins=True)
IND_Rev_18.rename(columns={1: 'IND_Rev_Jan_18', 2: 'IND_Rev_Feb_18', 3: 'IND_Rev_Mar_18',
                           4: 'IND_Rev_Apr_18', 5: 'IND_Rev_May_18', 6: 'IND_Rev_Jun_18',
                           7: 'IND_Rev_Jul_18', 8: 'IND_Rev_Aug_18', 9: 'IND_Rev_Sep_18',
                           10: 'IND_Rev_Oct_18', 11: 'IND_Rev_Nov_18', 12: 'IND_Rev_Dec_18'
                           }, inplace=True)
BLK_RNs_18 = pd.crosstab(index=Production_18.PROFILE_ID,
                         columns=Production_18.MONTH_SORT,
                         values=Production_18.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_18', margins=True)
BLK_RNs_18.rename(columns={1: 'BLK_RNs_Jan_18', 2: 'BLK_RNs_Feb_18', 3: 'BLK_RNs_Mar_18',
                           4: 'BLK_RNs_Apr_18', 5: 'BLK_RNs_May_18', 6: 'BLK_RNs_Jun_18',
                           7: 'BLK_RNs_Jul_18', 8: 'BLK_RNs_Aug_18', 9: 'BLK_RNs_Sep_18',
                           10: 'BLK_RNs_Oct_18', 11: 'BLK_RNs_Nov_18', 12: 'BLK_RNs_Dec_18'
                           }, inplace=True)
BLK_Rev_18 = pd.crosstab(index=Production_18.PROFILE_ID,
                         columns=Production_18.MONTH_SORT,
                         values=Production_18.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_18', margins=True)
BLK_Rev_18.rename(columns={1: 'BLK_Rev_Jan_18', 2: 'BLK_Rev_Feb_18', 3: 'BLK_Rev_Mar_18',
                           4: 'BLK_Rev_Apr_18', 5: 'BLK_Rev_May_18', 6: 'BLK_Rev_Jun_18',
                           7: 'BLK_Rev_Jul_18', 8: 'BLK_Rev_Aug_18', 9: 'BLK_Rev_Sep_18',
                           10: 'BLK_Rev_Oct_18', 11: 'BLK_Rev_Nov_18', 12: 'BLK_Rev_Dec_18'
                           }, inplace=True)
IND_RNs_19 = pd.crosstab(index=Production_19.PROFILE_ID,
                         columns=Production_19.MONTH_SORT,
                         values=Production_19.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_19', margins=True)
IND_RNs_19.rename(columns={1: 'IND_RNs_Jan_19', 2: 'IND_RNs_Feb_19', 3: 'IND_RNs_Mar_19',
                           4: 'IND_RNs_Apr_19', 5: 'IND_RNs_May_19', 6: 'IND_RNs_Jun_19',
                           7: 'IND_RNs_Jul_19', 8: 'IND_RNs_Aug_19', 9: 'IND_RNs_Sep_19',
                           10: 'IND_RNs_Oct_19', 11: 'IND_RNs_Nov_19', 12: 'IND_RNs_Dec_19'
                           }, inplace=True)
IND_Rev_19 = pd.crosstab(index=Production_19.PROFILE_ID,
                         columns=Production_19.MONTH_SORT,
                         values=Production_19.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_19', margins=True)
IND_Rev_19.rename(columns={1: 'IND_Rev_Jan_19', 2: 'IND_Rev_Feb_19', 3: 'IND_Rev_Mar_19',
                           4: 'IND_Rev_Apr_19', 5: 'IND_Rev_May_19', 6: 'IND_Rev_Jun_19',
                           7: 'IND_Rev_Jul_19', 8: 'IND_Rev_Aug_19', 9: 'IND_Rev_Sep_19',
                           10: 'IND_Rev_Oct_19', 11: 'IND_Rev_Nov_19', 12: 'IND_Rev_Dec_19'
                           }, inplace=True)
BLK_RNs_19 = pd.crosstab(index=Production_19.PROFILE_ID,
                         columns=Production_19.MONTH_SORT,
                         values=Production_19.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_19', margins=True)
BLK_RNs_19.rename(columns={1: 'BLK_RNs_Jan_19', 2: 'BLK_RNs_Feb_19', 3: 'BLK_RNs_Mar_19',
                           4: 'BLK_RNs_Apr_19', 5: 'BLK_RNs_May_19', 6: 'BLK_RNs_Jun_19',
                           7: 'BLK_RNs_Jul_19', 8: 'BLK_RNs_Aug_19', 9: 'BLK_RNs_Sep_19',
                           10: 'BLK_RNs_Oct_19', 11: 'BLK_RNs_Nov_19', 12: 'BLK_RNs_Dec_19'
                           }, inplace=True)
BLK_Rev_19 = pd.crosstab(index=Production_19.PROFILE_ID,
                         columns=Production_19.MONTH_SORT,
                         values=Production_19.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_19', margins=True)
BLK_Rev_19.rename(columns={1: 'BLK_Rev_Jan_19', 2: 'BLK_Rev_Feb_19', 3: 'BLK_Rev_Mar_19',
                           4: 'BLK_Rev_Apr_19', 5: 'BLK_Rev_May_19', 6: 'BLK_Rev_Jun_19',
                           7: 'BLK_Rev_Jul_19', 8: 'BLK_Rev_Aug_19', 9: 'BLK_Rev_Sep_19',
                           10: 'BLK_Rev_Oct_19', 11: 'BLK_Rev_Nov_19', 12: 'BLK_Rev_Dec_19'
                           }, inplace=True)
IND_RNs_20 = pd.crosstab(index=Production_20.PROFILE_ID,
                         columns=Production_20.MONTH_SORT,
                         values=Production_20.IND_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='IND_RNs_FY_20', margins=True)
IND_RNs_20.rename(columns={1: 'IND_RNs_Jan_20', 2: 'IND_RNs_Feb_20', 3: 'IND_RNs_Mar_20',
                           4: 'IND_RNs_Apr_20', 5: 'IND_RNs_May_20', 6: 'IND_RNs_Jun_20',
                           7: 'IND_RNs_Jul_20', 8: 'IND_RNs_Aug_20', 9: 'IND_RNs_Sep_20',
                           10: 'IND_RNs_Oct_20', 11: 'IND_RNs_Nov_20', 12: 'IND_RNs_Dec_20'
                           }, inplace=True)
IND_Rev_20 = pd.crosstab(index=Production_20.PROFILE_ID,
                         columns=Production_20.MONTH_SORT,
                         values=Production_20.IND_ROOM_REVENUE,
                         aggfunc=sum, margins_name='IND_Rev_FY_20', margins=True)
IND_Rev_20.rename(columns={1: 'IND_Rev_Jan_20', 2: 'IND_Rev_Feb_20', 3: 'IND_Rev_Mar_20',
                           4: 'IND_Rev_Apr_20', 5: 'IND_Rev_May_20', 6: 'IND_Rev_Jun_20',
                           7: 'IND_Rev_Jul_20', 8: 'IND_Rev_Aug_20', 9: 'IND_Rev_Sep_20',
                           10: 'IND_Rev_Oct_20', 11: 'IND_Rev_Nov_20', 12: 'IND_Rev_Dec_20'
                           }, inplace=True)
BLK_RNs_20 = pd.crosstab(index=Production_20.PROFILE_ID,
                         columns=Production_20.MONTH_SORT,
                         values=Production_20.BLK_ROOM_NIGHTS,
                         aggfunc=sum, margins_name='BLK_RNs_FY_20', margins=True)
BLK_RNs_20.rename(columns={1: 'BLK_RNs_Jan_20', 2: 'BLK_RNs_Feb_20', 3: 'BLK_RNs_Mar_20',
                           4: 'BLK_RNs_Apr_20', 5: 'BLK_RNs_May_20', 6: 'BLK_RNs_Jun_20',
                           7: 'BLK_RNs_Jul_20', 8: 'BLK_RNs_Aug_20', 9: 'BLK_RNs_Sep_20',
                           10: 'BLK_RNs_Oct_20', 11: 'BLK_RNs_Nov_20', 12: 'BLK_RNs_Dec_20'
                           }, inplace=True)
BLK_Rev_20 = pd.crosstab(index=Production_20.PROFILE_ID,
                         columns=Production_20.MONTH_SORT,
                         values=Production_20.BLK_ROOM_REVENUE,
                         aggfunc=sum, margins_name='BLK_Rev_FY_20', margins=True)
BLK_Rev_20.rename(columns={1: 'BLK_Rev_Jan_20', 2: 'BLK_Rev_Feb_20', 3: 'BLK_Rev_Mar_20',
                           4: 'BLK_Rev_Apr_20', 5: 'BLK_Rev_May_20', 6: 'BLK_Rev_Jun_20',
                           7: 'BLK_Rev_Jul_20', 8: 'BLK_Rev_Aug_20', 9: 'BLK_Rev_Sep_20',
                           10: 'BLK_Rev_Oct_20', 11: 'BLK_Rev_Nov_20', 12: 'BLK_Rev_Dec_20'
                           }, inplace=True)
PRO_IN_1 = [IND_RNs_09, IND_Rev_09, BLK_RNs_09, BLK_Rev_09,
            IND_RNs_10, IND_Rev_10, BLK_RNs_10, BLK_Rev_10,
            IND_RNs_11, IND_Rev_11, BLK_RNs_11, BLK_Rev_11,
            IND_RNs_12, IND_Rev_12, BLK_RNs_12, BLK_Rev_12,
            IND_RNs_13, IND_Rev_13, BLK_RNs_13, BLK_Rev_13,
            IND_RNs_14, IND_Rev_14, BLK_RNs_14, BLK_Rev_14,
            IND_RNs_15, IND_Rev_15, BLK_RNs_15, BLK_Rev_15,
            IND_RNs_16, IND_Rev_16, BLK_RNs_16, BLK_Rev_16,
            IND_RNs_17, IND_Rev_17, BLK_RNs_17, BLK_Rev_17,
            IND_RNs_18, IND_Rev_18, BLK_RNs_18, BLK_Rev_18,
            IND_RNs_19, IND_Rev_19, BLK_RNs_19, BLK_Rev_19,
            IND_RNs_20, IND_Rev_20, BLK_RNs_20, BLK_Rev_20]
Merge_PRO = functools.reduce(lambda left, right:
                             pd.merge(left, right, on='PROFILE_ID', how='outer'), PRO_IN_1)
account_profile_with_production = pd.merge(Profile, Merge_PRO,
                                           left_on='Accounts Account ID',
                                           right_on='PROFILE_ID', how='left')
account_profile_with_production.to_excel(r'output\account_profile_with_production.xlsx', index=None)
print('Done!!!!!')
