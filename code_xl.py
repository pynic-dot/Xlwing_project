import xlwings as xw
import pandas as pd
import numpy as np
import os
import calendar
import time
import datetime
from datetime import time
from pandas.tseries.offsets import CustomBusinessHour
from business_duration import businessDuration


dir_list = os.listdir('C://Users/akumar/Desktop/New folder/Raw_data_Feed_here')
path_for_excel = 'C://Users/akumar/Desktop/New folder/Raw_data_Feed_here/'+dir_list[0]
dateParser = lambda x:pd.to_datetime(x,format='%Y-%m-%d %H:%M:%S')
data_set =pd.read_excel(path_for_excel, parse_dates=["Created","Closed"], date_parser=dateParser)

unit_hour='hour'
tt_Wh= pd.Series([],dtype=float)
for x in range(len(data_set['Created'])):
    start_d = data_set["Created"][x]
    end_d = data_set["Closed"][x]
    tt_Wh[x]=(businessDuration(startdate=start_d,enddate=end_d,unit=unit_hour))

data_set.insert(len(data_set.columns),"Total_Business_Hrs",tt_Wh)

data_set["Total_Hrs"] = (data_set['Closed']-data_set['Created']).astype('timedelta64[h]')

unit_hour='day'
Total_Workig_days= pd.Series([],dtype=float)
for i in range(len(data_set['Created'])):
    start_d = data_set["Created"][i]
    end_d = data_set["Closed"][i]
    Total_Workig_days[i]=(businessDuration(startdate=start_d,enddate=end_d,unit=unit_hour))

data_set.insert(len(data_set.columns),"Total_BusinessWorkig_days",Total_Workig_days)

data_set.fillna(0,inplace=True)


TTH_15HRS_dedc = pd.Series([],dtype=float)
for rn in range(len(data_set.index)):
    if int(data_set["Total_BusinessWorkig_days"][rn]) > 0:
        TTH_15HRS_dedc[rn]= float(data_set["Total_Business_Hrs"][rn])-int(data_set["Total_BusinessWorkig_days"][rn])*15
    elif int(data_set["Total_BusinessWorkig_days"][rn]) == 0:
        TTH_15HRS_dedc[rn]= float(data_set["Total_Business_Hrs"][rn])-15

data_set.insert(len(data_set.columns),'Total_Bsn_hrs_15Hr_Diduction',TTH_15HRS_dedc)

a=data_set[['Owner','Total_Bsn_hrs_15Hr_Diduction','Ticket Number']].groupby(['Owner']).mean()['Total_Bsn_hrs_15Hr_Diduction']

b=data_set[['Owner','Ticket Number']].groupby(['Owner']).count()['Ticket Number']

c = data_set[['Owner',"Total_Hrs"]].groupby(['Owner']).sum()["Total_Hrs"]

new_col=pd.DataFrame(pd.concat([a,b,c],axis = 1))

new_col.rename(columns={'Total_Bsn_hrs_15Hr_Diduction':"Total_avg_time",'Ticket Number':"NumberOfTickets"},inplace=True)


def code_excel():
    wb = xw.Book.caller()
    wb.sheets[0].range('A1').value = data_set
    new_sheet = wb.sheets.add( 'New_Sheet' , after= 'Sheet1' )
    new_sheet.range('A1').value = new_col
if __name__ == '__main__':
    xw.Book('Book1.xlsm').set_mock_caller()
    code_excel()
