from Imports import *

dateToday = ""
timeNow = ""
monthNowText = ""

yearNow = ""

def GetDateToday():
    global dateToday
    global monthNowText
    global monthNow
    global yearNow

    dateToday = datetime.datetime.today()
    monthNowText = dateToday.strftime("%B")
    monthNow = int(dateToday.month)
    yearNow = int(dateToday.year)
    dateToday = dateToday.strftime('%Y/%m/%d')

def GetTimeNow():
    global timeNow

    timeNow = datetime.datetime.today()
    timeNow = timeNow.strftime('%H:%M')