Calendar functions:

- DaysBetween(Date) As Long
    Returns the days between the calendar date and specified date

- DayOfYear() As Integer
    Returns the day of the current year

- GetMonthDays([Month]) As Integer
    Returns the total days of specified month

- GetMonthName(Month)
    Returns the name of specified month

- GetMoonPhaseDetail(Date) As Double
    Returns the deatail number of the moonphase from 0.00 to 0.99

- CalcMoonPhaseExact(Date) As MoonTypes
    Returns the moonphase of given date

- GetMoonPhaseInfo(Date, [Icon]) As String
    Returns the moonphase info for the specified date,
    where the icon parameter is optional

- GetQuarterInfo(Quarter, [Icon]) As String
    Returns the quarter info for the specified quarter,
    where the icon parameter is optional

- GetSeasonInfo(Season, [Icon])
    Returns the season info for the specified season,
    where the icon parameter is optional

- GetWeekDayName(WeekDayNumber) As String
    Returns the name of the week specified by daynumber

- GetZodiacInfo(ZodiacSign, [Icon]) As String
    Returns the zodiacsign info for the specified zodiacsign,
    where the icon parameter is optional

- IsDaySel(Day) As Boolean
    Returns True if that day is selected indeed


Calendar Subs:

- DayMarking(Day, MarkType, OnOff [TipText])
    Marks or Demarks a calendar day for the specified day
    MarkType sets the type marker 1 to 5 (defined by color)
    OnOff switcht the marke on or off
    TipText sets the tooltiptext of the specified marker

- DaySelect(Day, OnOff)
    Selects or Deselects a day for the specified day
    OnOff switcht the selection on or off

- Refresh()
    To refresh the whole calendar

- SetMarkColors([Color1], [Color2], [Color3], [Color4], [Color5])
    Sets the color of the markers
    Default color1 = Red
    Default color2 = Green
    Default color3 = Magenta
    Default color4 = Yellow
    Default color5 = Cyan
