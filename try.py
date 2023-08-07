from datetime import datetime, timedelta
import datetime
import calendar

date_today = datetime.datetime.now()
currentMonth = datetime.datetime.now().month
currentYear = datetime.datetime.now().year
obj = calendar.Calendar()

print( obj.monthdatescalendar(currentYear, currentMonth)[0][0]+timedelta(days=6))
    