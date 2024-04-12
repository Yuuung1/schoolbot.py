import datetime

print(datetime.datetime.now())
print(datetime.date.today())

date_time = datetime.datetime.now()
current_time = date_time.time()
print(current_time)

print(datetime.time(10,30,11))
print(datetime.datetime(2023, 12, 11, 10, 14, 50))

date = "06/05/22 12:06:58"
data_obj = datetime.datetime.strptime(date, "%m/%d/%y %H:%M:%S")
print(data_obj)
today = datetime.date.today()
print(d.strftime('%A'))