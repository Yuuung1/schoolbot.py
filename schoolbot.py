import telebot
from telebot import types
import webbrowser
import openpyxl
from openpyxl import load_workbook
import datetime

bot = telebot.TeleBot('6576190355:AAG9VaxezSQn3F1aERId0MklelSb-5ECPxs')
fn = 'timetable.xlsx'
wb = load_workbook(fn)
uch = wb['ученики']


markupbuttonmenu = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
menu = types.KeyboardButton('Меню')
markupbuttonmenu.add(menu)


markupbuttonmenuteach = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
timetablesettings = types.KeyboardButton("Изменить расписание уроков")
zvsettings = types.KeyboardButton("Изменить расписание звонков")
markupbuttonmenuteach.add(timetablesettings, zvsettings)


markupmenu = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
todaytimetable = types.KeyboardButton("Расписание на сегодня")
weektimetable = types.KeyboardButton("Расписание на неделю")
callstimetable = types.KeyboardButton("Расписание звонков")
markupmenu.add(todaytimetable, weektimetable, callstimetable)


today = datetime.date.today()


@bot.message_handler(commands=['start'])
def start(message):
    mess = f'Приветствую тебя, {message.from_user.first_name}! Я - твоё карманное расписание уроков. Я помогу тебе соореентироваться в уроках и звонках.'
    bot.send_message(message.chat.id, mess, parse_mode='html')
    markup1 = types.ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    student = types.KeyboardButton('Я ученик')
    father = types.KeyboardButton('Я родитель')
    markup1.add(student, father)
    bot.send_message(message.chat.id, 'Чтобы начать пользоваться ботом, тебе нужно зарегестрироваться.', reply_markup=markup1)


@bot.message_handler(commands=['site'])
def site(message):
    webbrowser.open('http://226school.ru')


@bot.message_handler(commands=['help', 'support'])
def help(message):
    bot.send_message(message.chat.id, '<b>Если у вас возникли проблемы с ботом обратитесь к</b> <em><u>@obhoyuuung</u></em>', parse_mode='html')


@bot.message_handler(commands=['menu'])
def menu(message):
    bot.send_message(message.chat.id,'Приветствую!\nНаш бот поможет тебе разобраться в твоем раписании\nГотов узнать что тебя ждет сегодня?', reply_markup=markupmenu)


@bot.message_handler()
def main(message):
    if message.text.lower() == 'привет':
        bot.send_message(message.chat.id, f'Привет, {message.from_user.first_name}')
    if message.text.lower() == 'id':
        bot.reply_to(message, f'ID, {message.from_user.id}')
    if message.text.lower() == 'я ученик':
        bot.reply_to(message, f'Привет, {message.from_user.first_name}, рад познакомится. Чтобы зарегестрироваться напиши своё ФИО и класс через запятую. Пример - Иван Иванов, 11А')
        bot.register_next_step_handler(message, rega)
    if message.text.lower() == 'я родитель':
        bot.reply_to(message, f'Здравствуйте, {message.from_user.first_name}, рад познакомится. Чтобы зарегестрироваться напиши своё ФИО и класс ребенка через запятую. Пример - Герасимов Александр, 11А')
        bot.register_next_step_handler(message, rega)
    if message.text.lower() == 'меню':
        bot.send_message(message.chat.id,'Приветствую!\nНаш бот поможет тебе разобраться в твоем раписании\nГотов узнать что тебя ждет сегодня?', reply_markup=markupmenu)
    if message.text.lower() == 'расписание на сегодня':
        workbook = openpyxl.open('timetable.xlsx', read_only=True)
        uchbook = workbook['ученики']
        for i in range(1,1000):
            if uchbook['A'+str(i)].value == str(message.from_user.id):
                cl = uchbook['B'+str(i)].value
        if today.strftime('%A') == 'Monday':
            sheetcl = workbook[str(cl)]
            day = 'понедельник'
            if sheetcl['A1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['A1'].value
            if sheetcl['A2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['A2'].value
            if sheetcl['A3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['A3'].value
            if sheetcl['A4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['A4'].value
            if sheetcl['A5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['A5'].value
            if sheetcl['A6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['A6'].value
            if sheetcl['A7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['A7'].value
        if today.strftime('%A') == 'Tuesday':
            day = 'вторник'
            sheetcl = workbook[str(cl)]
            if sheetcl['B1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['B1'].value
            if sheetcl['B2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['B2'].value
            if sheetcl['B3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['B3'].value
            if sheetcl['B4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['B4'].value
            if sheetcl['B5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['B5'].value
            if sheetcl['B6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['B6'].value
            if sheetcl['B7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['B7'].value
        if today.strftime('%A') == 'Wednesday':
            day = 'среда'
            sheetcl = workbook[str(cl)]
            if sheetcl['C1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['C1'].value
            if sheetcl['C2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['C2'].value
            if sheetcl['C3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['C3'].value
            if sheetcl['C4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['C4'].value
            if sheetcl['C5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['C5'].value
            if sheetcl['C6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['C6'].value
            if sheetcl['C7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['C7'].value
        if today.strftime('%A') == 'Thursday':
            day = 'четверг'
            sheetcl = workbook[str(cl)]
            if sheetcl['D1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['D1'].value
            if sheetcl['D2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['D2'].value
            if sheetcl['D3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['D3'].value
            if sheetcl['D4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['D4'].value
            if sheetcl['D5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['D5'].value
            if sheetcl['D6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['D6'].value
            if sheetcl['D7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['D7'].value
        if today.strftime('%A') == 'Friday':
            day = 'пятница'
            sheetcl = workbook[str(cl)]
            if sheetcl['E1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['E1'].value
            if sheetcl['E2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['E2'].value
            if sheetcl['E3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['E3'].value
            if sheetcl['E4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['E4'].value
            if sheetcl['E5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['E5'].value
            if sheetcl['E6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['E6'].value
            if sheetcl['E7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['E7'].value
        if today.strftime('%A') == 'Saturday':
            day = 'суббота'
            sheetcl = workbook[str(cl)]
            if sheetcl['F1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['F1'].value
            if sheetcl['F2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['F2'].value
            if sheetcl['F3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['F3'].value
            if sheetcl['F4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['F4'].value
            if sheetcl['F5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['F5'].value
            if sheetcl['F6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['F6'].value
            if sheetcl['F7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['F7'].value
        if today.strftime('%A') == 'Sunday':
            day = 'воскресенье'
            sheetcl = workbook[str(cl)]
            if sheetcl['G1'].value == None:
                less1 = 'Нету'
            else:
                less1 = sheetcl['G1'].value
            if sheetcl['G2'].value == None:
                less2 = 'Нету'
            else:
                less2 = sheetcl['G2'].value
            if sheetcl['G3'].value == None:
                less3 = 'Нету'
            else:
                less3 = sheetcl['G3'].value
            if sheetcl['G4'].value == None:
                less4 = 'Нету'
            else:
                less4 = sheetcl['G4'].value
            if sheetcl['G5'].value == None:
                less5 = 'Нету'
            else:
                less5 = sheetcl['G5'].value
            if sheetcl['G6'].value == None:
                less6 = 'Нету'
            else:
                less6 = sheetcl['G6'].value
            if sheetcl['G7'].value == None:
                less7 = 'Нету'
            else:
                less7 = sheetcl['G7'].value
        bot.send_message(message.chat.id,f'Сегодня <b>{day}</b>!\nВаше расписание на сегодня:\n1. {less1}\n2. {less2}\n3. {less3}\n4. {less4}\n5. {less5}\n6. {less6}\n7. {less7}', parse_mode='html', reply_markup=markupbuttonmenu)
    if message.text.lower() == 'расписание на неделю':
        workbook = openpyxl.open('timetable.xlsx', read_only=True)
        uchbook = workbook['ученики']
        for i in range(1,1000):
            if uchbook['A'+str(i)].value == str(message.from_user.id):
                cl = uchbook['B'+str(i)].value
        if cl == '5А':
            bot.send_photo(message.chat.id, open('5А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '6А':
            bot.send_photo(message.chat.id, open('6А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '7А':
            bot.send_photo(message.chat.id, open('7А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '8А':
            bot.send_photo(message.chat.id, open('8А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '9А':
            bot.send_photo(message.chat.id, open('9А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '10А':
            bot.send_photo(message.chat.id, open('10А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '11А':
            bot.send_photo(message.chat.id, open('11А.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '5Б':
            bot.send_photo(message.chat.id, open('5Б.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '6Б':
            bot.send_photo(message.chat.id, open('5Б.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '7Б':
            bot.send_photo(message.chat.id, open('7Б.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '8Б':
            bot.send_photo(message.chat.id, open('8Б.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '9Б':
            bot.send_photo(message.chat.id, open('9Б.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '5В':
            bot.send_photo(message.chat.id, open('5В.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '6В':
            bot.send_photo(message.chat.id, open('6В.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
        if cl == '7В':
            bot.send_photo(message.chat.id, open('7В.jpg', 'rb'), f'Расписание {cl} класса на неделю.', reply_markup=markupbuttonmenu)
    if message.text.lower() == 'расписание звонков':
        workbook = openpyxl.open('timetable.xlsx', read_only=True)
        zvbook = workbook['звонки']
        zvsbbook = workbook['звонкисб']
        zv1, zv2, zv3, zv4, zv5, zv6, zv7 = zvbook['A1'].value, zvbook['A2'].value, zvbook['A3'].value, zvbook['A4'].value, zvbook['A5'].value, zvbook['A6'].value, zvbook['A7'].value
        zvs1, zvs2, zvs3, zvs4, zvs5, zvs6, zvs7 = zvsbbook['A1'].value, zvsbbook['A2'].value, zvsbbook['A3'].value, zvsbbook['A4'].value, zvsbbook['A5'].value, zvsbbook['A6'].value, zvsbbook['A7'].value
        bot.send_message(message.chat.id,f'Расписание звонков:\n\n1. {zv1}\n2. {zv2}\n3. {zv3}\n4. {zv4}\n5. {zv5}\n6. {zv6}\n7. {zv7}\n\nРасписание звонков на субботу:\n\n1. {zvs1}\n2. {zvs2}\n3. {zvs3}\n4. {zvs4}\n5. {zvs5}\n6. {zvs6}\n7. {zvs7}', reply_markup=markupbuttonmenu)


def rega(message):
    if len(message.text.split(', '))==2:
        reg = message.text.split(', ')
        bot.send_message(message.chat.id, f'Аккаунт {reg[0]} зарегестрирован в {reg[1]} классе', reply_markup=markupbuttonmenu)
        uch.append([str(message.from_user.id), reg[1]])
        wb.save(fn)
        wb.close()
    else:
        bot.send_message(message.chat.id, f'Регистрация не прошла!\nВводите данные как в примере(через запятую): Иван Иванов, 5А')
        bot.register_next_step_handler(message, rega)


bot.polling(none_stop=True)