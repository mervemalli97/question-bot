import telebot
from telebot import types
import random
from openpyxl import load_workbook
from keep_alive import keep_alive

KEY = "5040911548:AAGM8aqtZnBCIK8pH2IlPoxPYXxncqMbrzY"
bot = telebot.TeleBot(KEY, parse_mode=None)

qs = load_workbook('book.xlsx', 'r')
q = qs.active

qs_length = q['C1'].value
subjects = []
subjects_count = []
for x in range(qs_length):
    y = 'A' + str(x)
    if q[y].value not in subjects:
        subjects.append(q[y].value)
        subjects_count.append(x)


@bot.message_handler(commands=["hi", "hello"])
def greet(message):
    try:
        send = "to get a random question press /random, to get the subjects in the database please press /subjects"
        markup = types.ReplyKeyboardMarkup(row_width=1)
        itembtn1 = types.KeyboardButton('/random')
        itembtn2 = types.KeyboardButton('/subjects')
        markup.row(itembtn1, itembtn2)
        bot.send_message(message.chat.id,
                         send,
                         reply_markup=markup)
    except Exception as e:
        print(e)

@bot.message_handler(regexp="random*")
def rand(message):
    try:
        r = 1
        while r in subjects_count:
            r = random.randrange(1, qs_length)
            if '_s' in message.text:
              if (int(message.text[9:]) == len(subjects_count)-1):
                r = random.randrange(subjects_count[int(message.text[9:])], qs_length+1)
              else:
                print(subjects_count[int(message.text[9:])]+1)
                r = random.randrange(subjects_count[int(message.text[9:])], subjects_count[int(message.text[9:])+1]+1)
        subject_of_qa = len(subjects_count) - 1
        for x in subjects_count:
            if r < x:
                subject_of_qa = subjects_count.index(x) - 1
                break
        question = "B" + str(r)
        subj = "B" + str(subjects_count[subject_of_qa])        
        send = q[subj].value + "\n" + q[question].value
        markup = types.ReplyKeyboardMarkup(row_width=1)
        itembtn1 = types.KeyboardButton('/a' + str(r))
        itembtn2 = types.KeyboardButton('/q' + str(r + 1))
        itembtn3 = types.KeyboardButton('/random')
        itembtn4 = types.KeyboardButton('/random_s' + str(subject_of_qa))
        markup.row(itembtn1, itembtn2)
        markup.row(itembtn3, itembtn4)
        bot.send_message(message.chat.id,
                         send,
                         reply_markup=markup)
    except Exception as e:
        print(e)


@bot.message_handler(commands=["subjects"])
def get_subjects(message):
    try:
        send = ""
        count = 1
        for s in subjects:
            if s == None:
              continue
            subject_name = "B" + str(subjects_count[count - 1])
            send = send + "/s" + str(count-1) + " " + str(
                q[subject_name].value) + "\n"
            count += 1
        bot.send_message(message.chat.id, send)
    except Exception as e:
        print(e)


@bot.message_handler(regexp="[aqs][0-9]*")
def get_this(message):
    try:
        subject_of_qa = len(subjects_count) - 1
        if (message.text.startswith("/a") | message.text.startswith("/q")):
            for x in subjects_count:
                if int(message.text[2:]) < x:
                    subject_of_qa = subjects_count.index(x) - 1
                    break
        markup = types.ReplyKeyboardMarkup(row_width=1)
        if message.text.startswith("/s"):
            count = subjects.index(message.text[1:])
            subject_name = "B" + str(subjects_count[count])
            subject_qcount = str(qs_length - subjects_count[count])
            if len(subjects_count) > int(message.text[2:]) + 1:
                subject_qcount = str(subjects_count[count + 1] -
                                     subjects_count[count] - 1)
            send = q[subject_name].value + " (" + subject_qcount + ")"
            itembtn1 = types.KeyboardButton("/q" +
                                            str(subjects_count[count] + 1))
            itembtn2 = types.KeyboardButton('/random')
            itembtn3 = types.KeyboardButton('/random_s' + message.text[2:])
            markup.row(itembtn1)
            markup.row(itembtn2, itembtn3)
            bot.send_message(message.chat.id, send, reply_markup=markup)
        elif message.text.startswith("/a"):
            answer = "C" + message.text[2:]
            subj = "B" + str(subjects_count[subject_of_qa])   
            send = q[subj].value + "\n" + q[answer].value
            itembtn1 = types.KeyboardButton("/q" + message.text[2:])
            itembtn2 = types.KeyboardButton("/q" +
                                            str(int(message.text[2:]) + 1))
            itembtn3 = types.KeyboardButton('/random')
            itembtn4 = types.KeyboardButton('/random_s' + str(subject_of_qa))
            markup.row(itembtn1, itembtn2)
            markup.row(itembtn3, itembtn4)
            bot.send_message(message.chat.id, send, reply_markup=markup)
        elif message.text.startswith("/q"):
            quest = "B" + message.text[2:]
            subj = "B" + str(subjects_count[subject_of_qa])     
            send = q[subj].value + "\n" + q[quest].value
            itembtn1 = types.KeyboardButton("/a" + message.text[2:])
            itembtn2 = types.KeyboardButton("/q" +
                                            str(int(message.text[2:]) + 1))
            itembtn3 = types.KeyboardButton('/random')
            itembtn4 = types.KeyboardButton('/random_s' + str(subject_of_qa))
            markup.row(itembtn1, itembtn2)
            markup.row(itembtn3, itembtn4)
            bot.send_message(message.chat.id, send, reply_markup=markup)
        else:
            pass
    except Exception as e:
        print(e)

keep_alive()
bot.polling()
