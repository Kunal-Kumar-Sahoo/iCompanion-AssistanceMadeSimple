from __future__ import print_function
import ctypes
import datetime
import json
import os
import os.path
import pickle
import smtplib
import time
import webbrowser
import platform
from termcolor import colored
from urllib.request import urlopen
import pyjokes
import pyttsx3
import pytz
import requests
import speech_recognition as sr
import wikipedia
import winshell
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from mysql.connector import MySQLConnection
import csv
import matplotlib.pyplot as plt
from ecapture import ecapture as ec


def virtualAssistant():
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
    MONTHS = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
              "november", "december"]
    DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    DAY_EXTENTIONS = ["rd", "th", "st", "nd"]

    def credAccess():
        with open("userCred.txt", 'rt') as fin:
            creds = fin.readlines()
            username = creds[1]
            USERNAME = username[0:len(username) - 1]
            email_id = creds[3]
            USER_EMAIL_ID = email_id[0:len(email_id) - 1]
            password = creds[5]
            USER_EMAIL_PASS = password[0:len(password) - 1]
            sql = creds[7]
            USER_SQL_PASS = sql[0:len(sql) - 1]
            return USERNAME, USER_EMAIL_ID, USER_EMAIL_PASS, USER_SQL_PASS

    def speak(text):
        engine = pyttsx3.init("sapi5")  # object creation

        """ RATE"""
        rate = engine.getProperty('rate')  # getting details of current speaking rate
        engine.setProperty('rate', 125)  # setting up new voice rate

        """VOLUME"""
        volume = engine.getProperty('volume')  # getting to know current volume level (min=0 and max=1)
        engine.setProperty('volume', 1.0)  # setting up volume level  between 0 and 1

        """VOICE"""
        voices = engine.getProperty('voices')  # getting details of current voice
        # engine.setProperty('voice', voices[0].id)  #changing index, changes voices. o for male
        engine.setProperty('voice', voices[1].id)  # changing index, changes voices. 1 for female

        engine.say(text)
        engine.runAndWait()

    recognitionMode = int(input('''By which mode would you like to give commands : 
        1 - Voice Mode 
        2 - Script Mode\n'''))

    def take_command():
        global query, endTime
        if recognitionMode == 1:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                print("Listening...")
                r.pause_threshold = 1
                audio = r.listen(source)
            try:
                print("Recognizing...")
                query = r.recognize_google(audio, language='en-in')
                print(f"User said: {query}\n")
            except Exception as e:
                print("Say that again please...")
                return "None"

        elif recognitionMode == 2:
            try:
                query = input('Enter Command : ')
            except Exception as e:
                print(e)
        return query.lower()

    def authenticate_google():
        """Shows basic usage of the Google Calendar API.
        Prints the start and name of the next 10 events on the user's calendar.
        """
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('calendar', 'v3', credentials=creds)

        return service

    def activity(today, initTime, endTime):

        deltaTime = endTime - initTime
        duration = deltaTime.total_seconds() / 60

        importedData = []

        with open('screenActivity.csv', 'r') as file:
            reader = csv.reader(file)
            for row in reader:
                importedData.append(row)
        dict = {}
        for i in range(len(importedData)):
            if len(importedData) != 0:
                dict[importedData[i][0]] = round(float(importedData[i][1]), 2)
        for i in dict:
            if i == today:
                print(dict[i])
                dict[i] = duration + dict[i]
                break
        else:
            dict[today] = duration
        print(dict)
        with open('screenActivity.csv', 'w') as file:
            for key in dict.keys():
                file.write("%s,%s\n" % (key, dict[key]))
        x = []
        y = []
        with open('screenActivity.csv', 'r') as csvfile:
            plots = csv.reader(csvfile, delimiter=',')
            for row in plots:
                x.append((row[0]))
                y.append(float(row[1]))
        plt.plot(x, y)
        plt.xticks(rotation=30)
        plt.xlabel('Date')
        plt.ylabel('Duration(sec)')
        plt.title('Screen Activity')
        plt.show()

    def get_events(day, service):
        # Call the Calendar API
        date = datetime.datetime.combine(day, datetime.datetime.min.time())
        end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
        utc = pytz.UTC
        date = date.astimezone(utc)
        end_date = end_date.astimezone(utc)
        events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(),
                                              timeMax=end_date.isoformat(), singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])
        if not events:
            speak('No upcoming events found.')
        else:
            speak(f"You have {len(events)} events on this day.")
            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                print(start, event['summary'])
                start_time = str(start.split("T")[1].split("-")[0])
                if int(start_time.split(":")[0]) < 12:
                    start_time = start_time + "am"
                else:
                    start_time = str(int(start_time.split(":")[0]) - 12) + start_time.split(":")[1]
                    start_time = start_time + "pm"

                speak(event["summary"] + " at " + start_time)

    def get_date(text):
        text = text.lower()
        today = datetime.date.today()

        if text.count("today") > 0:
            return today

        day = -1
        day_of_week = -1
        month = -1
        year = today.year

        for word in text.split():
            if word in MONTHS:
                month = MONTHS.index(word) + 1
            elif word in DAYS:
                day_of_week = DAYS.index(word)
            elif word.isdigit():
                day = int(word)
            else:
                for ext in DAY_EXTENTIONS:
                    found = word.find(ext)
                    if found > 0:
                        try:
                            day = int(word[:found])
                        except:
                            pass

        if month < today.month and month != -1:  # if the month mentioned is before the current month set the year to the next
            year = year + 1

        if month == -1 and day != -1:  # if we didn't find a month, but we have a day
            if day < today.day:
                month = today.month + 1
            else:
                month = today.month

        # if we only found a dta of the week
        if month == -1 and day == -1 and day_of_week != -1:
            current_day_of_week = today.weekday()
            dif = day_of_week - current_day_of_week

            if dif < 0:
                dif += 7
                if text.count("next") >= 1:
                    dif += 7

            return today + datetime.timedelta(dif)

        if day != -1:
            return datetime.date(month=month, day=day, year=year)

    def note(text):
        date = datetime.datetime.now()
        file_name = str(date).replace(":", "-") + "-note.txt"
        with open(file_name, "w") as f:
            f.write(text)
        speak("I've made a note of that.")
        osCommandString = f"notepad.exe {file_name}"
        os.system(osCommandString)

    def sendEmail(to, content):
        global USER_EMAIL_ID, USER_EMAIL_PASS
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()

        server.login(USER_EMAIL_ID, USER_EMAIL_PASS)
        server.sendmail(USER_EMAIL_ID, to, content)
        server.close()

    def history(text):
        date_time = datetime.datetime.now().strftime("%d-%m-%Y, %H:%M:%S")
        f = open("Command history.txt", "a")
        f.write(f"{date_time} : {text}\n")

    def contactBook(USER_SQL_PASS):
        global emailID
        mydb = MySQLConnection(host="localhost", user="root", database="contact_book")
        mycursor = mydb.cursor()
        name1 = str(input("Enter the name : "))
        mycursor.execute("SELECT * FROM contact_table")
        myresult = mycursor.fetchall()
        for x in myresult:
            if name1 == x[0]:
                emailID = x[1]
                break
        else:
            print("Record not found")
            insertChoice = input("Would you like to add contact ? ")
            if insertChoice.lower().startswith("y"):
                mydb1 = MySQLConnection(host="localhost", user="root",  database="contact_book")
                mycursor1 = mydb1.cursor()
                mail = str(input("Enter the e-mail ID of the new contact : "))
                sql = "INSERT INTO contact_table(name, email_id) VALUES(%s,%s)"
                val = (name1, mail)
                mycursor1.execute(sql, val)
                mydb1.commit()
                print(mycursor1.rowcount, "record inserted.")

            mydb = MySQLConnection(host="localhost", user="root", database="contact_book")
            mycursor = mydb.cursor()
            mycursor.execute("SELECT * FROM contact_table")
            myresult = mycursor.fetchall()
            for x in myresult:
                if name1 == x[0]:
                    emailID = x[1]
                    break
        return emailID

    def wishMe(name):
        hour = int(datetime.datetime.now().hour)
        if hour >= 0 and hour < 12:
            speak("Good Morning!" + str(name))
        elif hour >= 12 and hour < 16:
            speak("Good Afternoon!" + str(name))
        else:
            speak("Good Evening!" + str(name))

    if __name__ == "__main__":
        USERNAME, USER_EMAIL_ID, USER_EMAIL_PASS, USER_SQL_PASS = credAccess()
        wishMe(USERNAME)
        initTime = datetime.datetime.now()
        today = str(datetime.date.today())
        WAKE = assname = "computer"
        SERVICE = authenticate_google()
        '''bat = psutil.sensors_battery()
        if 60 <= bat[0] <= 100:
            print(colored(f"Outstanding performance: {bat[0]} % battery remaining", "green"))
        elif 30 <= bat[0] < 60:
            print(colored(f"Good performance: {bat[0]} % battery remaining", "yellow"))
        else:
            print(colored(f"Plug-in required: {bat[0]} % battery remaining"), "red")'''
        spec = platform.uname()
        print("System = ", spec[0])
        print("Host Name = ", spec[1])
        print("Release(Windows) = ", spec[2])
        print("PC's Version = ", spec[3])
        print("Machine = ", spec[4])
        print("PC's Processor= ", spec[5])
        print("Starting...")
        speak("Initialized")
        while True:
            text = take_command()

            if text.count(WAKE) > 0:

                text = text.replace(WAKE, "")

                CALENDAR_STRS = ["what do i have", "do i any have plans"]
                for phrase in CALENDAR_STRS:
                    if phrase in text:
                        date = get_date(text)
                        if date:
                            get_events(date, SERVICE)
                        else:
                            speak("I don't understand")

                NOTE_STRS = ["make a note", "write this down", "remember this"]
                for phrase in NOTE_STRS:
                    if phrase in text:
                        speak("What would you like me to write down?")
                        note_text = take_command()
                        note(note_text)

                if "wikipedia" in text:
                    try:
                        speak('Searching Wikipedia...')
                        text = text.replace("wikipedia", "")
                        results = wikipedia.summary(text, sentences=2)
                        speak("According to Wikipedia")
                        print(results)
                        speak(results)
                    except Exception as e:
                        print(e)

                elif 'open youtube' in text:
                    speak("Here you go to Youtube")
                    webbrowser.open("https://www.youtube.com")

                elif 'search' in text:
                    speak("Searching")
                    search = text.replace("search", "")
                    link = f"https://www.google.com.tr/search?q={search}"
                    webbrowser.open(link)

                elif 'open stackoverflow' in text:
                    speak("Here you go to Stack Over flow. Happy coding!")
                    webbrowser.open("https://www.stackoverflow.com")

                elif 'play' in text:
                    music = text.replace("play", "")
                    link = f"https://music.youtube.com/search?q={music}"
                    webbrowser.open(link)

                elif 'time' in text:
                    strTime = datetime.datetime.now().strftime("%H:%M:%S")
                    print(strTime)
                    speak(f"the time is {strTime}")

                elif 'narrate' in text:
                    speak('What shall I narrate?')
                    string = input("What shall I narrate : ")
                    speak(string)

                elif 'email' in text or 'send email' in text:
                    try:
                        speak("What should I say?")
                        content = take_command()
                        speak("whom should i send")
                        to = contactBook(USER_SQL_PASS)
                        sendEmail(to, content)
                        speak("Email has been sent !")
                    except Exception as e:
                        print(e)
                        speak("Sorry, I am not able to send this email")

                elif 'how are you' in text:  # 3
                    speak("I am fine. Thanks for asking")
                    speak(f"How are you?")
                    text = take_command()
                    if 'fine' in text or "good" in text:
                        speak("It's good to know that your fine")

                elif "change name" in text:
                    speak("What would you like to call me")
                    WAKE = take_command()
                    speak("Thanks for giving me special name")

                elif 'exit' in text or 'good bye' in text:
                    endTime = datetime.datetime.now()
                    print(f"Time duration of Usage : {endTime - initTime}")
                    speak("Thanks for giving me your time")
                    activity(today, initTime, endTime)
                    exit()

                elif 'joke' in text:
                    joke = pyjokes.get_joke()
                    print(joke)
                    speak(joke)

                elif "why you came to world" in text:  # 4
                    speak("To help you out, further thanks to team JKS")

                elif "who are you" in text:  # 1
                    speak("I am iCompanion your personal assistant")

                elif 'tell me something about you' in text:  # 2
                    speak("")

                elif 'change background' in text:
                    ctypes.windll.user32.SystemParametersInfoW(20,
                                                               0,
                                                               "C:\\Windows\\Web\\Wallpaper\\Theme1",
                                                               0)
                    speak("Background changed successfully")

                elif 'news' in text:
                    try:
                        jsonObj = urlopen(
                            '''http://newsapi.org/v2/top-headlines?country=in&apiKey=b34c76c69a4048dfa815774ae73ce139''')
                        data = json.load(jsonObj)
                        speak('So, here I have few latest news for you across the world')
                        i = 1
                        for item in data['articles']:
                            if i <= 5:
                                print(str(i) + '. ' + item['title'] + '\n')
                                speak(item['title'] + '\n')
                                i += 1
                    except Exception as e:
                        print(str(e))
                        speak('Here are some headlines from the Times of India,Happy reading')
                        webbrowser.open("https://timesofindia.indiatimes.com/home/headlines")
                        time.sleep(6)

                elif 'lock window' in text:
                    speak("locking the device")
                    ctypes.windll.user32.LockWorkStation()

                elif 'shutdown' in text:
                    endTime = datetime.datetime.now()
                    print(f"Time duration of Usage : {endTime - initTime}")
                    speak("Hold On a Sec ! Your system is on its way to shut down")
                    os.system("shutdown -s")
                    activity(today, initTime, endTime)

                elif 'empty recycle bin' in text:
                    winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                    speak("Recycle Bin Recycled")

                elif "don't listen" in text or "stop listening" in text:
                    speak("for how much time you want me to stop listening commands")
                    a = int(take_command())
                    time.sleep(a)

                elif 'make a stopwatch' in text or 'stopwatch' in text:
                    def countdown(t):
                        while t > 0:
                            print(t)
                            t -= 1
                            time.sleep(1)

                    speak("For how much time should I set the timer?")
                    seconds = int(input("For how much time should I set the timer: "))
                    countdown(seconds)
                    print(colored("Time's Up"), "red")
                    speak("Time's Up")

                elif "where is" in text:
                    text = text.replace("where is", "")
                    location = text
                    speak("User asked to Locate")
                    speak(location)
                    webbrowser.open(f"https://www.google.nl/maps/place/{location}")

                elif "restart" in text:
                    endTime = datetime.datetime.now()
                    print(f"Time duration of Usage : {endTime - initTime}")
                    os.system("shutdown -r")
                    activity(today, initTime, endTime)

                elif "hibernate" in text or "sleep" in text:
                    speak("Hibernating")
                    os.system("shutdown -h")

                elif "log off" in text or "sign out" in text:
                    speak("Make sure all the application are closed before sign-out")
                    time.sleep(5)
                    os.system("shutdown -l")
                    activity(today, initTime, endTime)

                elif "weather" in text:
                    api_key = "6c7e7e30ff6df9bc6b22fb28c227ff24"
                    base_url = "https://api.openweathermap.org/data/2.5/weather?"
                    speak("what is the city name")
                    city_name = take_command()
                    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
                    response = requests.get(complete_url)
                    x = response.json()
                    if x["cod"] != "200":
                        data = response.json()
                        main = data['main']
                        temp = main['temp']
                        temperature = round((temp - 273.15), 2)
                        humidity = main['humidity']
                        pressure = main['pressure']
                        report = data['weather']
                        print(f"{city_name:-^30}")
                        print(f"{temperature}°C in {city_name}");
                        speak(f"{temperature}°C in {city_name}")
                        print(f"{humidity}% humidity");
                        speak(f"{humidity}% humidity")
                        print(f"Pressure is {pressure} hPa");
                        speak(f"Pressure is {pressure} hPa")
                        print(f"Report says that there is {report[0]['description']}");
                        speak(f"Report says that there is {report[0]['description']}")
                    else:
                        speak("City Not Found")

                elif "reboot" in text:
                    endTime = datetime.datetime.now()
                    print(f"Time duration of Usage : {endTime - initTime}")
                    i = 3
                    while i >= 1:
                        speak("Rebooting in")
                        speak(i)
                        i -= 1
                    speak("Rebooting now")
                    activity(today, initTime, endTime)
                    virtualAssistant()

                elif "history" in text:
                    osCommandString = f"notepad.exe Command history.txt"
                    os.system(osCommandString)
                    speak("Would you like me to clear your command history ?")
                    confirmation = take_command()
                    if confirmation.lower().startswith("y"):
                        open("Command history.txt", 'w').close()

                elif 'exit' in text or 'good bye' in text:
                    endTime = datetime.datetime.now()
                    print(f"Time duration of Usage : {endTime - initTime}")
                    speak("Thanks for giving me your time")
                    activity(today, initTime, endTime)
                    exit()
                elif 'clear screen' in text:
                    os.system('cls')
                elif "camera" in text or "take a photo" in text:
                    ec.capture(0, f"{assname} Camera ", "img.jpg")

                history(text)


if __name__ == "__main__":
    try:
        virtualAssistant()
    except Exception as e:
        print(e)


