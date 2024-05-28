# -*- coding: utf-8 -*-
"""
Created on Tue Feb  6 06:53:49 2024

@author: BHANUPRAKASH
"""
import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import wikipedia
import pyjokes
import pandas as pd
from docx import Document
import random
import pyautogui as pg
#import os
#import emojis


listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def talk(text):
    engine.say(text)
    engine.runAndWait()

def take_command():
    try:
        with sr.Microphone() as source:
            print('listening...')
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            command = command.lower()
            if 'veda ' in command:
                command = command.replace('veda', '')
                print(command)
    except:
        pass
    return command


def run_veda():
    command = take_command()
    print(command)
    
        
    
    if  'hello' in command:
        s =('sir','boss','commandor')
        for i in range (1):
            a=random.choice(s)            
            talk('hello'+a)
        
                
    elif 'play' in command:
        song = command.replace('play', '')
        talk('playing ' + song)
        pywhatkit.playonyt(song)
        
    elif 'find' in command:
        file = command.replace('find', '')
        talk('sure. finding' + file)
        with open ('C:/Users/BHANUPRAKASH/OneDrive/Desktop/coding/VIDEO-object.py' , 'r' ) as file:
            exec(file.read())
            
        
    #elif 'what' in command:
        #gk = command.replace('what is', '')
        #talk (gk)
       # pygkquiz
        
            
        
    elif 'message' in command:
        msg = command.replace('send', '')
        talk('sending ' + msg)
        pywhatkit.sendwhatmsg_to_group_instantly(msg)
        
    elif 'time' in command:
        time = datetime.datetime.now().strftime('%I:%M %p')
        talk('Current time is ' + time)
        
    elif 'heck ' in command:
        person = command.replace('who the heck is', '')
        info = wikipedia.summary(person, 1)
        print(info)
        talk(info)
    elif 'random '  in command:
        s=('good','bad','worst')
        for i in range (1):
            a=random.choice(s)
            talk("you are a "+a)
        
        
    elif 'you date' in command:
        talk('sorry, I have a headache')
    
    elif 'are you single' in command:
        talk('I am in a relationship with wifi')
        
    elif 'your name' in command:
        talk('I am veda voice assistant')
        
    elif 'joke' in command:
        talk(pyjokes.get_joke())
        
    else:
        talk('Please say the command again.')
        def update_excel(command):
            try:
                # Read existing Excel file or create a new DataFrame
                df = pd.read_excel('commands.xlsx')
            except FileNotFoundError:
                df = pd.DataFrame()

            # Add a new row with the command
            df = df.append({'Command': command}, ignore_index=True)

            # Save the DataFrame to Excel
            df.to_excel('commands.xlsx', index=False)

        # Example usage:
        update_excel(command)
        

while True:
    run_veda()