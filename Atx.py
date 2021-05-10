import os
import subprocess
import win32com.client as wc
import webbrowser
import wikipedia
import datetime
import speech_recognition as sr
import pyautogui as pag
import time

    
def Greeting():
    hour = int(datetime.datetime.now().hour)
    
    v=[1,0]
    x=0
    if hour>= 4 and hour<12 :
        speak("Good morning")
    elif hour>=12 and hour<18 :
        speak("Good Afternoon")
    elif hour>=18 and hour<=20 :
        speak("Good Evening")
    else:
        speak("Its time to sleep")
        speak("if you want to work, otherwise go and sleep well")
    speak("How can I  help you, Sir")
"""
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
#print(voices[2].id)
engine.setProperty('voice',voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
"""


def speak(text):
    sp = wc.Dispatch("SAPI.SpVoice")
    sp.Speak(text)







def takeCommand():

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening......")
        r.pause_threshold=1
        audio = r.listen(source)

    try:
        print("Recognizing......")

        query = r.recognize_google(audio, language="en-in")
        print( query)

    except Exception as e:
        print("............")

        return "...."

    return query

def res():
   
    chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
    cu=["cuims", "college id", "college login", "uims", "ums"]
    while True:
        query = takeCommand().lower()

        
            

        if 'wikipedia' in query:

            speak('Searching wikipedia....')

            query = query.replace("wikipedia", "")

            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            
            print(results)
            speak(results)
            

        elif 'who is' in query:

            speak('looking for the results')

            query = query.replace("wikipedia", "")

            results = wikipedia.summary(query, sentences=1)
            speak("According to wikipedia")
            
            print(results)
            speak(results)
            
        
        elif 'what is ' in query:

            try:
                speak('looking for the results')

                query = query.replace("wikipedia", "")
        
                results = wikipedia.summary(query.upper(), sentences=5)
                speak("According to wikipedia")
            
                print(results)
                speak(results)

            except:
                speak("Try different id or word")
            


        elif 'open youtube' in query :
            speak("Okay!")
            webbrowser.get(chrome_path).open('http://youtube.com/')
            
        elif 'geeksforgeeks' in query :
            speak("Sure! opening geeks for geeks")
            webbrowser.get(chrome_path).open('https://practice.geeksforgeeks.org/')

        
        elif query in cu:
            
            speak("Sure! opening C U I M S")
            webbrowser.get(chrome_path).open("https://uims.cuchd.in/uims/")

        elif 'who is your founder' in query or 'who created you' in query or 'who made you in ' in query:
            speak("Ankit Bishwas")
                
        elif 'open chrome'  in query:
            speak("Sure! opening google chrome")
           # webbrowser.open("google.com")
        
            #webbrowser.get(chrome_path).open('http://google.com/')
            subprocess.Popen("C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")

            
        elif 'terminate' in query or'stop' in query or "exit" in query or "shut up" in query:
            speak("Okay bye!")
            break
        elif 'search' in  query :
            query = query.replace("search", "")
            
            #for j in search(query, tld="co.in", num=10, stop=10, pause=2):
             #   print(j)
            webbrowser.get(chrome_path).open('https://google.com/?#q=' + query)
        elif "who are you" in query :
            speak("I am your personal assistant, ATX")

        elif "what is your name" in query:
            speak("My name is ATX")

        elif "aap kaun hai" in query or "tum kon ho"  in query:
            speak("main ATX hu")
        elif "who am i" in query or "what is my name" in query:
            speak("apka naam ankit hai")
        elif "thank you" in query:
            speak(" you are welcome")

        elif "lock" in query:
            speak("enter your passcode ")

            qr = takeCommand().lower()
            if qr=="31":
                os.system("rundll32.exe user32.dll, LockWorkStation")
                break
            else:
                res()

        elif "shutdown" in query or "power off" in query:
            speak("enter passcode")
            xr = takeCommand().lower()
            if xr=="99":
                speak("Shutting down your pc")
                os.system("shutdown /s /t 3")
                break
  
            else:
                res()
             
        

        elif "...." not in query :
           
            
            #for j in search(query, tld="co.in", num=10, stop=10, pause=2):
             #   print(j)
            speak("Ok, searching for"+ query)
            webbrowser.get(chrome_path).open('https://google.com/?#q=' + query)

        elif "open cortana" in query or "open Cortana" in query:
            pag.hotkey("win", "x")
        else:
            speak("Say something")
            


    
if __name__=="__main__":
    count=0
    xm=0
    print("A_________T________X")
    print()
    speak("Hi! I am ATX, version 1.002")
    speak("who are you?")
    xx = takeCommand().lower()
    coll=["i am", "my name is", "myself", "it's ", "its", "it is", 'its me', 'it is me']
    xm=0
    for word in coll:
        if word in xx:
            xx=xx.replace(word, "")
    #xx= xx.replace("I am", "")
    for count in range(3):
        if(xm==1):
            break
        speak("hello " + xx )
        speak("enter passcode")
        xt = takeCommand().lower()
    #xx = str(xt.split(" ", ""))
        
        if "98" in xt:
            xm=1
            Greeting()
            res()
            
        
        else:
            if count==2:
                speak("you are not autherized")
            else:   speak("try again")

            

        
        
        
            
    
