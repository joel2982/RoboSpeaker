import win32com.client as wincom
if __name__ == '__main__':
    while True:
        text = input("What do you want to say: ")
        speak = wincom.Dispatch("SAPI.SpVoice")
        if text == 'n':
            speak.Speak("I'll take your leave. Good Bye! ") 
            break  
        speak.Speak(text)   