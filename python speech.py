import win32com.client
import speech_recognition as sr

def speak(text):
    engine = win32com.client.Dispatch("SAPI.SpVoice")
    engine.Speak(text)

def listen():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source, timeout=5)

    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio)
        print(f"You said: {query}")
        return query.lower()
    except sr.UnknownValueError:
        print("Sorry, I didn't catch that. Can you repeat?")
        return None
    except sr.RequestError as e:
        print(f"Could not request results from Google Speech Recognition service; {e}")
        return None

def assistant(query):
    if "hello" in query:
        speak("Hello! How can I assist you today?")
    elif "your name" in query:
        speak("I am your Python voice assistant. myself TAMIL!")
    elif "how are you" in query:
        speak("I'm just a computer program, but I'm here and ready to help!")
    elif "thank you" in query:
        speak("You're welcome!")
    elif "will shadows win the kani tamil event" in query:
        speak("yes definitely!")
    elif "had your lunch" in query:
        speak("yes! it was tasty")
    elif "what's the time" in query or "tell me the time" in query:
        # Include code to get the current time
        speak("Sorry, I am not programmed to tell the time yet.")
    elif "exit" in query:
        speak("Goodbye!")
        exit()
    else:
        speak("I'm sorry, I don't understand that command.")

if __name__ == "__main__":
    speak("Hello! I am your Python voice assistant.")
    
    while True:
        query = listen()
        if query:
            assistant(query)
