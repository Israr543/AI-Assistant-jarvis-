import pyttsx3
import speech_recognition as sr # For speech-to-text functionality                      type: ignore
import datetime # For time-related operations
import pyaudio # To handle audio input/output                                       type: ignore
import wikipedia # For searching Wikipedia                                 type: ignore
import webbrowser # To open web pages
import os # To interact with the operating system
import random # For random operations, like playing random music
import smtplib  # For sending emails
import subprocess # To execute system-level command 


# Initialize pyttsx3 for text-to-speech
engine = pyttsx3.init('sapi5')      # the technology for voice recognition and synthesis provided by Microsoft. It can be used to convert Text into Speech
voices = engine.getProperty('voices')
#print(voices[0].id)
engine.setProperty('voices', voices[0].id)


# Function to convert text to speech
def speak(audio):
    ''' convert text to speech'''
    engine.say(audio)
    engine.runAndWait()     

# Function to greet the user
def wishMe():
     ''' Greets the user based on the time of the day and engages in a brief interaction '''
     hour = int(datetime.datetime.now().hour)
     if hour >= 0 and hour < 12:
         speak("Good Morning!")
     elif hour >= 12 and hour < 18:
         speak("Good Afternoon!")
     else:
         speak("Good Evening!")

     speak("I am Jarvis, your assistant. Please tell me how I can help you.")
     speak("But before that, how are you today?")
    
     # Take user's response
     user_response = takeCommand().lower()
    
     if "fine" in user_response or "good" in user_response or "okay" in user_response:
         speak("That's great to hear! How can I assist you?")
     elif "not" in user_response or "bad" in user_response:
         speak("I'm sorry to hear that. I hope I can help make your day better.")
     else:
         speak("Got it. How can I assist you?")

# Function to take voice input from the user
def takeCommand():
    '''  
      it take microphone input from the user and convert into text
    '''         
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.adjust_for_ambient_noise(source)  # Adjust for ambient noise
        r.pause_threshold = 1    # Wait time before recognizing the input
        audio = r.listen(source)

    try:
        print("Recogniziting....")
        query = r.recognize_google(audio, language='en-US')  # Convert audio to text
        print(f"user said: {query}\n")

    except Exception as e:
        #print(e)
        print("Say that again please....")
        return "None"    

    return query



# Function to send an email
def sendEmail(to, content):
     """Sends an email using the SMTP protocol."""
     try:
         server = smtplib.SMTP('smtp.gmail.com', 587)  # Connect to Gmail SMTP server
         server.ehlo()  # Identify with the server
         server.starttls()  # Start TLS encryption
         server.login('53530@students.riphah.edu.pk','arbaz@123')  # Login credentials
         server.sendmail('53530@students.riphah.edu.pk', to, content)  # Send the email
         server.close()
         speak("Email has been sent!")
     except Exception as e:
          speak("Sorry, I couldn't send the email.")
          print(f"Error: {e}")


# Function to create a folder at a specified path
def create_folder(folder_path):
    """Creates a folder if it doesn't already exist."""
    try:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            speak(f"Folder has been created at {folder_path}")
        else:
            speak(f"A folder already exists at {folder_path}")
    except Exception as e:
        speak("An error occurred while creating the folder.")
        print(f"Error: {e}")


# Function to display system information
def system_info():
    """Displays basic system information."""
    import platform
    system = platform.system()
    release = platform.release()
    version = platform.version()
    architecture = platform.architecture()[0]
    speak(f"This system is running {system}, version {release}, {architecture}.")
    print(f"System: {system}\nVersion: {release}\nArchitecture: {architecture}")


# Function to take a screenshot
def take_screenshot():
    """Captures a screenshot and saves it to the desktop."""
    from PIL import ImageGrab
    speak("Taking a screenshot.")
    try:
        screenshot = ImageGrab.grab()
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop_path):
            os.makedirs(desktop_path)
        file_path = os.path.join(desktop_path, "screenshot.png")
        screenshot.save(file_path, "PNG")
        speak(f"Screenshot saved to {file_path}.")
    except Exception as e:
        speak("I encountered an issue while taking the screenshot.")
        print(f"Error: {e}")


# Function to open a predefined application
def open_application(app_name):
    """Opens common applications by name."""
    apps = {
        "notepad": "notepad.exe",
        "calculator": "calc.exe",
        "command prompt": "cmd.exe",
        "explorer": "explorer.exe",
        "settings": "ms-settings:"  # Settings app
    }
    app_name = app_name.strip().lower()  # Normalize input
    if app_name in apps:
        speak(f"Opening {app_name}.")
        try:
            os.system(apps[app_name])
        except Exception as e:
            speak(f"Sorry, I couldn't open {app_name}.")
            print(f"Error: {e}")
    else:
        speak(f"I couldn't find {app_name}. Please add it to my applications list.")


# Function to close a predefined application
def close_application(app_name):
    """Closes common applications by name."""
    try:
        speak(f"Closing {app_name}.")
        os.system(f"taskkill /f /im {app_name}.exe")
    except Exception as e:
        speak(f"Sorry, I couldn't close {app_name}.")
        print(f"Error: {e}")


# Function to lock the PC
def lock_pc():
    """Locks the computer."""
    speak("Locking the computer.")
    os.system("rundll32.exe user32.dll,LockWorkStation")

#they not work  # Function to shutdown the computer
def shutdown_pc():
     """Shutdown the PC."""
     speak("Are you sure you want to shut down the computer? Say yes to confirm.")
     confirmation = takeCommand().lower()  # Ask for confirmation

     if "yes" in confirmation:
         speak("Shutting down the computer.")
         try:
             # Shutdown command with a delay of 1 second
             subprocess.run(["shutdown", "/s", "/t", "1"], check=True)
         except subprocess.CalledProcessError as e:
             speak("I encountered an error while shutting down the computer.")
             print(f"CalledProcessError: {e}")
         except Exception as e:
             speak("An unexpected error occurred while attempting to shut down.")
             print(f"Error: {e}")
     else:
         speak("Shutdown operation canceled.")

#they not work  # Function to restart the computer
def restart_pc():
     """Restart the PC."""
     speak("Are you sure you want to restart the computer? Say yes to confirm.")
     confirmation = takeCommand().lower()
     if "yes" in confirmation:
         speak("Restarting the computer.")
         try:
             # Restart command with a delay of 1 second
             subprocess.run(["shutdown", "/r", "/t", "1"], check=True)
         except subprocess.CalledProcessError as e:
             speak("I encountered an error while restarting the computer.")
             print(f"CalledProcessError: {e}")
         except Exception as e:
             speak("An unexpected error occurred while attempting to restart.")
             print(f"Error: {e}")
     else:
         speak("Restart operation canceled.")

 

# Main logic to handle user queries

if __name__ == "__main__":
    wishMe()
    if 1:
        query = takeCommand().lower()

     # logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('searching wikipedia....')
            query = query.replace("wikipedia", "").strip()
            try:#error handling
                print(f"Searching for: {query}")#error handling
                results = wikipedia.summary(query, sentences=2) # it can take two sentence from wikipedia
                print(f"Results: {results}")#error handling
                speak("According to Wikipedia")
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:#error handling
                speak("There are multiple results for this query. Please be more specific.")#error handling
                print(f"Disambiguation Error: {e}")#error handling
            except wikipedia.exceptions.PageError:#error handling
                speak("I could not find any results for your query.")#error handling
            except Exception as e:#error handling
                speak("Something went wrong while searching Wikipedia.")#error handling
                print(f"Error: {e}")#error handling

        elif 'open google' in query:
          speak("What should I search on Google?")
          search_query = takeCommand().lower()
          if search_query != "none":  # Ensure a valid query is captured
              webbrowser.open(f"https://www.google.com/search?q={search_query}")
              speak(f"Here are the results for {search_query} on Google.")
          else:
             speak("I didn't catch that. Opening Google home page instead.")
             webbrowser.open("https://www.google.com")    

        elif 'open netflix' in query:
          webbrowser.open("netflix.com")

        elif 'open spotify' in query:
          webbrowser.open("spotify.com")

        elif 'open notepad' in query:
          open_application("notepad")

        elif 'close notepad' in query:
         close_application("notepad")

        elif 'open calculater' in query:
         open_application("calculater")

        elif 'close calculater' in query:
         close_application("calculater")

        elif 'open cmd prompt ' in query:
         open_application("cmd prompt")

        elif 'close command prompt' in query:
         close_application("command prompt")

        elif 'open explorer' in query:
         open_application("explorer")

        elif 'close explorer' in query:
         close_application("explorer")

        elif 'open camera' in query:
         open_application("camera")

        elif 'close camera' in query:
         close_application("camera")

        elif 'open setting' in query:
         open_application("setting")

        elif 'close setting' in query:
         close_application("setting")

        elif ' lock  computer' in query:
         lock_pc()

        elif 'shutdown computer' in query:
         shutdown_pc()

        elif 'restart computer' in query:
         restart_pc()

        

        elif 'play music' in query:
          Music = r'C:\Users\Arbaz\Music'

          if os.path.exists(Music):
             audios = [audio for audio in os.listdir(Music) if audio.endswith(('.mp3','.wav'))]
             if audios:
                random_song = random.choice(audios)#pick a random song
                print(f"playing: {random_song}")
                os.startfile(os.path.join(Music, random_song)) #play random song
                speak(f"playing {random_song}")
             else:
              speak("no music filefound in the specified directry")
              print("no music file found in the directry")
          else:
             speak("the music doest not exist.")
             print("the specified folder path does not exist")

          # Adjust volume
        elif 'volume high' in query:
            engine.setProperty('volume', 1.0)  # Set volume to 100%
            speak("Volume set to high.")

        elif 'volume low' in query:
            engine.setProperty('volume', 0.3)  # Set volume to 30%
            speak("Volume set to low.")

        elif 'volume medium' in query:
            engine.setProperty('volume', 0.7)  # Set volume to 70%
            speak("Volume set to medium.")    

        elif 'tell me what time ' in query:
           strTime = datetime.datetime.now().strftime("%I:%M:%S %p")
           speak(f"sir, the time is {strTime}")

        elif 'open vscode' in query:
           vscodePath = " E:\\Microsoft VS Code\\Code.exe "
           os.startfile(vscodePath)

        elif 'system information' in query:
          system_info()

        elif 'take screenshot' in query:
         take_screenshot()

        elif 'create test folder' in query:
            # Predefined folder creation logic
            folder_path = r"C:\Users\Arbaz\NewTestFolder"
            create_folder(folder_path)

        elif 'create folder ' in query:
             speak("Where should I create the folder? Please provide the full path.")
             folder_path = takeCommand().strip()  # Capture the folder path from the user
             print(f"Captured folder path: {folder_path}")  # Debugging statement

              # Validate the input
             if folder_path and folder_path.lower() != "none":
                 if os.path.isabs(folder_path):  # Check if the path is absolute
                     create_folder(folder_path)
                 else:
                     speak("The folder path provided is not valid. Please provide a full absolute path.")
                     print("The provided path is not absolute.")
             else:
                 speak("I didn't catch that or the input is invalid. Please try again.")



        elif 'send email' in query:
        
             try:
                  # Ask for recipient email
                 speak("Please tell me the recipient's email address.")
                 recipient_email = takeCommand().lower()

                  # Verify or re-confirm email
                 if '@' not in recipient_email or '.' not in recipient_email:
                  speak("The email address doesn't seem valid. Can you please spell it?")
                  recipient_email = takeCommand()

                  # Ensure email is valid
                 if '@' not in recipient_email or '.' not in recipient_email:
                   speak("I couldn't understand the email address. Please try again.")
                 else:
                  # Ask for subject
                   speak("What should be the subject of the email?")
                   subject = takeCommand()

                 # Ask for email body
                 speak("What should I say in the email?")
                 body = takeCommand()

                 # Construct Gmail compose link
                 gmail_url = (
                 f"https://mail.google.com/mail/?view=cm&fs=1&to={recipient_email}"
                 f"&su={subject}&body={body}"
                 )

                # Open Gmail in browser
                 webbrowser.open(gmail_url)
                 speak(f"I have opened Gmail with the email addressed to {recipient_email}. Please review and send it manually.")

             except Exception as e:
                 speak("I encountered an error while trying to open Gmail.")
                 print(f"Error: {e}")

         #if we send mail to specific person
        elif 'email to arbaz' in query:
             try:
                 speak("What should i say")
                 content = takeCommand()
                 to = "53530@students.riphah.edu.pk"
                 sendEmail(to, content)
                 speak("Email has been sent!")

             except Exception as e:
                print(e)
                speak("sorry !!.I am not able to send this mail")

              



             
           
          