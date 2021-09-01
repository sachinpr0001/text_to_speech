from gtts import gTTS
import docx
import os
import speech_recognition as sr
import pyttsx3
i = 1
def main():
    try:
        doc = docx.Document('input.docx')
        data = ""
        fullText = []
        for para in doc.paragraphs:
            fullText.append(para.text)
            data = '\n'.join(fullText)
 
        mytext = data
        language = 'en'

        myobj = gTTS(text=mytext, lang=language, slow=False)
        myobj.save("output_audio.mp3")
        os.system("output_audio.mp3")
 
    except IOError:
        print('There was an error opening the file!')
        return

r = sr.Recognizer()
def SpeakText(command):
    engine = pyttsx3.init()
    engine.say(command)
    engine.runAndWait()
while(1):
    try:
        
        with sr.Microphone() as source2:
            r.adjust_for_ambient_noise(source2, duration=0.2)
            if i == 1:
                print('[NOTE : The word document must be named "input".]')
                print('-------------------------------------------------------')
                print('Please speak "Read Document pdd 174"')
                SpeakText('Please speak, read document pdd 174')
                i=2
            audio2 = r.listen(source2)
            MyText = r.recognize_google(audio2)
            MyText = MyText.lower()
            
            if MyText == "read document pdd 174":
                print("File Generated Succesfully. Opening the output file now.")
                SpeakText("File Generated Succesfully. Opening the output file now")
                if __name__ == '__main__':
                    main()
                break
    except sr.RequestError as e:
        print("Could not request results; {0}".format(e))
    except sr.UnknownValueError:
        print("Could not hear you. Please speak again.")
        SpeakText("Could not hear you. Please speak again.")



