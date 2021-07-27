from xlrd import open_workbook
from random import randint
import pyttsx3
import time
from plyer import notification

class DailySentence:
    def __init__(self,filename):
        self.filename=filename
        self.engine = pyttsx3.init()
       

    def setting_voice_rate(self):
        rate = self.engine.getProperty('rate')   # getting details of current speaking rate
        # print (rate)                        #printing current voice rate
        self.engine.setProperty('rate', 120)     # setting up new voice rate

    def setting_volume_level(self):
        """VOLUME"""
        volume = self.engine.getProperty('volume')  #getting to know current volume level (min=0 and max=1)
        # print (volume)                          #printing current volume level
        self.engine.setProperty('volume',1.0)    # setting up volume level  between 0 and 

    def setting_voice_gender(self):
        """VOICE"""
        #voices = engine.getProperty('voices')       #getting details of current voice
        #engine.setProperty('voice', voices[0].id)  #changing index, changes voices. o for male
        #engine.setProperty('voice', voices[1].id)   #changing index, changes voices. 1 for female
    
    def save_voice(self):
         """Saving Voice to a file"""
        # On linux make sure that 'espeak' and 'ffmpeg' are installed
        # engine.save_to_file('Hello World', 'test.mp3')
        # engine.runAndWait()

    def set_notification(self):
        wb=open_workbook('sentence.xlsx')
        for sheet in wb.sheets():
            number_of_rows = sheet.nrows-1
            number_of_columns = sheet.ncols-1
            while(True):
                row=randint(0,number_of_rows)
                print(sheet.cell(row,number_of_columns).value)
                # for row in range(number_of_rows):
                #     for col in range(number_of_columns):
                #         print(sheet.cell(row,col).value)
                self.engine.say(sheet.cell(row,number_of_columns).value)
                self.engine.runAndWait()
                notification.notify(
                #title of the notification,
                title = "",
                #the body of the notification
                message = sheet.cell(row,number_of_columns).value,  
                #creating icon for the notification
                #we need to download a icon of ico file format
                app_icon = "",
                # the notification stays for 50sec
                timeout  = 500
                )
                time.sleep(60*15)

           

    


class_obj=DailySentence('sentence.xlsx')
class_obj.setting_voice_rate()
class_obj.set_notification()



    
   
        

# class_obj=DailySentence('sentence.xlsx')


