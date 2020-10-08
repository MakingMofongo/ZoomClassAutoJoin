## Code to Automatically open Zoom program to join a meeting and record meeting is required


import pyautogui
import time
import schedule
import socket
import speech_recognition as sr
import threading
import logging as log
import urllib.request
import pyaudio
import win32gui as W
import os
#from win32gui import GetWindowText, GetForegroundWindow

pyautogui.FAILSAFE=False



log.basicConfig(level=log.INFO, filename='Logger.log', filemode='w', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

#####   Joining Zoom Meeting   ###################

#######################################################################################
c_id="0"
c_pass="0"
ongoing=False
t=0



#esc clicked to ensure that the win key will open up correctly in the next step
def start(meet_id,password,*kwargs):
    print("function triggered at:",ct())
    log.info('Join meeting function triggered')
    global c_id
    global c_pass
    global ongoing
    global t
    c_id=meet_id
    c_pass=password

    #Starting Screen Recording on Nvidia
    #also checks if recording on nvidia started sucessfully, otherwise, starts recording on game bar 
    GameBarRecording=False
    try:
        pyautogui.hotkey('alt','f9')
        time.sleep(3)
        x,y=pyautogui.locateCenterOnScreen('RecordingHasStarted.png')
        print(ct(),"recording started")
    except TypeError:
        GameBarRecording=True
        log.warning('Nvidia recording aint working, will start windows game bar recording')
    
    time.sleep(2)
    pyautogui.press('esc',interval=0.1)
    pyautogui.press('win',interval=0.1)
    time.sleep(1)
    pyautogui.write('zoom')
    time.sleep(1)
    pyautogui.press('enter',interval=0.5)
    print(ct(),"zoom opened")


    #time delay to factor for zoom app to load up, good buffer is like 10 sec but its case specific
    time.sleep(5)


    pyautogui.moveTo(5,5)
    try:
        x,y = pyautogui.locateCenterOnScreen('joinButton.png')
        pyautogui.click(x,y)
    except TypeError:
        pyautogui.click(653,250)

    time.sleep(2)
    pyautogui.write(meet_id)
    pyautogui.press('enter',interval=1)
   
   
    time.sleep(1)
    if password!='0':
        pyautogui.write(password)
        pyautogui.press('enter',interval = 1)


    ongoing=True
    print(ct(),"meeting joined")

    time.sleep(5)

    try:
        x,y = pyautogui.locateCenterOnScreen('JoinAudio.png')
        pyautogui.click(x,y)
        pyautogui.press('esc',interval=1)
    except TypeError:
        pass

    if GameBarRecording==True:
        
        x=0
        while meeting_active==False and x<30:
            time.sleep(10)
            x=x+1

        pyautogui.hotkey('win','alt','r')
        log.info('Windows GameBar recording started')
    



    #### recording time amount + reconnect if internet lost during meeting
    t=0
    t=t+x #time taken for starting windows game bar recording, >0 if meeting took a while to get started AND nvidia recording didnt work
    #t=t+ kwargs.get('t_c',None)  idk how to do this yet              
    while t<482:   #1t= 10s, 482t = 80 mins
        if meeting_active()==False & t>30 & t<360: #check between first 5-60 mins
            # (OBSOLETE) if VSopen() != True :
            os.system("TASKKILL /F /IM Zoom.exe")
            reconnect()
            t=t+3
        time.sleep(20)              
        t=t+2



    ## ending screen recording)
    if GameBarRecording==False:
        pyautogui.hotkey('alt','f9')
        print(ct(),"Nvidia recording stopped")
        log.info('Nvidia recording stopped')
    
    ## closing Zoom
    if VSopen() != True:
        os.system("TASKKILL /F /IM Zoom.exe")
    ongoing=False
    print(ct(),"zoom closed")
    


def meeting_active():
#    pyautogui.click(240,240)
#    pyautogui.moveTo(300,300,1)
#    match=pyautogui.pixelMatchesColor(300,869,(26,26,26))
    if (W.GetWindowText(W.GetForegroundWindow())) == "Zoom Meeting" :
        match = True
    else:
        match = False    
    
    if match==False:
        print(ct(),"Meeting not active")
        log.warning('MEETING NOT ACTIVE')
    elif match==True:
        log.info('Meeting is active')
    return match



def VSopen():
    
    if "Visual Studio Code" in W.GetWindowText(W.GetForegroundWindow()) : 
        return True
    else:
        return False




#Check if internet connection is NOT working
def is_not_connected():
    try:
        # connect to the host -- tells us if the host is actually
        # reachable
        socket.create_connection(("www.google.com", 80))
        socket.create_connection(("www.microsoft.com",80))
        return False
    except OSError:
        pass
    print("internet not working at:",ct())
    log.warning('Internet not working')
    return True

def connect(host='http://google.com'):
    try:
        urllib.request.urlopen(host) #Python 3.x
        return True
    except:
        return False

#this function takes <= 30 seconds
def reconnect():
    
    log.info('rejoining meeting')
    #deal with the zoom dialouge box for internet connection lost
    



   # i have no idea why this exists
   # if pyautogui.pixelMatchesColor(608,332,(14,113,235),tolerance=2)!=True:
    #    t_c=t
     #   start(c_id,c_pass,t_c)



 #rejoin the last connected zoom meeting
    time.sleep(2)
    pyautogui.press('esc',interval=0.1)
    pyautogui.press('win',interval=0.1)
    time.sleep(1)
    pyautogui.write('zoom')
    time.sleep(1)
    pyautogui.press('enter',interval=0.5)
    print(ct(),"zoom opened")


    #time delay to factor for zoom app to load up, good buffer is like 10 sec but its case specific
    time.sleep(5)

    pyautogui.moveTo(5,5)
    time.sleep(1)       
    
    try:
        x,y = pyautogui.locateCenterOnScreen('joinButton.png')
        pyautogui.click(x,y)
    except TypeError:
        pyautogui.click(653,250)
    


    pyautogui.press('enter',interval=1)
    pyautogui.write(c_id)
    pyautogui.press('enter',interval=1)


    time.sleep(3)
    pyautogui.press('enter',interval=1)
    pyautogui.write(c_pass)
    pyautogui.press('enter',interval = 1)

    print(ct(),"succesfully reconnected")
    log.info('Succesfully Reconnected')

def ct():
    c_time = time.localtime()
    current_time = time.strftime("(%H:%M:%S)", c_time)
    return current_time   







########################################################################################


################ Auto Scheduler #######################################################


#######################################################################################
#test_id=input("test id:")
#test_pass=input("test pass:")
#start(test_id,test_pass)

#Start of code exec
#print(W.GetWindowText(W.GetForegroundWindow()))














#input ids and pass
print("Check if mic and speaker are set")
time.sleep(2)

p=pyaudio.PyAudio()
SpeakerInfo=p.get_default_output_device_info()
if SpeakerInfo["name"] != "CABLE Input (VB-Audio Virtual C" :
    print("WARNING: default audio device is '",SpeakerInfo["name"])

else:
    print("default audio device is '",SpeakerInfo["name"])
cbs_id=input("CBS ID:")
pvk_id=input("PVK ID:")
mpsk_id=input("MPSK ID:")
all_pass=input("PASS:")
time.sleep(1)
print(ct(),"Input succesful")

#schedule meetings
schedule.every().day.at("08:45").do(start,cbs_id,all_pass) #CBS
schedule.every().day.at("10:10").do(start,pvk_id,all_pass) #PVK
schedule.every().day.at("11:35").do(start,mpsk_id,all_pass) #MPSK
print(ct(),"scheduling successful")

log.info('Input and Scheduling successful')

def test():
    test=False
    test=input("Test? (True?):")

    if test=="True":   
        #test_id=input("Test ID:")
        test_pass=input("Test Pass:")
        
        start('3666504520',test_pass)
    else:
        print('Ok, testing cancelled.')


#loop checks for time every ten seconds
def CheckForTime():
    while True:
        if connect()==False:
            time.sleep(10)
            continue
        schedule.run_pending()

        time.sleep(10)

##############################################
########## Raise Hand Code ###################

r=sr.Recognizer()
mic = sr.Microphone()
print("Listening for name")
log.info('Voice Recognition code block executed')


def ra():


    response = {
        "success": True,
        "error": None,
        "transcription": "nada"
    }
 


    with mic as source:
        audio = r.listen(source, phrase_time_limit=3)
        r.adjust_for_ambient_noise(source)

    try:
        response["transcription"] = r.recognize_google(audio)
    except sr.RequestError:
        # API was unreachable or unresponsive
        response["success"] = False
        response["error"] = "API unavailable"
        print("Google Speech API unavailable")
    except sr.UnknownValueError:
        # speech was unintelligible
        response["error"] = "Unable to recognize speech"

    return response["transcription"]


def raise_hand():
    pyautogui.hotkey('alt','y')
    print(ct(),"hand raised")
    log.info('hand raised')
    time.sleep(5)
    pyautogui.hotkey('alt','y')
    print(ct(),"hand down")
    log.info('hand down')


#check for name being called out and call raise_hand() if true
def CheckForName():
    while True:
        phrase=ra()
        callout=False
        Name=["Abdul Rashid","Abdul","Adil Rashid","Rashida","Rasheed","Isa","Rashid"]
        x=0
        if phrase != "nada":
            print(ct(),phrase)
        
        while x<7:
            if Name[x] in phrase:
                callout=True
                break
            x=x+1
      
        if callout==True:
            raise_hand()

thread1 = threading.Thread(target=CheckForName, args=(), daemon=True)
thread1.start()
test()
CheckForTime()