##################################################
import time
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import pandas as pd
import pygame
from pygame import mixer
from tkinter import *
from tkinter import filedialog
from datetime import datetime

import random
import os
import matplotlib
import soundfile as sf
import matplotlib.pyplot as plt

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,
NavigationToolbar2Tk)
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

#----------------------OAUth initialisation----------------#
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets'] #changed scope from below
SERVICE_ACCOUNT_FILE = 'registration-keys.json'
creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

#-----------------------------------------------------------#

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '16wytMNFyD6_uhyz3aE0KHj2R5CPYG9hVshWwd753aJE'
service = build('sheets', 'v4', credentials=creds)
SAMPLE_RANGE_NAME = 'Sheet1!A1:G'

#------------- Call the Sheets API ---------------------------#
sheet = service.spreadsheets()

#-----------------------------------------------------------#

# Main Screen
root = ThemedTk(theme='adapta')
root.title("Speaker/Language change detection")
height = root.winfo_screenheight()

# getting screen's width in pixels
width = root.winfo_screenwidth()
#print("\n width x height = %d x %d (in pixels)\n" %(width, height))

root.geometry("{}x{}".format(width, height))
#root.geometry("heightxwidth")
#root.minsize(2000,1500)
#root.maxsize(2000,1500)
root.resizable("False","False")
#root.tk.call('tk', 'scaling', '-displayof', '.', 100.0/72.0)
#root.iconbitmap("/img/Listen.ico")

# Intial Values
response = ''
Song_name = ''
Song_title = ''

song_no = 0
path = ''
######################################################################################################## Excelsheet Creation ##################################################################################################

# Creation of the temp Records excel sheet.
#outWorksheet = xlsxwriter.Workbook("./records/record.xlsx")
#outSheet = outWorksheet.add_worksheet()
#outSheet.write("A1", "Mother Language")
#outSheet.write("B1", "AudioName")
#outSheet.write("C1", "Response")
#outSheet.write("D1", "Date&Time")
#outWorksheet.close()


########################################################################################################### Functions #########################################################################################################
#16wytMNFyD6_uhyz3aE0KHj2R5CPYG9hVshWwd753aJE
############ Start Functions, to initiate the program

def start1():
    global song_title,path
    global entry1,qno,canvas,entry,name
    global Label1,start_btn1,start_btn2,Label1
    global start_time
    global filename1,filename2
    global session
    session = '1'
    path = "./Speaker_change_data/"
    qno = int(entry1.get())
    entry1.destroy()
    name = entry2.get()
    entry2.destroy()
    start_btn1.destroy()
    start_btn2.destroy()
    T1.destroy()
    T2.destroy()
    T3.destroy()
    T4.destroy()
    Label1.destroy()
    Label.destroy()
    #frame4.destroy()
    starting1()
    filename1 = os.listdir(path)
    filename1 = sorted(filename1)
    #song_title = random.sample(filename,21)
    #print(song_title)
    song_title = filename1

    play_btn['state']=tk.NORMAL
    #plot_btn['state']=tk.NORMAL

def start2():
    global song_title,path
    global entry1,qno,canvas,entry,name
    global Label1,start_btn1,start_btn2,Label1
    global start_time
    global session
    session = '2'
    path = "./Language_change_data/"
    qno = int(entry1.get())
    entry1.destroy()
    name = entry2.get()
    entry2.destroy()
    start_btn1.destroy()
    start_btn2.destroy()
    T1.destroy()
    T2.destroy()
    T3.destroy()
    T4.destroy()
    Label1.destroy()
    Label.destroy()
    #frame4.destroy()
    starting2()
    filename2 = os.listdir(path)
    filename2 = sorted(filename2)
    #song_title = random.sample(filename,21)
    song_title = filename2
    play_btn['state']=tk.NORMAL
    #plot_btn['state']=tk.NORMAL
############ Function to Play the AudioFiles

def play_song1():
    global song_title
    global path,canvas
    global Song_name
    global start_time
    global replay
    global Label2, Label6
    replay = 0
    start_time = time.time()
    Label2 = tk.Label(root, text="Did you detect the Speaker Change Point?", font=10, bg="white")
    Label2.pack(side=BOTTOM)
    Label2.place(x=width/3, y=(height/2)-(height/24))
    Label6 = tk.Label(root, text="Question : "+str(song_no+qno)+"/"+str(len(song_title)), font=20, bg="white")
    Label6.pack(side=BOTTOM)
    Label6.place(x=4*(width/5), y=(height/8))
    if qno+song_no==(len(song_title)+1):
        ending()
    else:
        try:
            Song_name = str(song_title[song_no+qno-1])
            data, fs = sf.read(path+str(Song_name))
            fig = Figure(figsize = (10, 3),dpi = 80)
            plot1 = fig.add_subplot(111)
            plot1.plot(data)
            canvas = FigureCanvasTkAgg(fig,master = root)
            canvas.draw()
            canvas.get_tk_widget().pack()
            #toolbar = NavigationToolbar2Tk(canvas,root)
            #toolbar.update()
            #canvas.get_tk_widget().pack()
            mixer.init()
            mixer.music.load(path+str(Song_name))
            mixer.music.play()
            play_btn['state'] = tk.DISABLED
            yes_btn['state'] = tk.NORMAL
            no_btn['state'] = tk.NORMAL
            replay_btn['state'] = tk.NORMAL
            next_btn['state'] = tk.DISABLED
        except Exception as e:
            print(e)

def play_song2():
    global song_title
    global path,canvas
    global Song_name
    global start_time
    global replay
    global Label2, Label6
    replay = 0
    start_time = time.time()
    Label2 = tk.Label(root, text="Did you detect the Language Change Point?", font=10, bg="white")
    Label2.pack(side=BOTTOM)
    Label2.place(x=width/3, y=(height/2)-(height/24))
    Label6 = tk.Label(root, text="Question : "+str(song_no+qno)+"/"+str(len(song_title)), font=20, bg="white")
    Label6.pack(side=BOTTOM)
    Label6.place(x=4*(width/5), y=(height/8))
    if qno+song_no==(len(song_title)+1):
        ending()
    else:
        try:
            Song_name = str(song_title[song_no+qno-1])
            data, fs = sf.read(path+str(Song_name))
            fig = Figure(figsize = (10, 3),dpi = 80)
            plot1 = fig.add_subplot(111)
            plot1.plot(data)
            canvas = FigureCanvasTkAgg(fig,master = root)
            canvas.draw()
            canvas.get_tk_widget().pack()
            #toolbar = NavigationToolbar2Tk(canvas,root)
            #toolbar.update()
            #canvas.get_tk_widget().pack()
            mixer.init()
            mixer.music.load(path+str(Song_name))
            mixer.music.play()
            play_btn['state'] = tk.DISABLED
            yes_btn['state'] = tk.NORMAL
            no_btn['state'] = tk.NORMAL
            replay_btn['state'] = tk.NORMAL
            next_btn['state'] = tk.DISABLED
        except Exception as e:
            print(e)

def replay_song():
            mixer.init()
            mixer.music.load(path+str(Song_name))
            mixer.music.play()
            play_btn['state'] = tk.DISABLED
            replay_btn['state'] = tk.NORMAL
            yes_btn['state'] = tk.NORMAL
            no_btn['state'] = tk.NORMAL
            next_btn['state'] = tk.DISABLED
            global replay
            replay +=1

############ Funtion to record the 'Yes' responses

def yes():
    global response
    global song_no,Label5
    global end_time
    global time_dur, time_stamp
    mixer.music.stop()
    end_time=time.time()
    Label2.destroy()
    Label5 = tk.Label(root, text="Click Next", font=10, bg="white")
    Label5.pack(side=BOTTOM)
    Label5.place(x=width/2.5, y=(height/2)-(height/24))
    response = "Yes"
    play_btn['state'] = tk.DISABLED
    replay_btn['state'] = tk.DISABLED
    yes_btn['state'] = tk.DISABLED
    no_btn['state'] = tk.DISABLED
    next_btn['state'] = tk.NORMAL
    time_dur = end_time-start_time
    now = datetime.now()
    time_stamp = now.strftime("%d/%m/%Y %H:%M:%S")
    song_no += 1
    submit_btn()


############ Funtion to record the 'No' responses

def no():
    global response
    global song_no,Label5
    global end_time
    global time_dur,time_stamp
    mixer.music.stop()
    end_time=time.time()
    Label2.destroy()
    Label5 = tk.Label(root, text="Click Next", font=10, bg="white")
    Label5.pack(side=BOTTOM)
    Label5.place(x=width/2.5, y=(height/2)-(height/24))
    response = "No"
    no_btn['state'] = tk.DISABLED
    yes_btn['state'] = tk.DISABLED
    replay_btn['state'] = tk.DISABLED
    next_btn['state'] = tk.NORMAL
    song_no += 1
    time_dur = end_time-start_time
    now = datetime.now()
    time_stamp = now.strftime("%d/%m/%Y %H:%M:%S")
    submit_btn()

############ Export the Response and Audio Filename in Excel Spreadsheet

def submit_btn():
    Label5.destroy()
    canvas.get_tk_widget().pack_forget()
    #print(Song_name, response,time_dur,replay)
    aoa = [[name,session,Song_name,response,time_dur,replay,time_stamp]]
    request = sheet.values().append(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="Sheet1!A1", valueInputOption="USER_ENTERED",insertDataOption="INSERT_ROWS", body={"values":aoa}).execute()

def next1():
    play_btn['state'] = tk.NORMAL
    no_btn['state'] = tk.DISABLED
    yes_btn['state'] = tk.DISABLED
    replay_btn['state'] = tk.DISABLED
    next_btn['state'] = tk.DISABLED

def next2():
    play_btn['state'] = tk.NORMAL
    no_btn['state'] = tk.DISABLED
    yes_btn['state'] = tk.DISABLED
    replay_btn['state'] = tk.DISABLED
    next_btn['state'] = tk.DISABLED

########### Function to Export the Excel Sheet

#def export():
#    global df1
#    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
#    df1.to_excel(export_file_path, index=False, header=True)
    #os.remove("./records/record.xlsx")

########### Exporting the Excelsheet

def ending():
    global next_btn
    global saveAs
    play_btn.destroy()
    replay_btn.destroy()
    yes_btn.destroy()
    no_btn.destroy()
    Label2.destroy()
    frame5.destroy()
    Label3.destroy()
    Label5.destroy()
    Label6.destroy()
    #frame4 = Frame(root, bg="white", bd=4, width=500, height=200,borderwidth=2,relief="groove")
    #frame4.place(x=220,y=250)
    end = tk.Label(text="Experiment is Complete Your Response has been Recorded", font=("Proxima Nova", 12),bg="white")
    end.place(x=width/3, y=(height/2)-(height/24))
    next_btn.destroy()
    #saveAs = ttk.Button(text='Export Results', command=export)
    #saveAs.place(x=420,y=340)


########################################################################################################### GUI CODE###########################################################################################################

#~~~~~~ GUI TEXT ~~~~~~~#

global T,T1,T2,T3,T4,frame4
frame1 = Frame(root,bg = "teal",width=width,height=height/10,relief="raised",bd=4)
frame1.pack(side=TOP)
#frame2 = Frame(root,bg = "white",bd=4,width=2000,height=500,relief="raised")
#frame2.pack(side=BOTTOM)

T = tk.Label(root, text="Listening Test",font=("Proxima Nova",35),justify=CENTER,bg ="teal",fg="white")
T.pack(side=TOP)
T.place(x=width/2.5, y=height/40)

T1 = tk.Label(root, text="Step 1 : Enter sentence you want to start from (starting from 1)",font=("Proxima Nova",12),justify=LEFT,bg = "white")
T1.pack(side=TOP)
T1.place(x=width/3, y=height/8)
T2 = tk.Label(root, text="Step 2 : Click 'Start' and listen to a fixed set of audio files",font=("Proxima Nova",12),justify=LEFT,bg = "white")
T2.pack(side=TOP)
T2.place(x=width/3, y=(height/8)+(height/24))
T3 = tk.Label(root, text="Step 3 : If you detect the speaker/ language change point, give your response in the form of 'YES/NO'",font=("Proxima Nova",12),justify=LEFT,bg = "white")
T3.pack(side=TOP)
T3.place(x=width/3, y=(height/8)+(height/12))
#frame4 = Frame(root, bg="white", bd=4, width=500, height=200,borderwidth=2,relief="groove")
#frame4.place(x=500,y=700)
T4 = tk.Label(root, text="(Use headphones or earphones if available!)",font=("Proxima Nova",12),justify=LEFT,bg = "white",fg="red")
T4.pack(side=TOP)
T4.place(x=width/3, y=(height/8)+(height/8))


#~~~~~ Buttons ~~~~~~~#

global Label1,Label

Label = tk.Label(root, text="Enter your name",font=10,bg="white")
Label.pack(side=BOTTOM)
Label.place(x=width/2.5, y=(height/3))
Label.grid_columnconfigure(0,weight=1)
entry2 = ttk.Entry(root, width=10)
entry2.grid_columnconfigure(1,weight=1)
entry2.place(x=width/2.5, y=(height/3)+(height/24))
Label1 = tk.Label(root, text="Enter Sentence you want to start from",font=10,bg="white")
Label1.pack(side=BOTTOM)
Label1.place(x=width/2.5, y=(height/3)+(height/12))
Label1.grid_columnconfigure(0,weight=1)
entry1 = ttk.Entry(root, width=10)
entry1.grid_columnconfigure(1,weight=1)
entry1.place(x=width/2.5, y=(height/3)+(height/8))
start_btn1 = ttk.Button(root, text="Start session 1 [Speaker change]", command=start1, state=tk.NORMAL)
start_btn1.place(x=width/2.5, y=(height/3)+(height/6))
start_btn1.grid_columnconfigure(2,weight=1)
start_btn2 = ttk.Button(root, text="Start session 2 [Language change]", command=start2, state=tk.NORMAL)
start_btn2.place(x=width/2.5, y=(height/3)+(height/6)+(height/24))
start_btn2.grid_columnconfigure(2,weight=1)

def starting1():
    global play_btn, yes_btn,no_btn,next_btn,Label3,frame5,replay_btn, plot_btn
    frame5 = Frame(root, bg="white", bd=4, width=width/3.5, height=height/4, borderwidth=2, relief="groove")
    frame5.place(x=width/3, y=height/2)
    Label3 = tk.Label(root, text="Please Listen Carefully. Do not wait between replays", font=10, bg="white")
    Label3.pack(side=BOTTOM)
    Label3.place(x=width/3, y=(height/2)-(height/12))
    #plot_btn = tk.Button(root, text="Plot", command=plot, state=tk.DISABLED)
    #plot_btn.pack(side=BOTTOM)
    #plot_btn.place(x=970, y=200)
    play_btn = ttk.Button(root, text="Play", command=play_song1, state=tk.DISABLED)
    play_btn.pack(side=TOP,fill=BOTH,expand="true")
    play_btn.place(x=width/2.5, y=(height/2)+(height/24))
    replay_btn = ttk.Button(root, text="replay", command=replay_song, state=tk.DISABLED)
    replay_btn.pack(side=TOP,fill=BOTH,expand="true")
    replay_btn.place(x=width/2, y=(height/2)+(height/24))
    yes_btn = ttk.Button(root, text="Yes", command=yes,state=tk.DISABLED)
    yes_btn.pack(side=LEFT)
    yes_btn.place(x=width/2.5, y=(height/2)+(height/12))
    no_btn = ttk.Button(root, text="No", command=no, state=tk.DISABLED)
    no_btn.pack(side=RIGHT)
    no_btn.place(x=width/2, y=(height/2)+(height/12))
    next_btn = ttk.Button(root, text="Next", command=next1, state=tk.DISABLED)
    next_btn.pack(side=BOTTOM)
    next_btn.place(x=width/2.25, y=(height/2)+(height/6))

def starting2():
    global play_btn, yes_btn,no_btn,next_btn,Label3,frame5,replay_btn, plot_btn
    frame5 = Frame(root, bg="white", bd=4, width=width/3.5, height=height/4, borderwidth=2, relief="groove")
    frame5.place(x=width/3, y=height/2)
    Label3 = tk.Label(root, text="Please Listen Carefully. Do not wait between replays", font=10, bg="white")
    Label3.pack(side=BOTTOM)
    Label3.place(x=width/3, y=(height/2)-(height/12))
    #plot_btn = tk.Button(root, text="Plot", command=plot, state=tk.DISABLED)
    #plot_btn.pack(side=BOTTOM)
    #plot_btn.place(x=970, y=200)
    play_btn = ttk.Button(root, text="Play", command=play_song2, state=tk.DISABLED)
    play_btn.pack(side=TOP,fill=BOTH,expand="true")
    play_btn.place(x=width/2.5, y=(height/2)+(height/24))
    replay_btn = ttk.Button(root, text="replay", command=replay_song, state=tk.DISABLED)
    replay_btn.pack(side=TOP,fill=BOTH,expand="true")
    replay_btn.place(x=width/2, y=(height/2)+(height/24))
    yes_btn = ttk.Button(root, text="Yes", command=yes,state=tk.DISABLED)
    yes_btn.pack(side=LEFT)
    yes_btn.place(x=width/2.5, y=(height/2)+(height/12))
    no_btn = ttk.Button(root, text="No", command=no, state=tk.DISABLED)
    no_btn.pack(side=RIGHT)
    no_btn.place(x=width/2, y=(height/2)+(height/12))
    next_btn = ttk.Button(root, text="Next", command=next2, state=tk.DISABLED)
    next_btn.pack(side=BOTTOM)
    next_btn.place(x=width/2.25, y=(height/2)+(height/6))

root.mainloop()

#################################################################################################### Program End ###################################################################################################################
