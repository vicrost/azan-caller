import tkinter as tk
from tkinter import Grid
import xlrd
import os
from xlrd import cellname,xldate_as_tuple,xldate_as_datetime
from datetime import date
from datetime import datetime , timedelta
from playsound import playsound
from tkinter import font
import pygame




#set time and current date
current_time = datetime.strftime(datetime.now(), "%I:%M %p")  #time.strftime("%I:%M %p")
todays_date = date.today()
month = todays_date.strftime('%m')

def current_date(todays_date):
    format_date  = todays_date.strftime("%d/%m/%Y")
    return format_date


format_date = current_date(todays_date)

#function to display current time
def display_time():
    current_time = datetime.strftime(datetime.now(), "%I:%M %S %p")
    lbl_display_time.config(text=current_time)
    lbl_display_time.after(1000,display_time)

#function to display current date
def display_date():
    todays_date = date.today()
    month = todays_date.strftime('%m')
    format_date = todays_date.strftime("%d/%m/%Y")
    lbl_display_date.config(text=format_date)
    lbl_display_date.after(1000,display_date)






#open workbook
workbook = xlrd.open_workbook(r'prayer_timings.xlsx')
worksheet = workbook.sheet_by_name('Sheet %s'% month)


    
def get_date(format_date):
    #number of rows in sheet
    rows = worksheet.nrows
    #get date from excel file and compare with current date
    raw_date = ''
    for i in range(-1,rows):
        raw_date = worksheet.cell_value(i,1)
        conv_date = datetime(*xlrd.xldate_as_tuple(raw_date, workbook.datemode))
        excel_date = conv_date.strftime("%d/%m/%Y")
        if format_date == excel_date:
            cell_location = (i,1)
            break
    return [cell_location , excel_date]
        

def calc_countdown():
    ''' This function calculates the remaining time before the next prayer '''

    

    #converts prayer times to datetime object
    struc_prayertime_list = [datetime.strptime(i, "%I:%M %p") for i in azan_time]

    #gets the current time and converts it to datetime object
    current_time = datetime.strptime(datetime.strftime(datetime.now(), "%I:%M:%S %p"), "%I:%M:%S %p")

    #gets the next prayer time
    nearest_time = next((i for i in struc_prayertime_list if i > current_time), datetime.strptime('01-02 04:18 AM', "%m-%d %I:%M %p"))

    #calculates the time left before next prayer, converts it to string and stores it in a variable
    time_left_text.set("{} hours {} minutes {} seconds".format(*str(nearest_time - current_time).split(":")))

    random_lbl.after(1000, calc_countdown)





#get time for different prayers
def fajr_time(cell_location):
    fajr = worksheet.cell_value(cell_location[0][0],3)
    return fajr

def zuhr_time(cell_location):
    zuhr = worksheet.cell_value(cell_location[0][0],5)
    return zuhr

def asr_time(cell_location):
    asr = worksheet.cell_value(cell_location[0][0],6)
    return asr

def maghrib_time(cell_location):
    maghrib = worksheet.cell_value(cell_location[0][0],7)
    return maghrib

def isha_time(cell_location):
    isha = worksheet.cell_value(cell_location[0][0],8)
    return isha

def update_time():
    ''' this function will run continuoulsy and its function is to keep updating the prayer times
    as the day and date changes    flow : get the correct cell_location using get_Date()  --- set prayer times
    using cell_location --- append prayer times(variables) to a list called azan_time--- return azan_time '''

    cell_location = get_date(format_date)

    
    #set prayer times
    fajr = fajr_time(cell_location)
    fajr="0"+fajr
    zuhr  = zuhr_time(cell_location)
    asr = asr_time(cell_location)
    maghrib = maghrib_time(cell_location)
    isha = isha_time(cell_location)
    
    #store prayer times in a list
    azan_time = [fajr,zuhr,asr,maghrib,isha]

    #checks if any of the time is in PM. IF so, a '0' is added to the beginning
    azan_time.remove(zuhr)  #temporarily remove zuhr time
    for x in range(0,len(azan_time)):
        temp_time = azan_time[x]
        if temp_time[-2]+temp_time[-1] == 'PM':
            azan_time[x] = '0'+azan_time[x]

    azan_time.insert(1,zuhr) # add zuhr time back

    random_lbl.after(1000,update_time)

    return azan_time  #return azan time so it can be used



#function to check if it is time for prayer
def call_azan():
    current_time = datetime.strftime(datetime.now(), "%I:%M %p")
    for i in range(0,len(azan_time)):
        pygame.mixer.init()
        if current_time ==  azan_time[i] and pygame.mixer.music.get_busy() == False:
            azan_path = os.path.join("azan caller","azan_sound.mp3")
            pygame.mixer.music.load(azan_path)
            pygame.mixer.music.play()
            
            #playsound(azan_path)
        else:
            continue
    random_lbl.after(1000,call_azan)



#GUI
window  = tk.Tk()
window.title("Azan Caller")
Grid.rowconfigure(window, 0, weight=1)
Grid.columnconfigure(window, 0, weight=1)



#set up window
height = window.winfo_screenheight()
width = window.winfo_screenwidth()
window.geometry(f'{int(width/2)}x{int(height/2)}')
frm_window = tk.Frame(window,bg="#800000")
frm_window.pack(fill="both",expand=True)

# add icon to page
window.iconbitmap(r'images\icon.ico')

# a label just to run the functions continuously after 1 second
random_lbl = tk.Label(frm_window,width=2)

#create fonts
heading_font = font.Font(family = "ShareTech",size=25,weight='bold')
subheading_font = font.Font(family="ShareTech",size=14)
body_font = font.Font(family="ShareTech",size=10,weight="bold")


# title section
frm_title = tk.Frame(frm_window,relief=tk.SOLID,borderwidth=3,bg="#000000")
frm_title.grid(row=0, column=0, sticky="new",pady=(0,5))
lbl_title = tk.Label(master=frm_title, text="Azan Caller For Abu Dhabi",font=heading_font,fg="#008000")
lbl_title.pack(side='top',fill='x')

                #prayer timings table
frm_table = tk.Frame(frm_window,bg="#000000",relief=tk.SOLID,borderwidth=3)
frm_table.grid(row=1,column=0, sticky="nsew",padx=(20,20))
tbl_title  = tk.Label(frm_table,text="Today's Prayer Times",relief=tk.SOLID,font=subheading_font,borderwidth=2)
tbl_title.grid(row=0,column=0,columnspan=5,sticky="nsew")

#prayer names
prayer_names = ['Fajr','Zuhr','Asr','Maghrib','Isha']
for j in range(0,len(prayer_names)):
    name = prayer_names[j]
    cell = tk.Label(frm_table,text=name,relief=tk.SOLID,font=body_font,borderwidth=2)
    cell.grid(row=1,column=j,sticky="nsew")

#prayer times
azan_time = update_time()
for i in range(0,len(azan_time)):
    prayer_time = azan_time[i]
    cell_time = tk.Label(frm_table,text=prayer_time,relief=tk.SOLID,font=body_font,borderwidth=2)
    cell_time.grid(row=2,column=i,sticky="nsew")
        
#body section
frm_body = tk.Frame(frm_window,relief=tk.SOLID,borderwidth=5,)
frm_body.grid(row=2, column=0, sticky="nsew",pady=(10,10),padx=(20,20))

lbl_date = tk.Label(frm_body,text="Today's Date: ",font=body_font)
lbl_date.grid(row=0,column=0,sticky="nsew",pady=(0,5))
lbl_display_date = tk.Label(frm_body,font=body_font)
lbl_display_date.grid(row=0,column=1,sticky="nsew",pady=(0,5))


lbl_time = tk.Label(frm_body,text="Current Time: ",font=body_font)
lbl_time.grid(row=1,column=0,sticky="nsew",pady=(0,5))
lbl_display_time = tk.Label(frm_body,font=body_font)
lbl_display_time.grid(row=1,column=1,sticky="nsew",pady=(0,5))

time_left_text = tk.StringVar()
lbl_countdown = tk.Label(frm_body,text="Next Prayer in : ",font=body_font)
lbl_countdown.grid(row=2,column=0,sticky="nsew",pady=(0,5))
lbl_display_countdown = tk.Label(frm_body,textvariable=time_left_text,font=body_font,fg="red")
lbl_display_countdown.grid(row=2,column=1,sticky="nsew",pady=(0,5))

#footer
frm_footer  = tk.Frame(frm_window,relief=tk.SOLID,borderwidth=3,bg="#000000")
frm_footer.grid(row=3, column=0, sticky="sew",pady=(10,0))
lbl_footer = tk.Label(frm_footer,text="@Copyright AbdulRahman Tijani 2020",font=body_font,fg="#008000")
lbl_footer.pack(side="bottom",fill='x')


                    #make the structure resizable
#columns
for i in range(frm_window.grid_size()[0]):
    frm_window.grid_columnconfigure(i, weight=1)

for i in range(frm_table.grid_size()[0]):
    frm_table.grid_columnconfigure(i, weight=1)

for i in range(frm_body.grid_size()[0]):
    frm_body.grid_columnconfigure(i,weight=1)

#row
for i in range(frm_window.grid_size()[1]):
    frm_window.grid_rowconfigure(i, weight=1)

for i in range(frm_table.grid_size()[1]):
    frm_table.grid_rowconfigure(i, weight=1)


for i in range(frm_body.grid_size()[1]):
    frm_body.grid_rowconfigure(i,weight=1)



#run functions
display_date()
display_time()
update_time()
call_azan()
calc_countdown()



window.mainloop()

