import pandas as pd
import sys
import win32api, win32con
import ctypes
from tkinter import Tk
from time import sleep

#VARIABLES YOU CAN CHANGE!!!!!!!!!!!!!!!!
dlay = 1			#delay between mouse and keyboard actions in seconds
SaveName = "result.txt"	#Saved File Name

#2D variable locations
mnloc = ([1011,253])	#Notes loc to paste material id
mnback = ([343,67])
mdrun = ([1097,403]) #Distributor loc to paste material id
nsdrag = ([213,407])	#note drag start loc
nedrag = ([1779,949])	#note drag end loc
dsdrag = ([237,421])	#distributor drag start loc
dedrag = ([1159,963])	#distributor drag end loc
mdback = ([277,371])
mdtc = ([917,509])

#VARIABLES YOU SHOULD NOT CHANGE!!!!!!!!!!!!!!
cb = Tk()	#cb variable using tkinter to manipulate clipboard contents for copy and paste
user32 = ctypes.windll.user32	#used for keyboard events

#FUNCTIONS!!!!
#place the input data onto the clipboard, can then be pasted onto screen as needed with ctrl-v aka paste function
def copy(data):
	cb.withdraw()
	cb.clipboard_clear()
	cb.clipboard_append(data)
	cb.update()

#drag select with mouse between inputs and copy the selected data, try to return clipboard data else return a space
def mousecopy(start,end):
	result = None	#return variable
	count = 0		#count for repeat attempts
	
	while (result == None):
		count += 1
		win32api.SetCursorPos(start)#move mouse to loc
		sleep(dlay)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,start[0],start[1],0,0) #click left mouse button
		sleep(dlay)
		win32api.SetCursorPos(end)#move mouse to loc
		sleep(dlay)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,end[0],end[1],0,0) #release left mouse button
		sleep(dlay)
		user32.keybd_event(0x11, 0, 0, 0) #Ctrl
		sleep(dlay)	
		user32.keybd_event(0x43, 0, 0, 0) #c
		sleep(dlay)
		user32.keybd_event(0x11, 0, 2, 0) #~Ctrl
		sleep(dlay)
		cb.update()	#update the clipboard
		click(start)
		try:													#try to get the data on the clipboard, if there was nothing to copy this will fail and move to the except block
			result = cb.selection_get(selection = "CLIPBOARD")	#if successful result will not be None and this will break the loop
		except:													#If the above code failed check if the number of attempts (aka count) is sufficient and if so set result to a space to break the loop
			if count > 2:
				result = " "
	return result	#return the data in result

#paste the contents of the clipboard to the current focused widget/location		
def paste():
	sleep(dlay)
	user32.keybd_event(0x11, 0, 0, 0) #Ctrl
	sleep(dlay)
	user32.keybd_event(0x56, 0, 0, 0) #v
	sleep(dlay)
	user32.keybd_event(0x11, 0, 2, 0) #~Ctrl

#click the input location once with the mouse
def click(loc):
	win32api.SetCursorPos(loc)	#move mouse to loc
	sleep(dlay)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,loc[0],loc[1],0,0)	#click left mouse button
	sleep(dlay)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,loc[0],loc[1],0,0)		#release left mouse button

def tclick(loc):
	win32api.SetCursorPos(loc)	#move mouse to loc
	sleep(0.1)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,loc[0],loc[1],0,0)	#click left mouse button
	sleep(0.05)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,loc[0],loc[1],0,0)		#release left mouse button
	sleep(0.05)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,loc[0],loc[1],0,0)	#click left mouse button
	sleep(0.05)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,loc[0],loc[1],0,0)		#release left mouse button
	sleep(0.05)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,loc[0],loc[1],0,0)	#click left mouse button
	sleep(0.05)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,loc[0],loc[1],0,0)		#release left mouse button
	sleep(0.05)
	user32.keybd_event(0x08, 0, 0, 0) #Backspace press
	sleep(0.05)
	user32.keybd_event(0x08, 0, 2, 0) #Backspace press	
	
	
	
#Alt Tab to working window
def alttab():
	user32.keybd_event(0x12, 0, 0, 0) #Alt
	sleep(dlay)
	user32.keybd_event(0x09, 0, 0, 0) #Tab
	sleep(dlay)
	user32.keybd_event(0x09, 0, 2, 0) #~Tab
	sleep(dlay)
	user32.keybd_event(0x12, 0, 2, 0) #~Alt

#alt tab tab to second working window
def alttabtab():
	user32.keybd_event(0x12, 0, 0, 0) #Alt
	sleep(dlay)
	user32.keybd_event(0x09, 0, 0, 0) #Tab
	sleep(dlay)
	user32.keybd_event(0x09, 0, 2, 0) #~Tab
	sleep(dlay)
	user32.keybd_event(0x09, 0, 0, 0) #Tab
	sleep(dlay)
	user32.keybd_event(0x09, 0, 2, 0) #~Tab
	sleep(dlay)
	user32.keybd_event(0x12, 0, 2, 0) #~Alt

#MAIN!!!!!!!!
if __name__ == "__main__":
	
	#alt tab to arrange windows for program to alt tab
	alttab()	#swap to the notes window
	alttabtab()	#swap to the distributor window
	alttab()	#swap to the notes window

	FileName = sys.argv[1]	#Get the name of the excel file, should be fed in from the command line
	sheet = 0				#Excel Sheet to use

	df = pd.read_excel(io=FileName, sheet_name=sheet)	#read in the excel file
	result = open(SaveName,"w")	#Open the file to save the results to

	rmax,cmax = df.shape	#rmax = max number of rows, cmax is discarded and is only there to break up the tuple return of df.shape


	for i in range(0,1):	#loop through each row in the excel file
		
		status = df.loc[i,"Status"]	#get the status column
		
		if(status[0] == "X" or status[1] == "X"):	#if the status column is a XOO or OXO do the following code
			
			matID = df.loc[i,"Material"]	#get the material id from the excel file at row i  (i is the counter that is incremeted in the loop! starts at row 0 then 1,2,3, etc.)
			copy(matID)	#copy the material id onto the clipboard so we can use control v to paste it
			
			result.write(df.loc[i,"Status"] + "\t" + matID + "\t" + df.loc[i,"MRP Area"] + "\t" + df.loc[i, "Description"] + "\t\t")	#write the data in coluns 1 through 4 from the excel file at row i to our new file
			paste()		#paste the material id in the text field using control v 
			user32.keybd_event(0x0D, 0, 0, 0) #Enter press
			sleep(0.05)
			user32.keybd_event(0x0D, 0, 2, 0) #Enter release
			sleep(dlay)
			click(mnloc)

			
			note = mousecopy(nsdrag,nedrag)	#drag select with mouse for notes data to copy
			result.write(note + "\t\t")	#write the copied notes data to the new file
			click(mnback)
			
			user32.keybd_event(0x08, 0, 0, 0) #Backspace press
			sleep(0.05)
			user32.keybd_event(0x08, 0, 2, 0) #Backspace press
			
			alttab()	#Swap to the distributor window
			
			copy(matID)		#put the material id back on the clipboard
			paste()		#paste the material id in the text field using control v
			click(mdrun)	#move the mouse and click Run Report
			sleep(dlay)
			
			note = mousecopy(dsdrag,dedrag)	#drag select with mouse for distributor data to copy
			result.write(note + "\n")	#write the copied distributor data to the new file
			tclick(mdtc)
			
			alttab()	#Swap to the notes window
	
	cb.destroy() #destroy the variable we created for manipulating the clipboard, api says to do so
		



