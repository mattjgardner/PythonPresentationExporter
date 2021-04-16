import sys  
import os
from os import path  
import win32com.client
import easygui

##select the PowerPoint you wish to export to image using the fileopenbox
##PowerPoint should open, export all the images to a folder of the same name
##as your presentation in the same folder as the presentation then close.
file = easygui.fileopenbox()


papp = win32com.client.Dispatch("Powerpoint.Application")  
papp.Visible = 1
pres = papp.Presentations.Open(file)


filename = os.path.splitext(file)[0]


if path.exists(filename) == True:
    pass
else:
    os.mkdir(filename)

i = 0
noOfSlides = len(pres.Slides)
for i in range(noOfSlides):
    slideName = 'slide' + str(i) + '.png'
    newfile = filename + "\\" + slideName
    pres.Slides[i].Export(newfile, "PNG")
    i += 1
    
    
pres.Close()
papp.Quit()

pres = None
papp = None

