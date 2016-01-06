from __future__ import unicode_literals
from __future__ import print_function

import csv

# pywinauto-fullcam.py
# 20160107-001
# crgentle@gmail.com

# Setup:
# 1) Install Canopy and use the canopy CLI to "pip install pywinauto"
# 2) Change the variables in the block below
# 3) Run
# 4) ???
# 5) Profit!
#
#

##############################################
#
# Change these variables to suit your local environment
#
#


# Application to launch - you can see I didn't test to see if the double-\s were really required...

fullCamPath = r"C:\Program Files (x86)\FullCAM Public Release 2015\FullCam.exe"

# Where the *.plo files are stored

plotFilePath = "Z:\\Desktop\\fullcam-data\\"

# A list of plot file names

plotFileList = ["High Reafforestation", "Med Reafforestation", "Low Reafforestation"]

destinationDir = "Z:\\Desktop\\dest-dir\\"

inputCsv = plotFilePath + "TESTchris20151228.csv"

# !!! Probably not a good idea to change anything after this line !!!
#
#####################################################

try:
    from pywinauto import application
except ImportError:
    import os.path
    pywinauto_path = os.path.abspath(__file__)
    pywinauto_path = os.path.split(os.path.split(pywinauto_path)[0])[0]
    import sys
    sys.path.append(pywinauto_path)
    from pywinauto import application

app = application.Application() 
app.start(fullCamPath) 

app.FullCAMLicenseAndDisclaimer.IAcceptTheLicenseAndDisclaimer.ClickInput()

app.FullCAMLicenseAndDisclaimer.WaitNot('visible')


app.AboutFullCAM.WaitNot('visible')
  
    

for fname in plotFileList:    
    with open(inputCsv, 'rb') as csvfile:
        taskReader = csv.DictReader(csvfile)
        mainWin = app.FullCAM
    
        # Open the *.plo files listed above
        
        mainWin.TypeKeys("^o")
        
        # Wait for the dialog to open
        
        app.OpenFullCAMDocument.Wait("exists enabled visible ready")
        
        win = app.OpenFullCAMDocument
        # Type the file name 
        win.TypeKeys(plotFilePath + fname + ".plo", with_spaces=True)
        win.Open.ClickInput()
    
        for row in taskReader:
            print(row['ID'], row['Latitude'], row['Longitude'], row['Property'])
    
            plotWin = app[fname + '.plo']
            plotWin.SetFocus()
            plotWin.Wait("exists enabled visible ready")
            
            # Enter lat/long and click out of each text field after we do or the application won't recognise a change in the values
    
            plotWin.Edit3.SetEditText(row['Latitude'])
            plotWin.Edit2.Click()
    
            plotWin.Edit2.SetEditText(row['Longitude'])
            plotWin.Edit3.Click()
    
            plotWin.Wait("exists enabled visible ready")
                        
            # Press the download data button
            
            plotWin["Download Spatial Data"].ClickInput()
            app.Info.Wait("exists enabled visible ready", timeout=60)
            
            # Click ok - thanks for downloading, but you need me to OK every time?
            
            app.Info.Ok.ClickInput()
            app.Info.WaitNot('visible')
            
            # Now run simulation by pressing F9
            
            mainWin.TypeKeys('{F9}')
            
            simWindow = app[fname + ".plo - Output 1"]
            simWindow.Wait("exists enabled visible ready")
            
            # It took forever to narrow down the coordinates of the save dialog and teh keyboard shortcut 
            # wasn't wired up by the application developers so I couldn't shortcut with ^s
    
            #            simWindow.TToolBar.ClickInput(coords=(675,5))
            simWindow.TToolBar.ClickInput(coords=(320,5))
            
            # wait for the dsave dialog to appear
            
            app.SaveSimulationResults.Wait("exists enabled visible ready")
            
            # Save to the destination dir with a filename of <ID> - <plo name>.xls
            
            app.SaveSimulationResults.TypeKeys(destinationDir + row['ID'] + " - " + fname + ".xls", with_spaces=True)
            
            app.SaveSimulationResults.Save.ClickInput()
            app.SaveSimulationResults.WaitNot('visible')
            
            # Get out of the simulation window
            simWindow.TypeKeys('{ESC}')
                
            
            
            
            
            
            
            
