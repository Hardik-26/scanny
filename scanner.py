#|==============================================================================================
#|                                                                                         Scanner Module                                                                                                         |
#|==============================================================================================

# Author : Hardik Shah.
# GitHub : https://github.com/Hardik-26

#|---------------------------------------------------------------------------------------------------------------------------------------------------------------------
# With this module you will be able to use your flatbed scanner using python.
# This module uses PowerShell script to execute the scan command.
# you can say that this module works as an API for Powershell which can comunicate with the scanner using WIA (Windows Image Acquisition).
#
# INSTRUCTIONS-
# Make sure your flatbed scanner is pluged in and ready to scan
# well thats it.
#
# HOW TO USE-
# import scanner
# scanner.StartScan("Path"( where you want your image file to be saved) , "ImageName" )
# Ex-  scanner.StartScan('C:\Users\\Admin\\Desktop','TestIamge')
#
# THANK YOU
#|---------------------------------------------------------------------------------------------------------------------------------------------------------------------
#|==============================================================================================


# PRE-REQUISITES--------------------------------------------------
_=False #To check if scanner is calibrated or not
Printer_Pixel_Data={}
#|--------------------------------------------------------------------

#|==============================================================================================
# CALIBRATE SCANNER------------------------------------------
def Calibrate():
    def start():
        _=True #To check if scanner is calibrated or not
        # IMPORTS------------------
        import os
        from PIL import Image 
        #------------------------------
        # Start Scan--------
        path= os.curdir
        StartScan(path,'TEMP')
        # Read Image-------
        with Image.open("TEMP.png") as img:
            dpi=round((img.info['dpi'])[1])
        global pixel_len
        pixel_len=2.54/dpi  #Length of one pixel in cm
        #os.remove('TEMP.png')
        print(' CALIBRTION COMPLETE')
        return pixel_len
    if _==True:
        print('!! RECALIBRATION WILL DELETE THE DAT FROM THE PREVIOUS CALIBRTION !!')
        choice=input('Do you want to proceed? (y/n) : ')
        if choice=='y' or 'Y':
            start()
        else:
            pass
    else:
        start()

#|==============================================================================================


# Start Scanning-----------------------------------------------------------------------------------------

def StartScan(Path,ImageName):

    # IMPORTS------------------
    import subprocess
    import os
    #------------------------------

    if os.path.exists(Path)==False:
        raise OSError(" Path Not Found / Path Does Not Exist ")
    else:
        pass

    #POWERSHELL SCRIPT ---------------------------------------------------------------------------
    Powershell_code='''# .Net methods for hiding/showing the console in the background
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
function Hide-Console
{
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
Hide-Console
$deviceManager = new-object -ComObject WIA.DeviceManager
$device = $deviceManager.DeviceInfos.Item(1).Connect()    

$wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"

foreach ($item in $device.Items) { 
    $image = $item.Transfer($wiaFormatPNG) 
}    

if($image.FormatID -ne $wiaFormatPNG)
{
    $imageProcess = new-object -ComObject WIA.ImageProcess
    $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
    $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatPNG
    $image = $imageProcess.Apply($image)
}

$image.SaveFile("'''+ str(Path)+'\\'+str(ImageName)+'.png'+'"'+')'

    Powershell_code= str(Powershell_code)
    
#-------------------------------------------------------------------------------------------------------
    
    powershell_file= open('scanner.ps1','w')
    powershell_file.write(Powershell_code)
    powershell_file.close()
    process = subprocess.Popen('powershell.exe -ExecutionPolicy RemoteSigned -file "scanner.ps1"', stdout= subprocess.PIPE) # Execute Powershell file
    process.communicate()
    os.remove('scanner.ps1')


#|==============================================================================================

# Get size of the scanner bed----------------------------

def size():
    # IMPORTS------------------
    import subprocess
    import os
    from PIL import Image
    #------------------------------
    if _==True: #To check if scanner is calibrated or not
        with Image.open("TEMP.png") as img:   # open image using PIL
            width,height= img.size # get image size
            dpi=round((img.info['dpi'])[1]) #Get DPI
        os.remove('TEMP.png')
    else:
        path= os.curdir
        StartScan(path,'TEMP')    # scan a temp image
        with Image.open("TEMP.png") as img:   # open image using PIL
            width,height= img.size # get image size
            dpi=round((img.info['dpi'])[1]) #Get DPI
        os.remove('TEMP.png')
    print(width,'px',',',height,'px')
    print(width/dpi,'in',',',height/dpi,'in')
#|==============================================================================================

# Get mesurements of scanned objects--------------------------

def getlength():
    
    if _==False:    #To check if scanner is calibrated or not
        print(' You need to calibrate your scanner for this process')
        print('\n Follow the calibration process in README.md to calibrate your scanner.\n')
    else:
        pass
    
    # IMPORTS------------------
    import os
    import numpy as np
    import matplotlib.pyplot as plt
    try:
        import cv2
    except ModuleNotFoundError as error:
        print(' This Process Requires Open-cv2.')
        print(' Please Install Open-cv2 -> "pip install opencv-python"')
        print(error)
    #-----------------------------

    
 #
