#|==============================================================================================
#|                                         ScanneModule                                        |
#|==============================================================================================

''' Copyright (c) 2022 Hardik Shah.

    Copyright (C) 2007 Free Software Foundation, Inc. <https://fsf.org/>
    Everyone is permitted to copy and distribute verbatim copies
    of this license document, but changing it is not allowed. '''

#|-----------------------------------------------------------------------------------------------------------------------------------------

# Author : Hardik Shah.
# GitHub : https://github.com/Hardik-26

#|-----------------------------------------------------------------------------------------------------------------------------------------

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
# Ex-  scanner.StartScan('C:\\Users\\Admin\\Desktop','TestIamge')
#
# THANK YOU
#|-------------------------------------------------------------------------------------------------------------------------------------------

#|==============================================================================================


# PRE-REQUISITES--------------------------------------------------
_calibration=False #To check if scanner is calibrated or not
#|----------------------------------------------------------------

#|==============================================================================================

# CALIBRATE SCANNER------------------------------------------
''' To get the pixel size in cm so that we can scan any object
    and measure it's dimensions (Height & Width).
    Brief description of algo:
    1. scan a 'temp' image, get its size and DPI.
    2. Let DPI be 100, therefore 100 pixels per inch or 100 pixels per 2.57 cm.
    3. do basic calculations to get the size of one pixel in cm.'''

def Calibrate():
    def start():
        global _calibration
        _calibration=True #To check if scanner is calibrated or not
        # IMPORTS----------------------
        import os
        from PIL import Image 
        #------------------------------
        # Start Scan--------
        path= os.curdir
        StartScan(path,'TEMP')
        # Read Image-------
        with Image.open("TEMP.png") as img:
            dpi=round((img.info['dpi'])[1]) # Get DPI of the Image
        global _pixel_len
        _pixel_len=2.54/dpi  #Length of one pixel in cm
        #os.remove('TEMP.png') #Test
        print(' CALIBRTION COMPLETE')
        return _pixel_len
    if _calibration==True:
        print('!! RECALIBRATION WILL DELETE THE DATA FROM THE PREVIOUS CALIBRTION !!')
        choice=input('Do you want to proceed? (y/n) : ')
        if choice=='y' or 'Y':
            start()
        else:
            pass
    else:
        start()

#|=========================================================================================================================================

# Start Scanning-----------------------------------------------------------------------------------------
'''To initiate the scan process, This funtion when called will-
1. create a Powershell script to communicate with the flatbed scanner using WIA
2. execute the PS script.
3. save the image in the given location.'''

def StartScan(Path,ImageName):
    # IMPORTS----------------------
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


#|=========================================================================================================================================

# Get size of the scanner bed----------------------------------------------------
''' This process will get the size of the flatbed scanner in pixels and in cm.
Algo:
1. scan a Temp image using the StartScan funtion.
2. get the image size in pixels.
3. if the scanner has been calibrated then give size in cm also.'''

def size():
    # IMPORTS----------------------
    import subprocess
    import os
    from PIL import Image
    #------------------------------
    if _calibration==True: #To check if scanner is calibrated or not
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
#|=========================================================================================================================================
