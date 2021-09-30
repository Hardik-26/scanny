#|==============================================================================================
#|                                                                                         Scanner Module                                                                                                         |
#|==============================================================================================

# Module By : Hardik Shah.

#|---------------------------------------------------------------------------------------------------------------------------------------------------------------------
# With this module you will be able to use your flatbed scanner using python.
# This module uses PowerShell script to execute the scan command.
# you can say that this module works as an API for Powershell.
#
# INSTRUCTIONS-
# Make sure your flatbed scanner is pluged in and ready to scan
# Change the Execution Policy in Powershell for your computer.
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


# PRE REQUSITS--------------------------------------------------
_=False #To check if scanner is calibrated or not
Printer_Pixel_Data={}
#|--------------------------------------------------------------------


# CALIBRATE SCANNER------------------------------------------
def Calibrate():
    _=True #To check if scanner is calibrated or not
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
    #------------------------------

    # Start Scan--------
    path= os.curdir
    StartScan(path,'TEMP')
    # Read Image-------
    image = cv2.imread("TEMP.png")
    # convert to grayscale--------
    #grayscale = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # perform edge detection---------
    edges = cv2.Canny(image, 400, 650)
    # detect lines in the image using hough lines technique-----
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, 10, np.array([]), 500, 50)
    line_length= abs(((lines[0])[0])[1]-((lines[0])[0])[3])  #---- I can't explain this line in comments,
                                                                                  #---if you dont understand this line then mail me @dennishardik.2673@gmail.com
                                                                                  #---i will try to explain it to you. ☺
    #print(line_length)

    global pixel_len
    pixel_len=5/line_lenght
    print(' CALIBRTION COMPLETE')

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
    Powershell_code='''$deviceManager = new-object -ComObject WIA.DeviceManager
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
    path= os.curdir
    StartScan(path,'TEMP')    # scan a temp image
    with Image.open("TEMP.png") as img:   # open image using PIL
        width,height= img.size # get image size
    os.remove('TEMP.png')
    print(width,'px',',',height,'px')

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
    #------------------------------
    
   #
