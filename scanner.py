#|================================================================================================================================================================================
#|                                                                                         Scanner Module                                                                                                         |
#|================================================================================================================================================================================

# Author : Hardik Shah.
# GitHub : https://github.com/Hardik-26

#|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
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
#|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#|================================================================================================================================================================================


# PRE-REQUISITES--------------------------------------------------
_=False #To check if scanner is calibrated or not
#|----------------------------------------------------------------

#|==============================================================================================

# CALIBRATE SCANNER------------------------------------------
''' To get the pixel size in cm so that we can scan any object
    and measure it's dimensions (Height & Width).
    Brief description of algo:
    1. scan a temp image, get its size and DPI.
    2. Let DPI be 100 therefore 100 pixels per inch or 100 pixels per 2.57 cm.
    3. do basic calculations to get the size of one pixel in cm.'''

def Calibrate():
    def start():
        global _
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
            dpi=round((img.info['dpi'])[1]) # Get DPI of the Image
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
'''To initiate the scan process, This funtion when called will-
1. create a Powershell script to communicate with the flatbed scanner using WIA
2. execute the PS script.
3. save the image in the given location.'''

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

# Get size of the scanner bed----------------------------------------------------
''' This process will get the size of the flatbed scanner in pixels and in cm.
Algo:
1. scan a Temp image using the StartScan funtion.
2. get the image size in pixels.
3. if the scanner has been calibrated then give size in cm also.'''

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
''' This funtion will read a image for objects draw a square around it
    and get the dimentions of the object in cm.
    This will be done using the help of Open-CV.'''

def MeasureObject(ImagePath='',__='Leave ImagePath empty to scan a new image and mesure that.'):
    Temp = None # Refrencing Temp Becaues why not.
    if _==False:    #To check if scanner is calibrated or not
        print(' You need to calibrate your scanner for this process')
        raise SystemError(' Use "scanner.Calibarate()" to calibrate Scanner')
    else:
        pass
    
# IMPORTS--------------------------------------------------------------------------------------------------
    import os
    import numpy as np
    from PIL import Image
    try:
        import cv2
    except ModuleNotFoundError as error:
        print(' This Process Requires Open-cv2.')
        print(' Please Install Open-cv2 -> "pip install opencv-python"')
        print(error)
#-------------------------------------------------------------------------------------------------------------
    Home_path=os.environ["HOMEPATH"]+'\Desktop' # Get HomePath of User
    def Measure(image_path):
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5,5), 0)
        thresh = cv2.threshold(blur, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        x,y,w,h = cv2.boundingRect(thresh)  # get cordinates for object
        cv2.rectangle(image, (x, y), (x + w, y + h), (36,255,12), 2) #Draw Box Around Object
        width=w*pixel_len #Converting pixels to cm
        height=h*pixel_len #Converting pixels to cm
        if Temp==True: #checking If image was resized or not
            cv2.putText(image, "Width={}cm,Height={}cm".format(round(width,1)*4,round(height,1)*4), #if image resized then w*4 & h*4
                          (round(w/2)-50,y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,0), 2)
        else:
            cv2.putText(image, "Width={}cm,Height={}cm".format(round(width,2),round(height,2)),
                              (round(w/2)-50,y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,0), 2)
        cv2.imshow("image", image)
        cv2.waitKey(0)
        return image
#-------------------------------------------------------------------------------------------------------------

    if ImagePath=='':       #To check if to scan or not
        Temp = True #To checking If image was resized or not
        StartScan(Home_path,'Image')
        with Image.open(Home_path+'\Image.png') as img:
            width,height= img.size
            image_c=img.crop((10,10,width-10,height-10))
            image_T = image_c.resize((round(image_c.width/4), round(image_c.height/4)), Image.ANTIALIAS)#Resizing Image to Fit in CV2
            image_T.save(Home_path+'\ResizedImage.png')
        Measure(Home_path+'\ResizedImage.png')
        os.remove(Home_path+'\ResizedImage.png')
        print("Image Saved On Desktop.")
        print("Image Location: ",Home_path+'\Image.png')
    else:
            if os.path.exists(ImagePath)==False:
                raise OSError(" Path Not Found / Path Does Not Exist ")
            else:
                Measure(ImagePath)
#|================================================================================================================================================================================
# TO BE CONTINUED...☺
