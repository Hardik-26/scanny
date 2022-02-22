# Hinding Console-->
#-------------------------------------------------------------------
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

#-------------------------------------------------------------------

# Selecting device using WIA--->                           # WIA- "Windows Image Acquisition"
$deviceManager = new-object -ComObject WIA.DeviceManager
$device = $deviceManager.DeviceInfos.Item(1).Connect()    

#Png format--->
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

#Saving Image--------> # .png file only
$image.SaveFile("C:\Users\Admin\Desktop\TestImage.png")