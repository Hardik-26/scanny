# Scanny

[![PyPI version](https://badge.fury.io/py/scanny.svg)](https://badge.fury.io/py/scanny)
[![License](https://img.shields.io/badge/License-GNU%20General%20Public%20License%20v3.0-blue)](LICENSE)
[![Downloads](https://static.pepy.tech/personalized-badge/scanny?period=total&units=none&left_color=grey&right_color=brightgreen&left_text=Total%20Downloads)](https://pepy.tech/project/scanny)

--------------------------------------------------------------------------------------------------- 

   ### Module for Python. This Module will enable a user to use any flatbed scanner using python.

--------------------------------------------------------------------------------------------------- 


### Copyright (c) 2024 Hardik Shah.


--------------------------------------------------------------------------------------------------- 
### Author : Hardik Shah.
### License : GNU General Public License v3.0

--------------------------------------------------------------------------------------------------- 
### Brief Description :
With this module you will be able to use your flatbed scanner using python.<br>
This module uses PowerShell script to execute the scan command.<br>
you can say that this module works as an API for Powershell which can communicate with the scanner using WIA (Windows Image Acquisition).<br>


--------------------------------------------------------------------------------------------------- 
### INSTRUCTIONS-
Make sure your flatbed scanner is pluged in and ready to scan <br>
well thats it.

---------------------------------------------------------------------------------------------------
## HOW TO USE-

`import scanny` <br>

### Scan A image- 
``` 
>>> scanny.StartScan("Path"( where you want your image file to be saved) , "ImageName" ) 
Eg- scanny.StartScan('C:\\Users\\Admin\\Desktop','TestIamge')
```

### Calibrate the scanner-
(Optional, only for backward compatibility purposes)-
```
>>> scanny.Calibrate()
```

### Get Scanner Bed Size-
```
>>> scanny.size()
```

### MISC-
```
>>> scanny._calibration
>>> scanny._pixel_len
```
---------------------------------------------------------------------------------------------------
### Other Information-
> Currently this python package only works for Windows 8 & Above. <br>
> Linux and MAC-OS compatibility comming soon in next version: 0.0.2. <br>

---------------------------------------------------------------------------------------------------

# Thank you. ☺
