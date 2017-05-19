' Create a variable that holds the IP address for the printer.
' The following code shows you how to create an IP variable and set it to a value:
Dim IPpath
IPpath = "10.12.2.35"

' Create a mapped printer variable and set it to the IP printer.
' The following code adds the IP printer to the list of available devices in the user's Control Panel:
Set printer = CreateObject("WScript.Network") 
printer.AddWindowsPrinterConnection IPpath

' Close the VBScript file. The "Quit" function stops the VBScript process and closes the file.
' The following code completes the script:
WScript.Quit