' VBSCRIPT TO SET PAGEFILE TO SYSTEM MANAGED
'
' VARIABLES:
'
' "D:" = Page File set to D Drive
' ""   = Disable Page File
'
' 0 and 0  = System Managed Size
' "" or "" = Use Existing Size
'
' DebugEnable = Shows Message Prompts
' ErrorNotify = Shows Script Error
'
' NOTES:
'
' If existing Page file is disabled,
' no New sizes set, and NewLocation 
' enabled - sets to System Managed.
'
' If Page File not yet created
' (requires reboot) even if you
' just set a new page file - it will
' be set to System Managed if you
' dont set New Sizes.
'

Option Explicit

'******** CHANGE THIS SETTING **********

Const NewLocation    = "C:"

'// OPTIONAL

Const DebugEnable    = True
Const ErrorNotify    = True
Const NewInitSize    = "0"
Const NewMaxSize     = "0"

'***************************************

Dim objWMIService
Dim colPageFiles
Dim objPageFile
Dim PFsize, PFRead

Const SystemSize     = 1
Const CustomSize     = 2
Const NoPageFile     = 3
const HKLM           = &H80000002

'******* MAIN SCRIPT *********

'// EXISTING PAGE FILE SIZES
PFRead = GetPageFileSize()

On Error Goto 0
On Error Resume Next

'// CHECK FOR NEW CUSTOM SIZES
If IsNumeric(NewInitSize) And IsNumeric(NewMaxSize) Then
   PFRead = NewInitSize & " " & NewMaxSize
End If

'// CHECK PAGE FILE INFO
If Len(NewLocation) = 2 And Right(NewLocation, 1) = ":" And _
   IsNumeric(NewLocation) = False Then
    If Len(PFRead) Then
        If PFRead = "0 0" Then
            PFSize = SystemSize
        Else
            PFSize = CustomSize
        End If
    Else
        PFRead = "0 0"
        PFSize = SystemSize
    End If
Else
    PFSize = NoPageFile
End if

'// SET PAGE FILE
Select Case PFSize
    Case 1
        If SetPageFile(NewLocation & "\pagefile.sys " & PFRead) Then
            If DebugEnable Then _
            MsgBox "PAGE FILE SET OKAY:" & vbCrlf & vbCrlf & _
                   "System Managed Page File" & vbCrlf & _
                   NewLocation & "\pagefile.sys " & PFRead
        Else
            If DebugEnable Then _
            MsgBox "Error Setting PageFile"
        End If
    Case 2
        If SetPageFile(NewLocation & "\pagefile.sys " & PFRead) Then
            If DebugEnable Then _
            MsgBox "PAGE FILE SET OKAY:" & vbCrlf & vbCrlf & _
                   "Custom Page File" & vbCrlf & _
                   NewLocation & "\pagefile.sys " & PFRead
        Else
            If DebugEnable Then _
            MsgBox "Error Setting PageFile"
        End If
    Case 3
        If SetPageFile("") Then
            If DebugEnable Then _
            MsgBox "PAGE FILE DISABLED"
        Else
            If DebugEnable Then _
            MsgBox "Error Setting PageFile"
        End If
End Select

'// NOTIFY IF SCRIPT ERROR
If Err And ErrorNotify Then
    MsgBox "Script Error Occurred:" & vbCrLf & _
    Err.Number & ": " & Err.Description
End If

'******** FUNCTIONS **********

'// GET PAGE FILE SIZE
Function GetPageFileSize()
    Dim iInit, iMax
    On Error Goto 0: On Error Resume Next
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colPageFiles = objWMIService.ExecQuery("Select InitialSize, MaximumSize from Win32_PageFile")
    For each objPageFile in colPageFiles
        iInit = objPageFile.InitialSize
        iMax  = objPageFile.MaximumSize
    Next
    If iInit <> "" And iMax <> "" And Err = 0 Then
        GetPageFileSize = iInit & " " & iMax
    End If
End Function

'// SET PAGE FILE AND SIZE
Function SetPageFile(pValue)
    Dim sKey, oReg
    Const HKLM = &H80000002
    On Error Goto 0: On Error Resume Next
    sKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    oReg.CreateKey HKLM, sKey
    oReg.SetMultiStringValue HKLM, sKey, "PagingFiles", Array(pValue)
    If Err = 0 Then SetPageFile = True
End Function