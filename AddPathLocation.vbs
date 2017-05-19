Set WshShell = WScript.CreateObject("WScript.Shell") 
Set WshEnv = WshShell.Environment("SYSTEM") 
WshEnv("Path") = WshEnv("Path") & ";C:\Program Files\Java\jdk1.7.0_06\bin"
WshEnv("CLASSPATH") = WshEnv("CLASSPATH") & ";T:\Java"