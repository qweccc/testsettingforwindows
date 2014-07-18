strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48) 
For Each objItem in colItems 
    Wscript.Echo "" & objItem.Model
Next