 strComputer = "."
 Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

 Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 For Each objComputer in colSettings 
    computerName = CStr(objComputer.Name)
 Next
 Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystemProduct") 
 For Each objItem in colItems 
    serialNumber = CStr(objItem.IdentifyingNumber)
 Next

info = serialNumber & " " & computerName 

Set fso = CreateObject("Scripting.FileSystemObject")
set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
strDesktop = strDesktop & "\serial_number_and_computer_name.txt"
Set MyFile = fso.CreateTextFile(strDesktop, True)
MyFile.WriteLine(info)
MyFile.Close	
