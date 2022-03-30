 'Machine Details to HTML File.
'Written 29/03/2011 by Gareth Pullen
'Modified 04/10/2011 to change the NetPath (new DC)
'Modified 17/02/2012 to add a Notes Inputbox.

'**************************'
'*** Defining Constants ***'
'**************************'
CONST NetPath = "\\<FQDN>\Laptops Checklist\Auto\" 'This is the base folder to copy to, + Year (must have trailing \).
CopyPath = NetPath & Year(Date) & "\" 'Append the year to get the right folder.
'***************************'
'*** Start of Functions: ***
'***************************'

Function MachineDetails(Details)
'Pulls the make and model (according to WMI) of the machine.
'Also pulls the CPU info of the machine.
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
Set ColSpeed = objWMIService.ExecQuery("select * from Win32_Processor")
For each objItem in colItems
Select case Details
Case "Make"
MachineDetails = objItem.Manufacturer
Case "Model"
MachineDetails = objItem.Model
Case "RAM"
MachineDetails = (((objItem.TotalPhysicalMemory) /1048576) & " MB")
Case "CPU Speed"
for each CPUInfo in ColSpeed
MaxSpeed = CPUInfo.MaxClockSpeed
CurrentSpeed = CPUInfo.CurrentClockSpeed
if MaxSpeed = CurrentSpeed then
MachineDetails = CurrentSpeed
else
MachineDetails = "Current speed; " & CurrentSpeed & "</br>" & vbCrLf & "Max Speed; " & MaxSpeed
End if
Next
Case "CPU Model"
for each CPUInfo in ColSpeed
CPUName = CPUInfo.Name
MachineDetails = CPUInfo.Name
Next
Case else
MachineDetails = "Please choose some details"
End select
Next

End Function

'Function to get MAC Address.
'Written by Gareth Pullen - 29/03/2011
Function GetMacAddress
dim MACAddys 'Used to store the addresses before returning.
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2") '"." means local machine.
Set colItems = objWMIService.ExecQuery _
("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objItem in colItems
MACAddys = MACAddys & objItem.Description & "</br>" & vbCrLf & "<strong>" & " MAC Address: " & "</strong>" & objItem.MACAddress & "</br>" & vbCrLf
Next

GetMacAddress = MACAddys 'Got to return the addresses...

End Function

Function UserHostName(UserORHost)
'This gets the username or hostname of the current user and machine, and returns it.
'Call like: "UserHostName("Hostname")" or "UserHostName("Username")" and it will return the result:
'Example: Hostname = UserHostName("Hostname")
dim WshNetwork, Hostname, Username
Set WshNetwork = WScript.CreateObject("WScript.Network")
Select case UserOrHost
Case "Hostname"
Hostname = WshNetwork.ComputerName
UserHostName = Hostname
Case "Username"
Username = WshNetwork.UserName
UserHostName = Username
Case else
UserHostName = "Error - specify user or host name"
End select

End Function

Function AskUsersName
dim UsersName
'Done as a function as it (should) clear the memory afterwards.
Do
UsersName = InputBox("Who is this machine for ?:" & vbCrLf & "(Required input)", "Name ?")
if UsersName = "" then 'Need some way to abort...
QuitScript = MsgBox("Quit ?", vbYesNo, "Quit ?") 'Can't evaluate this directly...
if QuitScript = vbYes then
Wscript.Quit '*** Remember this will quit here! ***
end if
end if
loop while UsersName = ""
AskUsersName = UsersName
End Function

Function BiosInfo
'This function reads information from the BIOS using the WMI interface.
dim BInfo
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colBIOS = objWMIService.ExecQuery _
("Select * from Win32_BIOS")

For each objBIOS in colBIOS
BInfo = "<strong>" & "Manufacturer: " & "</strong>" & objBIOS.Manufacturer & "</br>" & vbCrLf
BInfo = BInfo & "<strong>" & "Serial number: " & "</strong>" & objBIOS.SerialNumber & "</br>"
Next
BiosInfo = BInfo
End Function


Function WriteToHTML(Title, Body, FName)
'Write to a HTML file.
'Call as: WriteToHTML "Title", "Body", "File-Name"
Set objFSO = CreateObject("Scripting.FileSystemObject") ' set + Call library to allow FS Manipulation
Set objNewFile = objFSO.CreateTextFile(FName)
dim Main, Top, Tail, Complete
Top = "<HTML>" & VbCrLf & "<Head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</Head>" & vbCrLf & "<Body>" & vbCrLf & "<H1>" & Title & "</H1>" & vbCrLf
Tail = "</body>" & vbCrLf & "</HTML>"

Complete = Top & Body & Tail
objNewFile.WriteLine Complete

End Function

Function CheckCreateFolder(Path)
'Function to check if a folder exists, and if not create it.
'End result = Folder will exist.
'Takes "Path" as the folder path to check.
Set objFSO = CreateObject("Scripting.FileSystemObject") ' set + Call library to allow FS Manipulation
if objFSO.Folderexists(Path) then
CheckCreateFolder = "True" 'Return true that it exists.
else
'objFSO.CreateFolder(Path)
Set objShell = CreateObject("Wscript.Shell")
objShell.Run "cmd /c mkdir " & """" & Path & """"
CheckCreateFolder = "True" 'Folder now exists.
end if

End Function

Function CopyFile(FileLocation, FileDestination, OverWrite, Cleanup)
'File Copy Function.
'Takes input file name(with path), destination, overwrite (True, False, or D-Version), "Cleanup" - True or False, delete the original after copying.
'Example call:
'CopyFile "C:\Test.html", "C:\Testing.html", "D-Vers-IfExist", False
Set objFSO = CreateObject("Scripting.FileSystemObject") ' set + Call library to allow FS Manipulation

if not objFSO.FileExists(FileLocation) then 'Sanity check - does the file exist ? (True means no!)
CopyFile = "Error! File doesn't exist!"
end if
Select Case OverWrite

Case "Yes"
if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
ObjFso.CopyFile FileLocation, FileDestination, OverWrite
else
CopyFile = "Error! File doesn't exist!"
end if

Case "No"
if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
ObjFso.CopyFile FileLocation, FileDestination, OverWrite
else
CopyFile = "Error! File doesn't exist!"
end if

Case "D-Version"
FileDestination = Replace(FileDestination, ".html", " - " & (Replace(Date, "/", "-") & ".html"))
if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
ObjFso.CopyFile FileLocation, FileDestination, False
else
CopyFile = "Error! File doesn't exist!"
end if

Case "D-Vers-IfExist" 'Date version if file exists...
if objFSO.FileExists(FileDestination) then
FileDestination = Replace(FileDestination, ".html", " - " & (Replace(Date, "/", "-") & ".html"))
if objFSO.FileExists(FileDestination) then 'Does a date-named version exist ?
'If so, add time to the file name.
if (len(minute(Time))) = 1 then
Min = "0" & minute(Time)
else
Min = minute(Time)
end if
FileDestination = Replace(FileDestination, ".html", " - " & (Hour(Time) & "." & Min & ".html"))
if objFSO.FileExists(FileDestination) then 'Does a hour/minute-named version exist ?
'If so, add seconds ?
FileDestination = Replace(FileDestination, ".html", "." & (Second(Time) & ".html"))
End if
else
'If not, copy it over...
ObjFso.CopyFile FileLocation, FileDestination, False
end if
if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
ObjFso.CopyFile FileLocation, FileDestination, False
else
CopyFile = "Error! File doesn't exist!"
end if
else 'File doesn't exist in destination
if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
ObjFso.CopyFile FileLocation, FileDestination, False
else
CopyFile = "Error! File doesn't exist!"
end if
End if

End Select

'Cleanup the old file
if Cleanup then 'Check to see if we need to cleanup.
if objFso.FileExists(FileLocation) then 'Check the file exists
objFso.DeleteFile(FileLocation) 'Delete the file.
end if
end if

end function

Function MemType(NumberType)
'Simple function for a Case statement.

Select case NumberType

Case 0
MemType = "Unknown"

Case 1
MemType = "Other"

Case 2
MemType = "SIP"

Case 3
MemType = "DIP"

Case 4
MemType = "ZIP"

Case 5
MemType = "SOJ"

Case 6
MemType = "Proprietary"

Case 7
MemType = "SIMM"

Case 8
MemType = "DIMM"

Case 9
MemType = "TSOP"

Case 10
MemType = "PGA"

Case 11
MemType = "RIMM"

Case 12
MemType = "SODIMM"

Case 13
MemType = "SRIMM"

Case 14
MemType = "SMD"

Case 15
MemType = "SSMP"

Case 16
MemType = "QFP"

Case 17
MemType = "TQFP"

Case 18
MemType = "SOIC"

Case 19
MemType = "LCC"

Case 20
MemType = "PLCC"

Case 21
MemType = "BGA"

Case 22
MemType = "FPBGA"

Case 23
MemType = "LGA"
End select

End Function


Function MachineRAM
'Get the currently installed RAM sticks
'Along with the total possible sticks.
'Taken from: http://www.wisesoft.co.uk/scripts/vbscript_installed_memory_modules-free_slots.aspx
'Hacked about by Gareth Pullen - 31/03/2011

strMemory = ""
i = 1
SET objWMIService = GETOBJECT("winmgmts:\\.\root\cimv2")
SET colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")

FOR EACH objItem In colItems

IF strMemory <> "" THEN strMemory = strMemory & vbcrlf
strMemory = strMemory & "Bank" & i & " : " & (objItem.Capacity / 1048576) & " Mb" & vbCrLf & "Type: " & MemType(objItem.MemoryType) & "</br>"
i = i + 1
NEXT
installedModules = i - 1

SET colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemoryArray")

FOR EACH objItem in colItems
totalSlots = objItem.MemoryDevices
NEXT
RetRep = "<strong> Total Slots: </strong>" & totalSlots & "</br>" & vbCrLf
RetRep = RetRep & "<strong> Free Slots: </strong>" & (totalSlots - installedModules) & "</br>" & vbcrlf
RetRep = RetRep & "<strong> Installed Modules: </strong>" & "</br>" & vbcrlf & strMemory & "</br>"

'Testing
'MsgBox(RetRep)
MachineRAM = RetRep
End Function

Function GroupMembers(GroupName)
'Function to return the members of a specified group.
'Written by Gareth Pullen 06/04/2011
'Call like: GroupMembers("Administrators")

GroupMems = ""
strComputer = "."
Set colGroups = GetObject("WinNT://" & strComputer & "")
colGroups.Filter = Array("group")
For Each objGroup In colGroups
if objGroup.Name = GroupName then
For Each objUser in objGroup.Members
GroupMems = GroupMems & objUser.Name & "</br>" & vbCrLf
Next
end if
Next
GroupMembers = GroupMems
End Function

'*******************************'
'****** End of Functions. ******'
'*******************************'

'********************************'
'** Script proper starts here: **'
'********************************'

dim Report, WordArray, NameOfUser, MachineModel, FName, Title, Notes

Notes = ""

NameOfUser = AskUsersName 'Handily also acts as "Do you want to run now ?"
Notes = InputBox("Any notes ?", "Extra Notes")
MachineModel = MachineDetails("Model")

Report = Report & "<strong>" & "This laptop is for: " & "</strong>" & NameOfUser & "</br>" & vbCrLf
Report = Report & "<strong>" & "Date: " & "</strong>" & Date() & "<strong> Time: </strong>" & Time & "</br>" & vbCrLf
Report = Report & "<strong>" & "Hostname: " & "</strong>" & UserHostName("Hostname") & "</br>" & vbCrLf
Report = Report & "<strong>" & "Script run by: " & "</strong>" & UserHostName("Username") & "</br>" & vbCrLf & "</br>"
if Notes <> "" then '"Notes" If
Report = Report & "<strong>" & "Misc Notes: " & "</strong>" & Notes & "</br>" & VbCrLf
End if 'End of "Notes" if.
Report = Report & "</br> <strong>" & "*** Machine Details: ***" & "</strong>" & "</br>" & vbCrLf
Report = Report & "<strong>" & "Make: " & "</strong>" & MachineDetails("Make") & "</br>" & vbCrLf
Report = Report & "<strong>" & "Model: " & "</strong>" & MachineModel & "</br>" & vbCrLf
Report = Report & "<strong>" & "RAM: " & "</strong>" & MachineRAM & "</br>" & vbCrLf
'Report = Report & "RAM: " & MachineDetails("RAM") & vbCrLf
Report = Report & "<strong>" & "CPU Info." & "</strong>" & vbCrLf & "</br>" & "CPU Model: " & MachineDetails("CPU Model") & "</br>" & VbCrLf
Report = Report & "<strong>" & "CPU Speed: " & "</strong>" & MachineDetails("CPU Speed") & "MHz" & "</br>" & vbCrLf
Report = Report & BiosInfo & "</br>" & vbCrLf
Report = Report & "<strong>" & "Network cards: " & "</br> </strong>" & GetMacAddress & "</br>" & vbCrLf
'Admin Group enumeration:
Report = Report & "<Strong> Admin Group Members: </strong> </br>" & GroupMembers("Administrators") & "</br>" & vbCrLf

'Setup vars for HTML:
Title = NameOfUser & " - " & MachineModel
FName = "C:\" & Title & ".html"
'Check if the local (C:\) file exists, if so ask what to do:
Set objFSO = CreateObject("Scripting.FileSystemObject") ' set + Call library to allow FS Manipulation
if objFSO.FileExists(Fname) then
Set WshShell = WScript.CreateObject("WScript.Shell")
FEReturn = WshShell.Popup("The file " & FName & " already exists" & VbCrLf & "Delete it ?" & VbCrLf & "('No' will quit the script)", , "File Exists!", 4+48)
if FEReturn = vbYes then
objFSO.DeleteFile(FName)
else
WScript.Quit
end if
end if
'Write it to HTML
WriteToHTML Title, Report, FName

'Check the network folder exists (and create it if not)
CheckCreateFolder(CopyPath)

'File:
File = NameOfUser & " - " & MachineModel & ".html"
FDest = CopyPath & File
'Copy the file
CopyFile FName, FDest, "D-Vers-IfExist", True 'Last option = Cleanup.

MsgBox("File Created. Thanks." & VbCrLf & "End of script")
'Testing:
'MsgBox(Report)
