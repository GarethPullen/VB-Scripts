'Written by Gareth Pullen, 26/05/10 to install Altiris agent quickly.
'Written as a bet (Dan Meridew bet I couldn't do it!)
'Modified 04/06/10 to cleanup the logic, tidy, and add the "PacManEnable" Constant to dis/en-able the PacMan prompts.

on error resume next
Dim WshShell, BTN, loopVar, GUID, strWinDir, Success, SWDSolnAgent, SWDSolnAgentUpgrade
Set WshShell = WScript.CreateObject("WScript.Shell" )
Set objFSO = CreateObject("Scripting.FileSystemObject")   ' set + Call library to allow FS Manipulation
Const OverwriteExisting = True
Set Success = 0
Set ObjLink = WshShell.CreateShortcut("C:\AA.lnk")
'** Change these if the file locations change:
SWDSolnAgent = "\\<FQDN>\nscap\Bin\Win32\X86\SWDSolnAgent\AeXSWDSolnAgent.exe"
SWDSolnAgentUpgrade = "\\ukdacans01\nscap\Bin\Win32\X86\SWDSolnAgent\AeXSWDSolnAgentUpgrade.exe"
LogonServer = WSHShell.ExpandEnvironmentStrings("%LOGONSERVER%")
PingLogonServer = Right(LogonServer, Len(LogonServer)-2)
AEXMSI = LogonServer & "\NETLOGON\AeXNSCInstSvc.msi"
HTTPLocation = "http://<ServerURL>/altiris/ns/NSCap/bin/win32/x86/NS%20Client%20Package/AeXNSC.exe"
'** End of the file-variables.
'Set to FALSE to disable PacMan prompts:
Const PacManEnable = True

'**********************************'
'*****Ping Script starts here.*****'
'**********************************'
'Ping Shamelessly stolen from here: http://customerfx.com/pages/crmdeveloper/2008/05/19/ping-a-remote-server-using-vbscript.aspx
Function Ping(Target)
Dim results

    On Error Resume Next

    Set shell = CreateObject("WScript.Shell")
    
    ' Send 1 echo request, waiting 2 seconds for result 
    Set exec = shell.Exec("ping -n 1 -w 2000 " & Target)
    results = LCase(exec.StdOut.ReadAll)
    
    Ping = (InStr(results, "reply from") > 0)
End Function

'**********************************'
'*****Gareth's Functions here.*****'
'**********************************'

Sub PacMan 
dim Window
    set Window = CreateObject("InternetExplorer.Application")
    Window.Visible = 1
    Window.Height = 550
    Window.Width = 600
    Window.Navigate "http://www.google.com/pacman"
End Sub

Function InstallSleep(FileName, Time)
'Takes input of "FileName, Time to sleep (in miliseconds - 1000 = 1 second).
dim InstFuncShell
Set InstFuncShell = WScript.CreateObject("WScript.Shell" )
if objFSO.fileexists(FileName) then
    InstFuncShell.Run FileName, true
    WScript.sleep Time ' sleep $Time to let it work.
    'MsgBox(Success) ' Debug
    InstallSleep = True
Else
    InstallSleep = False
End If 
Set InstFuncShell = Nothing
End Function

Sub Download(Url, Destination)
'Download code taken from: http://blog.netnerds.net/2007/01/vbscript-download-and-save-a-binary-file/
'Hacked about by Gareth Pullen - 04/06/10
' Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    objXMLHTTP.open "GET", Url, false
    objXMLHTTP.send()
    If objXMLHTTP.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary
        objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0    'Set the stream position to the start
        Set objFSO = Createobject("Scripting.FileSystemObject")
        If objFSO.Fileexists(Destination) Then objFSO.DeleteFile Destination
        Set objFSO = Nothing
        objADOStream.SaveToFile Destination
        objADOStream.Close
        Set objADOStream = Nothing
    End if
    Set objXMLHTTP = Nothing
End Sub

Sub QuitCleanup
if objFSO.fileexists("C:\AeXSWDSolnAgentUpgrade.exe") then
    objFSO.DeleteFile("C:\AeXSWDSolnAgentUpgrade.exe")
End If  'Delete AeXSWDSolnAgentUpgrade.exe
if objFSO.fileexists("C:\AeXSWDSolnAgent.exe") then
    objFSO.DeleteFile("C:\AeXSWDSolnAgent.exe")
End If  'Delete AeXSWDSolnAgent.exe
if objFSO.fileexists("C:\AeXNSCInstSvc.msi") then
    objFSO.DeleteFile("C:\AeXNSCInstSvc.msi")
End If  'Delete AeXNSCInstSvc.msi   
Set WshShell = Nothing
Set loopVar = Nothing
Set GUID = Nothing
Set strWinDir = Nothing
Set Success = Nothing
Set ObjFSO = Nothing
Set BTN = Nothing
Set client = Nothing
Set ObjLink = Nothing
WScript.Quit
End Sub

'**********************************'
'****Script proper starts here.****'
'**********************************'

Return = WshShell.Popup("Start Altiris install ?", , "Start ?", 4+32)
if Return = vbNo then
    BTN = WshShell.Popup("Cancelled, exiting",10,"Exiting!",16)
    QuitCleanup
Else
    If Ping(PingLogonServer) then ' Only true if it gets a response.
        objFSO.CopyFile AEXMSI, "C:\", OverWriteExisting
        if objFSO.fileexists("C:\AeXNSCInstSvc.msi") then
            BTN = WshShell.Popup("File copied, running", 5)
            WshShell.Run "C:\AeXNSCInstSvc.msi -qb ns=""<ServerURL>""", true
            Return = MsgBox("Please hit Yes when the install is done" & vbCR & "IMPORTANT! I can't tell when it's done, don't hit Yes until the install has completed!", vbYesNo, "Install done ?")
            If Return = vbYes then
                Wscript.Sleep 5000 ' Sleep 5 seconds to make sure it's done...
                if objFSO.FileExists("C:\AeXNSC.exe") = False then
                Download HTTPLocation, "C:\AeXNSC.exe"
                End If
                BTN = WshShell.Popup("Found file, executing", 5)
                WshShell.run "C:\AeXNSC.exe", 1, true
                If PacManEnable then
                    PM = MsgBox("While that installs, how about a nice game of PacMan ?", vbYesNo, "PacMan ?")
                    If PM = vbYes then
                        PacMan
                    End If  'PacMan end.
                End If '"If PacManEnable = True" End if.
                Wscript.sleep 15000 ' sleep 15 seconds to let it work.
                Success = Success +1
                'MsgBox(Success) ' Debug
                'Write keys, then restart service (hopefully will pickup new server key)
                WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Altiris\Altiris Agent\Servers\" ,"UKDACANS01", "REG_SZ"
                WshShell.RegWrite "HKLM\Software\Altiris\eXpress\NS Client\NSs\ukdacans01\NSWeb" , "http://ukdacans01/Altiris" , "REG_SZ"
                strComputer = "."
                Set objWMIService = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
                Set colServiceList = objWMIService.ExecQuery _
                ("Select * from Win32_Service where Name='AeXNSClient' OR Name='CarbonCopy32' OR Name='Altiris Agent'")
                For each objService in colServiceList 'start of loop.
                    errReturn = objService.StopService()
                    Wscript.Sleep 3000 ' sleep 3 seconds to let it stop.
                    errReturn = objService.StartService()
                    Wscript.Sleep 10000
                    'Sleep 10 seconds to let the services start. - 5 not long enough.
                Next
                    'Loop starts here to check for GUID, loop until GUID is set.
                do
                    'MsgBox(GUID) 'Debug
                    Dim client 
                    Set client=WScript.CreateObject ("Altiris.AeXNSClient" ) 
                    ignoreBlockouts=1 
                    sendIfUnchanged = 1 
                    client.SendBasicInventory sendIfUnchanged, ignoreBlockouts 
                    client.UpdatePolicies ignoreBlockouts
                    client.ClientPolicyMgr.Refresh 
                    'Update policy & send, refresh config, sleep 5 then check GUID
                    Wscript.sleep 5000
                    GUID = WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Altiris\Altiris Agent\MachineGuid")
                Loop while GUID = "" 'End loop when GUID is set.
                'Copy files
                objFSO.CopyFile SWDSolnAgent, "C:\", OverwriteExisting
                objFSO.CopyFile SWDSolnAgentUpgrade, "C:\", OverwriteExisting
                if InstallSleep("C:\AeXSWDSolnAgent.exe", 15000) then
                    Success = Success +1
                    'MsgBox(Success for AeXSWDSolnAgent) ' Debug
                Else
                    MsgBox("Copy AeXSWDSolnAgent failed!")
                    QuitCleanup
                End If '"if AeXSWDSolnAgent.exe" end.
                if InstallSleep("C:\AeXSWDSolnAgentUpgrade.exe", 15000) then
                    Success = Success +1
                    'MsgBox(Success for AeXSWDSolnAgentUpgrade) ' Debug
                Else
                    MsgBox("Copy AeXSWDSolnAgentUpgrade Failed!")
                    QuitCleanup
                End If  '"if AeXSWDSolnAgentUpgrade.exe" end.
                Wscript.sleep 5000
                If Success = 3 then ' Success will only be 3 if it all worked...
                    BTN = WshShell.Popup("Running software portal", 5)
                    AA = "C:\Program Files\Altiris\Altiris Agent\SWRAgentUtils.exe" 
                    ObjLink.Description = "Altiris SP"
                    ObjLink.TargetPath = AA
                    ObjLink.Arguments = "/ShowSoftwarePortal"
                    ObjLink.Save
                    WshShell.run("C:\AA.lnk")
                    if objFSO.fileexists("C:\AA.lnk") then    
                        ObjFSO.DeleteFile "C:\AA.lnk"
                    End If '"AA.lnk" delete end.
                End If  'Run software portal end.
            Else '"Is install done" else.
                MsgBox("You hit No. Script exiting.")
                QuitCleanup
            End If ' Install Done End If
        Else 'Download "AeXNSCInstSvc.msi
            MsgBox("Error! Can't download AeXNSCInstSvc.msi" & vbCR & "from - " & PingLogonServer)
            QuitCleanup
        End If 'Download "AeXNSCInstSvc.msi
    Else 'Ping Else
        MsgBox("Can't contact " & LogonServer & vbCR & "Please check network connection!")
        QuitCleanup
    End If  'Ping End if.
    QuitCleanup
End If ' First "Should install start ?" Question.
