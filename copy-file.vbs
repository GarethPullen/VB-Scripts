
Function CopyFile(FileLocation, FileDestination, OverWrite, Cleanup)
'File Copy Function.
'Takes input file name(with path), destination, overwrite (True, False, or D-Version), "Cleanup" - True or False, delete the original after copying.
'Example call:
'CopyFile "C:\Test.doc", "C:\Testing.doc", "D-Vers-IfExist", False
Set objFSO = CreateObject("Scripting.FileSystemObject")   ' set + Call library to allow FS Manipulation

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
    FileDestination = Replace(FileDestination, ".doc", " - " & (Replace(Date, "/", "-") & ".doc"))
    if objFSO.FileExists(FileLocation) then 'Sanity check - does the file (still) exist ?
        ObjFso.CopyFile FileLocation, FileDestination, False
    else
        CopyFile = "Error! File doesn't exist!"
    end if

Case "D-Vers-IfExist"   'Date version if file exists...
        if objFSO.FileExists(FileDestination) then
            FileDestination = Replace(FileDestination, ".doc", " - " & (Replace(Date, "/", "-") & ".doc"))
                if objFSO.FileExists(FileDestination) then 'Does a date-named version exist ?
                    'If so, add time to the file name.
                    FileDestination = Replace(FileDestination, ".doc", " - " & (Hour(Time) & "." & Minute(Time) & ".doc"))
                    if objFSO.FileExists(FileDestination) then 'Does a hour/minute-named version exist ?
                        'If so, add seconds ?
                        FileDestination = Replace(FileDestination, ".doc", "." & (Second(Time) & ".doc"))
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

CopyFile "C:\Test.doc", "C:\Testing.doc", "D-Vers-IfExist", False
