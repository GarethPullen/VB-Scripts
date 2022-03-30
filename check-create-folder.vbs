
Function CheckCreateFolder(Path)
'Function to check if a folder exists, and if not create it.
'End result = Folder will exist.
'Takes "Path" as the folder path to check.
Set objFSO = CreateObject("Scripting.FileSystemObject")   ' set + Call library to allow FS Manipulation
if objFSO.Folderexists(Path) then
    CheckCreateFolder = "True" 'Return true that it exists.
else
    'objFSO.CreateFolder(Path)
    Set objShell = CreateObject("Wscript.Shell") 
    objShell.Run "cmd /c mkdir " & """" & Path & """"
    CheckCreateFolder = "True" 'Folder now exists.
end if

End Function

CONST NetPath = "<FQDN>\IT\Laptops Checklist\Auto\"    'This is the base folder to copy to, + Year (must have trailing \).
CopyPath = NetPath & Year(Date) & "\"   'Append the year to get the right folder.
CheckCreateFolder(CopyPath)
