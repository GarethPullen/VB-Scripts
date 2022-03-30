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
        GroupMems = GroupMems & objUser.Name & vbCrLf
    Next
end if
Next
GroupMembers = GroupMems
End Function

MsgBox("Admin: " & vbCrLf &  GroupMembers("Administrators"))
