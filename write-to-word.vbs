
Function WriteToWord(rArgs) 'Array.
'Word Function.
'Array inputs:
'1 - Text to write, 2 - File Name   - Mandatory!
'3 - Font Name, 4 - Font Size   'Optional.
'Returns file name

Set objWord = CreateObject("Word.Application")  'Create the Word object

objWord.Visible = ShowWord  'Show the Word window - False means hide.
Set objDoc = objWord.Documents.Add()    'Create new document
Set objSelection = objWord.Selection    'Needed to write.
'Set variable for Array:
dim InText, FileName
'These will be changed if they're passed to the function, otherwise these are 'defaults':
objSelection.Font.Name = "Arial"
objSelection.Font.Size = "12"

Select Case UBound(rArgs)   '***This returns /HOW MANY/ items are in the array, so 0-4.
'*** Due to the UBound returning /HOW MANY/ items are in the array, you have to do_
'*** The previous items in each new one, hence the below Case.

Case 0  'Array starts at 0, so this is 1 argument.
    'If only one argument passed then it's wrong - needs a file name *AND* some text to write.
    WriteToWord = "Error! Too few parameters!" 'Using function name quits the function, thus 'error checking'
Case 1
    'File Name + 'Take the text'
    InText = rArgs(0)
    FileName = rArgs(1)
Case 2
    'Font Name (if specified) (+ Above...)
    InText = rArgs(0)
    FileName = rArgs(1)
    objSelection.Font.Name = rArgs(2)
Case 3
    'Font Size (if specified) (+ Above...)
    InText = rArgs(0)
    FileName = rArgs(1)
    objSelection.Font.Name = rArgs(2)
    objSelection.Font.Size = rArgs(3)
End Select  'End Case

'Now write the text:
objSelection.TypeText InText

'Save the Document:
objDoc.SaveAs(FileName)

' Quit word!
objWord.Quit
'Return the file name
WriteToWord = FileName

End Function

dim ArInput

ReDim ArInput(3)
ArInput(0) = "Text to write"
ArInput(1) = "C:\Test-Doc.doc"
ArInput(2) = "Arial"
ArInput(3) = "14"

'MsgBox("Input: " & vbCrLf & ArInput)
WriteToWord(ArInput)
MsgBox("Written file: " & ArInput(1))
