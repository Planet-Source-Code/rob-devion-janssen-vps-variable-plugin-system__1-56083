Attribute VB_Name = "modFiles"
Option Explicit

Public Function fstrGetTextFromFile(strFilename As String) As String
On Error GoTo errhandler

    Dim strTemp As String
    
    If fysnExist(strFilename) Then
    
        strTemp = Space$(FileLen(strFilename))
        Close #1
        Open strFilename For Binary As #1
        Get #1, , strTemp
        Close #1
        fstrGetTextFromFile = Trim$(strTemp)
    Else
        fstrGetTextFromFile = "Sorry, Could not fetch a server file named " & strFilename
    End If
    
    Exit Function
    
errhandler:
    fstrGetTextFromFile = "Sorry, Could not fetch a server file named " & strFilename
    
End Function

Public Function fysnExist(strFilename As String) As Boolean
On Error GoTo errhandler

    Dim lngTemp As Long
    
    lngTemp = FileLen(strFilename)
    fysnExist = True
    Exit Function
    
errhandler:
    fysnExist = False
End Function
