Attribute VB_Name = "modCommon"
Option Explicit

Public Parent As Object
Public strPath As String

Public Type T_DIRS
    SourcePath As String
    DestPath As String
End Type

Public Dirs() As T_DIRS
Public Dircount As Long
Public mysnIncrement As Boolean
Public BackupInterval As Long

Public TotalFiles As Long

Public Sub TellParent(strString As String)
    
    If Not Parent Is Nothing Then
        CallByName Parent, "log", VbMethod, strString
    End If

End Sub

Public Sub CountFiles()

    Dim lngCounter As Long
    Dim strTemp As String
    
    For lngCounter = 1 To Dircount
        
        strTemp = Dir(Dirs(lngCounter).SourcePath & "\*.*")
        
        TotalFiles = 0
        
        While strTemp <> ""
            TotalFiles = TotalFiles + 1
            strTemp = Dir
        Wend
        
    Next lngCounter

End Sub

