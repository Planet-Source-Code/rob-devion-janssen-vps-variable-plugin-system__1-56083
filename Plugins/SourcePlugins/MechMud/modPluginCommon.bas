Attribute VB_Name = "modPluginCommon"
Option Explicit

' This is the msg-feedback for the VPS.
' Dont touch if you are using the VPS System...

Public Parent As Object ' For VPS!

Public Sub TellParent(strString As String)
    
    ' Standalone version
    
    If Not Parent Is Nothing Then
        CallByName Parent, "log", VbMethod, "MechMud: " & strString
    End If

End Sub

