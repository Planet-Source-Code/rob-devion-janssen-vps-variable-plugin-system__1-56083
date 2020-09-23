VERSION 5.00
Begin VB.Form frmCore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VPS-CORE-0"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2085
   Icon            =   "frmCore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   2085
   Visible         =   0   'False
End
Attribute VB_Name = "frmCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Form objects can be referenced by objects.

Public Sub Log(strString As String)

    Open App.Path & "\VPS.Log " For Append As #1
    Print #1, strString
    Close #1
    ConWrite strString & vbCrLf
    DoEvents
    
End Sub

Public Sub Logfile(strString As String)

    Open App.Path & "\VPS.Log " For Append As #1
    Print #1, strString
    Close #1
    DoEvents
    
End Sub

Public Sub Logscreen(strString As String)

    ConWrite strString & vbCrLf
    DoEvents
    
End Sub

Public Sub goIntoWaitState()
    DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 1 Then
     If KeyCode = vbKeyC Then
        modMain.App_Close
        Unload frmCore
        End
     End If
    End If
    
    If KeyCode = vbKeyEscape Then
        modMain.App_Close
    End If
End Sub

