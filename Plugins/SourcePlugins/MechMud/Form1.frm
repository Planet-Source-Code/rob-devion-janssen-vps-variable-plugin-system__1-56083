VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6915
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   7200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtLog_Change()
    txtLog.SelStart = Len(txtLog.Text)
End Sub

Public Sub Log(strTExt As String)
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = strTExt & vbCrLf
End Sub

