VERSION 5.00
Begin VB.Form frmSelDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select directory..."
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      TabIndex        =   2
      Top             =   4440
      Width           =   1185
   End
   Begin VB.DirListBox objDir 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   90
      TabIndex        =   1
      Top             =   480
      Width           =   4020
   End
   Begin VB.DriveListBox objDrive 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4050
   End
End
Attribute VB_Name = "frmSelDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    strPath = objDir.Path
    Unload Me
End Sub

Private Sub objDrive_Change()
    objDir.Path = objDrive.Drive
End Sub

