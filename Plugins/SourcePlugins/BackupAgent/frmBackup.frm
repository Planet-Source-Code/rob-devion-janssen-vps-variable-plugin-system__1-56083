VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backing up data...."
   ClientHeight    =   435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   435
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrBackup 
      Enabled         =   0   'False
      Left            =   225
      Top             =   15
   End
   Begin MSComctlLib.ProgressBar objProgress 
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngTimer As Long
Public lngTmpTimer As Long

Public intTemp As Integer

Private Sub tmrBackup_Timer()
    
    intTemp = intTemp + 1
       
    If lngTmpTimer >= 1 Then
        lngTmpTimer = lngTmpTimer - 1 ' Minus one.
    End If
    
    If modCommon.TotalFiles = 0 Then Exit Sub
    
    If lngTmpTimer = 0 Then
        TellParent "Backup Agent: Initiating Backup of directories..."
        lngTmpTimer = lngTimer ' Reset timer
        frmBackup.Visible = True
        frmBackup.Left = 250
        frmBackup.Top = 250
        frmBackup.objProgress.Min = 1
        frmBackup.objProgress.Max = CSng(modCommon.TotalFiles)
        frmBackup.objProgress.Value = 1
        BackupFiles
    End If

End Sub

Public Sub BackupFiles()

   On Error Resume Next
    
    Dim lngCounter As Long
    Dim strTemp As String
    Dim lngProg As Long
    Dim strDateDir As String
    
    strDateDir = Format(Now, "DMMYYYY")
    
    For lngCounter = 1 To Dircount
        strTemp = Dir(Dirs(lngCounter).SourcePath & "\*.*")
        While strTemp <> ""
            lngProg = lngProg + 1
            frmBackup.objProgress.Value = CSng(lngProg)
            
            If modCommon.mysnIncrement Then
                MkDir Dirs(lngCounter).DestPath & "\" & strDateDir
                FileCopy Dirs(lngCounter).SourcePath & "\" & strTemp, Dirs(lngCounter).DestPath & "\" & strDateDir & "\" & strTemp
            Else
                FileCopy Dirs(lngCounter).SourcePath & "\" & strTemp, Dirs(lngCounter).DestPath & "\" & strTemp
            End If
                        
            DoEvents
            
            strTemp = Dir
        Wend
    Next lngCounter
    
    frmBackup.Visible = False
End Sub

