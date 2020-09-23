VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackupAgent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Agent Configuration Tool"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackupAgent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2010
      TabIndex        =   13
      Text            =   "45"
      Top             =   5880
      Width           =   765
   End
   Begin VB.CheckBox chkIncremental 
      Alignment       =   1  'Right Justify
      Caption         =   "Create incremental directors (Uses date for dir.)"
      Height          =   450
      Left            =   105
      TabIndex        =   11
      Top             =   6195
      Width           =   2700
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete selected directory"
      Height          =   315
      Left            =   3690
      TabIndex        =   10
      Top             =   5505
      Width           =   2445
   End
   Begin VB.CommandButton cmdFinishSetup 
      Caption         =   "Finish Setup"
      Height          =   315
      Left            =   2835
      TabIndex        =   9
      Top             =   6990
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   6225
      TabIndex        =   8
      Top             =   5505
      Width           =   1140
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   300
      Index           =   1
      Left            =   7020
      TabIndex        =   7
      Top             =   5130
      Width           =   330
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   300
      Index           =   0
      Left            =   7020
      TabIndex        =   6
      Top             =   4800
      Width           =   330
   End
   Begin VB.TextBox txtDest 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2010
      TabIndex        =   5
      Top             =   5130
      Width           =   4950
   End
   Begin VB.TextBox txtSource 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2010
      TabIndex        =   4
      Top             =   4800
      Width           =   4950
   End
   Begin MSComctlLib.ListView lvwDirs 
      Height          =   3735
      Left            =   75
      TabIndex        =   1
      Top             =   960
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label5 
      Caption         =   "minutes"
      Height          =   270
      Left            =   2835
      TabIndex        =   14
      Top             =   5895
      Width           =   675
   End
   Begin VB.Label Label4 
      Caption         =   "Time interval of backups:"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   5895
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "Destination Directory:"
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "Source Directory:"
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   $"frmBackupAgent.frx":000C
      Height          =   1020
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7260
   End
End
Attribute VB_Name = "frmBackupAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If txtDest.Text = "" Then Exit Sub
    If txtSource.Text = "" Then Exit Sub
    
    lvwDirs.ListItems.Add , "key" & Timer, txtSource.Text
    lvwDirs.ListItems(lvwDirs.ListItems.Count).ListSubItems.Add , "key" & Timer & "_1", txtDest.Text
    txtDest.Text = ""
    txtSource.Text = ""
    
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    frmSelDir.Show vbModal
    If Index = 0 Then
        txtSource.Text = strPath
    Else
        txtDest.Text = strPath
    End If
End Sub

Private Sub cmdDel_Click()
    If Not lvwDirs.SelectedItem Is Nothing Then
        lvwDirs.ListItems.Remove lvwDirs.SelectedItem.Index
    End If
End Sub

Private Sub cmdFinishSetup_Click()

    If Not IsNumeric(txtInterval.Text) Then
        MsgBox "The value you entered in the interval textbox is not a numeric value.", vbOKOnly, "Not a numeric value."
        Exit Sub
    End If

    Dim vbMsg As VbMsgBoxResult
    
    vbMsg = MsgBox("Are you sure you wish to finish the setup?" & vbCrLf & "You can only return to this window if you delete BackupAgent.ini", vbYesNo, "Finish setup?")

    If vbMsg = vbYes Then
    
        Dim IntCount As Integer
        
        Open App.Path & "\BackupAgent.ini" For Output As #1
            Print #1, "[BackupAgent]" & vbCrLf
            Print #1, "DirCount=" & CStr(lvwDirs.ListItems.Count) & vbCrLf
            Print #1, "Interval=" & CStr(txtInterval.Text) & vbCrLf
            If chkIncremental.Value = 1 Then
                Print #1, "Increment=1" & vbCrLf
            Else
                Print #1, "Increment=0" & vbCrLf
            End If
            
            Print #1, vbCrLf
            
            For IntCount = 1 To lvwDirs.ListItems.Count
                Print #1, "[Path" & CStr(IntCount) & "]" & vbCrLf
                Print #1, "Source=" & lvwDirs.ListItems(IntCount).Text & vbCrLf
                Print #1, "Dest=" & lvwDirs.ListItems(IntCount).SubItems(1) & vbCrLf
                Print #1, vbCrLf
            Next IntCount
            
        Close #1

        TellParent "* Backup Agent has finished setting up."
        TellParent ""
        Unload frmBackupAgent
    End If
    
End Sub

Private Sub Form_Load()

    lvwDirs.ColumnHeaders.Add , "Source", "Source Dir", 3615
    lvwDirs.ColumnHeaders.Add , "Dest", "Destination Dir", 3615

End Sub

Private Sub lvwDirs_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lvwDirs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not Item Is Nothing Then
        txtSource.Text = Item.Text
        txtDest.Text = Item.SubItems(1)
    End If
End Sub
