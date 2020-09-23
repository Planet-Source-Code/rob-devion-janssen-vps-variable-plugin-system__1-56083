VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MechMud Zone Creator"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   885
      Left            =   11550
      TabIndex        =   30
      Top             =   9675
      Width           =   4125
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   ".. No zone selected .."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   33
         Top             =   495
         Width           =   2910
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zone ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   345
         TabIndex        =   32
         Top             =   495
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse-over zone Info"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   31
         Top             =   195
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   8595
      Left            =   11550
      TabIndex        =   6
      Top             =   1065
      Width           =   4140
      Begin VB.CommandButton cmdRoomOption 
         Caption         =   "Import Room"
         Height          =   270
         Index           =   2
         Left            =   270
         TabIndex        =   36
         Top             =   8145
         Width           =   1110
      End
      Begin VB.CommandButton cmdRoomOption 
         Caption         =   "Export Room"
         Height          =   270
         Index           =   1
         Left            =   1500
         TabIndex        =   35
         Top             =   8145
         Width           =   1155
      End
      Begin VB.CommandButton cmdRoomOption 
         Caption         =   "Save Room"
         Height          =   270
         Index           =   0
         Left            =   2775
         TabIndex        =   34
         Top             =   8145
         Width           =   1140
      End
      Begin VB.TextBox txtExitID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2610
         TabIndex        =   25
         Text            =   "1"
         Top             =   6885
         Width           =   915
      End
      Begin VB.TextBox txtExitID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2610
         TabIndex        =   24
         Text            =   "1"
         Top             =   6390
         Width           =   915
      End
      Begin VB.TextBox txtExitID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2610
         TabIndex        =   23
         Text            =   "1"
         Top             =   5910
         Width           =   915
      End
      Begin VB.TextBox txtExitID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2610
         TabIndex        =   22
         Text            =   "1"
         Top             =   5460
         Width           =   915
      End
      Begin VB.CheckBox chkExits 
         BackColor       =   &H00808080&
         Caption         =   "West"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   375
         TabIndex        =   21
         Top             =   6660
         Width           =   3525
      End
      Begin VB.CheckBox chkExits 
         BackColor       =   &H00808080&
         Caption         =   "South"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   375
         TabIndex        =   20
         Top             =   6180
         Width           =   3525
      End
      Begin VB.CheckBox chkExits 
         BackColor       =   &H00808080&
         Caption         =   "East"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   390
         TabIndex        =   19
         Top             =   5700
         Width           =   3525
      End
      Begin VB.CheckBox chkExits 
         BackColor       =   &H00808080&
         Caption         =   "North"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   390
         TabIndex        =   18
         Top             =   5220
         Width           =   3525
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00808080&
         Caption         =   "This zone is a portal to other world map"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   390
         TabIndex        =   16
         Top             =   4605
         Width           =   3615
      End
      Begin VB.TextBox txtHostileLevel 
         Alignment       =   1  'Right Justify
         Height          =   240
         Left            =   2610
         TabIndex        =   15
         Text            =   "1"
         Top             =   4350
         Width           =   915
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00808080&
         Caption         =   "This zone is a hostile zone."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   390
         TabIndex        =   13
         Top             =   4065
         Width           =   3195
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00808080&
         Caption         =   "This zone is a recall point."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   390
         TabIndex        =   12
         Top             =   3780
         Width           =   3195
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H00808080&
         Caption         =   "This zone is a starting point."
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   390
         TabIndex        =   11
         Top             =   3495
         Width           =   3195
      End
      Begin VB.TextBox Text2 
         Height          =   2445
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   735
         Width           =   3675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Room ID for West exit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   645
         TabIndex        =   29
         Top             =   6930
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Room ID for South exit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   645
         TabIndex        =   28
         Top             =   6420
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Room ID for East exit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   660
         TabIndex        =   27
         Top             =   5940
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Room ID for North exit:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   660
         TabIndex        =   26
         Top             =   5475
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Exits (if Worldzone exit enter ID name of world) :"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   8
         Left            =   255
         TabIndex        =   17
         Top             =   4965
         Width           =   3765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hostile level (1 to 25):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   675
         TabIndex        =   14
         Top             =   4365
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zone options:"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   6
         Left            =   255
         TabIndex        =   10
         Top             =   3270
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zone description:"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   5
         Left            =   255
         TabIndex        =   8
         Top             =   495
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoneblock  information"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   3900
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   11535
      TabIndex        =   2
      Top             =   -45
      Width           =   4155
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1695
         TabIndex        =   5
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "World name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   225
         TabIndex        =   4
         Top             =   555
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zone information"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   195
         Width           =   1305
      End
   End
   Begin VB.PictureBox picBox 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10515
      Left            =   45
      ScaleHeight     =   10455
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   60
      Width           =   11460
      Begin VB.CommandButton cmdZoneMap 
         BackColor       =   &H00C0FFFF&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenZone 
         Caption         =   "Open ZoneMap"
      End
      Begin VB.Menu mnuSavezone 
         Caption         =   "Save Zonemap"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRoomOption_Click(Index As Integer)
    cmdZoneMap(CurrentRoom).BackColor = RGB(0, 255, 0)
End Sub

Private Sub cmdZoneMap_Click(Index As Integer)
    If CurrentRoom <> 0 Then
        If cmdZoneMap(CurrentRoom).BackColor <> RGB(0, 255, 0) Then
            cmdZoneMap(CurrentRoom).BackColor = cmdZoneMap(0).BackColor
        End If
    End If
    
    Label1(4).Caption = "Zoneblock  information for zoneID: " & CStr(Index)
    CurrentRoom = Index
    cmdZoneMap(Index).BackColor = RGB(255, 0, 0)
End Sub

Private Sub cmdZoneMap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblInfo.Caption = CStr(Index)
End Sub
