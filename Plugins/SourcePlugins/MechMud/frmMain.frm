VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MECH Mud Server"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   1905
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wskPlayers 
      Index           =   0
      Left            =   540
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   75
      Top             =   45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2002
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub wskPlayers_DataArrival(index As Integer, ByVal bytesTotal As Long)
    
    Dim lngTemp As Long
    Dim strTemp As String
    
    wskPlayers(index).GetData strTemp
    lngTemp = flngFindPlayerId(wskPlayers(index).SocketHandle)
    
    If InStr(1, strTemp, Chr(13)) Then
        If InStr(1, strTemp, Chr(10)) Then
            strTemp = udtPlayers(lngTemp).strNetworkData
            udtPlayers(lngTemp).strNetworkData = ""
        End If
    Else
        udtPlayers(lngTemp).strNetworkData = udtPlayers(lngTemp).strNetworkData & strTemp
        Exit Sub
    End If
    
    Select Case udtPlayers(lngTemp).NetworkState
        Case Login
            Select Case udtPlayers(lngTemp).StateStep
                Case 1
                    udtPlayers(lngTemp).strUsernameTemp = strTemp
                    udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
                    wskPlayers(index).SendData vbCrLf & SERVER_PASSWORD_TEXT
                Case 2
                    udtPlayers(lngTemp).strPasswordTemp = strTemp
                    udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
                    wskPlayers(index).SendData SERVER_AUTH_TEXT
                    If Not CheckPlayer(lngTemp) Then
                        udtPlayers(lngTemp).NetworkState = NewChar
                        udtPlayers(lngTemp).StateStep = 1
                        wskPlayers(index).SendData vbCrLf & "Welcome new recruit..." & vbCrLf & "Please take a minute to fill in these forms..." & vbCrLf & vbCrLf & "Your nickname: "
                        Exit Sub
                    End If
                    
                    If CheckAuth(lngTemp) Then
                        wskPlayers(index).SendData vbCrLf & "Clearance granted!" & vbCrLf
                        wskPlayers(index).SendData vbCrLf & "Teleporting you to your last known location..." & vbCrLf
                        udtPlayers(lngTemp).StateStep = 1
                        udtPlayers(lngTemp).NetworkState = Ingame
                    Else
                        wskPlayers(index).SendData vbCrLf & "Username or password invalid!" & vbCrLf
                        udtPlayers(lngTemp).StateStep = 1
                        wskPlayers(index).SendData vbCrLf & SERVER_USERNAME_TEXT
                    End If
                    
            End Select
        Case Ingame
            modCommandParse.Parsecommand strTemp, lngTemp
        Case NewChar
            RegisterNewChar lngTemp, strTemp, index
    End Select
    strTemp = ""
End Sub

Private Sub wskPlayers_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim lngTemp As Long
    
    lngTemp = flngFindPlayerId(wskPlayers(index).SocketHandle)
    
    If lngTemp <= 1 Then
        TellParent "Winsock Player sockets - Error on unknown player."
    Else
        TellParent "Winsock Player sockets error affecting (" & udtPlayers(lngTemp).strUsername & ") with ID: " & CStr(wskPlayers(index).SocketHandle) & " had an error."
        wskPlayers(index).Close
        udtPlayers(lngTemp).sckConnected = False
    End If
    
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
    
    If modPlayers.CheckPlayerSlots Then
        intPlayerCount = intPlayerCount + 1
        Load wskPlayers(intPlayerCount)
        wskPlayers(intPlayerCount).Accept requestID
        DoEvents
        modPlayers.AddPlayer wskPlayers(intPlayerCount).SocketHandle
        TellParent "Incoming connection accepted ID:(" & CStr(requestID) & ") at slot " & CStr(intPlayerCount)
    Else
        TellParent "Player limit of " & SERVER_MAX_USERS & " has been reached."
        TellParent "Incoming connections are disconnected immediately."
        DoEvents
        wskServer.Close
        wskServer.Listen
    End If
    
End Sub

Private Sub wskServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    TellParent "Winsock Listen Error: " & CStr(Number) & " - " & Description
    wskServer.Close
    wskServer.Listen
    TellParent "Winsock Listen Server restarted."
End Sub
