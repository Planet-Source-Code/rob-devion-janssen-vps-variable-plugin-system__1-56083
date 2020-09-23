Attribute VB_Name = "modNetworking"
Option Explicit

Public Enum enumNetWorkState_template
    Login
    Ingame
    NewChar
End Enum

Public ConnectedClients As Integer

Public Sub initServerNetwork()
    TellParent "Initialising MechMUD Network components..."
    frmMain.wskServer.LocalPort = SERVER_PORT
    frmMain.wskServer.Listen
End Sub

Public Function fintFindID(ConnID As Long) As Integer

    On Error Resume Next
    
    Dim intCount As Integer
    
    For intCount = 1 To frmMain.wskPlayers.Count - 1
        If frmMain.wskPlayers(intCount).SocketHandle = ConnID Then
            fintFindID = intCount
            Exit Function
        End If
    Next
    
    fintFindID = 0
    
End Function

Public Function flngFindPlayerId(SockHandle As Long) As Long

    Dim intCount As Integer
    
    For intCount = 1 To intPlayerCount
        If udtPlayers(intCount).sckConnID = SockHandle Then
            flngFindPlayerId = intCount
            Exit Function
        End If
    Next
    
    flngFindPlayerId = 0
    
End Function


Public Sub SendMessageToClient(ConnID As Long, strMessage As String)

    Dim IntTemp As Integer
    
    IntTemp = fintFindID(ConnID)
    
    If IntTemp = 0 Then Exit Sub
    
    If frmMain.wskPlayers(IntTemp).State = sckConnected Then
        frmMain.wskPlayers(IntTemp).SendData strMessage & Chr(0)
    End If

End Sub

Public Function CheckAuth(lngTemp As Long) As Boolean

    With udtPlayers(lngTemp)
    
    If .strUsername <> .strUsernameTemp Then
        CheckAuth = False
        Exit Function
    End If
    
    If .strPassword <> modMD5.CalculateMD5(.strPasswordTemp) Then
        CheckAuth = False
        Exit Function
    End If
    
    CheckAuth = True
    
    End With
    
End Function
