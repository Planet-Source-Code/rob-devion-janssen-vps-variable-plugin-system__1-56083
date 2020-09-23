Attribute VB_Name = "modPlayers"
Option Explicit

' Player data

Public Type udtPlayer_Template
    ' Network
    sckConnID As Long
    sckConnected As Boolean
    NetworkState As Integer
    StateStep As Integer
    strNetworkData As String
    email As String
    
    ' User common
    strUsername As String
    strPassword As String ' = MD5 Hash.
    
    strUsernameTemp As String
    strPasswordTemp As String
    
    strNickname As String
    ysnAdmin As Boolean
    ysnBanned As Boolean
    ClassID As Integer
    ClanID As Integer
    
    strRank As String
    
    ' Player common
    
    lngCredits As Long
    intHP As Integer
    intMaxHP As Integer
    ysnAlive As Boolean
    
End Type

' -------------

Public udtPlayers() As udtPlayer_Template
Public intPlayerCount As Integer

Public Sub AddPlayer(sckConnID As Long)

    If CheckPlayerSlots Then
        ReDim Preserve udtPlayers(1 To intPlayerCount)
        
        With udtPlayers(intPlayerCount)
            .NetworkState = Login
            .sckConnected = True
            .sckConnID = sckConnID
            .StateStep = 1 'Get username
        End With
    End If

    SendMessageToClient sckConnID, SERVER_WELCOME
    SendMessageToClient sckConnID, SERVER_USERNAME_TEXT
    
End Sub

Public Function CheckPlayerSlots() As Boolean

    If intPlayerCount >= SERVER_MAX_USERS Then
        CleanUpArray
        If intPlayerCount >= SERVER_MAX_USERS Then
            CheckPlayerSlots = False
            Exit Function
        End If
    Else
        CheckPlayerSlots = True
        Exit Function

    End If
    
End Function

Public Sub CleanUpArray()
    ' Cleans up the array if there are broken players in it
End Sub

Public Sub LoadPlayerData(strString As String, lngTemp As Long)
    If fysnExist(strString) Then
        Close #1
        Open strString For Random As #1
            Get #1, , udtPlayers(lngTemp)
        Close #1
    Else
        TellParent "ERROR: Could not load playerdata for " & udtPlayers(lngTemp).strUsername
    End If
End Sub

Public Sub SavePlayerData(strString As String, lngTemp As Long)
    
    Dim strTemp As String
       
    Close #1
    If fysnExist(strString) Then
        Kill strString
    End If
    
    Open strString For Random As #1
        Put #1, , udtPlayers(lngTemp)
    Close #1

End Sub

Function FilterString(text As String, ValidChars As String) As String
    Dim i As Long, result As String
    For i = 1 To Len(text)
        If InStr(ValidChars, Mid$(text, i, 1)) Then
            result = result & Mid$(text, i, 1)
        End If
    Next
    FilterString = result
End Function

Function CheckPlayer(lngTemp As Long) As Boolean

    Dim strTemp As String
    
    strTemp = udtPlayers(lngTemp).strUsername
    
    strTemp = FilterString(strTemp, "abdefghijklmnopqrstuvwxyzABDEFGHIJKLMNOPQRSTUVWXYZ1234567890")
    
    If fysnExist(App.Path & "\mechmuddata\" & strTemp & ".data") Then
        modPlayers.LoadPlayerData App.Path & "\mechmuddata\" & strTemp & ".data", lngTemp
        CheckPlayer = True
    Else
        CheckPlayer = False
    End If
    
End Function

Public Sub RegisterNewChar(lngTemp, strText As String, index As Integer)
    
    Select Case udtPlayers(lngTemp).StateStep
        Case 1 ' Nickname
            udtPlayers(lngTemp).strNickname = strText
            frmMain.wskPlayers(index).SendData vbCrLf & "Email address: "
            udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
        Case 2 ' email
            udtPlayers(lngTemp).email = strText
            frmMain.wskPlayers(index).SendData vbCrLf & modFiles.fstrGetTextFromFile(App.Path & "\mechmuddata\Playerclasses.txt")
            udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
        Case 3 ' playerClass
            If IsNumeric(strText) Then
                udtPlayers(lngTemp).ClassID = CInt(strText)
                frmMain.wskPlayers(index).SendData vbCrLf & vbCrLf & "Almost there recruit!" & vbCrLf & vbCrLf & "Do you wish to ally with a clan now (Y/N)? "
                udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
            Else
                frmMain.wskPlayers(index).SendData vbCrLf & vbCrLf & modFiles.fstrGetTextFromFile(App.Path & "\mechmuddata\Playerclasses.txt")
            End If
        Case 4 ' Clan
            If LCase(Left(strText, 1)) = "y" Then
                frmMain.wskPlayers(index).SendData vbCrLf & "Okay.. which?" & vbCrLf & modFiles.fstrGetTextFromFile(App.Path & "\mechmuddata\clans.txt")
                udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
            Else
                frmMain.wskPlayers(index).SendData vbCrLf & "Okay.. no problem.." & vbCrLf
                udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 2
            End If
            
        Case 5 ' clan
            If InStr(1, LCase$(strText), "x") Then
                frmMain.wskPlayers(index).SendData vbCrLf & "Okay.. no problem.." & vbCrLf
            Else
                udtPlayers(lngTemp).ClanID = CInt(strText)
            End If
            
            udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
    
        Case 6 ' ..
            udtPlayers(lngTemp).ClassID = CInt(strText)
            frmMain.wskPlayers(index).SendData vbCrLf & vbCrLf & "Almost there recruit!" & vbCrLf & vbCrLf & "Do you wish to ally with a clan now (Y/N)? "
            udtPlayers(lngTemp).StateStep = udtPlayers(lngTemp).StateStep + 1
    
    End Select

End Sub
