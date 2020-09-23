Attribute VB_Name = "modServerCommon"
Option Explicit

Public Const SERVER_NAME As String = "MechMUD by Devion"
Public Const SERVER_VERSION As String = "v0.1alpha"
Public Const SERVER_PORT As Integer = 2002

Public Const SERVER_MAX_USERS As Integer = 1400 ' max users

Public SERVER_WELCOME As String

Public SERVER_USERNAME_TEXT As String
Public SERVER_PASSWORD_TEXT As String
Public SERVER_AUTH_TEXT As String

Public Sub InitServerCommons()
    TellParent "Initialising Server commons..."
    SERVER_WELCOME = modFiles.fstrGetTextFromFile(App.Path & "\MechMudData\ServerWelcome.txt")
    SERVER_USERNAME_TEXT = "Please enter your username: "
    SERVER_PASSWORD_TEXT = ".. checking..." & vbCrLf & ".. Requiring password..." & vbCrLf & vbCrLf & "Enter password: "
    SERVER_AUTH_TEXT = vbCrLf & "..checking..." & vbCrLf
    
End Sub
