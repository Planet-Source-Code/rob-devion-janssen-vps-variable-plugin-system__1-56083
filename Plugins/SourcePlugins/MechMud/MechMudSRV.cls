VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MechMudSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MechMud is originally meant as a plugin for my VPS system.
' Though with the modPluginCommon mod you can also set it as stand-alone.
' No logging though without VPS.
'

' Plugin Common
Public Function fStrGetName() As String
    fStrGetName = "MechMud Server VPS Plugin"
End Function

Public Function fstrGetFunction() As String
    fstrGetFunction = "InitMechMud"
End Function

Public Function fClose() As String
    fClose = "Ok"
End Function

Public Function RegisterParent(frmObject As Object) As String
    Set Parent = frmObject
    RegisterParent = "Ok"
End Function

' -----'

Public Sub InitMechMud()

    Load frmMain
    frmMain.Visible = False

    TellParent modServerCommon.SERVER_NAME & " " & modServerCommon.SERVER_VERSION
    TellParent "Server running on port " & modServerCommon.SERVER_PORT
    
    ' Dev
    Load Form1
    Form1.Visible = True
    
    InitServerCommons
    initServerNetwork
    
    
End Sub
