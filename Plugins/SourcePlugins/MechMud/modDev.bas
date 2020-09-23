Attribute VB_Name = "modTempDevelopment"
Option Explicit

' For Testing

Dim Server As New MechMudSRV

Public Sub Main()
    
    Server.RegisterParent Form1
    Server.InitMechMud

End Sub
