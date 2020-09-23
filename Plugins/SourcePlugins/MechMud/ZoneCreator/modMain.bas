Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Load frmMain
    frmMain.cmdZoneMap(0).Left = -465 * 2
    frmMain.cmdZoneMap(0).Top = -465 * 2
    Call modZones.InitZonesDisplay
    frmMain.Show 0
End Sub
