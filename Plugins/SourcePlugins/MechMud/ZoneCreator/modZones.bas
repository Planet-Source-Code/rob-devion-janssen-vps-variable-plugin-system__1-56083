Attribute VB_Name = "modZones"
Option Explicit

Public Type T_ZONEINFO_TEMPLATE
    ZoneID As Integer ' Index of Room
    ZoneDescription As String
    ExitAreas As String ' Formated as N1E1S1W1 or N0E0S0W0
    SafeArea As Boolean ' No encouters
    HostileLevel As Integer ' The higher, the worser it gets
    RecallPoint As Integer
    ExitPointToOtherWorld As Boolean
    StartingArea As Boolean
End Type

Public T_ZONEINFO() As T_ZONEINFO_TEMPLATE
Public ZoneInfoCounter As Integer
Public CurrentRoom As Integer

Public Sub AddZone()
    
    
End Sub


Public Sub InitZonesDisplay()

    Dim lngCountX As Long
    Dim lngCountY As Long
    Dim lngOffset As Long
    Dim intCount As Integer
    
    lngOffset = 455
    intCount = 1
    
    ReDim T_ZONEINFO(1 To 483)
    
    For lngCountX = 1 To 23
        For lngCountY = 1 To 21
            Load frmMain.cmdZoneMap(intCount)
            frmMain.cmdZoneMap(intCount).Left = lngCountX * lngOffset
            frmMain.cmdZoneMap(intCount).Top = lngCountY * lngOffset
            frmMain.cmdZoneMap(intCount).Visible = True
            T_ZONEINFO(intCount).ZoneID = intCount
            intCount = intCount + 1
        Next
    Next
        
    ZoneInfoCounter = 483 ' Static

End Sub
