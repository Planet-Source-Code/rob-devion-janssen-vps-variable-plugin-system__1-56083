Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    modConsole.ConAcquire
    Con80Line "*"
    frmCore.Log "VPS - Variable Plugin System (C) Devion" & vbCrLf
    frmCore.Log ""
    frmCore.Log "Date of execution: " & Now
    frmCore.Log ""
    frmCore.Log "CPU Type: " & SystemInfo.CPUVersion
    frmCore.Log "Available Maximum Plugin memory: " & Round((SystemInfo.MemoryFree / 60) / 1024000, 0) & " MB"
    frmCore.Log "Available System memory: " & Round(SystemInfo.MemoryTotal / 1024000, 0) & " MB"
    frmCore.Log "OS Detected: " & SystemInfo.WinName & SystemInfo.WinVersion
    frmCore.Log ""
    ConWrite "Generating seed...": modCore.InitSeed
    frmCore.Log ""
    modCore.InitMemory 5000
    frmCore.Log ""
    frmCore.Log "Initialising plugins... (DIR: /plugins)"
    frmCore.Log ""
    modCore.InitPlugins
    frmCore.Log ""
    frmCore.Log "Start plugin inits..."
    frmCore.Log ""
    modCore.StartPluginInits
    frmCore.Log ""
    frmCore.Log "Init complete..."
    frmCore.Log ""
    frmCore.Log ""
    frmCore.Log "To close down VPS and plugins close the window or press CTRL+C."
    Load frmCore
    frmCore.goIntoWaitState
    frmCore.Log ""
End Sub

Public Sub Con80Line(strChar As String)

    Dim IntCount As Integer
    
    ConWrite vbCrLf
    For IntCount = 1 To 80
        ConWrite strChar
    Next IntCount
    ConWrite vbCrLf
    
End Sub

Public Sub App_Close()
    frmCore.Log "Closing application"
    frmCore.Log " - Unloading plugins..."
    modCore.UnloadPlugins
    frmCore.Log " - Freeing memory stack..."
    modCore.UnloadMemory
    modConsole.ConRelease
End Sub
