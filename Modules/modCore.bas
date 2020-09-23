Attribute VB_Name = "modCore"
Option Explicit

Public SystemInfo As New cSystemInfo
Public strSeed As String

Public Type T_OBJECTS
    objPLUGIN As Object
    strFunction As String
    strFriendlyName As String
    strFilename As String
    
End Type

Public Type T_MEMORY
    MemoryOffset As Long
    MemoryLength As Long
    Memoryname As String
End Type

Public ObjectLibrary() As T_OBJECTS

Public Memory() As Byte
Public MemoryLookup() As T_MEMORY

Public intPluginCount As Integer
Public intMemoryCount As Integer

Public intErrors As Integer

Public Sub InitPlugins()
On Error Resume Next

    Dim strtemp As String
    Dim strTemp2 As String
    Dim IntCount As Integer
    
    strtemp = Dir(App.Path & "\plugins\*.DLL")
    
    While strtemp <> ""
        IntCount = IntCount + 1
        strtemp = Dir
    Wend
    
    strtemp = Dir(App.Path & "\plugins\*.DLL")
    
    If IntCount = 0 Then GoTo PluginComplete
    
    frmCore.Log "Plugin(s) DLL found: " & CStr(IntCount)
    frmCore.Log "Proceeding to initialize plugins..."
    frmCore.Log ""
    
    While strtemp <> ""
        strTemp2 = App.Path & "\plugins\" & strtemp ' Prefix the path.
        SetupNewPlugin strTemp2, LCase$("VPS." & Left(strtemp, InStr(1, strtemp, ".") - 1))
        strtemp = Dir
        DoEvents
    Wend
    
PluginComplete:
    
    If intPluginCount <> 0 Then
        
        IntCount = CStr(UBound(ObjectLibrary))
        If intPluginCount = 0 Then
            frmCore.Log "Found plugins but none could be loaded..."
            frmCore.Log "Exiting VPS..."
            frmCore.Log ""
        Else
            frmCore.Log ""
            frmCore.Log "Loaded " & IntCount & " plugin(s) with " & CStr(intErrors) & " errors while initializing them."
        End If
        
    End If
    
    frmCore.Log ""
    frmCore.Log "Plugin load complete."
    DoEvents
    
End Sub

Public Sub SetupNewPlugin(strFilename As String, DLLNAME As String)
    intErrors = 0
    intPluginCount = intPluginCount + 1
    ReDim Preserve ObjectLibrary(1 To intPluginCount)
    
    modDLL.Register strFilename
    

    With ObjectLibrary(intPluginCount)
        Set .objPLUGIN = CreateObject(DLLNAME)
        .strFilename = strFilename
        .strFriendlyName = CallByName(.objPLUGIN, "fstrGetName", VbMethod)
        .strFunction = CallByName(.objPLUGIN, "fstrGetFunction", VbMethod)
        'On Error Resume Next
        CallByName .objPLUGIN, "RegisterParent", VbMethod, frmCore
        'On Error GoTo ErrHandler
    End With
    
    frmCore.Log " - Plugin found: " & DLLNAME & " (" & ObjectLibrary(intPluginCount).strFriendlyName & ")"
    
    Exit Sub
ErrHandler:
    intErrors = intErrors + 1
    intPluginCount = intPluginCount - 1
    DoEvents
End Sub

Public Sub InitMemory(lngMemoryLength As Long)

    Dim lngTemp As Long
    
    frmCore.Log "Initializing memory stack... (" & Round(lngMemoryLength / 1024000, 0) & " MB)"
    ReDim Memory(1 To lngMemoryLength)
    
    
    frmCore.Log "M_Alloc: " & lngMemoryLength & " bytes allocated..."
End Sub

Public Sub UnloadMemory()
    ReDim Memory(0 To 0)
End Sub

Public Sub UnloadPlugins()
    On Error Resume Next
    
    Dim intcounter As Integer
    Dim strtemp As String
  
    ReDim ObjectLibrary(0 To 0)
    
End Sub

Public Sub InitSeed()
    strSeed = modMD5.CalculateMD5(CStr(Timer) & "--VPS--" & CStr(Timer) & "--SEEDER")
    frmCore.Log "(" & strSeed & ")"
    frmCore.Log ""
    
End Sub

Public Sub StartPluginInits()

    Dim IntCount As Integer
    
    For IntCount = 1 To intPluginCount
        frmCore.Log " - Plugin_init : " & ObjectLibrary(IntCount).strFriendlyName
        frmCore.Log ""
        Call CallByName(ObjectLibrary(IntCount).objPLUGIN, ObjectLibrary(IntCount).strFunction, VbMethod)
        frmCore.Log ""
    Next IntCount

End Sub
