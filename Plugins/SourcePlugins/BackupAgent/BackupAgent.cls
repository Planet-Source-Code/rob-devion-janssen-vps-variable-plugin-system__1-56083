VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackupAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'fstrGetName (no args) return result should be a string giving the friendly name of the plugin
'fstrGetFunction (no args) return result should be a string giving the function that initializes the plugin
'fClose (no args) allows plugin to close itself before VPS terminates
'registerparent (refObject) allows the plugin to be able to return calls.


Public Function fStrGetName() As String
    fStrGetName = "Backup Agent"
End Function

Public Function fstrGetFunction() As String
    fstrGetFunction = "InitBackup"
End Function

Public Function fClose() As String
    fClose = "Ok"
End Function

Public Function RegisterParent(frmObject As Object) As String
    Set Parent = frmObject
    RegisterParent = "Ok"
End Function

' -----'

Public Sub InitBackup()
On Error GoTo Errhandler

        Dim IntCount As Long
        
        If Dir$(App.Path & "\BackupAgent.ini") <> "" Then
            'Load ini
            
            modCommon.Dircount = CLng(modIniHandling.f_mfncGetFromIni("BackupAgent", "DirCount", App.Path & "\BackupAgent.ini"))
            modCommon.mysnIncrement = IIf(f_mfncGetFromIni("BackupAgent", "Increment", App.Path & "\BackupAgent.ini") = "1", True, False)
            modCommon.BackupInterval = CLng(modIniHandling.f_mfncGetFromIni("BackupAgent", "Interval", App.Path & "\BackupAgent.ini"))
            
'           [PathX]
'           Source = Path
'           Dest = Path
    

            
            ReDim Dirs(1 To modCommon.Dircount)
            
            For IntCount = 1 To modCommon.Dircount
                Dirs(IntCount).SourcePath = modIniHandling.f_mfncGetFromIni("Path" & CStr(IntCount), "Source", App.Path & "\BackupAgent.ini")
                Dirs(IntCount).DestPath = modIniHandling.f_mfncGetFromIni("Path" & CStr(IntCount), "Dest", App.Path & "\BackupAgent.ini")
            Next IntCount
            
            TellParent "Backup Agent loaded " & CStr(modCommon.Dircount) & " directories for backup..."
            TellParent "Increment Backup is " & IIf(modCommon.mysnIncrement = True, "On", "Off")
            TellParent "Backup Interval is " & CStr(modCommon.BackupInterval)
            
        Else
        
            TellParent "Backup Agent was not yet configured. Config window shown."
            Load frmBackupAgent
            frmBackupAgent.Show vbModal
            
            If Dir$(App.Path & "\BackupAgent.ini") = "" Then
                TellParent "Backup Agent configuration failed? No BackupAgent.ini file found."
                Exit Sub
            End If
            
            modCommon.Dircount = CLng(modIniHandling.f_mfncGetFromIni("BackupAgent", "DirCount", App.Path & "\BackupAgent.ini"))
            modCommon.mysnIncrement = IIf(f_mfncGetFromIni("BackupAgent", "Increment", App.Path & "\BackupAgent.ini") = "1", True, False)
            modCommon.BackupInterval = CLng(modIniHandling.f_mfncGetFromIni("BackupAgent", "Interval", App.Path & "\BackupAgent.ini"))
            
'           [PathX]
'           Source = Path
'           Dest = Path
    
            ReDim Dirs(1 To modCommon.Dircount)
            
            For IntCount = 1 To modCommon.Dircount
                Dirs(IntCount).SourcePath = modIniHandling.f_mfncGetFromIni("Path" & CStr(IntCount), "Source", App.Path & "\BackupAgent.ini")
                Dirs(IntCount).DestPath = modIniHandling.f_mfncGetFromIni("Path" & CStr(IntCount), "Dest", App.Path & "\BackupAgent.ini")
            Next IntCount
            
            TellParent "Backup Agent loaded " & CStr(modCommon.Dircount) & " directories for backup..."
            TellParent "Increment Backup is " & IIf(modCommon.mysnIncrement = True, "On", "Off")
            TellParent "Backup Interval is " & CStr(modCommon.BackupInterval)
        End If
    
        TellParent ""
        TellParent "Preparing backup..."
        modCommon.CountFiles
        TellParent "Total files for backup: " & CStr(modCommon.TotalFiles)
        TellParent ""
        Load frmBackup
        
        frmBackup.tmrBackup.Interval = 1000
        frmBackup.lngTimer = CLng(Round(modCommon.BackupInterval * 60))
        frmBackup.tmrBackup.Enabled = True
        TellParent "Backup Agent Initialised and running in messageloop."
    
        Exit Sub
Errhandler:
        
        MsgBox "Error in here."
        
    
End Sub
