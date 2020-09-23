Attribute VB_Name = "modIniHandling"
'**************************************
'Windows API/Global Declarations for :.INI read/write routines
'**************************************

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function f_mfncGetFromIni(s_strSectionHeader As String, s_strVariableName As String, s_strFileName As String) As String
    Dim s_strReturn As String
    s_strReturn = String(255, Chr(0))
    f_mfncGetFromIni = Left$(s_strReturn, GetPrivateProfileString(s_strSectionHeader, ByVal s_strVariableName, "", s_strReturn, Len(s_strReturn), s_strFileName))
End Function

Private Function f_mfncParses_string(s_strIn As String, intOffset As Integer, strDelimiter As String) As String
      
    If Len(s_strIn) = 0 Or intOffset = 0 Then
        f_mfncParses_string = ""
        Exit Function
    End If
    
    'Declare local variables
    Dim i_intStartPos As Integer
    ReDim i_intDelimPos(10) As Integer
    Dim i_intStrLen As Integer
    Dim i_intNoOfDelims As Integer
    Dim i_intCount As Integer
    Dim s_strQuotationMarks As String
    Dim i_intInsideQuotationMarks As Integer
    
    s_strQuotationMarks = Chr(34) & Chr(147) & Chr(148)
    i_intInsideQuotationMarks = False

    For i_intCount = 1 To Len(s_strIn)
        'If character is a double-quote then toggle the In Quotation flag

        If InStr(s_strQuotationMarks, Mid$(s_strIn, i_intCount, 1)) <> 0 Then
            i_intInsideQuotationMarks = (Not i_intInsideQuotationMarks)
        End If
        If (Not i_intInsideQuotationMarks) And (Mid$(s_strIn, i_intCount, 1) = strDelimiter) Then
        i_intNoOfDelims = i_intNoOfDelims + 1
        If (i_intNoOfDelims Mod 10) = 0 Then
            ReDim Preserve i_intDelimPos(i_intNoOfDelims + 10)
        End If
        i_intDelimPos(i_intNoOfDelims) = i_intCount
    End If
Next i_intCount

If intOffset > (i_intNoOfDelims + 1) Then
    f_mfncParses_string = ""
    Exit Function
End If

If intOffset = 1 Then
    i_intStartPos = 1
End If

If intOffset = (i_intNoOfDelims + 1) Then

   If Right$(s_strIn, 1) = strDelimiter Then
        i_intStartPos = -1
        i_intStrLen = -1
        f_mfncParses_string = ""
        Exit Function
    Else
        i_intStrLen = Len(s_strIn) - i_intDelimPos(intOffset - 1)
    End If
End If

If i_intStartPos = 0 Then
    i_intStartPos = i_intDelimPos(intOffset - 1) + 1
End If

If i_intStrLen = 0 Then
    i_intStrLen = i_intDelimPos(intOffset) - i_intStartPos
End If

f_mfncParses_string = Mid$(s_strIn, i_intStartPos, i_intStrLen)
End Function

Public Function f_mfncWriteIni(s_strSectionHeader As String, s_strVariableName As String, s_strValue As String, s_strFileName As String) As Integer
  f_mfncWriteIni = WritePrivateProfileString(s_strSectionHeader, s_strVariableName, s_strValue, s_strFileName)
End Function
 
 

