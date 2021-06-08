Attribute VB_Name = "IniTools"
Option Explicit

Declare PtrSafe Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringW" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As String, _
  ByVal lpDefault As String, _
  lpReturnedString As Any, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long
  
'Only UTF-16LE
Public Function ReadIniValue(Section As String, _
                             KeyName As String, _
                             DefValue As String, _
                             IniFile As String) As String
  Const MaxSize = 2048
  Dim RetVal As Long
  Dim Buf(MaxSize - 1) As Byte
  
  RetVal = GetPrivateProfileString(StrConv(Section, vbUnicode), _
                                   StrConv(KeyName, vbUnicode), _
                                   StrConv(DefValue, vbUnicode), _
                                   Buf(0), _
                                   MaxSize, _
                                   StrConv(IniFile, vbUnicode))
  If RetVal > 0 Then
    ReadIniValue = Left(Buf, RetVal)
  End If

End Function

'Only UTF-16LE
Function ReadIniArray(Section As String, _
                      KeyName As String, _
                      DefValue As String, _
                      IniFile As String, _
                      Separator As String) As String()
  Dim RawLines As Variant
  Dim Lines() As String
  Dim Count As Integer
  Dim I As Variant
  Dim S As String
  
  RawLines = Split(ReadIniValue(Section, KeyName, DefValue, IniFile), Separator)
  If IsArrayEmpty(RawLines) Then Exit Function
  
  Count = 0
  ReDim Lines(UBound(RawLines))
  For Each I In RawLines
    S = Trim(I)
    If S <> "" Then
      Lines(Count) = S
      Count = Count + 1
    End If
  Next
  If Count > 0 Then
    ReDim Preserve Lines(Count - 1)
  Else
    Erase Lines
  End If
  ReadIniArray = Lines

End Function
