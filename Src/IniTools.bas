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
Sub ReadIniArray(Section As String, _
                 KeyName As String, _
                 DefValue As String, _
                 IniFile As String, _
                 Separator As String, _
                 ByRef Lines As Dictionary)
  Dim RawLines As Variant
  Dim I As Variant
  Dim S As String
  
  RawLines = Split(ReadIniValue(Section, KeyName, DefValue, IniFile), Separator)
  If Not IsArrayEmpty(RawLines) Then
    For Each I In RawLines
      S = Trim(I)
      If (S <> "") And Not Lines.Exists(S) Then
        Lines.Add S, Empty
      End If
    Next
  End If

End Sub

Sub GetCurrentFolders(Doc As ModelDoc2, ByRef Lines As Dictionary)

  Dim I As Variant
  Dim AFeature As Feature
  Dim Name As String
  
  For Each I In Doc.FeatureManager.GetFeatures(True)
    Set AFeature = I
    Name = AFeature.Name
    If (AFeature.GetTypeName = "FtrFolder") And (Not Name Like "*__EndTag__*") And Not Lines.Exists(Name) Then
      Lines.Add Name, Empty
    End If
  Next

End Sub
