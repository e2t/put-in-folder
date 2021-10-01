Attribute VB_Name = "Main"
Option Explicit

Const SettingsName = "Settings.ini"
Const SectionName = "Main"
Const KeyName = "Lines"

Public FSO As FileSystemObject
Public Lines As Dictionary 'as String

Dim swApp As Object
Dim CurrentDoc As ModelDoc2
Dim SelectedCount As Integer
Dim SettingsPath As String

Sub Main()
  
  Set swApp = Application.SldWorks
  Set CurrentDoc = swApp.ActiveDoc
  If CurrentDoc Is Nothing Then Exit Sub
  If CurrentDoc.GetType <> swDocASSEMBLY Then Exit Sub
  
  SelectedCount = CurrentDoc.SelectionManager.GetSelectedObjectCount2(0)
  If SelectedCount <= 0 Then
    MsgBox "Выделите компоненты для перемещения!", vbCritical
    Exit Sub
  End If
  
  Set FSO = New FileSystemObject
  SettingsPath = FSO.BuildPath(swApp.GetCurrentMacroPathFolder, SettingsName)
  If Not FSO.FileExists(SettingsPath) Then
    CreateDefaultSettingsFile SettingsPath
  End If
  
  Set Lines = New Dictionary
  Lines.Add "(В новую папку...)", Empty
  ReadIniArray SectionName, KeyName, "There is not a line", SettingsPath, ",", Lines
  GetCurrentFolders CurrentDoc, Lines
  
  MainForm.Show
  
End Sub

Function MainFormInit() 'hide

  Dim I As Variant
    
  For Each I In Lines
    MainForm.ListBoxName.AddItem I
  Next
  MainForm.ListBoxName.ListIndex = 0

End Function

Sub CreateDefaultSettingsFile(SettingsPath As String)

  Dim IniFile As TextStream
  
  Set IniFile = FSO.CreateTextFile(SettingsPath, False, True)
  IniFile.WriteLine "[" & SectionName & "]"
  IniFile.WriteLine KeyName & " = " & _
    "Примененные сборки, Примененные, Стандартные изделия, Покупные, Прочее, Сварные швы"
  IniFile.Close

End Sub

Function OpenSettingsFile() 'hide

  Shell "notepad """ & SettingsPath & """", vbNormalFocus

End Function

Function RunIfSelected() 'hide

  Dim FolderName As String

  If MainForm.ListBoxName.ListIndex = 0 Then
    FolderName = InputBox("Имя новой папки:")
  ElseIf MainForm.ListBoxName.ListIndex > 0 Then
    FolderName = MainForm.ListBoxName.Text
  Else
    MsgBox "Выберите раздел!", vbCritical
  End If
  
  If FolderName <> "" Then
    MainForm.Hide
    Run FolderName
    ExitApp
  End If

End Function

Sub Run(FolderName As String)

  Dim Feat As Feature
  Dim I As Integer
  Dim ComponentsToMove() As Component2
  Dim Asm As AssemblyDoc
  
  Set Feat = SearchFeature(FolderName, CurrentDoc.FeatureManager)
  If Feat Is Nothing Then
    Set Feat = CurrentDoc.FeatureManager.InsertFeatureTreeFolder2(swFeatureTreeFolder_Containing)
    Feat.Name = FolderName
  Else
    ReDim ComponentsToMove(SelectedCount - 1)
    For I = 0 To SelectedCount - 1
      Set ComponentsToMove(I) = CurrentDoc.SelectionManager.GetSelectedObjectsComponent4(I + 1, 0)
    Next
    Set Asm = CurrentDoc
    Asm.ReorderComponents ComponentsToMove, Feat, swReorderComponents_LastInFolder
  End If

End Sub

Function ExitApp() 'hide

  Unload MainForm
  End

End Function
