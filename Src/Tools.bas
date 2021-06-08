Attribute VB_Name = "Tools"
Option Explicit

Function SearchFeature(FeatureName As String, FeatMgr As FeatureManager) As Feature

  Dim I As Variant
  Dim Feat As Feature

  For Each I In FeatMgr.GetFeatures(True)
    Set Feat = I
    If Feat.Name = FeatureName Then
      Set SearchFeature = Feat
      Exit For
    End If
  Next
  
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean

  IsArrayEmpty = True
  On Error Resume Next
  IsArrayEmpty = LBound(anArray) > UBound(anArray)

End Function
