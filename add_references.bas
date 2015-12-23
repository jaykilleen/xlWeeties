Sub AddReference()
  Dim References As Variant
  Dim VBAEditor As VBIDE.VBE
  Dim vbProj As VBIDE.VBProject
  Dim chkRef As VBIDE.Reference
  Dim BoolExists As Boolean

  Set VBAEditor = Application.VBE
  Set vbProj = ActiveWorkbook.VBProject

  On Error Resume Next
    vbProj.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"
    vbProj.References.AddFromFile "C:\WINDOWS\system32\winhttp.dll"
    vbProj.References.AddFromFile "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
  On Error GoTo 0

  Set vbProj = Nothing
  Set VBAEditor = Nothing
End Sub