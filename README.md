# xlWeeties

Simply copy and paste the code below into a new module named 'xlWeeties' in your 'personal.xlsb' file.

```vbnet
  Public VBProj As VBIDE.VBProject
  Public VBComp As VBIDE.VBComponent
  Public CodeMod As CodeModule
  
  Sub xlWeeties()
    
    Set VBProj = ActiveWorkbook.VBProject
    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    Set CodeMod = VBComp.CodeModule
    
    Dim strResult As String
    Dim objHTTP As Object
    Dim URL As String
    
    Dim NewModuleName As String
      
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    URL = "https://raw.githubusercontent.com/jaykilleen/xlWeeties/master/hello_world.bas"
    NewModuleName = ExtractModuleNameFromURL(URL)
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objHTTP.send ("keyword=php")
    strResult = objHTTP.responseText
    Debug.Print strResult
    
    For Each VBComp In VBProj.VBComponents
      If VBComp.Type = vbext_ct_Document Then
  
          With CodeMod
            .DeleteLines 1, .CountOfLines
          End With
          Else
            VBProj.VBComponents.Remove VBComp
          End If
      Next VBComp
      
    Call AddModuleToProject(NewModuleName, strResult)
    
  End Sub
  Sub AddModuleToProject(NewModuleName, Code As String)
    
    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    VBComp.Name = NewModuleName
    
    Set CodeMod = VBComp.CodeModule
    
    With CodeMod
      .InsertLines .CountOfLines + 1, Code
    End With
  End Sub
  
  Function ExtractModuleNameFromURL(s As String)
  
    Dim iLastBackslash As Integer
  
    iLastBackslash = InStrRev(s, "/")
    
    s = Replace(Right(s, (Len(s) - iLastBackslash)), ".bas", "")
    
  End Function
```

