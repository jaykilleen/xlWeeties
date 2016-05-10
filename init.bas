Public vbProj As VBIDE.VBProject
Public VBComp As VBIDE.VBComponent
Public CodeMod As CodeModule
Public wbProject As Workbook
  
Sub xlWeeties()
  
  Set vbProj = ActiveWorkbook.VBProject
  Set VBComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
  Set CodeMod = VBComp.CodeModule
    
  Dim strResult As String
  Dim objHTTP As Object
  Dim URL As String
      
  Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
  URL = "https://raw.githubusercontent.com/jaykilleen/xlWeeties/master/hello_world.bas"
  objHTTP.Open "GET", URL, False
  objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
  objHTTP.SetRequestHeader "Content-type", "application/x-www-form-urlencoded"
  objHTTP.Send ("keyword=php")
  strResult = objHTTP.responseText
  Debug.Print strResult
    
  For Each VBComp In vbProj.VBComponents
    If VBComp.Type = vbext_ct_Document Then
  
        With CodeMod
          .DeleteLines 1, .CountOfLines
        End With
        Else
          vbProj.VBComponents.Remove VBComp
        End If
    Next VBComp
  
  Call CreateProject(URL)
  Call AddModuleToProject(URL, strResult)
    
End Sub
Sub CreateProject(URL)
  
  Set vbProj = ActiveWorkbook.VBProject
  
  Dim iLastBackslash As Integer
  Dim NewWorkbookName As String
  Dim Path As String: Path = "D:\Projects\excel\"
  
  iLastBackslash = InStrRev(URL, "/")
  NewWorkbookName = Replace(Right(URL, (Len(URL) - iLastBackslash)), ".bas", "")
  
  'Adding New Workbook
  Workbooks.Add
  
  'Create the folder
  MkDir Path & NewWorkbookName
  'Saving the Workbook
  ActiveWorkbook.SaveAs Filename:=Path & NewWorkbookName & ".xlsx"
  
  Set wbProj = Workbooks(NewWorkbookName & ".xlsx")
  
End Sub
Sub AddModuleToProject(URL, Code As String)
    
  Dim iLastBackslash As Integer
  Dim NewModuleName As String
  Set VBComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    
  iLastBackslash = InStrRev(URL, "/")
  NewModuleName = Replace(Right(URL, (Len(URL) - iLastBackslash)), ".bas", "")
    
  VBComp.Name = NewModuleName
    
  Set CodeMod = VBComp.CodeModule
    
  With CodeMod
    .InsertLines .CountOfLines + 1, Code
  End With
End Sub
  
