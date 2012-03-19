Dim included    

Call Include


Sub Include
  If (Not included) Then
    Dim libPath, fso, fld, f
    libPath = modulePath() & "lib\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(libPath)
    If fld.Files.Count <> 0 Then
      For Each f In fld.Files
        Call ExecuteGlobal(fso.OpenTextFile(libPath & f.Name, 1).ReadAll)
      Next
    End If
    Set fso = Nothing
    included = True
  End If 
End Sub