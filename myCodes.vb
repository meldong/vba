Sub Open_Workbook_Dialog()

    Dim fileName As Variant
    
    fileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")
    
    If fileName <> False Then
        Workbooks.Open FileName:=fileName, _
        UpdateLinks:=3, _
        
        
    End If
    
End Sub
