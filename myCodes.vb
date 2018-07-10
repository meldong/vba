Sub Open_Workbook_Dialog()

    Dim fileName As Variant
    
    fileName = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")
    
    If fileName <> False Then
        Workbooks.Open fileName
    End If
    
End Sub

Sub Create_Object_Demo()
    
    ' file handeling
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\myTXT.txt", True)
    a.WriteLine ("This is a test.")
    a.Close
    
    ' excel handeling
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlApp.Visible = False
    xlSheet.Cells(1, 1).Value = "This is column A, row 1"
    xlSheet.SaveAs "C:\myXLS.xlsx"
    Debug.Print xlApp.Version
    xlApp.Quit
    Set xlApp = Nothing
    
End Sub
