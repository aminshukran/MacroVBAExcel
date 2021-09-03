Attribute VB_Name = "Module1"

Public Sub importtextfile()
    Dim textfile As Workbook
    Dim openfiles() As Variant
    Dim i As Integer
    
    openfiles = getfiles
    
    Application.ScreenUpdating = False
    
    For i = 1 To Application.CountA(openfiles)
    
        Set textfile = Workbooks.Open(openfiles(i))
        
        textfile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        ActiveSheet.Name = textfile.Name
        
        Application.CutCopyMode = False
        
        textfile.Close
    
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Public Function getfiles() As Variant

    getfiles = Application.GetOpenFilename(Title:="Select file(s) to import", MultiSelect:=True)
    
    
End Function
