Attribute VB_Name = "pdfmaker"
Sub ConvertExcelToPDF()
    ' On Error GoTo ErrorHandler
    
    Dim FolderPath As String
    Dim FileName As String
    Dim filePath As String
    Dim NewFilePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim TotalFiles As Integer
    Dim ProcessedFiles As Integer
    
    ' Set the folder path where the Excel files are located
    FolderPath = "C:\path\to\folder\"
    
    
    ' Get the first Excel file in the folder
    FileName = Dir(FolderPath & "*.xlsx")
    
    ' Count the total number of files to process
    TotalFiles = 0
    Do While FileName <> ""
        TotalFiles = TotalFiles + 1
        FileName = Dir
    Loop
    MsgBox "Total Files: " & TotalFiles
    ' Reset the FileName to get the first Excel file again
    FileName = Dir(FolderPath & "*.xlsx")
    
    ' Initialize the count of processed files
    ProcessedFiles = 0
    
    ' Loop through all Excel files in the folder
    Do While FileName <> ""
        ' Build the full file path
        filePath = FolderPath & FileName
        ' Optional change the folder location
        ' FilePath2 = "C:\path\to\another\folder\"
        
        ' Build the new file path for the PDF
        NewFilePath = Replace(filePath, ".xlsx", ".pdf", , , vbTextCompare)
        
        ' Check if the PDF file already exists

            ' Open the Excel file
            Set wb = Workbooks.Open(filePath)
            
            ' Check if the workbook has at least one sheet
            If wb.Sheets.Count > 0 Then
                ' Get the first sheet in the workbook
                Set ws = wb.Sheets(1)
                
                ' Save the first sheet as PDF
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=NewFilePath
                
                ' Close the Excel file
                wb.Close SaveChanges:=False
                Set wb = Nothing
                
                ' Increment the count of processed files
                ProcessedFiles = ProcessedFiles + 1
                
            Else
                MsgBox "The workbook does not contain any sheets. Skipping conversion."
                wb.Close SaveChanges:=False
                Set wb = Nothing
            End If
        
        ' Reset the FileName to get the next Excel file in the folder
        FileName = Dir
    Loop
    
    MsgBox "Conversion to PDF completed successfully! Total Files: " & TotalFiles & ", Processed Files: " & ProcessedFiles
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
    Resume Next
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

