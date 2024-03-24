Attribute VB_Name = "GenerateReturn"
Sub GenerateReturn()
        On Error GoTo ErrorHandler
    
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    Dim wbData As Workbook
    Dim wsData As Worksheet
    Dim studentName As String
    Dim studentID As String
    Dim newFileName As String
    Dim destinationPath As String
    
    ' Set the path of the template file
    Dim templatePath As String
    templatePath = "C:\path\to\file\template.xlsx"
    
    ' Set the path of the data file containing Candidates
    Dim dataFilePath As String
    dataFilePath = "C:\path\to\file\candidates.xlsx"
    
    ' Set the destination path where the generated grade sheets will be saved
    destinationPath = "C:\out\path"
    
   ' Open the data file containing student names and IDs
    Set wbData = Workbooks.Open(dataFilePath)
    Set wsData = wbData.Sheets(1)
    
    ' Loop through the list of wards in the data file
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow ' Assuming candidates start from row 2, change if needed
        ' Open the template file
        Set wbTemplate = Workbooks.Open(templatePath)
        Set wsTemplate = wbTemplate.Sheets(1)
        
        ' Get data from data file
        ' Can Add more columns if needed
        Ward = wsData.Cells(i, 1).Value
        Name = wsData.Cells(i, 2).Value
        Code = wsData.Cells(i, 3).Value
        Electorate = wsData.Cells(i, 4).Value
        Limit = wsData.Cells(i, 5).Value
        
        ' Add in the data
        On Error Resume Next
        wsTemplate.Range("N4").Value = Code
        wsTemplate.Range("D10").Value = Ward
        wsTemplate.Range("D14").Value = Electorate
        wsTemplate.Range("D18").Value = Name
        wsTemplate.Range("M20").Value = Limit
        On Error GoTo 0
        
        If Err.Number <> 0 Then
            MsgBox "Error setting values in the template file: " & Err.Description
            wbTemplate.Close SaveChanges:=False
            Exit Sub
        End If
        
        ' Generate a unique file name for each return
        newFileName = Ward & Name & ".xlsx"
        
        ' Check if a file with the same name already exists in the destination folder
        If FileExists(destinationPath & newFileName) Then
            MsgBox "A file with the name '" & newFileName & "' already exists in the destination folder."
            wbTemplate.Close SaveChanges:=False
            Exit Sub
        End If
        
        ' Save the return with ward name in the desired destination folder
        wbTemplate.SaveCopyAs destinationPath & newFileName
        
        ' Close the newly created gradesheet
        wbTemplate.Close SaveChanges:=False
    Next i
    
    ' Close the data file
    wbData.Close SaveChanges:=False
    
    ' Release object references
    Set wsTemplate = Nothing
    Set wbTemplate = Nothing
    Set wsData = Nothing
    Set wbData = Nothing
    
    MsgBox "Grade sheets generated successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function


