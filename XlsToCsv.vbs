REM XlsToCsv.vbs D:\src\lfillMon\logs\ D:\src\lfillMon\out.csv
if WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: ExcelToCsv <xls/xlsx source folder> <csv destination file>"
    Wscript.Quit
End If

' src_folder = "D:\src\lfillMon\logs\"
src_folder = WScript.Arguments.Item(0)
dest_file = WScript.Arguments.Item(1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
' dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
Set objOutFile = objFSO.CreateTextFile(dest_file,True)
objOutFile.Write("location,date,Sample type,Purging time mins,CH4 v/v%,CO2 %v/v,O2 %v/v,Balance %v/v,CO ppmv,H2S ppmv,Barometric Pressure mb,Relative Pressure mb" & vbCrLf)

For Each oFile In objFSO.GetFolder(src_folder).Files
  If UCase(objFSO.GetExtensionName(oFile.Name)) = "XLS" Then
    convert(oFile.Name)
  End if
Next

objOutFile.Close

Sub convert (filename)
    src_file = src_folder & filename

    Dim oExcel
    Set oExcel = CreateObject("Excel.Application")

    Dim oBook
    Set oBook = oExcel.Workbooks.Open(src_file)
    ' Set oSheet = oBook.Worksheets("20.09.2017")

    Set currentWorkSheet = oExcel.ActiveWorkbook.Worksheets(1)
    Set Cells = currentWorksheet.Cells
    usedRowsCount = currentWorkSheet.UsedRange.Rows.Count
    For row = 1 to (usedRowsCount-1)
        If Cells(row ,4).Value = "Final" Then
        '    WScript.Echo (Cells(row ,2).Value)
        objOutFile.Write(Cells(row ,2).Value & "," _
        & Cells(row ,3).Value  & "," _
        & Cells(row ,4).Value & "," _
        & Cells(row ,5).Value & "," _
        & Cells(row ,8).Value & "," _
        & Cells(row ,14).Value  & "," _
        & Cells(row ,16).Value  & "," _
        & Cells(row ,18).Value  & "," _
        & Cells(row ,20).Value  & "," _
        & Cells(row ,22).Value  & "," _
        & Cells(row ,26).Value & "," _
        & Cells(row ,28).Value & vbCrLf)
        End If
    Next	
    oExcel.Quit
End Sub
