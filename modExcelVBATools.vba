Option Explicit

'MIT License
'
'Copyright (c) 2016 W. Jacob Harris
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


'Place in Module1 for whatever notebook (must be xlsm file)

'Note some features require ADO 2.x and Microsoft Scripting Runtime

Sub Button1_Click()
    Application.ScreenUpdating = False 'Suppress screen updates for speed
    Application.DisplayAlerts = False 'Disable prompts to ensure we're not harassing the user.
    Application.Calculation = xlCalculationManual
    Call ParseData("SourceData", "ParsedDataForColC", "C")
    Call ParseData("ParsedDataForColC", "ParsedDataFinal", "E") 'Final sheet
    'Call RemoveSheet("ParsedDataForColC") 'Optional removal of temporary sheet
    Application.DisplayAlerts = True 'Re-enable prompts
    Application.ScreenUpdating = True 'Disable suppression of screen updates
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Sub TestInsertColumn()
    Application.ScreenUpdating = False 'Suppress screen updates for speed
    Application.DisplayAlerts = False 'Disable prompts to ensure we're not harassing the user.
    Application.Calculation = xlCalculationManual
    
    
    
    Dim wksheetSource As Worksheet
    Set wksheetSource = Worksheets("SampleColumnInsert")
    
    InsertCellsToLeft wksheetSource, "C", "Lattitude"
    SetFormula wksheetSource, "C", "=SimpleLookup(SQLLookupSource!A:C, ""A"", A{R}, ""B"")" 'Adds formula to lookup lattitude
    
    InsertCellsToLeft wksheetSource, "D", "Longitude"
    SetFormula wksheetSource, "D", "=SimpleLookup(SQLLookupSource!A:C, ""A"", A{R}, ""C"")" 'Adds formula to lookup longitude
    
    wksheetSource.Calculate
    Set wksheetSource = Nothing
    
    
    Application.DisplayAlerts = True 'Re-enable prompts
    Application.ScreenUpdating = True 'Disable suppression of screen updates
    Application.Calculation = xlCalculationAutomatic
    
    
End Sub

Public Sub RemoveSheet(worksheetName As String)
    Dim wksht As Worksheet
    Set wksht = Worksheets(worksheetName)
    wksht.Delete
    Set wksht = Nothing
    
End Sub

'''Takes sourceWorksheet and creates a target worksheet by splitting a column
Public Sub ParseData(sourceWorksheet As String, targetWorksheet As String, columnToSplit As String)
    On Error Resume Next
    Dim wksheetSource As Worksheet
    Dim wksheetTarget As Worksheet
    
    Set wksheetSource = Worksheets(sourceWorksheet) 'Get a worksheet by the name requested
    Set wksheetTarget = Worksheets(targetWorksheet) 'Get a worksheet by the name requested
    
    If wksheetTarget Is Nothing Then
        Set wksheetTarget = Worksheets.Add(After:=wksheetSource)
        If wksheetTarget Is Nothing Then
            Call MsgBox("Could not create necessary worksheet", vbCritical, "Error")
            Exit Sub
        End If
        wksheetTarget.Name = targetWorksheet
        
    End If
    
    'Clear target worksheet
    wksheetTarget.UsedRange.cells.ClearContents
    
    'Copy data from source to target so we can work
    wksheetSource.UsedRange.cells.Copy
    wksheetTarget.Range("A1").PasteSpecial xlPasteValues
    
    On Error GoTo Err
    
    'Now start w/ row 2 (row 1 is column labels)
    'First get width of rows
    Dim columns As Integer
    Dim rows As Integer
    Dim r As Integer
    Dim splitCell As Range
    Dim splitCellValue As String
    Dim splitValues() As String
    Dim element As Variant
    Dim first As Boolean
    columns = wksheetTarget.UsedRange.columns.Count
    rows = wksheetTarget.UsedRange.rows.Count
    
    'Set active cell to 2nd row.
    wksheetTarget.Activate
    wksheetTarget.Range("A2").Select
        
    r = 2 'Current Row
    Do While r <= (rows) 'Subtract off the column header and loop through rows
        Set splitCell = wksheetTarget.Range(columnToSplit & CStr(r)) 'Build what cell we're on
    
        splitCellValue = CStr(splitCell.Value)
        
        'If we have a , embedded then lets split it, if not we move on to next row
        If InStr(1, splitCellValue, ",") > 0 Then
            splitValues = Split(splitCellValue, ",")
            first = True
            For Each element In splitValues
                If Trim(CStr(element)) <> "" Then 'Element must contain a value to be valid.
                    If first = False Then 'On first row skip and just assign first element, for other rows, copy
                        InsertRowBelow wksheetTarget, r
                        r = r + 1
                        rows = rows + 1
                        Set splitCell = wksheetTarget.Range(columnToSplit & CStr(r)) 'Build what cell we're on
                        splitCell.Value = element
                    Else
                        first = False
                        Set splitCell = wksheetTarget.Range(columnToSplit & CStr(r)) 'Build what cell we're on
                        splitCell.Value = element
                    End If
                End If
            Next
            
        End If
        r = r + 1
    Loop
    
Err:
    If Err.Number > 0 Then
        Call MsgBox("Error:" & Err.Description & " " & CStr(Err.Number), vbCritical, "Error")
    End If
    
    Set wksheetTarget = Nothing
    Set wksheetSource = Nothing
    Set splitCell = Nothing
End Sub

'Inserts a copy of the row below the current one
Public Sub InsertRowBelow(Worksheet As Worksheet, row As Integer)
    Worksheet.Range("A" & CStr(row)).Select
    ActiveCell.Offset(1).EntireRow.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.EntireRow.Copy
    ActiveCell.Offset(1).EntireRow.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

Public Function GetSQLValue(SQLQuery As String, Optional showError As Boolean = False)
    Dim retVal As Variant
    Application.Volatile
    On Error GoTo Err
    
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim cnnStr As String
    
    cnnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Application.ActiveWorkbook.Path & "\" & Application.ActiveWorkbook.Name & "';" & _
             "Extended Properties=""Excel 12.0;ReadOnly=true;HDR=YES;IMEX=1;"";"
    
    cnn.Open cnnStr

    rs.Open SQLQuery, cnn, adOpenStatic, adLockReadOnly, adCmdText
    
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        If rs.Fields.Count > 1 Then
            retVal = CVErr(xlErrNA)
        Else
            retVal = rs.Fields(0).Value
        End If
    
    End If
    
    Set cnn = Nothing
    Set rs = Nothing

    GetSQLValue = retVal
    Exit Function
Err:
    
    If showError Then
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    End If
    GetSQLValue = CVErr(xlErrNA)
End Function

'''Requires Microsoft Scripting Runtime
Public Function CountDistinctValues(cells As Range, Optional excludeHidden As Boolean = True, Optional showError As Boolean = False)
    Dim retVal As Variant
    Application.Volatile
    On Error GoTo Err
    Dim map As New Scripting.Dictionary
    Dim selectedRange As Range
    Dim cell As Range
    
    If excludeHidden Then
        Set selectedRange = cells.SpecialCells(xlCellTypeVisible)
    Else
        Set selectedRange = cells
    End If
    
    retVal = 0
    
    For Each cell In selectedRange
        If (Not excludeHidden) Or (Not cell.EntireRow.Hidden) Then
            If Not map.Exists(cell.Value) Then
                map.Add cell.Value, 1
            Else
                map.Item(cell.Value) = map.Item(cell.Value) + 1
            End If
        End If

    Next cell
    
    retVal = CLng(map.Count)
    
    Set map = Nothing
    Set selectedRange = Nothing
    Set cell = Nothing
    
    CountDistinctValues = retVal
    Exit Function
Err:
    Set map = Nothing
    Set selectedRange = Nothing
    Set cell = Nothing
    
    If showError Then
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    End If
    CountDistinctValues = CVErr(xlErrNA)
End Function

'''Requires Microsoft Scripting Runtime
Public Function CountFrequency(cells As Range, searchValue As Variant, Optional excludeHidden As Boolean = True, Optional showError As Boolean = False)
    Dim retVal As Variant
    Application.Volatile
    On Error GoTo Err
    Dim map As New Scripting.Dictionary
    Dim selectedRange As Range
    Dim cell As Range
    
    If excludeHidden Then
        Set selectedRange = cells.SpecialCells(xlCellTypeVisible)
    Else
        Set selectedRange = cells
    End If
    
    retVal = 0
    
    For Each cell In selectedRange
        If (Not excludeHidden) Or (Not cell.EntireRow.Hidden) Then
            If cell.Value = searchValue Then
                retVal = retVal + 1
            End If
        End If

    Next cell
    
    If map.Exists(searchValue) Then
        retVal = CLng(map.Item(searchValue))
    End If
    
    Set map = Nothing
    Set selectedRange = Nothing
    Set cell = Nothing
    
    CountFrequency = retVal
    Exit Function
Err:
    Set map = Nothing
    Set selectedRange = Nothing
    Set cell = Nothing
    
    If showError Then
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    End If
    CountFrequency = CVErr(xlErrNA)
End Function

'''Requires Microsoft Scripting Runtime
Public Function SimpleLookup(sourceCells As Range, searchColumn As String, searchValue As Variant, returnColumn As String, Optional excludeHidden As Boolean = True, Optional showError As Boolean = False)
    Dim retVal As Variant
    Application.Volatile
    On Error GoTo Err
    Dim selectedRange As Range
    Dim cell As Range
    Dim row As Range
    
    If excludeHidden Then
        Set selectedRange = sourceCells.SpecialCells(xlCellTypeVisible)
    Else
        Set selectedRange = sourceCells
    End If
    
    retVal = 0
    
    For Each row In selectedRange.rows
        If (Not excludeHidden) Or (Not row.EntireRow.Hidden) Then
            'check if searchColumn is within range.
            Set cell = row.columns(searchColumn)
            If Not cell Is Nothing Then
                If cell.Value = searchValue Then
                    Set cell = row.columns(returnColumn)
                    SimpleLookup = cell.Value
                    Exit Function
                End If
            End If
            
            
        End If

    Next row
    
    Set selectedRange = Nothing
    Set cell = Nothing
    
    SimpleLookup = Empty
    Exit Function
Err:
    Set selectedRange = Nothing
    Set cell = Nothing
    
    If showError Then
        MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    End If
    SimpleLookup = CVErr(xlErrNA)
End Function


'Inserts a blank column to the left of the column specified.
Public Sub InsertCellsToLeft(Worksheet As Worksheet, column As String, Optional NewTitle As String = "")
    Worksheet.Range(column & ":" & column).Select
    ActiveCell.Offset(1).EntireColumn.Insert Shift:=xlRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(1).cells(0).Value = NewTitle

    Application.CutCopyMode = False
End Sub

'Sets a formula for each cell.
Public Sub SetFormula(Worksheet As Worksheet, column As String, formula As String)
    Dim activeRange As Range
    Dim cell As Range
    Dim row As Range
    Dim targetFormula As String
    
    
    Set activeRange = Worksheet.UsedRange.Offset(1)
    
    For Each row In activeRange.rows
        Set cell = row.columns(column)
        'Sample formula, {R} is a row placeholder and {C} is a column placeholder
        '=SimpleLookup(SQLLookupSource!A:C, "A", SimpleLookupExample!A{R}, "B")
        targetFormula = Replace(formula, "{R}", row.row)
        targetFormula = Replace(targetFormula, "{C}", row.row)
        
        cell.formula = targetFormula
        
        Set cell = Nothing
    Next row
    
    
    
    Set activeRange = Nothing

    Application.CutCopyMode = False
End Sub



