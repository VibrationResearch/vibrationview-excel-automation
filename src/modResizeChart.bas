Attribute VB_Name = "modResizeChart"
Option Explicit

Sub SetChartDataLength(chartObj As Chart, newLength As Integer, Optional startRow As Integer = 1)
    Dim datasheet As Worksheet
    Dim seriesObj As Series
    Dim xValues As Variant, yValues As Variant
    Dim xColumn As Long, yColumn As Long
    Dim xSheetName As String, ySheetName As String
    
    ' Loop through each series in the chart
    For Each seriesObj In chartObj.SeriesCollection
        ' Get formula components - this is more reliable than XValues/Values properties
        ParseSeriesFormula seriesObj.formula, xSheetName, xColumn, ySheetName, yColumn
        
        ' If we have valid column information
        If yColumn > 0 Then
            ' Get the worksheet containing data
            On Error Resume Next
            Set datasheet = Worksheets(ySheetName)
            On Error GoTo 0
            
            ' Make sure we have a valid worksheet reference
            If Not datasheet Is Nothing Then
                ' Update Y values
                seriesObj.Values = datasheet.Range( _
                    datasheet.Cells(startRow, yColumn), _
                    datasheet.Cells(startRow + newLength - 1, yColumn))
                
                ' Update X values if we have them
                If xColumn > 0 Then
                    Set datasheet = Worksheets(xSheetName)
                    seriesObj.xValues = datasheet.Range( _
                        datasheet.Cells(startRow, xColumn), _
                        datasheet.Cells(startRow + newLength - 1, xColumn))
                End If
            End If
        End If
    Next seriesObj
End Sub

' Helper function to parse SERIES formula
Sub ParseSeriesFormula(formula As String, ByRef xSheetName As String, ByRef xColumn As Long, _
                       ByRef ySheetName As String, ByRef yColumn As Long)
    ' Initialize
    xSheetName = ""
    xColumn = 0
    ySheetName = ""
    yColumn = 0
    
    ' Debug
    Debug.Print "Formula: " & formula
    
    ' SERIES formula format: =SERIES(name,xvalues,yvalues,order)
    Dim parts As Variant
    Dim xPart As String, yPart As String
    
    On Error Resume Next
    
    ' Remove =SERIES( from beginning and ) from end
    formula = Mid(formula, 9, Len(formula) - 9)
    
    ' Split by commas not inside quotes
    parts = Split(formula, ",")
    
    ' We need at least 4 parts
    If UBound(parts) < 3 Then Exit Sub
    
    ' Get X and Y parts
    xPart = parts(1)
    yPart = parts(2)
    
    ' Parse X part if not empty
    If xPart <> "" And xPart <> "1" Then
        ParseRangePart xPart, xSheetName, xColumn
    End If
    
    ' Parse Y part
    ParseRangePart yPart, ySheetName, yColumn
    
    ' Debug
    Debug.Print "X Sheet: " & xSheetName & ", X Column: " & xColumn
    Debug.Print "Y Sheet: " & ySheetName & ", Y Column: " & yColumn
    
    On Error GoTo 0
End Sub

' Helper function to parse range part of formula
Sub ParseRangePart(rangePart As String, ByRef sheetName As String, ByRef columnNum As Long)
    Dim r As Range
    
    ' Remove quotes if present
    If Left(rangePart, 1) = """" Then
        rangePart = Mid(rangePart, 2, Len(rangePart) - 2)
    End If
    
    ' Try to convert to range
    On Error Resume Next
    Set r = Range(rangePart)
    
    If Not r Is Nothing Then
        sheetName = r.Worksheet.Name
        columnNum = r.Column
    Else
        ' Try to extract sheet name and range manually
        Dim bangPos As Integer
        bangPos = InStr(rangePart, "!")
        
        If bangPos > 0 Then
            sheetName = Left(rangePart, bangPos - 1)
            ' Remove sheet name brackets if present
            If Left(sheetName, 1) = "'" Then
                sheetName = Mid(sheetName, 2, Len(sheetName) - 2)
            End If
            
            ' Try to determine column from remaining part
            Dim rangePortion As String
            rangePortion = Mid(rangePart, bangPos + 1)
            
            ' This is simplified - would need more robust parsing for complex ranges
            ' Just getting first column reference
            Dim colLetter As String
            colLetter = ""
            
            Dim i As Integer
            For i = 1 To Len(rangePortion)
                If IsLetter(Mid(rangePortion, i, 1)) Then
                    colLetter = colLetter & Mid(rangePortion, i, 1)
                Else
                    Exit For
                End If
            Next i
            
            ' Convert column letter to number
            If colLetter <> "" Then
                columnNum = ColumnLetterToNumber(colLetter)
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub

' Helper function to check if character is letter
Function IsLetter(c As String) As Boolean
    IsLetter = (UCase(c) >= "A" And UCase(c) <= "Z")
End Function

' Helper function to convert column letter to number
Function ColumnLetterToNumber(colLetter As String) As Long
    Dim result As Long
    Dim i As Integer, c As String
    
    result = 0
    For i = 1 To Len(colLetter)
        c = Mid(colLetter, i, 1)
        result = result * 26 + (Asc(UCase(c)) - Asc("A") + 1)
    Next i
    
    ColumnLetterToNumber = result
End Function

