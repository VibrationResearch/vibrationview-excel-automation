Attribute VB_Name = "modData"
Option Explicit

Private Sub cmdVector()
    Dim vectorlength As Integer
    Dim datasheet As Worksheet
    Set datasheet = Worksheets("Chart Data")
    vectorlength = GetData(VV_WAVEFORMAXIS, datasheet.Range("A2..E2"))
    SetChartDataLength Charts("Chart Time"), vectorlength, 2
    
    vectorlength = GetData(VV_FREQUENCYAXIS, datasheet.Range("H2..L2"))
    SetChartDataLength Charts("Chart freq"), vectorlength, 2
    
    vectorlength = GetData(VV_TIMEHISTORYAXIS, datasheet.Range("O2..S2"))
    SetChartDataLength Charts("Chart History"), vectorlength, 2
    
    GetData VV_WAVEFORMDEMAND, datasheet.Range("F2..G2")
    GetData VV_FREQUENCYDEMAND, datasheet.Range("M2..N2")
    GetData VV_REARINPUTHISTORY1, datasheet.Range("P2..S2")
    
End Sub

Private Sub cmdRearInputs()
On Error Resume Next
Dim nLen As Integer
Dim vctr As VibrationVIEWLib.vvVector
Dim arry(7) As Single
Dim sLabel As String
Dim lp As Integer
    For lp = 0 To 7
        ActiveSheet.Cells(1, lp + 2) = Vibview.RearInputLabel(lp)
        ActiveSheet.Cells(2, lp + 2) = Vibview.RearInputUnit(lp)
    Next
    Vibview.RearInput arry
    ActiveSheet.Range("B3:I3") = arry()

End Sub

Public Function GetData(Vector As VibrationVIEWLib.vvVector, Target As Range)
    Dim nLen As Integer
    Dim nVectors As Integer
    ' Ask VibrationVIEW how long the dataset is
    nLen = Vibview.vectorlength(Vector)
    
    ' Determine from target range how many columns we are requesting
    nVectors = Target.Columns.Count
    
    ' Clear all rows in the target columns except the first row
    Target.Worksheet.Range(Target.Worksheet.Cells(2, Target.Column), _
                           Target.Worksheet.Cells(Target.Worksheet.Rows.Count, Target.Column + nVectors - 1)).Clear
    
    If nLen > 0 Then
        ' Dimension our global array to accommodate the data
        ReDim arry(nLen, nVectors)
    
        ' Get the data from VibrationVIEW
        Vibview.Vector arry, Vector
    
        ' Paste the data into the target range after resizing the range to fit
        Target.Rows.Resize(nLen, nVectors) = arry
    End If
    
    GetData = nLen
    
End Function

