Attribute VB_Name = "modParams"
Option Explicit

Private Sub cmdUpdateParams()
Dim lp As Integer
Dim request As String

Application.Cursor = xlWait

For lp = 1 To 300 'max number of fields in the worksheet
    request = ActiveSheet.Cells(lp, 2) ' column 2 has list of field names
    If Len(request) > 0 Then
       'ask VibrationVIEW to interpret the report field, put result in column 3
        ActiveSheet.Cells(lp, 3) = Vibview.ReportField(request)
    End If
Next
Application.Cursor = xlDefault
       
End Sub

