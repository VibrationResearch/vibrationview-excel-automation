Attribute VB_Name = "modTEDS"
Option Explicit

Private Sub cmdUpdateTEDS1()
cmdUpdateTEDS 0
End Sub
Private Sub cmdUpdateTEDS2()
cmdUpdateTEDS 1
End Sub
Private Sub cmdUpdateTEDS3()
cmdUpdateTEDS 2
End Sub
Private Sub cmdUpdateTEDS4()
cmdUpdateTEDS 3
End Sub


Private Sub cmdUpdateTEDS(channel As Integer)
    On Error Resume Next
    ' to get all the teds fields you need a 3 column array, 100 rows is just a guess which should exceed any teds report
    Dim array1(1 To 100, 1 To 3) As String
    
    ActiveSheet.Range("A1:C103").Clear
    ActiveSheet.Cells(1, 1) = "Channel"
    ActiveSheet.Cells(1, 2) = channel + 1 ' channel number starts with 0
    ' request channel 1 teds data
    Vibview.Teds channel, array1
    If Err Then ' no data .. clear old data out
        ActiveSheet.Range("A2:C103").Clear
    Else  ' paste new data into the worksheet
        ActiveSheet.Range("A2:C103") = array1
    End If
End Sub
