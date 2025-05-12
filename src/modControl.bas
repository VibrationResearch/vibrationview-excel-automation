Attribute VB_Name = "modControl"
Option Explicit

Private sData As String
Private sTest As String

Private Sub cmdRunTest_Click()
Dim sTest As String

sTest = Application.GetOpenFilename("Sine Profiles (*.vsp), *.vsp,Random Profiles (*.vrp), *.vrp,Shock Profiles (*.vkp), *.vkp,Data Replay Profiles (*.vfp), *.vfp")

If Len(sTest) > 0 And Len(Dir(sTest)) > 0 Then
    Application.Cursor = xlWait
    Vibview.RunTest sTest
    Application.Cursor = xlDefault
End If
End Sub
Private Sub cmdStart_Click()
    Vibview.StartTest
End Sub
Private Sub cmdStop_Click()
    Vibview.StopTest
End Sub
Private Sub cmdResume_Click()
    If Vibview.CanResumeTest() Then
        Vibview.ResumeTest
    End If
End Sub

Private Sub cmdEdit_click()
On Error Resume Next
sTest = Application.GetOpenFilename("Sine Profiles (*.vsp), *.vsp,Random Profiles (*.vrp), *.vrp,Shock Profiles (*.vkp), *.vkp,Data Replay Profiles (*.vfp), *.vfp")

If Len(sTest) > 0 And Len(Dir(sTest)) > 0 Then
    Application.Cursor = xlWait
    Vibview.EditTest sTest
    Application.Cursor = xlDefault
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbOKOnly, Err.Description
    End If
End If
End Sub

Private Sub cmdSave_Click()

sData = Application.GetSaveAsFilename("", "Random Data (*.vrd), *.vrd")

If Len(sData) > 0 Then
    Vibview.SaveData sData
End If
End Sub
Private Sub cmdReadChannels_Click()
Dim arry(3) As Single
Dim sLabel As String
Dim lp As Integer
    For lp = 0 To 3
        ActiveSheet.Cells(16, lp + 4) = Vibview.ChannelLabel(lp)
        ActiveSheet.Cells(17, lp + 4) = Vibview.ChannelUnit(lp)
    Next
    Vibview.channel arry
    ActiveSheet.Range("D18:G18") = arry()
End Sub

Private Sub cmdReadDemand_Click()
Dim arry(0) As Single

    Vibview.Demand arry
    ActiveSheet.Range("D21:D21") = arry()
 
End Sub
Private Sub cmdReadControl_Click()
Dim arry(0) As Single

    Vibview.Control arry
    ActiveSheet.Range("D22:D22") = arry()

End Sub
Private Sub cmdReadStatus_Click()
Dim sStatus As String
Dim nStopCodeIndex As Long
    Vibview.Status sStatus, nStopCodeIndex
    ActiveSheet.Range("D24") = sStatus
    ActiveSheet.Range("e24") = nStopCodeIndex
   
End Sub


