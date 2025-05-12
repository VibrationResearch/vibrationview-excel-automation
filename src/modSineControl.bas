Attribute VB_Name = "modSineControl"
Option Explicit

Private Sub cmdAmplitudeMultiplier()
    Vibview.DemandMultipler = Range("AmplitudeMultiplier")
End Sub

Private Sub cmdGetAmplitude()
    Range("AmplitudeMultiplier") = Vibview.DemandMultipler
End Sub
Private Sub cmdFreqMult()
    Vibview.SweepMultiplier = Range("SweepMultiplier")
End Sub


Private Sub cmdGetSweepRate()
    Range("SweepMultiplier") = Vibview.SweepMultiplier
End Sub

Private Sub cmdResonanceHold()
    Vibview.SweepResonanceHold
End Sub
Private Sub cmdGetFreq()
    Range("SineFrequency") = Vibview.SineFrequency
End Sub

Private Sub cmdSetFreq()
    Vibview.SineFrequency = Range("SineFrequency")
End Sub

Private Sub cmdSweepDown()
    Vibview.SweepDown
End Sub

Private Sub cmdSweepHold()
    Vibview.SweepHold
End Sub

Private Sub cmdSweepStepDown()
    Vibview.SweepStepDown
End Sub

Private Sub cmdSweepStepUp()
    Vibview.SweepStepUp
End Sub

Private Sub cmdSweepUp()
    Vibview.SweepUp
End Sub




