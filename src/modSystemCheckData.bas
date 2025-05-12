Attribute VB_Name = "modSystemCheckData"
Option Explicit

Private Sub cmdVector_SystemCheck()
' GetData is a macro located in modMain
    GetData VV_WAVEFORMAXIS, ActiveSheet.Range("a2..e2")
    GetData VV_FREQUENCYAXIS, ActiveSheet.Range("f2..j2")
End Sub

