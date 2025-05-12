Attribute VB_Name = "modRandomControl"
Option Explicit

Private Sub cmdSetModifier()
    Dim comboIndex As Long
    comboIndex = Range("RandomModifierType").Value
    
    ' Get the actual text value from the list
    Dim textValue As String
    textValue = Range("RANDOMLIST").Cells(comboIndex, 1).Value
    
    ' Convert the text value to a BYTE (using ASCII value)
    Dim actualmodifier As Byte
    actualmodifier = Asc(textValue)
    
    Vibview.ActiveTest.ActiveScheduleLevel.SetScheduleLevel Range("RandomTime").Value, Range("RandomModifier").Value, actualmodifier, Range("RandomLoop").Value
End Sub

