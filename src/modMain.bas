Attribute VB_Name = "modMain"
Option Explicit

' Our one and only VibrationVIEW
' Use "early binding" (New), VibrationVIEW will open as soon as first accessed
Global Vibview As New VibrationVIEWLib.VibrationVIEW
Global VibviewTransient As New VibrationVIEWLib.TransientControl
' Array defined globally for convieniance in passing data
Global arry() As Single

