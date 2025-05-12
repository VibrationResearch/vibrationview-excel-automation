# VibrationVIEW Excel Integration

## Overview

This Excel workbook provides a comprehensive interface to the VibrationVIEW automation API, allowing you to control vibration tests, manage test profiles, collect data, and monitor system status directly from Microsoft Excel. This integration enables advanced automation capabilities, data collection, and custom reporting for vibration testing applications. It may be used as distributed or as an example to build your own interface to VibrationVIEW.

## Requirements

- Microsoft Excel
- VibrationVIEW software installed
- VibrationVIEW automation option (VR9604) - OR - VibrationVIEW may be run in Simulation mode without any additional hardware or software

## Modules Overview

The workbook contains several VBA modules, each providing different functionality:

### modMain.bas

The main module that establishes global references to the VibrationVIEW application. It provides:

- Global declaration of the VibrationVIEW and TransientControl objects
- Early binding to ensure VibrationVIEW opens when first accessed

```vba
Global Vibview As New VibrationVIEWLib.VibrationVIEW
Global VibviewTransient As New VibrationVIEWLib.TransientControl
Global arry() As Single
```

### modControl.bas

Provides core test control functions:

- Running tests from different profile types (Sine, Random, Shock, Data Replay)
- Starting, stopping, and resuming tests
- Editing test profiles
- Saving test data
- Reading channel information, demand values, control values, and system status

### modData.bas

Handles data acquisition and display:

- Retrieving vector data for time, frequency, and waveform displays
- Reading rear input values and labels
- Populating worksheet ranges with vibration data
- Setting up charts with dynamic data

### modParams.bas

Provides access to VibrationVIEW report fields:

- Retrieves parameter values from VibrationVIEW
- Updates worksheets with current test parameters

### modRandomControl.bas

Specialized controls for random vibration tests:

- Setting modifiers for random tests
- Handling schedule levels

### modSineControl.bas

Specialized controls for sine vibration tests:

- Setting and getting amplitude multipliers
- Setting and getting frequency values
- Controlling sine sweep operations (up, down, hold, step)
- Resonance hold functionality

### modResizeChart.bas

Utilities for dynamically resizing charts based on data:

- Adjusting chart series lengths
- Parsing chart formulas
- Helper functions for column conversion

### modTEDS.bas

Handles Transducer Electronic Data Sheet (TEDS) information:

- Reading TEDS data for connected sensors
- Displaying TEDS information in worksheets

### Create a reference to the VibrationVIEW application

VBA menu Tools .. References
add a reference to VibrationVIEW Type Library (C:\Program Files\VibrationVIEW <version>\VibrationVIEW.tlb)

in a module include:
Global Vibview As New VibrationVIEWLib.VibrationVIEW

A status is provided vibview.IsReady - no commands are accepted until this status==1

## Usage Examples

### Opening and Running a Test Profile 

````vba
' Open and run a test profile
Dim testProfile As String
testProfile = Application.GetOpenFilename("Sine Profiles (*.vsp), *.vsp,Random Profiles (*.vrp), *.vrp")

If Len(testProfile) > 0 And Len(Dir(testProfile)) > 0 Then
    Vibview.RunTest testProfile
End If

### Reading Channel Data

```vba
' Read channel data
Dim channelValues(3) As Single

Vibview.Channel channelValues
ActiveSheet.Range("D18:G18") = channelValues()
````

### Controlling a Sine Sweep

```vba
' Set sweep parameters
Vibview.SweepMultiplier = Range("SweepMultiplier")
Vibview.SineFrequency = Range("SineFrequency")

' Start sweep
Vibview.SweepUp

' Hold at current frequency
Vibview.SweepHold

' Continue sweep
Vibview.SweepUp

' Step through frequencies
Vibview.SweepStepUp
```

### Getting Test Status

```vba
Dim testStatus As String
Dim stopCodeIndex As Long

Vibview.Status testStatus, stopCodeIndex
ActiveSheet.Range("D24") = testStatus
ActiveSheet.Range("E24") = stopCodeIndex
```

## Chart Data Integration

The workbook includes functionality to dynamically update Excel charts with data from VibrationVIEW:

1. Time-domain data charts
2. Frequency-domain data charts
3. Historical data charts

The `SetChartDataLength` function dynamically resizes chart series to accommodate changing data.

## TEDS Sensor Information

The workbook can read and display TEDS sensor information:

```vba
' Read TEDS data from channel 0
Dim tedsData(1 To 100, 1 To 3) As String
Vibview.Teds 0, tedsData
ActiveSheet.Range("A2:C103") = tedsData
```

## Notes

- Error handling is implemented in most functions to handle common issues
- The workbook uses early binding for improved performance and IntelliSense support
- Multiple test types are supported (Sine, Random, Shock, Data Replay)
- Data can be saved to external files for further analysis

## Troubleshooting

- If VibrationVIEW fails to open, verify the installation and registration
- Ensure your hardware is properly connected and recognized by VibrationVIEW before attempting to control tests -OR- Use simulation mode
- Check stopcode messages in the VibrationVIEW application if commands fail to execute

## License

This Excel integration is provided as an example and can be modified to suit your specific testing needs.
