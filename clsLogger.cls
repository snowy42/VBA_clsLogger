Option Explicit

'==============================================================================
' Logging Class Module
'==============================================================================
' Description: A class module designed to provide a log file output with simple
'              event adding.
' Author: Matthew Snow
' Version: 1.0
' Last Modified: 15 Feb 2024
'==============================================================================


'==============================================================================
' Class Variables
'==============================================================================
Private mLogText As String
Private mLogFilePath As String
Private mLogFileName As String
Private mTimer As Double
Private fNow As String


'==============================================================================
' Public Properties
'==============================================================================
Public Property Get logText() As String
    ' get the current log text after events have been added
    '   (useful for debugging)
    logText = mLogText
End Property


'==============================================================================
' Public Methods
'==============================================================================
Public Sub SetFileTitle(ByVal vTitle As String, Optional ByVal vIncludeTime As Boolean = True)
    '==========================================================================
    ' USAGE:
    ' SetFileTitle("New File Name")
    '   = adjusts the output file name to "New File Name_*timestamp*"
    ' SetFileTitle("FileNameNoTimestamp", false)
    '   = adjusts the output file name to "FileNameNoTimestamp" and
    '     omits the timestamp
    '==========================================================================
    ' sets the name of the file.  It will always export to the workbook path
    ' optional variable to include a timestamp on the log file name (preferable
    ' so thatlog files aren't overwritten)
    
    mLogFileName = vTitle & "_Log" & IIf(vIncludeTime, "_" & fNow, "") & ".txt"
    mLogFilePath = LocalPath(ThisWorkbook.path) & "\" & mLogFileName
End Sub

Public Sub AddEvent(eventText As String)
    '==========================================================================
    ' USAGE:
    ' SetFileTitle("Successfully opened word document")
    '   = will add an event to the log like this:
    '      123.4 ms            Successfully opened word document
    '==========================================================================
    ' add an event to the log to be included in the output file
    ' includes time (in ms) since the class was initialized
    
    Dim fTimer As String
    fTimer = Format((Timer - mTimer) * 1000, "0.0") & " ms"
    
    Dim tabSpace As Integer
    tabSpace = 20
    mLogText = mLogText & fTimer
    Dim i As Integer
    For i = Len(fTimer) To 20
        mLogText = mLogText & " "
    Next i
    
    mLogText = mLogText & eventText & vbCrLf
End Sub

Public Sub CommitLog()
    '==========================================================================
    ' USAGE:
    ' CommitLog()
    '==========================================================================
    ' finalises the log object and creates a text file in the workbook path with
    ' all the log data including a header and footer
    
    mLogText = mLogFileName & vbCrLf & mLogText & "--------------------------------------------------" & vbCrLf & _
                "Log complete, process took:" & "   " & Format((Timer - mTimer), "0.0") & " seconds" & vbCrLf & _
                "Log file saved to: " & mLogFilePath
    On Error Resume Next
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open mLogFilePath For Append As fileNumber
    Print #fileNumber, mLogText
    Close fileNumber
    On Error GoTo 0
    Debug.Print "Log file saved to: " & mLogFilePath & vbCrLf
End Sub


'==============================================================================
' Private Methods
'==============================================================================
Private Sub Class_Initialize()
    ' Initializes the class, starts a timer and sets the default filename
    fNow = Format(Now, "yyyyMMdd_HHmmss")
    
    Dim fNowLong As String
    fNowLong = Format(Now, "MMMM dd, yyyy - HH:mm:ss")
    
    mLogFileName = "Log_" & fNow & ".txt"
    mLogFilePath = LocalPath(ThisWorkbook.path) & "\" & mLogFileName
    mLogText = "Start: " & fNowLong & vbCrLf & _
                "__________________________________________________" & vbCrLf
    mTimer = Timer
End Sub

Private Function LocalPath(datPath As String) As String
    ' Allows use to include MS OneDrive locations
    Dim oneDrivePart As String
    datPath = Replace(datPath, "/", "\")
    oneDrivePart = "https:\\d.docs.live.net\"
    If InStr(datPath, oneDrivePart) Then
        datPath = Replace(datPath, oneDrivePart, "")
        datPath = Right(datPath, Len(datPath) - InStr(1, datPath, "\"))
        datPath = Environ$("OneDriveConsumer") & "\" & datPath
    End If
    LocalPath = datPath
End Function
