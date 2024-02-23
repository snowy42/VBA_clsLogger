# Logging Class Module

## Description

The Logging Class Module is designed to provide a log file output with simple event adding functionality. It allows you to easily log events with timestamps and export them to a text file for tracking and debugging purposes.

## Author

Matthew Snow

## Version

1.0

## Last Modified

15 Feb 2024

## Class Variables

- `mLogText`: String variable to store the log text.
- `mLogFilePath`: String variable representing the full path to the log file.
- `mLogFileName`: String variable representing the log file name.
- `mTimer`: Double variable for measuring time intervals.
- `fNow`: String variable storing the current date and time.

## Public Properties

### `logText`

- **Description**: Get the current log text after events have been added.
- **Usage**: 
  ```vba
  `logText = mLogText`
  ```

## Public Methods
### `SetFileTitle`
- **Description**: Set the name of the log file and optionally include a timestamp.
- **Usage**:
  ```vba
  SetFileTitle("New File Name")
  SetFileTitle("FileNameNoTimestamp", False)
  ```
### `AddEvent`
- **Description**: Add an event to the log with a timestamp.
- **Usage**:
  ```vba
  AddEvent("Successfully opened word document")
  ```
### `CommitLog`
- **Description**: Finalize the log object, create a text file, and save all log data.
- **Usage**:
  ```vba
  CommitLog()
  ```

## Private Methods
### `Class_Initialize`
- **Description**: Initialize the class, start a timer, and set the default filename.
### `LocalPath`
- **Description**: Convert the path to a local path, handling MS OneDrive locations.

## Usage Example
```vba
  ' Create a new instance of the Logging Class
Dim log As New LoggingClass

' Set the log file title with a timestamp
log.SetFileTitle "LogFile"

' Add events to the log
log.AddEvent "Event 1"
log.AddEvent "Event 2"
log.AddEvent "Event 3"

' Commit the log to a text file
log.CommitLog
```

## Output Example
```
LogFile_Log_20240223_150121.txt
Start: February 23, 2024 - 15:01:21
__________________________________________________
3.9 ms               Event 1
1160.2 ms            Event 2
2707.0 ms            Event 3
--------------------------------------------------
Log complete, process took:   4.2 seconds
Log file saved to: C:\Users\Snowy\OneDrive\Documents\_Upwork\_Common Files\LogFile_Log_20240223_150121.txt
```

## Notes
- The log file is saved to the workbook path by default.
- Timestamps include milliseconds for accurate time tracking.
