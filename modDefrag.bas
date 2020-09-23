Attribute VB_Name = "modDefrag"
' ***************************************************************************
' Routine:   modDefrag
'
' Purpose:   Perform a defrag of one or more drives using two external
'            Microsoft utilities written by Mark Russinovichand.  These
'            are freeware and can be downloaded from the Microsoft
'            Sysinternals web site .
'
' Warning:   Read the documentation on the web site prior to using these
'            utilities.  They may not be compatible with non NT based
'            operating systems.
'
'            Contig.exe v1.60  dtd 01-Feb-2011
'            Defrags files only.
'            http://www.microsoft.com/technet/sysinternals/FileAndDisk/Contig.mspx
'
'            PageDfrg.exe v2.32  dtd 1-Nov-2006
'            Defrags windows page file and registry hives.  PageDefrag works
'            on 32-bit Windows NT 4.0, Windows 2000, Windows XP, Server 2003.
'            Pagedefrag does not work on Windows Vista or newer (32-bit) and
'            not on any version of 64-bit Windows.
'            http://www.microsoft.com/technet/sysinternals/FileAndDisk/PageDefrag.mspx
'            http://en.wikipedia.org/wiki/PageDefrag
'
' Thanks:    To David Sherlock whose suggestions concerning this project is
'            greatly appreciated.
'
'            To Ruturaaj for his suggestions for capturing Contig data.
'
' Good information from Alfred Hellmüller:
'
'     Disk Defragmenters as well as Wipers work with numerous
'     repeating write operations. Flash memories like USB Sticks or
'     SD Cards do have a limited life of about 10,000 to 10e6
'     (high reliability) erase and write operations.  Even read
'     operations are degrading the data quality response. Increasing
'     Bit failures.
'
'     Some areas of a Flash are extremely exposed such as those
'     that holds index structures of the file system. The 10,000
'     operations can be reached in a short time because any write
'     operation includes a delete operation in advance.
'
'     Conclusion:  We should avoid any Write operations on USB sticks
'     whenever possible.  Both, Defragmenter and Wiper applications
'     are high grade toxic Procedures for Flash memories.
'
'     References:
'        How To Use the Remote Shutdown Tool to Shut Down and Restart
'        a Computer in Windows 2000
'        http://support.microsoft.com/kb/317371
'
'        Storage Search.com
'        http://www.storagesearch.com/reliability.html
'
'        Maximizing Performance and Reliability in Flash Memory Devices
'        by Randy Martin, QNX Software Systems
'        Ecnmag.com - March 01, 2006
'        http://www.ecnmag.com/maximizing-performance-and-reliability.aspx?meid=580
'
'        Google:   flash memories write operations limit
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Created module
'              Added file size check for files less than 2 bytes.
'              Added log file reference
' 02-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Fixed a bug where the system detected a file but would
'              indicate that it did not exist becuase it was being held by
'              an internal process.
' 07-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated above documentation concerning Flash Memory devices.
' 25-Oct-2009  Kenneth Ives  kenaso@tx.rr.com
'              Added a pause processing  to the defrag process.
'              Added byte and file count to display.
'              Added SetTimer() and KillTimer() APIs to control elapsed time.
' 23-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Added API call if Contig takes longer to process a file
'              - Added option to log just path\name of file
' 07-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Deactivated all updated locations on form prior to rebooting
'              - Removed reference to MS Shutdown tool.  Too slow.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Constants
' ***************************************************************************
  Private Const MODULE_NAME          As String = "modDefrag"
  Private Const KB_1                 As Long = &H400&   ' 1024
  Private Const INVALID_HANDLE_VALUE As Long = -1

' ***************************************************************************
' Type Structures
' ***************************************************************************
  ' The FILETIME structure is a 64-bit value representing the number of
  ' 100-nanosecond intervals since January 1, 1601.
  Private Type FILETIME
       dwLowDateTime  As Long
       dwHighDateTime As Long
  End Type

  ' The WIN32_FIND_DATA structure describes a file found by the FindFirstFile
  ' or FindNextFile function. If a file has a long filename, the complete
  ' name appears in the cFileName field, and the 8.3 format truncated version
  ' of the name appears in the cAlternate field. Otherwise, cAlternate is empty.
  Private Type WIN32_FIND_DATA
       dwFileAttributes As Long          ' file attributes
       ftCreationTime   As FILETIME      ' File creation date and time
       ftLastAccessTime As FILETIME      ' File last accessed date and time
       ftLastWriteTime  As FILETIME      ' File last modified date and time
       nFileSizeHigh    As Long          ' file sizes over 2GB (2,147,483,647)
                                         ' (nFileSizeHigh * MAXDWORD) + nFileSizeLow
       nFileSizeLow     As Long          ' file sizes under 2GB (2,147,483,647)
       dwReserved0      As Long
       dwReserved1      As Long
       cFilename        As String * 260  ' full file name w/o path
       cAlternate       As String * 14   ' short file name w/o path
  End Type
    
' ***************************************************************************
' Local API Declares
' ***************************************************************************
  ' The FindFirstFile function searches a directory for a file whose name
  ' matches the specified filename. FindFirstFile examines subdirectory names
  ' as well as filenames.  The FindFirstFile function opens a search handle
  ' and returns information about the first file whose name matches the
  ' specified pattern. Once the search handle is established, you can use the
  ' FindNextFile function to search for other files that match the same
  ' pattern. When the search handle is no longer needed, close it by using
  ' the FindClose function.  The FindFirstFile function searches for files by
  ' name only; it cannot be used for attribute-based searches.
  Private Declare Function FindFirstFile Lib "kernel32" _
          Alias "FindFirstFileA" (ByVal lpFileName As String, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' The FindNextFile function continues a file search from a previous
  ' call to the FindFirstFile function.
  Private Declare Function FindNextFile Lib "kernel32" _
          Alias "FindNextFileA" (ByVal hFindFile As Long, _
          lpFindFileData As WIN32_FIND_DATA) As Long

  ' The FindClose function closes the specified search handle. The
  ' FindFirstFile and FindNextFile functions use the search handle
  ' to locate files with names that match a given name.  Always close
  ' the file handle when finished.
  Private Declare Function FindClose Lib "kernel32" _
          (ByVal hFindFile As Long) As Long

  ' ZeroMemory is used for clearing contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Global API Declares
' ***************************************************************************
  ' The SetTimer function creates a timer with the specified time-out value.
  ' If the function succeeds and the hWnd parameter is NULL, the return value
  ' is a value identifying the new timer.
  Public Declare Function SetTimer Lib "user32" _
         (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, _
         ByVal lpTimerFunc As Long) As Long
  
  ' The KillTimer function destroys the specified timer.  If the function
  ' succeeds, the return value is nonzero.
  Public Declare Function KillTimer Lib "user32" _
         (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Currency)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m cur ByteCount
' Variable name:     mcurByteCount
'
' ***************************************************************************
  Private mlngTimerID   As Long      ' Timer ID handle
  Private mcurByteCount As Currency
  Private mcurFileCount As Currency
  Private mfrmName      As Form

' ***************************************************************************
' Routine:       BeginDefrag
'
' Description:   Select the drive or drives that are to be defragged.
'
' Parameters:    strTarget - One or more drives to be processed
'                frm - Link to the main form of this application
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-JUL-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added log file and reboot reference
' ***************************************************************************
Public Sub BeginDefrag(ByVal strTarget As String, _
                       ByVal frm As Form)

    Dim lngIndex    As Long
    Dim strDrive    As String
    Dim strRecord   As String
    Dim strFileSys  As String
    Dim objDiskInfo As cDiskInfo
    
    ' must have error handler enabled, as all
    ' disks do not return all information
    Const ROUTINE_NAME As String = "BeginDefrag"

    On Error GoTo BeginDefrag_Error

    On Local Error Resume Next
    
    Set mfrmName = frm     ' mfrmName referenced in more than one routine
    
    StartTimer   ' Start API Timer
    
    If gblnLogData Then
        BuildLogFile     ' Create log file if it does not exist
    End If
    
    mcurByteCount = 0@   ' Initialize counters
    mcurFileCount = 0@
    
    strFileSys = vbNullString                   ' Clear holding areas
    mfrmName.lblMsg(2).Caption = vbNullString
    mfrmName.lblStats(1).Caption = "0"
    mfrmName.lblStats(3).Caption = "0"
    mfrmName.lblStats(4).Caption = "Elapsed time:  0:00:00"
    
    strTarget = QualifyPath(strTarget)           ' Append backslash to target
    strDrive = QualifyPath(Left$(strTarget, 2))  ' capture drive letter
    strDrive = UCase$(strDrive)                  ' Convert to uppercase
    
    ' Capture file system used
    Set objDiskInfo = New cDiskInfo
    objDiskInfo.GetVolumeInfo strDrive, strFileSys
    Set objDiskInfo = Nothing
            
    ' Display system type in upper right corner of form
    mfrmName.lblMsg(2).Caption = "File system:  " & strFileSys
    
    DoEvents
    If gblnLogData Then
        strRecord = GetTimeStamp & Space$(2) & "STARTED:  " & "Defrag of " & strTarget & _
                    vbTab & "File system:  " & strFileSys
        UpdateLogFile strRecord
    End If

    SearchAndProcess strTarget  ' Begin defrag process
    CloseAllFiles               ' Close any open files
    mfrmName.StopAnimation         ' Verify animation has been stopped
    StopTimer                   ' Stop API timer
        
    DoEvents
    If gblnStopProcessing Then
        
        If gblnLogData Then
            strRecord = GetTimeStamp & Space$(4) & "ABEND:  " & _
                        "User may have opted to stop processing." & vbTab & _
                        "Completed " & Format$(mcurByteCount, "#,##0") & " bytes (" & _
                        DisplayNumber(mcurByteCount) & ")"
            UpdateLogFile strRecord
        End If
    
    Else
        
        DoEvents
        If gblnLogData Then
            strRecord = GetTimeStamp & " FINISHED:  " & "Defrag of " & strTarget & _
                        vbTab & "Completed " & Format$(mcurByteCount, "#,##0") & _
                        " bytes (" & DisplayNumber(mcurByteCount) & ")"
            UpdateLogFile strRecord
        End If
    
        ' See if user wants to
        ' defrag windows page file
        DoEvents
        If gblnDoPageFile Then
        
            If gblnLogData Then
                strRecord = GetTimeStamp & "PAGE FILE:  " & _
                            "Windows page file will be defragmented immediately after rebooting"
                UpdateLogFile strRecord
            End If
                        
            ' Verify checkboxes and command buttons
            ' have been disabled because system is
            ' about to be rebooted
            With mfrmName
                ' Disable all checkboxes
                .chkLogData.Enabled = False
                .chkPageFile.Enabled = False
                .chkReboot.Enabled = False
            
                ' Disable all command buttons
                For lngIndex = 0 To .cmdChoice.Count - 1
                    .cmdChoice(lngIndex).Enabled = False
                Next lngIndex
                
                ' Display page defrag message on second line
                With .txtProgress
                    .BackColor = &H8000&     ' Forest green background
                    .ForeColor = &HFFFFFF    ' White lettering
                    .FontBold = True         ' Make Lettering bold
                    .FontSize = 16           ' Increase font size
                    .Text = vbNewLine & Space$(10) & "Setting page defrag switches"
                End With
            End With
            
            Wait 2000    ' need time to display page defrag message
            
            ' Set the flag to defrag windows page
            ' file after the reboot using Microsoft's
            ' Sysinternals utility "PageDfrg.exe"
            ' http://www.microsoft.com/technet/sysinternals/FileAndDisk/PageDefrag.mspx
            Shell gstrPageDfrg, vbHide
            DoEvents
            
            ' Display reboot message on second line
            DoEvents
            With mfrmName.txtProgress
                .BackColor = &HFF&       ' Bright red background
                .ForeColor = &HFFFFFF    ' White lettering
                .Text = vbNewLine & Space$(10) & "Reboot process has started"
            End With
    
            Wait 2000           ' need time to display reboot message
            DoShutdownProcess   ' Begin shutdown process
            
        ElseIf gblnReboot Then
            
            ' If reboot is checked then restart this machine
            ' if there was a successful completion
            '
            ' Verify checkboxes and command buttons
            ' have been disabled because system is
            ' about to be rebooted
            With mfrmName
                ' Disable all checkboxes
                .chkLogData.Enabled = False
                .chkPageFile.Enabled = False
                .chkReboot.Enabled = False
            
                ' Disable all command buttons
                For lngIndex = 0 To .cmdChoice.Count - 1
                    .cmdChoice(lngIndex).Enabled = False
                Next lngIndex
        
                ' Display reboot message on second line
                With .txtProgress
                    .BackColor = &HFF&       ' Bright red background
                    .ForeColor = &HFFFFFF    ' White lettering
                    .FontBold = True         ' Make Lettering bold
                    .FontSize = 16           ' Increase font size
                    .Text = vbNewLine & Space$(10) & "Reboot process has started"
                End With
            End With
            
            Wait 2000           ' need time to display reboot message
            DoShutdownProcess   ' Begin shutdown process
            
        Else
        
            ' General information message to
            ' remind user to reboot this machine
            InfoMsg "If defragging finished successfully, this machine should be " & vbNewLine & _
                    "rebooted so as to properly align internal folder and file pointers."
        End If
    End If
    
BeginDefrag_CleanUp:
    On Error GoTo 0  ' Nullify this error trap
    Exit Sub

BeginDefrag_Error:
    If gblnLogData Then
        strRecord = GetTimeStamp & Space$(4) & "ABEND:  " & Err.Description & _
                    "[" & MODULE_NAME & "." & ROUTINE_NAME & "]" & vbNewLine & _
                    "Completed " & Format$(mcurByteCount, "#,##0") & " bytes (" & _
                    DisplayNumber(mcurByteCount) & ")"
        UpdateLogFile strRecord
    End If
        
    Err.Clear
    Resume BeginDefrag_CleanUp
    
End Sub

Public Function StartTimer()
    
    ' Set number of milliseconds to equal one second.
    ' We are not interested in precision timing into
    ' the thousandths.
    glngMilliseconds = 1000
    
    If mlngTimerID <> 0 Then
        mlngTimerID = KillTimer(0, mlngTimerID)
    End If
    
    ' Start timer and capture the new timer ID handle.
    ' AddressOf Operator - Creates a procedure delegate
    ' instance that references the specific procedure.
    ' In other words, this will run in the background
    ' while not impeding the applications other functions.
    '
    ' WARNING!  You must call StopTimer() routine when
    ' the StartTimer() routine is no longer needed else
    ' this application will crash and sometimes force
    ' the user to power down and reboot.
    mlngTimerID = SetTimer(0, 0, glngMilliseconds, AddressOf TrackTime)

End Function

Private Sub TrackTime()
    mfrmName.ElapsedTime   ' Update elapsed time on form
End Sub

' Always stop the timer before closing
' your application or VB will crash.
Public Function StopTimer()
    mlngTimerID = KillTimer(0, mlngTimerID)
End Function


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       SearchAndProcess
'
' Description:   Parses a path and processes the selected files and folders.
'                Makes recursive calls so as to search subfolders.
'
' Parameters:    strTarget - Current folder being processed.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 25-JUL-2008  Kenneth Ives  kenaso@tx.rr.com
'              - Added file size check for files less than 2 bytes.
'              - Added log file reference
' 23-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Added API call if Contig takes longer to process a file
'              - Added option to log just path\name of file
' ***************************************************************************
Private Sub SearchAndProcess(ByVal strTarget As String)

    ' Called by BeginDefrag()
    
    Dim strFile      As String           ' current file name
    Dim strRecord    As String           ' log file record
    Dim strPathFile  As String           ' current folder or filename
    Dim strSubFolder As String           ' current folder or filename
    Dim hFile        As Long             ' folder or file handle
    Dim lngHwnd      As Long             ' Current process handle
    Dim curFilesize  As Currency         ' Size of current file
    Dim typWFD       As WIN32_FIND_DATA  ' Folder or file data structure
    Dim objDiskInfo  As cDiskInfo        ' class object

    On Error GoTo SearchAndProcess_Error

    ' See if user wants to stop processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If
    
    Set objDiskInfo = New cDiskInfo     ' Instantate class object
    ZeroMemory typWFD, Len(typWFD)      ' clear type structure
    strTarget = QualifyPath(strTarget)  ' Add a trailing backslash if missing
    strPathFile = strTarget & "*.*"     ' Prepare search criteria
    lngHwnd = 0                         ' Init handle value
    
    ' Capture first item
    hFile = FindFirstFile(strPathFile & Chr$(0), typWFD)

    ' If we have a valid handle then continue processing
    If hFile <> INVALID_HANDLE_VALUE Then
    
        Do
            ' See if user wants to stop processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do
            End If
                    
            ' Collect all the folder names
            DoEvents
            If (typWFD.dwFileAttributes And vbDirectory) Then

                strSubFolder = TrimStr(typWFD.cFilename)
                
                ' See if user wants to stop processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do
                End If
                    
                ' make sure this is not a subfolder identifier
                If strSubFolder <> "." And strSubFolder <> ".." Then
                    SearchAndProcess strTarget & strSubFolder
                End If
            
            Else
                ' This must be a file.
                DoEvents
                strFile = TrimStr(typWFD.cFilename)
                strPathFile = strTarget & strFile
                curFilesize = CalcFileSize(strPathFile)
                                
                ' See if user wants to stop processing
                DoEvents
                If gblnStopProcessing Then
                    Exit Do
                End If
                                     
                mcurByteCount = mcurByteCount + curFilesize   ' Update byte counter
                mcurFileCount = mcurFileCount + 1             ' Number of files accessed
                
                ' Update stats
                mfrmName.lblStats(1).Caption = Format$(mcurFileCount, "#,##0")
                mfrmName.lblStats(3).Caption = Format$(mcurByteCount, "#,##0") & "  (" & _
                                               objDiskInfo.DisplayNumber(mcurByteCount) & ")"
                mfrmName.txtProgress.Text = ShrinkToFit(strPathFile, 230)
                
                ' file size is greater than one byte
                If curFilesize > 1 Then
                    
                    ' Defrag all files on selected drive or in selected
                    ' folder using Microsoft's Sysinternals utility "Contig.exe"
                    ' http://www.microsoft.com/technet/sysinternals/FileAndDisk/Contig.mspx
                    DoEvents
                    If gblnLogData Then
                        
                        strRecord = GetTimeStamp & Space$(5) & "FILE:  " & strPathFile
                        UpdateLogFile strRecord
                        
                    End If
                        
                    ' Runs quietly while executing
                    Shell "contig.exe -q " & strPathFile, vbHide
                    Wait 500  ' Pause half a second
                    
                    ' See if Contig is still active
                    lngHwnd = FindProcessByCaption("contig")
                    
                    ' Handle will be captured only if the file is
                    ' fragmented to the degree that contig.exe will
                    ' take longer than a partial second to perform
                    ' its function
                    If lngHwnd > 0 Then
                        DoEvents
                        StopProcessByName "contig"   ' Kill all occurances of this process
                    End If
                        
                    ' See if user wants to stop processing
                    DoEvents
                    If gblnStopProcessing Then
                        GoTo SearchAndProcess_CleanUp
                    End If
                        
                Else
                    
                    DoEvents
                    If gblnLogData Then
                        strRecord = GetTimeStamp & " EXCLUDED:  " & strPathFile
                        UpdateLogFile strRecord
                    End If
    
                End If
            End If
GetNextFile:
            ' See if program has been paused then loop
            ' until resume or exit has been selected.
            ' Excessive DoEvents are here for a reason.
            DoEvents
            Do While gblnPgmPaused
                Wait 5000   ' Five second delay
            Loop
                    
            ' See if user wants to stop processing
            DoEvents
            If gblnStopProcessing Then
                Exit Do
            End If
            
        Loop While FindNextFile(hFile, typWFD)

    End If

SearchAndProcess_CleanUp:
    FindClose hFile                  ' always close file handles when not in use
    Set objDiskInfo = Nothing        ' Free class object from memory
    ZeroMemory typWFD, Len(typWFD)   ' clear type structure
    
    ' See if user wants to stop processing
    DoEvents
    If gblnStopProcessing Then
         
        If lngHwnd > 0 Then
            StopProcessByHandle lngHwnd   ' Stop Contig.exe, if active
            lngHwnd = 0                   ' Reset process handle
        End If
    End If

    On Error GoTo 0  ' nullify this error trap
    Exit Sub

SearchAndProcess_Error:
    DoEvents
    If gblnLogData Then
    
        ' Capture error data for log file
        strRecord = GetTimeStamp & "****ERROR:  " & strPathFile & _
                    " - " & Err.Description & " (" & CStr(Err.Number) & ")"
        Err.Clear                 ' Clear error code
        UpdateLogFile strRecord   ' Update log file
        
    Else
        Err.Clear      ' Clear error code
    End If
    
    GoTo GetNextFile   ' Process next file
    
End Sub

Private Function GetTimeStamp() As String

    ' Called by BeginDefrag()
    '           SearchAndProcess()
    
    ' Format system date and time for use in log file
    ' Ex:  "08.08.2008 05:33:44  "
    GetTimeStamp = Format$(Now(), "mm.dd.yyyy") & " " & _
                   Format$(Now(), "hh:nn:ss") & Space$(2)
    
End Function

Private Sub UpdateLogFile(ByVal strRecord As String)

    ' Called by BeginDefrag()
    '           SearchAndProcess()
    
    Dim hFile As Long
    
    Const ROUTINE_NAME As String = "UpdateLogFile"

    On Error GoTo UpdateLogFile_Error
    
    hFile = FreeFile
    Open gstrLogFile For Append As #hFile
    Print #hFile, strRecord
    DoEvents
    
UpdateLogFile_CleanUp:
    Close #hFile
    DoEvents
    
    On Error GoTo 0
    Exit Sub

UpdateLogFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume UpdateLogFile_CleanUp

End Sub

' ***************************************************************************
' Routine:       BuildLogFile
'
' Description:   Create a log file and folder if they do not exist.
'
' Parameters:    blnCreateNew - OPTIONAL - Create a new log file.
'                    DEFAULT - FALSE do not create a file.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Sub BuildLogFile()

    ' Called by BeginDefrag()
    
    Dim hFile As Long
    
    Const ROUTINE_NAME As String = "BuildLogFile"

    On Error GoTo BuildLogFile_Error
    
    CreateLogFileName   ' Get name of log file
    
    ' if log file does not exist
    ' then create a log file
    If Not IsPathValid(gstrLogFile) Then
        
        hFile = FreeFile                        ' Capture first free file handle
        Open gstrLogFile For Output As #hFile   ' Create an empty file
        Print #hFile, "Defrag Log File"         ' Write both title lines
        Print #hFile, "dd.mm.yyyy hh:mm:ss" & Space$(20) & "Description"
        Print #hFile, String$(80, "*")          ' Write row of asteriks
        Close #hFile                            ' Close file
    
    End If

BuildLogFile_CleanUp:
    DoEvents
    On Error GoTo 0
    Exit Sub

BuildLogFile_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    gblnStopProcessing = True
    Resume BuildLogFile_CleanUp
    
End Sub

Private Sub CreateLogFileName()

    ' Called by BuildLogFile()
    
    ' create a log file name using Julian data
    ' Example:  02/01/2004  -->  DF_04032.log
    Dim strFilename As String
    
    gstrLogFolder = QualifyPath(App.Path) & "Log"
    
    ' Create log folder if it does not exist
    If Not IsPathValid(gstrLogFolder) Then
        MkDir gstrLogFolder
    End If
    
    gstrLogFolder = QualifyPath(gstrLogFolder)      ' Append backslash
    strFilename = "DF_" & GetJulianDate() & ".log"  ' name of log file
    gstrLogFile = gstrLogFolder & strFilename       ' full path to log file
    
End Sub

' ***************************************************************************
' Routine:       GetJulianDate
'
' Description:   This procedure takes a normal date format (that is, mm/dd/yyyy)
'                and converts it to the appropriate Julian date (yyddd).
'
'                Most government agencies and contractors require the use of
'                Julian dates. A Julian date starts with a two-digit year,
'                and then counts the number of days from January 1 of that
'                year.
'
'                The Julian Date is returned in string format so as to
'                display all 5 digits to include any leading zeroes. If it
'                were returned in numeric format the Julian Date might be
'                truncated.  For example the Julian Date "00001"
'                (Jan 1, 2000) would appear as 1 because all leading
'                zeroes would be dropped.
'
' Parameters:    datDate - Date to be converted
'
' Returns:       Formatted date in string format to display all 5 digits to
'                include any leading zeroes.
'                Ex:  8/17/2007 --> "07229"
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 17-Aug-2007  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function GetJulianDate(Optional ByVal datDate As Date = Empty) As String

    ' Called by CreateLogFileName()
    
    ' If data variable has been emptied
    ' only time (midnight) will be available
    If datDate = "12:00:00 AM" Then
        datDate = Now()   ' Use system time
    End If
    
    GetJulianDate = Format$(datDate, "yy") & _
                    Format$(DatePart("y", datDate), "000")

End Function

' **************************************************************************
' Routine:       DisplayNumber
'
' Description:   Return a string representing the value in string format
'                to requested number of decimal positions.
'
'                    Bytes  Bytes
'                    KB     Kilobytes
'                    MB     Megabytes
'                    GB     Gigabytes
'                    TB     Terabytes
'                    PB     Petabytes
'
'                Ex:  75231309824 -> 70.1 GB
'
' Parameters:    dblCapacity - value to be reformatted
'                lngDecimals - [OPTIONAL] number of decimal positions.
'                     Valid values are 0-5.  Change to meet special needs.
'                     Default value = 1 decimal position
'
' Returns:       Reformatted string representation
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-DEC-2001  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 12-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated output
' ***************************************************************************
Public Function DisplayNumber(ByVal dblCapacity As Double, _
                     Optional ByVal lngDecimals As Long = 1) As String
  
    ' Called by BeginDefrag()
    '           SearchAndProcess()
    
    Dim intCount As Long
    Dim dblValue As Double
    
    Const MAX_DECIMALS As Long = 5   ' Change to meet special needs
    
    On Error GoTo DisplayNumber_Error
    
    dblValue = dblCapacity   ' I do this for debugging purposes
    intCount = 0             ' Counter for KB_1
    DisplayNumber = vbNullString
    
    If dblValue > 0 Then
        
        ' Must be a positive value
        If lngDecimals < 1 Then
            lngDecimals = 0
        End If
        
        ' Maximum of 5 decimal positions.
        If lngDecimals > MAX_DECIMALS Then
            lngDecimals = MAX_DECIMALS
        End If
    
        ' Loop thru input value and determine how
        ' many times it can be divided by 1024 (1 KB)
        Do While dblValue > (KB_1 - 1)
            dblValue = dblValue / KB_1
            intCount = intCount + 1
        Loop
        
        If lngDecimals = 0 Then
            ' Format value with no decimal positions
            DisplayNumber = Format$(Fix(dblValue), "0")
        Else
            ' Format value with requested decimal positions
            DisplayNumber = FormatNumber(dblValue, lngDecimals)
        End If
        
        DisplayNumber = DisplayNumber & " " & _
                        Choose(intCount + 1, "Bytes", "KB", "MB", "GB", "TB", "PB")
    Else
    
        ' No value was passed to this routine
        If lngDecimals = 0 Then
            DisplayNumber = "0 Bytes"     ' Format value with no decimal positions
        Else
            DisplayNumber = "0.0 Bytes"   ' Format value with one decimal position
        End If
    
    End If

DisplayNumber_Error:

End Function

' ***************************************************************************
' Routine:       DoShutdownProcess
'
' Description:   Reboot system
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 07-Dec-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Sub DoShutdownProcess()
    
    ' Called by BeginDefrag()
    
    Dim objReboot As cReboot
    
    ' Reset text box properties to normal
    DoEvents
    With mfrmName.txtProgress
        .BackColor = &H80000005   ' Normal Windows background
        .ForeColor = &H80000008   ' Normal Windows text
        .FontBold = False         ' Turn off bold font
        .FontSize = 9             ' Normal font size
        .Text = vbNullString                ' Remove text
    End With
        
    Wait 2000                     ' Pause for two seconds
    Set objReboot = New cReboot   ' Instantiate class object
    objReboot.ShutdownSystem      ' Reboot this system
    Set objReboot = Nothing       ' Free class object from memory
    
End Sub


