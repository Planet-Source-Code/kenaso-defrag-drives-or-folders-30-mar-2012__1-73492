Attribute VB_Name = "modFileSize"
' ***************************************************************************
'  Module:     modFileSize.bas
'
'  Purpose:    This module calculates the size of a file.  Can handle
'              file sizes greater than 2gb.
'
' Reference:   Richard Newcombe  22-Jan-2007
'              Getting Past the 2 Gb File Limit
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
'
'              How To Seek Past VBA's 2GB File Limit
'              http://support.microsoft.com/kb/189981
'
' Description: The descriptions in this module are excerts from Richard
'              Newcombe's article.
'
'              When working in the IDE, any numbers that are entered are
'              limited to a Long variable type. Actually, as far I've
'              found, the IDE uses Longs for most numeric storage within
'              the projects that you write.
'
'              Okay, so what's the problem with Longs? Well, by definition
'              they are a signed 4-byte variable, in hex &H7FFFFFFF, with a
'              lower limit of -2,147,483,648 and an upper limit of
'              2,147,483,647 (2 Gb). &H80000000 stores the sign of the
'              value. Even when you enter values in Hex, they are stored in
'              a Long.
'
'              Working with random access files, you quite often use a Long
'              to store the filesize and current position, completely
'              unaware that if the file you access is just one byte over
'              the 2 Gb size, you can cause your application to corrupt the
'              file when writing to it.
'
'              Unfortunately, there is no quick fix for this. To get around
'              the problem, you need to write your own file handling
'              module, one that uses windows APIs to open, read, write, and
'              close any file.
'
'              The API's expect the Low and High 32-bit values in unsigned
'              format. Also, the APIs return unsigned values. So, the first
'              thing you have to do is decide on a variable type that you
'              can use to store values higher than 2 Gb. After some serious
'              thought, I decided to use a Currency type (64-bit scaled
'              integer) this gives you a 922,337 gig upper file limit, way
'              bigger that the largest hard drive available today.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              http://www.codeguru.com/vb/controls/vb_file/directory/article.php/c12917__1/
' 13-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated CalcFileSize() routine.
' ***************************************************************************
Option Explicit

' ********************************************************************
' Constants
' ********************************************************************
  Private Const MODULE_NAME          As String = "modFileSize"
  Private Const INVALID_HANDLE_VALUE As Long = -1
  Private Const GB_4                 As Currency = (2 ^ 32)  ' 4294967296
  
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
' API Declares
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
' ***                           Methods                                   ***
' ***************************************************************************

' ***************************************************************************
' Routine:       CalcFileSize
'
' Description:   This routine is used to open a file as read only and
'                calculate it's size.
'
' WARNING:       Always make a backup of the files that are to be processed.
'
' Parameters:    strFileName  - Name of file
'
' Returns:       Fle size in bytes
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 22-Jan-2007  Richard Newcombe
'              Wrote routine
' 13-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated file size calculations.
' ***************************************************************************
Public Function CalcFileSize(ByVal strFilename As String) As Currency

    Dim hHandle     As Long
    Dim curFilesize As Currency
    Dim typWFD      As WIN32_FIND_DATA
    
    Const ROUTINE_NAME As String = "CalcFileSize"
    
    On Error GoTo CalcFileSize_Error
    
    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Function
    End If
    
    curFilesize = 0@
    ZeroMemory typWFD, Len(typWFD)                 ' Empty type structure
    hHandle = FindFirstFile(strFilename, typWFD)   ' Load file stats into type structure
    
    ' If valid file handle retrieved
    If hHandle <> INVALID_HANDLE_VALUE Then
    
        FindClose hHandle  ' Close file handle
        
        ' Calculate file size
        With typWFD
            If .nFileSizeLow < 0 Then
                curFilesize = GB_4 + .nFileSizeLow
            Else
                curFilesize = CCur(.nFileSizeLow)
            End If
            
            curFilesize = (GB_4 * .nFileSizeHigh) + curFilesize
        End With

    End If
    
CalcFileSize_CleanUp:
    If hHandle > 0 Then
        FindClose hHandle   ' Verify file handle is closed
    End If
    
    CalcFileSize = curFilesize   ' Return file size
    On Error GoTo 0
    Exit Function

CalcFileSize_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    curFilesize = 0@
    Resume CalcFileSize_CleanUp

End Function

                              
