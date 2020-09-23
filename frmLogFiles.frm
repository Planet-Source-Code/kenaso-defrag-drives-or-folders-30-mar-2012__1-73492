VERSION 5.00
Begin VB.Form frmLogFiles 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   3810
   ControlBox      =   0   'False
   Icon            =   "frmLogFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   3810
   Begin VB.ListBox lstLogFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   45
      TabIndex        =   5
      Top             =   135
      Width           =   3660
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   2190
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save all entries and leave"
      Top             =   4665
      Width           =   750
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Save all entries and leave"
      Top             =   4665
      Width           =   750
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Return to Options screen without saving"
      Top             =   4665
      Width           =   750
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No log files found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   4
      Top             =   1770
      Width           =   3240
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   4875
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmLogFiles
'
' Description:   This form displays the list of log files that have been
'                generated.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Verified screen would not be displayed during initial load
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
  Private Const LB_SETTABSTOPS        As Long = &H192
                
' ***************************************************************************
' API Declares
' ***************************************************************************
  ' SetFileAttributes Function sets the attributes for a file or directory.
  ' If the function succeeds, the return value is nonzero.
  Private Declare Function SetFileAttributes Lib "kernel32" _
          Alias "SetFileAttributesA" _
          (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

  ' The SendMessage function sends the specified message to a window or
  ' windows. The function calls the window procedure for the specified
  ' window and does not return until the window procedure has processed
  ' the message.
  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
          (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
          lParam As Any) As Long


Private Sub cmdLog_Click(Index As Integer)

    Dim intIndex  As Integer
    Dim strFile   As String
    Dim strFolder As String
    
    On Error GoTo cmdLog_Click_Error

    Select Case Index
    
           Case 0    ' view log file
                ' prepare folder name
                strFolder = QualifyPath(gstrLogFolder)
                
                ' capture name of selected file
                For intIndex = 0 To (lstLogFiles.ListCount - 1)
                    If lstLogFiles.Selected(intIndex) Then
                        strFile = lstLogFiles.List(intIndex)  ' Capture listbox data
                        strFile = Trim$(Left$(strFile, 12))   ' Strip out just file name
                        Exit For    ' exit For..Next loop
                    End If
                Next intIndex
                                
                If Len(Trim$(strFile)) > 0 Then
                    
                    ' display log file using default text editor
                    DisplayFile strFolder & strFile, frmLogFiles
                End If
                
           Case 1    ' delete a log file
                ' prepare the folder name
                strFolder = QualifyPath(gstrLogFolder)
                
                ' capture name of selected file
                For intIndex = 0 To (lstLogFiles.ListCount - 1)
                    If lstLogFiles.Selected(intIndex) Then
                        strFile = lstLogFiles.List(intIndex)  ' Capture listbox data
                        strFile = Trim$(Left$(strFile, 12))   ' Strip out just file name
                        Exit For    ' exit For..Next loop
                    End If
                Next intIndex
                
                If ResponseMsg("Delete this log file?") = vbYes Then
                    ' reset file attributes
                    SetFileAttributes strFolder & strFile, FILE_ATTRIBUTE_NORMAL
                    Kill strFolder & strFile  ' delete log file
                    Reset_frmLogFiles         ' refresh file list
                End If
                                
           Case Else  ' leave
                frmLogFiles.Hide   ' hide this form
                frmMain.Show       ' show main form
    End Select

cmdLog_Click_CleanUp:
    On Error GoTo 0
    Exit Sub

cmdLog_Click_Error:
    ErrorMsg "frmLogFiles", "cmdLog_Click", Err.Description
    Resume cmdLog_Click_CleanUp
    
End Sub

Private Sub Form_Initialize()

    ' Make sure this form is hidden
    ' during initial load
    frmLogFiles.Hide
    DoEvents
    
End Sub

Private Sub Form_Load()

    Dim objkeyedit As cKeyEdit
    
    DisableX frmLogFiles   ' Disable "X" in upper right corner of form
    
    With frmLogFiles
        .Caption = "Defrag Log Files"
        
        Set objkeyedit = New cKeyEdit
        objkeyedit.CenterCaption frmLogFiles   ' Center form window caption
        Set objkeyedit = Nothing
        
        ' center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Hide
    End With

    Reset_frmLogFiles
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Based on the unload code the system passes, we determine what to do.
    '
    ' Unloadmode codes
    '     0 - Close from the control-menu box or Upper right "X"
    '     1 - Unload method from code elsewhere in the application
    '     2 - Windows Session is ending
    '     3 - Task Manager is closing the application
    '     4 - MDI Parent is closing
    Select Case UnloadMode
           Case 0    ' return to main form
                frmMain.Show
                frmLogFiles.Hide
    
           Case Else
                ' Fall thru. Something else is shutting us down.
    End Select

End Sub

Public Sub Reset_frmLogFiles()

    Dim strMsg  As String
    
    strMsg = vbNullString
    gstrLogFolder = QualifyPath(gstrLogFolder)
    
    If IsPathValid(gstrLogFolder) Then
        ' See if there is a log file in this folder
        strMsg = Dir$(gstrLogFolder & "*.log")
    End If
    
    ' Find any log files?
    If Len(Trim$(strMsg)) > 0 Then
        strMsg = vbNullString                    ' Yes, load log file names
    Else
        strMsg = "No log files found"  ' No log files found
    End If

    LoadListbox strMsg

End Sub

Private Sub LoadListbox(Optional strMsg As String = vbNullString)

    Dim lngIndex   As Long
    Dim astrData() As String
    Dim alngTabs() As Long
  
    Erase alngTabs()
    
    With frmLogFiles
        
        If Len(Trim$(strMsg)) = 0 Then
            
            ReDim alngTabs(0 To 1)  ' Array for 2 additional columns
            alngTabs(0) = 60        ' Tab stop for second column
            alngTabs(1) = 150       ' Tab stop for third column (currently not used)
            
            ' Set listbox tab stops
            Call SendMessage(.lstLogFiles.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
            Call SendMessage(.lstLogFiles.hwnd, LB_SETTABSTOPS, 2, alngTabs(0))
            
            ' Load log file list with file names
            GetLogFiles astrData()    ' Load array with names of log files
            SortFileNames astrData()  ' Sort with newest date first (descending)
            
            .lblMsg.Visible = False
            
            With .lstLogFiles
            
                .Clear   ' Always empty listbox before refilling
                
                ' Load file names into listbox
                For lngIndex = 0 To UBound(astrData) - 1
                    .AddItem astrData(lngIndex)
                Next lngIndex
                
                .Visible = True  ' Display listbox
                .ListIndex = 0   ' Highlight first item in listbox
            End With
            
            ' Enable the View & Delete buttons
            .cmdLog(0).Enabled = True
            .cmdLog(0).Visible = True
            .cmdLog(1).Enabled = True
            .cmdLog(1).Visible = True
        Else
            ' Display a message for no log files.
            .lstLogFiles.Visible = False
            .lblMsg.Visible = True
            .lblMsg.Caption = strMsg
            
            ' Disable the View & Delete buttons.
            .cmdLog(0).Enabled = False
            .cmdLog(0).Visible = False
            .cmdLog(1).Enabled = False
            .cmdLog(1).Visible = False
        End If
    End With
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub lstLogFiles_DblClick()
    cmdLog_Click 0
End Sub

' ***************************************************************************
' Routine:       SortFileNames
'
' Description:   Bubble sort, sometimes shortened to bubblesort, also known
'                as exchange sort, is a simple sorting algorithm. It works by
'                repeatedly stepping through the list to be sorted, comparing
'                two items at a time and swapping them if they are in the
'                wrong order. The pass through the list is repeated until no
'                swaps are needed, which means the list is sorted. The
'                algorithm gets its name from the way smaller elements
'                "bubble" to the top (i.e. the beginning) of the list via the
'                swaps. Because it only uses comparisons to operate on
'                elements, it is a comparison sort. This is the easiest
'                comparison sort to implement.
'
'                This function performs a bubble sort on an array of string
'                data while keeping track of the element indices. The indices
'                (pointers) will be returned to arrange the original data
'                in a sorted format.
'
' Parameters:    astrData() - Array to be sorted
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 13-Feb-2008  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Sub SortFileNames(ByRef astrData() As String)

    Dim lngIndex  As Long    ' loop counter
    Dim lngLow    As Long    ' lowest number of elements in the array
    Dim lngHigh   As Long    ' highest number of elements in the array
    Dim blnSorted As Boolean ' Flag to determine is data is sorted

    ' An error occurred or user opted to STOP processing
    DoEvents
    If gblnStopProcessing Then
        Exit Sub
    End If

    lngHigh = UBound(astrData)
    lngLow = LBound(astrData)
    
    ' File names will be sorted in descending order
    Do
        blnSorted = True  ' assume array is sorted
        
        For lngIndex = (lngLow + 1) To lngHigh
            
            If StrComp(astrData(lngIndex - 1), astrData(lngIndex), vbTextCompare) = -1 Then
                SwapData astrData(lngIndex - 1), astrData(lngIndex)
                blnSorted = False  ' Set flag denoting array is not sorted
            End If
            
        Next lngIndex

    Loop Until blnSorted

End Sub

' ***************************************************************************
' Routine:       SwapData
'
' Description:   Swap data with each other.  I wrote this function since
'                BASIC stopped having its own SWAP function.  I use this
'                for swapping strings, type structures, numbers with
'                decimal values, etc.
'
' Parameters:    vntValue1 - Incoming data to be swapped with Value2
'                vntValue2 - Incoming data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-NOV-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function SwapData(ByRef vntData1 As Variant, _
                          ByRef vntData2 As Variant)

    Dim vntHold As Variant

    vntHold = Empty   ' Always start with empty variants

    vntHold = vntData1
    vntData1 = vntData2
    vntData2 = vntHold

    vntHold = Empty   ' Always empty variants when not needed

End Function

Private Sub GetLogFiles(ByRef astrData() As String)

    Dim lngIndex  As Long
    Dim strName   As String
    Dim objFile   As File
    Dim objFiles  As Files
    Dim objFolder As Folder
    Dim objFSO    As Scripting.FileSystemObject
    
    Const INCREMENT As Long = 25
    
    Erase astrData()            ' Always start with an empty array
    
    ReDim astrData(INCREMENT)   ' Size receiving array
    lngIndex = 0                ' Initialize array index pointer
    
    Set objFSO = New Scripting.FileSystemObject      ' Instantiate scripting object
    Set objFolder = objFSO.GetFolder(gstrLogFolder)  ' Capture folder object
    Set objFiles = objFolder.Files                   ' Capture list of files in folder
    
    ' Loop thru list of files
    ' and select only log files
    For Each objFile In objFiles
        
        strName = objFile.Name   ' Capture name of file
            
        ' Capture just log file name extensions
        If StrComp(".log", Right$(strName, 4), vbTextCompare) = 0 Then
                    
            ' Add file name to array
            astrData(lngIndex) = strName & vbTab & objFile.DateCreated
            
            ' Increment array index
            lngIndex = lngIndex + 1
        
            ' Increase array size if needed
            If lngIndex Mod INCREMENT = 0 Then
                ReDim Preserve astrData(lngIndex + INCREMENT)
            End If
            
        End If
        
    Next objFile
    
    ' Resize arrray to just what was used
    ReDim Preserve astrData(lngIndex)
    
    ' Always free objects from
    ' memory when not needed
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFolder = Nothing
    Set objFSO = Nothing

End Sub
