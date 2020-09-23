VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6060
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   6450
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picManifest 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   0
      ScaleHeight     =   6180
      ScaleWidth      =   6945
      TabIndex        =   8
      Top             =   -90
      Width           =   6945
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Log files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5190
         TabIndex        =   27
         ToolTipText     =   "Terminate this application"
         Top             =   3075
         Width           =   1065
      End
      Begin VB.CommandButton cmdFolders 
         Height          =   345
         Left            =   5970
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   1290
         Width           =   315
      End
      Begin VB.PictureBox picDefrag 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5340
         Picture         =   "frmMain.frx":0544
         ScaleHeight     =   465
         ScaleWidth      =   645
         TabIndex        =   23
         Top             =   450
         Width           =   645
      End
      Begin MSComCtl2.Animation aniProgress 
         Height          =   375
         Left            =   270
         TabIndex        =   22
         Top             =   4890
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   393
         FullHeight      =   25
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Pause"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   4050
         TabIndex        =   4
         ToolTipText     =   "Pause this application"
         Top             =   3075
         Width           =   1065
      End
      Begin VB.CheckBox chkReboot 
         Caption         =   "Reboot when finished"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   3750
         Width           =   2235
      End
      Begin VB.CheckBox chkLogData 
         Caption         =   "Use log file"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkPageFile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   150
         TabIndex        =   2
         Top             =   3255
         Width           =   3375
      End
      Begin MSComCtl2.Animation aniDefrag 
         Height          =   870
         Left            =   5310
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   105
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   1535
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   72
         FullHeight      =   58
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5190
         TabIndex        =   6
         ToolTipText     =   "Terminate this application"
         Top             =   3585
         Width           =   1065
      End
      Begin VB.CommandButton cmdChoice 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Stop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4020
         TabIndex        =   5
         ToolTipText     =   "Stop processing"
         Top             =   3585
         Width           =   1065
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4020
         TabIndex        =   7
         ToolTipText     =   "Start the defrag process"
         Top             =   3585
         Width           =   1065
      End
      Begin VB.TextBox txtProgress 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0986
         Top             =   1845
         Width           =   6135
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Select drive or folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   26
         Top             =   1050
         Width           =   1890
      End
      Begin VB.Label lblFolders 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFolders"
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1290
         Width           =   5745
      End
      Begin VB.Label lblProgressBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   300
         TabIndex        =   24
         Top             =   4905
         Width           =   5835
      End
      Begin VB.Label lblOperSysInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3285
         TabIndex        =   21
         Top             =   5595
         Width           =   2955
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4260
         TabIndex        =   20
         Top             =   1035
         Width           =   1950
      End
      Begin VB.Label lblStats 
         BackStyle       =   0  'Transparent
         Caption         =   "Byte count:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   19
         Top             =   4500
         Width           =   960
      End
      Begin VB.Label lblStats 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1215
         TabIndex        =   18
         Top             =   4500
         Width           =   5010
      End
      Begin VB.Label lblStats 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1215
         TabIndex        =   17
         Top             =   4230
         Width           =   2130
      End
      Begin VB.Label lblStats 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elapsed time:  0:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   3360
         TabIndex        =   16
         Top             =   4230
         Width           =   2745
      End
      Begin VB.Label lblStats 
         BackStyle       =   0  'Transparent
         Caption         =   "File count:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   15
         Top             =   4230
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   150
         X2              =   6255
         Y1              =   4095
         Y2              =   4095
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   180
         TabIndex        =   14
         Top             =   5370
         Width           =   2280
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         TabIndex        =   12
         Top             =   240
         Width           =   4980
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
         Left            =   2085
         TabIndex        =   11
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Current file being processed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   10
         Top             =   1620
         Width           =   2235
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmMain
'
' Description:   Perform a defragmentation of all files on all local logical
'                disks using external utilities.  Download the following
'                freeware utilities from Microsoft:
'
'            Place them in the same folder as this application.
'
'            Contig.exe v1.6  dtd 01-Feb-2011
'            A command-line utility that enables you to defrag individual
'            files, now supports defragmentation of NTFS metadata files,
'            including the MFT.
'            http://technet.microsoft.com/en-us/sysinternals/bb897428.aspx
'            http://en.wikipedia.org/wiki/Contig_(defragmentation_utility)
'
'            PageDfrg.exe v2.32  dtd 1-Nov-2006
'            Defrags windows page file and registry hives.  PageDefrag works
'            on 32-bit Windows NT 4.0, Windows 2000, Windows XP, Server 2003.
'            Pagedefrag does not work on Windows Vista or newer (32-bit) and
'            not on any version of 64-bit Windows.
'            http://technet.microsoft.com/en-us/sysinternals/bb897426
'            http://en.wikipedia.org/wiki/PageDefrag
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 16-Mar-2007  Kenneth Ives  kenaso@tx.rr.com
' 25-Oct-2009  Kenneth Ives  kenaso@tx.rr.com
'              - Added a pause processing  to the defrag process.
'              - Added byte and file count to display.
'              - Updated elpased time counter.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (String)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m str Target
' Variable name:     mstrTarget
'
' ***************************************************************************
  Private mstrTarget       As String
  Private mlngPrevAmt      As Long
  Private mlngMilliseconds As Long
  
Private Sub chkLogData_Click()

    ' If log data is checked then
    ' use a log file located in the
    ' application folder
    gblnLogData = CBool(chkLogData.Value)
    
End Sub

Private Sub chkPageFile_Click()
    
    gblnDoPageFile = CBool(chkPageFile.Value)
    
    ' If opted to defrag page file
    ' then a reboot is automatic
    If gblnDoPageFile Then
        chkReboot.Value = vbChecked
        chkReboot.Enabled = False
    Else
        ' If page file check box is manually
        ' unchecked then the reboot check box
        ' will be made available
        chkReboot.Enabled = True
        chkReboot.Value = vbUnchecked
        gblnReboot = False
    End If
    
End Sub

Private Sub chkReboot_Click()

    ' If reboot is checked then
    ' reboot this machine when
    ' finished defragging
    DoEvents
    gblnReboot = CBool(chkReboot.Value)
    DoEvents
    
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
    
           Case 0  ' Begin defrag process
                If Not IsPathValid(mstrTarget) Then
                    InfoMsg "Cannot identify target."
                    Exit Sub
                End If
                
                With frmMain
                    .cmdChoice(0).Visible = False   ' Disable GO button
                    .cmdChoice(1).Visible = True    ' Enable STOP button
                    .cmdChoice(2).Caption = Chr$(38) & "Pause"
                    .cmdChoice(2).Visible = True    ' Enable PAUSE button
                    .cmdChoice(3).Enabled = False   ' Disable LOG FILE button
                    .cmdChoice(4).Enabled = False   ' Disable EXIT button
                    .chkLogData.Enabled = False
                    .cmdFolders.Enabled = False
                    .lblStats(1).Caption = "0"
                    .lblStats(3).Caption = "0"
                    .lblStats(4).Caption = "Elapsed:  0:00:00"
                    .picDefrag.Visible = False       ' Hide defrag icon
                    .aniDefrag.Visible = True        ' Show defrag animation
                    .aniDefrag.Play                  ' Start defrag animation
                    .lblProgressBar.Visible = False  ' Hide empty progressbar
                    .aniProgress.Visible = True      ' Show progressbar animation
                    .aniProgress.Play                ' Start progressbar animation
                End With
                
                gblnPgmPaused = False
                gblnStopProcessing = False
                mlngPrevAmt = -1
                cmdChoice(1).SetFocus   ' Focus on STOP button
                Wait 1000
                
                BeginDefrag mstrTarget, frmMain
                ResetControls
                
           Case 1  ' Stop processing
                StopProcessing
                    
           Case 2  ' Pause processing
                PauseProcessing
                    
           Case 3  ' View log files
                gblnStopProcessing = False  ' reset STOP flag
                frmMain.Hide
                
                With frmLogFiles
                    .Reset_frmLogFiles
                    .Show
                End With
                    
           Case Else  ' Exit this application
                ResetControls
                TerminateProgram
    End Select
    
End Sub

' ***************************************************************************
' Routine:       cmdFolders_Click
'
' Description:   Displays the folder dialog box.  Allows the user to select
'                a folder to be wiped.  At this point, only the files in the
'                upper level folder will be targeted.  If there are no
'                subfolders or no files, then the folder will be deleted.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 04-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Update the amount of folder name to be displayed
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Added functionality to disallow the user to wipe a drive
'              starting at the root if that drive contains the Windows
'              operating system.
'              Added title when browsing for a folder.
' ***************************************************************************
Private Sub cmdFolders_Click()
    
    Dim strFolder As String     ' holds the folder name selected
    Dim objBrowse As cBrowse    ' class to display the folder dialog box
    
    On Error GoTo Cancel_Pressed
    
    With frmMain
        .lblStats(1).Caption = "0"
        .lblStats(3).Caption = "0"
        .lblStats(4).Caption = "Elapsed:  0:00:00"
    End With
    
    strFolder = vbNullString
    lblFolders.Caption = vbNullString
    
    Set objBrowse = New cBrowse   ' instantiate the class module
    strFolder = objBrowse.BrowseForFolder(frmMain, "Select drive\folder to defragment")
    Set objBrowse = Nothing       ' Free class object from memory
    
    ' see if a folder was selected
    If Len(Trim$(strFolder)) > 0 Then
        
        mstrTarget = QualifyPath(strFolder)
        lblFolders.Caption = ShrinkToFit(mstrTarget, 50)
    
    End If
    
cmdFolders_CleanUp:
    Set objBrowse = Nothing  ' Free class objects from memory
    On Error GoTo 0
    Exit Sub
    
Cancel_Pressed:
    ' Most likely the user selected
    ' CANCEL on the dialog box
    Resume cmdFolders_CleanUp
        
End Sub

Private Sub Form_Load()
    
    ResetControls
    
    With frmMain
        .Caption = PGM_NAME & gstrVersion
        .lblTitle.Caption = PGM_NAME
        .lblAuthor.Caption = AUTHOR_NAME
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                "No warranties or guarantees implied or intended."
        .lblOperSysInfo.Caption = gstrSysInfo
        .lblFolders.Caption = vbNullString
        .txtProgress.Text = vbNullString
        .chkPageFile.Value = vbUnchecked
        .chkPageFile.Caption = "Defrag windows page file" & vbNewLine & _
                               "Not available on 64-bit systems"
                 
        ' If running 64-bit or 32-bit Vista or Windows 7
        ' operating system then disable defrag pagefile checkbox
        If gblnOperSystem64 Or _
           gblnVistaOrNewer Then
           
            .chkPageFile.Enabled = False
        
        ' See if running 32-bit operating system on
        ' Windows NT or newer. Not Vista or Windows 7.
        ElseIf gblnWinNTorNewer Then
            .chkPageFile.Enabled = True
        End If
        
        .chkLogData.Value = vbUnchecked
        chkLogData_Click
        
        .picDefrag.Visible = True                 ' Show defrag icon
        .aniDefrag.Visible = False                ' Hide defrag animation
        .aniDefrag.Open App.Path & "\Defrag.avi"  ' Load defrag animation file
        
        .lblProgressBar.Visible = True                  ' Show empty progressbar
        .aniProgress.Visible = False                    ' Hide progressbar animation
        .aniProgress.Open App.Path & "\TimeLapse.avi"   ' Load progressbar animation file
        
        ' Center form on screen
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        .Show vbModeless  ' Reduce flicker
        .Refresh
    End With
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' "X" in upper right corner was selected
    If UnloadMode = 0 Then
        StopProcessing
        TerminateProgram
    End If
    
End Sub

Private Sub StopProcessing()
    
    DoEvents
    gblnStopProcessing = True
    
    StopAnimation   ' Stop progressbar animation
    StopTimer       ' Stop API timer or VB will crash!
    
    DoEvents
    With frmMain
        
        With .txtProgress
            ' Display message on second line
            .BackColor = &HFF0000    ' Bright blue background
            .ForeColor = &HFFFFFF    ' White lettering
            .FontBold = True         ' Make Lettering bold
            .FontSize = 16           ' Increase font size
            .Text = vbNewLine & Space$(15) & "Closing defrag process"
        End With
        
        Wait 3000       ' Pause for 3 seconds
            
        With .txtProgress
            .BackColor = &H80000005   ' Regular Windows background
            .ForeColor = &H80000008   ' Regular Windows lettering
            .FontBold = False         ' Turn off bold setting
            .FontSize = 9             ' Regular font size
            .Text = vbNullString                ' Remove text
        End With
    
    End With
    
    DoEvents
    ResetControls

End Sub

Private Sub PauseProcessing()
    
    DoEvents
    gblnPgmPaused = Not gblnPgmPaused       ' Toggle flag
    
    DoEvents
    With frmMain
        If gblnPgmPaused Then
            
            StopAnimation                   ' Stop progressbar animation
            StopTimer                       ' Stop API timer or VB will crash!
            
            .cmdChoice(1).Enabled = False   ' Disable STOP button
            .cmdChoice(2).Caption = Chr$(38) & "Resume"
            .cmdChoice(2).SetFocus          ' Focus on RESUME button
            .cmdChoice(3).Enabled = True    ' Enable LOG FILE button
            .cmdChoice(4).Enabled = True    ' Enable EXIT button
            
            With .txtProgress
                ' Display message on second line
                .BackColor = &HFF0000       ' Bright blue background
                .ForeColor = &HFFFFFF       ' White lettering
                .FontBold = True            ' Make Lettering bold
                .FontSize = 16              ' Increase font size
                .Text = vbNewLine & Space$(15) & "Pausing defrag process"
                
                Wait 3000   ' Pause for 3 seconds
                
                .BackColor = &H80000005     ' Regular Windows background
                .ForeColor = &H80000008     ' Regular Windows lettering
                .FontBold = False           ' Turn off bold setting
                .FontSize = 9               ' Regular font size
                .Text = vbNullString                  ' Remove text
            End With
        Else
            .cmdChoice(2).Caption = Chr$(38) & "Pause"
            .cmdChoice(3).Enabled = False   ' Disable LOG FILE button
            .cmdChoice(4).Enabled = False   ' Disable EXIT button
            .cmdChoice(1).Enabled = True    ' Enable STOP button
            .cmdChoice(1).SetFocus          ' Focus on STOP button
            .picDefrag.Visible = False      ' Hide defrag icon
            .aniDefrag.Visible = True       ' Show defrag animation
            .aniDefrag.Play                 ' Start defrag animation
            .lblProgressBar.Visible = False ' Hide empty progressbar
            .aniProgress.Visible = True     ' show progressbar animation
            .aniProgress.Play               ' Start progressbar animation
            StartTimer                      ' Start API timer
        End If
    End With
    
    DoEvents
    
End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub
  

' ***************************************************************************
' ***                 Global routines for this form                       ***
' ***************************************************************************

' ***************************************************************************
' Routine:       ElapsedTime
'
' Description:   Formats time display
'
' Reference:     Karl E. Peterson, http://vb.mvps.org/
'
' Returns:       Formatted output
'                Ex:  12:34:56    <- 12 hours 34 minutes 56 seconds
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Aug-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Sub ElapsedTime()

    Dim lngDays         As Long
    Dim lngMilliseconds As Long
    Dim strElapsed      As String
    
    Const ONE_DAY As Long = 86400000   ' Number of milliseconds in a day
    
    strElapsed = "Elapsed:  "
    mlngMilliseconds = mlngMilliseconds + glngMilliseconds
    lngMilliseconds = mlngMilliseconds
    
    lngDays = Fix(lngMilliseconds / ONE_DAY)   ' Calculate number of days
        
    ' If one or more days has passed
    ' then prefix output string
    If lngDays > 0 Then
        strElapsed = strElapsed & CStr(lngDays) & " day(s)  "
        lngMilliseconds = lngMilliseconds - (ONE_DAY * lngDays)
    End If

    ' Continue formatting output string as HH:MM:SS
    strElapsed = strElapsed & Format$(DateAdd("s", (lngMilliseconds \ 1000), #12:00:00 AM#), "HH:MM:SS")
    
    frmMain.lblStats(4).Caption = strElapsed

End Sub

Public Sub ResetControls()

    DoEvents
    mlngMilliseconds = 0
    StopTimer                  ' Stop API timer or VB will crash!
    
    DoEvents
    With frmMain
        .cmdChoice(0).Visible = True    ' Enable GO button
        .cmdChoice(1).Visible = False   ' Disable STOP button
        .cmdChoice(2).Visible = False   ' Disable PAUSE\RESUME button
        .cmdChoice(2).Caption = Chr$(38) & "Pause"
        .cmdChoice(3).Enabled = True    ' Enable LOG FILE button
        .cmdChoice(4).Enabled = True    ' Enable EXIT button
        .cmdFolders.Enabled = True      ' Disable BROWSE FOR FOLDERS button
        .lblStats(1).Caption = "0"
        .lblStats(3).Caption = "0"
        .lblStats(4).Caption = "Elapsed:  0:00:00"
    
        ' If running 64-bit or 32-bit Vista or Windows 7
        ' operating system then disable defrag pagefile checkbox
        If gblnOperSystem64 Or _
           gblnVistaOrNewer Then
           
            .chkPageFile.Enabled = False
        
        ' See if running 32-bit operating system on
        ' Windows NT or newer. Not Vista or Windows 7.
        ElseIf gblnWinNTorNewer Then
            .chkPageFile.Enabled = True
            .chkPageFile.Value = vbUnchecked
        End If
        
        .chkReboot.Enabled = True
        .chkReboot.Value = vbUnchecked
        
        .chkLogData.Enabled = True
        .chkLogData.Value = vbUnchecked
        
        With .txtProgress
            .BackColor = &H80000005     ' Regular Windows background
            .ForeColor = &H80000008     ' Regular Windows lettering
            .FontBold = False           ' Turn off bold setting
            .FontSize = 9               ' Regular font size
            .Text = vbNullString                  ' Remove text
        End With
    End With
        
    StopAnimation
    DoEvents

End Sub

Public Sub StopAnimation()

    DoEvents
    With frmMain
        .picDefrag.Visible = True       ' Show defrag icon
        .aniDefrag.Visible = False      ' Hide defrag animation
        .aniDefrag.Stop                 ' Stop defrag animation
        .lblProgressBar.Visible = True  ' Show empty progressbar
        .aniProgress.Visible = False    ' Hide progressbar animation
        .aniProgress.Stop               ' Stop progressbar animation
    End With
    DoEvents

End Sub

