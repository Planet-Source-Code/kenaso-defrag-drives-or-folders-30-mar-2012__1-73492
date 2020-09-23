VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cReboot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Module:        clsReboot
'
' Description:   This module is used to shutdown, logoff or reboot a Windows
'                operating system. Has been tested on Windows 9x, NT, ME,
'                2000, XP and Vista.
'
' References:
'
'    PRB: ExitWindowsEx API Does Not cReboot Windows NT
'    http://support.microsoft.com/default.aspx?scid=kb;en-us;176695
'
'    DOC: ExitWindows Function Declaration Incorrect in API Viewer
'    http://support.microsoft.com/kb/168796/EN-US/
'
'    Steve McMahon  steve@vbaccelerator.com
'    http://www.vbaccelerator.com/home/VB/Tips/How_to_Shutdown_the_System_in_Windows_9x_and_NT/article.asp
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-SEP-2002  Kenneth Ives  kenaso@tx.rr.com
' ***************************************************************************
Option Explicit

' ***************************************************************************
' To Shutdown Windows:
' ***************************************************************************
  Private Const MODULE_NAME              As String = "clsReboot"

  ' Used with NT Shutdown privileges
  Private Const SE_SHUTDOWN_NAME         As String = "SeShutdownPrivilege"
  Private Const SE_PRIVILEGE_ENABLED     As Long = &H2
  Private Const TOKEN_QUERY              As Long = &H8
  Private Const TOKEN_ADJUST_PRIVILEGES  As Long = &H20

  ' Exit Windows Constants
  Private Const EWX_LOGOFF       As Long = &H0
  Private Const EWX_SHUTDOWN     As Long = &H1
  Private Const EWX_REBOOT       As Long = &H2
  Private Const EWX_FORCE        As Long = &H4
  Private Const EWX_POWEROFF     As Long = &H8
  Private Const EWX_FORCEIFHUNG  As Long = &H10

  ' These entries must follow above definitions
  Private Const EWX_FORCEDOWN_9X As Long = EWX_POWEROFF Or _
                                           EWX_SHUTDOWN Or _
                                           EWX_FORCE
                                           
  Private Const EWX_FORCEDOWN_NT As Long = EWX_POWEROFF Or _
                                           EWX_SHUTDOWN Or _
                                           EWX_FORCEIFHUNG
  ' To Report API errors
  Private Const FORMAT_MESSAGE_FROM_SYSTEM     As Long = &H1000
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type LUID
      LowPart  As Long
      HighPart As Long
  End Type
  
  Private Type LUID_AND_ATTRIBUTES
      pLuid      As LUID
      Attributes As Long
  End Type
  
  Private Type TOKEN_PRIVILEGES
      PrivilegeCount     As Long
      Privileges(0 To 0) As LUID_AND_ATTRIBUTES
  End Type

' ****************************************************************************
' Enumerations
' ****************************************************************************
  Public Enum enumExitProcess
      enuReboot     ' 0
      enuLogoff     ' 1
      enuShutdown   ' 2
      enuPowerOff   ' 3
  End Enum

' ****************************************************************************
' API Declares
' ****************************************************************************
  ' The GetCurrentProcess function returns a pseudohandle for the current
  ' process.
  Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

  ' The OpenProcessToken function opens the access token associated with
  ' a process.
  Private Declare Function OpenProcessToken Lib "advapi32.dll" _
          (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, _
          TokenHandle As Long) As Long

  ' Closes an open object handle.
  Private Declare Function CloseHandle Lib "kernel32" _
         (ByVal hObject As Long) As Long

  ' The LookupPrivilegeValue function retrieves the locally unique identifier
  ' (LUID) used on a specified system to locally represent the specified
  ' privilege name.
  Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
          Alias "LookupPrivilegeValueA" _
          (ByVal lpSystemName As String, ByVal lpName As String, _
          lpLuid As LUID) As Long
  
  ' The AdjustTokenPrivileges function enables or disables privileges in the
  ' specified access token. Enabling or disabling privileges in an access
  ' token requires TOKEN_ADJUST_PRIVILEGES access.
  Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" _
          (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, _
          NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, _
          PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

  ' The ExitWindowsEx function either logs off, shuts down, or shuts down and
  ' restarts the system. The ExitWindowsEx function returns as soon as it has
  ' initiated the shutdown. The shutdown or logoff then proceeds asynchronously.
  ' During a shutdown or log-off operation, applications that are shut down are
  ' allowed a specific amount of time to respond to the shutdown request. If
  ' the time expires, Windows displays a dialog box that allows the user to
  ' forcibly shut down the application, to retry the shutdown, or to cancel the
  ' shutdown request. If the EWX_FORCE value is specified, Windows always forces
  ' applications to close and does not display the dialog box.  The EWX_FORCE
  ' flag can cause the applications to lose data. Therefore, you should only
  ' use this flag in an emergency.
  Private Declare Function ExitWindowsEx Lib "user32" _
          (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long

  ' The FormatMessage function formats a message string. The function requires
  ' a message definition as input. The message definition can come from a
  ' buffer passed into the function. It can come from a message table resource
  ' in an already-loaded module. Or the caller can ask the function to search
  ' the system's message table resource(s) for the message definition. The
  ' function finds the message definition in a message table resource based on
  ' a message identifier and a language identifier. The function copies the
  ' formatted message text to an output buffer, processing any embedded insert
  ' sequences if requested.
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
          (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
          ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
          Arguments As Long) As Long


' ***************************************************************************
' ****                           Methods                                 ****
' ***************************************************************************

' ***************************************************************************
' Routine:       ShutdownSystem
'
' Description:   This is the main routine called by an outside reference.
'
' Parameters:    lngAction - Numeric representation of type of shutdown
'                            process to perform.  Defaults to zero.
'
'                     0 - EWX_REBOOT
'                     1 - EWX_POWERDOWN
'                     2 - EWX_LOGOFF
'
' Special Note:  On some Windows 9x operating systems, if the action
'                requested is LOGOFF then the PC will power off.  This does
'                not happen to all machines, just some of them.
'
'                Example:  CompUSA 500 mhz and Dell 200 mhz will power off
'                          Dell 500 mhz will Logoff.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 05-SEP-2002  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 29-MAR-2005  Kenneth Ives  kenaso@tx.rr.com
'              Added a forced shutdown for NT if system is hung
' ***************************************************************************
Public Sub ShutdownSystem(Optional ByVal lngAction As enumExitProcess = enuReboot)
Attribute ShutdownSystem.VB_Description = "Shutdown the operating system."

    Dim lngShutdownFlags As Long
    Dim blnWindowsNT     As Boolean
    Dim objOperSys       As cOperSystem

    Set objOperSys = New cOperSystem       ' Instantiate class object
    blnWindowsNT = objOperSys.bWindowsNT   ' Determine if Windows NT based operating system
    Set objOperSys = Nothing               ' Free class object from memory
    
    lngShutdownFlags = 0   ' Init shutdown flag

    ' Determine type of shutdown requested
    Select Case lngAction
           Case enuReboot:   lngShutdownFlags = EWX_REBOOT
           Case enuLogoff:   lngShutdownFlags = EWX_LOGOFF
           Case enuShutdown: lngShutdownFlags = EWX_POWEROFF
           Case enuPowerOff: lngShutdownFlags = EWX_POWEROFF
           Case Else:        lngShutdownFlags = EWX_REBOOT
    End Select

    ' If running under NT or better, the shutdown
    ' privledges need to be adjusted to allow the
    ' ExitWindowsEx() call.  If an adjust fails
    ' on a NT system, a shutdown will fail.
    If blnWindowsNT Then
        If Not AdjustShutdownPrivledges() Then
            Exit Sub
        End If
    End If

    ' Stop some known stubborn applications
    StopStubbornPgms
    DoEvents
    
    ' Windows NT based operating system
    If blnWindowsNT Then
                            
        ' Start shutdown process
        If ExitWindowsEx(lngShutdownFlags, 0&) = 0 Then
        
            ' If the system did not shutdown cleanly then
            ' force a shutdown and lose any data that was
            ' not previously saved.  At this point, it is
            ' too late to attempt saving anything.  Some
            ' processes have already started their shutdown
            ' events.
            DoEvents
            ExitWindowsEx EWX_FORCEDOWN_NT, 0&
        
        End If
        
    ' not Windows NT
    Else
        ' Start shutdown process
        If ExitWindowsEx(lngShutdownFlags, 0&) = 0 Then
        
            ' If the system did not shutdown cleanly then
            ' force a shutdown and lose any data that was
            ' not previously saved.  At this point, it is
            ' too late to attempt saving anything.  Some
            ' processes have already started their shutdown
            ' events.
            DoEvents
            ExitWindowsEx EWX_FORCEDOWN_9X, 0&
        
        End If
    End If
    
End Sub



' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       AdjustShutdownPrivledges
'
' Description:   When using an NT based operating system, you must make sure
'                that the user has proper authority to shut the system down.
'                Here, we grant temporary authority for this action.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 09-AUG-1998  Steve McMahon  steve@vbaccelerator.com
' http://www.vbaccelerator.com/home/VB/Tips/How_to_Shutdown_the_System_in_Windows_9x_and_NT/article.asp
' 29-MAR-2005  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function AdjustShutdownPrivledges() As Boolean

    ' Under NT we must enable the SE_SHUTDOWN_NAME
    ' privilege in the process we're trying to
    ' shutdown from, otherwise a to try to shutdown
    ' has no effect!
    
    Dim hProcess    As Long
    Dim hToken      As Long
    Dim lngTP_Old   As Long
    Dim lngRetValue As Long
    Dim typLUID     As LUID
    Dim typTP       As TOKEN_PRIVILEGES
    Dim typTP_OLD   As TOKEN_PRIVILEGES
    
    Const ROUTINE_NAME As String = "AdjustShutdownPrivledges"

    On Error GoTo AdjustShutdownPrivledges_Error

    ' Find the LUID of the Shutdown privilege token
    lngRetValue = LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, typLUID)
    
    ' If we get the authorization
    If (lngRetValue <> 0) Then
    
        ' Get the current process handle
        hProcess = GetCurrentProcess()
        
        If (hProcess <> 0) Then
            ' Open token for adjusting and querying
            ' (if we can - user may not have rights)
            lngRetValue = OpenProcessToken(hProcess, _
                                           TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, _
                                           hToken)
            
            If (lngRetValue <> 0) Then
            
                ' Ok we can now adjust the shutdown priviledges
                With typTP
                     .PrivilegeCount = 1
                     With .Privileges(0)
                          .Attributes = SE_PRIVILEGE_ENABLED
                          .pLuid.HighPart = typLUID.HighPart
                          .pLuid.LowPart = typLUID.LowPart
                     End With
                End With
                
                ' Now allow this process to shutdown the system
                lngRetValue = AdjustTokenPrivileges(hToken, 0, _
                                                    typTP, Len(typTP), _
                                                    typTP_OLD, lngTP_Old)
            
                If (lngRetValue <> 0) Then
                    AdjustShutdownPrivledges = True
                Else
                    InfoMsg "Cannot enable shutdown.  Cannot adjust privileges " & _
                            "to shutdown this system. [" & WinError(Err.LastDllError) & "]"
                End If
            
                ' Remember to close the handle when finished with it
                CloseHandle hToken
            Else
                InfoMsg "Cannot enable shutdown. Cannot open process token" & _
                        " to shutdown this system. [" & WinError(Err.LastDllError) & "]"
            End If
        Else
            InfoMsg "Cannot determine current process. " & _
                    "[" & WinError(Err.LastDllError) & "]"
        End If
    Else
        InfoMsg "Cannot find SE_SHUTDOWN_NAME privilege value." & _
                "[" & WinError(Err.LastDllError) & "]"
    End If

AdjustShutdownPrivledges_CleanUp:
    On Error GoTo 0
    Exit Function

AdjustShutdownPrivledges_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume AdjustShutdownPrivledges_CleanUp

End Function

Private Function WinError(ByVal lngLastDLLError As Long) As String

    Dim strBuffer As String
    Dim lngCount  As Long
    
    ' Return the error message associated with LastDLLError:
    strBuffer = String$(256, 0)
    lngCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                             FORMAT_MESSAGE_IGNORE_INSERTS, _
                             0, lngLastDLLError, 0&, _
                             strBuffer, Len(strBuffer), ByVal 0)
    If lngCount Then
        WinError = Left$(strBuffer, lngCount)
    End If
    
End Function


