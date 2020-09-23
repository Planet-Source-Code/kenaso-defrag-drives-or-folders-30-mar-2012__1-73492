VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOperSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
'  Module:        clsOperSystem
'
'  Description:   Determines the operating system.  When this class is
'                 initialized, the GetWindowsVersion() routine is called.
'                 This will instantly set the flags to their appropriate
'                 value.  This module does not test with operating sytems
'                 earlier than Windows NT4.
'
'                 Identifying the current operating system is usually not
'                 the best way to determine whether a particular operating
'                 system feature is present. This is because the operating
'                 system may have had new features added in a redistributable
'                 DLL. Rather than using GetVersionEx to determine the
'                 operating system platform or version number, test for the
'                 presence of the feature itself.
'                 http://msdn.microsoft.com/en-us/library/ms724832(VS.85).aspx
'
' -----------------------------------------------------------------------
' The base version information for various Windows versions is:
'
' Version                     Number  Major  Minor  Other info
' -----------------------------------------------------------------------
' Windows 8                    ?.?      ?      ?    Official release due in Fall of 2012
' Windows 7                    6.1      6      1    OSVERIONINFOEX.wProductType = VER_NT_WORKSTATION
' Windows Server 2008 R2       6.1      6      1    OSVERIONINFOEX.wProductType <> VER_NT_WORKSTATION
' Windows Server 2008          6.0      6      0    OSVERSIONINFOEX.wProductType <> VER_NT_WORKSTATION
' Windows Vista                6.0      6      0    OSVERSIONINFOEX.wProductType = VER_NT_WORKSTATION
' Windows Server 2003 R2       5.2      5      2    GetSystemMetrics(SM_SERVERR2) <> 0
' Windows Server 2003          5.2      5      2    GetSystemMetrics(SM_SERVERR2) = 0
' Windows Home Server          5.2      5      2    OSVERSIONINFOEX.wSuiteMask = VER_SUITE_WH_SERVER
' Windows XP Pro x64 Edition   5.2      5      2    (OSVERSIONINFOEX.wProductType = VER_NT_WORKSTATION) And
'                                                   (SYSTEM_INFO.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64)
' Windows XP                   5.1      5      1
' Windows 2000                 5.0      5      0
' Windows NT4                  4.0      4      0
' -----------------------------------------------------------------------
'
' Reference:   Randy Birch
'              GetVersionEx: Windows Version Info (Wrapper Routines)
'              http://vbnet.mvps.org/code/helpers/iswinversion.htm
'
'              Karl Perterson's Classic VB
'              http://www.mvps.org/vb/
'
'              GetVersionEx Function
'              http://msdn2.microsoft.com/en-us/library/ms724451(VS.85).aspx
'
'              Getting the System Version (Source code)
'              http://msdn2.microsoft.com/en-us/library/ms724429(VS.85).aspx
'
'              Operating System Version
'              http://msdn.microsoft.com/en-us/library/windows/desktop/ms724832%28v=vs.85%29.aspx
'              
'              OSVersionInfoEx Structure
'              http://msdn2.microsoft.com/en-us/library/ms724833(VS.85).aspx
'
'              GetSystemMetrics Function
'              http://msdn2.microsoft.com/en-us/library/ms724385(VS.85).aspx
'
'              GetSystemInfo Function
'              http://msdn2.microsoft.com/en-us/library/ms724381(VS.85).aspx
'
'              GetProductInfo Function
'              http://msdn2.microsoft.com/en-us/library/ms724358.aspx
'
' Important:   Developers working in the Windows Vista environment should be
'              aware that if this code is run in the VB IDE and you are
'              notified that the operating system is other than Windows Vista,
'              do the following:
'
'              1.  Close the VB IDE (Visual BASIC Integrated Development Environment)
'              2.  Locate the file VB.EXE
'              3.  Right mouse click and select Properties >> Compatibility tab
'              4.  Remove any checkmarks and close the window
'              5.  Restart the VB IDE
'
'  -----------------------------------------------------------------------
'
'  Example:    Set objOS = new cOperSystem  ' Instantiate class object
'
'              ' Is this Windows XP?
'              If objOS.bWinXP then
'                  Msgbox objOS.VersionName & vbNewLine & _
'                          "Version " & objOS.VersionNumber & " " & _
'                          objOS.ServicePack
'              Else
'                  Msgbox "This is not Windows XP"
'              End If
'
'              Set objOS = Nothing  ' Free class object from memory
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 15-Oct-2006  Kenneth Ives  kenaso@tx.rr.com
'              Converted to a class, renamed some of the variables,
'              Combined some of the routines for speed.  When this class is
'              initialized, the GetWindowsVersion() routine is called.  This
'              will instantly set the flags to their appropriate value.
' 16-Jan-2008  Kenneth Ives  kenaso@tx.rr.com
'              Dropped support for any operating systems earlier than Windows
'              NT4 Servers and Windows 2000.
' 17-Apr-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated for additional Windows Vista versions.
'              Updated 64-bit testing routines.
' 30-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated for Windows 2000 Data Center Server
' 20-Aug-2008  Kenneth Ives  kenaso@tx.rr.com
'              Updated testing logic for enhanced systems
' 04-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote module.
'              Added testing for additional Windows operating systems.
' 05-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Maria AAM for pointing out that I needed to add an
'              additional property named bOperSystem64() denoting a 64-bit
'              operating system in general
' 20-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated logic in TestForOldestVersion() routine
'              - Added property bWinXPSP3() for Windows XP Service Pack 3
'              - Updated IsWinXP() routine
'              - Updated TestForOtherServers() routine
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module. 
'              - Changed call to RemoveTrailingNulls() to TrimStr module 
'                due to speed and accuracy.
'              - Updated TestForOldestVersion() routine logic.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  ' used by OSVERSIONINFOEX.PlatformID
  Private Const VER_PLATFORM_WIN32_NT                As Long = 2

  ' used by OSVERSIONINFOEX.wProductType
  Private Const VER_NT_WORKSTATION                   As Long = &H1     ' Windows Vista, XP Pro, XP Home Edition, or 2000 Pro
  Private Const VER_NT_DOMAIN_CONTROLLER             As Long = &H2     ' Windows Server 2008, Windows Server 2003, or Windows 2000 Server
  Private Const VER_NT_SERVER                        As Long = &H3     ' Windows Server 2008, Windows Server 2003, or Windows 2000 Server

  ' used by OSVERSIONINFOEX.wSuiteMask
  Private Const VER_SUITE_SMALLBUSINESS              As Long = &H1     ' Microsoft Small Business Server w        As once installed on the system, but may have been upgraded to another version of Windows
  Private Const VER_SUITE_ENTERPRISE                 As Long = &H2     ' Windows Server 2008 Enterprise, Windows Server 2003, Enterprise Edition, or Windows 2000 Advanced Server is installed
  Private Const VER_SUITE_BACKOFFICE                 As Long = &H4     ' Windows backoffice suite
  Private Const VER_SUITE_TERMINAL                   As Long = &H10    ' Terminal Services is installed
  Private Const VER_SUITE_EMBEDDEDNT                 As Long = &H40    ' Windows XP Embedded is installed
  Private Const VER_SUITE_DATACENTER                 As Long = &H80    ' Windows Server 2008 Datacenter, Windows Server 2003, Datacenter Edition, or Windows 2000 Datacenter Server is installed
  Private Const VER_SUITE_PERSONAL                   As Long = &H200   ' Windows Vista Home Premium, Windows Vista Home Basic, or Windows XP Home Edition is installed
  Private Const VER_SUITE_BLADE                      As Long = &H400   ' Windows Server 2003, Web Edition is installed
  Private Const VER_SUITE_STORAGE_SERVER             As Long = &H2000  ' Windows Storage Server 2003 R2 or Windows Storage Server 2003 is installed
  Private Const VER_SUITE_COMPUTE_SERVER             As Long = &H4000  ' Windows Server 2003, Compute Cluster Edition is installed
  Private Const VER_SUITE_WH_SERVER                  As Long = &H8000  ' Windows home server is installed
  Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED   As Long = &H20    ' Microsoft Small Business Server is installed with the restrictive client license in force

  Private Const SM_TABLETPC                          As Long = 86
  Private Const SM_MEDIACENTER                       As Long = 87
  Private Const SM_STARTER                           As Long = 88
  Private Const SM_SERVERR2                          As Long = 89

  Private Const PROCESSOR_ARCHITECTURE_INTEL         As Long = 0       ' 32-bit
  Private Const PROCESSOR_ARCHITECTURE_IA64          As Long = 6       ' 64-bit
  Private Const PROCESSOR_ARCHITECTURE_AMD64         As Long = 9       ' 64-bit

  ' GetProductInfo possible values
  Private Const PRODUCT_ULTIMATE                     As Long = &H1     ' Ultimate Edition
  Private Const PRODUCT_HOME_BASIC                   As Long = &H2     ' Home Basic Edition
  Private Const PRODUCT_HOME_PREMIUM                 As Long = &H3     ' Home Premium Edition
  Private Const PRODUCT_ENTERPRISE                   As Long = &H4     ' Enterprise Edition
  Private Const PRODUCT_HOME_BASIC_N                 As Long = &H5     ' Home Basic Edition
  Private Const PRODUCT_BUSINESS                     As Long = &H6     ' Business Edition
  Private Const PRODUCT_STANDARD_SERVER              As Long = &H7     ' Server Standard Edition (full installation)
  Private Const PRODUCT_DATACENTER_SERVER            As Long = &H8     ' Server Datacenter Edition (full installation)
  Private Const PRODUCT_SMALLBUSINESS_SERVER         As Long = &H9     ' Small Business Server
  Private Const PRODUCT_SMALLBUSINESS_SERVER_PREMIUM As Long = &H19    ' Small Business Server Premium Edition
  Private Const PRODUCT_ENTERPRISE_SERVER            As Long = &HA     ' Server Enterprise Edition (full installation)
  Private Const PRODUCT_STARTER                      As Long = &HB     ' Starter Edition
  Private Const PRODUCT_DATACENTER_SERVER_CORE       As Long = &HC     ' Server Datacenter Edition (core installation)
  Private Const PRODUCT_STANDARD_SERVER_CORE         As Long = &HD     ' Server Standard Edition (core installation)
  Private Const PRODUCT_ENTERPRISE_SERVER_CORE       As Long = &HE     ' Server Enterprise Edition (core installation)
  Private Const PRODUCT_ENTERPRISE_SERVER_IA64       As Long = &HF     ' Server Enterprise Edition for Itanium-based Systems
  Private Const PRODUCT_BUSINESS_N                   As Long = &H10    ' Business Edition
  Private Const PRODUCT_WEB_SERVER                   As Long = &H11    ' Web Server Edition (full installation)
  Private Const PRODUCT_CLUSTER_SERVER               As Long = &H12    ' Cluster Server Edition
  Private Const PRODUCT_HOME_SERVER                  As Long = &H13    ' Home Server Edition
  Private Const PRODUCT_HOME_PREMIUM_N               As Long = &H1A    ' Home Premium Edition
  Private Const PRODUCT_ENTERPRISE_N                 As Long = &H1B    ' Enterprise Edition
  Private Const PRODUCT_ULTIMATE_N                   As Long = &H1C    ' Ultimate Edition
   
' ***************************************************************************
' Type structures
' ***************************************************************************
  ' The OSVERSIONINFOEX data structure contains operating system version
  ' information. The information includes major and minor version numbers,
  ' a build number, a platform identifier, and information about product
  ' suites and the latest Service Pack installed on the system. This structure
  ' is used with the GetVersionEx and VerifyVersionInfo functions.
  Private Type OSVERSIONINFOEX
      OSVSize           As Long          ' size of this data structure (in bytes)
      dwVerMajor        As Long          ' ex: 5
      dwVerMinor        As Long          ' ex: 01
      dwBuildNumber     As Long          ' ex: 2600
      PlatformID        As Long          ' Identifies operating system platform
      szCSDVersion      As String * 128  ' ex: "Service Pack 3"
      wServicePackMajor As Integer
      wServicePackMinor As Integer
      wSuiteMask        As Integer
      wProductType      As Byte
      wReserved         As Byte
  End Type

  ' The SYSTEM_INFO structure contains information about the current computer
  ' system.  This includes the architecture and type of the processor, the
  ' number of processors in the system, the page size, and other such information.
  Private Type OEM_ID
      wProcessorArchitecture As Integer
      wReserved              As Integer
  End Type
  
  Private Type SYSTEM_INFO
      dwOemID                     As OEM_ID
      dwPageSize                  As Long
      lpMinimumApplicationAddress As Long
      lpMaximumApplicationAddress As Long
      dwActiveProcessorMask       As Long
      dwNumberOfProcessors        As Long
      dwProcessorType             As Long
      dwAllocationGranularity     As Long
      wProcessorLevel             As Integer
      wProcessorRevision          As Integer
  End Type

  ' Windows operating system formatted information
  Private Type OS_INFORMATION
      PlatformID   As Long
      VersionName  As String
      VersionNo    As String
      BuildNo      As String
      ServicePack  As String
  End Type

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' This function obtains extended information about the version of the
  ' operating system that is currently running.
  Private Declare Function GetSystemMetrics Lib "user32" _
          (ByVal nIndex As Long) As Long

  ' This function obtains extended information about the version of the
  ' operating system that is currently running.  If the function succeeds,
  ' the return value is a nonzero value.
  Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
          (LpVersionInformation As Any) As Long

  ' The GetProcAddress function returns the address of the specified
  ' exported dynamic-link library (DLL) function. If the function call fails,
  ' the return value is null.
  Private Declare Function GetProcAddress Lib "kernel32" _
          (ByVal hModule As Long, ByVal lpProcName As String) As Long

  ' Retrieves a module handle for the specified module. The module must have
  ' been loaded by the calling process.
  Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" _
         (ByVal lpModuleName As String) As Long
  
  ' This function returns information about the current system.
  Private Declare Sub GetSystemInfo Lib "kernel32" _
          (lpSystemInfo As SYSTEM_INFO)
  
  ' Retrieves information about the current system running Windows 2000 Pro or
  ' newer.  It is equivalent to the GetSystemInfo function.
  Private Declare Sub GetNativeSystemInfo Lib "kernel32" _
          (lpSystemInfo As SYSTEM_INFO)

  ' Retrieves the product type for the operating system on the local
  ' computer, and maps the type to the product types supported by the
  ' specified operating system.
  Private Declare Function GetProductInfo Lib "kernel32" _
          (ByVal dwOSMajorVersion As Long, ByVal dwOSMinorVersion As Long, _
          ByVal dwSpMajorVersion As Long, ByVal dwSpMinorVersion As Long, _
          pdwReturnedProductType As Long) As Long

  ' ZeroMemory is used for clearing the contents of a type structure.
  Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
          (Destination As Any, ByVal Length As Long)

' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Boolean)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m bln WindowsNT
' Variable name:     mblnWindowsNT
'
' ***************************************************************************
  Private mtypOSInfo                        As OS_INFORMATION
  Private mtypOSVIEX                        As OSVERSIONINFOEX
  Private mtypSI                            As SYSTEM_INFO
  Private mstrVerName                       As String
  Private mstrSrvPack                       As String
  
  Private mblnWindowsNT                     As Boolean
  Private mblnWinNT4orNewer                 As Boolean
  Private mblnWin2000orNewer                As Boolean
  Private mblnWinXPorNewer                  As Boolean
  Private mblnWinVistaOrNewer               As Boolean
  
  Private mblnWin2000                       As Boolean
  Private mblnWin2000Pro                    As Boolean
  Private mblnWin2000Workstation            As Boolean
  Private mblnWin2000Server                 As Boolean
  Private mblnWin2000DatacenterSvr          As Boolean
  Private mblnWin2000AdvancedSvr            As Boolean
  
  Private mblnWinXP                         As Boolean
  Private mblnWinXPSP2                      As Boolean
  Private mblnWinXPSP3                      As Boolean
  Private mblnWinXPHomeEdition              As Boolean
  Private mblnWinXPPro                      As Boolean
  Private mblnWinXPMediaCenter              As Boolean
  Private mblnWinXPStarter                  As Boolean
  Private mblnWinXPTabletPC                 As Boolean
  Private mblnWinXPEmbedded                 As Boolean
  
  Private mblnWinVista                      As Boolean
  Private mblnWinVistaSP1                   As Boolean
  Private mblnWinVistaHomeBasic             As Boolean
  Private mblnWinVistaHomeEdition           As Boolean
  Private mblnWinVistaHomePremium           As Boolean
  Private mblnWinVistaHomeServer            As Boolean
  Private mblnWinVistaUltimate              As Boolean
  Private mblnWinVistaBusiness              As Boolean
  Private mblnWinVistaEnterprise            As Boolean
  Private mblnWinVistaWorkstation           As Boolean
  Private mblnWinVistaStarter               As Boolean
  Private mblnWindows7                      As Boolean
  
  Private mblnWinServer2003                 As Boolean
  Private mblnWin2003ServerR2               As Boolean
  Private mblnWin2003StorageServer          As Boolean
  Private mblnBladeServer                   As Boolean
  Private mblnWebServer                     As Boolean
  Private mblnWinHomeServer                 As Boolean
  Private mblnClusterServer                 As Boolean
  Private mblnComputeClusterServer          As Boolean
  Private mblnWinServer2008                 As Boolean
  Private mblnWinServer2008R2               As Boolean
  Private mblnDataCenterServer              As Boolean
  Private mblnDataCenterServerCore          As Boolean
  
  Private mblnBackOfficeServer              As Boolean
  Private mblnDomainController              As Boolean
  Private mblnEnterpriseServer              As Boolean
  Private mblnEnterpriseServerCore          As Boolean
  Private mblnTerminalServer                As Boolean
  Private mblnSmallBusinessServer           As Boolean
  Private mblnSmallBusinessServerPremium    As Boolean
  Private mblnSmallBusinessRestrictedServer As Boolean
  Private mblnStandardServer                As Boolean
  Private mblnStandardServerCore            As Boolean
  
  ' 64-Bit Versions
  Private mblnWinVista64                    As Boolean
  Private mblnWinXPPro64                    As Boolean
  Private mblnDataCenterItanium64           As Boolean
  Private mblnEnterpriseItanium64           As Boolean
  Private mblnDataCenter64                  As Boolean
  Private mblnEnterprise64                  As Boolean
  Private mblnStandard64                    As Boolean
  Private mblnComputeServer64               As Boolean
  Private mblnDatacenterServer64            As Boolean
  Private mblnEnterpriseServer64            As Boolean
  Private mblnWebBladeServer64              As Boolean
  Private mblnStandardServer64              As Boolean
  Private mblnOperSystem64                  As Boolean


' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Get VersionName() As String
    VersionName = mtypOSInfo.VersionName
End Property

Public Property Get VersionNumber() As String
    VersionNumber = mtypOSInfo.VersionNo
End Property

Public Property Get BuildNumber() As String
    BuildNumber = mtypOSInfo.BuildNo
End Property

Public Property Get VersionData() As String

    With mtypOSInfo
        VersionData = .VersionName & " " & _
                      .VersionNo & "." & _
                      .BuildNo
    End With
    
    If mblnWinNT4orNewer Then
        VersionData = VersionData & " " & ServicePack
    End If

End Property

Public Property Get ServicePack() As String
    ServicePack = mtypOSInfo.ServicePack
End Property

Public Property Get PlatformID() As Long
    PlatformID = mtypOSInfo.PlatformID
End Property

Public Property Get bWindowsNT() As Boolean
    bWindowsNT = mblnWindowsNT
End Property

Public Property Get bWinNT4orNewer() As Boolean
    bWinNT4orNewer = mblnWinNT4orNewer
End Property

Public Property Get bWin2000() As Boolean
    bWin2000 = mblnWin2000
End Property

Public Property Get bWin2000Pro() As Boolean
    bWin2000Pro = mblnWin2000Pro
End Property

Public Property Get bWin2000Workstation() As Boolean
    bWin2000Workstation = mblnWin2000Workstation
End Property

Public Property Get bWin2000Server() As Boolean
    bWin2000Server = mblnWin2000Server
End Property

Public Property Get bWin2000DatacenterSvr() As Boolean
    bWin2000DatacenterSvr = mblnWin2000DatacenterSvr
End Property

Public Property Get bWin2000AdvancedSvr() As Boolean
    bWin2000AdvancedSvr = mblnWin2000AdvancedSvr
End Property

Public Property Get bWin2000orNewer() As Boolean
    bWin2000orNewer = mblnWin2000orNewer
End Property

Public Property Get bWinXP() As Boolean
    bWinXP = mblnWinXP
End Property

Public Property Get bWinXPSP2() As Boolean
    bWinXPSP2 = mblnWinXPSP2
End Property

Public Property Get bWinXPSP3() As Boolean
    bWinXPSP3 = mblnWinXPSP3
End Property

Public Property Get bWinXPHomeEdition() As Boolean
    bWinXPHomeEdition = mblnWinXPHomeEdition
End Property

Public Property Get bWinXPProfessional() As Boolean
    bWinXPProfessional = mblnWinXPPro
End Property

Public Property Get bWinXPMediaCenter() As Boolean
    bWinXPMediaCenter = mblnWinXPMediaCenter
End Property

Public Property Get bWinXPStarter() As Boolean
    bWinXPStarter = mblnWinXPStarter
End Property

Public Property Get bWinXPTabletPC() As Boolean
    bWinXPTabletPC = mblnWinXPTabletPC
End Property

Public Property Get bWinXPEmbedded() As Boolean
    bWinXPEmbedded = mblnWinXPEmbedded
End Property

Public Property Get bWinXPorNewer() As Boolean
    bWinXPorNewer = mblnWinXPorNewer
End Property

Public Property Get bWindows7() As Boolean
    bWindows7 = mblnWindows7
End Property

Public Property Get bWinVistaStarter() As Boolean
    bWinVistaStarter = mblnWinVistaStarter
End Property

Public Property Get bWinServer2003() As Boolean
    bWinServer2003 = mblnWinServer2003
End Property

Public Property Get bWin2003ServerR2() As Boolean
    bWin2003ServerR2 = mblnWin2003ServerR2
End Property

Public Property Get bWin2003StorageServer() As Boolean
    bWin2003StorageServer = mblnWin2003StorageServer
End Property

Public Property Get bComputeClusterServer() As Boolean
    bComputeClusterServer = mblnComputeClusterServer
End Property

Public Property Get bWinVista() As Boolean
    bWinVista = mblnWinVista
End Property

Public Property Get bWinVistaSP1() As Boolean
    bWinVistaSP1 = mblnWinVistaSP1
End Property

Public Property Get bWinVistaHomeBasic() As Boolean
    bWinVistaHomeBasic = mblnWinVistaHomeBasic
End Property

Public Property Get bWinVistaHomeEdition() As Boolean
    bWinVistaHomeEdition = mblnWinVistaHomeEdition
End Property

Public Property Get bWinVistaHomePremium() As Boolean
    bWinVistaHomePremium = mblnWinVistaHomePremium
End Property

Public Property Get bWinVistaHomeServer() As Boolean
    bWinVistaHomeServer = mblnWinVistaHomeServer
End Property

Public Property Get bWinVistaUltimate() As Boolean
    bWinVistaUltimate = mblnWinVistaUltimate
End Property

Public Property Get bWinVistaBusiness() As Boolean
    bWinVistaBusiness = mblnWinVistaBusiness
End Property

Public Property Get bWinVistaEnterprise() As Boolean
    bWinVistaEnterprise = mblnWinVistaEnterprise
End Property

Public Property Get bWinVistaWorkstation() As Boolean
    bWinVistaWorkstation = mblnWinVistaWorkstation
End Property

Public Property Get bWinServer2008() As Boolean
    bWinServer2008 = mblnWinServer2008
End Property

Public Property Get bWinServer2008R2() As Boolean
    bWinServer2008R2 = mblnWinServer2008R2
End Property

Public Property Get bWinVistaOrNewer() As Boolean
    bWinVistaOrNewer = mblnWinVistaOrNewer
End Property

Public Property Get bDataCenterServer() As Boolean
    bDataCenterServer = mblnDataCenterServer
End Property

Public Property Get bDataCenterServerCore() As Boolean
    bDataCenterServerCore = mblnDataCenterServerCore
End Property

Public Property Get bBackOfficeServer() As Boolean
    bBackOfficeServer = mblnBackOfficeServer
End Property

Public Property Get bBladeServer() As Boolean
    bBladeServer = mblnBladeServer
End Property

Public Property Get bWebServer() As Boolean
    bWebServer = mblnWebServer
End Property

Public Property Get bWinHomeServer() As Boolean
    bWinHomeServer = mblnWinHomeServer
End Property

Public Property Get bClusterServer() As Boolean
    bClusterServer = mblnClusterServer
End Property

Public Property Get bDomainController() As Boolean
    bDomainController = mblnDomainController
End Property

Public Property Get bEnterpriseServer() As Boolean
    bEnterpriseServer = mblnEnterpriseServer
End Property

Public Property Get bEnterpriseServerCore() As Boolean
    bEnterpriseServerCore = mblnEnterpriseServerCore
End Property

Public Property Get bSmallBusinessServer() As Boolean
    bSmallBusinessServer = mblnSmallBusinessServer
End Property

Public Property Get bSmallBusinessServerPremium() As Boolean
    bSmallBusinessServerPremium = mblnSmallBusinessServerPremium
End Property

Public Property Get bSmallBusinessRestrictedServer() As Boolean
    bSmallBusinessRestrictedServer = mblnSmallBusinessRestrictedServer
End Property

Public Property Get bStandardServer() As Boolean
    bStandardServer = mblnStandardServer
End Property

Public Property Get bStandardServerCore() As Boolean
    bStandardServerCore = mblnStandardServerCore
End Property

Public Property Get bTerminalServer() As Boolean
    bTerminalServer = mblnTerminalServer
End Property

Public Property Get bWinVista64() As Boolean
    bWinVista64 = mblnWinVista64
End Property

Public Property Get bWinXPPro64() As Boolean
    bWinXPPro64 = mblnWinXPPro64
End Property

Public Property Get bDatacenterItanium64() As Boolean
    bDatacenterItanium64 = mblnDataCenterItanium64
End Property

Public Property Get bEnterpriseItanium64() As Boolean
    bEnterpriseItanium64 = mblnEnterpriseItanium64
End Property

Public Property Get bDatacenter64() As Boolean
    bDatacenter64 = mblnDataCenter64
End Property

Public Property Get bEnterprise64() As Boolean
    bEnterprise64 = mblnEnterprise64
End Property

Public Property Get bStandard64() As Boolean
    bStandard64 = mblnStandard64
End Property

Public Property Get bComputeServer64() As Boolean
    bComputeServer64 = mblnComputeServer64
End Property

Public Property Get bDatacenterServer64() As Boolean
    bDatacenterServer64 = mblnDatacenterServer64
End Property

Public Property Get bEnterpriseServer64() As Boolean
    bEnterpriseServer64 = mblnEnterpriseServer64
End Property

Public Property Get bWebBladeServer64() As Boolean
    bWebBladeServer64 = mblnWebBladeServer64
End Property

Public Property Get bStandardServer64() As Boolean
    bStandardServer64 = mblnStandardServer64
End Property

Public Property Get bOperSystem64() As Boolean
    bOperSystem64 = mblnOperSystem64
End Property


' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
'  Routine:      GetWindowsVersion
'
'  Description:  Get the current windows operating system version.
'                Called by Class_Initialize() routine when this class is
'                instantiated.  Each routine may perform additional tests
'                to further refine the type of operating system.
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-May-2009  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote module.
'              Added testing for additional Windows operating systems.
' 05-Sep-2009  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Maria AAM for pointing out that I needed to add an
'              additional property named bOperSystem64() denoting a 64-bit
'              operating system in general.
' ***************************************************************************
Private Sub GetWindowsVersion()

    mtypOSVIEX.OSVSize = Len(mtypOSVIEX)   ' Size of this data structure
    
    ' Make the call to load basic information
    ' into the data structure
    If GetVersionEx(mtypOSVIEX) <> 0 Then
        
        ' Test for specific features
        With mtypOSVIEX
            
            ' Is this an NT based Windows operating system?
            If (.PlatformID = VER_PLATFORM_WIN32_NT) Then
            
                ' See if GetNativeSystemInfo is supported
                ' (Windows 2000 Professional or newer)
                If GetProcAddress(GetModuleHandle("kernel32.dll"), "GetNativeSystemInfo") <> 0 Then
                    
                    GetNativeSystemInfo mtypSI  ' GetNativeSystemInfo is supported
                Else
                    GetSystemInfo mtypSI        ' GetNativeSystemInfo is NOT supported
                End If
                        
                IsWindows7
                IsWinServer2008
                IsWinServer2008R2
                IsWinVista
                IsWinXP
                IsWin2000
                TestForOldestVersion
                
                ' If operating system is Windows NT
                ' based then test for enhanced systems
                If mblnWindowsNT Then
                
                    IsSmallBusinessServer
                    IsSmallBusinessRestrictedServer
                    IsTerminalServer
                    IsBackOfficeServer
                    IsEnterpriseServer
                End If
               
                ' If operating system is Windows 2000 or
                ' newer then test for enhanced systems
                If mblnWin2000orNewer Or _
                   mblnWinXPorNewer Or _
                   mblnWinVistaOrNewer Then
                
                    IsDomainController
                    IsWin2000DatacenterSvr
                    IsWin2000AdvancedSvr
                    TestForOtherServers
                End If
                
                ' Save system information
                mtypOSInfo.PlatformID = .PlatformID
                mtypOSInfo.VersionName = "MS " & mstrVerName
                mtypOSInfo.VersionNo = .dwVerMajor & "." & Format$(.dwVerMinor, "00")
                mtypOSInfo.BuildNo = Format$((.dwBuildNumber And &HFFFF&), "0000")
                
                If Len(Trim$(mstrSrvPack)) > 0 Then
                    mtypOSInfo.ServicePack = mstrSrvPack
                Else
                    mtypOSInfo.ServicePack = TrimStr(.szCSDVersion)
                End If
        
            End If
        End With
    End If
    
End Sub

Private Function IsWindows7() As Boolean

    ' Returns TRUE if running Windows 7
    
    With mtypOSVIEX
        mblnWindows7 = (.dwVerMajor = 6) And _
                       (.dwVerMinor = 1) And _
                       (.wProductType = VER_NT_WORKSTATION)
    End With
    
    If mblnWindows7 Then
        mstrVerName = "Windows 7"
    End If
    
End Function

Private Function IsWinServer2008R2() As Boolean

    ' Returns TRUE if running Windows Server 2008 Release 2
    
    With mtypOSVIEX
        mblnWinServer2008R2 = (.dwVerMajor = 6) And _
                              (.dwVerMinor = 1) And _
                              (.wProductType <> VER_NT_WORKSTATION)
    End With
                
    If mblnWinServer2008R2 Then
        mstrVerName = "Windows Server 2008 Release 2"
    End If
    
End Function

Private Function IsWinServer2008() As Boolean

    ' Returns TRUE if running Windows Server 2008
    
    With mtypOSVIEX
        mblnWinServer2008 = (.dwVerMajor = 6) And _
                            (.dwVerMinor = 0) And _
                            (.wProductType <> VER_NT_WORKSTATION)
    End With
                
    If mblnWinServer2008 Then
        mstrVerName = "Windows Server 2008"
    End If

End Function

Private Function IsWinVista() As Boolean

    ' True if using Windows Vista
    ' Refine the search for a specific flavor of Vista
    ' More than one item may be flagged as TRUE

    Dim lngProduct As Long
    Dim strAppend  As String
    
    strAppend = vbNullString
    lngProduct = 0
    
    With mtypOSVIEX
        If (.dwVerMajor = 6) And _
           (.dwVerMinor = 0) And _
           (.wProductType = VER_NT_WORKSTATION) Then
                
            mblnWinVista = True
            mstrVerName = "Windows Vista "
                        
            If (.wSuiteMask And VER_SUITE_PERSONAL) Then
                mblnWinVistaHomeEdition = True
                strAppend = "Home Edition"
            
            Else
                mblnWinVistaWorkstation = True
                strAppend = "Workstation"
            End If
            
            If (mtypSI.dwOemID.wProcessorArchitecture And PROCESSOR_ARCHITECTURE_AMD64) Then
                mblnWinVista64 = True
                mblnOperSystem64 = True
                strAppend = "x64"
            End If
            
            If StrComp(Left$("Service", 7), TrimStr(.szCSDVersion), vbTextCompare) = 0 Then
                mblnWinVistaSP1 = True
                mstrSrvPack = TrimStr(.szCSDVersion)
            End If
            
            If GetProductInfo(.dwVerMajor, .dwVerMinor, 0&, 0&, lngProduct) <> 0 Then
    
                Select Case lngProduct
                       Case PRODUCT_ULTIMATE, PRODUCT_ULTIMATE_N
                            mblnWinVistaUltimate = True
                            strAppend = "Ultimate Edition"
    
                       Case PRODUCT_HOME_PREMIUM, PRODUCT_HOME_PREMIUM_N
                            mblnWinVistaHomePremium = True
                            strAppend = "Home Premium Edition"
                       
                       Case PRODUCT_HOME_BASIC, PRODUCT_HOME_BASIC_N
                            mblnWinVistaHomeBasic = True
                            strAppend = "Home Basic Edition"
                       
                       Case PRODUCT_ENTERPRISE, PRODUCT_ENTERPRISE_N
                            mblnWinVistaEnterprise = True
                            strAppend = "Enterprise Edition"
                       
                       Case PRODUCT_BUSINESS, PRODUCT_BUSINESS_N
                            mblnWinVistaBusiness = True
                            strAppend = "Business Edition"
                       
                       Case PRODUCT_STARTER
                            mblnWinVistaStarter = True
                            strAppend = "Starter Edition"
                       
                       Case PRODUCT_HOME_SERVER
                            mblnWinVistaHomeServer = True
                            strAppend = "Home Server"
                       
                       Case PRODUCT_CLUSTER_SERVER
                            mblnClusterServer = True
                            strAppend = "Cluster Server Edition"
                       
                       Case PRODUCT_DATACENTER_SERVER
                            mblnDataCenterServer = True
                            strAppend = "Datacenter Edition"
                       
                       Case PRODUCT_DATACENTER_SERVER_CORE
                            mblnDataCenterServerCore = True
                            strAppend = "Datacenter Edition (core installation)"
                       
                       Case PRODUCT_ENTERPRISE_SERVER
                            mblnEnterpriseServer = True
                            strAppend = "Enterprise Server Edition"
                       
                       Case PRODUCT_ENTERPRISE_SERVER_CORE
                            mblnEnterpriseServerCore = True
                            strAppend = "Enterprise Server Edition (core installation)"
                       
                       Case PRODUCT_ENTERPRISE_SERVER_IA64
                            mblnEnterpriseItanium64 = True
                            mblnOperSystem64 = True
                            strAppend = "Enterprise Edition for Itanium-based Systems"
                       
                       Case PRODUCT_SMALLBUSINESS_SERVER
                            mblnSmallBusinessServer = True
                            strAppend = "Small Business Server"
                       
                       Case PRODUCT_SMALLBUSINESS_SERVER_PREMIUM
                            mblnSmallBusinessServerPremium = True
                            strAppend = "Small Business Server Premium Edition"
                            
                       Case PRODUCT_STANDARD_SERVER
                            mblnStandardServer = True
                            strAppend = "Standard Server Edition"
                       
                       Case PRODUCT_STANDARD_SERVER_CORE
                            mblnStandardServerCore = True
                            strAppend = "Standard Server Edition (core installation)"
                       
                       Case PRODUCT_WEB_SERVER
                            mblnWebServer = True
                            strAppend = "Web Server Edition"
                End Select
            
            End If
            
            If (mtypSI.dwOemID.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                strAppend = strAppend & ", 64-bit"
                
            ElseIf (mtypSI.dwOemID.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_INTEL) Then
                strAppend = strAppend & ", 32-bit"
            End If
        
            mstrVerName = mstrVerName & Trim$(strAppend)
            
        End If
    End With
    
End Function

Private Sub IsWinXP()

    ' Returns TRUE if running Windows XP
    ' Then refine the search for a specific flavor of Vista
    ' More than one item may be flagged as TRUE

    Dim strAppend As String
    
    strAppend = vbNullString
    
    With mtypOSVIEX
        If (.dwVerMajor = 5) And _
           (.dwVerMinor = 1) Then
                    
            mblnWinXP = True
            mstrVerName = "Windows XP "
                    
            If (.wSuiteMask And VER_SUITE_PERSONAL) Then
                mblnWinXPHomeEdition = True
                strAppend = "Home Edition"
            
            ElseIf GetSystemMetrics(SM_MEDIACENTER) Then
                mblnWinXPMediaCenter = True
                strAppend = "Media Center"
            
            ElseIf GetSystemMetrics(SM_STARTER) Then
                mblnWinXPStarter = True
                strAppend = "Starter Edition"
            
            ElseIf GetSystemMetrics(SM_TABLETPC) Then
                mblnWinXPTabletPC = True
                strAppend = "Tablet PC"
            
            ElseIf (.wSuiteMask And VER_SUITE_EMBEDDEDNT) Then
                mblnWinXPEmbedded = True
                strAppend = "Embedded"
            
            Else
                mblnWinXPPro = True
                strAppend = "Professional"
            End If
            
            If StrComp(Left$("Service", 7), TrimStr(.szCSDVersion), vbTextCompare) = 0 Then
                
                mstrSrvPack = TrimStr(.szCSDVersion)
            
                If StrComp("Service Pack 2", mstrSrvPack, vbTextCompare) = 0 Then
                    mblnWinXPSP2 = True
                End If
                
                If StrComp("Service Pack 3", mstrSrvPack, vbTextCompare) = 0 Then
                    mblnWinXPSP3 = True
                End If
                
            End If
            
            mstrVerName = mstrVerName & Trim$(strAppend)
        End If
    End With
    
End Sub

Private Sub IsWin2000()

    ' Returns TRUE if running Windows 2000
    
    Dim strAppend As String
    
    strAppend = vbNullString
    
    With mtypOSVIEX
        If (.dwVerMajor = 5) And _
           (.dwVerMinor = 0) Then
        
            mblnWin2000 = True
            mstrVerName = "Windows 2000 "
            
            If (mtypOSVIEX.wProductType = VER_NT_WORKSTATION) Then
                
                mblnWin2000Workstation = True
                mblnWin2000Pro = True
                strAppend = "Professional"
            
            ElseIf (mtypOSVIEX.wProductType = VER_NT_SERVER) Then
            
                mblnWin2000Server = True
                strAppend = "Server"
            End If
            
            mstrVerName = mstrVerName & Trim$(strAppend)
        
        End If
    End With
        
End Sub

Private Sub TestForOldestVersion()

    With mtypOSVIEX
        ' Windows NT based operating system
        If (.dwVerMajor >= 4) Then
            mblnWindowsNT = True
        End If
        
        If (.dwVerMajor >= 6) Then
            mblnWinVistaOrNewer = True      ' Windows Vista or newer
                        
        ElseIf (.dwVerMajor = 5) Then
               
            If (.dwVerMinor = 0) Then
                mblnWin2000orNewer = True   ' Windows 2000 or newer
            
            ElseIf (.dwVerMinor >= 1) Then
                mblnWinXPorNewer = True     ' Windows XP or newer
            End If
            
        ElseIf (.dwVerMajor = 4) Then
            mblnWinNT4orNewer = True        ' Windows NT4 or newer
        End If
    End With
        
End Sub

Private Sub IsSmallBusinessServer()

    ' Called if Windows NT operating system
    ' Returns TRUE if Microsoft Small Business Server components have
    ' been installed
    
    ' You should not rely upon only the VER_SUITE_SMALLBUSINESS flag to
    ' determine whether Small Business Server has been installed on the
    ' system, as both this flag and the VER_SUITE_SMALLBUSINESS_RESTRICTED
    ' flag are set when this product suite is installed. If you upgrade
    ' this installation to Windows Server, Standard Edition, the
    ' VER_SUITE_SMALLBUSINESS_RESTRICTED flag will be cleared; however, the
    ' VER_SUITE_SMALLBUSINESS flag will remain set. In this case, this
    ' indicates that Small Business Server was once installed on this system.
    ' If this installation is further upgraded to Windows Server, Enterprise
    ' Edition, the VER_SUITE_SMALLBUSINESS flag will remain set.
    '
    ' http://msdn.microsoft.com/en-us/library/ms724833(VS.85).aspx
    
    mblnSmallBusinessServer = (mtypOSVIEX.wSuiteMask And VER_SUITE_SMALLBUSINESS)

    
End Sub

Private Sub IsSmallBusinessRestrictedServer()

    ' Called if Windows NT operating system
    ' Returns TRUE if Microsoft Small Business Restricted Server
    ' components have been installed with the restrictive client
    ' license in force
    
    mblnSmallBusinessRestrictedServer = (mtypOSVIEX.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)

End Sub

Private Sub IsTerminalServer()

    ' Called if Windows NT operating system
    ' Returns TRUE if Terminal Services is installed
    
    mblnTerminalServer = (mtypOSVIEX.wSuiteMask And VER_SUITE_TERMINAL)
    
End Sub

Private Sub IsBackOfficeServer()

    ' Called if Windows NT operating system
    ' Returns TRUE if Microsoft BackOffice Server components
    ' have been installed
    
    mblnBackOfficeServer = (mtypOSVIEX.wSuiteMask And VER_SUITE_BACKOFFICE)

End Sub

Private Sub IsEnterpriseServer()

    ' Called if Windows NT operating system
    ' Returns TRUE if Windows NT 4.0 Enterprise Edition,
    ' Windows 2000 Advanced Server, or Windows Server 2003
    ' then Windows Enterprise Edition is installed
    
    With mtypOSVIEX
        mblnEnterpriseServer = (.wProductType = VER_NT_SERVER) And _
                               (.wSuiteMask And VER_SUITE_ENTERPRISE)
    End With
    
End Sub

Private Sub IsDomainController()
  
    ' Called if Windows 2000 or newer
    ' Returns TRUE if the server is a Windows Domain Controller
    ' (Win 2000 or newer), including under active directory
        
    mblnDomainController = (mtypOSVIEX.wProductType = VER_NT_DOMAIN_CONTROLLER)
    
    If mblnDomainController Then
        mstrVerName = "Windows Domain Controller"
    End If
    
End Sub

Private Sub IsWin2000DatacenterSvr()

    ' Called if Windows 2000 or newer
    ' Returns TRUE if running Windows 2000 Datacenter Server
    
    mblnWin2000DatacenterSvr = (mtypOSVIEX.wSuiteMask And VER_SUITE_DATACENTER)
    
    If mblnWin2000DatacenterSvr Then
        mstrVerName = "Windows 2000 Datacenter Server"
    End If

End Sub
            
Private Sub IsWin2000AdvancedSvr()

    ' Called if Windows 2000 or newer
    ' Returns TRUE if running Windows 2000 Advanced Server
    
    With mtypOSVIEX
        mblnWin2000AdvancedSvr = ((.wProductType = VER_NT_SERVER) Or _
                                  (.wProductType = VER_NT_DOMAIN_CONTROLLER)) And _
                                 (.wSuiteMask And VER_SUITE_ENTERPRISE)
    End With
    
    If mblnWin2000AdvancedSvr Then
        mstrVerName = "Windows 2000 Advanced Server"
    End If

End Sub

Private Sub TestForOtherServers()

    ' Called if Windows 2000 or newer
    
    With mtypOSVIEX
        If (.dwVerMajor = 5) And _
           (.dwVerMinor = 2) Then
                
            ' Because the version numbers for Windows Server 2003 R2 and
            ' Windows Server 2003 are identical, you must also check the
            ' value of the SM_SERVERR2 setting.  If it returns a nonzero
            ' value, the operating system is Windows Server 2003 R2;
            ' otherwise, it is Windows Server 2003.
            If GetSystemMetrics(SM_SERVERR2) <> 0 Then
                mblnWin2003ServerR2 = True
                mstrVerName = "Windows Server 2003 R2, "
                            
            ' Additional testing
            ElseIf (.wSuiteMask And VER_SUITE_COMPUTE_SERVER) Then
                mblnComputeClusterServer = True
                mstrVerName = "Windows 2003 Compute Cluster Server Edition"
            
            ElseIf (.wSuiteMask = VER_SUITE_STORAGE_SERVER) Then
                mblnWin2003StorageServer = True
                mstrVerName = "Windows Storage Server 2003"
                
            ElseIf (.wSuiteMask = VER_SUITE_WH_SERVER) Then
                mblnWinHomeServer = True
                mstrVerName = "Windows Home Server"
                
            ElseIf (.wSuiteMask And VER_SUITE_BLADE) Then
                mblnBladeServer = True
                mstrVerName = "Windows Server 2003 Web Edition"
            
            ElseIf (.wProductType = VER_NT_WORKSTATION) And _
                   (mtypSI.dwOemID.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
                   
                mblnWinXPPro64 = True
                mblnOperSystem64 = True
                mstrVerName = "Windows XP Professional x64 Edition"

            Else
                mblnWinServer2003 = True
                mstrVerName = "Windows Server 2003"
            End If
            
            ' Test for additional 64-bit server type
            If (.wProductType <> VER_NT_WORKSTATION) Then
            
                ' See if running Itanium (Intel) based architecture
                If (mtypSI.dwOemID.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64) Then
            
                    mblnOperSystem64 = True
            
                    If (.wSuiteMask And VER_SUITE_DATACENTER) Then
                        mblnDataCenterItanium64 = True
                        mstrVerName = "Datacenter x64 Edition for Itanium-based Systems"
                    
                    ElseIf (.wSuiteMask And VER_SUITE_ENTERPRISE) Then
                        mblnEnterpriseItanium64 = True
                        mstrVerName = "Enterprise x64 Edition for Itanium-based Systems"
                    End If
            
                ' See if running AMD based architecture
                ElseIf (mtypSI.dwOemID.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64) Then
            
                    mblnOperSystem64 = True
            
                    If (.wSuiteMask And VER_SUITE_DATACENTER) Then
                        mblnDataCenter64 = True
                        mstrVerName = "Datacenter x64 Edition"
                        
                    ElseIf (.wSuiteMask And VER_SUITE_ENTERPRISE) Then
                        mblnEnterprise64 = True
                        mstrVerName = "Enterprise x64 Edition"
                        
                    Else
                        mblnStandard64 = True
                        mstrVerName = "Standard x64 Edition"
                    End If
                
                Else
            
                    If (.wSuiteMask And VER_SUITE_COMPUTE_SERVER) Then
                        mblnComputeServer64 = True
                        mstrVerName = "Compute Cluster x64 Edition"
                    
                    ElseIf (.wSuiteMask And VER_SUITE_DATACENTER) Then
                        mblnDatacenterServer64 = True
                        mstrVerName = "Datacenter x64 Edition"
                        
                    ElseIf (.wSuiteMask And VER_SUITE_ENTERPRISE) Then
                        mblnEnterpriseServer64 = True
                        mstrVerName = "Enterprise x64 Edition"
                        
                    ElseIf (.wSuiteMask And VER_SUITE_BLADE) Then
                        mblnWebBladeServer64 = True
                        mstrVerName = "Web Blade Server x64 Edition"
                        
                    Else
                        mblnStandardServer64 = True
                        mstrVerName = "Standard Server x64 Edition"
                    End If
                End If
                
            End If
        End If
    End With
    
End Sub

Private Sub ResetVariables()
    
    ' Make sure all flags are set to FALSE
    ' before determining Windows version.
    mblnWindowsNT = False
    mblnWinNT4orNewer = False
    mblnWin2000orNewer = False
    mblnWinXPorNewer = False
    mblnWinVistaOrNewer = False
   
    mblnWin2000 = False
    mblnWin2000Pro = False
    mblnWin2000Workstation = False
    mblnWin2000Server = False
    mblnWin2000DatacenterSvr = False
    mblnWin2000AdvancedSvr = False
   
    mblnWinXP = False
    mblnWinXPSP2 = False
    mblnWinXPSP3 = False
    mblnWinXPHomeEdition = False
    mblnWinXPPro = False
    mblnWinXPMediaCenter = False
    mblnWinXPStarter = False
    mblnWinXPTabletPC = False
    mblnWinXPEmbedded = False
                 
    mblnWinServer2003 = False
    mblnWin2003ServerR2 = False
    mblnBladeServer = False
    mblnClusterServer = False
    mblnComputeClusterServer = False
    mblnDataCenterServer = False
    mblnDataCenterServerCore = False
    mblnWebServer = False
    mblnWinHomeServer = False
    mblnWin2003StorageServer = False
   
    mblnWinVista = False
    mblnWinVistaSP1 = False
    mblnWinVistaHomeBasic = False
    mblnWinVistaHomeEdition = False
    mblnWinVistaHomePremium = False
    mblnWinVistaHomeServer = False
    mblnWinVistaUltimate = False
    mblnWinVistaBusiness = False
    mblnWinVistaEnterprise = False
    mblnWinVistaWorkstation = False
    mblnWinServer2008 = False
    mblnWinServer2008R2 = False
    mblnWinVistaStarter = False
    mblnWindows7 = False

    mblnBackOfficeServer = False
    mblnDomainController = False
    mblnEnterpriseServer = False
    mblnEnterpriseServerCore = False
    mblnTerminalServer = False
    mblnSmallBusinessServer = False
    mblnSmallBusinessServerPremium = False
    mblnSmallBusinessRestrictedServer = False
    mblnStandardServer = False
    mblnStandardServerCore = False
                                     
    mblnWinVista64 = False
    mblnWinXPPro64 = False
    mblnDataCenterItanium64 = False
    mblnEnterpriseItanium64 = False
    mblnDataCenter64 = False
    mblnEnterprise64 = False
    mblnStandard64 = False
    mblnComputeServer64 = False
    mblnDatacenterServer64 = False
    mblnEnterpriseServer64 = False
    mblnWebBladeServer64 = False
    mblnStandardServer64 = False
    mblnOperSystem64 = False
    
    ' Empty type structures
    ZeroMemory mtypOSInfo, Len(mtypOSInfo)
    ZeroMemory mtypOSVIEX, Len(mtypOSVIEX)
    ZeroMemory mtypSI, Len(mtypSI)
        
    mstrVerName = vbNullString
    mstrSrvPack = vbNullString
    
End Sub

Private Sub Class_Initialize()

    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objOS = new cOperSystem
            
    ResetVariables

    ' When this class is instantiated,
    ' capture windows version information.
    GetWindowsVersion

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objOS = Nothing
            
    ResetVariables

End Sub

