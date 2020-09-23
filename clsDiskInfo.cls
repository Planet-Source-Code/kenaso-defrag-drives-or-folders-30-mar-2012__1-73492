VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDiskInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ***************************************************************************
' Several of these routines were found at a site that is
' maintained by Randy Birch.
' http://vbnet.mvps.org/index.html
'
' Runs only on Windows NT and newer, for two reasons:
'
' The standard device input/output control codes are available only
' on Windows NT and newer.
'
' On Windows NT based operating systems, an application can use the
' DeviceIoControl API function to perform direct input and output
' operations on, or retrieve information about, a floppy disk drive,
' hard disk drive, tape drive, or CD-ROM drive. This page demonstrates
' how to retrieveinformation about the installed  physical drives on
' the system. It uses the CreateFile function to obtain a device handle
' to the physical drives, and then uses DeviceIoControl with the
' IOCTL_DISK_GET_DRIVE_GEOMETRY control code to fill a DISK_GEOMETRY
' structure with information about the drive.
'
' The IOCTL_DISK_GET_DRIVE_GEOMETRY control code returns information
' about the physical disk's geometry: type, number of cylinders,
' tracks per cylinder, sectors per track, and bytes per sector. The
' DISK_GEOMETRY structure describes the geometry of disk devices and
' media. Note that the first parameter of the DISK_GEOMETRY type is
' defined as a LARGE_INTEGER, requiring the VB Currency data type to
' be used instead of a traditional Long.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Replaced FileExists() and PathExists() routines with
'              IsPathValid() routine.
' 15-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Added 3 holding areas to type structure DRIVE_INFORMATION
'              - Updated SwapBytes(), GetMfrInfo() and GetDriveInfo() routines.
'              - Corrected formatting of volume serial number in
'                GetVolumeInfo() routine
'              - Added minor documentation
' 21-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Added GetDriveLetters() routine
' 15-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rearranged parameters and made return parameters optional in
'              GetMfrInfo() routine.
' 10-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Replaced type structure DISK_INFORMATION with multiple
'                properties so user can access disk information more easily.
'              - Updated GetDriveLetters(), GetClusterSize(), DisplayNumber()
'                routines.
'              - Added GetDiskNumber() routine.
'              - Removed some obsolete code.
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Rewrote GetDriveInfo(), GetClusterSize() routines.
'              - Updated SpecificTypeOfDrive() to catch USB devices.
'              - Updated GetDiskNumber() routine to capture USB information.
'              - Renamed GetMediaSize() routine to GetDriveTypeExtra().
'              - Updated GetMediaInfo() routine. Moved from Public to Private.
'              - Rewrote GetDiskGeometryEx() routine.
' 20-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              - Updated GetMfrInfo(), GetPhysicalDriveInfo() routines.
' 20-Nov-2011  Kenneth Ives  kenaso@tx.rr.com
'              Added new function to determine if CD drive is physically
'              mounted or just a drive letter (Virtual).  See IsCDMounted()
'              routine.
' 31-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for noticing that I truncated
'              the longer hard drive serial numbers as used with SSD-drives.
'              See GetMfrInfo() routine and step through the process.
' 26-Mar-2012  Kenneth Ives  kenaso@tx.rr.com
'              - Deleted RemoveTrailingNulls() routine from this module. 
'              - Changed call to RemoveTrailingNulls() to TrimStr module 
'                due to speed and accuracy.
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const MODULE_NAME                      As String = "clsDiskInfo"
  Private Const MAX_SIZE                         As Long = 260
  Private Const ERROR_FILE_NOT_FOUND             As Long = 2
  Private Const FORMAT_MESSAGE_FROM_SYSTEM       As Long = &H1000
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS    As Long = &H200
  Private Const FILE_DEVICE_CD_ROM               As Long = &H2
  Private Const FILE_DEVICE_DVD                  As Long = &H33
  Private Const FILE_DEVICE_DISK                 As Long = &H7
  Private Const FILE_DEVICE_VIRTUAL_DISK         As Long = &H24
  Private Const FILE_SHARE_READ                  As Long = &H1
  Private Const FILE_SHARE_WRITE                 As Long = &H2
  Private Const FILE_ANY_ACCESS                  As Long = 0
  Private Const OPEN_EXISTING                    As Long = 3
  Private Const GENERIC_WRITE                    As Long = &H40000000
  Private Const GENERIC_READ                     As Long = &H80000000
  Private Const INVALID_HANDLE_VALUE             As Long = -1
  Private Const IOCTL_DISK_BASE                  As Long = FILE_DEVICE_DISK
  Private Const IOCTL_STORAGE_CHECK_VERIFY       As Long = &H2D4800
  Private Const IOCTL_STORAGE_MEDIA_REMOVAL      As Long = &H2D4804
  Private Const IOCTL_STORAGE_GET_MEDIA_TYPES_EX As Long = &H2D0C04
  Private Const KB_1                             As Long = &H400&   ' 1024
  Private Const KB_2                             As Long = &H800&   ' 2048
  Private Const METHOD_BUFFERED                  As Long = 0
  Private Const TEN_THOUSAND                     As Currency = 10000@

  ' Registry
  Private Const KEY_READ                         As Long = &H20009
  Private Const REG_BINARY                       As Long = 3
  Private Const HKEY_LOCAL_MACHINE               As Long = &H80000002

  ' Disk serial and model info
  Private Const VER_PLATFORM_WIN32_NT            As Long = 2
  Private Const IDENTIFY_BUFFER_SIZE             As Long = 512
  Private Const OUTPUT_DATA_SIZE                 As Long = IDENTIFY_BUFFER_SIZE + 16
  
  ' Valid values for the bCommandReg member of IDEREGS.
  Private Const IDE_ID_FUNCTION                  As Long = &HEC ' Returns ID sector for ATA.
  Private Const IDE_EXECUTE_SMART_FUNCTION       As Long = &HB0 ' Performs SMART cmd.
                                                                ' Requires valid bFeaturesReg,
                                                                ' bCylLowReg, and bCylHighReg

  ' Cylinder register values required when issuing SMART command
  Private Const SMART_CYL_LOW                    As Long = &H4F
  Private Const SMART_CYL_HI                     As Long = &HC2
  
  ' Vendor specific feature register defines
  ' for SMART "sub commands"
  Private Const SMART_ENABLE_SMART_OPERATIONS    As Long = &HD8

  ' IOCTL commands
  Private Const DFP_GET_VERSION                  As Long = &H74080
  Private Const DFP_SEND_DRIVE_COMMAND           As Long = &H7C084
  Private Const DFP_RECEIVE_DRIVE_DATA           As Long = &H7C088

' ***************************************************************************
' Enumerations
' ***************************************************************************
  Public Enum enumTypeOfRemovable
      eFloppy            ' 0
      eUSB_Drive         ' 1
      eBothTypes         ' 2
  End Enum

  Public Enum enumDiskSpace
      eFreespace         ' 0
      eTotalSize         ' 1
      eUsedSpace         ' 2
      eAvailSpace        ' 3
  End Enum

  Public Enum enumDriveType
      eUnknown           ' 0 Unknown drive type
      eBadRoot           ' 1 No root directory
      eRemovable         ' 2 Floppy or flash drive
      eFixed             ' 3 Local hard drive
      eNetwork           ' 4 Shared Network drive
      eCDRom             ' 5 CD-Rom drive (CD or DVD)
      eRamdisk           ' 6 Virtual memory disk
  End Enum

  Public Enum enumMediaSize
      eUnknownMedia      ' 0
      eF5_1Pt2_512       ' 1
      eF3_1Pt44_512      ' 2
      eF3_2Pt88_512      ' 3
      eF3_20Pt8_512      ' 4
      eF3_720_512        ' 5
      eF5_360_512        ' 6
      eF5_320_512        ' 7
      eF5_320_1024       ' 8
      eF5_180_512        ' 9
      eF5_160_512        ' 10
      eRemovableMedia    ' 11
      eFixedMedia        ' 12
  End Enum

  Public Enum enumIDE_DRIVE_NUMBER
      ePrimaryMaster     ' 0
      ePrimarySlave      ' 1
      eSecondaryMaster   ' 2
      eSecondarySlave    ' 3
      eTertiaryMaster    ' 4
      eTertiarySlave     ' 5
      eQuartiaryMaster   ' 6
      eQuartiarySlave    ' 7
  End Enum

  ' Status Flags Values
  Public Enum enumSTATUS_FLAGS
      ePrefailureWarranty = &H1
      eOnLineCollection = &H2
      ePerformanceAttribute = &H4
      eErrorRateAttribute = &H8
      eEventCouontAttribute = &H10
      eSELF_PRESERVING_ATTRIBUTE = &H20
  End Enum

' ***************************************************************************
' Type structures
' ***************************************************************************
  Private Type OSVERSIONINFO
      OSVSize       As Long
      dwVerMajor    As Long
      dwVerMinor    As Long
      dwBuildNumber As Long
      PlatformID    As Long
      szCSDVersion  As String * 128
  End Type

  Private Type PREVENT_MEDIA_REMOVAL
      PreventMediaRemoval  As Byte
  End Type

  Private Type ATTR_DATA
      AttrID         As Byte
      AttrName       As String
      AttrValue      As Byte
      ThresholdValue As Byte
      WorstValue     As Byte
      StatusFlags    As enumSTATUS_FLAGS
  End Type

  Private Type DRIVE_INFO
      bDriveType     As Byte
      SerialNumber   As String
      Model          As String
      FirmWare       As String
      Cylinders      As Long
      Heads          As Long
      SecPerTrack    As Long
      BytesPerSector As Long
      BytesperTrack  As Long
      NumAttributes  As Byte
      Attributes()   As ATTR_DATA
  End Type

  Private Type DEVICE_MEDIA_INFO
      Cylinders            As Double
      MediaType            As Long
      TracksPerCylinder    As Long
      SectorsPerTrack      As Long
      BytesPerSector       As Long
      NumberMediaSides     As Long
      MediaCharacteristics As Long
  End Type

  Private Type GET_MEDIA_TYPES
      DeviceType     As Long
      MediaInfoCount As Long
      MediaInfo(10)  As DEVICE_MEDIA_INFO
  End Type

  ' GETVERSIONOUTPARAMS contains the data returned
  ' from the Get Driver Version function
  Private Type GETVERSIONOUTPARAMS
      bVersion      As Byte  ' Binary driver version.
      bRevision     As Byte  ' Binary driver revision
      bReserved     As Byte  ' Not used
      bIDEDeviceMap As Byte  ' Bit map of IDE devices
      fCapabilities As Long  ' Bit mask of driver capabilities
      dwReserved(3) As Long  ' For future use
  End Type

  ' IDE registers
  Private Type IDEREGS
      bFeaturesReg     As Byte  ' Used for specifying SMART "commands"
      bSectorCountReg  As Byte  ' IDE sector count register
      bSectorNumberReg As Byte  ' IDE sector number register
      bCylLowReg       As Byte  ' IDE low order cylinder value
      bCylHighReg      As Byte  ' IDE high order cylinder value
      bDriveHeadReg    As Byte  ' IDE drive/head register
      bCommandReg      As Byte  ' Actual IDE command
      bReserved        As Byte  ' reserved for future use - must be zero
  End Type

  ' SENDCMDINPARAMS contains the input parameters for the
  ' Send Command to Drive function
  Private Type SENDCMDINPARAMS
      cBufferSize   As Long     ' Buffer size in bytes
      irDriveRegs   As IDEREGS  ' Structure with drive register values.
      bDriveNumber  As Byte     ' Physical drive number to send command to (0,1,2,3).
      bReserved(2)  As Byte     ' Bytes reserved
      dwReserved(3) As Long     ' DWORDS reserved
      bBuffer()     As Byte     ' Input buffer.
  End Type

  ' Status returned from driver
  Private Type DRIVERSTATUS
      bDriverError  As Byte     ' Error code from driver, or 0 if no error
      bIDEStatus    As Byte     ' Contents of IDE Error register
                                ' Only valid when bDriverError is SMART_IDE_ERROR
      bReserved(1)  As Byte
      dwReserved(1) As Long
  End Type

  ' Structure returned by SMART IOCTL commands
  Private Type SENDCMDOUTPARAMS
      cBufferSize  As Long          ' Size of Buffer in bytes
      DRIVERSTATUS As DRIVERSTATUS  ' Driver status structure
      bBuffer()    As Byte          ' Buffer of arbitrary length for data read from drive
  End Type

  Private Type IDSECTOR
      wGenConfig                 As Integer
      wNumCyls                   As Integer
      wReserved                  As Integer
      wNumHeads                  As Integer
      wBytesPerTrack             As Integer
      wBytesPerSector            As Integer
      wSectorsPerTrack           As Integer
      wVendorUnique(2)           As Integer
      sSerialNumber(19)          As Byte
      wBufferType                As Integer
      wBufferSize                As Integer
      wECCSize                   As Integer
      sFirmwareRev(7)            As Byte
      sModelNumber(39)           As Byte
      wMoreVendorUnique          As Integer
      wDoubleWordIO              As Integer
      wCapabilities              As Integer
      wReserved1                 As Integer
      wPIOTiming                 As Integer
      wDMATiming                 As Integer
      wBS                        As Integer
      wNumCurrentCyls            As Integer
      wNumCurrentHeads           As Integer
      wNumCurrentSectorsPerTrack As Integer
      ulCurrentSectorCapacity    As Long
      wMultSectorStuff           As Integer
      ulTotalAddressableSectors  As Long
      wSingleWordDMA             As Integer
      wMultiWordDMA              As Integer
      bReserved(127)             As Byte
  End Type

  Private Type DISK_GEOMETRY
      Cylinders         As Currency  ' LARGE_INTEGER (8 bytes)
      MediaType         As Long
      TracksPerCylinder As Long
      SectorsPerTrack   As Long
      BytesPerSector    As Long
  End Type

  Private Type MBR
      Signature As Byte  ' Specifies the signature value, which uniquely
                         ' identifies the disk. The Mbr member of the union
                         ' is used to specify the disk signature data for
                         ' a disk formatted with a Master Boot Record (MBR)
                         ' format partition table. Any other value indicates
                         ' that the partition is not a boot partition.  This
                         ' member is valid when PartitionStyle is PARTITION_STYLE_MBR.
      CheckSum  As Long  ' Specifies the checksum for the master boot record.
                         ' The Mbr member of the union is used to specify the
                         ' disk signature data for a disk formatted with a
                         ' Master Boot Record (MBR) format partition table.
                         ' This member is valid when PartitionStyle is PARTITION_STYLE_MBR.
  End Type

  Private Type GUID
      DiskId As Long     ' Specifies the GUID that uniquely identifies the
                         ' disk. The Gpt member of the union is used to
                         ' specify the disk signature data for a disk formatted
                         ' with a GUID Partition Table (GPT) format partition
                         ' table. This member is valid when PartitionStyle is
                         ' PARTITION_STYLE_GPT. The GUID data type is described
                         ' on the LPGUID reference page.
  End Type

  Private Type DISK_PARTITION_INFO
      SizeOfPartitionInfo As Long   ' Contains the size in bytes of the partition.
      PartitionStyle      As Long   ' Takes a PARTITION_STYLE enumerated value that
                                    ' specifies type of partition table the disk contains.
      MBRFIX              As MBR
      Gpt                 As GUID
  End Type


  Private Type DISK_GEOMETRY_EX
      Geometry As DISK_GEOMETRY
      DiskSize As Currency             ' LARGE_INTEGER (8 bytes) prior to formatting
      Data     As DISK_PARTITION_INFO  ' Data(0 To 56) As BytePointer to a
                                       ' variable length area containing a
                                       ' DISK_PARTITION_INFO structure
                                       ' followed by a DISK_DETECTION_INFO
                                       ' structure.
  End Type

' ***************************************************************************
' API Declares (Module)
' ***************************************************************************
  ' The CopyMemory function copies a block of memory from one location to
  ' another. For overlapped blocks, use the MoveMemory function.
  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
          (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

  ' PathFileExists function determines whether a path to a file system
  ' object such as a file or directory is valid.  Returns nonzero if
  ' file exists.
  Private Declare Function PathFileExists Lib "shlwapi" _
          Alias "PathFileExistsA" (ByVal pszPath As String) As Long

  ' The FormatMessage function formats a message string. The function
  ' requires a message definition as input. The message definition can
  ' come from a buffer passed into the function. It can come from a
  ' message table resource in an already-loaded module. Or the caller
  ' can ask the function to search the system’s message table resource(s)
  ' for the message definition. The function finds the message definition
  ' in a message table resource based on a message identifier and a
  ' language identifier. The function copies the formatted message text
  ' to an output buffer, processing any embedded insert sequences if requested.
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
          (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, _
          ByVal dwLanguageId As Long, ByVal lpBuffer As String, _
          ByVal nSize As Long, Arguments As Any) As Long

  ' The CreateFile function creates or opens the following objects and
  ' returns a handle that can be used to access the object:
  '
  '       Files
  '       pipes
  '       mailslots
  '       communications resources
  '       disk devices (Windows NT only)
  '       consoles
  '       directories (open only)
  Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
          (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
          ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, _
          ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
          ByVal hTemplateFile As Long) As Long

  ' CloseHandle invalidates the specified object handle, decrements
  ' the object’s handle count, and performs object retention checks.
  ' Once the last handle to an object is closed, the object is removed
  ' from the operating system.
  Private Declare Function CloseHandle Lib "kernel32" _
          (ByVal hObject As Long) As Long

  ' The DeviceIoControl function sends a control code directly to a
  ' specified device driver, causing the corresponding device to
  ' perform the specified operation.
  Private Declare Function DeviceIoControl Lib "kernel32" _
          (ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
          lpInBuffer As Any, ByVal nInBufferSize As Long, _
          lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
          lplngBytesRet As Long, lpOverlapped As Any) As Long

  ' The GetLogicalDriveStrings function fills a buffer with strings that
  ' specify valid drives in the system.
  Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
          Alias "GetLogicalDriveStringsA" _
          (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

  ' The GetDriveType function determines whether a disk drive is a removable,
  ' fixed, CD-ROM, RAM disk, or network drive.
  Private Declare Function GetDriveType Lib "kernel32" _
          Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

  ' retrieves information describing a disk volume and the file system it
  ' uses. This information includes things such as the volume label and the
  ' disk's serial number.
  Private Declare Function GetVolumeInformation Lib "kernel32" _
          Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
          ByVal lpDrvVolNameBuffer As String, ByVal nDrvVolNameSize As Long, _
          lpDrvVolSerialNumber As Long, lpMaximumComponentLength As Long, _
          lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
          ByVal nFileSystemNameSize As Long) As Long

  ' The GetDiskFreeSpace function retrieves information about the
  ' specified disk, including the amount of free space on the disk
  Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
          (ByVal lpRootPathName As String, lpDrvSectorsPerCluster As Long, _
          lpDrvBytesPerSector As Long, lpNumberOfDrvFreeClusters As Long, _
          lpTtoalNumberOfClusters As Long) As Long

  ' The GetDiskFreeSpaceEx function obtains information about the
  ' amount of space available on a disk volume: the total amount of
  ' space, the total amount of free space, and the total amount of
  ' free space available to the user associated with the calling thread.
  Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
          (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, _
          lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

  ' This function obtains extended information about the version of the OS
  ' that is currently running. CeGetVersionEx is the RAPI version of this
  ' function.  Nonzero indicates success. Zero indicates failure.
  Private Declare Function GetVersionEx Lib "kernel32" _
          Alias "GetVersionExA" (LpVersionInformation As Any) As Long

  ' ZeroMemory fills a location in memory with zeros. The function sets
  ' each byte starting at the given memory location to zero. The memory
  ' location is identified by a pointer to the memory address.
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" _
          (ByRef Destination As Any, ByVal Length As Long)

  ' RegOpenKeyEx function opens the specified key. Returns zero if successful.
  Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
          Alias "RegOpenKeyExA" (ByVal hKey As Long, _
          ByVal lpSubKey As String, ByVal ulOptions As Long, _
          ByVal samDesired As Long, phkResult As Long) As Long
  
  ' RegQueryValueEx function retrieves the type and data for a specified value
  ' name associated with an open registry key.
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
          Alias "RegQueryValueExA" (ByVal hKey As Long, _
          ByVal lpValueName As String, ByVal lpReserved As Long, _
          lpType As Long, lpData As Any, lpcbData As Long) As Long
  
  ' RegCloseKey function releases the handle of the specified key.
  Private Declare Function RegCloseKey Lib "advapi32.dll" _
          (ByVal hKey As Long) As Long
  
' ***************************************************************************
' Module Variables
'
'                    +-------------- Module level designator
'                    |  +----------- Data type (Type Structure)
'                    |  |     |----- Variable subname
'                    - --- ---------
' Naming standard:   m typ DG
' Variable name:     mtypDG
'
' ***************************************************************************
  Private mtypDG                   As DISK_GEOMETRY_EX
  Private mstrDriveType            As String
  Private mstrMfgModel             As String
  Private mstrMfgSerial            As String
  Private mstrUSBSerial            As String
  Private mstrDrvVolName           As String
  Private mstrMfgFirmware          As String
  Private mstrDrvVolSerial         As String
  Private mstrDrvFileSystem        As String
  Private mstrPartitionData        As String
  Private mstrDriveTypeExtra       As String
  Private mstrDrvFormattedSize     As String
  Private mlngDrvFreeClusters      As Long
  Private mlngDrvTotalClusters     As Long
  Private mlngDrvBytesPerSector    As Long
  Private mlngDrvBytesPerCluster   As Long
  Private mlngDrvSectorsPerCluster As Long
  Private mcurDrvUsedSpace         As Currency
  Private mcurDrvFreeSpace         As Currency
  Private mcurTotalDrvSpace        As Currency
  Private mcurDrvAvailableSpace    As Currency
  
  ' Disk geometry (not partition)
  Private mstrDiskFormattedSize    As String
  Private mlngDiskTracksPerCyl     As Long
  Private mlngDiskBytesPerSector   As Long
  Private mlngDiskSectorsPerTrack  As Long
  Private mcurDiskCylinders        As Currency
  Private mcurTotalDiskSpace       As Currency
  
' ***************************************************************************
' ****                      Properties                                   ****
' ***************************************************************************

Public Property Let StopProcessing(ByVal blnData As Boolean)
    
    ' used when assigning a value to the property, on the left
    ' side of an assignment.    Syntax:  X.StopProcessing = True
    gblnStopProcessing = blnData
    DoEvents

End Property

Public Property Get StopProcessing() As Boolean
    
    ' used when retrieving value of a property, on the right
    ' side of an assignment.   Syntax:  Debug.Print X.StopProcessing
    StopProcessing = gblnStopProcessing
    DoEvents

End Property

' ===========================================================================
' Disk information (whole disk) - Disk geometry -  Read-only
' ===========================================================================
Public Property Get TotalDiskSpace() As Currency
    TotalDiskSpace = mcurTotalDiskSpace
End Property

Public Property Get DiskCylinders() As Currency
    DiskCylinders = mcurDiskCylinders
End Property

Public Property Get DiskTracksPerCyl() As Long
    DiskTracksPerCyl = mlngDiskTracksPerCyl
End Property

Public Property Get DiskSectorsPerTrack() As Long
    DiskSectorsPerTrack = mlngDiskSectorsPerTrack
End Property

Public Property Get DiskBytesPerSector() As Long
    DiskBytesPerSector = mlngDiskBytesPerSector
End Property

Public Property Get DiskFormattedSize() As String
    DiskFormattedSize = mstrDiskFormattedSize
End Property

' ===========================================================================
' Manufacturer data - Cannot be changed - Read-only
' ===========================================================================
Public Property Get MfgSerial() As String
    MfgSerial = mstrMfgSerial
End Property

Public Property Get MfgModel() As String
    MfgModel = mstrMfgModel
End Property

Public Property Get MfgFirmware() As String
    MfgFirmware = mstrMfgFirmware
End Property

' ===========================================================================
' Drive information (partition) Read-only
' ===========================================================================
Public Property Get DriveType() As String
    DriveType = mstrDriveType
End Property

Public Property Get DriveTypeExtra() As String
    DriveTypeExtra = mstrDriveTypeExtra
End Property

Public Property Get DrvFormattedSize() As String
    DrvFormattedSize = mstrDrvFormattedSize
End Property

Public Property Get PartitionData() As String
    PartitionData = mstrPartitionData
End Property

Public Property Get DrvVolName() As String
    DrvVolName = mstrDrvVolName
End Property

Public Property Get DrvVolSerial() As String
    DrvVolSerial = mstrDrvVolSerial
End Property

Public Property Get DrvFileSystem() As String
    DrvFileSystem = mstrDrvFileSystem
End Property

Public Property Get DrvBytesPerSector() As Long
    DrvBytesPerSector = mlngDrvBytesPerSector
End Property

Public Property Get DrvSectorsPerCluster() As Long
    DrvSectorsPerCluster = mlngDrvSectorsPerCluster
End Property

Public Property Get DrvBytesPerCluster() As Long
    DrvBytesPerCluster = mlngDrvBytesPerCluster
End Property

Public Property Get DrvFreeClusters() As Long
    DrvFreeClusters = mlngDrvFreeClusters
End Property

Public Property Get DrvTotalClusters() As Long
    DrvTotalClusters = mlngDrvTotalClusters
End Property

Public Property Get DrvFreeSpace() As Currency
    DrvFreeSpace = mcurDrvFreeSpace
End Property

Public Property Get DrvAvailableSpace() As Currency
    DrvAvailableSpace = mcurDrvAvailableSpace
End Property

Public Property Get DrvUsedSpace() As Currency
    DrvUsedSpace = mcurDrvUsedSpace
End Property

Public Property Get TotalDrvSpace() As Currency
    TotalDrvSpace = mcurTotalDrvSpace
End Property


' ***************************************************************************
' ****                      Methods                                      ****
' ***************************************************************************

' ***************************************************************************
' Routine:       GetDriveInfo
'
' Description:   Captures information about a specific drive and updates
'                appropriate property values.
'
' Parameters:    strDrive - drive to be queried
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Sub GetDriveInfo(ByVal strDrive As String)

    Dim lngDiskNo     As Long      ' Physical disk number
    Dim lngMediaType  As Long
    Dim curTotalSpace As Currency
    
    Const ROUTINE_NAME As String = "GetDriveInfo"

    On Error GoTo GetDriveInfo_Error

    ' Verify data was passed
    If Len(Trim$(strDrive)) = 0 Then
        Exit Sub
    End If
    
    ResetVariables  ' Initialize module variables
    
    strDrive = QualifyPath(Left$(strDrive, 1), ":")   ' Format drive letter
    mstrDriveType = GetDriveDescription(strDrive)     ' Partition description
    
    ' Test for CD-Rom, DVD or Virtual CD
    If StrComp("DVD", Mid$(mstrDriveType, 1, 3), vbBinaryCompare) = 0 Or _
       StrComp("CD-ROM", Mid$(mstrDriveType, 1, 6), vbBinaryCompare) = 0 Or _
       StrComp("Virtual", Mid$(mstrDriveType, 1, 7), vbBinaryCompare) = 0 Then
       
        ' If this is a CD or DVD drive
        ' then exit this routine because
        ' no more data is available
        GoTo GetDriveInfo_CleanUp
        
    End If
           
    GetDiskSpaceInfo strDrive   ' Partition space information
    GetVolumeInfo strDrive      ' Partition volume information
    GetMediaInfo strDrive       ' Capture disk geometry
    
    ' Never pass a module level variable
    ' as a parameter within its own module.
    ' Pointers may become confused and
    ' erroneous data can occur.
    curTotalSpace = mcurTotalDrvSpace                    ' Total partition space (ex:  75,231,309,824)
    mstrDrvFormattedSize = DisplayNumber(curTotalSpace)  ' Formatted partition space (ex:  70.1 GB)
    lngDiskNo = GetDiskNumber(strDrive)                  ' Get disk number assigned by BIOS (0-7)
    GetMfrInfo lngDiskNo                                 ' Get disk manufacturer info (Model, Serial, Firmware rev.)
    
    ' Is operating system Windows XP or newer?
    If IsWinXPorNewer Then
    
        ' If additional description of device has
        ' not been collected then get the data
        If Len(Trim$(mstrDriveTypeExtra)) = 0 Then
            lngMediaType = mtypDG.Geometry.MediaType  ' (0-12 valid value)
            mstrDriveTypeExtra = GetDriveTypeExtra(lngMediaType)
        End If
        
        ' Per whole disk totals (not partition)
        mcurTotalDiskSpace = CCur(mtypDG.DiskSize * TEN_THOUSAND)           ' Total space (ex:  80,026,361,856)
        mcurDiskCylinders = CCur(mtypDG.Geometry.Cylinders * TEN_THOUSAND)  ' Total cylinder count
        mlngDiskTracksPerCyl = mtypDG.Geometry.TracksPerCylinder            ' Total tracks per cylinder
        mlngDiskSectorsPerTrack = mtypDG.Geometry.SectorsPerTrack           ' Total sectors per track
        mlngDiskBytesPerSector = mtypDG.Geometry.BytesPerSector             ' Total bytes per sector
        curTotalSpace = mcurTotalDiskSpace                                  ' Total total disk space (ex:  80,026,361,856)
        mstrDiskFormattedSize = DisplayNumber(curTotalSpace)                ' Formatted total disk space (ex:  74.5 GB)
        
    End If
       
GetDriveInfo_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetDriveInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume GetDriveInfo_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetDiskNumber
'
' Description:   Use WMI (Windows Management Instrumentation) to obtain a
'                drive's partition information and the disk number as
'                assigned by the BIOS.  See enumIDE_DRIVE_NUMBER.
'
' Parameters:    strDrive - Drive letter to be queried
'
' Returns:       Disk number (0-7)
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Public Function GetDiskNumber(ByVal strDrive As String) As Long

    ' Called by GetDriveInfo()
    
    Dim blnFoundIt      As Boolean
    Dim strQuery        As String
    Dim strDiskNo       As String   ' Disk number
    Dim strUSBExtra     As String   ' USB enhanced description
    Dim strUSBSerialNo  As String   ' USB serial number
    Dim objWMI          As Object
    Dim objDrive        As Object
    Dim objDrives       As Object
    Dim objPartition    As Object
    Dim objPartitions   As Object
    Dim objLogicalDisk  As Object
    Dim objLogicalDisks As Object
    
    On Error Resume Next
        
    Set objWMI = Nothing   ' Verify objects are not active
    Set objDrive = Nothing
    Set objDrives = Nothing
    Set objPartition = Nothing
    Set objPartitions = Nothing
    Set objLogicalDisk = Nothing
    Set objLogicalDisks = Nothing
    
    blnFoundIt = False                                                ' Preset drive flag to FALSE
    strDrive = QualifyPath(Left$(strDrive, 1), ":")                   ' Format drive letter
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")                 ' Instantiate WMI object
    Set objDrives = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive") ' Get list of all physical drives
    
    ' Parse each physical drive
    For Each objDrive In objDrives
        
        ' Use this data if USB device
        strUSBExtra = objDrive.Caption
        strUSBSerialNo = objDrive.PNPDeviceID
        
        ' Associate partitions to physical drives
        strQuery = vbNullString
        strQuery = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & objDrive.DeviceID & _
                   "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        
        ' Get list of all partitions on this physical drive
        Set objPartitions = objWMI.ExecQuery(strQuery)
        
        ' Parse each partition
        For Each objPartition In objPartitions
            
            ' Associate logical disks to partitions
            strQuery = vbNullString
            strQuery = "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & objPartition.DeviceID & _
                       "'} WHERE AssocClass = Win32_LogicalDiskToPartition"
        
            ' Get list of all logical disks on this partition
            ' (ex:  C:, D:, etc.)
            Set objLogicalDisks = objWMI.ExecQuery(strQuery)
            
            ' Parse each logical disk
            For Each objLogicalDisk In objLogicalDisks
            
                ' Verify this is the drive we are looking for.
                '
                ' Example:  objLogicalDisk.DeviceID = "C:"
                '           objPartition.Caption = "Disk #0, Partition #0"
                If StrComp(Left$(objLogicalDisk.DeviceID, 2), strDrive, vbTextCompare) = 0 Then
                
                    mstrPartitionData = objPartition.Caption   ' Save Partition info
                    
                    ' Capture disk number assigned by BIOS (0-7)
                    ' Search for first comma then back up one character
                    strDiskNo = Mid$(mstrPartitionData, InStr(1, mstrPartitionData, Chr$(44)) - 1, 1)
                    
                    ' Validate data
                    If IsNumeric(strDiskNo) Then
                        Select Case Val(strDiskNo)
                               Case 0 To 7
                                    ' Good data
                                    GetDiskNumber = Val(strDiskNo)
                                    
                                    ' If USB device then save manufacturer
                                    ' enhanced description and serial number
                                    If objLogicalDisk.DriveType = 2 Then
                                        mstrDriveTypeExtra = strUSBExtra
                                        mstrUSBSerial = FormatSerialNo(strUSBSerialNo)
                                    End If
        
                               Case Else
                                    ' Bad data (invalid number)
                                    GetDiskNumber = -1
                                    mstrPartitionData = vbNullString
                        End Select
                    Else
                        ' Bad data (not numeric)
                        GetDiskNumber = -1
                        mstrPartitionData = vbNullString
                    End If
                    
                    blnFoundIt = True  ' Found matching drive letter
                    Exit For           ' Exit FOR..NEXT loop
                
                End If
                
            Next objLogicalDisk
            
            If blnFoundIt Then
                Exit For   ' Exit FOR..NEXT loop
            End If
            
        Next objPartition
        
        If blnFoundIt Then
            Exit For   ' Exit FOR..NEXT loop
        End If
        
    Next objDrive

    strQuery = vbNullString
    Set objLogicalDisks = Nothing   ' Always free objects from memory when not needed
    Set objLogicalDisk = Nothing
    Set objPartitions = Nothing
    Set objPartition = Nothing
    Set objDrives = Nothing
    Set objDrive = Nothing
    Set objWMI = Nothing
    
    On Error GoTo 0   ' Nullify this error trap

End Function

' ***************************************************************************
' Routine:       GetDriveDescription
'
' Description:   Make an API call to a specific drive and determine type
'                of device
'
' Parameters:    strDrive - Drive letter to be queried
'
' Returns:       Brief desription of drive (ex: "Remote (network) drive")
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Public Function GetDriveDescription(ByVal strDrive As String) As String
Attribute GetDriveDescription.VB_Description = "Capture the volume description.  Example:  Physical hard drive"

    ' Called by GetDriveInfo()
    
    Dim lngDriveType As enumDriveType

    Const ROUTINE_NAME As String = "GetDriveDescription"

    On Error GoTo GetDriveDescription_Error

    GetDriveDescription = vbNullString
    strDrive = QualifyPath(Left$(strDrive, 1), ":")   ' Format drive letter
    lngDriveType = GetDriveType(strDrive)             ' Get drive number constant

    ' Test for the type of drive
    Select Case lngDriveType

           Case eUnknown       ' 0
                GetDriveDescription = "Drive type cannot be determined"

           Case eBadRoot       ' 1
                GetDriveDescription = "Root directory does not exist"

           Case eRemovable     ' 2

                Select Case LCase$(Left$(strDrive, 1))
                       Case "a", "b"
                            GetDriveDescription = "Floppy drive"
                       Case Else
                            GetDriveDescription = "USB drive"
                End Select

           Case eFixed         ' 3
                GetDriveDescription = "Physical hard drive"

           Case eNetwork       ' 4
                GetDriveDescription = "Remote (network) drive"

           Case eCDRom         ' 5
                ' a default description in case the call below fails
                GetDriveDescription = "Optical Drive (CD or DVD)"

                ' now test specifics
                Select Case IsCDRomDrive(strDrive)
                       Case FILE_DEVICE_CD_ROM:       GetDriveDescription = "CD-ROM drive"
                       Case FILE_DEVICE_DVD:          GetDriveDescription = "DVD drive"
                       Case FILE_DEVICE_VIRTUAL_DISK: GetDriveDescription = "Virtual CD drive"
                End Select

           Case eRamdisk       ' 6
                GetDriveDescription = "Ramdisk"
   End Select

GetDriveDescription_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GetDriveDescription_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    GetDriveDescription = vbNullString
    Resume GetDriveDescription_CleanUp

End Function

Public Function IsCDRomDrive(ByVal strDrive As String) As Long
Attribute IsCDRomDrive.VB_Description = "Calls GetMediaType() to capture the media identifier code."

    ' Called by GetDriveInfo()
    
    Const ROUTINE_NAME As String = "IsCDRomDrive"

    On Error GoTo IsCDRomDrive_Error

    strDrive = QualifyPath(Left$(strDrive, 1), ":")   ' Format drive letter
    IsCDRomDrive = GetDriveType(strDrive)             ' Capture drive type
    
    ' if the drive type returned indicates a
    ' CD (DRIVE_CDROM) perform a DeviceIoControl
    ' test to determine if a CD or a DVD.
    If IsCDRomDrive = enumDriveType.eCDRom Then
     
        ' See if this drive is physically
        ' mounted or just a drive letter
        If IsCDMounted(strDrive) Then
            ' Obtain DeviceIoControl media
            ' code for the drive
            IsCDRomDrive = GetMediaType(strDrive)
        Else
            IsCDRomDrive = &H24  ' Virtual drive
        End If
    Else
        IsCDRomDrive = 0
    End If

IsCDRomDrive_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

IsCDRomDrive_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume IsCDRomDrive_CleanUp
    
End Function

Public Function IsCDMounted(ByVal strDrive As String) As Boolean

    ' 1. Read registry    Path - HKEY_LOCAL_MACHINE\SYSTEM\MountedDevices
    '                     Subkey - \DosDevices\E:   <- drive letter is dynamic
    ' 2. Verify a specific drive exist
    ' 3. Read binary data to determine if this is a physically mounted CD-Rom drive
    '
    ' Called by IsCDRomDrive()
    
    Dim hKey       As Long    ' Hive handle
    Dim lngIdx     As Long    ' Loop index
    Dim lngDummy   As Long    ' Just a temp hold variable
    Dim lngBuffer  As Long    ' Length of subkey data
    Dim hPathValue As Long    ' Hive path value (must return greater than zero)
    Dim strData    As String  ' Formatted data
    Dim strSubKey  As String  ' Complete subkey name
    Dim abytData() As Byte    ' Hold binary data
    
    Const REG_PATH As String = "SYSTEM\MountedDevices"
    Const SUB_KEY  As String = "\DosDevices\"

    On Error GoTo IsCDMounted_Error

    IsCDMounted = False      ' Preset flag to FALSE
    Erase abytData()         ' Always start with an empty array
    strData = vbNullString   ' Verify variable is empty
    hPathValue = 0           ' Init return value
    lngBuffer = 0
    
    strDrive = UCase$(QualifyPath(Left$(strDrive, 1), ":"))   ' Format drive letter
    hKey = HKEY_LOCAL_MACHINE                                 ' Registry hive to be accessed
    strSubKey = SUB_KEY & strDrive                            ' Prepare subkey name - ex:  \DosDevices\E:
    
    ' Does registry path exist?  If so,
    ' hPathValue will be greater than zero.
    Call RegOpenKeyEx(hKey, REG_PATH, 0&, KEY_READ, hPathValue)
        
    If hPathValue > 0 Then
    
        ' Does subkey name exist?  If so,
        ' lngBuffer will be greater than zero.
        Call RegQueryValueEx(hPathValue, strSubKey, 0&, lngDummy, 0&, lngBuffer)
        
        ' Test buffer for a successful read
        If lngBuffer > 0 Then
            
            ReDim abytData(lngBuffer - 1)   ' Size receiving byte array
            
            ' Capture binary data from registry
            ' subkey and store in a byte array
            Call RegQueryValueEx(hPathValue, strSubKey, 0&, REG_BINARY, abytData(0), lngBuffer)
            
            strData = StrConv(abytData(), vbUnicode)   ' Convert byte array to string
            strData = Replace(strData, Chr$(0), "")    ' Remove all null values
            
            ' Verify some data was captured
            If Len(strData) > 0 Then
                
                ' Search data string
                For lngIdx = 1 To Len(strData) - 4
                      
                    ' Look for word "CDRom"
                    If StrComp(Mid$(strData, lngIdx, 5), "cdrom", vbTextCompare) = 0 Then
                        IsCDMounted = True   ' Found it
                        Exit For             ' exit FOR..NEXT loop
                    End If
                
                Next lngIdx
            End If
            
        End If
    End If
    
IsCDMounted_CleanUp:
    Erase abytData()         ' Always empty arrays when not needed
    RegCloseKey hPathValue   ' Close any open keys
    RegCloseKey hKey

    On Error GoTo 0   ' Nullify this error trap
    Exit Function

IsCDMounted_Error:
    Err.Clear
    Resume IsCDMounted_CleanUp

End Function

' ***************************************************************************
' Routine:       GetClusterSize
'
' Description:   Determine cluster size of a specific partition.
'
' Parameters:    strDrive - Drive letter to be queried
'
' Returns:       Cluster size
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Public Function GetClusterSize(ByVal strDrive As String) As Long

    ' Called by GetDriveInfo()
    
    Const ROUTINE_NAME As String = "GetClusterSize"

    On Error GoTo GetClusterSize_Error

    GetClusterSize = 0
    GetDiskSpaceInfo strDrive
    GetClusterSize = mlngDrvBytesPerCluster

GetClusterSize_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GetClusterSize_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume GetClusterSize_CleanUp
    
End Function
                         
' ***************************************************************************
' Routine:       GetDriveLetters
'
' Description:   Capture a list of all available drive letters
'                (Ex: A:\, C:\, D:\, etc.)
'
' Returns:       String array filled with drive letters
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 21-Jul-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' 12-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated parsing of drive letters
' ***************************************************************************
Public Function GetDriveLetters() As String()

    Dim avntDrives As Variant
    Dim lngIndex   As Long
    Dim lngCount   As Long
    Dim lngLength  As Long
    Dim strDrvList As String
    Dim astrTemp() As String
    
    avntDrives = Empty          ' Always start with empty variants
    Erase astrTemp()            ' Always start with empty arrays
    strDrvList = vbNullString   ' Start with empty string
    lngLength = 0
    lngCount = 0
    
    ' Call API with an empty string. This will
    ' return length required for data string.
    lngLength = GetLogicalDriveStrings(Len(strDrvList), strDrvList)
    
    If lngLength > 0 Then
        
        strDrvList = Space$(lngLength)                ' Preload with spaces
        GetLogicalDriveStrings lngLength, strDrvList  ' Call API again to load string with drive letters
        avntDrives = Split(strDrvList, Chr$(0))       ' Transfer drive letters to an array
        
        ReDim astrTemp(26)   ' Size return array to hold upto 26 drive letters
        
        ' Loop thru list of drive letters
        ' and transfer to string array
        For lngIndex = 0 To UBound(avntDrives) - 1
            
            If Len(Trim$(avntDrives(lngIndex))) > 0 Then
                astrTemp(lngIndex) = CStr(avntDrives(lngIndex))
                lngCount = lngCount + 1   ' Increment counter
            End If
            
        Next lngIndex
    
    End If
    
    ReDim Preserve astrTemp(lngCount)  ' Size array to what was used
    GetDriveLetters = astrTemp()       ' Return array of drive letters
    
    avntDrives = Empty   ' Always empty variants when not needed
    Erase astrTemp()     ' Alwyas empty arrays when not needed
    
End Function

' ***************************************************************************
' Routine:       GetVolumeInfo
'
' Description:   Make an API call to a specific drive and capture the volume
'                information.  This data is created during formatting.
'
' Parameters:    strDrive   - Drive letter to be queried
'                strFileSys - Optional - Returns file system type
'                             Ex:  FAT32, NTFS, HPFS, CDFS (CD-Rom File System)
'                strVolName - Optional - Returns volume label
'                strVolSN   - Optional - Returns volume serial number
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 15-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              - Corrected formatting of volume serial number
'              - Replaced multiple IF statements with a CASE statement
' 15-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Made return parameters optional
' ***************************************************************************
Public Sub GetVolumeInfo(ByVal strDrive As String, _
                Optional ByRef strFileSys As String = vbNullString, _
                Optional ByRef strVolName As String = vbNullString, _
                Optional ByRef strVolSerial As String = vbNullString)
Attribute GetVolumeInfo.VB_Description = "Capture the volume information."

    ' Called by GetDriveInfo()
    '           GetClusterSize()
    
    Dim lngVolSn    As Long
    Dim lngHiWord   As Long
    Dim lngLoWord   As Long
    Dim lngSysFlags As Long

    Const ROUTINE_NAME As String = "GetVolumeInfo"

    On Error GoTo GetVolumeInfo_Error

    ' Use spaces and not null values.  Reduces your errors.
    strVolName = Space$(MAX_SIZE)
    strFileSys = Space$(MAX_SIZE)
    strDrive = Left$(strDrive, 1) & ":\"   ' Format drive letter

    ' Make API call to capture volume information
    If GetVolumeInformation(strDrive, strVolName, MAX_SIZE, _
                            lngVolSn, 0, lngSysFlags, _
                            strFileSys, MAX_SIZE) <> 0 Then

        strVolName = TrimStr(strVolName)   ' save volume label (name)

        If Len(Trim$(strVolName)) = 0 Then
            strVolName = "<No_Label>"
        End If

        strFileSys = TrimStr(Trim$(strFileSys))   ' Save file system type
        GetHiLoWord lngVolSn, lngLoWord, lngHiWord            ' Process both sides of volume serial number
        
        ' Format volume serial number (0311-B12C)
        strVolSerial = UCase$(Right$("0000" & Hex$(lngHiWord), 4) & "-" & _
                              Right$("0000" & Hex$(lngLoWord), 4))
    End If

GetVolumeInfo_CleanUp:
    mstrDrvVolName = strVolName
    mstrDrvVolSerial = strVolSerial
    mstrDrvFileSystem = strFileSys
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetVolumeInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strVolName = vbNullString
    strVolSerial = vbNullString
    strFileSys = vbNullString
    Resume GetVolumeInfo_CleanUp

End Sub

' ***************************************************************************
' Routine:       SpecificTypeOfDrive
'
' Description:   Determine the type of drive we are looking for.
'                Return codes are:
'
'                    Public Enum enumDriveType
'                        eUnknown         ' 0 Unknown drive type
'                        eBadRoot         ' 1 No root directory
'                        eRemovable       ' 2 Floppy or flash drive
'                        eFixed           ' 3 Local hard drive
'                        eNetwork         ' 4 Shared Network drive
'                        eCDRom           ' 5 CD-Rom drive (DVD or other)
'                        eRamdisk         ' 6 Virtual memory disk
'                    End Enum
'
' Parameters:    strDrive - drive to be queried (ex:  C:)
'                lngTypeNeeded - Type of drive needed
'                lngTypeOfRemovable - Optional - Type of removable drive.
'                   This parameter is ignored if not a removable drive.
'                            0 = eFloppy
'                            1 = eUSB_Drive
'                            2 = eBoth types
'                   Default = eFloppy
'
' Returns:       TRUE if or drive is found else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 03-MAR-2000  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated to distinquish between floppy and flash drives
' ***************************************************************************
Public Function SpecificTypeOfDrive(ByVal strDrive As String, _
                                    ByVal lngTypeNeeded As enumDriveType, _
                           Optional ByVal lngTypeOfRemovable As enumTypeOfRemovable = eFloppy) As Boolean

    Dim strDescription As String
    Dim lngDriveType   As enumDriveType

    Const ROUTINE_NAME As String = "SpecificTypeOfDrive"

    On Error GoTo SpecificTypeOfDrive_Error

    SpecificTypeOfDrive = False  ' Preset return code to FALSE
    
    ' Verify drive data was passed
    If Len(Trim$(strDrive)) = 0 Then
        Exit Function
    End If

    ' Verify drive letter is in proper format
    strDrive = Left$(strDrive, 1) & ":"

    ' Make API call to get drive number constant
    lngDriveType = GetDriveType(strDrive)

    ' Evaluate results of API call
    Select Case lngDriveType
           
           Case eRemovable
                ' Get decription of drive type
                strDescription = GetDriveDescription(strDrive)
            
                Select Case lngTypeOfRemovable
                       
                       Case eFloppy      ' 0
                            ' Verify this is a floppy drive
                            If StrComp(strDescription, "Floppy drive", vbTextCompare) = 0 Then
                                
                                ' Verify floppy drive is accessible
                                If IsDeviceReady(strDrive) Then
                                    SpecificTypeOfDrive = True
                                End If
                            End If
                       
                       Case eUSB_Drive   ' 1
                            ' Verify this is a USB device
                            If StrComp(strDescription, "USB drive", vbTextCompare) = 0 Then
                                
                                ' Verify USB device is accessible
                                If IsDeviceReady(strDrive) Then
                                    SpecificTypeOfDrive = True
                                End If
                            End If
                            
                       Case eBothTypes   ' 2
                            SpecificTypeOfDrive = True
                End Select
    
           Case lngTypeNeeded   ' Is this the device we are looking for?
                SpecificTypeOfDrive = True
    End Select

SpecificTypeOfDrive_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

SpecificTypeOfDrive_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume SpecificTypeOfDrive_CleanUp

End Function

' ***************************************************************************
' Routine:       GetDiskSpaceInfo
'
' Description:   Capture the space information about a selected drive.
'
' Parameters:    strDrive - drive letter to be queried
'                lngChoice - [OPTIONAL] Determine type of space info to
'                           return in bytes.  Default is amount of free space.
'
' Returns:       Drive space information in bytes
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 09-JAN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function GetDiskSpaceInfo(ByVal strDrive As String, _
                        Optional ByVal lngChoice As enumDiskSpace = eFreespace) As Currency
Attribute GetDiskSpaceInfo.VB_Description = "Capture the space information about a selected drive."

    ' Called by GetDriveInfo()
    '           GetClusterSize()
    
    Dim lngDrvFreeClusters      As Long        ' API GetDiskFreeSpace()
    Dim lngDrvTotalClusters     As Long        ' API GetDiskFreeSpace()
    Dim lngDrvBytesPerSector    As Long        ' API GetDiskFreeSpace()
    Dim lngDrvSectorsPerCluster As Long        ' API GetDiskFreeSpace()
    Dim curTotalBytes           As Currency    ' API GetDiskFreeSpaceEx()
    Dim curDrvFreeSpace         As Currency    ' API GetDiskFreeSpaceEx()
    Dim curDrvAvailableSpace    As Currency    ' API GetDiskFreeSpaceEx()

    Const ROUTINE_NAME As String = "GetDiskSpaceInfo"
    
    On Error GoTo GetDiskSpaceInfo_Error

    ' Verify drive letter is in proper format
    strDrive = Left$(strDrive, 1) & ":"

    If HasLargeDriveSupport(strDrive) Then

        ' get drive info using GetDiskFreeSpace
        If GetDiskFreeSpace(strDrive, lngDrvSectorsPerCluster, _
                            lngDrvBytesPerSector, lngDrvFreeClusters, _
                            lngDrvTotalClusters) <> 0 Then
    
            ' capture data
            mlngDrvSectorsPerCluster = lngDrvSectorsPerCluster
            mlngDrvBytesPerSector = lngDrvBytesPerSector
            mlngDrvBytesPerCluster = lngDrvSectorsPerCluster * lngDrvBytesPerSector
            mlngDrvFreeClusters = lngDrvFreeClusters
            mlngDrvTotalClusters = lngDrvTotalClusters

        End If
    
        ' get drive info using GetDiskFreeSpaceEx
        If GetDiskFreeSpaceEx(strDrive, curDrvAvailableSpace, curTotalBytes, curDrvFreeSpace) <> 0 Then

            ' adjust by multiplying the returned value by 10,000
            ' to accommodate for the decimal places the currency
            ' data type returns.
            mcurTotalDrvSpace = curTotalBytes * TEN_THOUSAND                     ' total disk space
            mcurDrvFreeSpace = curDrvFreeSpace * TEN_THOUSAND                    ' free space
            mcurDrvUsedSpace = (curTotalBytes - curDrvFreeSpace) * TEN_THOUSAND  ' Used space
            mcurDrvAvailableSpace = curDrvAvailableSpace * TEN_THOUSAND          ' available space
        
            Select Case lngChoice
                   Case eFreespace:  GetDiskSpaceInfo = mcurDrvFreeSpace       ' free space
                   Case eTotalSize:  GetDiskSpaceInfo = mcurTotalDrvSpace      ' total disk size
                   Case eUsedSpace:  GetDiskSpaceInfo = mcurDrvUsedSpace       ' Used space
                   Case eAvailSpace: GetDiskSpaceInfo = mcurDrvAvailableSpace  ' available space
                   Case Else:        GetDiskSpaceInfo = 0
            End Select
        End If

    ' Does not support drives greater than 2gb
    Else
        ' get drive info using GetDiskFreeSpace
        If GetDiskFreeSpace(strDrive, lngDrvSectorsPerCluster, lngDrvBytesPerSector, _
                            lngDrvFreeClusters, lngDrvTotalClusters) <> 0 Then

            ' capture the data
            mlngDrvSectorsPerCluster = lngDrvSectorsPerCluster
            mlngDrvBytesPerSector = lngDrvBytesPerSector
            mlngDrvBytesPerCluster = lngDrvSectorsPerCluster * lngDrvBytesPerSector
            mlngDrvFreeClusters = lngDrvFreeClusters
            mlngDrvTotalClusters = lngDrvTotalClusters

            ' Calculate space
            mcurTotalDrvSpace = CCur((lngDrvSectorsPerCluster * lngDrvBytesPerSector) * lngDrvTotalClusters)
            mcurDrvFreeSpace = CCur((lngDrvSectorsPerCluster * lngDrvBytesPerSector) * lngDrvFreeClusters)
            mcurDrvUsedSpace = (mcurTotalDrvSpace - mcurDrvFreeSpace)
        
            Select Case lngChoice
                   Case eFreespace:  GetDiskSpaceInfo = mcurDrvFreeSpace   ' free space
                   Case eTotalSize:  GetDiskSpaceInfo = mcurTotalDrvSpace  ' total disk size
                   Case eUsedSpace:  GetDiskSpaceInfo = mcurDrvUsedSpace   ' used space
                   Case eAvailSpace: GetDiskSpaceInfo = mcurDrvFreeSpace   ' space available
                   Case Else:        GetDiskSpaceInfo = 0
            End Select
        
        End If
    End If

GetDiskSpaceInfo_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GetDiskSpaceInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume GetDiskSpaceInfo_CleanUp

End Function

Public Function IsDeviceReady(ByVal strDrive As String) As Boolean
Attribute IsDeviceReady.VB_Description = "Determines if a device is accessible."

    Dim hDevice   As Long   ' device handle
    Dim lngTemp   As Long   ' Number of bytes to return
    Dim lngReturn As Long   ' return value (0=FALSE, nonZero = TRUE)

    Const ROUTINE_NAME As String = "IsDeviceReady"

    On Error GoTo IsDeviceReady_Error

    lngTemp = 0
    lngReturn = 0   ' Preset to FALSE
    
    ' Is operating system Windows XP or newer?
    If IsWinXPorNewer Then

        ' Verify drive letter is in proper format
        strDrive = Left$(strDrive, 1) & ":"

        ' hDevice will return INVALID_HANDLE_VALUE
        ' when no more physical drives are located.
        hDevice = CreateFile("\\.\" & strDrive, _
                             GENERIC_READ, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, OPEN_EXISTING, 0&, 0&)

        ' If a valid handle, pass that to
        ' obtain the physical disk geometry
        If hDevice <> INVALID_HANDLE_VALUE Then

            ' If the operation succeeds and the device media is accessible,
            ' DeviceIoControl returns a nonzero value
            lngReturn = DeviceIoControl(hDevice, IOCTL_STORAGE_CHECK_VERIFY, _
                                        0&, 0&, ByVal 0&, 0&, lngTemp, ByVal 0&)
        End If

        CloseHandle hDevice   ' Always close device handle
    Else
        InfoMsg "This routine is NOT designed to execute on operating" & vbNewLine & _
                "systems earlier than Windows XP." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If


IsDeviceReady_CleanUp:
    IsDeviceReady = CBool(lngReturn)   ' return either TRUE or FALSE
    On Error GoTo 0                    ' Nullify this error trap
    Exit Function

IsDeviceReady_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    lngReturn = 0   ' Set to FALSE
    Resume IsDeviceReady_CleanUp

End Function

Public Sub DeviceLock(ByVal strDrive As String, _
                      ByVal blnLockDevice As Boolean)
Attribute DeviceLock.VB_Description = "Determines if a device is accessible."

    Dim hHandle     As Long
    Dim lngBytesRet As Long
    Dim lngRetValue As Long
    Dim typPMR      As PREVENT_MEDIA_REMOVAL

    Const ROUTINE_NAME As String = "DeviceLock"

    On Error GoTo DeviceLock_Error

    ' Is operating system Windows XP or newer?
    If IsWinXPorNewer Then

        ' Verify drive letter is in proper format
        strDrive = Left$(strDrive, 1) & ":"

        'obtain a handle to the device
        'using the correct device syntax
        hHandle = CreateFile("\\.\" & strDrive, _
                             GENERIC_READ, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, OPEN_EXISTING, 0&, 0&)

        If hHandle <> INVALID_HANDLE_VALUE Then

            'assign the blnLockDevice value to the
            'PREVENT_MEDIA_REMOVAL type
            typPMR.PreventMediaRemoval = CByte(Abs(blnLockDevice))

            'If the operation succeeds,
            'DeviceIoControl returns a nonzero value
            lngRetValue = DeviceIoControl(hHandle, _
                                          IOCTL_STORAGE_MEDIA_REMOVAL, _
                                          typPMR, Len(typPMR), _
                                          ByVal 0&, 0&, lngBytesRet, ByVal 0&)
        End If
    
    Else
        InfoMsg "This routine is NOT designed to execute on operating" & vbNewLine & _
                "systems earlier than Windows XP." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

DeviceLock_CleanUp:
    CloseHandle hHandle
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

DeviceLock_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume DeviceLock_CleanUp

End Sub

' ***************************************************************************
' Routine:       GetMfrInfo
'
' Description:   Capture the hard drive manufacturer model information,
'                hard coded serial number and firmware information.
'
' Parameters:    lngDiskNo - Numeric indicator of which drive to query
'                strSerial - [Optional] (Return) Hard drive serial number
'                strModel - [Optional] (Return) Hard drive model number
'                strFirmWare - [Optional] (Return) Hard drive firmware
'                                revision data
'
' Returns:       Returns disk drive information
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 09-JAN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 15-May-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified output
' 15-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rearranged parameters and their returns as optional
' 15-Dec-2010  Kenneth Ives  kenaso@tx.rr.com
'              Rearranged parameters and their returns as optional
' 31-Jan-2012  Kenneth Ives  kenaso@tx.rr.com
'              Thanks to Alfred Hellmüller for noticing that I truncated
'              the longer hard drive serial numbers as used with SSD-drives.
'              See GetMfrInfo() routine and step through the process.
' ***************************************************************************
Public Sub GetMfrInfo(ByVal lngDiskNo As enumIDE_DRIVE_NUMBER, _
             Optional ByRef strSerial As String = vbNullString, _
             Optional ByRef strModel As String = vbNullString, _
             Optional ByRef strFirmware As String = vbNullString)
                      
    ' Called by GetDriveInfo()
    
    Dim typDrvInfo As DRIVE_INFO
   
    Const ROUTINE_NAME As String = "GetMfrInfo"

    On Error GoTo GetMfrInfo_Error
    
    ZeroMemory typDrvInfo, Len(typDrvInfo)  ' Empty type structure
    
    strModel = vbNullString    ' Verify return variables are empty
    strSerial = vbNullString
    strFirmware = vbNullString
    
    ' Capture physical hard drive information
    typDrvInfo = GetPhysicalDriveInfo(lngDiskNo)
    
    ' Parse information to be returned
    With typDrvInfo
        Select Case .bDriveType
            
               Case 0   ' Unidentified or USB device
                    strSerial = IIf(Len(mstrUSBSerial) = 0, "<Unknown>", mstrUSBSerial)
                    strModel = IIf(Len(mstrUSBSerial) = 0, "<Unknown>", vbNullString)
                    strFirmware = IIf(Len(mstrUSBSerial) = 0, "<Unknown>", vbNullString)
                        
               Case 1   ' Good data
                    strSerial = Trim$(.SerialNumber)
                    strModel = Trim$(.Model)
                    strFirmware = Trim$(.FirmWare)
                        
               Case 2
                    strSerial = "[ATAPI drive - Data not available]"
                    strModel = "[ATAPI drive - Data not available]"
                    strFirmware = "[ATAPI drive - Data not available]"
                        
               Case Else  ' No data to process
        End Select
    End With

GetMfrInfo_CleanUp:
    mstrMfgSerial = strSerial                ' Mfg hard drive serial number
    mstrMfgModel = strModel                  ' Mfg hard drive model identification
    mstrMfgFirmware = strFirmware            ' Mfg hard drive firmware identification
    ZeroMemory typDrvInfo, Len(typDrvInfo)   ' Empty type structure
    
    On Error GoTo 0   ' Nullify this error trap
    Exit Sub

GetMfrInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    strModel = vbNullString
    strSerial = vbNullString
    strFirmware = vbNullString
    Resume GetMfrInfo_CleanUp

End Sub

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
  
    ' Called by GetDriveInfo()
    
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
' Routine:       CreateDirStructure
'
' Description:   Create nested folder structure
'
' Parameters:    strCompletePath - Fully qualified path of complete path
'                                  structure
'
' Returns:       TRUE if successful else FALSE
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 04-Feb-2010  Karl E. Peterson
'              Creating Nested Folders
'              http://visualstudiomagazine.com/articles/2010/02/04/creating-nested-folders.aspx
' 13-Nov-2010  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Public Function CreateDirStructure(ByVal strCompletePath As String) As Boolean
   
    Dim lngBase       As Long
    Dim lngIndex      As Long
    Dim astrFolders() As String
    
    Const ROUTINE_NAME As String = "CreateDirStructure"

    On Error GoTo CreateDirStructure_Error

    Erase astrFolders()                             ' Always start with empty arrays
    strCompletePath = QualifyPath(strCompletePath)  ' Append backslash if needed
    lngBase = 0
    
    ' Split incoming folder into subfolders
    astrFolders() = Split(strCompletePath, "\")
    
    ' Ignore first entry in array since it will
    ' be root of the path (ex:  C: )
    For lngIndex = 1 To UBound(astrFolders) - 1
        astrFolders(lngIndex) = astrFolders(lngIndex - 1) & "\" & astrFolders(lngIndex)
    Next lngIndex
    
    ' If the input path is UNC, the base element
    ' will be empty and the second "\", so we need
    ' to adjust where we start creating folders
    If Len(Trim$(astrFolders(0))) = 0 And UBound(astrFolders) > 0 Then
    
        If astrFolders(1) = "\" Then
            lngBase = 4  ' fourth element becomes base path
        End If
    
    End If
    
    For lngIndex = lngBase To UBound(astrFolders) - 1
        
        ' Check if this folder already exists
        If Not IsPathValid(astrFolders(lngIndex)) Then
            
            ' Create folder if it does not exist
            MkDir astrFolders(lngIndex)
            
            ' See if folder was created
            If Not IsPathValid(astrFolders(lngIndex)) Then
                Exit For
            End If
        End If
        
    Next lngIndex
    
CreateDirStructure_CleanUp:
    Erase astrFolders()                               ' Always empty arrays when not needed
    CreateDirStructure = IsPathValid(strCompletePath)   ' See if all folders were created
    On Error GoTo 0                                   ' Nullify this error trap
    Exit Function

CreateDirStructure_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume CreateDirStructure_CleanUp
    
End Function



' ***************************************************************************
' ****               Internal Procedures and Functions                   ****
' ***************************************************************************

' ***************************************************************************
' Routine:       HasLargeDriveSupport
'
' Description:   Determines whether the method supports a call to the
'                GetDiskFreeSpaceEx API.  If the call fails (no entry point
'                in kernel32) then the function returns false, indicating
'                that the application should use GetDiskFreeSpace API
'                instead.  GetDiskFreeSpace is limited to partitions less
'                than 2gb.
'
' Parameters:    strDrive - (Input) Drive letter to be evaluated
'
' Returns:       TRUE - Use GetDiskFreeSpaceEx for disk space info
'                FALSE - Use GetDiskFreeSpace for disk space info
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 09-JAN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' ***************************************************************************
Private Function HasLargeDriveSupport(ByVal strDrive As String) As Boolean

    ' Called by GetDiskSpaceInfo()
    
    Dim curDummy1 As Currency   ' dummy holding variables
    Dim curDummy2 As Currency
    Dim curDummy3 As Currency

    On Error GoTo HasLargeDriveSupport_Error

    strDrive = QualifyPath(strDrive)

    ' get the drive's disk parameters
    If GetDiskFreeSpaceEx(strDrive, curDummy1, curDummy2, curDummy3) <> 0 Then
        HasLargeDriveSupport = True   ' GetDiskFreeSpaceEx supported
    Else
        HasLargeDriveSupport = False  ' Use GetDiskFreeSpace instead
    End If

HasLargeDriveSupport_Exit:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

HasLargeDriveSupport_Error:
    ' No large drive support
    ' Use GetDiskFreeSpace instead
    Err.Clear
    HasLargeDriveSupport = False
    Resume HasLargeDriveSupport_Exit

End Function

' ***************************************************************************
' Routine:       GetMediaInfo
'
' Description:   Capture information concerning a particular device.
'
' Parameters:    strDrive - Drive letter to be queried
'
' Returns:       Returns device information
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 01-NOV-2001  Randy Birch  http://vbnet.mvps.org/index.html
'              Routine created
' 09-JAN-2004  Kenneth Ives  kenaso@tx.rr.com
'              Modified and documented
' 19-Jul-2008  Kenneth Ives  kenaso@tx.rr.com
'              Modified logic
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated routine and documented.
' ***************************************************************************
Private Function GetMediaInfo(ByVal strDrive As String) As Boolean

    ' Called by GetDriveInfo()
    
    Dim hDevice As Long
    
    Const ROUTINE_NAME As String = "GetMediaInfo"

    On Error GoTo GetMediaInfo_Error

    GetMediaInfo = False             ' Preset to FALSE
    ZeroMemory mtypDG, Len(mtypDG)   ' Empty type structure
    hDevice = INVALID_HANDLE_VALUE   ' -1
    
    ' Is operating system Windows XP or newer?
    If IsWinXPorNewer Then
            
        ' Verify drive letter is in proper format
        strDrive = Left$(strDrive, 1) & ":"

        ' hDevice will return INVALID_HANDLE_VALUE (-1)
        ' when a physical drive is not located.
        hDevice = CreateFile("\\.\" & strDrive, _
                             GENERIC_READ Or GENERIC_WRITE, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             ByVal 0&, OPEN_EXISTING, 0&, 0&)
                               
        ' If a valid handle, pass that to
        ' obtain the physical disk geometry
        If hDevice <> INVALID_HANDLE_VALUE Then

            ' Obtain DISK_GEOMETRY_EX type information
            mtypDG = GetDiskGeometryEx(hDevice)
            GetMediaInfo = True
            
        Else
            
            ' hDevice was invalid, so determine the error
            If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
                ' Something went wrong
            Else
                InfoMsg GetSystemMessage(Err.LastDllError) & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
            End If

        End If

    Else
        InfoMsg "This routine is NOT designed to execute on operating" & vbNewLine & _
                "systems earlier than Windows XP." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If
        
GetMediaInfo_CleanUp:
    CloseHandle hDevice  ' Always close the device handle
    On Error GoTo 0      ' Nullify this error trap
    Exit Function

GetMediaInfo_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ZeroMemory mtypDG, Len(mtypDG)   ' Empty type structure
    Resume GetMediaInfo_CleanUp

End Function

' ***************************************************************************
' Routine:       GetDiskGeometryEx
'
' Description:   Capture disk information to a type structure
'
' Reference:     Geforce - Senior Developer
'                http://www.vbib.be/PNphpBB2-viewtopic-t-5836.phtml
'
' Parameters:    hDevice - handle designation of drive being queried
'
' Returns:       Disk information in a type structure
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 14-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Rewrote routine
' ***************************************************************************
Private Function GetDiskGeometryEx(ByVal hDevice As Long) As DISK_GEOMETRY_EX

    ' Called by GetMediaInfo()
    
    Dim lngNotUsed     As Long
    Dim lngControlCode As Long
    
    Const ROUTINE_NAME As String = "GetDiskGeometryEx"

    On Error GoTo GetDiskGeometryEx_Error

    ' Start with empty return type structure
    ZeroMemory GetDiskGeometryEx, Len(GetDiskGeometryEx)

    ' Another handle check
    If hDevice <> INVALID_HANDLE_VALUE Then

        ' Calculate control code needed by API DeviceIoControl
        lngControlCode = GetControlCode(IOCTL_DISK_BASE, &H28, _
                                        METHOD_BUFFERED, FILE_ANY_ACCESS)
            
        ' Returned structure information is passed directly to
        ' type structure of this function
        If DeviceIoControl(hDevice, lngControlCode, ByVal 0&, 0&, _
                           GetDiskGeometryEx, Len(GetDiskGeometryEx), _
                           lngNotUsed, ByVal 0&) = 0 Then
        
            ' An error occurred. Verify structure is empty.
            ZeroMemory GetDiskGeometryEx, Len(GetDiskGeometryEx)
        
        End If
    End If
    
GetDiskGeometryEx_CleanUp:
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GetDiskGeometryEx_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    ZeroMemory GetDiskGeometryEx, Len(GetDiskGeometryEx)
    Resume GetDiskGeometryEx_CleanUp
    
End Function

Private Function GetControlCode(ByVal lngDeviceType As Long, _
                                ByVal lngFunction As Long, _
                                ByVal lngMethod As Long, _
                                ByVal lngAccess As Long) As Long

    ' Called by GetDiskGeometryEx()
    
    GetControlCode = ((lngDeviceType * (2 ^ 16)) Or _
                      (lngAccess * (2 ^ 14)) Or _
                      (lngFunction * (2 ^ 2)) Or _
                      lngMethod)
    
End Function

' ***************************************************************************
' Routine:       GetHiLoWord
'
' Description:   Get the low and high word of a long integer.
'
' Parameters:    lngValue - number to be evaluated
'                LOWORD - Low order of a word to be returned
'                HIWORD - High order of a word to be returned
'
' Returns:       Low and high order of a word (long integer)
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 09-Dec-2003  How to Mimic HIWORD, LOWORD, HIBYTE, LOBYTE C Macros in VB
'              http://support.microsoft.com/default.aspx?scid=kb;en-us;112651
' ***************************************************************************
Private Function GetHiLoWord(ByVal lngValue As Long, _
                             ByRef LOWORD As Long, _
                             ByRef HIWORD As Long) As Boolean

    ' Called by GetVolumeInfo()
    
    On Error GoTo GetHiLoWord_Error

    ' This is the LOWORD of the input value.
    ' LOWORD now equals 65,535 or &HFFFF
    LOWORD = lngValue And &HFFFF&
      
    ' This is the HIWORD of the input value.
    ' HIWORD now equals 30,583 or &H7777
    HIWORD = lngValue \ &H10000 And &HFFFF&
    
    ' Successful finish
    GetHiLoWord = True
    Exit Function

GetHiLoWord_Error:
    Err.Clear
    LOWORD = 0
    HIWORD = 0
    GetHiLoWord = False

End Function

Private Function GetPhysicalDriveInfo(ByVal lngDiskNo As enumIDE_DRIVE_NUMBER) As DRIVE_INFO
    
    ' Called by GetMfrInfo()
    
    Dim hDrive     As Long
    Dim typDrvInfo As DRIVE_INFO
    
    ZeroMemory typDrvInfo, Len(typDrvInfo)  ' Empty type structure
    hDrive = SmartOpen(lngDiskNo)           ' Capture drive handle
    
    If hDrive <> INVALID_HANDLE_VALUE Then
    
        If SmartGetVersion(hDrive) Then
        
            With typDrvInfo
                .bDriveType = 0
                .NumAttributes = 0
                ReDim .Attributes(0)
                .bDriveType = 1
            End With
            
            If SmartCheckEnabled(hDrive, lngDiskNo) Then
            
                If IdentifyDrive(hDrive, IDE_ID_FUNCTION, lngDiskNo, typDrvInfo) Then
                    GetPhysicalDriveInfo = typDrvInfo    ' Return data
                End If
                
            End If  ' SmartCheckEnabled
        End If      ' SmartGetVersion
    End If          ' hDrive
    
    CloseHandle hDrive                      ' Always close handle
    ZeroMemory typDrvInfo, Len(typDrvInfo)  ' Empty type structure
   
End Function

Private Function SmartOpen(ByVal lngDiskNo As enumIDE_DRIVE_NUMBER) As Long

    ' Called by GetPhysicalDriveInfo() [sub call of GetMfrInfo()]
    '
    ' Open SMART to allow DeviceIoControl
    ' communications and return SMART handle
    SmartOpen = CreateFile("\\.\PhysicalDrive" & CStr(lngDiskNo), _
                           GENERIC_READ Or GENERIC_WRITE, _
                           FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                           ByVal 0&, OPEN_EXISTING, 0&, 0&)
   
End Function

Private Function SmartGetVersion(ByVal hDrive As Long) As Boolean
   
    ' Called by GetPhysicalDriveInfo() [sub call of GetMfrInfo()]

    Dim lngReturn As Long
    Dim typGVOP   As GETVERSIONOUTPARAMS
    
    ZeroMemory typGVOP, Len(typGVOP)  ' Empty type structure
    
    SmartGetVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, _
                                      ByVal 0&, 0, _
                                      typGVOP, Len(typGVOP), _
                                      lngReturn, ByVal 0&)
   
    ZeroMemory typGVOP, Len(typGVOP)  ' Empty type structure
    
End Function

Private Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   ByVal lngDiskNo As enumIDE_DRIVE_NUMBER) As Boolean
   
    ' Called by GetPhysicalDriveInfo() [sub call of GetMfrInfo()]
    '
    ' Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
    ' lngDiskNo = 0 for first physical disk, 1=second physical disk, etc.
    
    Dim typSCIP          As SENDCMDINPARAMS
    Dim typSCOP          As SENDCMDOUTPARAMS
    Dim lngBytesReturned As Long
    
    ZeroMemory typSCIP, Len(typSCIP)  ' Empty type structure
    ZeroMemory typSCOP, Len(typSCOP)
    
    With typSCIP
        .cBufferSize = 0
        
        With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
            .bDriveHeadReg = &HA0
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
        End With
        
        .bDriveNumber = lngDiskNo
    End With
    
    SmartCheckEnabled = CBool(DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, _
                                              typSCIP, Len(typSCIP) - 4, _
                                              typSCOP, Len(typSCOP) - 4, _
                                              lngBytesReturned, ByVal 0&))
                                      
    ZeroMemory typSCIP, Len(typSCIP)  ' Empty type structures
    ZeroMemory typSCOP, Len(typSCOP)
    
End Function

Private Function IdentifyDrive(ByVal hDrive As Long, _
                               ByVal bytIDCmd As Byte, _
                               ByVal lngDiskNo As enumIDE_DRIVE_NUMBER, _
                               ByRef typDrvInfo As DRIVE_INFO) As Boolean
                               
    ' Called by GetPhysicalDriveInfo() [sub call of GetMfrInfo()]
    '
    ' Send an IDENTIFY command to the drive
    ' lngDiskNo = 0 for first physical disk, 1=second physical disk, etc.
    ' bytIDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
    
    Dim typSCIP     As SENDCMDINPARAMS
    Dim typIDSEC    As IDSECTOR
    Dim lngReturned As Long
    Dim abytTemp()  As Byte
    
    IdentifyDrive = True                  ' Preset flag to FALSE
    ZeroMemory typSCIP, Len(typSCIP)      ' Empty type structures
    ZeroMemory typIDSEC, Len(typIDSEC)
    ReDim abytTemp(OUTPUT_DATA_SIZE - 1)  ' Size temp hold array
    
    With typSCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(lngDiskNo)
        
        With .irDriveRegs
            .bFeaturesReg = 0
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = 0
            .bCylHighReg = 0
            .bDriveHeadReg = &HA0           ' compute drive number
            .bCommandReg = CByte(bytIDCmd)  ' command can either be IDE
                                            ' identify or ATAPI identify.
        End With
    End With
    
    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, _
                       typSCIP, Len(typSCIP) - 4, _
                       abytTemp(0), OUTPUT_DATA_SIZE, _
                       lngReturned, ByVal 0&) <> 0 Then
                
        ' Load type structure with pertinent data
        CopyMemory typIDSEC, abytTemp(16), Len(typIDSEC)
        
        ' Convert to legible data
        With typDrvInfo
            .Model = ByteArrayToString(SwapBytes(typIDSEC.sModelNumber()))
            .SerialNumber = ByteArrayToString(SwapBytes(typIDSEC.sSerialNumber()))
            .FirmWare = ByteArrayToString(SwapBytes(typIDSEC.sFirmwareRev()))
        End With
        
        IdentifyDrive = True
    
    End If
    
    ZeroMemory typSCIP, Len(typSCIP)    ' Empty type structures
    ZeroMemory typIDSEC, Len(typIDSEC)
    
End Function

Private Function SwapBytes(ByRef abytData() As Byte) As Byte()
   
    ' Called by IdentifyDrive() [sub call of GetMfrInfo()]
    
    Dim lngIndex As Long
    
    ' Swap data
    For lngIndex = LBound(abytData) To UBound(abytData) Step 2
        SwapByte abytData(lngIndex), abytData(lngIndex + 1)
    Next lngIndex
    
    SwapBytes = abytData()  ' Return new data
    
End Function

' ***************************************************************************
' Routine:       SwapByte
'
' Description:   Swap data (byte, integer, or long) with each other without
'                using a temp.  Using Math to do this takes a substantial
'                amount more processing than Logic Gates. The logic gates
'                foundation is in all processors. Which you could argue that
'                so is Math, but running a math process with a temp variable
'                in fact uses a considerable amount more processor cycles
'                than logic gates.
'
'                Using 3 Xor's over 3 Basic Equations greatly speeds up the
'                application. (especially when dealing with millions of
'                equations/Logic Gates)
'
'                This routine works with byte, integer and long values.
'                Change the parameter data type accordingly.
'
' Parameters:    AA - data to be swapped with Value2
'                BB - data to be swapped with Value1
'
' Returns:       Swapped data
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 10-Mar-2003  Ulli
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=43903&lngWId=1
' ***************************************************************************
Private Sub SwapByte(ByRef AA As Byte, _
                     ByRef BB As Byte)

    ' Called by SwapBytes()
    
    AA = AA Xor BB
    BB = BB Xor AA
    AA = AA Xor BB

End Sub

Private Function IsWinXPorNewer() As Boolean

    ' Called by GetDriveInfo()
    '           GetMediaInfo()
    '           IsDeviceReady()
    '           DeviceLock()
    '           GetDiskGeometryEx()
    
    Dim typOSVI As OSVERSIONINFO
    
    ZeroMemory typOSVI, Len(typOSVI)  ' Empty type structure
    
    With typOSVI
        .OSVSize = Len(typOSVI)
    
        If GetVersionEx(typOSVI) = 1 Then
        
            IsWinXPorNewer = (.PlatformID = VER_PLATFORM_WIN32_NT) And _
                             (.dwVerMajor = 5) And _
                             (.dwVerMinor >= 1)
        End If
    End With
    
    ZeroMemory typOSVI, Len(typOSVI)  ' Empty type structure
    
End Function

Private Function GetDriveTypeExtra(ByVal lngMediaType As enumMediaSize) As String

    ' Called by GetDriveInfo()
    '
    ' Additional information about the drive type
    Select Case lngMediaType
           Case eUnknownMedia:   GetDriveTypeExtra = "<Unknown>"
           Case eF5_1Pt2_512:    GetDriveTypeExtra = "5.25, 1.2MB, 512 bytes/sector"
           Case eF3_1Pt44_512:   GetDriveTypeExtra = "3.5, 1.44MB, 512 bytes/sector"
           Case eF3_2Pt88_512:   GetDriveTypeExtra = "3.5, 2.88MB, 512 bytes/sector"
           Case eF3_20Pt8_512:   GetDriveTypeExtra = "3.5, 20.8MB, 512 bytes/sector"
           Case eF3_720_512:     GetDriveTypeExtra = "3.5, 720KB, 512 bytes/sector"
           Case eF5_360_512:     GetDriveTypeExtra = "5.25, 360KB, 512 bytes/sector"
           Case eF5_320_512:     GetDriveTypeExtra = "5.25, 320KB, 512 bytes/sector"
           Case eF5_320_1024:    GetDriveTypeExtra = "5.25, 320KB, 1024 bytes/sector"
           Case eF5_180_512:     GetDriveTypeExtra = "5.25, 180KB, 512 bytes/sector"
           Case eF5_160_512:     GetDriveTypeExtra = "5.25, 160KB, 512 bytes/sector"
           Case eRemovableMedia: GetDriveTypeExtra = "Removable media other than floppy"
           Case eFixedMedia:     GetDriveTypeExtra = "Fixed hard disk"
           Case Else:            GetDriveTypeExtra = "<Unknown>"
    End Select

End Function

Private Function GetSystemMessage(ByVal lngMsgID As Long) As String

    ' Called by GetMediaInfo()
    
    Dim lngRetValue As Long
    Dim strBuffer   As String

    strBuffer = Space$(MAX_SIZE)

    lngRetValue = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                FORMAT_MESSAGE_IGNORE_INSERTS, _
                                0&, lngMsgID, 0, strBuffer, _
                                Len(strBuffer), ByVal 0)

    If (lngRetValue <> 0) Then
        GetSystemMessage = Left$(strBuffer, lngRetValue)
    Else
        GetSystemMessage = "Unknown error"
    End If

End Function

Private Function GetMediaType(ByVal strDrive As String) As String

    ' Called by IsCDRomDrive()
    
    Dim hHandle      As Long
    Dim lngStatus    As Long
    Dim lngReturn    As Long
    Dim lngNullValue As Long
    Dim typGMT       As GET_MEDIA_TYPES

    Const ROUTINE_NAME As String = "GetMediaType"

    On Error GoTo GetMediaType_Error

    GetMediaType = vbNullString
    ZeroMemory typGMT, Len(typGMT)  ' Empty type structure
        
    ' Is operating system Windows XP or newer?
    If IsWinXPorNewer Then
    
        ' Verify drive letter is in proper format
        strDrive = Left$(strDrive, 1) & ":"

        ' get a handle to drive
        hHandle = CreateFile("\\.\" & strDrive, _
                             GENERIC_READ Or GENERIC_WRITE, _
                             FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                             lngNullValue, OPEN_EXISTING, 0, lngNullValue)
        
        If hHandle <> INVALID_HANDLE_VALUE Then
        
           lngStatus = DeviceIoControl(hHandle, _
                                       IOCTL_STORAGE_GET_MEDIA_TYPES_EX, _
                                       lngNullValue, 0, typGMT, _
                                       KB_2, lngReturn, ByVal 0)
           
           If lngStatus <> 0 Then
               GetMediaType = typGMT.DeviceType
           End If
        
        End If
        
    Else
        InfoMsg "This routine is NOT designed to execute on operating" & vbNewLine & _
                "systems earlier than Windows XP." & _
                vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End If

GetMediaType_CleanUp:
    CloseHandle hHandle
    On Error GoTo 0   ' Nullify this error trap
    Exit Function

GetMediaType_Error:
    ErrorMsg MODULE_NAME, ROUTINE_NAME, Err.Description
    Resume GetMediaType_CleanUp

End Function

Private Function FormatSerialNo(ByVal strData As String) As String

    ' Called by GetDiskNumber()
    
    Dim avntData   As Variant
    Dim avntSerial As Variant
     
    On Error Resume Next
    
    avntData = Empty    ' Always start with empty variants
    avntSerial = Empty
    
    avntData = Split(strData, "\")   ' Split on backslash
    
    ' Split on ampersand in last element
    avntSerial = Split(avntData(UBound(avntData)), "&")
  
    ' Save serial number
    If UBound(avntSerial) > 0 Then
        FormatSerialNo = avntSerial(UBound(avntSerial) - 1)
    Else
        FormatSerialNo = avntSerial(UBound(avntSerial))
    End If

    avntData = Empty   ' Always empty variants when not needed
    avntSerial = Empty
    On Error GoTo 0
    
End Function

' ***************************************************************************
' Routine:       IsPathValid
'
' Description:   Determines whether a path to a file system object such as
'                a file or directory is valid. This function tests the
'                validity of the path. A path specified by Universal Naming
'                Convention (UNC) is limited to a file only; that is,
'                \\server\share\file is permitted. A UNC path to a server
'                or server share is not permitted; that is, \\server or
'                \\server\share. This function returns FALSE if a mounted
'                remote drive is out of service.
'
'                Requires Version 4.71 and later of Shlwapi.dll
'
' Reference:     http://msdn.microsoft.com/en-us/library/bb773584(v=vs.85).aspx
'
' Syntax:        IsPathValid("C:\Program Files\Desktop.ini")
'
' Parameters:    strName - Path or filename to be queried.
'
' Returns:       True or False
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 02-Nov-2009  Kenneth Ives  kenaso@tx.rr.com
'              Wrote routine
' ***************************************************************************
Private Function IsPathValid(ByVal strName As String) As Boolean

   IsPathValid = CBool(PathFileExists(strName))
   
End Function
 
' ***************************************************************************
' Routine:       QualifyPath
'
' Description:   Adds a trailing character to the path, if missing.
'
' Parameters:    strPath - Current folder being processed.
'                strChar - Optional - Specific character to append.
'                          Default = "\"
'
' Returns:       Fully qualified path with a specific trailing character
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' Unknown      Randy Birch
'              http://vbnet.mvps.org/index.html
' 14-MAY-2002  Kenneth Ives  kenaso@tx.rr.com
'              Modified/documented
' ***************************************************************************
Private Function QualifyPath(ByVal strPath As String, _
                    Optional ByVal strChar As String = "\") As String

    strPath = Trim$(strPath)
    
    If StrComp(Right$(strPath, 1), strChar, vbTextCompare) = 0 Then
        QualifyPath = strPath
    Else
        QualifyPath = strPath & strChar
    End If
    
End Function

' ***************************************************************************
' Routine:       ByteArrayToString
'
' Description:   Converts a byte array to string data
'
' Parameters:    abytData - array of bytes
'
' Returns:       Data string
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 25-Aug-2004  Kenneth Ives  kenaso@tx.rr.com
'              Routine created
' ***************************************************************************
Private Function ByteArrayToString(ByRef abytData() As Byte) As String

    ByteArrayToString = StrConv(abytData(), vbUnicode)

End Function

Private Sub ResetVariables()

  ZeroMemory mtypDG, Len(mtypDG)   ' Empty type structure

  mstrMfgModel = vbNullString      ' Initialize variables
  mstrDriveType = vbNullString
  mstrMfgSerial = vbNullString
  mstrUSBSerial = vbNullString
  mstrDrvVolName = vbNullString
  mstrMfgFirmware = vbNullString
  mstrDrvVolSerial = vbNullString
  mstrDrvFileSystem = vbNullString
  mstrPartitionData = vbNullString
  mstrDriveTypeExtra = vbNullString
  mstrDrvFormattedSize = vbNullString
  
  mlngDrvFreeClusters = 0
  mlngDrvTotalClusters = 0
  mlngDiskTracksPerCyl = 0
  mlngDrvBytesPerSector = 0
  mlngDrvBytesPerCluster = 0
  mlngDiskSectorsPerTrack = 0
  mlngDrvSectorsPerCluster = 0
  
  mcurDrvUsedSpace = 0@
  mcurDrvFreeSpace = 0@
  mcurTotalDrvSpace = 0@
  mcurDiskCylinders = 0@
  mcurDrvAvailableSpace = 0@

End Sub

Private Sub Class_Initialize()
    
    ' Whenever a class is instantiated (activated), this routine
    ' is called automatically if it exist with code inside.
    ' Ex:  Set objDiskInfo = new cDiskInfo
            
    ResetVariables           ' Initialize module variables
    StopProcessing = False   ' Preset property flag

End Sub

Private Sub Class_Terminate()

    ' Whenever a class is deactivated (freed from memory), this
    ' routine is called automatically if it exist with code inside.
    ' Ex:  Set objDiskInfo = Nothing
        
    ResetVariables           ' Initialize module variables
    
End Sub
