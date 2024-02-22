Attribute VB_Name = "ModFileProperty"
'Option Explicit
'
'Const VS_FFI_SIGNATURE = &HFEEF04BD
'Const VS_FFI_STRUCVERSION = &H10000
'Const VS_FFI_FILEFLAGSMASK = &H3F&
'Const VS_FF_DEBUG = &H1
'Const VS_FF_PRERELEASE = &H2
'Const VS_FF_PATCHED = &H4
'Const VS_FF_PRIVATEBUILD = &H8
'Const VS_FF_INFOINFERRED = &H10
'Const VS_FF_SPECIALBUILD = &H20
'Const VOS_UNKNOWN = &H0
'Const VOS_DOS = &H10000
'Const VOS_OS216 = &H20000
'Const VOS_OS232 = &H30000
'Const VOS_NT = &H40000
'Const VOS__BASE = &H0
'Const VOS__WINDOWS16 = &H1
'Const VOS__PM16 = &H2
'Const VOS__PM32 = &H3
'Const VOS__WINDOWS32 = &H4
'Const VOS_DOS_WINDOWS16 = &H10001
'Const VOS_DOS_WINDOWS32 = &H10004
'Const VOS_OS216_PM16 = &H20002
'Const VOS_OS232_PM32 = &H30003
'Const VOS_NT_WINDOWS32 = &H40004
'Const VFT_UNKNOWN = &H0
'Const VFT_APP = &H1
'Const VFT_DLL = &H2
'Const VFT_DRV = &H3
'Const VFT_FONT = &H4
'Const VFT_VXD = &H5
'Const VFT_STATIC_LIB = &H7
'Const VFT2_UNKNOWN = &H0
'Const VFT2_DRV_PRINTER = &H1
'Const VFT2_DRV_KEYBOARD = &H2
'Const VFT2_DRV_LANGUAGE = &H3
'Const VFT2_DRV_DISPLAY = &H4
'Const VFT2_DRV_MOUSE = &H5
'Const VFT2_DRV_NETWORK = &H6
'Const VFT2_DRV_SYSTEM = &H7
'Const VFT2_DRV_INSTALLABLE = &H8
'Const VFT2_DRV_SOUND = &H9
'Const VFT2_DRV_COMM = &HA
'Private Type VS_FIXEDFILEINFO
'   dwSignature As Long
'   dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
'   dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
'   dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
'   dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
'   dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
'   dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
'   dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
'   dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
'   dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
'   dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
'   dwFileFlagsMask As Long        ' = &h3F for version "0.42"
'   dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
'   dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
'   dwFileType As Long             ' e.g. VFT_DRIVER
'   dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
'   dwFileDateMS As Long           ' e.g. 0
'   dwFileDateLS As Long           ' e.g. 0
'End Type
'Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
'Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
'Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
'
''Dim Filename As String, Directory As String, FullFileName As String
'Dim StrucVer As String, FileVer As String, ProdVer As String
'Dim FileFlags As String, FileOS As String, FileType As String, FileSubType As String
'
'Public Function VersionExe(ArchivoFullPath As String) As String
'   Dim rc As Long, lDummy As Long, sBuffer() As Byte
'   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
'   Dim lVerbufferLen As Long
'
'   '*** Get size ****
'   lBufferLen = GetFileVersionInfoSize(ArchivoFullPath, lDummy)
'   If lBufferLen < 1 Then
'      VersionExe = ""
'      Exit Function
'   End If
'
'   '**** Store info to udtVerBuffer struct ****
'   ReDim sBuffer(lBufferLen)
'   rc = GetFileVersionInfo(ArchivoFullPath, 0&, lBufferLen, sBuffer(0))
'   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
'   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
'
'   '**** Determine Structure Version number - NOT USED ****
'   StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)
'
'   '**** Determine File Version number ****
'   FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
'
'   '**** Determine Product Version number ****
'   ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
'
'
'    Dim aa
'    aa = Split(FileVer, ".")
'    'VerInfo = Format(aa(0), "000") & "." & Format(aa(1), "000") & "." & Format(aa(2), "000") & "." & Format(aa(3), "0000")
'    VersionExe = Format(aa(0), "00") & "." & Format(aa(1), "00") & "." & Format(aa(3), "0000")
'
'End Function
