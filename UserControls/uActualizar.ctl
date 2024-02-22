VERSION 5.00
Begin VB.UserControl uActualizar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ScaleHeight     =   2205
   ScaleWidth      =   2460
   Begin VB.TextBox txtNoDefinido 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   570
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "-Sin Fecha-"
      Top             =   660
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.TextBox txtFechaEXE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   570
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   195
      Width           =   1350
   End
   Begin VB.Frame fraOculto 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   2205
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2460
      Begin VB.TextBox txtFechaSQL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   660
         Width           =   1305
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   810
         Left            =   795
         Picture         =   "uActualizar.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label lblNoCoincide 
         BackColor       =   &H80000001&
         Caption         =   "Hay actualizaciones disponible"
         ForeColor       =   &H80000016&
         Height          =   330
         Left            =   150
         TabIndex        =   3
         Top             =   990
         Width           =   2340
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65000
      Left            =   900
      Top             =   795
   End
End
Attribute VB_Name = "uActualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const VS_FFI_SIGNATURE = &HFEEF04BD
Const VS_FFI_STRUCVERSION = &H10000
Const VS_FFI_FILEFLAGSMASK = &H3F&
Const VS_FF_DEBUG = &H1
Const VS_FF_PRERELEASE = &H2
Const VS_FF_PATCHED = &H4
Const VS_FF_PRIVATEBUILD = &H8
Const VS_FF_INFOINFERRED = &H10
Const VS_FF_SPECIALBUILD = &H20
Const VOS_UNKNOWN = &H0
Const VOS_DOS = &H10000
Const VOS_OS216 = &H20000
Const VOS_OS232 = &H30000
Const VOS_NT = &H40000
Const VOS__BASE = &H0
Const VOS__WINDOWS16 = &H1
Const VOS__PM16 = &H2
Const VOS__PM32 = &H3
Const VOS__WINDOWS32 = &H4
Const VOS_DOS_WINDOWS16 = &H10001
Const VOS_DOS_WINDOWS32 = &H10004
Const VOS_OS216_PM16 = &H20002
Const VOS_OS232_PM32 = &H30003
Const VOS_NT_WINDOWS32 = &H40004
Const VFT_UNKNOWN = &H0
Const VFT_APP = &H1
Const VFT_DLL = &H2
Const VFT_DRV = &H3
Const VFT_FONT = &H4
Const VFT_VXD = &H5
Const VFT_STATIC_LIB = &H7
Const VFT2_UNKNOWN = &H0
Const VFT2_DRV_PRINTER = &H1
Const VFT2_DRV_KEYBOARD = &H2
Const VFT2_DRV_LANGUAGE = &H3
Const VFT2_DRV_DISPLAY = &H4
Const VFT2_DRV_MOUSE = &H5
Const VFT2_DRV_NETWORK = &H6
Const VFT2_DRV_SYSTEM = &H7
Const VFT2_DRV_INSTALLABLE = &H8
Const VFT2_DRV_SOUND = &H9
Const VFT2_DRV_COMM = &HA
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
   dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
   dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
   dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
   dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
   dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
   dwFileFlagsMask As Long        ' = &h3F for version "0.42"
   dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             ' e.g. VFT_DRIVER
   dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           ' e.g. 0
   dwFileDateLS As Long           ' e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

Private mUbicacionExeOrigen As String


Public Function VersionExe(ArchivoFullPath As String) As String
'    Dim StrucVer As String, FileVer As String, ProdVer As String
'    Dim FileFlags As String, FileOS As String, FileType As String, FileSubType As String
    Dim FileVer As String

   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(ArchivoFullPath, lDummy)
   If lBufferLen < 1 Then
      VersionExe = ""
      Exit Function
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(ArchivoFullPath, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

'   '**** Determine Structure Version number - NOT USED ****
'   StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)

   '**** Determine File Version number ****
   FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)

'   '**** Determine Product Version number ****
'   ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)

    Dim aa
    aa = Split(FileVer, ".")
    VersionExe = Format(aa(0), "00") & "." & Format(aa(1), "00") & "." & Format(aa(3), "0000")
End Function




Private Sub RevisaVersion()
    txtFechaSQL = VersionExe(afull(mUbicacionExeOrigen))
    txtFechaEXE = VersionExe(afull(App.Path))
    
    
    If txtFechaEXE = "" Then
        fraOculto.Visible = False
    Else
        'fraOculto.Visible = True
        fraOculto.Visible = txtFechaSQL > txtFechaEXE
    End If
    txtNoDefinido.Visible = (Trim(txtFechaSQL) = "")
End Sub

Public Property Let UbicacionEXE(CarpetaFullPath As String)
    mUbicacionExeOrigen = CarpetaFullPath
    If mUbicacionExeOrigen = "" Then
        Timer1.enabled = False
    Else
        Timer1.enabled = True
        RevisaVersion
    End If
End Property

Private Sub cmdActualizar_Click()
    Dim ss As String, co As String
    co = copi()
    If co = "" Then
        MsgBox "no encuentro el copiador"
    Else
        ss = co & afull(mUbicacionExeOrigen) & " , " & afull(App.Path)
        If MsgBox("¿Cerrar programa y actualizar?", vbYesNo + vbInformation, "Informe") = vbYes Then
            Shell ss
            End
        End If
    End If
End Sub

Private Function copi() As String
    If VersionExe(App.Path & "\actualizar.exe  ") > "" Then
        copi = App.Path & "\actualizar.exe  "
    ElseIf VersionExe(mUbicacionExeOrigen & "\actualizar.exe ") > "" Then
        copi = mUbicacionExeOrigen & "\actualizar.exe "
    End If
End Function
Private Function afull(ByVal dire As String) As String
    dire = Trim(dire)
    If dire = "" Then Exit Function
    If Right(dire, 1) <> "\" Then dire = dire & "\"
    afull = dire & App.EXEName & ".EXE"
End Function

Private Sub UserControl_DblClick()
    RevisaVersion
End Sub
Private Sub fraOculto_DblClick()
    RevisaVersion
End Sub
Private Sub Timer1_Timer()
    Dim x As Long
     x = (x Mod 10) + 1
    If x = 10 Then RevisaVersion
End Sub

