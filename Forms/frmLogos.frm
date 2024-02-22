VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLogos 
   Caption         =   "Logos"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   Icon            =   "frmLogos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRe 
      Caption         =   "Refrescar"
      Height          =   345
      Left            =   825
      TabIndex        =   9
      Top             =   9645
      Width           =   885
   End
   Begin VB.CommandButton cmdBuscaLogoFull 
      Height          =   360
      Left            =   60
      Picture         =   "frmLogos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9090
      Width           =   405
   End
   Begin VB.CommandButton cmdBuscaLogoSimple 
      Height          =   360
      Left            =   60
      Picture         =   "frmLogos.frx":25C4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   405
   End
   Begin VB.CommandButton cmdGrabaLogoSimple 
      Height          =   360
      Left            =   480
      Picture         =   "frmLogos.frx":42BE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   390
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2550
      Top             =   4365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGrabaLogoFull 
      Height          =   360
      Left            =   480
      Picture         =   "frmLogos.frx":45D0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9090
      Width           =   375
   End
   Begin VB.PictureBox picLogoFull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3720
      Left            =   90
      ScaleHeight     =   3720
      ScaleWidth      =   4560
      TabIndex        =   4
      Top             =   5355
      Width           =   4560
   End
   Begin VB.PictureBox picLogoSimple 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   90
      ScaleHeight     =   3510
      ScaleWidth      =   4590
      TabIndex        =   3
      Top             =   435
      Width           =   4590
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   45
      TabIndex        =   2
      Top             =   9645
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Logo Grande"
      Height          =   270
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   5100
      Width           =   1710
   End
   Begin VB.Label Label1 
      Caption         =   "Logo simple"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "frmLogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLogoSimple_Path_y_Nombre As String
Private mLogoFull_Path_y_Nombre As String
Private Const sFilter As String = "Bmp (Mapa bit)|*.bmp|Jpg (Imagen Debinida)|*.jp*|Todos|*.*|"


Private Sub cmdBuscaLogoFull_Click()
    On Error GoTo ufa
    CommonDialog1.Filter = sFilter
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName > "" Then
        mLogoFull_Path_y_Nombre = CommonDialog1.FileName
        picLogoFull.Picture = LoadPicture(CommonDialog1.FileName)
    End If
fin:
    Exit Sub
ufa:
    mLogoFull_Path_y_Nombre = ""
End Sub

Private Sub cmdBuscaLogoSimple_Click()
    On Error GoTo ufa
    CommonDialog1.Filter = sFilter
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName > "" Then
        mLogoSimple_Path_y_Nombre = CommonDialog1.FileName
        picLogoSimple.Picture = LoadPicture(CommonDialog1.FileName)
    End If
fin:
    Exit Sub
ufa:
    mLogoSimple_Path_y_Nombre = ""
End Sub


Private Sub cmdGrabaLogoFull_Click()
    Dim rs As New ADODB.Recordset
    
    If mLogoFull_Path_y_Nombre = "" Then
        MsgBox "No hay imagen para guardar.", vbExclamation
        Exit Sub
    End If

    With rs
        .Open "select imgLogoSimple, imgLogoFull from datosempresa d inner join bs on bs.idempresa = d.idempresa", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        AS_Arch_2_Base !imgLogoFull, mLogoFull_Path_y_Nombre
        .Update
    End With
    
    Set rs = Nothing
    MsgBox "Imagen guardada.", vbInformation
End Sub


Private Sub cmdGrabaLogoSimple_Click()
    Dim rs As New ADODB.Recordset
    
    If mLogoSimple_Path_y_Nombre = "" Then
        MsgBox "No hay imagen para guardar.", vbExclamation
        Exit Sub
    End If
    
    With rs
        .Open "select imgLogoSimple, imgLogoFull from datosempresa d inner join bs on bs.idempresa = d.idempresa", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        AS_Arch_2_Base !imgLogoSimple, mLogoSimple_Path_y_Nombre
        .Update
    End With
    
    Set rs = Nothing
    MsgBox "Imagen guardada.", vbInformation
End Sub


Private Sub cmdRe_Click()
    FrmPrincipal.Cargarlogo
End Sub

Private Sub cmdSalir_Click()
    FrmPrincipal.Cargarlogo
    Unload Me
End Sub


Private Sub Form_Load()
    On Error GoTo ufa
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select imgLogoSimple, imgLogoFull from datosempresa d inner join bs on bs.idempresa = d.idempresa", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic

        'simple
        AS_Base_2_Arch !imgLogoSimple, Archivo(False)
        picLogoSimple.Picture = LoadPicture(Archivo(False))
    
        'full
        AS_Base_2_Arch !imgLogoFull, Archivo(True)
        picLogoFull.Picture = LoadPicture(Archivo(True))
    
        .Close
    End With
    
ufa:
    Set rs = Nothing
End Sub

Public Function loadLogoSimple()
    On Error GoTo ufa
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select imgLogoSimple, imgLogoFull from datosempresa d inner join bs on bs.idempresa = d.idempresa", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic

        'simple
        AS_Base_2_Arch !imgLogoSimple, Archivo(False)
        Set loadLogoSimple = LoadPicture(Archivo(False))
        FileSystem.Kill Archivo(False)
    
'        'full
'        AS_Base_2_Arch !imgLogoFull, Archivo(True)
'        picLogoFull.Picture = LoadPicture(Archivo(True))
    
        .Close
    End With
    
ufa:
    Set rs = Nothing
End Function


Public Function loadLogoFull()
    On Error GoTo ufa
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select imgLogoSimple, imgLogoFull from datosempresa d inner join bs on bs.idempresa = d.idempresa", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic

        'simple
        AS_Base_2_Arch !imgLogoFull, Archivo(True)
        Set loadLogoFull = LoadPicture(Archivo(True))
        FileSystem.Kill Archivo(True)
    
'        'full
'        AS_Base_2_Arch !imgLogoFull, Archivo(True)
'        picLogoFull.Picture = LoadPicture(Archivo(True))
    
        .Close
    End With
    
ufa:
    Set rs = Nothing
End Function

Private Function Archivo(logofull As Boolean)
    Archivo = App.Path & "\" & IIf(logofull, "tmplogoFull.jpg", "tmplogosimple.jpg")
End Function

