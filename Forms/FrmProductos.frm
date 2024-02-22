VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.Ocx"
Begin VB.Form frmProductos 
   Caption         =   "Producto"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "FrmProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk4 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   95
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chk3 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   94
      Top             =   5000
      Width           =   255
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   93
      Top             =   4720
      Width           =   255
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9360
      TabIndex        =   92
      Top             =   4440
      Width           =   255
   End
   Begin VB.CheckBox chkFactu 
      Caption         =   "Facturable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   91
      Top             =   2205
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox txtDesc 
      Height          =   660
      Left            =   1320
      TabIndex        =   90
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1164
      _Version        =   393217
      TextRTF         =   $"FrmProductos.frx":08CA
   End
   Begin VB.ComboBox cboCParte 
      Height          =   315
      Left            =   4590
      Style           =   2  'Dropdown List
      TabIndex        =   87
      Top             =   2175
      Width           =   4635
   End
   Begin VB.CommandButton cmdNFactor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9240
      Picture         =   "FrmProductos.frx":094D
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   1860
      Width           =   615
   End
   Begin VB.ComboBox cboConversion 
      Height          =   315
      Left            =   4590
      Style           =   2  'Dropdown List
      TabIndex        =   85
      Top             =   1860
      Width           =   4635
   End
   Begin VB.CheckBox chkTiene_Cuenta 
      Alignment       =   1  'Right Justify
      Caption         =   "Cta Contable"
      Height          =   330
      Left            =   5085
      TabIndex        =   10
      Top             =   1470
      Width           =   1215
   End
   Begin VB.ComboBox cboEstado 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":0ED7
      Left            =   2025
      List            =   "FrmProductos.frx":0EE1
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Tag             =   "31"
      Top             =   6105
      Width           =   2295
   End
   Begin VB.ComboBox cboManejaStock 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":0EED
      Left            =   2010
      List            =   "FrmProductos.frx":0EEF
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Tag             =   "17"
      Top             =   4410
      Width           =   1095
   End
   Begin Gestion.ucCoDe uCuenta 
      Height          =   315
      Left            =   6315
      TabIndex        =   11
      Top             =   1485
      Width           =   3615
      _ExtentX        =   8281
      _ExtentY        =   450
      CodigoWidth     =   1000
   End
   Begin VB.TextBox txtAlias 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5685
      TabIndex        =   9
      Tag             =   "2"
      Top             =   1185
      Width           =   3435
   End
   Begin VB.TextBox txtExistencia 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5490
      TabIndex        =   29
      Tag             =   "19"
      Top             =   4410
      Width           =   1110
   End
   Begin VB.CommandButton cmdguardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar Como"
      Height          =   765
      Left            =   8070
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmProductos.frx":0EF1
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5790
      Width           =   1140
   End
   Begin VB.CommandButton cmdSubgrupo 
      DisabledPicture =   "FrmProductos.frx":17BB
      Height          =   315
      Left            =   1320
      Picture         =   "FrmProductos.frx":1B45
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1515
      Width           =   435
   End
   Begin VB.CommandButton cmdGrupo 
      DisabledPicture =   "FrmProductos.frx":1ECF
      Height          =   315
      Left            =   1320
      Picture         =   "FrmProductos.frx":2259
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1185
      Width           =   435
   End
   Begin VB.TextBox txtSubGrupo 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1755
      TabIndex        =   6
      Top             =   1500
      Width           =   615
   End
   Begin VB.TextBox txtGrupo 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1755
      TabIndex        =   3
      Top             =   1185
      Width           =   615
   End
   Begin Gestion.ucBotonera ucMenu 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   1620
      Left            =   2040
      TabIndex        =   46
      Top             =   7365
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   2858
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.ComboBox cboMedidas 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "0"
      Top             =   1860
      Width           =   1050
   End
   Begin VB.CheckBox chkformula 
      Alignment       =   1  'Right Justify
      Caption         =   "Tiene Formula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4995
      TabIndex        =   43
      Top             =   5835
      Width           =   1620
   End
   Begin VB.ComboBox cmbmoneda 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":25E3
      Left            =   4800
      List            =   "FrmProductos.frx":25ED
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Tag             =   "31"
      Top             =   3975
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9420
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Imagen BMP JPG |*.jpg;*.bmp"
   End
   Begin VB.TextBox txtobserv 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2025
      TabIndex        =   42
      Tag             =   "30"
      Top             =   6735
      Width           =   7455
   End
   Begin VB.TextBox txtcontrol 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2010
      TabIndex        =   28
      Tag             =   "29"
      Top             =   5325
      Width           =   1335
   End
   Begin VB.TextBox txtletra 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4335
      TabIndex        =   38
      Tag             =   "28"
      Top             =   5820
      Width           =   615
   End
   Begin VB.CommandButton cmbgrafico 
      DisabledPicture =   "FrmProductos.frx":25F9
      Height          =   300
      Left            =   6240
      Picture         =   "FrmProductos.frx":363B
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6420
      Width           =   390
   End
   Begin VB.TextBox txtgrafico 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2025
      TabIndex        =   41
      Tag             =   "27"
      Top             =   6435
      Width           =   4215
   End
   Begin VB.TextBox txtelaboracion 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2025
      TabIndex        =   37
      Tag             =   "26"
      Top             =   5805
      Width           =   1335
   End
   Begin VB.TextBox txtdep1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8100
      TabIndex        =   33
      Tag             =   "22"
      Top             =   4410
      Width           =   1110
   End
   Begin VB.TextBox txtdep2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8100
      TabIndex        =   34
      Tag             =   "23"
      Top             =   4695
      Width           =   1110
   End
   Begin VB.TextBox txtdep3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8100
      TabIndex        =   35
      Tag             =   "24"
      Top             =   4980
      Width           =   1110
   End
   Begin VB.TextBox txtdep4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8100
      TabIndex        =   36
      Tag             =   "25"
      Top             =   5265
      Width           =   1110
   End
   Begin VB.TextBox txtsaldo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5490
      TabIndex        =   32
      Tag             =   "21"
      Top             =   5310
      Width           =   1110
   End
   Begin VB.TextBox txtpedidos 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5490
      TabIndex        =   31
      Tag             =   "20"
      Top             =   5010
      Width           =   1110
   End
   Begin VB.TextBox txtExistenciaCalculada 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5490
      TabIndex        =   30
      Tag             =   "19"
      Top             =   4710
      Width           =   1110
   End
   Begin VB.ComboBox cmbactivo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":467D
      Left            =   2010
      List            =   "FrmProductos.frx":4687
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Tag             =   "31"
      Top             =   7020
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtiva 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1305
      TabIndex        =   16
      Tag             =   "18"
      Text            =   "0,21"
      Top             =   3975
      Width           =   555
   End
   Begin VB.ComboBox cmbserie 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":4693
      Left            =   8175
      List            =   "FrmProductos.frx":469D
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Tag             =   "17"
      Top             =   3630
      Width           =   690
   End
   Begin VB.TextBox txtpedidomin 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2010
      TabIndex        =   27
      Tag             =   "15"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox txtstockmin 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2010
      TabIndex        =   26
      Tag             =   "14"
      Top             =   4740
      Width           =   1335
   End
   Begin VB.TextBox txtcostoprov 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1305
      TabIndex        =   15
      Tag             =   "13"
      Top             =   3675
      Width           =   1215
   End
   Begin VB.TextBox txtprelis4 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7500
      TabIndex        =   20
      Tag             =   "11"
      Top             =   2970
      Width           =   1335
   End
   Begin VB.TextBox txtprelis3 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7500
      TabIndex        =   19
      Tag             =   "10"
      Top             =   2670
      Width           =   1335
   End
   Begin VB.TextBox txtprelis2 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4800
      TabIndex        =   18
      Tag             =   "9"
      Top             =   2970
      Width           =   1335
   End
   Begin VB.TextBox txtprelis1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Tag             =   "8"
      Top             =   2670
      Width           =   1335
   End
   Begin VB.TextBox txtcosto 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1305
      TabIndex        =   13
      Tag             =   "7"
      Top             =   2970
      Width           =   1215
   End
   Begin VB.TextBox txtcodigo 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1335
      TabIndex        =   0
      Tag             =   "2"
      Top             =   120
      Width           =   2130
   End
   Begin VB.ComboBox cboGrupo 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   1185
      Width           =   2670
   End
   Begin VB.TextBox txtdescripcion 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4710
      TabIndex        =   1
      Tag             =   "3"
      Top             =   135
      Width           =   4395
   End
   Begin VB.ComboBox cmbproducto 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "FrmProductos.frx":46A9
      Left            =   4800
      List            =   "FrmProductos.frx":46BC
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Tag             =   "12"
      Top             =   3630
      Width           =   2175
   End
   Begin VB.TextBox txtcostobase 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1305
      TabIndex        =   12
      Tag             =   "5"
      Top             =   2670
      Width           =   1215
   End
   Begin VB.ComboBox cmbcalculo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":4714
      Left            =   1755
      List            =   "FrmProductos.frx":471E
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Tag             =   "6"
      Top             =   3315
      Width           =   765
   End
   Begin VB.TextBox txtcodbarras 
      BackColor       =   &H80000009&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4800
      TabIndex        =   21
      Tag             =   "16"
      Top             =   3300
      Width           =   4050
   End
   Begin VB.ComboBox cboSubgrupo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmProductos.frx":472A
      Left            =   2370
      List            =   "FrmProductos.frx":472C
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   1500
      Width           =   2670
   End
   Begin VB.Label Label37 
      Caption         =   "Descripción Presupuesto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   89
      Top             =   480
      Width           =   1290
   End
   Begin VB.Label Label36 
      Caption         =   "Factor para Produccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2430
      TabIndex        =   88
      Top             =   2205
      Width           =   2490
   End
   Begin VB.Label Label35 
      Caption         =   "Factor para Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2415
      TabIndex        =   84
      Top             =   1890
      Width           =   2490
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9930
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9915
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9945
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label34 
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   83
      Top             =   6105
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Maneja Stock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   82
      Top             =   4455
      Width           =   1410
   End
   Begin VB.Label Label10 
      Caption         =   "Alias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   5115
      TabIndex        =   81
      Top             =   1185
      Width           =   795
   End
   Begin VB.Label lblExistenciaCalculada 
      Caption         =   "Existencia Calculada:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3570
      TabIndex        =   80
      Top             =   4710
      Width           =   1920
   End
   Begin VB.Label Label5 
      Caption         =   "Existencia Fisica:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3930
      TabIndex        =   79
      Top             =   4410
      Width           =   1740
   End
   Begin VB.Label Label33 
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3990
      TabIndex        =   78
      Top             =   3990
      Width           =   975
   End
   Begin VB.Label Label32 
      Caption         =   "Observaciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   585
      TabIndex        =   77
      Top             =   6735
      Width           =   1815
   End
   Begin VB.Label Label31 
      Caption         =   "Cant. a controlar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   510
      TabIndex        =   76
      Top             =   5325
      Width           =   1575
   End
   Begin VB.Label Label30 
      Caption         =   "Letra Act.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3465
      TabIndex        =   75
      Top             =   5805
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "Imagen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1290
      TabIndex        =   74
      Top             =   6435
      Width           =   735
   End
   Begin VB.Label Label28 
      Caption         =   "Tiempo Elaboración:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   73
      Top             =   5805
      Width           =   1950
   End
   Begin VB.Label Label21 
      Caption         =   "Deposito 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7065
      TabIndex        =   72
      Top             =   4695
      Width           =   1500
   End
   Begin VB.Label Label20 
      Caption         =   "Deposito 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7065
      TabIndex        =   71
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "Deposito 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7065
      TabIndex        =   70
      Top             =   5265
      Width           =   1590
   End
   Begin VB.Label Label23 
      Caption         =   "Deposito 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7065
      TabIndex        =   69
      Top             =   4410
      Width           =   1275
   End
   Begin VB.Label Label15 
      Caption         =   "Disponible:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4455
      TabIndex        =   68
      Top             =   5310
      Width           =   1155
   End
   Begin VB.Label Label7 
      Caption         =   "Pedidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4665
      TabIndex        =   67
      Top             =   5010
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Activo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1395
      TabIndex        =   66
      Top             =   7035
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Iva:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   915
      TabIndex        =   65
      Top             =   3990
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Posee serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   7020
      TabIndex        =   64
      Top             =   3660
      Width           =   1335
   End
   Begin VB.Label Label27 
      Caption         =   "Pedido mínimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   63
      Top             =   4995
      Width           =   1575
   End
   Begin VB.Label Label26 
      Caption         =   "Stock mínimo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   750
      TabIndex        =   62
      Top             =   4740
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "Costo Prov.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   195
      TabIndex        =   61
      Top             =   3675
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Precio Lista 4:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6255
      TabIndex        =   60
      Top             =   2970
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "Precio Lista 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6255
      TabIndex        =   59
      Top             =   2685
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Precio Lista 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3525
      TabIndex        =   58
      Top             =   3000
      Width           =   1545
   End
   Begin VB.Label Label9 
      Caption         =   "Precio Lista 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3525
      TabIndex        =   57
      Top             =   2670
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Costo Promedio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   735
      TabIndex        =   56
      Top             =   2955
      Width           =   540
   End
   Begin VB.Label Label16 
      Caption         =   "Calculo s/Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   255
      TabIndex        =   55
      Top             =   3330
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Grupo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   690
      TabIndex        =   54
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label12 
      Caption         =   "Descripción;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3555
      TabIndex        =   53
      Top             =   135
      Width           =   1290
   End
   Begin VB.Label Label10 
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   52
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   51
      Top             =   3645
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "U de Medida:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   105
      TabIndex        =   50
      Top             =   1860
      Width           =   1320
   End
   Begin VB.Label Label13 
      Caption         =   "Costo Base:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   210
      TabIndex        =   49
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label18 
      Caption         =   "Código de barras:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3165
      TabIndex        =   48
      Top             =   3315
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "Sub Grupo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   285
      TabIndex        =   47
      Top             =   1500
      Width           =   1095
   End
End
Attribute VB_Name = "frmProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numero As Long
Dim WithEvents grupo As LiCodigo
Attribute grupo.VB_VarHelpID = -1
Dim WithEvents SubGrupo As LiCodigo
Attribute SubGrupo.VB_VarHelpID = -1
Dim mens As String
Private uMedidas() As datCD
Private uFactor() As datCD

Private Sub cboMedidas_Change()
    CargoFactores OBTC
End Sub

Private Sub cboMedidas_Click()
    CargoFactores OBTC
End Sub

Private Function OBTC() As Long
If cboMedidas.ListIndex = -1 Then Exit Function
    OBTC = nSinNull(obtenerDeSQL("select tipo from unidadesmedida where umcodigo=" & uMedidas(cboMedidas.ListIndex).dCodigo))
End Function

Private Sub cmbgrafico_Click()
    CommonDialog1.Action = 1
    txtgrafico = CommonDialog1.FileName
End Sub

Private Sub cmdguardar_Click()

Dim rs As New ADODB.Recordset, ManejaStock As Long, tiene_Cuenta_PRO As Long
Dim factu As Long
If ON_ERROR_HABILITADO Then On Error GoTo ufaErr

mens = InputBox("Ingrese el código del nuevo Producto")
If mens = "" Then
    MsgBox "Debe ingresar algun código", 48, "Atencion"
    Exit Sub
Else
    rs.Open "select * from producto where codigo='" & Trim(mens) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "El codigo ya existe", 48, "ATENCION"
        Exit Sub
    Else
        
    
        If FaltanCosas() Then Exit Sub
        
        Dim Serie As Long, formula As Long
        If gEMPR_Maneja_series Then
            Serie = IIf(Trim(cmbserie) = "NO", 0, 1)
        Else
            Serie = 0
        End If
'        ManejaStock = IIf(Trim(cboManejaStock) = "NO", 0, 1)
        ManejaStock = ComboCodigo(cboManejaStock)
        formula = IIf(chkformula.Value = 0, 0, 1)
        factu = IIf(chkFactu.Value = 0, 0, 1)
    
    
        If Not IsEmpty(obtenerDeSQL("select codigo from producto where activo = 1 and codigo = '" & CodigoGrabarAlternativo() & "'")) Then
            MsgBox " Grupo + Subgrupo + Codigo  ya existe "
            Exit Sub
        End If
        
        If chkTiene_Cuenta.Value = 0 Then
             tiene_Cuenta_PRO = 0
        Else
             tiene_Cuenta_PRO = 1
        End If
        
        'DataEnvironment1.dbo_PRODUCTOS "A", CodigoGrabarAlternativo(), grupo.codigo, SubGrupo.codigo, _
             Trim(txtdescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, s2n(txtcostobase), IIf(cmbcalculo.Text = "SI", 1, 0), s2n(txtcosto), s2n(txtprelis1), _
            s2n(txtprelis2), s2n(txtprelis3), s2n(txtprelis4), ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text)), s2n(txtcostoprov), s2n(txtstockmin), _
            s2n(txtpedidomin), Trim(txtcodbarras), Serie, s2n(txtiva), s2n(txtExistencia), _
            s2n(txtelaboracion), Trim(txtgrafico), Trim(txtletra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)), formula, txtAlias, Date, UsuarioSistema!codigo, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO
        ABMProducto "A", CodigoGrabarAlternativo(), Trim(txtCodigo), grupo.codigo, SubGrupo.codigo, _
        Trim(txtDescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, uFactor(cboConversion.ListIndex).dCodigo, uFactor(cboCParte.ListIndex).dCodigo, s2n(txtcostobase, 4), IIf(cmbcalculo.Text = "SI", 1, 0), s2n(txtcosto, 4), s2n(txtprelis1, 4), _
        s2n(txtprelis2, 4), s2n(txtprelis3, 4), s2n(txtprelis4, 4), ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text)), s2n(txtcostoprov, 4), s2n(txtstockmin, 4), _
        s2n(txtpedidomin, 4), Trim(txtcodbarras), Serie, s2n(txtIva, 4), s2n(txtExistencia), _
        s2n(txtelaboracion), Trim(txtgrafico), Trim(txtLetra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbMoneda.Text)), formula, txtAlias, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO, txtDesc.TextRTF, factu, Alma

        'MsgBox "La operación fue realizada con éxito"
        ucMenu.AceptarOk ("codigo = '" & CodigoGrabarAlternativo() & "'")
    
    End If
    Set rs = Nothing
End If


GoTo fin
ufaErr:
    ufa "err en alta", Me.Name ', Err
fin:
End Sub

Private Sub cmdNFactor_Click()
    FrmABMUnidadFactor.Show
End Sub

Private Sub Form_Load()
    Set grupo = New LiCodigo
    Set SubGrupo = New LiCodigo
    
    grupo.init cboGrupo, txtGrupo, "GruposProducto", False, True, cmdGrupo, "activo = 1 "
    SubGrupo.init cboSubGrupo, txtSubGrupo, "SubGruposProducto", False, True, cmdSubgrupo, "activo=1"
    comboSql cboEstado, "select Estado, codigo from ProductoEstado order by codigo"
    
    CargaCombo cmbproducto, "TIPOPRODUCTOS", "descripcion", "codigo", ""
    CargaCombo cmbMoneda, "MONEDAS", "descripcion", "codigo", ""
    CargoMedidas
    comboArray cboManejaStock, Array("SI", "NO"), Array(1, 0)
   
    ucMenu.init True, True, True, False, True, "select * from Producto where activo = 1", DataEnvironment1.Sistema, True
    ucMenu.MsgConfirmaEliminar = "Esta seguro de querer eliminar este registro ?"
    ucMenu.MsgConfirmaSalir = "Cerrar Formulario ?"
    
    If gEMPR_FormulaEsVirtual Then
        GeneraExistenciaCalculada
        txtExistenciaCalculada.Visible = True
        lblExistenciaCalculada.Visible = True
    End If
    
    uCuenta.ini "select Descripcion from Cuentas where cuenta = '###' ", "Select cuenta as [ Cuenta          ], Descripcion as [ Descripcion                             ] from cuentas where activo = 1 and imputable = 1 ", True
    
    cmbserie.Visible = gEMPR_Maneja_series
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, False, True
End Sub


Sub LimpioControles()
    On Error Resume Next
    FrmBorrarCbo Me

    grupo.codigo = ""
    SubGrupo.codigo = ""
    txtCodigo = ""
    txtDescripcion = ""
    txtDesc.TextRTF = ""
    cboMedidas.ListIndex = 0
    cboConversion.ListIndex = 0
    cboCParte.ListIndex = 0
    txtcostobase = "0.0000"
    txtcosto = "0.0000"
    txtprelis1 = "0.0000"
    txtprelis2 = "0.0000"
    txtprelis3 = "0.0000"
    txtprelis4 = "0.0000"
    cmbproducto.ListIndex = 2
    txtcostoprov = "0.0000"
    txtstockmin = "0"
    txtpedidomin = "0"
    txtcodbarras = ""


    cmbserie.ListIndex = BuscarenComboS(cmbserie, IIf(gEMPR_Default_ProductoConSerie, "SI", "NO"))
    cboManejaStock.ListIndex = BuscarenComboS(cboManejaStock, "SI")
    
    cmbMoneda.ListIndex = 2 'selecciona por defecto la moneda en pesos

    txtIva = "0.21"
    txtExistencia = ""
    txtExistenciaCalculada = ""
    txtpedidos = "0"
    txtsaldo = "0.00"
    txtdep1 = "0"
    txtdep2 = "0"
    txtdep3 = "0"
    txtdep4 = "0"
    txtgrafico = ""
    txtLetra = ""
    txtcontrol = "0"
    txtobserv = ""

    chkformula.Value = 0
    chkFactu.Value = 0
    txtAlias = ""
    chkTiene_Cuenta.Value = 0
    uCuenta.clear
End Sub

Private Sub CargoProducto()
Dim rsped As New ADODB.Recordset
Dim pendiente As Variant

  With ucMenu.rs
    'If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        If Not IsNull(!grupo) Then
            grupo.codigo = !grupo
        End If
        If Not IsNull(!SubGrupo) Then
            SubGrupo.codigo = !SubGrupo
        End If
    'End If
    txtAlias = sSinNull(!Alias)
    
    If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        txtCodigo = Mid(!codigo, 7)  ' Limpio grupo y subgripo
    Else
        txtCodigo = !codigo
    End If
    
    
    txtDescripcion = !DESCRIPCION
    txtDesc.TextRTF = sSinNull(!DescGeneral)
    
    cboMedidas.ListIndex = OBM(nSinNull(!UMedida))
    cboConversion.ListIndex = OBF(nSinNull(!uFactor))
    cboCParte.ListIndex = OBF(nSinNull(!Uparte))
    
    cboEstado.ListIndex = BuscarEnCombo(cboEstado, !estado)
    
    
    If Not IsNull(!costobase) Then
        txtcostobase = !costobase
    End If
    
    cmbcalculo.ListIndex = BuscarenComboS(cmbcalculo, IIf(!CALCSINCOSTO, "SI", "NO"))
    If Not IsNull(!COSTOPROM) Then
        txtcosto = !COSTOPROM
    End If
    If Not IsNull(!precio) Then
        txtprelis1 = s2n(!precio, 4)
    End If
    If Not IsNull(!precio2) Then
        txtprelis2 = !precio2
    End If
    If Not IsNull(!precio3) Then
        txtprelis3 = !precio3
    End If
    If Not IsNull(!precio4) Then
        txtprelis4 = !precio4
    End If
    rsped.Open "Select saldo from itempedidocliente i inner join pedidos_clientes p on i.pedido=p.numero where i.producto='" & !codigo & "' and i.saldo<>0 and p.activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    pendiente = 0
    If Not rsped.EOF Then
        Do While Not rsped.EOF
            pendiente = s2n(pendiente) + s2n(rsped!saldo)
            rsped.MoveNext
        Loop
    End If
    txtpedidos = s2n(pendiente)
    
    
    rsped.Close
    Set rsped = Nothing
    If Not IsNull(!TIPOPROD) Then
        cmbproducto.ListIndex = BuscarenComboS(cmbproducto, ObtenerDescripcion("tipoproductos", !TIPOPROD))
    End If
    If Not IsNull(!COSTOPROV) Then
        txtcostoprov = !COSTOPROV
    End If
    If Not IsNull(!STOCKMIN) Then
        txtstockmin = !STOCKMIN
    End If
    If Not IsNull(!pedmin) Then
        txtpedidomin = !pedmin
    End If
    If Not IsNull(!CODIGOBARRA) Then
        txtcodbarras = !CODIGOBARRA
    End If
    
    cmbserie.ListIndex = BuscarenComboS(cmbserie, IIf(!Serie, "SI", "NO"))
    cboManejaStock.ListIndex = BuscarenComboS(cboManejaStock, IIf(nSinNull(!ManejaStock), "SI", "NO"))
    
    If Not IsNull(!Iva) Then
        txtIva = !Iva
    End If
   
    If Not IsNull(!existencia) Then
        If gEMPR_FormulaEsVirtual Then
            txtExistencia = !existencia
            txtExistenciaCalculada = !ExistenciaCalculada
            txtsaldo = s2n(!ExistenciaCalculada) - s2n(pendiente)
        Else
            txtExistencia = !existencia
            txtsaldo = s2n(!existencia) - s2n(pendiente)
        End If
    End If
    
    If Not IsNull(!dep1) Then
        txtdep1 = !dep1
    End If
    If Not IsNull(!dep2) Then
        txtdep2 = !dep2
    End If
    If Not IsNull(!dep3) Then
        txtdep3 = !dep3
    End If
    If Not IsNull(!dep4) Then
        txtdep4 = !dep4
    End If
    If Not IsNull(!TIEMPOELABORACION) Then
        txtelaboracion = !TIEMPOELABORACION
    End If
    If Not IsNull(!grafico) Then
        txtgrafico = !grafico
    End If
    If Not IsNull(!letra) Then
        txtLetra = !letra
    End If
    If Not IsNull(!CANTCONTROL) Then
        txtcontrol = !CANTCONTROL
    End If
    If Not IsNull(!observaciones) Then
        txtobserv = !observaciones
    End If
    If Not IsNull(!moneda) Then
        cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, ObtenerDescripcion("Monedas", !moneda))
    End If
    cmbactivo.ListIndex = BuscarenComboS(cmbactivo, IIf(!PUEDOFAC, "SI", "NO"))
    
    If Not IsNull(!formula) Then
        chkformula.Value = IIf(!formula, vbChecked, vbUnchecked)
    End If
    If Not IsNull(!facturable) Then
        chkFactu.Value = IIf(!facturable, vbChecked, vbUnchecked)
    Else
        chkFactu.Value = 0
    End If
    
    chkTiene_Cuenta = b2k(nSinNull(!tiene_Cuenta))
    
    uCuenta.codigo = sSinNull(!Cuenta)
    
    chk1.Value = 0
    chk2.Value = 0
    chk3.Value = 0
    chk4.Value = 0
    If s2n(!almacen) = 1 Then
        chk1.Value = 1
    ElseIf s2n(!almacen) = 2 Then
        chk2.Value = 1
    ElseIf s2n(!almacen) = 3 Then
        chk3.Value = 1
    ElseIf s2n(!almacen) = 4 Then
        chk4.Value = 1
    End If
    
  End With
End Sub

Sub HabilitoControles(habilito As Boolean)

    cboEstado.enabled = habilito
    grupo.enabled = habilito
    cmdGrupo.enabled = habilito
    cmdSubgrupo.enabled = habilito
    SubGrupo.enabled = habilito
    txtCodigo.enabled = habilito
    txtDescripcion.enabled = habilito
    txtDesc.enabled = habilito
    cboMedidas.enabled = habilito
    cboConversion.enabled = habilito
    cboCParte.enabled = habilito
    txtcostobase.enabled = habilito
    cmbcalculo.enabled = habilito
    txtcosto.enabled = habilito
    txtprelis1.enabled = habilito
    txtprelis2.enabled = habilito
    txtprelis3.enabled = habilito
    txtprelis4.enabled = habilito
    cmbproducto.enabled = habilito
    txtcostoprov.enabled = habilito
    txtstockmin.enabled = habilito
    txtpedidomin.enabled = habilito
    txtcodbarras.enabled = habilito
    cmbserie.enabled = habilito
    cboManejaStock.enabled = habilito
    chkformula.enabled = habilito
    txtIva.enabled = habilito
    txtelaboracion.enabled = habilito
    txtgrafico.enabled = habilito
    txtLetra.enabled = habilito
    txtcontrol.enabled = habilito
    txtobserv.enabled = habilito
    cmbactivo.enabled = habilito
    cmbMoneda.enabled = habilito
    txtAlias.enabled = habilito
    uCuenta.enabled = habilito
    chkTiene_Cuenta.enabled = habilito
    chkFactu.enabled = habilito
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set grupo = Nothing
    Set SubGrupo = Nothing
End Sub


Private Sub grupo_cambio(codigo As Variant)
    revisoproducto
End Sub

Private Sub SubGrupo_cambio(codigo As Variant)
    revisoproducto
End Sub

Private Sub txtcodigo_GotFocus()
    GotFocusPinto txtCodigo
End Sub

Private Sub txtcodigo_LostFocus()
    revisoproducto
End Sub

Private Sub txtcostobase_LostFocus()
    If Not IsNumeric(txtcostobase) Then
        txtcostobase = "0"
        txtcostobase.SetFocus
    End If
End Sub
Private Sub txtcosto_LostFocus()
    Dim rs As New ADODB.Recordset
    If Not IsNumeric(txtcosto) Then
        txtcosto = "0"
        txtcosto.SetFocus
    Else
        rs.Open "select * from PorcentajeListas where activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!ID) Or IsEmpty(rs!ID) Then
        Else
            txtprelis1.Text = txtcosto.Text - (txtcosto.Text * (rs!lista1 / 100))
            txtprelis2.Text = txtcosto.Text - (txtcosto.Text * (rs!lista2 / 100))
            txtprelis3.Text = txtcosto.Text - (txtcosto.Text * (rs!lista3 / 100))
            txtprelis4.Text = txtcosto.Text - (txtcosto.Text * (rs!lista4 / 100))
        End If
    End If
End Sub

Private Sub txtGrupo_LostFocus()
    revisoproducto
End Sub

Private Sub txtprelis1_LostFocus()
    If Not IsNumeric(txtprelis1) Then
        txtprelis1 = "0"
        txtprelis1.SetFocus
    End If
End Sub
Private Sub txtprelis2_LostFocus()
    If Not IsNumeric(txtprelis2) Then
        txtprelis2 = "0"
        txtprelis2.SetFocus
    End If
End Sub
Private Sub txtprelis3_LostFocus()
    If Not IsNumeric(txtprelis3) Then
        txtprelis3 = "0"
        txtprelis3.SetFocus
    End If
End Sub
Private Sub txtprelis4_LostFocus()
    If Not IsNumeric(txtprelis4) Then
        txtprelis4 = "0"
        txtprelis4.SetFocus
    End If
End Sub
Private Sub txtcostoprov_LostFocus()
    If Not IsNumeric(txtcostoprov) Then
        txtcostoprov = "0"
        txtcostoprov.SetFocus
    End If
End Sub
Private Sub txtstockmin_LostFocus()
    If Not IsNumeric(txtstockmin) Then
        txtstockmin = "0"
        txtstockmin.SetFocus
    End If
End Sub
Private Sub txtpedidomin_LostFocus()
    If Not IsNumeric(txtpedidomin) Then
        txtpedidomin = "0"
        txtpedidomin.SetFocus
    End If
End Sub
Private Sub txtiva_LostFocus()
    If Not IsNumeric(txtIva) Then
        txtIva = "0"
        txtIva.SetFocus
    End If
End Sub
Private Sub txtexistencia_LostFocus()
    If Not IsNumeric(txtExistencia) Then
        txtExistencia = "0"
        txtExistencia.SetFocus
    End If
End Sub
Private Sub txtpedidos_LostFocus()
    If Not IsNumeric(txtpedidos) Then
        txtpedidos = "0"
        txtpedidos.SetFocus
    End If
End Sub
Private Sub txtsaldo_LostFocus()
    If Not IsNumeric(txtsaldo) Then
        txtsaldo = "0"
        txtsaldo.SetFocus
    End If
End Sub

Private Sub txtdep1_LostFocus()
    If Not IsNumeric(txtdep1) Then
        txtdep1 = "0"
        txtdep1.SetFocus
    End If
End Sub
Private Sub txtdep2_LostFocus()
    If Not IsNumeric(txtdep2) Then
        txtdep2 = "0"
        txtdep2.SetFocus
    End If
End Sub
Private Sub txtdep3_LostFocus()
    If Not IsNumeric(txtdep3) Then
        txtdep3 = "0"
        txtdep3.SetFocus
    End If
End Sub
Private Sub txtdep4_LostFocus()
    If Not IsNumeric(txtdep4) Then
        txtdep4 = "0"
        txtdep4.SetFocus
    End If
End Sub
Private Sub txtcontrol_LostFocus()
    If Not IsNumeric(txtcontrol) Then
        txtcontrol = "0"
        txtcontrol.SetFocus
    End If
End Sub

Private Sub txtcosto_GotFocus()
    txtcosto.SelStart = 0
    txtcosto.SelLength = Len(txtcosto.Text)
End Sub

Private Sub txtprelis1_GotFocus()
    txtprelis1.SelStart = 0
    txtprelis1.SelLength = Len(txtprelis1.Text)
End Sub

Private Sub txtprelis2_GotFocus()
    txtprelis2.SelStart = 0
    txtprelis2.SelLength = Len(txtprelis2.Text)
End Sub

Private Sub txtprelis3_GotFocus()
    txtprelis3.SelStart = 0
    txtprelis3.SelLength = Len(txtprelis3.Text)
End Sub

Private Sub txtprelis4_GotFocus()

    txtprelis4.SelStart = 0
    txtprelis4.SelLength = Len(txtprelis4.Text)

End Sub

Private Sub txtcostoprov_GotFocus()

    txtcostoprov.SelStart = 0
    txtcostoprov.SelLength = Len(txtcostoprov.Text)

End Sub

Private Sub txtstockmin_GotFocus()

    txtstockmin.SelStart = 0
    txtstockmin.SelLength = Len(txtstockmin.Text)

End Sub

Private Sub txtpedidomin_GotFocus()

    txtpedidomin.SelStart = 0
    txtpedidomin.SelLength = Len(txtpedidomin.Text)

End Sub
Private Sub txtcostobase_GotFocus()

    txtcostobase.SelStart = 0
    txtcostobase.SelLength = Len(txtcostobase.Text)

End Sub
Private Sub txtcodbarras_GotFocus()

    txtcodbarras.SelStart = 0
    txtcodbarras.SelLength = Len(txtcodbarras.Text)

End Sub

Private Sub txtiva_GotFocus()

    txtIva.SelStart = 0
    txtIva.SelLength = Len(txtIva.Text)

End Sub
Private Sub txtexistencia_GotFocus()

    txtExistencia.SelStart = 0
    txtExistencia.SelLength = Len(txtExistencia.Text)

End Sub
Private Sub txtsaldo_GotFocus()

    txtsaldo.SelStart = 0
    txtsaldo.SelLength = Len(txtsaldo.Text)

End Sub
Private Sub txtdep2_GotFocus()

    txtdep2.SelStart = 0
    txtdep2.SelLength = Len(txtdep2.Text)

End Sub
Private Sub txtdep3_GotFocus()

    txtdep3.SelStart = 0
    txtdep3.SelLength = Len(txtdep3.Text)

End Sub

Private Sub txtdep4_GotFocus()

    txtdep4.SelStart = 0
    txtdep4.SelLength = Len(txtdep4.Text)

End Sub

Private Sub txtdep1_GotFocus()

    txtdep1.SelStart = 0
    txtdep1.SelLength = Len(txtdep1.Text)

End Sub
Private Sub txtelaboracion_GotFocus()

    txtelaboracion.SelStart = 0
    txtelaboracion.SelLength = Len(txtelaboracion.Text)

End Sub
Private Sub txtgrafico_GotFocus()

    txtgrafico.SelStart = 0
    txtgrafico.SelLength = Len(txtgrafico.Text)

End Sub
Private Sub txtletra_GotFocus()

    txtLetra.SelStart = 0
    txtLetra.SelLength = Len(txtLetra.Text)

End Sub
Private Sub txtcontrol_GotFocus()

    txtcontrol.SelStart = 0
    txtcontrol.SelLength = Len(txtcontrol.Text)

End Sub
Private Sub txtobserv_GotFocus()
    
    txtobserv.SelStart = 0
    txtobserv.SelLength = Len(txtobserv.Text)

End Sub

Private Sub txtpedidos_GotFocus()
    
    txtpedidos.SelStart = 0
    txtpedidos.SelLength = Len(txtpedidos.Text)

End Sub

Private Sub txtDescripcion_LostFocus()
    If txtDescripcion = "" Then
        MsgBox "Debe ingresar una descripción"
    End If
End Sub

Private Function CodigoGrabar() As String
    If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        CodigoGrabar = Left(grupo.codigo & "   ", 3) & Left(SubGrupo.codigo & "   ", 3) & Trim(txtCodigo)
    Else
        CodigoGrabar = Trim(txtCodigo)
    End If
End Function
Private Function CodigoGrabarAlternativo() As String
    If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        CodigoGrabarAlternativo = Left(grupo.codigo & "   ", 3) & Left(SubGrupo.codigo & "   ", 3) & Trim(mens)
    Else
        CodigoGrabarAlternativo = Trim(mens)
    End If
End Function
Private Function FaltanCosas() As Boolean
    FaltanCosas = True
    
    If Trim(txtCodigo) = "" Or Trim(txtDescripcion) = "" Then
        che "Faltan codigo y descripcion "
        Exit Function
    End If
    If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        If grupo.codigo = "" Or SubGrupo.codigo = "" Then
            che "falta Grupo y SubGrupo"
            Exit Function
        End If
    End If
    
    'If s2n(txtIva) = 0 Then
    '    MsgBox "Falta porcentaje de IVA (ej: 0.21)...", vbExclamation
    '    Exit Function
    'End If
    
    FaltanCosas = False
End Function

Private Sub revisoproducto()
Dim resu
    
    
    If ucMenu.estado = Not ucbEditando Then Exit Sub
    If ucMenu.estado = ucbMostrando Then Exit Sub
    If ucMenu.estado = ucbocioso Then Exit Sub
   
    
    If VerParametro(BS_CON_CODPROD_COMPUESTO) Then
        If grupo.codigo = "" Or SubGrupo.codigo = "" Or Trim(txtCodigo) = "" Then Exit Sub
    Else
        If Trim(txtCodigo) = "" Then Exit Sub
    End If
        
    resu = obtenerDeSQL("select codigo, activo, idProducto from producto where  codigo = '" & CodigoGrabar() & "'")
    
    If IsEmpty(resu) Then Exit Sub
    
    If resu(1) = True Then
        che "Producto ya existe"
        Exit Sub
    Else
        If confirma("Producto figura como borrado" & vbCrLf & "¿Desea activar el viejo?") Then
            DataEnvironment1.Sistema.Execute _
                "update producto set activo = 1 where idproducto = " & resu(2)
                
            ucMenu.CancelarEdicion
            ucMenu.rs.Requery

            ucMenu.BuscarOK ("idProducto = '" & resu(2) & "'")
            CargoProducto
        Else
            che "Codigo producto no puede duplicarse"
        End If
    End If
End Sub

Private Function NuevoCodeBar() As String
    Dim ultimo
    If s2n(VerParametro(BS_ALTAPROD_GENERACODEBAR)) = 1 Then  ' metodo LauMamyBlue
    
        If IsNumeric(txtCodigo) Then
            txtcodbarras = Format(txtCodigo, "0000000000") & "0000"
        Else
            ultimo = ultimoCodeBar
            txtcodbarras = Format(ultimo + 1, "7700000000") & "0000"
        End If
    End If
End Function

Private Function ultimoCodeBar() As Long
    Dim ultimo As String, proximo As Long
    ultimo = obtenerDeSQL("select max(codigobarra) from producto ")
    ultimoCodeBar = s2n(Mid(ultimo, 3, 8))
End Function

Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    If FaltanCosas() Then Exit Sub
    
    Dim Serie As Long, formula As Long, ManejaStock As Long, tiene_Cuenta_PRO As Long
    Dim factu As Long
    
    Serie = IIf(Trim(cmbserie) = "NO", 0, 1)
    ManejaStock = ComboCodigo(cboManejaStock)
    formula = IIf(chkformula.Value = 0, 0, 1)
    factu = IIf(chkFactu.Value = 0, 0, 1)
    
    If Trim(txtcodbarras) = "" Then NuevoCodeBar

    If Not IsEmpty(obtenerDeSQL("select codigo from producto where codigo = '" & CodigoGrabar() & "'")) Then
        MsgBox " Grupo + Subgrupo + Codigo  ya existe "
        Exit Sub
    End If
    
    If chkTiene_Cuenta = 0 Then
        tiene_Cuenta_PRO = 0
    Else
        tiene_Cuenta_PRO = 1
    End If
    
    'DataEnvironment1.dbo_PRODUCTOS "A", CodigoGrabar(), grupo.codigo, SubGrupo.codigo, _
        Trim(txtdescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, s2n(txtcostobase, 4), IIf(cmbcalculo.Text = "SI", 1, 0), s2n(txtcosto, 4), s2n(txtprelis1, 4), _
        s2n(txtprelis2, 4), s2n(txtprelis3, 4), s2n(txtprelis4, 4), ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text)), s2n(txtcostoprov, 4), s2n(txtstockmin, 4), _
        s2n(txtpedidomin, 4), Trim(txtcodbarras), Serie, s2n(txtiva, 4), s2n(txtExistencia), _
        s2n(txtelaboracion), Trim(txtgrafico), Trim(txtletra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)), formula, txtAlias, Date, UsuarioSistema!codigo, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO
    ABMProducto "A", CodigoGrabar(), Trim(txtCodigo), grupo.codigo, SubGrupo.codigo, _
        Trim(txtDescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, uFactor(cboConversion.ListIndex).dCodigo, uFactor(cboCParte.ListIndex).dCodigo, s2n(txtcostobase, 4), IIf(cmbcalculo.Text = "SI", 1, 0), s2n(txtcosto, 4), s2n(txtprelis1, 4), _
        s2n(txtprelis2, 4), s2n(txtprelis3, 4), s2n(txtprelis4, 4), ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text)), s2n(txtcostoprov, 4), s2n(txtstockmin, 4), _
        s2n(txtpedidomin, 4), Trim(txtcodbarras), Serie, s2n(txtIva, 4), s2n(txtExistencia), _
        s2n(txtelaboracion), Trim(txtgrafico), Trim(txtLetra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbMoneda.Text)), formula, txtAlias, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO, txtDesc.TextRTF, factu, Alma

    'MsgBox "La operación fue realizada con éxito.", vbInformation, "Producto guardado"
    ucMenu.AceptarOk ("codigo = '" & CodigoGrabar() & "'")
    
Exit Sub
ufaErr:
    MsgBox "Error al guardar", vbCritical, "Error en producto"
End Sub

Private Sub ucMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    If FaltanCosas() Then Exit Sub
    
    Dim Serie As Long, formula As Long, ManejaStock As Long, tiene_Cuenta_PRO As Long
    Dim factu As Long
    Serie = IIf(Trim(cmbserie) = "NO", 0, 1)
    ManejaStock = ComboCodigo(cboManejaStock)  'IIf(Trim(cboManejaStock) = "NO", 0, 1)
    formula = IIf(chkformula.Value = 0, 0, 1)
    factu = IIf(chkFactu.Value = 0, 0, 1)
    
    If Trim(txtcodbarras) = "" Then NuevoCodeBar
    
    If chkTiene_Cuenta = 0 Then
        tiene_Cuenta_PRO = 0
    Else
        tiene_Cuenta_PRO = 1
    End If
    
    'DataEnvironment1.dbo_PRODUCTOS "M", CodigoGrabar(), grupo.codigo, SubGrupo.codigo, _
        Trim(txtdescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, s2n(txtcostobase, 4), IIf(cmbcalculo = "SI", 1, 0), s2n(txtcosto, 4), s2n(txtprelis1, 4), _
        s2n(txtprelis2, 4), s2n(txtprelis3, 4), s2n(txtprelis4, 4), s2n(ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text))), s2n(txtcostoprov, 4), s2n(txtstockmin, 4), _
        s2n(txtpedidomin, 4), Trim(txtcodbarras), Serie, s2n(txtiva, 4), s2n(txtExistencia), _
        s2n(txtelaboracion), Trim(txtgrafico), Trim(txtletra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)), formula, txtAlias, Date, UsuarioSistema!codigo, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO
    ABMProducto "M", CodigoGrabar(), Trim(txtCodigo), grupo.codigo, SubGrupo.codigo, _
        Trim(txtDescripcion), uMedidas(cboMedidas.ListIndex).dCodigo, uFactor(cboConversion.ListIndex).dCodigo, uFactor(cboCParte.ListIndex).dCodigo, s2n(txtcostobase, 4), IIf(cmbcalculo = "SI", 1, 0), s2n(txtcosto, 4), s2n(txtprelis1, 4), _
        s2n(txtprelis2, 4), s2n(txtprelis3, 4), s2n(txtprelis4, 4), s2n(ObtenerCodigo("tipoproductos", Trim(cmbproducto.Text))), s2n(txtcostoprov, 4), s2n(txtstockmin, 4), _
        s2n(txtpedidomin, 4), Trim(txtcodbarras), Serie, s2n(txtIva, 4), s2n(txtExistencia), _
        s2n(txtelaboracion), Trim(txtgrafico), Trim(txtLetra), s2n(txtcontrol), Trim(txtobserv), IIf(cmbactivo = "SI", 1, 0), ObtenerCodigo("Monedas", Trim(cmbMoneda.Text)), formula, txtAlias, uCuenta.codigo, ManejaStock, ComboCodigo(cboEstado), tiene_Cuenta_PRO, txtDesc.TextRTF, factu, Alma
    
    
    'MsgBox "La operación fue realizada con éxito", vbInformation, "Producto actualizado"
    ucMenu.AceptarOk ("codigo = '" & CodigoGrabar() & "'")
   
Exit Sub
ufaErr:
    MsgBox "Error al guardar", vbCritical, "Error en producto"
End Sub
Private Sub ucMenu_BorrarControles()
    LimpioControles
End Sub

Private Sub ucMenu_Buscar()
    Dim resu As String
    resu = frmBuscar.MostrarSql("Select codigo as [ Codigo             ], alias as [ Alias                ], descripcion as [ Descripcion                                                                ] from producto where activo = 1")
    If resu > "" Then
        ucMenu.BuscarOK ("codigo = '" & resu & "'")
        CargoProducto
    End If
End Sub

Private Sub ucMenu_BuscarYa(que As Variant)
    Dim resu As String
    resu = obtenerDeSQL("select codigo from producto where alias = '" & que & "' and activo = 1")
    If sSinNull(resu) = "" Then Exit Sub
    ucMenu.BuscarOK ("codigo = '" & resu & "'")
    CargoProducto
End Sub


Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    'DataEnvironment1.dbo_PRODUCTOS "B", CodigoGrabar(), grupo.codigo, SubGrupo.codigo _
            , Trim(txtdescripcion), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
            , 0, "", 0, 0, 0, 0, "", "", 0, "", 0, 0, 0, "", Date, UsuarioActual(), "", 0, 0, 0
    ABMProducto "B", CodigoGrabar(), "", "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, 0, 0, 0, "", "", 0, "", 0, 0, 0, "", "", 0, 0, 0, "", 0, Alma
    grabaBitacora "B", s2n(Trim(txtCodigo)), "producto" ' OJO grababitacora es codigo numerico
    ucMenu.EliminarOK
    GoTo fin
ufaErr:
    ufa "err al eliminar", Me.Name ', Err
fin:
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitoControles sino
End Sub
Private Sub ucMenu_Modificar()
    txtCodigo.enabled = False
    'grupo.enabled = False
    'SubGrupo.enabled = False
End Sub

Private Sub ucMenu_Nuevo()
    'txtCodigo.SetFocus
End Sub

Private Sub ucMenu_SALIR()
    Unload Me
End Sub
Private Sub ucMenu_SeMovio()
    CargoProducto
End Sub

Private Function OBM(dCod As Long) As Long
Dim i As Long
OBM = 0
    For i = 0 To UBound(uMedidas)
        If dCod = uMedidas(i).dCodigo Then
            OBM = i
        End If
    Next
End Function

Private Function OBF(dCod As Long) As Long
Dim i As Long
OBF = 0
    For i = 0 To UBound(uFactor)
        If dCod = uFactor(i).dCodigo Then
            OBF = i
        End If
    Next
End Function

Private Sub CargoMedidas()
Dim rsMedidas As New ADODB.Recordset
Dim i As Long
rsMedidas.Open "select * from unidadesmedida order by umcodigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsMedidas
    If .EOF Or .BOF Then
        ReDim uMedidas(0)
        uMedidas(0).dCodigo = 0
        uMedidas(0).dDescripcion = "Sin Tipos"
        cboMedidas.AddItem uMedidas(0).dDescripcion
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            ReDim Preserve uMedidas(i)
            uMedidas(i).dCodigo = !umcodigo
            uMedidas(i).dDescripcion = sSinNull(!abreviatura)
            cboMedidas.AddItem sSinNull(!abreviatura)
            .MoveNext
        Next
    End If
    cboMedidas.ListIndex = 0
End With
End Sub

Private Sub CargoFactores(fTipo As Long)
Dim rsMedidas As New ADODB.Recordset
Dim i As Long
rsMedidas.Open "select * from umfactor where tipo=" & fTipo & " order by ufcodigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsMedidas
    If .EOF Or .BOF Then
    cboConversion.clear
    cboCParte.clear
        ReDim uFactor(0)
        uFactor(0).dCodigo = 0
        uFactor(0).dDescripcion = "Sin Factor"
        cboConversion.AddItem uFactor(0).dDescripcion
        cboCParte.AddItem uFactor(0).dDescripcion
    Else
    cboConversion.clear
    cboCParte.clear
    ReDim uFactor(0)
        .MoveFirst
        For i = 0 To .RecordCount - 1
            ReDim Preserve uFactor(i)
            uFactor(i).dCodigo = !ufcodigo
            uFactor(i).dDescripcion = !caracteristica
            cboConversion.AddItem !caracteristica
            cboCParte.AddItem !caracteristica
            .MoveNext
        Next
    End If
    cboConversion.ListIndex = 0
    cboCParte.ListIndex = 0
End With
End Sub

Public Function ABMProducto(pOPE As String, pCodigo As String, pCodigo2 As String, pGrupo As String, pSubgrupo As String, pDescripcion As String, pUnidad As Long, pFactor As Long, pParte As Long, pCosto As Double, pCsincosto As Long, pCpromedio As Double, pPrecio1 As Double, pPrecio2 As Double, pPrecio3 As Double, pPrecio4 As Double, pTipo As Long, pCproveedor As Double, pStockmin As Double, pStockpedidos As Double, pCBarra As String, pSerie As Long, PIVA As Double, pExistencia As Double, pTiempo As Double, pImagen As String, pLetra As String, pControl As Double, pObservacion As String, pPuedoFac As Long, pMoneda As Long, pFormula As Long, pAlias As String, pCuenta As String, pManejaStock As Long, pEstado As Long, pTieneCuenta As Long, DesG As String, facturable As Long, almacen As Integer) As Boolean
On Error GoTo malp
Dim ABMP  As String
ABMProducto = True
    Select Case pOPE
        Case "A":
            ABMP = "INSERT INTO PRODUCTO (CODIGO,_codigo,GRUPO,SUBGRUPO,DESCRIPCION, UMEDIDA,UFACTOR,UPARTE,COSTOBASE,CALCSINCOSTO,COSTOPROM,PRECIO,PRECIO2,PRECIO3, PRECIO4,TIPOPROD,COSTOPROV,STOCKMIN,PEDMIN, CODIGOBARRA,SERIE, IVA,EXISTENCIA,TIEMPOELABORACION,GRAFICO, LETRA,CANTCONTROL,PUEDOFAC,MONEDA,FORMULA,OBSERVACIONES,FECHA,alias, FECHA_ALTA, USUARIO_ALTA, ACTIVO, Cuenta, ManejaStock,estado,TIENE_CUENTA,DescGeneral,facturable,almacen) VALUES " _
                & "(" & ssTexto(pCodigo) & "," & ssTexto(pCodigo2) & "," & ssTexto(pGrupo) & "," & ssTexto(pSubgrupo) & "," & ssTexto(pDescripcion) & "," & pUnidad & "," & pFactor & "," & pParte & "," & x2s(pCosto) & "," & x2s(pCsincosto) & "," & x2s(pCpromedio) & "," & x2s(pPrecio1) & "," & x2s(pPrecio2) & "," & x2s(pPrecio3) & "," & x2s(pPrecio4) & "," & pTipo & "," & x2s(pCproveedor) & "," & x2s(pStockmin) & "," & x2s(pStockpedidos) _
                & "," & ssTexto(pCBarra) & "," & pSerie & "," & x2s(PIVA) & "," & x2s(pExistencia) & "," & x2s(pTiempo) & "," & ssTexto(pImagen) & "," & ssTexto(pLetra) & "," & x2s(pControl) & "," & pPuedoFac & "," & pMoneda & "," & pFormula & "," & ssTexto(pObservacion) & "," & ssFecha(Date) & ", " & ssTexto(pAlias) & ", " & ssFecha(Date) & " ,2,1," & ssTexto(pCuenta) & "," & pManejaStock & "," & pEstado & "," & pTieneCuenta & ",'" & Trim(DesG) & "'," & facturable & "," & almacen & ")"
            DataEnvironment1.Sistema.Execute ABMP
            MsgBox "Producto guardado.", vbInformation, "Alta de producto"
        Case "M":
            ABMP = " Update producto SET " _
                & " DESCRIPCION=" & ssTexto(pDescripcion) & ",GRUPO=" & ssTexto(pGrupo) & ",SUBGRUPO=" & ssTexto(pSubgrupo) & ", UMEDIDA=" & pUnidad & ", COSTOBASE=" & x2s(pCosto) & ", CALCSINCOSTO=" & x2s(pCsincosto) & ", COSTOPROM=" & x2s(pCpromedio) & ",PRECIO=" & x2s(pPrecio1) & ", FORMULA=" & (pFormula) & ",PRECIO2=" & x2s(pPrecio2) & ", PRECIO3=" & x2s(pPrecio3) & ",PRECIO4=" & x2s(pPrecio4) _
                & ", TIPOPROD=" & pTipo & ", COSTOPROV=" & x2s(pCproveedor) & ", STOCKMIN=" & x2s(pStockmin) & ", PEDMIN=" & x2s(pStockpedidos) & ", CODIGOBARRA=" & ssTexto(pCBarra) & ",SERIE=" & (pSerie) & ",IVA=" & x2s(PIVA) & ",EXISTENCIA=" & x2s(pExistencia) & ", TIEMPOELABORACION=" & x2s(pTiempo) & ",GRAFICO=" & ssTexto(pImagen) & ",LETRA=" & ssTexto(pLetra) & ", CANTCONTROL=" & x2s(pControl) _
                & ",PUEDOFAC=" & pPuedoFac & ", MONEDA=" & pMoneda & ",OBSERVACIONES=" & ssTexto(pObservacion) & ", alias = " & ssTexto(pAlias) & ", cuenta = " & ssTexto(pCuenta) & ", ManejaStock = " & pManejaStock & " , estado = " & pEstado & ",TIENE_CUENTA=" & pTieneCuenta & ",UFACTOR=" & pFactor & ",UPARTE=" & pParte & ",DescGeneral='" & Trim(DesG) & "',facturable=" & facturable & ", almacen=" & almacen _
                & " WHERE  CODIGO=" & ssTexto(pCodigo)
            DataEnvironment1.Sistema.Execute ABMP
            MsgBox "Producto Modificado.", vbInformation, "Modificacion de producto"
        Case "B":
            ABMP = " Update producto SET " _
                & " ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA=2 " _
                & " WHERE  CODIGO=" & ssTexto(pCodigo)
            DataEnvironment1.Sistema.Execute ABMP
            MsgBox "Producto Elimminado.", vbInformation, "Baja de producto"
    End Select
Exit Function
malp:
    MsgBox "Error en ABM producto", vbCritical
    ABMProducto = False
End Function

Private Function Alma() As Integer
    Alma = 0
    If chk1.Value = 1 Then
        Alma = 1
    ElseIf chk2.Value = 1 Then
        Alma = 2
    ElseIf chk3.Value = 1 Then
        Alma = 3
    ElseIf chk4.Value = 1 Then
        Alma = 4
    End If
End Function
