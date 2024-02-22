VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmAbmClientes1 
   Caption         =   "Clientes"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18525
   Icon            =   "FrmAbmClientes1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   18525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar Afip"
      Height          =   360
      Left            =   9495
      TabIndex        =   115
      Top             =   3915
      Width           =   1290
   End
   Begin VB.TextBox txtNroCliente 
      Height          =   285
      Left            =   7575
      TabIndex        =   113
      Top             =   3360
      Width           =   2400
   End
   Begin VB.TextBox txtPais 
      Height          =   285
      Left            =   2175
      TabIndex        =   111
      Top             =   3375
      Width           =   3015
   End
   Begin VB.TextBox txtCodPais 
      Height          =   285
      Left            =   1545
      TabIndex        =   110
      Top             =   3375
      Width           =   600
   End
   Begin VB.CommandButton cmdBuscarPais 
      Caption         =   "pais"
      Height          =   315
      Left            =   5235
      TabIndex        =   109
      Top             =   3375
      Width           =   540
   End
   Begin VB.CommandButton cmdBuscarCuit 
      Caption         =   "pais"
      Height          =   315
      Left            =   7800
      TabIndex        =   108
      Top             =   540
      Width           =   435
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2640
      TabIndex        =   107
      Top             =   1590
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   106
      Top             =   1590
      Width           =   735
   End
   Begin VB.TextBox txtRUC 
      Height          =   315
      Left            =   8940
      TabIndex        =   105
      Top             =   525
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddCtasVentas 
      Height          =   315
      Left            =   16080
      Picture         =   "FrmAbmClientes1.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   2835
      Width           =   345
   End
   Begin VB.CommandButton cmdDelCtasVentas 
      Height          =   315
      Left            =   16455
      Picture         =   "FrmAbmClientes1.frx":2284
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   2835
      Width           =   345
   End
   Begin VB.CommandButton cmdLocBusco 
      Caption         =   "Localidad"
      Height          =   300
      Left            =   9915
      TabIndex        =   99
      ToolTipText     =   "Seleccione primero la provincia."
      Top             =   915
      Width           =   1005
   End
   Begin VB.CheckBox chkPercIIBB_p 
      Alignment       =   1  'Right Justify
      Caption         =   "Con Perc IIBB Personal"
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
      Height          =   240
      Left            =   6390
      TabIndex        =   98
      Top             =   3735
      Width           =   2535
   End
   Begin VB.CheckBox chkPercGAN_p 
      Alignment       =   1  'Right Justify
      Caption         =   "Con Perc GAN Personal"
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
      Left            =   6405
      TabIndex        =   97
      Top             =   4035
      Width           =   2535
   End
   Begin VB.Frame frameIIBB 
      Caption         =   "Percepcion de IIBB"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   11340
      TabIndex        =   91
      Top             =   225
      Width           =   7080
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   105
         TabIndex        =   93
         Text            =   "0"
         Top             =   600
         Width           =   1635
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1860
         TabIndex        =   92
         Text            =   "0"
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label40 
         Caption         =   "<- Ingresar valores por ejemplo : 0.21 (= 21 %)"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3630
         TabIndex        =   96
         Top             =   555
         Width           =   3345
      End
      Begin VB.Label Label41 
         Caption         =   "Bases imponibles"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   135
         TabIndex        =   95
         Top             =   375
         Width           =   1605
      End
      Begin VB.Label Label42 
         Caption         =   "Coeficiente"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1860
         TabIndex        =   94
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.Frame frameGAN 
      Caption         =   "Percepcion de Ganancias"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   11355
      TabIndex        =   85
      Top             =   1695
      Width           =   7050
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1845
         TabIndex        =   87
         Text            =   "0"
         Top             =   510
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   90
         TabIndex        =   86
         Text            =   "0"
         Top             =   510
         Width           =   1635
      End
      Begin VB.Label Label38 
         Caption         =   "Coeficiente"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1845
         TabIndex        =   90
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label39 
         Caption         =   "Bases imponibles"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   135
         TabIndex        =   89
         Top             =   285
         Width           =   1605
      End
      Begin VB.Label Label43 
         Caption         =   "<- Ingresar valores por ejemplo : 0.21 (= 21 %)"
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   3645
         TabIndex        =   88
         Top             =   480
         Width           =   3315
      End
   End
   Begin Gestion.ucCuit uCuit 
      Height          =   315
      Left            =   6120
      TabIndex        =   84
      Top             =   525
      Width           =   1605
      _extentx        =   2831
      _extenty        =   556
   End
   Begin VB.TextBox txtCuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1425
      TabIndex        =   83
      Text            =   "0"
      Top             =   6165
      Width           =   1305
   End
   Begin VB.CheckBox chkTiene_Cuenta 
      Caption         =   "Usar Cuenta"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8010
      TabIndex        =   82
      Top             =   6135
      Width           =   1695
   End
   Begin VB.TextBox txtCuenta_Descrip 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2745
      TabIndex        =   81
      Top             =   6165
      Width           =   5220
   End
   Begin VB.CommandButton cmdCuenta 
      Caption         =   "Cuenta"
      Enabled         =   0   'False
      Height          =   300
      Left            =   435
      TabIndex        =   80
      Top             =   6165
      Width           =   975
   End
   Begin VB.CheckBox chkConPercIIBB 
      Alignment       =   1  'Right Justify
      Caption         =   "Aplicar Perc IIBB"
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
      Left            =   4320
      TabIndex        =   30
      Top             =   3720
      Width           =   1950
   End
   Begin Gestion.ucBotonera ucMenu 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   1410
      Left            =   11085
      TabIndex        =   79
      Top             =   5040
      Width           =   7440
      _extentx        =   13123
      _extenty        =   2487
      msgconfirmasalir=   "¿Quiere cerrar la ventana?"
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VB.ComboBox cmbprovincias 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "5"
      Top             =   1245
      Width           =   1815
   End
   Begin VB.CheckBox chkconsig 
      Alignment       =   1  'Right Justify
      Caption         =   "Consignatario"
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
      Left            =   2715
      TabIndex        =   28
      Top             =   3720
      Width           =   1545
   End
   Begin VB.CheckBox chkmay 
      Alignment       =   1  'Right Justify
      Caption         =   "Mayorista"
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
      Left            =   2940
      TabIndex        =   29
      Top             =   4035
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   7110
      TabIndex        =   6
      Top             =   915
      Width           =   2745
   End
   Begin VB.TextBox txtcodprov 
      Height          =   285
      Left            =   9300
      TabIndex        =   20
      Top             =   2325
      Width           =   1215
   End
   Begin VB.TextBox txtlimite 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6765
      TabIndex        =   19
      Top             =   2310
      Width           =   1065
   End
   Begin VB.CheckBox chketiqueta 
      Alignment       =   1  'Right Justify
      Caption         =   "Etiquetas"
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
      Left            =   1455
      TabIndex        =   27
      Top             =   4050
      Width           =   1155
   End
   Begin VB.TextBox txtweb 
      Height          =   285
      Left            =   5280
      TabIndex        =   23
      Top             =   3000
      Width           =   2400
   End
   Begin VB.TextBox txtmail 
      Height          =   285
      Left            =   1575
      TabIndex        =   22
      Top             =   2985
      Width           =   3090
   End
   Begin VB.TextBox txtdescuento2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9300
      MaxLength       =   3
      TabIndex        =   17
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox txtdescuento1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6825
      MaxLength       =   3
      TabIndex        =   16
      Top             =   1950
      Width           =   1005
   End
   Begin VB.TextBox txtdireccioncom 
      Height          =   285
      Left            =   1380
      TabIndex        =   31
      Top             =   4890
      Width           =   3360
   End
   Begin VB.TextBox txtfaxcom 
      Height          =   285
      Left            =   4290
      TabIndex        =   38
      Top             =   5535
      Width           =   2415
   End
   Begin VB.TextBox txttelcom 
      Height          =   285
      Left            =   1380
      TabIndex        =   37
      Top             =   5535
      Width           =   2415
   End
   Begin VB.TextBox txtbarriocom 
      Height          =   285
      Left            =   3810
      TabIndex        =   35
      Top             =   5190
      Width           =   2520
   End
   Begin VB.TextBox txtlocalidadcom 
      Height          =   285
      Left            =   7215
      TabIndex        =   33
      Top             =   4890
      Width           =   2220
   End
   Begin VB.TextBox txtcontactocom 
      Height          =   285
      Left            =   1380
      TabIndex        =   39
      Top             =   5835
      Width           =   4695
   End
   Begin VB.ComboBox cmbzonascom 
      Height          =   315
      Left            =   7215
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   5175
      Width           =   2265
   End
   Begin VB.ComboBox cmbprovinciacom 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   5190
      Width           =   1785
   End
   Begin VB.ComboBox cmblista 
      Height          =   315
      Left            =   3555
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1920
      Width           =   2130
   End
   Begin VB.TextBox txthorario 
      Height          =   285
      Left            =   6885
      TabIndex        =   40
      Top             =   5850
      Width           =   2355
   End
   Begin VB.CheckBox chkcertificado 
      Alignment       =   1  'Right Justify
      Caption         =   "Certificado"
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
      Left            =   150
      TabIndex        =   24
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CheckBox chkhabilitado 
      Alignment       =   1  'Right Justify
      Caption         =   "Habilitado"
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
      Left            =   150
      TabIndex        =   25
      Top             =   4065
      Width           =   1275
   End
   Begin VB.CheckBox chkcorreo 
      Alignment       =   1  'Right Justify
      Caption         =   "Correo"
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
      Left            =   1695
      TabIndex        =   26
      Top             =   3720
      Width           =   915
   End
   Begin VB.ComboBox cmbtransporte 
      Height          =   315
      Left            =   7125
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1605
      Width           =   2310
   End
   Begin VB.ComboBox cmbcategoria 
      Height          =   315
      Left            =   8865
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2685
      Width           =   2025
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1155
      TabIndex        =   4
      Top             =   885
      Width           =   3465
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   3375
      TabIndex        =   1
      Top             =   210
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   9165
      TabIndex        =   41
      Top             =   210
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   158466049
      CurrentDate     =   38052
   End
   Begin VB.ComboBox cmbivas 
      Height          =   315
      ItemData        =   "FrmAbmClientes1.frx":280E
      Left            =   8880
      List            =   "FrmAbmClientes1.frx":2815
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1245
      Width           =   2055
   End
   Begin VB.ComboBox cmbzonas 
      Height          =   315
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1260
      Width           =   1815
   End
   Begin VB.ComboBox cmbvendedores 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   1680
   End
   Begin VB.ComboBox cmbformaspagos 
      Height          =   315
      Left            =   1590
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2295
      Width           =   3615
   End
   Begin VB.TextBox txtcontacto 
      Height          =   285
      Left            =   1590
      TabIndex        =   21
      Top             =   2655
      Width           =   6090
   End
   Begin VB.TextBox txtfantasia 
      Height          =   285
      Left            =   2220
      TabIndex        =   2
      Top             =   525
      Width           =   3525
   End
   Begin VB.TextBox txtbarrio 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   1245
      Width           =   1965
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   1290
      TabIndex        =   11
      Top             =   1590
      Width           =   375
   End
   Begin VB.TextBox txtfax 
      Height          =   285
      Left            =   3945
      TabIndex        =   12
      Top             =   1590
      Width           =   1980
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtcodpostal 
      Height          =   300
      Left            =   5070
      TabIndex        =   5
      Top             =   885
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "?9999???"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtcodpostalcom 
      Height          =   300
      Left            =   5160
      TabIndex        =   32
      Top             =   4890
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "?9999???"
      PromptChar      =   "_"
   End
   Begin VSFlex7LCtl.VSFlexGrid gCtasVentas 
      Height          =   1290
      Left            =   11340
      TabIndex        =   102
      Top             =   3180
      Width           =   5460
      _cx             =   9631
      _cy             =   2275
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label49 
      Caption         =   "Nro Cliente:"
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
      Left            =   6240
      TabIndex        =   114
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label48 
      Caption         =   "Pais :"
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
      Left            =   630
      TabIndex        =   112
      Top             =   3375
      Width           =   975
   End
   Begin VB.Label Label47 
      Caption         =   "RUC : "
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
      Left            =   8400
      TabIndex        =   104
      Top             =   555
      Width           =   480
   End
   Begin VB.Label Label46 
      Caption         =   "CUENTAS PARA FACTURA VENTA"
      Height          =   345
      Left            =   11370
      TabIndex        =   103
      Top             =   2895
      Width           =   3195
   End
   Begin VB.Label Label37 
      Caption         =   "CP:"
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
      Left            =   4785
      TabIndex        =   78
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "Localidad:"
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
      Left            =   6045
      TabIndex        =   77
      Top             =   915
      Width           =   975
   End
   Begin VB.Label Label36 
      Caption         =   "Localidad:"
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
      Left            =   6225
      TabIndex        =   76
      Top             =   4905
      Width           =   975
   End
   Begin VB.Label Label35 
      Caption         =   "Cod.Proveedor:"
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
      Left            =   7860
      TabIndex        =   75
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label34 
      Caption         =   "Limite de Credito:"
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
      Left            =   5220
      TabIndex        =   74
      Top             =   2310
      Width           =   1650
   End
   Begin VB.Label Label33 
      Caption         =   "Web:"
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
      Left            =   4785
      TabIndex        =   73
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label32 
      Caption         =   "E-Mail :"
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
      Left            =   900
      TabIndex        =   72
      Top             =   2985
      Width           =   735
   End
   Begin VB.Label Label31 
      Caption         =   "Descuento2:"
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
      Left            =   8145
      TabIndex        =   71
      Top             =   1965
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Descuento1:"
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
      Left            =   5700
      TabIndex        =   70
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "Telefono/s:"
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
      Left            =   330
      TabIndex        =   69
      Top             =   5535
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "Barrio :"
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
      Left            =   3225
      TabIndex        =   68
      Top             =   5220
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "Provincia :"
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
      Left            =   450
      TabIndex        =   67
      Top             =   5190
      Width           =   1050
   End
   Begin VB.Label Label26 
      Caption         =   "Direccion :"
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
      Left            =   420
      TabIndex        =   66
      Top             =   4905
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "Zona:"
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
      Left            =   6660
      TabIndex        =   65
      Top             =   5220
      Width           =   615
   End
   Begin VB.Label Label23 
      Caption         =   "Fax:"
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
      Left            =   3825
      TabIndex        =   64
      Top             =   5550
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "Contacto :"
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
      Left            =   495
      TabIndex        =   63
      Top             =   5865
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Lista :"
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
      Left            =   3015
      TabIndex        =   62
      Top             =   1950
      Width           =   735
   End
   Begin VB.Label Label20 
      Caption         =   "Horario :"
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
      Left            =   6120
      TabIndex        =   61
      Top             =   5865
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   " Datos Comerciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   105
      TabIndex        =   60
      Top             =   4545
      Width           =   2145
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1845
      Left            =   30
      Top             =   4680
      Width           =   10965
   End
   Begin VB.Label Label18 
      Caption         =   "Transporte :"
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
      Left            =   6030
      TabIndex        =   59
      Top             =   1635
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Categoria :"
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
      Left            =   7905
      TabIndex        =   58
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "Contacto :"
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
      TabIndex        =   57
      Top             =   2655
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Fax:"
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
      Left            =   3510
      TabIndex        =   56
      Top             =   1590
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   4335
      Left            =   75
      Top             =   120
      Width           =   10980
   End
   Begin VB.Label Label14 
      Caption         =   "Tipo Iva :"
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
      Left            =   8040
      TabIndex        =   55
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Zona:"
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
      Left            =   5520
      TabIndex        =   54
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label12 
      Caption         =   "CP:"
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
      Left            =   4725
      TabIndex        =   53
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Codigo: "
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
      Height          =   270
      Left            =   150
      TabIndex        =   52
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Direccion:"
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
      Left            =   135
      TabIndex        =   51
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Razon Social:"
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
      Left            =   2055
      TabIndex        =   50
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Cuit : "
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
      Left            =   5745
      TabIndex        =   49
      Top             =   585
      Width           =   480
   End
   Begin VB.Label Label7 
      Caption         =   "Forma de Pago:"
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
      TabIndex        =   48
      Top             =   2310
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Provincia :"
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
      TabIndex        =   47
      Top             =   1275
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha"
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
      Left            =   8505
      TabIndex        =   46
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Barrio :"
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
      Left            =   2880
      TabIndex        =   45
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label Label3 
      Caption         =   "Telefono/s:"
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
      Left            =   165
      TabIndex        =   44
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Vendedor :"
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
      TabIndex        =   43
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Fantasia: "
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
      Left            =   165
      TabIndex        =   42
      Top             =   525
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAbmClientes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cadena_iu As String

Private Sub chkPercGAN_p_Click()
    If chkPercGAN_p Then
        frameGAN.enabled = True
    Else
        frameGAN.enabled = False
    End If
End Sub

Private Sub chkPercIIBB_p_Click()
    If chkPercIIBB_p Then
        frameIIBB.enabled = True
    Else
        frameIIBB.enabled = False
    End If
End Sub

Private Sub cmbprovincias_Click()
    cmbprovinciacom.ListIndex = cmbprovincias.ListIndex
End Sub

Private Sub cmdAddCtasVentas_Click()
Dim Res
Dim rr As Long
    Res = frmBuscar.MostrarSql("Select [Cuenta              ],[Descripcion                                    ] from Cuentas where imputable=1")
    If Res > "" Then
        gCtasVentas.AddItem ""
        rr = gCtasVentas.rows - 1
        gCtasVentas.TextMatrix(rr, 0) = Res
        gCtasVentas.TextMatrix(rr, 1) = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdBuscarCuit_Click()
If Trim(cmbprovincias.Text) = "EXTERIOR" Then
    Dim tmp
    tmp = frmBuscar.MostrarSql("SELECT CUITPAIS AS [CUIT         ], [DESCRIPCION          ] FROM FECUIT")
    If tmp > "" Then
        uCuit.Text = Format(tmp, "00-00000000-0")
    End If
End If
End Sub

Private Sub cmdBuscarPais_Click()
    Dim tmp
    tmp = frmBuscar.MostrarSql("SELECT CODIGO AS [CODIGO         ], [DESCRIPCION          ] FROM FEPAIS")
    If tmp > "" Then
        txtCodPais = tmp
        txtpais = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdConsultar_Click()
On Error Resume Next
Dim Afip As New FacturaElectronica, sCuit As String
Dim respuesta
Dim nEsGranEmpresa As Boolean, CodGE As Integer
sCuit = uCuit.Text
sCuit = Replace(sCuit, "-", "")
If Trim(sCuit) = "" Then Exit Sub

respuesta = Afip.ConsultarDatos(sCuit)
nEsGranEmpresa = Afip.EstaEnPadron(sCuit)

txtNombre = respuesta(0) & " " & respuesta(1)
txtfantasia = respuesta(2)
txtLocalidad = respuesta(3)
txtdireccion = respuesta(4)
'Respuesta(5) = ClienteTipoDomicilio
txtcodpostal = CInt(respuesta(6))
cmbprovincias.ListIndex = BuscarenComboS(cmbprovincias, ObtenerDescripcionS("provincias", BuscoCodigoProvincia(CStr(respuesta(7)))))
If nEsGranEmpresa Then
    CodGE = Afip.CodigoIvaGranEmpresa
    cmbivas.Text = obtenerDeSQL("select descripcion from ivas where codigo=" & CodGE)
Else
    If respuesta(8) = "Monotributo-Si" Then cmbivas.Text = "Monotributista"
    If respuesta(9) = "Inscripto-Si" Then cmbivas.Text = "Responsable Inscripto"
    If respuesta(10) = "Exento-Si" Then cmbivas.Text = "Exento"
End If


End Sub

Private Sub cmdCuenta_Click()
txtcuenta = BuscarCuenta(False, False)
If txtcuenta > "" Then txtCuenta_Descrip = obtenerDeSQL("select descripcion from cuentas where cuenta= " & txtcuenta)
End Sub

Private Sub cmdDelCtasVentas_Click()
    If gCtasVentas.rows > 0 Then
        If gCtasVentas.Row >= 0 Then
            gCtasVentas.RemoveItem gCtasVentas.Row
        End If
    End If
End Sub

Private Sub cmdLocBusco_Click()
Dim resu As Variant, CodigoProv As String
    If cmbprovincias.Text <> "" Then
        CodigoProv = "'" & obtenerDeSQL("select codigo from provincias where descripcion = '" & cmbprovincias.Text & "'") & "'"
        resu = frmBuscar.MostrarSql("select numero as NUM,localidad as LOCALIDAD , partido as PARTIDO from localidades where provincia = " & CodigoProv, , "Localidades", "-")
    Else
        resu = frmBuscar.MostrarSql("select numero as NUM,localidad as LOCALIDAD , partido as PARTIDO from localidades", , "Localidades", "-")
    End If
    If resu > "" Then
        txtLocalidad = obtenerDeSQL("select localidad from localidades where numero = " & resu)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, False, True
End Sub

Private Sub Form_Load()
   
    CargaCombo cmbzonas, "Zonas", "descripcion", "codigo", ""
    CargaCombo cmbprovincias, "provincias", "descripcion", "codigo", ""
    CargaCombo cmbprovinciacom, "provincias", "descripcion", "codigo", ""
    CargaCombo cmbzonascom, "Zonas", "descripcion", "codigo", ""
    CargaCombo cmbTransporte, "Transportes", "descripcion", "codigo", ""
    CargaCombo cmbvendedores, "usuarios", "descripcion", "codigo", ""
    CargaCombo cmbivas, "Ivas", "descripcion", "codigo", ""
    CargaCombo cmbcategoria, "categclie", "descripcion", "codigo", ""
    CargaCombo cmbformaspagos, "formaspago", "descripcion", "codigo", ""
    CargaCombo cmblista, "Listas", "descripcion", "codigo", ""
    dtFecha = Date
    
    ucMenu.init True, True, True, False, True, "select * from Clientes where activo = 1 order by codigo", DataEnvironment1.Sistema
    ucMenu.MsgConfirmaEliminar = "Elimina Cliente ? "
    ucMenu.MsgConfirmaSalir = "Cerrar formulario ? "
       
End Sub

Private Sub IngresoCuit1_CuitInvalido(Nro As String)
    MsgBox "Cuit invalido"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodigo_LostFocus()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset

    rs.Open "Select * from clientes where codigo=" & val(Trim(txtCodigo)), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "El codigo ya existe,verifiquelo", 48, "Atencion"
    End If

fin:
    Set rs = Nothing
    Exit Sub
ufaErr:
    'ufa
    Resume fin
End Sub

Private Sub txtcodpostal_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodpostalcom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodprov_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtdescuento1_LostFocus()
    If val(txtdescuento1) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtdescuento1.SetFocus
    End If
End Sub

Private Sub txtdescuento2_LostFocus()
    If val(txtdescuento2) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtdescuento2.SetFocus
    End If
End Sub

Private Sub txtDireccion_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtdireccioncom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfantasia_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfax_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfaxcom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txthorario_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 47 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub

Sub HabilitoTxt(habilito As Boolean)

    txtCodigo.Locked = habilito
    txtbarrio.Locked = habilito
    txtbarriocom.Locked = habilito
    Txtcontacto.Locked = habilito
    txtlimite.Locked = habilito
    txtcontactocom.Locked = habilito
    txtdescuento1.Locked = habilito
    txtdescuento2.Locked = habilito
    txtdireccion.Locked = habilito
    txtdireccioncom.Locked = habilito
    txtfantasia.Locked = habilito
    txtFax.Locked = habilito
    txtfaxcom.Locked = habilito
    txthorario.Locked = habilito
    txtLocalidad.Locked = habilito
    txtlocalidadcom.Locked = habilito
    txtmail.Locked = habilito
    txtNombre.Locked = habilito
    txttel.Locked = habilito
    txttelcom.Locked = habilito
    Text3.Locked = habilito
    Text6.Locked = habilito
    txtweb.Locked = habilito
    txtNroCliente.Locked = habilito
    cmbcategoria.Locked = habilito
    cmbformaspagos.Locked = habilito
    cmbivas.Locked = habilito
    cmblista.Locked = habilito
    cmbprovincias.Locked = habilito
    cmbprovinciacom.Locked = habilito
    cmbvendedores.Locked = habilito
    cmbzonas.Locked = habilito
    cmbzonascom.Locked = habilito
    cmbcategoria.Locked = habilito
    cmbTransporte.Locked = habilito
    chkcertificado.enabled = Not habilito
    chkhabilitado.enabled = Not habilito
    chkcorreo.enabled = Not habilito
    chketiqueta.enabled = Not habilito
    txtcodpostal.enabled = Not habilito
    txtcodpostalcom.enabled = Not habilito
    uCuit.enabled = Not habilito
    cmdCuenta.enabled = Not habilito
    chkTiene_Cuenta.enabled = Not habilito
    
    chkPercIIBB_p.enabled = Not habilito
    chkPercGAN_p.enabled = Not habilito
End Sub
Sub LimpioTxt()
    On Error Resume Next

    txtCodigo = ""
    txtbarrio = ""
    txtbarriocom = ""
    Txtcontacto = ""
    txtcontactocom = ""
    txtdescuento1 = "0.00"
    txtdescuento2 = "0.00"
    txtdireccion = ""
    txtdireccioncom = ""
    txtfantasia = ""
    txtFax = ""
    txtfaxcom = ""
    txthorario = ""
    txtLocalidad = ""
    txtlocalidadcom = ""
    txtmail = ""
    txtNombre = ""
    txtlimite = "0.00"
    txttel = ""
    txttelcom = ""
    Text3 = ""
    Text6 = ""
    txtweb = ""
    txtNroCliente = ""
    cmbcategoria.ListIndex = 0
    cmbformaspagos.ListIndex = 0
    cmbivas.ListIndex = 0
    cmblista.ListIndex = 0
    cmbprovincias.ListIndex = 1
    cmbprovinciacom.ListIndex = 1
    cmbvendedores.ListIndex = 0
    cmbzonas.ListIndex = -1
    cmbzonascom.ListIndex = -1
    cmbTransporte.ListIndex = 0
    chkcertificado.Value = 0
    chkhabilitado.Value = 1
    chkcorreo.Value = 0
    chketiqueta.Value = 0
'    MaskCuit.Mask = "  -       - "
    uCuit.Text = ""
    txtcodpostal.Mask = "       "
    txtcodpostalcom.Mask = "       "
    txtcodpostal.Mask = "?9999???"
    txtcodpostalcom.Mask = "?9999???"
    
    txtcuenta = 0
    txtCuenta_Descrip = ""
    chkTiene_Cuenta.Value = 0
    
    '16-11/07 'ret iibb personal para cada prov
    chkPercIIBB_p = 0
    frameIIBB.enabled = False
    chkPercGAN_p = 0
    frameGAN.enabled = False
    Text1.Text = 0
    Text4.Text = 0
    Text2.Text = 0
    Text5.Text = 0
    
    gCtasVentas.rows = 0
    
End Sub

Private Sub txtcodigo_GotFocus()
'    txtcodigo.SelStart = 0
'    txtcodigo.SelLength = Len(txtcodigo.text)
    frmPintoFoco Me
End Sub
Private Sub txtlocalidad_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtlocalidadcom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtlimite_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtnombre_LostFocus()
Dim rs As New ADODB.Recordset

    rs.Open "Select * from clientes where descripcion='" & Trim(txtNombre) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "La razon Social ya existe,es del cliente nro: " & rs!codigo
    End If
End Sub

Private Sub txtnombre_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub Txtweb_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub Txtnrocliente_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub Txtmail_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txttel_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txttelcom_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtcontacto_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtcontactocom_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtbarrio_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtbarriocom_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtDescuento1_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtDescuento2_GotFocus()
    frmPintoFoco Me
End Sub

Sub CargoRegistro()
    On Error Resume Next
    Dim tel
    ReDim tel(3)
    
    LimpioTxt
    With ucMenu.rs
        txtCodigo = !codigo
        txtbarrio = !barrio
        txtbarriocom = !barrio_comercial
        Txtcontacto = !contacto
        txtcontactocom = !contacto_comercial
        txtdireccion = !direccion
        txtdireccioncom = !direccion_comercial
        txtfantasia = !nombrefantasia
        txtFax = !fax
        txtdescuento1 = s2n(!descuento1)
        txtdescuento2 = s2n(!descuento2)
        txtlimite = s2n(!limitecredito)
        txtfaxcom = !fax_comercial
        txthorario = !horario
        txtLocalidad = !Localidad
        txtlocalidadcom = !localidad_comercial
        txtmail = !mail
        txtNombre = !DESCRIPCION
        tel = 0
        tel = Split(!Telefono, "-")
        If tel <> "" Then
            If UBound(tel) = 0 Then
                Text6 = tel(0)
            ElseIf UBound(tel) = 1 Then
                Text3 = tel(0)
                Text6 = tel(1)
            ElseIf UBound(tel) = 2 Then
                txttel = tel(0)
                Text3 = tel(1)
                Text6 = tel(2)
            End If
        End If
        'txttel = !Telefono
        TxtCodProv = !Proveedor
        txttelcom = !telefono_comercial
        txtweb = !web
        txtNroCliente = sSinNull(!NroCliente)
        cmbcategoria.ListIndex = BuscarenComboS(cmbcategoria, ObtenerDescripcion("categclie", !Categoria))
        cmbformaspagos.ListIndex = BuscarenComboS(cmbformaspagos, ObtenerDescripcion("formaspago", !formaPago))
        cmbivas.ListIndex = BuscarenComboS(cmbivas, ObtenerDescripcion("ivas", !Iva))
        cmbprovincias.ListIndex = BuscarenComboS(cmbprovincias, ObtenerDescripcionS("provincias", !Provincia))
        cmbprovinciacom.ListIndex = BuscarenComboS(cmbprovinciacom, ObtenerDescripcionS("provincias", !provincia_comercial))
        cmbzonas.ListIndex = BuscarenComboS(cmbzonas, ObtenerDescripcion("zonas", !zona))
        cmbTransporte.ListIndex = BuscarenComboS(cmbTransporte, ObtenerDescripcion("transportes", !Transporte))
        cmbvendedores.ListIndex = BuscarenComboS(cmbvendedores, ObtenerDescripcion("usuarios", !Vendedor))
        cmblista.ListIndex = BuscarenComboS(cmblista, ObtenerDescripcion("listas", !Lista))
        cmbzonascom.ListIndex = BuscarenComboS(cmbzonascom, ObtenerDescripcion("zonas", !zonacomercial))
        txtcodpostal = !codigopostal
        txtcodpostalcom = !codigopostal_comercial
'        MaskCuit = !cuit
        uCuit.Text = !CUIT
        chkcertificado.Value = b2k(!Certificado)
        chkconsig.Value = b2k(!consignatario)
        chkmay.Value = b2k(!mayorista)
        chkcorreo.Value = b2k(!Correo)
        chkhabilitado.Value = b2k(!puedofacturar)
        chkConPercIIBB.Value = b2k(!ConPercIIBB)
        chketiqueta.Value = b2k(!etiqueta)
        txtcuenta = s2n(!Cuenta)
        If txtcuenta > "" Then txtCuenta_Descrip = obtenerDeSQL("select descripcion from cuentas where cuenta= " & txtcuenta)
        chkTiene_Cuenta = b2k(!tiene_Cuenta)
        
        chkPercIIBB_p = b2k(!conperciibbper)
        If chkPercIIBB_p Then
            Text1 = obtenerDeSQL("select baseimponible from ClieTipoRetIB_Per where codclie = " & txtCodigo)
            Text4 = obtenerDeSQL("select coeficiente from ClieTipoRetIB_Per where codclie = " & txtCodigo)
        End If
        chkPercGAN_p = b2k(!conpercganper)
        If chkPercGAN_p Then
            Text2 = obtenerDeSQL("select baseimponible from ClieTipoRetgan_Per where codclie = " & txtCodigo)
            Text5 = obtenerDeSQL("select coeficiente from ClieTipoRetgan_Per where codclie = " & txtCodigo)
        End If
        CtasVentasSet sSinNull(!cuentasventas)
        txtRUC = sSinNull(!RUC)
    End With
End Sub

Private Function CtasVentasGet() As String
Dim i As Long, C As String
Se_Exedio:
    If gCtasVentas.rows > 0 Then
        C = ""
        For i = 0 To gCtasVentas.rows - 1
            If i = 0 Then
                C = "#" & gCtasVentas.TextMatrix(i, 0) & "#"
            Else
                C = C & ",#" & gCtasVentas.TextMatrix(i, 0) & "#"
            End If
        Next
    Else
        C = ""
    End If
    If Len(C) > 4000 Then
        MsgBox "Se exede de la cantidad de cuentas permitidas. se quitara la ultima una para continuar.", vbExclamation
        gCtasVentas.RemoveItem gCtasVentas.rows - 1
        GoTo Se_Exedio
    End If

CtasVentasGet = C
End Function

Private Function CtasVentasSet(C As String)
Dim i As Long, rr As Long, CT
C = Trim(C)
    If C > "" Then
        CT = Split(Replace(C, "#", ""), ",")
        iniCtasVentas
        For i = 0 To UBound(CT)
            gCtasVentas.AddItem ""
            rr = gCtasVentas.rows - 1
            gCtasVentas.TextMatrix(rr, 0) = CT(i)
            gCtasVentas.TextMatrix(rr, 1) = obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(CT(i)))
        Next
    Else
        iniCtasVentas
    End If
End Function

Private Function iniCtasVentas()
With gCtasVentas
    .rows = 0
    .cols = 0
    .cols = 2
    .ColWidth(0) = 1200
    .ColWidth(1) = 3500
End With
End Function

Private Function FaltanCosas(Optional modi As Boolean = False) As Boolean
    Dim tmp, s As String, iv As Long

    'nombre, codigo
    If s2n(txtCodigo) = 0 Or Trim(txtNombre) = "" Then
        che "falta cargar codigo y nombre cliente"
        FaltanCosas = True
        Exit Function
    End If
    
    iv = cmbivas.ListIndex
    If iv <> 0 And iv <> 1 And iv <> 2 Then ' 0 = CONSUMIDOR FINAL, 1 y 2= EXENTOS
        If uCuit.Text = "" Then
            che "CUIT Incorrecto"
            FaltanCosas = True
            Exit Function
        End If
        
        If modi = False Then
            s = "select codigo from clientes where activo = 1 and cuit = '" & uCuit.Text & "' and codigo <> " & s2n(txtCodigo)
            tmp = obtenerDeSQL(s)
            If s2n(tmp) > 0 Then
                che "El cuit ya existe en el cliente nro: " & tmp
                FaltanCosas = True
                Exit Function
            End If
        End If
    End If
    
    If chkTiene_Cuenta And txtcuenta = "" Then
        MsgBox "Este cliente utiliza cuenta contable propia pero no hay asignada ninguna cuenta.", vbCritical, "Informe"
        FaltanCosas = True
        Exit Function
    End If
    
    If cmbprovincias.Text = "" Then
        MsgBox "Debe ingresar la provincia.", , "ATENCION"
        FaltanCosas = True
        Exit Function
    End If
    
    If uCuit.digitoOK Then
    Else
        MsgBox "El cuit no es valido.", vbCritical
        FaltanCosas = True
        Exit Function
    End If
    
End Function

Private Sub GrabarCliente(Ope As String, Optional modi As Boolean = False)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr

    Dim Correo As Long
    Dim Puedo As Long
    Dim Certificado As Long
    Dim prov As Long
    Dim consig As Long
    Dim Mayor As Long
    Dim etiqueta As Long
    Dim limite As Double
    Dim ConPercIIBB As Long
    Dim tiene_Cuenta As Long
    Dim cliente_Cuenta As Long
            
    If FaltanCosas(modi) Then Exit Sub
            
    Correo = k2b(chkcorreo.Value)
    Puedo = k2b(chkhabilitado.Value)
    Certificado = k2b(chkcertificado.Value)
    consig = k2b(chkconsig.Value)
    Mayor = k2b(chkmay.Value)
    etiqueta = k2b(chketiqueta.Value)
    limite = s2n(txtlimite)
    ConPercIIBB = k2b(chkConPercIIBB.Value)
    cliente_Cuenta = s2n(txtcuenta)
    
    If chkTiene_Cuenta = 0 Then
        tiene_Cuenta = 0
    Else
        tiene_Cuenta = 1
    End If
    
    If Ope = "A" Then
        DataEnvironment1.dbo_CLIENTE "A", s2n(txtCodigo), Trim(txtNombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtdireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, val(Trim(TxtCodProv)), Trim(txttel) & "-" & Trim(Text3) & "-" & Trim(Text6), Trim(txtFax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), s2n(txtdescuento1), _
            s2n(txtdescuento2), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categclie", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, Mayor, _
            ConPercIIBB, etiqueta, _
            Date, UsuarioSistema!codigo, tiene_Cuenta, cliente_Cuenta
            
            DataEnvironment1.Sistema.Execute "update clientes set cuentasventas=" & ssTexto(CtasVentasGet) & ",RUC=" & ssTexto(txtRUC) & " where codigo=" & txtCodigo
            DataEnvironment1.Sistema.Execute "update clientes set pais=" & ssTexto(txtpais) & ",paisWSFEX=" & ssTexto(txtCodPais) & ",nrocliente=" & ssTexto(txtNroCliente) & " where codigo=" & s2n(txtCodigo)
            
            If chkPercIIBB_p Then
                cadena_iu = "update clientes set conperciibbper=1 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ClieTipoRetIB_Per (Codigo, CodClie, BaseImponible, Coeficiente) VALUES (1," & val(txtCodigo) & "," & x2s(Text1) & "," & x2s(Text4) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update clientes set conperciibbper=0 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetIB_Per where (CodClie = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
            
            If chkPercGAN_p Then
                cadena_iu = "update clientes set conpercganper=1 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ClieTipoRetGAN_Per (Codigo, CodClie, BaseImponible, Coeficiente) VALUES (1," & val(txtCodigo) & "," & x2s(Text2) & "," & x2s(Text5) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update clientes set conpercganper=0 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetGAN_Per where (CodClie = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
        
    ElseIf Ope = "M" Then
        DataEnvironment1.dbo_CLIENTE "M", s2n(txtCodigo), Trim(txtNombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtdireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, s2n(TxtCodProv), Trim(txttel) & "-" & Trim(Text3) & "-" & Trim(Text6), Trim(txtFax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), s2n(txtdescuento1), _
            s2n(txtdescuento2), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categclie", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, Mayor, _
            ConPercIIBB, etiqueta, _
            Date, UsuarioSistema!codigo, tiene_Cuenta, cliente_Cuenta
        grabaBitacora "M", s2n(txtCodigo), "Clientes"
            
            DataEnvironment1.Sistema.Execute "update clientes set cuentasventas=" & ssTexto(CtasVentasGet) & ",RUC=" & ssTexto(txtRUC) & " where codigo=" & txtCodigo
            DataEnvironment1.Sistema.Execute "update clientes set pais=" & ssTexto(txtpais) & ",paisWSFEX=" & ssTexto(txtCodPais) & ",nrocliente=" & ssTexto(txtNroCliente) & " where codigo=" & s2n(txtCodigo)
            
            If chkPercIIBB_p Then
                cadena_iu = "update clientes set ConPercIIBBper=1 where codigo = " & s2n(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetIB_Per where (CodClie = " & s2n(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ClieTipoRetIB_Per (Codigo, CodClie, BaseImponible, Coeficiente) VALUES (1," & s2n(txtCodigo) & "," & x2s(Text1) & "," & x2s(Text4) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update clientes set conperciibbper=0 where codigo = " & s2n(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetIB_Per where (CodClie = " & s2n(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
            
            
            If chkPercGAN_p Then
                cadena_iu = "update clientes set conpercganper=1 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetGAN_Per where (CodClie = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ClieTipoRetGAN_Per (Codigo, CodClie, BaseImponible, Coeficiente) VALUES (1," & val(txtCodigo) & "," & x2s(Text2) & "," & x2s(Text5) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update clientes set conpercganper=0 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ClieTipoRetGAN_Per where (CodClie = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
    End If
    MsgBox "La Operacion se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk "codigo = " & txtCodigo

fin:
    Exit Sub
ufaErr:
    ufa "Err al grabar", Me.Name & " " & txtCodigo ', Err
    Resume fin
End Sub

'*----------------------------- MENU ---------------------------------
Private Sub ucMenu_AceptarAlta()
    GrabarCliente "A"
End Sub
Private Sub ucMenu_AceptarModi()
    GrabarCliente "M", True
End Sub
Private Sub ucMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub ucMenu_Buscar()
    Dim resu As String
    'resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
    resu = frmBuscar.MostrarSql("select Codigo as [ Codigo              ], Descripcion [ Descripcion                                                             ],cuit [ CUIT              ],isnull(nombrefantasia,'') as [ Alias                                         ] from clientes where activo = 1 order by codigo ") ', arrayAnchos)
    If resu > "" Then
        txtCodigo = resu
        ucMenu.BuscarOK "codigo = " & txtCodigo
        CargoRegistro
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    DataEnvironment1.dbo_CLIENTE "B", s2n(txtCodigo), "", "", "", 0, 0, "", "", "", "", "", "", 0, "", "", "", 0, 0, 0, 0, "", "", "", "", 0, 0, 0, 0, "", "", "", "", 0, 0, 0, 0, "", "", 0, 0, 0, 0, Date, val(UsuarioSistema!codigo), 0, 0
    grabaBitacora "B", s2n(txtCodigo), "Clientes"
    MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
    ucMenu.EliminarOK
fin:
    Exit Sub
ufaErr:
    ufa "Err al eliminar", Me.Name & " " & s2n(txtCodigo) ', Err
    Resume fin
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt Not sino ' ta'reves
End Sub

Private Sub ucMenu_Modificar()
    txtCodigo.enabled = False
End Sub

Private Sub ucMenu_Nuevo()
    On Error Resume Next
    txtCodigo = nuevoCodigo("Clientes")
    txtCodigo.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
Private Sub ucMenu_SeMovio()
    CargoRegistro
End Sub
'*----------------------------- MENU ---------------------------------


' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
'   inhibo mov al cargar
' 17/8/4
'   codigo en CargarDatos() ? lo saque
' 18-10-4 from lorena-default date
' 30/11/4 los change x mayusculas, key press indiv x reempl enter x tab
'       se hacen ahora en 1 sola linea frmKeyPress
'       los gotfocus se reempl x instr con 1 solo parametro, el form


Private Sub uCuit_GotFocus()
    PintoFocoActivo
End Sub

Private Sub uCuit_LostFocus()

    Dim tmp, s As String

    s = "select codigo from clientes where activo = 1 and cuit = '" & uCuit.Text & "' "
    tmp = obtenerDeSQL(s)

    If Not IsEmpty(tmp) And tmp <> s2n(txtCodigo) Then
        MsgBox "El cuit ya existe,es del cliente nro: " & tmp
    End If
'
''    Dim rs As New ADODB.Recordset
''
''    rs.Open "Select * from clientes where cuit='" & MaskCuit & "'", daTaenvironment1.Sistema, adOpenStatic, adLockReadOnly
''    If Not rs.EOF Then
''        MsgBox "El cuit ya existe,es del cliente nro: " & rs!Codigo
''        MaskCuit.SetFocus
''    End If
'
End Sub
