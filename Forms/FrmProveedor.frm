VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmProveedor 
   Caption         =   "Proveedores"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   Icon            =   "FrmProveedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelCtasCompras 
      Height          =   315
      Left            =   5175
      Picture         =   "FrmProveedor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   7800
      Width           =   345
   End
   Begin VB.CommandButton cmdAddCtasCompras 
      Height          =   315
      Left            =   4800
      Picture         =   "FrmProveedor.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   7800
      Width           =   345
   End
   Begin VSFlex7LCtl.VSFlexGrid gCtasCompras 
      Height          =   1290
      Left            =   60
      TabIndex        =   105
      Top             =   8145
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
   Begin Gestion.ucCuit MaskCuit 
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Top             =   480
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   450
   End
   Begin VB.CheckBox chkRetIIBB_p 
      Alignment       =   1  'Right Justify
      Caption         =   "Con Ret IIBB Personal"
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
      Left            =   150
      TabIndex        =   30
      Top             =   4245
      Width           =   2535
   End
   Begin VB.CheckBox chkretGAN_p 
      Alignment       =   1  'Right Justify
      Caption         =   "Con Ret GAN Personal"
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
      Left            =   2820
      TabIndex        =   31
      Top             =   4245
      Width           =   2535
   End
   Begin VB.Frame frameIIBB 
      Caption         =   "Ingresos Brutos"
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
      Height          =   1290
      Left            =   30
      TabIndex        =   92
      Top             =   5025
      Width           =   10530
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5085
         TabIndex        =   98
         Text            =   "0"
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5085
         TabIndex        =   97
         Text            =   "0"
         Top             =   645
         Width           =   1635
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   5085
         TabIndex        =   96
         Text            =   "0"
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3330
         TabIndex        =   95
         Text            =   "0"
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3330
         TabIndex        =   94
         Text            =   "0"
         Top             =   645
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3330
         TabIndex        =   93
         Text            =   "0"
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label42 
         Caption         =   "Coeficiente"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5100
         TabIndex        =   104
         Top             =   135
         Width           =   1605
      End
      Begin VB.Label Label41 
         Caption         =   "Bases imponibles"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3360
         TabIndex        =   103
         Top             =   135
         Width           =   1605
      End
      Begin VB.Label Label40 
         Caption         =   "<- Ingresar valores como por                     ej: Compra bienes muebles, Coef = 1,75"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   6900
         TabIndex        =   102
         Top             =   390
         Width           =   3045
      End
      Begin VB.Label Label39 
         Caption         =   "Pago de servicios eventuales"
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
         Left            =   135
         TabIndex        =   101
         Top             =   930
         Width           =   3075
      End
      Begin VB.Label Label38 
         Caption         =   "Locaciones de obras y/o servicios"
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
         Left            =   135
         TabIndex        =   100
         Top             =   660
         Width           =   3240
      End
      Begin VB.Label Label26 
         Caption         =   "Compra bienes muebles"
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
         Left            =   135
         TabIndex        =   99
         Top             =   375
         Width           =   2475
      End
   End
   Begin VB.Frame frameGAN 
      Caption         =   "Ganancias"
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
      Height          =   1515
      Left            =   30
      TabIndex        =   79
      Top             =   6285
      Width           =   10530
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   4170
         TabIndex        =   85
         Text            =   "0"
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   4170
         TabIndex        =   84
         Text            =   "0"
         Top             =   645
         Width           =   1635
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   4170
         TabIndex        =   83
         Text            =   "0"
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   5925
         TabIndex        =   82
         Text            =   "0"
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   5925
         TabIndex        =   81
         Text            =   "0"
         Top             =   645
         Width           =   1635
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   5925
         TabIndex        =   80
         Text            =   "0"
         Top             =   930
         Width           =   1635
      End
      Begin VB.Label Label27 
         Caption         =   "Profesionales Liberales, Oficios"
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
         Left            =   135
         TabIndex        =   91
         Top             =   375
         Width           =   3825
      End
      Begin VB.Label Label28 
         Caption         =   "Enajenación bienes muebles y bienes de cambio"
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
         Left            =   120
         TabIndex        =   90
         Top             =   645
         Width           =   4050
      End
      Begin VB.Label Label29 
         Caption         =   "Locaciones de obra y/o servicios no ejecutados en relación de dependencia"
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
         Height          =   555
         Left            =   135
         TabIndex        =   89
         Top             =   915
         Width           =   3930
      End
      Begin VB.Label Label30 
         Caption         =   "<- Ingresar valores como por                     ej: Profesionales Liberales, Coef = 1,75"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   7575
         TabIndex        =   88
         Top             =   390
         Width           =   2910
      End
      Begin VB.Label Label31 
         Caption         =   "Bases imponibles"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4200
         TabIndex        =   87
         Top             =   135
         Width           =   1605
      End
      Begin VB.Label Label34 
         Caption         =   "Coeficiente"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5940
         TabIndex        =   86
         Top             =   135
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdCuenta 
      Caption         =   "Cuenta"
      Enabled         =   0   'False
      Height          =   300
      Left            =   195
      TabIndex        =   78
      Top             =   4590
      Width           =   975
   End
   Begin VB.TextBox txtCuenta_Descrip 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2505
      TabIndex        =   33
      Top             =   4590
      Width           =   5220
   End
   Begin VB.CheckBox chkTiene_Cuenta 
      Caption         =   "Usar Cuenta"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7770
      TabIndex        =   34
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtCuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1185
      TabIndex        =   32
      Text            =   "0"
      Top             =   4590
      Width           =   1305
   End
   Begin VB.ComboBox cboRetenerIIBB 
      Height          =   315
      ItemData        =   "FrmProveedor.frx":13DE
      Left            =   8595
      List            =   "FrmProveedor.frx":13E0
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Tag             =   "10"
      Top             =   3795
      Width           =   810
   End
   Begin VB.TextBox txtNroIIBB 
      Height          =   285
      Left            =   4620
      TabIndex        =   25
      Tag             =   "9"
      Top             =   3195
      Width           =   2055
   End
   Begin VB.ComboBox cboTipoIB 
      Height          =   315
      Left            =   2115
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   3885
      Width           =   2490
   End
   Begin VB.ComboBox cboTipoRetGan 
      Height          =   315
      Left            =   2115
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   3510
      Width           =   2490
   End
   Begin VB.TextBox txtdescuento2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3540
      MaxLength       =   5
      TabIndex        =   17
      Tag             =   "14"
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdcli 
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   2190
      Width           =   375
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmProveedor.frx":13E2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   5565
      Picture         =   "FrmProveedor.frx":1CAC
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   9495
      Width           =   975
   End
   Begin VB.CheckBox chkhabilitado 
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
      Height          =   270
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3195
      Width           =   1695
   End
   Begin VB.TextBox txtcodcli 
      Height          =   285
      Left            =   4680
      TabIndex        =   20
      Tag             =   "22"
      Top             =   2175
      Width           =   1095
   End
   Begin VB.ComboBox cmbFormaPagos 
      Height          =   315
      ItemData        =   "FrmProveedor.frx":2576
      Left            =   7275
      List            =   "FrmProveedor.frx":2583
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "15"
      Top             =   1515
      Width           =   2565
   End
   Begin VB.ComboBox cmbordcompra 
      Height          =   315
      ItemData        =   "FrmProveedor.frx":25B8
      Left            =   2715
      List            =   "FrmProveedor.frx":25C2
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Tag             =   "10"
      Top             =   3165
      Width           =   810
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   555
      Left            =   60
      Picture         =   "FrmProveedor.frx":25CE
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Primero"
      Top             =   9495
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   555
      Left            =   1965
      Picture         =   "FrmProveedor.frx":28D8
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Ultimo"
      Top             =   9495
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   555
      Left            =   1350
      Picture         =   "FrmProveedor.frx":2BE2
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Siguiente"
      Top             =   9495
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   555
      Left            =   675
      Picture         =   "FrmProveedor.frx":2EEC
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Anterior"
      Top             =   9495
      Width           =   675
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      DisabledPicture =   "FrmProveedor.frx":31F6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   9585
      Picture         =   "FrmProveedor.frx":3AC0
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton cmdmodificar 
      Caption         =   "&Modificar"
      DisabledPicture =   "FrmProveedor.frx":438A
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   4590
      Picture         =   "FrmProveedor.frx":4C54
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9495
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar"
      DisabledPicture =   "FrmProveedor.frx":551E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   2625
      Picture         =   "FrmProveedor.frx":5DE8
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   9495
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmProveedor.frx":66B2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   7635
      Picture         =   "FrmProveedor.frx":6F7C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      DisabledPicture =   "FrmProveedor.frx":7846
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8610
      Picture         =   "FrmProveedor.frx":8110
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmProveedor.frx":89DA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   3600
      Picture         =   "FrmProveedor.frx":92A4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9480
      Width           =   975
   End
   Begin VB.TextBox txtrf 
      Height          =   285
      Left            =   1050
      TabIndex        =   21
      Tag             =   "20"
      Top             =   2520
      Width           =   9015
   End
   Begin VB.ComboBox cmbcategoria 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "19"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox cmbprovext 
      Height          =   315
      ItemData        =   "FrmProveedor.frx":9B6E
      Left            =   8475
      List            =   "FrmProveedor.frx":9B78
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "17"
      Top             =   1185
      Width           =   975
   End
   Begin VB.TextBox txtdescuento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1275
      MaxLength       =   5
      TabIndex        =   16
      Tag             =   "14"
      Top             =   1845
      Width           =   1095
   End
   Begin VB.ComboBox cmbagente 
      Height          =   315
      ItemData        =   "FrmProveedor.frx":9B84
      Left            =   6300
      List            =   "FrmProveedor.frx":9B8E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "10"
      Top             =   465
      Width           =   735
   End
   Begin VB.TextBox txtsucursal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8250
      TabIndex        =   6
      Tag             =   "12"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtpais 
      Height          =   285
      Left            =   4500
      TabIndex        =   11
      Tag             =   "6"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox cmbprovincias 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "5"
      Top             =   1185
      Width           =   2775
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Left            =   930
      TabIndex        =   1
      Tag             =   "0"
      Top             =   105
      Width           =   1215
   End
   Begin VB.TextBox txtfax 
      Height          =   285
      Left            =   3675
      TabIndex        =   14
      Tag             =   "9"
      Top             =   1530
      Width           =   2055
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   1215
      TabIndex        =   13
      Tag             =   "8"
      Top             =   1530
      Width           =   2055
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   6990
      TabIndex        =   9
      Tag             =   "3"
      Top             =   870
      Width           =   2820
   End
   Begin VB.ComboBox cmbivas 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "11"
      Top             =   465
      Width           =   1440
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   3510
      TabIndex        =   2
      Tag             =   "1"
      Top             =   120
      Width           =   6330
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Tag             =   "2"
      Top             =   855
      Width           =   3255
   End
   Begin VB.ComboBox cmbtipocompra 
      Height          =   315
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "18"
      Top             =   1875
      Width           =   2175
   End
   Begin VB.TextBox txtmail 
      Height          =   285
      Left            =   1035
      TabIndex        =   22
      Tag             =   "21"
      Top             =   2850
      Width           =   4425
   End
   Begin VB.TextBox txtweb 
      Height          =   285
      Left            =   6015
      TabIndex        =   23
      Tag             =   "22"
      Top             =   2865
      Width           =   4050
   End
   Begin MSMask.MaskEdBox txtcodpostal 
      Height          =   300
      Left            =   5025
      TabIndex        =   8
      Top             =   855
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "?9999???"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskCuit2 
      Height          =   300
      Left            =   8550
      TabIndex        =   29
      Top             =   4140
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   13
      Format          =   "c"
      Mask            =   "99-99999999-9"
      PromptChar      =   "_"
   End
   Begin VB.Label Label35 
      Caption         =   "CUENTAS PARA FACTURA COMPRA"
      Height          =   345
      Left            =   90
      TabIndex        =   108
      Top             =   7875
      Width           =   3195
   End
   Begin VB.Label Label25 
      Caption         =   "Retener IIBB:"
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
      Left            =   8610
      TabIndex        =   77
      Top             =   3525
      Width           =   1425
   End
   Begin VB.Label Label19 
      Caption         =   "Nro de IIBB: "
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
      TabIndex        =   76
      Top             =   3195
      Width           =   1170
   End
   Begin VB.Label lblIB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4770
      TabIndex        =   75
      Top             =   3885
      Width           =   3555
   End
   Begin VB.Label lblRG 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4770
      TabIndex        =   74
      Top             =   3525
      Width           =   3525
   End
   Begin VB.Label Label32 
      Caption         =   "Tipo Para IIBB:"
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
      Index           =   2
      Left            =   195
      TabIndex        =   73
      Top             =   3885
      Width           =   1485
   End
   Begin VB.Label Label32 
      Caption         =   "Tipo Para Ret Gan:"
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
      Left            =   150
      TabIndex        =   72
      Top             =   3510
      Width           =   1785
   End
   Begin VB.Label Label24 
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
      Left            =   2415
      TabIndex        =   71
      Top             =   1875
      Width           =   1095
   End
   Begin VB.Label Label5 
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
      Left            =   4650
      TabIndex        =   68
      Top             =   885
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Cod.Cliente:"
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
      TabIndex        =   67
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "Requiere Orden de Compra:"
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
      TabIndex        =   66
      Top             =   3180
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   4995
      Left            =   45
      Top             =   30
      Width           =   10500
   End
   Begin VB.Label Label23 
      Caption         =   "Contacto:"
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
      TabIndex        =   57
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label21 
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
      Left            =   195
      TabIndex        =   56
      Top             =   2175
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "Proveedor Exterior:"
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
      Left            =   6675
      TabIndex        =   55
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label18 
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
      Left            =   195
      TabIndex        =   54
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "Nº Sucursal:"
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
      Left            =   7050
      TabIndex        =   53
      Top             =   480
      Width           =   1095
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
      Left            =   5790
      TabIndex        =   52
      Top             =   1530
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Left            =   195
      TabIndex        =   51
      Top             =   495
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "País:"
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
      Left            =   4020
      TabIndex        =   50
      Top             =   1215
      Width           =   615
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
      Left            =   210
      TabIndex        =   49
      Top             =   1200
      Width           =   975
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
      Left            =   210
      TabIndex        =   48
      Top             =   1530
      Width           =   1215
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
      Left            =   13560
      TabIndex        =   47
      Top             =   1920
      Width           =   615
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
      Left            =   2235
      TabIndex        =   46
      Top             =   135
      Width           =   1455
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
      Left            =   195
      TabIndex        =   45
      Top             =   855
      Width           =   975
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
      Height          =   375
      Left            =   195
      TabIndex        =   44
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
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
      Left            =   6000
      TabIndex        =   43
      Top             =   870
      Width           =   1215
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
      Left            =   2310
      TabIndex        =   42
      Top             =   480
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
      Left            =   3255
      TabIndex        =   41
      Top             =   1530
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Agente Retencion:"
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
      Left            =   4605
      TabIndex        =   40
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "Tipo de compra:"
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
      Left            =   4710
      TabIndex        =   39
      Top             =   1890
      Width           =   1695
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
      Index           =   0
      Left            =   210
      TabIndex        =   38
      Top             =   2865
      Width           =   735
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
      Left            =   5505
      TabIndex        =   37
      Top             =   2880
      Width           =   615
   End
End
Attribute VB_Name = "FrmProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 12/8/4
'6/1/5  le encaje un ignorar errores en limpiardatos()

Dim rsProv As New ADODB.Recordset
Dim Ope As String
Dim numero As Long


Private Sub cboTipoRetGan_Validate(cancel As Boolean)
    verRG
End Sub

Private Sub verRG()
    Dim t
    t = obtenerDeSQL("select Baseimponible, coeficiente from ProvTipoRetGan_per where codigo = '" & ComboCodigo(cboTipoRetGan) & "' ")
    lblRG = ""
    If IsEmpty(t) Then Exit Sub
    If s2n(t(0)) > 0 Then lblRG = "Base = " & t(0)
    If s2n(t(1)) > 0 Then lblRG = lblRG & " Coef = " & t(1)
End Sub

Private Sub chkretGAN_p_Click()
    If chkretGAN_p Then
        frameGAN.enabled = True
    Else
        frameGAN.enabled = False
    End If
End Sub

Private Sub chkRetIIBB_p_Click()
    If chkRetIIBB_p Then
        frameIIBB.enabled = True
    Else
        frameIIBB.enabled = False
    End If
End Sub



Private Sub cmdAddCtasCompras_Click()
Dim Res
Dim rr As Long
    Res = frmBuscar.MostrarSql("Select [Cuenta              ],[Descripcion                                    ] from Cuentas where imputable=1")
    If Res > "" Then
        gCtasCompras.AddItem ""
        rr = gCtasCompras.rows - 1
        gCtasCompras.TextMatrix(rr, 0) = Res
        gCtasCompras.TextMatrix(rr, 1) = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdDelCtasCompras_Click()
    If gCtasCompras.rows > 0 Then
        If gCtasCompras.Row >= 0 Then
            gCtasCompras.RemoveItem gCtasCompras.Row
        End If
    End If
End Sub

Private Function CtasComprasGet() As String
Dim i As Long, C As String
Se_Exedio:
    If gCtasCompras.rows > 0 Then
        C = ""
        For i = 0 To gCtasCompras.rows - 1
            If i = 0 Then
                C = "#" & gCtasCompras.TextMatrix(i, 0) & "#"
            Else
                C = C & ",#" & gCtasCompras.TextMatrix(i, 0) & "#"
            End If
        Next
    Else
        C = ""
    End If
    If Len(C) > 4000 Then
        MsgBox "Se exede de la cantidad de cuentas permitidas. se quitara la ultima una para continuar.", vbExclamation
        gCtasCompras.RemoveItem gCtasCompras.rows - 1
        GoTo Se_Exedio
    End If

CtasComprasGet = C
End Function

Private Function CtasComprasSet(C As String)
Dim i As Long, rr As Long, CT
C = Trim(C)
    If C > "" Then
        CT = Split(Replace(C, "#", ""), ",")
        iniCtasCompras
        For i = 0 To UBound(CT)
            gCtasCompras.AddItem ""
            rr = gCtasCompras.rows - 1
            gCtasCompras.TextMatrix(rr, 0) = CT(i)
            gCtasCompras.TextMatrix(rr, 1) = obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(CT(i)))
        Next
    Else
        iniCtasCompras
    End If
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub cmbprovincias_LostFocus()
    txtpais = "Argentina"
End Sub

Private Sub cmdAceptar_Click()
Dim tiene_cuenta_PROV As Long, cadena_iu As String
    If Trim(txtCodigo) = "" Then
        MsgBox "Debe cargar el codigo de proveedor", 48, "Atencion"
        txtCodigo.SetFocus
        Exit Sub
    Else
        If Trim(txtNombre) = "" Then
            MsgBox "Debe cargar el nombre de proveedor", 48, "Atencion"
            txtNombre.SetFocus
            Exit Sub
        Else
            If MaskCuit.Text = "__-________-_" Then
                MsgBox "Debe cargar el cuit de proveedor", 48, "Atencion"
                MaskCuit.SetFocus
                Exit Sub
            End If
        End If
    End If
    If MaskCuit.digitoOK Then
    Else
        MsgBox "Debe cargar un cuit valido", 48, "Atencion"
        Exit Sub
    End If
    If cmbprovincias.Text = "" Then
        MsgBox "Debe ingresar la provincia.", , "ATENCION"
        Exit Sub
    End If
    
    
    If chkTiene_Cuenta.Value = 0 Then
        tiene_cuenta_PROV = 0
        txtcuenta = ""
    Else
        tiene_cuenta_PROV = 1
        If Trim(txtcuenta) = "" Then
            tiene_cuenta_PROV = 0
        End If
    End If
    
    If Ope <> "" Then
    Dim g As String
        
        If Ope = "A" Then Ope = "M"
        
        If Ope = "A" Then
             DataEnvironment1.dbo_PROVEEDOR "A", val(Trim(txtCodigo)), Trim(txtNombre), _
                Trim(txtdireccion), Trim(txtLocalidad), UCase(Trim(txtcodpostal)), _
                ObtenerCodigoS("Provincias", Trim(cmbprovincias.Text)), Trim(txtpais), Trim(MaskCuit), _
                Trim(txttel), Trim(txtFax), IIf(Trim(cmbagente.Text) = "SI", 1, 0), ObtenerCodigo("Ivas", Trim(cmbivas.Text)), _
                Trim(txtsucursal), Trim(cmbFormaPagos.Text), s2n(Trim(txtDescuento), 4), s2n(Trim(txtdescuento2), 4), _
                IIf(Trim(cmbprovext.Text) = "SI", "S", "N"), _
                ObtenerCodigo("Tipocompras", cmbtipocompra.Text), ObtenerCodigo("ProvCategoria", Trim(cmbcategoria.Text)), _
                Trim(txtrf), Trim(txtmail), Trim(txtweb), IIf(Trim(cmbordcompra.Text) = "SI", 1, 0), IIf(chkhabilitado.Value = 1, 1, 0), val(Trim(txtcodcli)), ComboCodigo(cboTipoRetGan), ComboCodigo(cboTipoIB), ssStr(txtNroIIBB), (cboRetenerIIBB.ListIndex), _
                Date, UsuarioActual(), tiene_cuenta_PROV, txtcuenta
                DataEnvironment1.Sistema.Execute "update prov set cuentascompras=" & ssTexto(CtasComprasGet) & " where codigo=" & txtCodigo
            ''guardar iibb personales
            ''BuscarID(Trim(cmbtipocompra.Text))
            If chkRetIIBB_p Then
                cadena_iu = "update prov set conretiibbper=1 where codigo = " & s2n(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (0," & val(txtCodigo) & ",0,0)"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (1," & val(txtCodigo) & "," & x2s(Text1) & "," & x2s(Text4) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (2," & val(txtCodigo) & "," & x2s(Text2) & "," & x2s(Text5) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (3," & val(txtCodigo) & "," & x2s(Text3) & "," & x2s(Text6) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update prov set conretiibbper=0 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ProvTipoRetIB_Per where (CodProv = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
            ''guardar gan personales
            If chkretGAN_p Then
                cadena_iu = "update prov set conretganper=1 where codigo = " & s2n(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (0,0,0," & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (1," & x2s(Text7) & "," & x2s(Text10) & "," & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (2," & x2s(Text8) & "," & x2s(Text11) & "," & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (3," & x2s(Text9) & "," & x2s(Text12) & "," & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            Else
                cadena_iu = "update prov set conretganper=0 where codigo = " & val(txtCodigo)
                DataEnvironment1.Sistema.Execute cadena_iu
                cadena_iu = "DELETE FROM ProvTipoRetGan_Per where (Prov = " & val(txtCodigo) & ")"
                DataEnvironment1.Sistema.Execute cadena_iu
            End If
        Else
            If Ope = "M" Then
                 DataEnvironment1.dbo_PROVEEDOR "M", val(Trim(txtCodigo)), Trim(txtNombre), _
                    Trim(txtdireccion), Trim(txtLocalidad), UCase(Trim(txtcodpostal)), _
                    ObtenerCodigoS("Provincias", Trim(cmbprovincias.Text)), Trim(txtpais), Trim(MaskCuit), _
                    Trim(txttel), Trim(txtFax), IIf(Trim(cmbagente.Text) = "SI", 1, 0), ObtenerCodigo("Ivas", Trim(cmbivas.Text)), _
                    Trim(txtsucursal), Trim(cmbFormaPagos.Text), s2n(txtDescuento), s2n(txtdescuento2), _
                    IIf(Trim(cmbprovext.Text) = "SI", "S", "N"), _
                    ObtenerCodigo("Tipocompras", cmbtipocompra.Text), ObtenerCodigo("ProvCategoria", Trim(cmbcategoria.Text)), _
                    Trim(txtrf), Trim(txtmail), Trim(txtweb), IIf(Trim(cmbordcompra.Text) = "SI", 1, 0), IIf(chkhabilitado.Value = 1, 1, 0), val(Trim(txtcodcli)), ComboCodigo(cboTipoRetGan), ComboCodigo(cboTipoIB), Trim(txtNroIIBB), cboRetenerIIBB.ListIndex, _
                    Date, UsuarioActual(), tiene_cuenta_PROV, txtcuenta
                    DataEnvironment1.Sistema.Execute "update prov set cuentascompras=" & ssTexto(CtasComprasGet) & " where codigo=" & txtCodigo
            ''mod de iibb personal
            'BuscarID(Trim(cmbtipocompra.Text))
                If chkRetIIBB_p Then
                    cadena_iu = "update prov set conretiibbper=1 where codigo = " & val(txtCodigo)
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "DELETE FROM ProvTipoRetIB_Per where (CodProv = " & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (0," & val(txtCodigo) & ",0,0)"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (1," & val(txtCodigo) & "," & x2s(Text1) & "," & x2s(Text4) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (2," & val(txtCodigo) & "," & x2s(Text2) & "," & x2s(Text5) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetIB_Per (Codigo, CodProv, BaseImponible, Coeficiente) VALUES (3," & val(txtCodigo) & "," & x2s(Text3) & "," & x2s(Text6) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                Else
                    cadena_iu = "update prov set conretiibbper=0 where codigo = " & val(txtCodigo)
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "DELETE FROM ProvTipoRetIB_Per where (CodProv = " & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                End If
                ''mod gan personal
                If chkretGAN_p Then
                    cadena_iu = "update prov set conretganper=1 where codigo = " & val(txtCodigo)
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "DELETE FROM ProvTipoRetGan_Per where (Prov = " & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (0,0,0," & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (1," & x2s(Text7) & "," & x2s(Text10) & "," & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (2," & x2s(Text8) & "," & x2s(Text11) & "," & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "INSERT INTO ProvTipoRetGan_Per (Codigo, BaseImponible, Coeficiente, Prov) VALUES (3," & x2s(Text9) & "," & x2s(Text12) & "," & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                Else
                    cadena_iu = "update prov set conretganper=0 where codigo = " & val(txtCodigo)
                    DataEnvironment1.Sistema.Execute cadena_iu
                    cadena_iu = "DELETE FROM ProvTipoRetGan_Per where (Prov = " & val(txtCodigo) & ")"
                    DataEnvironment1.Sistema.Execute cadena_iu
                End If
                DataEnvironment1.dbo_GRABARBITACORA val(Trim(txtCodigo)), "Proveedores", UsuarioSistema!codigo, Date, Time, "M"
            End If
        End If
        
        MsgBox "La operación fue realizada con éxito"
        LimpioControles
        HabilitoControles (True)
        Call Habilitobotones(True, True, True, True, True, True)
    Else
        MsgBox "Operación no válida"
    End If

End Sub
Private Function BuscarID(Tipo As String) As Long
Dim sql As String
Dim rsCParam As New ADODB.Recordset
sql = "SELECT id FROM CuentasParam WHERE descripcion ='" & Tipo & "'"
rsCParam.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
If Not rsCParam.EOF Then
   BuscarID = rsCParam!ID
End If
End Function
Private Sub cmdBuscar_Click()
    Dim resu As String
    'resu = frmBuscar.MostrarCodigoDescripcionActivo("Prov")
    resu = frmBuscar.MostrarSql("select codigo as [ Codigo         ], cuit as [ CUIT            ], descripcion [ Descripcion                                                             ] from prov where activo = 1 order by codigo ")  ', arrayAnchos)
    If resu > "" Then
        txtCodigo = resu
        CargarDatos
        Call Habilitobotones(True, False, True, True, True, True)
    End If
End Sub

Private Sub cmdCancelar_Click()
Dim s As String
If Ope = "A" Then
    s = "delete from prov where codigo=" & s2n(txtCodigo)
    DataEnvironment1.Sistema.Execute s
End If
    LimpioControles
    Call HabilitoControles(True)
    Call Habilitobotones(True, True, False, False, False, True)
    Call HabilitoBotonesMoverse(False, False, False, False)
End Sub
Public Sub CargarDatos()
    Dim codigo

    If rsProv.State = 1 Then
        rsProv.Close
        Set rsProv = Nothing
    End If

    codigo = val(Trim(txtCodigo))
    rsProv.Open "select * from Prov where activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rsProv.EOF Then
        rsProv.MoveFirst
        rsProv.Find "Codigo= " & str(Trim(txtCodigo))
        CargoProveedor
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
    

End Sub


Private Sub cmdcli_Click()
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
    If resu > "" Then
        txtcodcli = resu
        txtcodcli.SetFocus
    End If

End Sub

Private Sub cmdCuenta_Click()
txtcuenta = BuscarCuenta(False, False)
If txtcuenta > "" Then txtCuenta_Descrip = obtenerDeSQL("select descripcion from cuentas where cuenta= " & txtcuenta)

End Sub

Private Sub cmdeliminar_Click()
Dim mensaje As String

    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        DataEnvironment1.dbo_PROVEEDOR "B", Trim(txtCodigo), "", "", "", "", "", "", "", "", "", 0, 0, "", "", 0, 0, 0, 0, 0, "", "", "", 0, 0, 0, 0, 0, "", 0, Date, UsuarioActual(), 0, 0
        DataEnvironment1.dbo_GRABARBITACORA val(Trim(txtCodigo)), "Prov", UsuarioSistema!codigo, Date, Time, "B"
        Call Habilitobotones(True, True, False, False, False, False)
        LimpioControles
        Call HabilitoControles(True)
        HabilitoBotonesMoverse False, False, False, False
    End If

End Sub

Private Sub cmdmodificar_Click()
    Ope = "M"
    Call HabilitoControles(False)
    Call Habilitobotones(True, False, False, True, True, True)
End Sub

Private Sub cmdnuevo_Click()
Dim rs As New ADODB.Recordset, s As String
Dim neww
    LimpioControles
    
    neww = obtenerDeSQL("select max(codigo) from prov")
    
    If IsNull(neww) Or IsEmpty(neww) Then
        txtCodigo = 1
    Else
        txtCodigo = neww + 1
    End If
    s = "insert into prov (codigo) values (" & txtCodigo & ")"
    DataEnvironment1.Sistema.Execute s
    'rs.Open "select max(codigo) + 1 as maxcodigo from Prov", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'If rs.EOF = True And rs.BOF = True Then
    '    txtCodigo = 1
    'Else
    '    If IsNull(rs!maxcodigo) Then
    '        txtCodigo = 1
    '    Else
    '        txtCodigo = rs!maxcodigo
    '    End If
    'End If
    
    
    txtCodigo.enabled = True
    txtCodigo.SetFocus
    'rs.Close
    'Set rs = Nothing
    
    Call HabilitoControles(False)
    Call Habilitobotones(False, False, False, False, True, True)
    chkhabilitado.Value = 1
    Ope = "A"
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    FrmKeyPress KeyAscii, True, True
'    If KeyAscii = 27 Then
'        Unload Me
'    End If
'End Sub

Private Sub Form_Load()
    LimpioControles
    CargaCombo cmbprovincias, "Provincias", "descripcion", "codigo", ""
'    CargaCombo cmbtipocompra, "CuentasParam", "descripcion", "id", "usocuenta= '" & ID_UsoCuenta_COMPRAS & "' "
    CargaCombo2 cmbtipocompra, "TipoCompras", "descripcion", "codigo", ""
    CargaCombo cmbcategoria, "ProvCategoria", "descripcion", "codigo", ""
    CargaCombo cmbivas, "ivas", "descripcion", "codigo", ""
    CargaCombo cmbFormaPagos, "formaspago", "descripcion", "codigo", ""
    
    comboSql cboTipoIB, "select descripcion, codigo from ProvTipoRetIB where activo = 1"
    comboSql cboTipoRetGan, "select descripcion, codigo from ProvTipoRetGan where activo = 1"
    
    comboArray cboRetenerIIBB, Array("No", "Si"), Array(0, 1)
    
    HabilitoControles (True)
    HabilitoBotonesMoverse False, False, False, False
    cmbcategoria.ListIndex = 0
End Sub

Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    cmdprimero.enabled = hab2
    cmdanterior.enabled = hab1
    cmdsiguiente.enabled = hab3
    cmdultimo.enabled = hab4
End Sub

Private Function iniCtasCompras()
With gCtasCompras
    .rows = 0
    .cols = 0
    .cols = 2
    .ColWidth(0) = 1200
    .ColWidth(1) = 3500
End With
End Function

Sub LimpioControles()
    On Error Resume Next
    
    FrmBorrarTxt Me
    
    'txtCodigo = ""
    'txtnombre = ""
    'txtDireccion = ""
    'txtLocalidad = ""
    cmbprovincias.ListIndex = 1
    'txtpais = ""
    'txttel = ""
    'txtfax = ""
    cmbagente.ListIndex = 1
    cmbivas.ListIndex = 0
    txtsucursal = "1"
    cmbFormaPagos.ListIndex = 0
    'txtDescuento = "0"
    'txtdescuento2 = "0"
'    cmbtipoprov.ListIndex = 0
    cmbprovext.ListIndex = 1
    cmbtipocompra.ListIndex = 0
    cmbcategoria.ListIndex = 2
    'txtrf = ""
    MaskCuit = "  -       - "
    txtcodpostal.Mask = "       "
    MaskCuit = "99-99999999-9"
    txtcodpostal.Mask = "?9999???"
    'txtmail = ""
    'txtweb = ""
    cmbordcompra.ListIndex = 1
    Ope = ""
    cboTipoIB.ListIndex = 0
    cboTipoRetGan.ListIndex = 0: verRG
    
    txtcuenta = 0
    txtCuenta_Descrip = ""
    chkTiene_Cuenta.Value = 0
    
    '16-11/07 'ret iibb y gan personal para cada prov
    chkRetIIBB_p = 0
    frameIIBB.enabled = False
    chkretGAN_p = 0
    frameGAN.enabled = False
    
    Text1.Text = 0
    Text2.Text = 0
    Text3.Text = 0
    Text4.Text = 0
    Text5.Text = 0
    Text6.Text = 0
    Text7.Text = 0
    Text8.Text = 0
    Text9.Text = 0
    Text10.Text = 0
    Text11.Text = 0
    Text12.Text = 0
    
    iniCtasCompras
End Sub

Sub CargoProveedor()
    On Error Resume Next
    Dim rs_ret As New ADODB.Recordset, conx As Integer, Aux As String
    txtCodigo = rsProv!codigo
    txtNombre = rsProv!DESCRIPCION
    
    txtNroIIBB = sSinNull(rsProv!NumIIBB)
    
    If Not IsNull(rsProv!direccion) Then
        txtdireccion = rsProv!direccion
    Else
        txtdireccion = ""
    End If
    If Not IsNull(rsProv!Localidad) Then
        txtLocalidad = rsProv!Localidad
    Else
        txtLocalidad = ""
    End If
    If Not IsNull(rsProv!codigopostal) Then
        txtcodpostal = rsProv!codigopostal
    Else
        txtcodpostal.Mask = "        "
        txtcodpostal.Mask = "?9999???"
    End If
    If Not IsNull(rsProv!Provincia) Then
        cmbprovincias.ListIndex = BuscarenComboS(cmbprovincias, ObtenerDescripcionS("provincias", rsProv!Provincia))
    Else
        cmbprovincias.ListIndex = -1
    End If
    If Not IsNull(rsProv!Pais) Then
        txtpais = rsProv!Pais
    Else
        txtpais = ""
    End If
    If Not IsNull(rsProv!CUIT) Then
        If rsProv!CUIT <> "" Then
            MaskCuit = rsProv!CUIT
        Else
            'MaskCuit.Mask = "  -       - "
            'MaskCuit.Mask = "99-99999999-9"
            MaskCuit = "  -       - "
            MaskCuit = "99-99999999-9"
            
        End If
        
    Else
        MaskCuit = "  -       - "
        MaskCuit = "99-99999999-9"
    End If
    If Not IsNull(rsProv!Telefono) Then
        txttel = rsProv!Telefono
    Else
        txttel = ""
    End If
    If Not IsNull(rsProv!fax) Then
        txtFax = rsProv!fax
    Else
        txtFax = ""
    End If
    If rsProv!agente = True Then
        cmbagente.ListIndex = 0
    Else
        cmbagente.ListIndex = 1
    End If
    If Not IsNull(rsProv!tipoiva) Then
        cmbivas.ListIndex = BuscarenComboS(cmbivas, ObtenerDescripcion("ivas", rsProv!tipoiva))
    Else
        cmbivas.ListIndex = -1
    End If
    If Not IsNull(rsProv!suc) Then
        txtsucursal = rsProv!suc
    Else
        txtsucursal = ""
    End If
    If Not IsNull(rsProv!pago) Then
        cmbFormaPagos.ListIndex = BuscarenComboS(cmbFormaPagos, rsProv!pago)
    Else
        cmbFormaPagos.ListIndex = -1
    End If
    If Not IsNull(rsProv!des) Then
        txtDescuento = rsProv!des
    Else
        txtDescuento = "0.00"
    End If
    If Not IsNull(rsProv!des2) Then
        txtdescuento2 = rsProv!des2
    Else
        txtdescuento2 = "0.00"
    End If
    If rsProv!exter = True Then
        cmbprovext.ListIndex = 0
    Else
        cmbprovext.ListIndex = 1
    End If
    If Not IsNull(rsProv!tipocom) Then
        'cmbtipocompra.ListIndex = BuscarenComboS(cmbtipocompra, ObtenerDescripcion("cuentasparam", rsProv!tipocom))
        cmbtipocompra.ListIndex = BuscarenComboS(cmbtipocompra, ObtenerDescripcion("TipoCompras", rsProv!tipocom))
    Else
        cmbtipocompra.ListIndex = -1
    End If
    
    If Not IsNull(rsProv!Categ) Then
        cmbcategoria.ListIndex = BuscarenComboS(cmbcategoria, ObtenerDescripcion("ProvCategoria", rsProv!Categ))
    Else
        cmbcategoria.ListIndex = -1
    End If
    
    cboTipoIB.ListIndex = BuscarEnCombo(cboTipoIB, rsProv!TipoRetIIBB)
    cboTipoRetGan.ListIndex = BuscarEnCombo(cboTipoRetGan, rsProv!TipoRetGan)
    verRG
    
    If Not IsNull(rsProv!rf) Then
        txtrf = rsProv!rf
    Else
        txtrf = ""
    End If
    If Not IsNull(rsProv!mail) Then
        txtmail = rsProv!mail
    Else
        txtmail = ""
    End If
    If Not IsNull(rsProv!web) Then
        txtweb = rsProv!web
    Else
        txtweb = ""
    End If
    If rsProv!ORDCOMPRA = True Then
        cmbordcompra.ListIndex = 0
    Else
        cmbordcompra.ListIndex = 1
    End If
    If rsProv!activo_pr = True Then
        chkhabilitado.Value = 1
    Else
        chkhabilitado.Value = 0
    End If
    
    CtasComprasSet sSinNull(rsProv!cuentascompras)
    
    txtcuenta = rsProv!Cuenta
    chkTiene_Cuenta.Value = b2k(rsProv!tiene_Cuenta)
    If txtcuenta > "" Then txtCuenta_Descrip = obtenerDeSQL("select descripcion from cuentas where cuenta= " & txtcuenta)
    cboRetenerIIBB.ListIndex = IIf(rsProv!reteneriibb, 1, 0)
    
    chkRetIIBB_p = b2k(rsProv!conretiibbper)
    If chkRetIIBB_p Then
        Aux = "select * FROM ProvTipoRetIB_Per where codprov=" & s2n(txtCodigo) & " order by codigo"
        rs_ret.Open Aux, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        With rs_ret
            If .EOF And .BOF Then
            Else
                For conx = 0 To 3
                    Select Case conx
                        Case 1:
                            Text1 = !BaseImponible
                            Text4 = !Coeficiente
                        Case 2:
                            Text2 = !BaseImponible
                            Text5 = !Coeficiente
                        Case 3:
                            Text3 = !BaseImponible
                            Text6 = !Coeficiente
                    End Select
                    .MoveNext
                Next
            End If
        End With
        Set rs_ret = Nothing
    Else
        Text1 = 0
        Text4 = 0
        Text2 = 0
        Text5 = 0
        Text3 = 0
        Text6 = 0
    End If
    
    
    chkretGAN_p = b2k(rsProv!conretganper)
    If chkretGAN_p Then
        Aux = "select * FROM ProvTipoRetGan_Per where prov=" & val(txtCodigo) & " order by codigo"
        rs_ret.Open Aux, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        With rs_ret
            If .EOF And .BOF Then
            Else
                For conx = 0 To 3
                    Select Case conx
                        Case 1:
                            Text7 = !BaseImponible
                            Text10 = !Coeficiente
                        Case 2:
                            Text8 = !BaseImponible
                            Text11 = !Coeficiente
                        Case 3:
                            Text9 = !BaseImponible
                            Text12 = !Coeficiente
                    End Select
                    .MoveNext
                Next
            End If
        End With
        Set rs_ret = Nothing
    Else
        Text7 = 0
        Text10 = 0
        Text8 = 0
        Text11 = 0
        Text9 = 0
        Text12 = 0
    End If
    
End Sub

Sub HabilitoControles(habilito As Boolean)
    txtCodigo.Locked = habilito
    txtNombre.Locked = habilito
    txtdireccion.Locked = habilito
    txtLocalidad.Locked = habilito
    txtcodpostal.enabled = Not habilito
    cmbprovincias.Locked = habilito
    txtpais.Locked = habilito
    MaskCuit.enabled = Not habilito
    txttel.Locked = habilito
    txtFax.Locked = habilito
    cmbagente.Locked = habilito
    cmbivas.Locked = habilito
    txtsucursal.Locked = habilito
    cmbFormaPagos.Locked = habilito
    txtDescuento.Locked = habilito
    txtdescuento2.Locked = habilito
    cmbprovext.Locked = habilito
    cmbtipocompra.Locked = habilito
    cmbcategoria.Locked = habilito
    txtrf.Locked = habilito
    txtmail.Locked = habilito
    txtweb.Locked = habilito
    cmbordcompra.Locked = habilito
    chkhabilitado.enabled = Not habilito
    chkTiene_Cuenta.enabled = Not habilito
    cmdCuenta.enabled = Not habilito
    
    chkRetIIBB_p.enabled = Not habilito
    chkretGAN_p.enabled = Not habilito
    
    Text1.Locked = habilito
    Text2.Locked = habilito
    Text3.Locked = habilito
    Text4.Locked = habilito
    Text5.Locked = habilito
    Text6.Locked = habilito
    
    Text7.Locked = habilito
    Text8.Locked = habilito
    Text9.Locked = habilito
    Text10.Locked = habilito
    Text11.Locked = habilito
    Text12.Locked = habilito
    
    
End Sub

Sub Habilitobotones(busco As Boolean, nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdbuscar.enabled = busco
    cmdnuevo.enabled = nuevo
    cmdmodificar.enabled = modifico
    cmdeliminar.enabled = elimino
    cmdAceptar.enabled = acepto
    cmdcancelar.enabled = Cancelo
End Sub

Private Sub cmdPrimero_Click()
    rsProv.MoveFirst
    txtCodigo = rsProv!codigo
    CargoProveedor
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsProv.MoveNext
    If Not rsProv.EOF Then
        txtCodigo = rsProv!codigo
        CargoProveedor
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsProv.MoveLast
    txtCodigo = rsProv!codigo
    CargoProveedor
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub cmdanterior_Click()
    rsProv.MovePrevious
    If Not rsProv.BOF Then
        txtCodigo = rsProv!codigo
        CargoProveedor
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rsProv.State = 1 Then
        rsProv.Close
        Set rsProv = Nothing
    End If
End Sub


Private Sub MaskCuit_LostFocus()
Dim rs As New ADODB.Recordset
Dim esta
'If cmdnuevo.enabled Then
    If MaskCuit = "__-________-_" Then Exit Sub
    esta = obtenerDeSQL("Select codigo from prov where cuit='" & MaskCuit & "' and activo=1")
    If IsEmpty(esta) Or IsNull(esta) Then
    Else
        MsgBox "El cuit ya existe,es del proveedor nro: " & esta
        MaskCuit = ""
    End If
    
    
    'rs.Open "Select * from prov where cuit='" & MaskCuit & "' and activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    'If Not rs.EOF Then
    '    MsgBox "El cuit ya existe,es del proveedor nro: " & rs!codigo
    '    MaskCuit.SetFocus
    'End If
'End If
End Sub


Private Sub txtcodcli_LostFocus()
Dim rsCli As New ADODB.Recordset
    If Trim(txtcodcli) <> "" Then
        rsCli.Open "Select * from clientes where codigo=" & val(Trim(txtcodcli)), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If rsCli.EOF Then
            MsgBox "El codigo de Cliente es inexistente,verifiquelo", 48, "Atencion"
            txtcodcli = ""
            txtcodcli.SetFocus
        End If
        rsCli.Close
        Set rsCli = Nothing
    End If
End Sub

Private Sub txtcodigo_GotFocus()

    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)

End Sub

Private Sub txtcodigo_LostFocus()
Dim rs As New ADODB.Recordset

    'rs.Open "Select * from prov where codigo=" & Val(Trim(txtCodigo)), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    'If Not rs.EOF Then
    '    MsgBox "El codigo ya existe,verifiquelo", 48, "Atencion"
    '    txtCodigo.SetFocus
    'End If
End Sub
'Private Sub Txtdescuento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If KeyAscii < 47 Or KeyAscii > 57 Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub
Private Sub Txtdescuento_GotFocus()
    txtDescuento.SelStart = 0
    txtDescuento.SelLength = Len(txtDescuento.Text)
End Sub
Private Sub txtdescuento_LostFocus()
    txtDescuento = s2n(txtDescuento, 4)
    If val(txtDescuento) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtDescuento.SetFocus
    End If
End Sub
'Private Sub txtdescuento2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If KeyAscii < 47 Or KeyAscii > 57 Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub
Private Sub txtDescuento2_GotFocus()
    txtdescuento2.SelStart = 0
    txtdescuento2.SelLength = Len(txtdescuento2.Text)
End Sub

Private Sub txtdescuento2_LostFocus()
    txtdescuento2 = s2n(txtdescuento2, 4)
    If val(txtdescuento2) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtdescuento2.SetFocus
    End If
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
    
End Sub
Private Sub txtnombre_GotFocus()

    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)

End Sub
Private Sub txtnombre_LostFocus()
Dim rs As New ADODB.Recordset
    If cmdnuevo.enabled Then
        rs.Open "Select * from prov where descripcion='" & Trim(txtNombre) & "'  and activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            MsgBox "La razon Social ya existe,es del proveedor nro: " & rs!codigo
        End If
    End If
End Sub
Private Sub Txtweb_GotFocus()

    txtweb.SelStart = 0
    txtweb.SelLength = Len(txtweb.Text)

End Sub
'Private Sub Txtsucursal_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If KeyAscii < 47 Or KeyAscii > 57 Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub
Private Sub Txtsucursal_GotFocus()

    txtsucursal.SelStart = 0
    txtsucursal.SelLength = Len(txtsucursal.Text)

End Sub
Private Sub txtrf_GotFocus()

    txtrf.SelStart = 0
    txtrf.SelLength = Len(txtrf.Text)

End Sub
Private Sub txtlocalidad_GotFocus()

    txtLocalidad.SelStart = 0
    txtLocalidad.SelLength = Len(txtLocalidad.Text)

End Sub

Private Sub txtcodcli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtcodcli_GotFocus()

    txtcodcli.SelStart = 0
    txtcodcli.SelLength = Len(txtcodcli.Text)

End Sub
Private Sub txtDireccion_GotFocus()
    txtdireccion.SelStart = 0
    txtdireccion.SelLength = Len(txtdireccion.Text)
End Sub

Private Sub txtpais_GotFocus()
    txtpais.SelStart = 0
    txtpais.SelLength = Len(txtpais.Text)
End Sub
Private Sub txttel_GotFocus()
    txttel.SelStart = 0
    txttel.SelLength = Len(txttel.Text)
End Sub
Private Sub txtfax_GotFocus()
    txtFax.SelStart = 0
    txtFax.SelLength = Len(txtFax.Text)
End Sub

Private Sub Txtmail_GotFocus()
    txtmail.SelStart = 0
    txtmail.SelLength = Len(txtmail.Text)
End Sub

' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
'   inhibo mov al cargar
' 23/11/4 modif pablo
' 29/11/4 " "  " " cuit lostfocus
' 2/12/4 error resume next en carga, por el cuit del ort
'

