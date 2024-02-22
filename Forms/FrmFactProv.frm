VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFactProv 
   Caption         =   "Factura de Proveedores"
   ClientHeight    =   7860
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "FrmFactProv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7860
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCanjear 
      Caption         =   "A Canje..."
      Height          =   420
      Left            =   10215
      TabIndex        =   113
      Top             =   645
      Width           =   1350
   End
   Begin VB.TextBox txtpais 
      Height          =   285
      Left            =   10100
      TabIndex        =   111
      Tag             =   "6"
      Top             =   1615
      Width           =   1455
   End
   Begin VB.TextBox txtLocalidad 
      Height          =   320
      Left            =   6420
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1995
   End
   Begin VB.TextBox txtDireccion 
      Height          =   320
      Left            =   1200
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   1200
      Width           =   4035
   End
   Begin VB.ComboBox cmbProvincia 
      Height          =   315
      Left            =   9540
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2115
   End
   Begin VB.TextBox txtCodMixto 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "0"
      Top             =   480
      Width           =   3300
   End
   Begin VB.TextBox txtNroIIBB 
      Height          =   285
      Left            =   7815
      TabIndex        =   9
      Top             =   1605
      Width           =   1425
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1575
      Width           =   2190
   End
   Begin Gestion.uNumDoc uNumDoc 
      Height          =   300
      Left            =   4245
      TabIndex        =   8
      Top             =   1590
      Width           =   2625
      _ExtentX        =   4736
      _ExtentY        =   529
   End
   Begin VB.Frame fraMesImputacion 
      Height          =   495
      Left            =   5145
      TabIndex        =   90
      Top             =   -30
      Width           =   3330
      Begin VB.TextBox txtanio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   555
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   150
         Width           =   735
      End
      Begin VB.ComboBox txtmes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmFactProv.frx":08CA
         Left            =   1830
         List            =   "FrmFactProv.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Año"
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
         Left            =   75
         TabIndex        =   94
         Top             =   150
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Mes"
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
         Left            =   1365
         TabIndex        =   93
         Top             =   150
         Width           =   615
      End
   End
   Begin Gestion.ucFecha uFeHa 
      Height          =   270
      Left            =   2355
      TabIndex        =   88
      Top             =   7425
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      FechaInit       =   4
   End
   Begin Gestion.ucFecha uFeDe 
      Height          =   270
      Left            =   1515
      TabIndex        =   87
      Top             =   7425
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   476
      FechaInit       =   5
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   7425
      Width           =   975
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   5265
      Left            =   0
      TabIndex        =   29
      Top             =   2055
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9287
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Factura"
      TabPicture(0)   =   "FrmFactProv.frx":091D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblImputaciones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbingresar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Forma de Pago"
      TabPicture(1)   =   "FrmFactProv.frx":0939
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Imputaciones Contables"
      TabPicture(2)   =   "FrmFactProv.frx":0955
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "uTipoCompra"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmbingresar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingresar"
         Enabled         =   0   'False
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   4800
         UseMaskColor    =   -1  'True
         Width           =   1065
      End
      Begin VB.Frame fraFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4350
         Left            =   75
         TabIndex        =   63
         Top             =   480
         Width           =   11475
         Begin VB.CommandButton cmdQuitarProv 
            Height          =   300
            Left            =   9180
            Picture         =   "FrmFactProv.frx":0971
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   2505
            Width           =   345
         End
         Begin VB.CommandButton cmdAgregarProv 
            Caption         =   "Agregar"
            Height          =   300
            Left            =   8145
            TabIndex        =   102
            Top             =   2505
            Width           =   1005
         End
         Begin VSFlex7LCtl.VSFlexGrid gIIBBProvincia 
            Height          =   1290
            Left            =   6525
            TabIndex        =   101
            Top             =   2925
            Width           =   4890
            _cx             =   8625
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
         Begin VB.ComboBox cmbtipocompra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9540
            Style           =   2  'Dropdown List
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtNoGrabado 
            Height          =   285
            Left            =   6915
            TabIndex        =   18
            Text            =   "0"
            Top             =   1425
            Width           =   1215
         End
         Begin VB.ComboBox cmbformapago 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   225
            Width           =   3840
         End
         Begin VB.TextBox txtPer3337 
            Height          =   285
            Left            =   1695
            TabIndex        =   22
            Text            =   "0"
            Top             =   2115
            Width           =   1215
         End
         Begin VB.TextBox txtIva 
            Height          =   285
            Left            =   1710
            TabIndex        =   19
            Text            =   "0"
            Top             =   1755
            Width           =   1215
         End
         Begin VB.TextBox txtImporte 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   1710
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1065
            Width           =   1215
         End
         Begin VB.TextBox txtNeto 
            Height          =   285
            Left            =   1710
            TabIndex        =   16
            Text            =   "0"
            Top             =   1410
            Width           =   1215
         End
         Begin VB.TextBox txtImpInt 
            Height          =   285
            Left            =   6900
            TabIndex        =   24
            Text            =   "0"
            Top             =   2145
            Width           =   1230
         End
         Begin VB.TextBox txtIva27 
            Height          =   285
            Left            =   4320
            TabIndex        =   20
            Text            =   "0"
            Top             =   1770
            Width           =   1215
         End
         Begin VB.TextBox txtExento 
            Height          =   285
            Left            =   4335
            TabIndex        =   17
            Text            =   "0"
            Top             =   1425
            Width           =   1215
         End
         Begin VB.TextBox txtRetenIva 
            Height          =   285
            Left            =   4320
            TabIndex        =   26
            Text            =   "0"
            Top             =   2475
            Width           =   1215
         End
         Begin VB.TextBox txtIBcapital 
            Height          =   285
            Left            =   1710
            TabIndex        =   27
            Text            =   "0"
            Top             =   2805
            Width           =   1215
         End
         Begin VB.TextBox txtRetenGan 
            Height          =   285
            Left            =   1710
            TabIndex        =   25
            Text            =   "0"
            Top             =   2460
            Width           =   1215
         End
         Begin VB.TextBox txtPer3431 
            Height          =   285
            Left            =   4320
            TabIndex        =   23
            Text            =   "0"
            Top             =   2115
            Width           =   1215
         End
         Begin VB.TextBox txtIva10 
            Height          =   285
            Left            =   6915
            TabIndex        =   21
            Text            =   "0"
            Top             =   1785
            Width           =   1215
         End
         Begin VB.TextBox txtcotiz 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4305
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   690
            Width           =   1215
         End
         Begin VB.ComboBox cmbmoneda 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "FrmFactProv.frx":0EFB
            Left            =   1695
            List            =   "FrmFactProv.frx":0EFD
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   675
            Width           =   1215
         End
         Begin VB.TextBox txtIBprovincia 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0"
            Top             =   2505
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker txtfvto 
            Height          =   300
            Left            =   6555
            TabIndex        =   11
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   188153857
            CurrentDate     =   37934
         End
         Begin VB.Label lbltipocompra 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Compra"
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
            Left            =   8280
            TabIndex        =   100
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "No Grabado"
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
            Index           =   1
            Left            =   5745
            TabIndex        =   99
            Top             =   1455
            Width           =   1215
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "F. Vto."
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
            Left            =   5775
            TabIndex        =   83
            Top             =   315
            Width           =   705
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Forma de Pago"
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
            Left            =   -30
            TabIndex        =   79
            Top             =   255
            Width           =   1605
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Per. RG3337"
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
            Left            =   0
            TabIndex        =   78
            Top             =   2175
            Width           =   1605
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iva 21%"
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
            Left            =   0
            TabIndex        =   77
            Top             =   1770
            Width           =   1605
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Neto"
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
            Left            =   0
            TabIndex        =   76
            Top             =   1440
            Width           =   1605
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Total"
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
            Left            =   0
            TabIndex        =   75
            Top             =   1110
            Width           =   1605
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Imp Int"
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
            Index           =   0
            Left            =   6180
            TabIndex        =   74
            Top             =   2175
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iva 27%"
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
            Left            =   2970
            TabIndex        =   73
            Top             =   1770
            Width           =   1215
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Exento"
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
            Left            =   2970
            TabIndex        =   72
            Top             =   1455
            Width           =   1215
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reten. Iva"
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
            Left            =   2955
            TabIndex        =   71
            Top             =   2475
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IB Capital"
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
            Left            =   0
            TabIndex        =   70
            Top             =   2850
            Width           =   1605
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "R. Ganancia"
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
            Left            =   0
            TabIndex        =   69
            Top             =   2490
            Width           =   1605
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Per. RG3431"
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
            Left            =   2970
            TabIndex        =   68
            Top             =   2100
            Width           =   1215
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iva 10.5%"
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
            Left            =   5955
            TabIndex        =   67
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblcotiz 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cotiz."
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
            Left            =   3600
            TabIndex        =   66
            Top             =   720
            Width           =   780
         End
         Begin VB.Label lblmoneda 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
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
            Left            =   750
            TabIndex        =   65
            Top             =   705
            Width           =   735
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IB provincia"
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
            Left            =   5550
            TabIndex        =   64
            Top             =   2505
            Width           =   1215
         End
      End
      Begin VB.Frame fraContado 
         BorderStyle     =   0  'None
         Height          =   4650
         Left            =   -74955
         TabIndex        =   57
         Top             =   435
         Width           =   11415
         Begin Gestion.ucRetCompras uRetCompras 
            Height          =   705
            Left            =   1725
            TabIndex        =   32
            Top             =   90
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   1244
         End
         Begin VB.TextBox txtTotalRetPago 
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtimpcheques 
            Enabled         =   0   'False
            Height          =   300
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1605
            Width           =   1320
         End
         Begin VB.TextBox txtefectivo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   45
            TabIndex        =   33
            Top             =   825
            Width           =   1335
         End
         Begin VB.TextBox txttransf 
            Enabled         =   0   'False
            Height          =   285
            Left            =   75
            TabIndex        =   39
            Top             =   4305
            Width           =   1215
         End
         Begin VB.TextBox txtcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4425
            TabIndex        =   43
            Tag             =   "2"
            Top             =   4335
            Width           =   3165
         End
         Begin VB.CommandButton cmbcuenta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cuenta"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3435
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtcodcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2175
            TabIndex        =   41
            Top             =   4335
            Width           =   1215
         End
         Begin VB.TextBox txtcodcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2355
            TabIndex        =   34
            Text            =   "1"
            Top             =   870
            Width           =   825
         End
         Begin VB.CommandButton cmbcaja 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caja"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3300
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   855
            Width           =   855
         End
         Begin VB.TextBox txtcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4290
            TabIndex        =   30
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   855
            Width           =   2535
         End
         Begin Gestion.ucCheques uCheques 
            Height          =   2700
            Left            =   1440
            TabIndex        =   37
            Top             =   1335
            Width           =   9930
            _ExtentX        =   14949
            _ExtentY        =   3307
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ret"
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
            Left            =   105
            TabIndex        =   82
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques"
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
            Left            =   45
            TabIndex        =   62
            Top             =   1305
            Width           =   870
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Efectivo"
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
            Left            =   60
            TabIndex        =   61
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transf."
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
            Left            =   75
            TabIndex        =   60
            Top             =   3990
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
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
            Left            =   1410
            TabIndex        =   59
            Top             =   4335
            Width           =   735
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caja"
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
            Left            =   1875
            TabIndex        =   58
            Top             =   870
            Width           =   855
         End
      End
      Begin Gestion.ucTipoCompra uTipoCompra 
         Height          =   3870
         Left            =   -74925
         TabIndex        =   40
         Top             =   645
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   6826
      End
      Begin VB.Label lblImputaciones 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Costos"
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
         Height          =   375
         Left            =   8640
         TabIndex        =   98
         Top             =   4800
         Width           =   1695
      End
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8775
      TabIndex        =   56
      Top             =   7860
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6255
      TabIndex        =   55
      Tag             =   "1"
      Top             =   7935
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnumcompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6900
      TabIndex        =   54
      Top             =   7995
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttipocompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5355
      TabIndex        =   53
      Top             =   7890
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtfechacompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7875
      TabIndex        =   52
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtserie 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.OptionButton optcontado 
      Caption         =   "Contado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton optctacte 
      Caption         =   "Cta. Cte."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame Framedevol 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   120
      TabIndex        =   38
      Top             =   -60
      Width           =   2535
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7410
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   7395
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7410
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7380
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7410
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   7440
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   188350465
      CurrentDate     =   37934
   End
   Begin Gestion.ucCoDe uProv 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   6510
      _ExtentX        =   9657
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCuit Cuit 
      Height          =   285
      Left            =   8685
      TabIndex        =   6
      Top             =   870
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
   End
   Begin VB.Label Label35 
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
      Left            =   9520
      TabIndex        =   112
      Top             =   1615
      Width           =   615
   End
   Begin VB.Label Label34 
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
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   220
      TabIndex        =   110
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label33 
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
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   5400
      TabIndex        =   109
      Top             =   1230
      Width           =   915
   End
   Begin VB.Label Label32 
      Caption         =   "Provincia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   8520
      TabIndex        =   108
      Top             =   1230
      Width           =   915
   End
   Begin VB.Label Label31 
      Caption         =   "NroIIBB"
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
      Left            =   7005
      TabIndex        =   96
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "IVA"
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
      TabIndex        =   95
      Top             =   1575
      Width           =   435
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "entre"
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
      Height          =   225
      Left            =   1035
      TabIndex        =   89
      Top             =   7395
      Width           =   645
   End
   Begin VB.Label lblIDDOC 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   10770
      TabIndex        =   86
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label27 
      Caption         =   "iddoc:"
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
      Height          =   225
      Left            =   10155
      TabIndex        =   85
      Top             =   105
      Width           =   645
   End
   Begin VB.Label lblcuit 
      Caption         =   "Cuit"
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
      Left            =   8115
      TabIndex        =   81
      Top             =   870
      Width           =   570
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      TabIndex        =   80
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "Serie"
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
      Left            =   8580
      TabIndex        =   51
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Factura:"
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
      TabIndex        =   50
      Top             =   1620
      Width           =   870
   End
   Begin VB.Label Label1 
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
      Left            =   2880
      TabIndex        =   49
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmFactProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 30/3/5
   

'------- POR REMITO -----------------
Private mPorRemito As Boolean
'------- POR REMITO -----------------


Private Const IVA_21 = 0.21
Private Const IVA_27 = 0.27
Private Const IVA_105 = 0.105
Private EsBusqueda As Boolean
'Private pAsiento As Prod_Asiento
Private pAsiento As Collection
Dim Ope As String
Dim rsmov As New ADODB.Recordset
Dim midDoc As Long

Private Sub cboIva_LostFocus()
    seteoLetra
    If Trim(cboIva.Text) = "MONOTRIBUTISTA" Then
        HabilitoControles2 False, True
    Else
        HabilitoControles2 True, True
    End If
End Sub

Private Sub cmbcaja_Click()
    Dim re
    re = frmBuscar.MostrarSql("select Codigo, Responsable  from cajas where activo = 1")
    If re > "" Then
        txtcodcaja = frmBuscar.resultado(1)
        txtcaja = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmbcuenta_Click()
    Dim re
    re = frmBuscar.MostrarSql("SELECT CTASBANK.CODIGO AS [ Codigo   ], CTASBANK.BANCO AS [Cod Banco ], BancosGrales.descripcion AS [ Nombre Banco                             ], CTASBANK.NUMERO, Monedas.descripcion AS [ Moneda        ] FROM CTASBANK LEFT OUTER JOIN                     BancosGrales ON CTASBANK.BANCO = BancosGrales.codigo LEFT OUTER JOIN                      Monedas ON Monedas.codigo = CTASBANK.MONEDA Where (CTASBANK.ACTIVO = 1)")
    If re > "" Then
        txtcodcuenta = frmBuscar.resultado(1)
        txtcuenta = frmBuscar.resultado(3) & " - " & frmBuscar.resultado(4)
    End If
End Sub

Private Sub cmbFormaPago_LostFocus()
    txtfvto = dtFecha + s2n(obtenerDeSQL("select dias from FormasPago where descripcion = '" & cmbformapago & "' and activo = 1"))
End Sub

Private Sub cmbingresar_Click()
    Dim Neto As Double
    
    On Error Resume Next
    If txtefectivo <> "" And txtefectivo <> "0" And txtcodcaja = "" Then
        MsgBox "Falta ingresar el código de caja del importe en efectivo", vbInformation
        txtcodcaja.SetFocus
        Exit Sub
    End If

    If txttransf <> "" And txttransf <> "0" And txtcodcuenta = "" Then
        MsgBox "Falta ingresar el código de cuenta del importe de transferencia", vbInformation
        txtcodcuenta.SetFocus
        Exit Sub
    End If

    If txtimporte <> "" And txtimporte <> "0" Then
        If s2n(txtimporte) <> s2n(s2n(s2n(txtneto) + s2n(txtNoGrabado) + s2n(txtIva) + s2n(txtper3337) + s2n(txtIva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtIva10) + s2n(txtIBcapital) + s2n(txtIBprovincia))) Then
            MsgBox "Los totales no concilian, hay una diferencia de: " & s2n((s2n(txtneto) + s2n(txtNoGrabado) + s2n(txtIva) + s2n(txtper3337) + s2n(txtIva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtIva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)) - s2n(txtimporte))
            Exit Sub
        End If
        
        If optcontado = True And s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf) <> s2n(txtimporte) Then
            MsgBox "El total de la forma de pago no coincide con el importe de la factura"
            Exit Sub
        End If
    Else
        MsgBox "Debe ingresar el importe de la factura"
        txtimporte.SetFocus
        Exit Sub
    End If
    
    If optcontado = True And ((txtimpcheques = "" And txtefectivo = "" And txttransf = "") Or (s2n(txtimpcheques) + s2n(txtefectivo) + s2n(txttransf) <> txtimporte)) Then
        MsgBox "No coinciden los totales en la forma de pago con el importe total a pagar"
        Exit Sub
    End If
    
    FrmCostosYContable.Tag = Me.Name
    vieneDE = Me.Name
    cmdok.SetFocus
    FrmCostosYContable.CargarImputacion s2n(txtneto) + s2n(txtexento), s2n(txtimporte), 1
    FrmCostosYContable.Show
End Sub

Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub cmdAgregarProv_Click()
Dim r, pRow As Long, rRepetido As Boolean
r = frmBuscar.MostrarSql("Select [CODIGO           ],[DESCRIPCION             ] from provincias where activo=1")
If r <> "" Then
    rRepetido = False
    For pRow = 1 To gIIBBProvincia.rows - 1
        If gIIBBProvincia.TextMatrix(pRow, 0) = r Then rRepetido = True
    Next
    
    If Trim(frmBuscar.resultado(1)) = "*" Then
        MsgBox "Esta tratando de agregar una jurisdiccion que puede ingresarlo en el campo IIBB capital." & Chr(13) & "Controle si es correcto agregarlo.", , "ATENCION"
    End If

    If Not rRepetido Then
        With gIIBBProvincia
            .AddItem ""
            pRow = .rows - 1
            .TextMatrix(pRow, 0) = frmBuscar.resultado(1)
            .TextMatrix(pRow, 1) = frmBuscar.resultado(2)
            .TextMatrix(pRow, 2) = dtFecha
            .TextMatrix(pRow, 3) = "0,00"
        End With
    End If
    If rRepetido Then
        MsgBox frmBuscar.resultado(2) & " ya existe en la lista.", vbInformation
    End If
End If

calGrillaProv
End Sub

Private Sub cmdBuscar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaBuscar
    
    Dim tabla, re, fech As Date, ssql As String, titu As String
    
    If optcontado = True Then
        tabla = "compras"
        titu = "Contado y canceladas"
    Else
        tabla = "transcom"
        titu = "Pendientes en Cuenta Corriente"
    End If
    ssql = " select Fecha as [Fecha ], tipoDoc as [Doc], NroDoc as [ Numero ], total as  [ Importe ], codPr as [ Prov ], descripcion as [ Razon social                           ], iddoc " & _
        " from " & tabla & " left join prov on codpr = prov.codigo  " & _
        " where " & tabla & ".activo = 1 and tipodoc = 'FAC' and fecha " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & _
        " order by fecha desc "
    
    re = frmBuscar.MostrarSql(ssql, , titu) ', , , , " " , "Anulada")
    If re > "" Then
        cmdCancelar_Click
        Call Habilitobotones(True, True, True, False, True, True)
        fech = CDate(frmBuscar.resultado(1))
        txttipocompra = frmBuscar.resultado(2)
        txtnumcompra = frmBuscar.resultado(3)

        rsmov.Open "select * from " & tabla & " where fecha = " & ssFecha(fech) & " and tipodoc = '" & txttipocompra & "' and nrodoc = " & val(txtnumcompra) & " and codpr=" & frmBuscar.resultado(5) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        CargoRegistro
        DetalleIIBB "C", txttipocompra, txtnumcompra, frmBuscar.resultado(5), rsmov!iddoc, s2n(txtcotiz)
        EsBusqueda = True
    Else
        EsBusqueda = False
    End If
    
fin:
    Set rsmov = Nothing
    Exit Sub
UfaBuscar:
    ufa "err la cargar", "buscar " & Me.Name ', Err.Number
    Resume fin
End Sub


Private Sub cmdCancelar_Click()
    LimpioControles
    HabilitoControles (False)
    FormadePago (False)
    Call Habilitobotones(True, True, False, False, True, False)
    uTipoCompra.Borrar
    EsBusqueda = False
    Set pAsiento = New Collection
    FrmCostosYContable.LimpioControles
'    FrmCostosYContable.InicioGrilla
    FrmCostosYContable.InicioGrillaCostos
End Sub

Private Sub cmdCanjear_Click()
Dim ss As String
If s2n(lblIDDOC) > 0 Then
    ss = "update transcom set formadepago=-1 where iddoc=" & s2n(lblIDDOC)
    
    DataEnvironment1.Sistema.Execute ss
    
    MsgBox "Factura colocada en canje...", vbInformation
Else
    MsgBox "Busque la factura...", vbExclamation
End If
End Sub

Private Sub cmdeliminar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelimina

    Dim mensaje As Long
    Dim rs As New ADODB.Recordset
    Dim conttot, cont, iddoc
    Dim impo As Double
    
    If Not PuedoCompras(dtFecha) Then
        'msg en funcion
        Exit Sub
    End If
    
    Dim sp
    sp = obtenerDeSQL("select * from relfnr_c d inner join rec_comp o on d.ndoc=o.nro where o.codpr=" & UpROV.codigo & " and o.activo=1 and d.tfac='FAC' and d.fact=" & s2n(uNumDoc.num))
    If IsNull(sp) Or IsEmpty(sp) Then
    Else
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If
    
    
    
    mensaje = MsgBox("Esta seguro que desea elimnar este registro?", vbYesNo, "Atencion")
    If mensaje = vbYes Then
    
    
    
    
        '*************************************************
        DE_BeginTrans
        
        If optcontado = True Then
        
            conttot = 0
            cont = 0
            rs.Open "select * from Chq_comp where tipodoc = 'FAC' and nrodoc = " & uNumDoc.num & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                conttot = conttot + 1
                If rs!estado <> "B" Then
                    cont = cont + 1
                End If
                rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            ' cont    = s2n(obtenerdesql("select count(??) from Chq_comp where tipodoc = 'FAC' and nrodoc = " & Val(txtfact) & " and activo = 1 and estado <>'B' " ))
            ' conttot = s2n(obtenerdesql("select count(??) from Chq_comp where tipodoc = 'FAC' and nrodoc = " & Val(txtfact) & " and activo = 1  " ))
            
            If conttot = cont Then
                    DataEnvironment1.dbo_INGCOMPRAS "B", 0, 0, 0, UpROV.codigo, "", "", 0, "", uNumDoc.num _
                        , 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, val(txtserie), 0, 0, 0, 0, 0, Date, UsuarioSistema!codigo, midDoc, 0, 0 _
                        , uNumDoc.letra, txtNroIIBB, uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo
                    
                    rs.Open "select * from Movibanc where fecha = " & ssFecha(dtFecha) & " and tipdoc = 'FAC' and nrodoc = " & uNumDoc.num & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    While Not rs.EOF
                        DataEnvironment1.dbo_INGCOMPRAMOVIBANC "B", 0, "", "" _
                            , 0, "", 0, 0, "", 0, rs!MovBanco, midDoc, Date, UsuarioSistema!codigo, 1
                        rs.MoveNext
                    Wend
                    rs.Close
                    Set rs = Nothing
                
                    rs.Open "select * from Movicaja where fecha = " & ssFecha(dtFecha) & " and tipodoc = 'FAC' and nrodoc = " & uNumDoc.num & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    While Not rs.EOF
                        DataEnvironment1.dbo_INGCOMPRAMOVICAJA "B", 0, rs!movimiento, "", "", 0, "" _
                            , 0, 0, 0, "", 0, "", 0, midDoc, Date, UsuarioSistema!codigo, 1
                        rs.MoveNext
                    Wend
                    rs.Close
                    Set rs = Nothing
                    
                    DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "B", 0, 0, 0, uNumDoc.num, TIPODOC_FAC_PROVEEDOR, 0, "", 0, 0, 0, 0, Date, UsuarioSistema!codigo, 0, 1, 0
                    DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "B", 0, 0, "", UpROV.codigo, uNumDoc.num, 0 _
                        , 0, "", "FDC", 0, 0, Date, UsuarioSistema!codigo, 0, 1, 1

                    grabaBitacora "B", UpROV.codigo, "compras"
            Else
                DE_RollbackTrans
                MsgBox "No se puede dar de baja dado que uno o mas cheques propios ya fueron debitados"
                Exit Sub
            End If
        Else  ' Cta Cte
            If Trim(cmbMoneda.Text) <> "Pesos" Then
                impo = s2n(txtcotiz.Text, 4)
                If impo = 0 Then impo = 1
                impo = s2n(txtimporte * impo)
            Else
                impo = s2n(txtimporte)
            End If
            If s2n(txtsaldo) = impo Then
                    DataEnvironment1.dbo_INGCOMPRASCTACTE "B", 0, 0, 0, UpROV.codigo, "", "", 0, "", uNumDoc.num, _
                        0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, val(txtserie), 0, 0, 0, 0, 0, 0, 0, Date, UsuarioSistema!codigo, midDoc, 0, 0 _
                        , uNumDoc.letra, txtNroIIBB, uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo
                    
                    grabaBitacora "B", UpROV.codigo, "dbo_INGCOMPRASCTACTE"
            Else
                DE_RollbackTrans
                MsgBox "No se puede anular este coprobante dado que fue parcialmente pagado"
                Exit Sub
            End If
        End If
        


'------- POR REMITO ----------------- Val(txtcodprov), "FAC", Val(txtfact), 0
        ItemFacturaRemitoCompra ItemFRC_BajaFactura, 0, UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, 0, 0
'------- POR REMITO -----------------
           
        
            'Baja Doc y asiento.
            'If siAsiento("AsientosCompras") Then
                If midDoc > 0 Then ' SI fue generado con este sistema, bajo asiento
                    If Not BorroDocumento(midDoc) Then
                        ufa "err al borrar documento", " middoc = " & midDoc

                        DE_RollbackTrans
                        GoTo FinElimina:
                    End If
                End If
            'Else
            '    DataEnvironment1.Sistema.Execute _
                    " update RegistroDocumentos " & _
                    " set activo = 0, usuario_baja = " & UsuarioActual() & " , fecha_baja = " & ssFecha(Date) & _
                    " where iddoc = '" & midDoc & "' "
            'End If
            DetalleIIBB "B", TIPODOC_FAC_PROVEEDOR, uNumDoc.num, UpROV.codigo, midDoc
            
        DE_CommitTrans
        '*************************************************

        MsgBox "La factura se ha anulado correctamente"
        LimpioControles
        HabilitoControles (True)
        Call Habilitobotones(True, True, False, False, True, False)
        cmbingresar.enabled = False
    End If
    
FinElimina:
    Set rs = Nothing
    Exit Sub
UFAelimina:
    DE_RollbackTrans
    ufa "Err al anular", Me.Name
    Resume FinElimina
End Sub

Private Sub cmdmodificar_Click()
    HabilitoControles (False)
    Call Habilitobotones(True, True, False, False, False, False)
    Ope = "M"
End Sub




Private Sub cmdnuevo_Click()
    Dim rs As New ADODB.Recordset
    EsBusqueda = False
    LimpioControles
    
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    
    HabilitoControles (True)
    FormadePago (False)
    Call Habilitobotones(False, False, False, True, True, False)
    
    uCheques.Borrar
    
    Ope = "A"
    '------- POR REMITO -----------------
    Dim tmpi
    If mPorRemito Then
        With frmRemitosCompraPendientes
            Set pAsiento = New Collection
            Set pAsiento = .mostrar
            If s2n(.Total) = 0 Then
                'Unload Me
            Else
                UpROV.codigo = .ProveedorCod
                txtneto = s2n(.Total)
                optctacte.Value = vbChecked
                dtFecha.SetFocus
            End If
        End With
    End If
    '------- POR REMITO -----------------
    
    optctacte.Value = vbChecked
    txtanio = Year(Date) 'Year(dtfecha)
    txtmes = Month(Date) 'Month(dtfecha)
    txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub


Public Function DetalleIIBB(oP As String, Optional dTipo As String = "", Optional dNumero As String = "0", Optional Dprov As Long = 0, Optional iddoc As Long, Optional zCotizacion As Double = 1)
Dim aTabla As String, aConsult As String
Dim i As Long
aTabla = " IIBBJURISDICCION "
With gIIBBProvincia
    Select Case oP
        Case "A":
            If .rows > 1 Then
                For i = 1 To .rows - 1
                    aConsult = "INSERT INTO " & aTabla & " (NRODOC,TIPODOC,FECHA,IMPORTE,CODJUR,JURISDICCION,ACTIVO,CODDOC,iddoc) VALUES (" _
                            & "" & ssTexto(dNumero) & "," & ssTexto(dTipo) & "," & ssFecha(.TextMatrix(i, 2)) & "," & x2s(s2n(s2n(.TextMatrix(i, 3)) * zCotizacion)) & " , " & ssTexto(.TextMatrix(i, 0)) & "," & ssTexto(.TextMatrix(i, 1)) & ",1," & Dprov & "," & iddoc & ")"
                    DataEnvironment1.Sistema.Execute aConsult
                Next
            End If
        Case "B":
            aConsult = "UPDATE " & aTabla & " SET ACTIVO=0 WHERE TIPODOC=" & ssTexto(dTipo) & " AND NRODOC=" & ssTexto(dNumero) & " and CODDOC=" & Dprov & " and iddoc=" & iddoc
            DataEnvironment1.Sistema.Execute aConsult
        Case "C":
            aConsult = "SELECT CODJUR,JURISDICCION ,FECHA,(IMPORTE /" & x2s(zCotizacion) & ") AS IMPORTE FROM " & aTabla & " WHERE ACTIVO=1 AND TIPODOC=" & ssTexto(dTipo) & " AND NRODOC=" & ssTexto(dNumero) & " and coddoc=" & Dprov & " and iddoc=" & iddoc
            LlenarGrilla gIIBBProvincia, aConsult, False
            If .rows > 1 Then
                .ColWidth(0) = 0
                .ColWidth(1) = 1500
                .ColWidth(2) = 1500
                .ColWidth(3) = 1500
            End If
        Case "I":
            .rows = 1
            .cols = 4
            .TextMatrix(0, 0) = "CODJUR"
            .TextMatrix(0, 1) = "JURISDICCION"
            .TextMatrix(0, 2) = "FECHA"
            .TextMatrix(0, 3) = "IMPORTE"
            
            .ColWidth(0) = 0
            .ColWidth(1) = 1500
            .ColWidth(2) = 1500
            .ColWidth(3) = 1500
            .Editable = flexEDKbdMouse
    End Select
End With
End Function

Private Sub cmdOk_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaOK

    Dim iddoc As Long, NroDoc As Long ', tdoc As String
    Dim rs As New ADODB.Recordset
    Dim AsientoCompra As New Asiento
    Dim fechapropio As Date ', valcartera
    Dim sAssert As String 'para seguimiento error
    Dim resu As String
    
    Dim NroPago As Long     ' para registrodocumentos, nro unico de pago
    Dim NroCertifIIBB As Long
    Dim NroCertifGan As Long
    
    Dim cueche  As Long  'cuenta bancaria del chequepropio
    Dim tiene_c
    Dim z As Double
    Dim k As Long
    Dim a As String
    
    Dim cYear As String, cMonth As String, cNrodoc As String, cCodMixto As String
    
    If Trim(cmbMoneda.Text) <> "Pesos" Then
        z = s2n(txtcotiz, 4)
        If z = 0 Then z = 1
    Else
        z = 1
    End If
    If UCase(cmbMoneda) <> "PESOS" Then
        If s2n(txtcotiz, 4) = 0 Then
            MsgBox "No se ingreso la cotización correspondiente.", , "ATENCION"
            Exit Sub
        End If
    End If
    
    If TrabaIva(dtFecha.Value) Then
        MsgBox "La fecha del comprobante esta dentro de las fechas trabadas para emision," & Chr(13) & "verifiquelo con su contadora.", , "ATENCION"
        Exit Sub
    End If
    
'    Dim ntipocompra As Long
''     'ntipocompra = s2n(ObtenerCodigo("Tipocompras", cmbtipocompra.Text))
    If Not PuedoCompras(dtFecha) Then
        'msg en funcion
        Exit Sub
    End If

    EsBusqueda = False
    If optcontado = False And optctacte = False Then
        MsgBox "Debe ingresar el tipo de pago, contado o Cta. Cte.", 48, "Atencion"
        Exit Sub
    Else
    
    If s2n(txtimporte) = 0 Then
        che "Falta Monto de factura"
        Exit Sub
    End If
    
    If uNumDoc.num = 0 Then
        MsgBox "Debe ingresar el número de factura"
         uNumDoc.SetFocus
        Exit Sub
    End If
    
    If Trim(UpROV.DESCRIPCION) = "" Then
        MsgBox "Debe ingresar proveedor"
        UpROV.SetFocus
        Exit Sub
    End If
    
    If Not optcontado And UpROV.codigo = 0 Then
        MsgBox "Debe ingresar proveedor"
        UpROV.SetFocus
        Exit Sub
    End If
    
    'If uProv.codigo > 0 Then
     '   resu =
    If ExisteFacCompraMSG(UpROV.codigo, uNumDoc.suc, uNumDoc.num) Then
     '   If resu > "" Then
     '       che "Documento existente " & resu
        uNumDoc.SetFocus
        Exit Sub
    End If
     '   End If
    'End If
    
    If cmbformapago.ListIndex = -1 And optctacte = True Then
        MsgBox "Debe ingresar la forma de pago"
        cmbformapago.SetFocus
        Exit Sub
    End If
    
    If s2n(txtefectivo) > 0 And s2n(txtcodcaja) = 0 Then
        che "Falta Nro Caja para efectivo"
        Exit Sub
    End If
    
    If s2n(txttransf) <> 0 And Trim(txtcuenta) = "" Then
        che "Falta cuenta transferencia"
        Exit Sub
    End If
    
    If cmbtipocompra.ListIndex = -1 Then
        MsgBox "Debe ingresar el tipo de compra"
        cmbtipocompra.SetFocus
        Exit Sub
    End If
    
    k = 1
    While k < gIIBBProvincia.rows
        If gIIBBProvincia.TextMatrix(k, 3) = 0 Then
            MsgBox "Hay una jurisdiccion que no posee importe.", , "ATENCION"
            Exit Sub
        End If
        If Trim(gIIBBProvincia.TextMatrix(k, 0)) = "*" Then
            If MsgBox("La jurisdiccion -CAPITAL FEDERAL- puede ingresarlo en el campo IIBB capital." & Chr(13) & "Controle si es correcto agregarlo." & Chr(13) & "Desea frenar y controlarlo?", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
                Exit Sub
            End If
        End If
        k = k + 1
    Wend
    
    
'    If optcontado = True And s2n(txtimporte) <> s2n(s2n(txtneto) + s2n(txtIva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)) Then
    Dim dife As Double
    dife = s2n((s2n(txtimporte)) - (s2n(s2n(txtneto) + s2n(txtIva) + s2n(txtper3337) + s2n(txtIva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtIva10) + s2n(txtIBcapital) + s2n(txtIBprovincia) + s2n(txtNoGrabado))))
    If dife <> 0 Then
        che "Los totales no concilian, hay una diferencia de: " & dife 's2n(s2n(txtimporte) - (s2n(txtneto) + s2n(txtIva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)))
        Exit Sub
    End If
    
    If optcontado = True And (s2n(txtimpcheques) <> 0) And uCheques.Total = 0 Then
         MsgBox "No se ingresaron los cheques correspondientes al pago"
        Exit Sub
    End If
    
    If s2n(txtimpcheques) <> uCheques.Total Then
        che "No coincide el total de cheques"
        Exit Sub
    End If
    
    If Not uCheques.FechasOk Then Exit Sub
    
    If optcontado = True And s2n(s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf) + uRetCompras.TotalRet - s2n(txtimporte)) <> 0 Then
        MsgBox "El total de la forma de pago no coincide con el importe de la factura"
        Exit Sub
    End If
    
    If gEMPR_ConSistContable Then
        If uTipoCompra.Diferencia <> 0 Or uTipoCompra.Total_a_Imputar = 0 Then
            che "Falta completar imputaciones contables"
            Exit Sub
        End If
    End If
    
    'If UCase(cmbmoneda) <> "PESO" Then
    '    MsgBox "No se ingreso la cotización correspondiente"
    '    Exit Sub
    'End If

    Dim strSql As String
    Dim cheque As Boolean, contado As Boolean
    Dim Total As Double
    Dim sumo As Double, i As Long
    Dim TextoAsientoComprobante As String
    NroDoc = uNumDoc.num
     
    TextoAsientoComprobante = "FC " & NroDoc
    With AsientoCompra
        'HEADER ASIENTO
        .nuevo "Fac " & UpROV.DESCRIPCION, dtFecha, TIPODOC_FAC_PROVEEDOR
        'DEBE
        For i = 1 To uTipoCompra.rows
            .AgregarItem uTipoCompra.imCuenta(i), uTipoCompra.imMonto(i) * z, 0  ', TextoAsientoComprobante
            'iva, iibb, perc, retgan
        Next i
'        .AgregarItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), sumaTxtIvas(), 0
'        .AgregarItem CuentaParam(ID_Cuenta_C_IIBB_COMPRA), s2n(txtIBcapital) + s2n(txtIBprovincia), 0
'        .AgregarItem CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(txtretengan), 0
'        .AgregarItem CuentaParam(ID_Cuenta_C_RET_IVA_CPRA), s2n(txtreteniva), 0
        

    End With
    
    Dim numIntP As Long
    
    '*************************************************
    DE_BeginTrans

    If optcontado = True Then
        NroPago = NuevoNroPago()
        If uRetCompras.retgan > 0 Then NroCertifGan = NuevoNroCertifGan()
        If uRetCompras.retIB > 0 Then NroCertifIIBB = NuevoNroCertifIIBB()
    End If
    
    
    iddoc = NuevoDocumento(TIPODOC_FAC_PROVEEDOR, NroDoc, UpROV.codigo, NroPago, NroCertifGan, NroCertifIIBB, uNumDoc.suc)
    midDoc = iddoc
    
    If Trim(Ope) <> "" Then
        If Ope = "A" Then
                   
            Dim porciva As Double, maximobanc As Long, maximocaja As Long, x As Long
            Dim valorcuenta As String, valorcuentacon As String, valorcartera As String
       
            sAssert = "1) %iva "
            rs.Open "select * from porcentajesiva where iva = " & ComboCodigo(cboIva) & " order by fecha_baja", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                If IsNull(rs!fecha_baja) Then
                    porciva = rs!PORCENTAJE
                Else
                    porciva = 0
                End If
                rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            If optcontado = True Then

                sAssert = "2) _INGCOMPRAS "
                
              DataEnvironment1.dbo_INGCOMPRAS "A", dtFecha, val(txtanio), val(txtmes), _
                    UpROV.codigo, UpROV.DESCRIPCION, CUIT.Text, ComboCodigo(cboIva), TIPODOC_FAC_PROVEEDOR, NroDoc, _
                    uNumDoc.suc, val(txtcodcuenta), s2n(txtimporte) * z, s2n(txtneto) * z, s2n(txtper3337) * z, s2n(txtIva) * z, s2n(txtIva27) * z, s2n(txtreteniva) * z, _
                    s2n(txtIva10) * z, s2n(txtimpint) * z, s2n(txtretengan) * z, s2n(txtper3431) * z, s2n(txtexento) * z, s2n(txtIBcapital) * z, s2n(txtIBprovincia) * z, _
                      1, val(txtserie), porciva, ObtenerCodigo("Tipocompras", cmbtipocompra.Text), 0, ObtenerCodigo("Monedas", cmbMoneda.Text), _
                    z, Date, UsuarioSistema!codigo, iddoc, uRetCompras.retIB, uRetCompras.retgan, _
                    uNumDoc.letra, txtNroIIBB, uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo
                
                If txtNoGrabado <> "" Then
                    txtNoGrabado = n2r(s2n(txtNoGrabado))
                Else
                    txtNoGrabado = 0
                End If
                
                cYear = Right(Year(dtFecha), 2)
                cMonth = Format(Month(dtFecha), "00")
                cNrodoc = NroDoc
                cCodMixto = cYear & cMonth & cNrodoc
                cCodMixto = txtCodMixto
                'a = "update compras set codmixto=" & ssTexto(cCodMixto) & ",nogravado=" & x2s(txtNoGrabado * z) & _
                        ",direccion='" & Trim(txtdireccion) & "',localidad='" & Trim(txtlocalidad) & "',provincia='" & ObtenerCodigoS("Provincias", Trim(CmbProvincia.Text)) & "',pais='" & Trim(txtpais) & "' " & _
                        " where codpr=" & UpROV.codigo & " and tipodoc='" & TIPODOC_FAC_PROVEEDOR & "' and NroDoc=" & NroDoc
                        
                DataEnvironment1.Sistema.Execute "update compras set codmixto=" & ssTexto(cCodMixto) & ",nogravado=" & x2s(txtNoGrabado * z) & _
                        ",direccion='" & Trim(txtdireccion) & "',localidad='" & Trim(txtLocalidad) & "',provincia='" & ObtenerCodigoS("Provincias", Trim(CmbProvincia.Text)) & "',pais='" & Trim(txtpais) & "' " & _
                        " where codpr=" & UpROV.codigo & " and tipodoc='" & TIPODOC_FAC_PROVEEDOR & "' and NroDoc=" & NroDoc
                    ' s2n(txtRetIIBBPago), s2n(txtRetGanPago)
    
                'SI REALIZO UNA TRANSFERENCIA
                If s2n(txttransf) <> 0 Then
                    rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not IsNull(rs!maxcodigo) Then
                        maximobanc = rs!maxcodigo + 1
                    Else
                        maximobanc = 1
                    End If
                    rs.Close
                    Set rs = Nothing
                    
                    sAssert = "3) dbo_INGCOMPRAMOVIBANC - SP MOVIBANC"
                    '
                    DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", val(txtcodcuenta), "S", "Transf. " & "Prov. " & ObtenerDescripcion("Prov", UpROV.codigo) _
                        , dtFecha, "E", 0, s2n(txttransf), TIPODOC_FAC_PROVEEDOR, uNumDoc.num, maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                    
                End If
                            
                'SI PAGO CON CHEQUES PROPIOS
                If s2n(txtimpcheques) <> 0 Then
                    If ExistenPropios Then
                        
                        rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximobanc = rs!maxcodigo + 1
                        Else
                            maximobanc = 1
                        End If
                        rs.Close
                        Set rs = Nothing
                        
                        sAssert = "4) dbo_INGCOMPRAMOVIBANC "
                        
                        For x = 1 To uCheques.rows  ' FrmCheques.grillapropios.rows - 1
                            'INGCOMPRAMOVIBANC es igual al STORE del INGCHEQUEMOVIBANC
                            'If FrmCheques.grillapropios.TextMatrix(x, 5) <> "" Then
                            If uCheques.chPropio(x) Then
                                sAssert = "4) dbo_INGCOMPRAMOVIBANC - SP"
                                
                                
                                If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                                    If uCheques.chNroInt(x) = 0 Then
                                        numIntP = nuevoCodigo("chq_Comp")
                                        ' cargo por 1ra vez
                                        DataEnvironment1.dbo_INGRESOCHEQUERA numIntP, 0, uCheques.chNumero(x), uCheques.chBancCod(x), uCheques.chBancCod(x), _
                                                 0, 0, "", 0, "C", 0, 0, Date, UsuarioSistema!codigo, 0, 0, 1
                                        uCheques.chSetearNroInt x, numIntP
                                    End If
                                End If
                               
                                'mod lito 20/7/6  cuenta = la del cheque
                                cueche = s2n(obtenerDeSQL("select cuentabancaria from chq_comp where codigo = " & uCheques.chNroInt(x)))
                                DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", cueche, "L", "Fac. " & uNumDoc.num & " Prov. " & ObtenerDescripcion("Prov", UpROV.codigo) _
                                     , dtFecha, "P", uCheques.chNroInt(x), uCheques.chMonto(x), TIPODOC_FAC_PROVEEDOR, uNumDoc.num, maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                                'INCREMENTO EL AUTOMATICO DE MOVIBANC
                                maximobanc = maximobanc + 1
                            End If
                        Next

                    End If
                End If
                            
                'SI PAGO CON CHEQUES DE TERCEROS
                'mod lito 20/7/6  cuenta = 0, no debe aparecer como mov bancario
                If s2n(txtimpcheques) <> 0 Then
                    If ExistenTerceros Then
                        
                        rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximobanc = rs!maxcodigo + 1
                        Else
                            maximobanc = 1
                        End If
                        rs.Close
                        Set rs = Nothing

                        sAssert = "5) dbo_INGCOMPRAMOVIBANC "

                        For x = 1 To uCheques.rows
                            If Not uCheques.chPropio(x) Then
                                sAssert = "5a) dbo_INGCOMPRAMOVIBANC - SP"
                                'INGCOMPRAMOVIBANC es igual al STORE del INGCHEQUEMOVIBANC
                                DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", 0, "T", "Fac. " & uNumDoc.num & "Prov. " & ObtenerDescripcion("Prov", UpROV.codigo) _
                                    , dtFecha, "C", uCheques.chNroInt(x), uCheques.chMonto(x), TIPODOC_FAC_PROVEEDOR, uNumDoc.num, maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                                'INCREMENTO EL AUTOMATICO DE MOVIBANC
                                maximobanc = maximobanc + 1
                                '
                            End If
                        Next

                    End If
                End If
                                        
                
                
                'ACA EMPIEZA LAS ALTAS A MOVICAJA
                
                                           
                'SI PAGO EN EFECTIVO
                If s2n(txtefectivo) <> 0 Then
                
                    sAssert = "6) MoviCaja "
                
                    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not IsNull(rs!maxcodigo) Then
                        maximocaja = rs!maxcodigo + 1
                    Else
                        maximocaja = 1
                    End If
                    rs.Close
                    Set rs = Nothing
                    
                    
                    valorcuenta = verCuentaContableCaja(val(txtcodcaja))
                    
                    
                    sAssert = "6) MoviCaja - asiento"
                    'haber EFECTIVO
                    AsientoCompra.AcumularItem valorcuenta, 0, s2n(txtefectivo)  ', TextoAsientoComprobante

                    sAssert = "6) MoviCaja -sp ingcompraMC"
                    '
                    DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", val(txtcodcaja), maximocaja, "E", "E", s2n(txtefectivo), "Fac. " & uNumDoc.num & "Prov. " & UpROV.codigo _
                        , dtFecha, 0, UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, valorcuenta, 0 _
                        , iddoc, Date, UsuarioSistema!codigo, z
                End If
                            
                'SI REALIZO UNA TRANSFERENCIA
                If s2n(txttransf) <> 0 Then
                
                    sAssert = "7) MoviCaja "
                
                    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not IsNull(rs!maxcodigo) Then
                        maximocaja = rs!maxcodigo + 1
                    Else
                        maximocaja = 1
                    End If
                    rs.Close
                    Set rs = Nothing
                
                    rs.Open "select cuenta_con from Ctasbank where codigo = " & val(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                    If Not rs.EOF Then
                        valorcuentacon = rs!cuenta_con
                    Else
                        valorcuentacon = ""
                    End If
                    rs.Close
                    Set rs = Nothing
                                    
                    rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not IsNull(rs!maxcodigo) Then
                        maximobanc = rs!maxcodigo + 1
                    Else
                        maximobanc = 1
                    End If
                    rs.Close
                    Set rs = Nothing
                                    
                    DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "T", "E", s2n(txttransf), "Fac. " & uNumDoc.num & "Prov. " & UpROV.codigo, _
                        dtFecha, 0, UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, valorcuentacon, maximobanc, _
                        iddoc, Date, UsuarioSistema!codigo, z
                    
                    sAssert = "7b) MoviCaja - ASIENTO "
                    'haber  TRANSFERENCIA
                    AsientoCompra.AcumularItem obtenerDeSQL("select cuenta_con from ctasbank where activo = 1 and codigo = '" & x2s(s2n(txtcodcuenta)) & "' "), 0, s2n(txttransf)
                End If
                           
                           
                'SI PAGO CON CHEQUES PROPIOS
'''             If txtimpcheques <> "" And txtimpcheques <> "0" Then
                If s2n(txtimpcheques) <> 0 Then ' al pedo , con existenpropios alcanza
                    If ExistenPropios() Then
                                                                                
                        rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximocaja = rs!maxcodigo + 1
                        Else
                            maximocaja = 1
                        End If
                        rs.Close
                        Set rs = Nothing
                        
                        rs.Open "select cuenta_con from Ctasbank where codigo = " & val(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                        If Not rs.EOF Then
                            valorcuentacon = rs!cuenta_con
                        Else
                            valorcuentacon = ""
                        End If
                        rs.Close
                        Set rs = Nothing
                        
                       
                        
                        rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximobanc = rs!maxcodigo + 1
                        Else
                            maximobanc = 1
                        End If
                        rs.Close
                        Set rs = Nothing
                        
                        sAssert = "8) dbo_INGCOMPRACHEQUEPROPIO "
                        
                        For x = 1 To uCheques.rows
                            
                            If uCheques.chPropio(x) Then
                                sAssert = "8a) cheque propio - sp"
                                DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "P", "E", uCheques.chMonto(x), "Fac. " & uNumDoc.num & " Prov. " & UpROV.codigo _
                                    , dtFecha, uCheques.chNroInt(x), UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, valorcuentacon, maximobanc _
                                    , iddoc, Date, UsuarioSistema!codigo, z
                                
                                fechapropio = uCheques.chFecha(x)
                                DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "A", uCheques.chNroInt(x), fechapropio, uCheques.chMonto(x) _
                                    , uNumDoc.num, TIPODOC_FAC_PROVEEDOR, UpROV.codigo, "T", uCheques.chFecha(x), dtFecha, Date, UsuarioSistema!codigo, 0, 0, 1, z, ObtenerCodigo("Monedas", cmbMoneda.Text)
                                
                                'INCREMENTO EL AUTOMATICO DE MOVIBANC
                                maximobanc = maximobanc + 1
                                
                                sAssert = "8b) cheque propio - asiento"
                                'haber CHEQUE PROPIO
                                
                                'AsientoCompra.AcumularItem sSinNull(obtenerDeSQL("select cuenta_con from ctasBank where  codigo = '" & uCheques.chBancCod(x) & "' and activo = 1")), 0, uCheques.chMonto(x)
                                AsientoCompra.AcumularItem uCheques.chCuenta(x), 0, uCheques.chMonto(x)
                                
                            End If
                        Next
                    
                    End If
                End If
                
                           
                'SI PAGO CON CHEQUES TERCEROS
'''             If txtimpcheques <> "" And txtimpcheques <> "0" Then
                If s2n(txtimpcheques) <> 0 Then
                    If ExistenTerceros() Then
                                                        
                        rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximocaja = rs!maxcodigo + 1
                        Else
                            maximocaja = 1
                        End If
                        rs.Close
                        Set rs = Nothing
                        
'                        rs.Open "select valores_cartera from Imputaciones", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                        If Not rs.EOF Then
'                            valcartera = rs!valores_cartera
'                        Else
'                            valcartera = ""
'                        End If
'                        rs.Close
'                        Set rs = Nothing
                        
                        rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not IsNull(rs!maxcodigo) Then
                            maximobanc = rs!maxcodigo + 1
                        Else
                            maximobanc = 1
                        End If
                        rs.Close
                        Set rs = Nothing
                        
                        sAssert = "9) dbo_INGCOMPRACHEQUETERCEROS "
                        For x = 1 To uCheques.rows
                            
                            If Not uCheques.chPropio(x) Then
                                sAssert = "9a) dbo_INGCOMPRACHEQUETERCEROS - SP "
                                DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "C", "E", uCheques.chMonto(x), "Fac. " & uNumDoc.num & " Prov. " & UpROV.codigo _
                                    , dtFecha, uCheques.chNroInt(x), UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, uCheques.chCuenta(x), maximobanc _
                                    , iddoc, Date, UsuarioSistema!codigo, z
                                
                                DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "A", uCheques.chNroInt(x), 0, "", UpROV.codigo, uNumDoc.num, 0 _
                                    , dtFecha, "T", "FDC", Date, UsuarioSistema!codigo, 0, 0, 1, ObtenerCodigo("Monedas", cmbMoneda.Text), z
                                
                                'INCREMENTO EL AUTOMATICO DE MOVIBANC
                                maximobanc = maximobanc + 1
                                
                                sAssert = "9b) dbo_INGCOMPRAMOVIBANC - Asiento"
                                'Haber Cheques 3ros
                                AsientoCompra.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, uCheques.chMonto(x)
                            End If
                        Next
                    
                    End If
                End If
                
                'retenciones pago (ret 3ros)
                AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_GAN_3ros), 0, uRetCompras.retgan  ' uRetCompras.CuentaGan
                AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_IB_3ros), 0, uRetCompras.retIB
                
            Else            ' *********  alta FC CUENTA CORRIENTE *******
            
            
                sAssert = "10a) dbo_INGCOMPRASCTACTE -ASIENTO "
                'HABER Deuda_a_Proveedores
                tiene_c = obtenerDeSQL("select tiene_cuenta from prov where codigo = " & UpROV.codigo)
                If tiene_c = 1 Then
                    AsientoCompra.AgregarItem obtenerDeSQL("select cuenta from prov where codigo = " & UpROV.codigo), 0, s2n(txtimporte) * z, TextoAsientoComprobante
                Else
                    AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(txtimporte) * z, TextoAsientoComprobante
                End If
                
                
                sAssert = "10b) dbo_INGCOMPRASCTACTE - SP CTACTE "
                
                DataEnvironment1.dbo_INGCOMPRASCTACTE "A", dtFecha, val(txtanio), val(txtmes) _
                    , UpROV.codigo, UpROV.DESCRIPCION, CUIT.Text, ComboCodigo(cboIva), TIPODOC_FAC_PROVEEDOR, uNumDoc.num _
                    , uNumDoc.suc, 0, s2n(txtimporte) * z, s2n(txtneto) * z, s2n(txtimporte) * z, txtfvto, s2n(txtIva) * z, s2n(txtper3337) * z, s2n(txtIva27) * z, s2n(txtreteniva) * z _
                    , s2n(txtIva10) * z, s2n(txtimpint) * z, s2n(txtretengan) * z, s2n(txtper3431) * z, s2n(txtexento) * z, s2n(txtIBcapital) * z, s2n(txtIBprovincia) * z, _
                    val(txtserie), porciva, ObtenerCodigo("Tipocompras", cmbtipocompra.Text), ObtenerCodigo("Formaspago", cmbformapago.Text), 0, ObtenerCodigo("Monedas", cmbMoneda.Text), _
                    z, 0, Date, UsuarioSistema!codigo, iddoc, 0, 0, _
                    uNumDoc.letra, txtNroIIBB, uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo
                    
                If txtNoGrabado <> "" Then
                    txtNoGrabado = s2n(txtNoGrabado)
                Else
                    txtNoGrabado = 0
                End If
                'resu = "update transcom set nogravado=" & x2s(txtNoGrabado) * z & " where codpr=" & UpROV.codigo & " and tipodoc='" & TIPODOC_FAC_PROVEEDOR & "' and NroDoc=" & NroDoc
                cCodMixto = txtCodMixto
                DataEnvironment1.Sistema.Execute "update transcom set codmixto=" & ssTexto(txtCodMixto) & ",nogravado=" & x2s(txtNoGrabado * z) & _
                        ",direccion='" & Trim(txtdireccion) & "',localidad='" & Trim(txtLocalidad) & "',provincia='" & ObtenerCodigoS("Provincias", Trim(CmbProvincia.Text)) & "',pais='" & Trim(txtpais) & "' " & _
                        " where codpr=" & UpROV.codigo & " and tipodoc='" & TIPODOC_FAC_PROVEEDOR & "' and NroDoc=" & NroDoc
                    ' 0,0   o s2n(txtRetGanPago), s2n(txtRetIIBBPago) ?
            End If
        End If
       
       If gEMPR_ConSistContable Then

            sAssert = "12) dbo_INGCENTROCOSTOS "
            
            If FrmCostosYContable.grillacostos.rows > 1 And FrmCostosYContable.grillacostos.TextMatrix(1, 1) > "" Then
            
                'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                For x = 1 To FrmCostosYContable.grillacostos.rows - 1
                    DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(FrmCostosYContable.grillacostos.TextMatrix(x, 0)), _
                    dtFecha, "FAC", val(uNumDoc.num), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) * z, s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3), 3) * z, (s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) + s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3))) * z, Date, 0, UsuarioSistema!codigo, 0, 1, "", FrmCostosYContable.grillacostos.TextMatrix(x, 4), UpROV.codigo
                    'txtfact
                Next
            End If
        End If
        
         sAssert = "13) xRemitos "
        
        '------- POR REMITO ---------  ver de meter la rutina y/o el loop dentro del modulo odel frm
        Dim ii As Long
        If mPorRemito Then
            With frmRemitosCompraPendientes
                ii = 1
                While .item(ii) > 0
                    ItemFacturaRemitoCompra ItemFRC_Alta, .item(ii), UpROV.codigo, TIPODOC_FAC_PROVEEDOR, uNumDoc.num, .cant(ii), .PrecioU(ii) * z
                    ii = ii + 1
                Wend
            End With
        End If
        '------- POR REMITO ---------
        
        
      
        sAssert = "15) ASIENTOS"
'        If AsientoCompra.Grabar(iddoc) = 0 Then
        If siAsiento("AsientosCompras") Then AsientoCompra.Grabar iddoc
'            DE_RollbackTrans
'            ufa "Err al grabar asiento ", Me.Name & " - " & sAssert
'            Exit Sub
'        End If
        DetalleIIBB "A", TIPODOC_FAC_PROVEEDOR, uNumDoc.num, UpROV.codigo, iddoc, z

        DE_CommitTrans
        '*************************************************

        'MsgBox "Operación Realizada con éxito", vbOKOnly
        MsgBox "FACTURA COMPRAS Realizada con éxito", vbOKOnly
        ImprimirFacProv
        cmbingresar.enabled = False
        LimpioControles
        HabilitoControles (False)
        Call Habilitobotones(True, True, False, False, True, False)
'        cmbingresar.Enabled = False
        
        uCheques.Borrar
        FrmCostosYContable.LimpioControles
'        FrmCostosYContable.InicioGrilla
        FrmCostosYContable.InicioGrillaCostos
        Unload FrmCostosYContable
        
    End If
End If


FinOK:
    Exit Sub
UfaOK:
    midDoc = 0
    DE_RollbackTrans
    uCheques.resetNroIntPropios
        
    
    ufa "Err al grabar factura", Me.Name & " - " & sAssert
    Resume FinOK
End Sub

Private Sub ImprimirFacProv()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAimprimir
    
    Dim stblChequesOPtmp As String
    Dim r As Long
    Dim str, sql, donde, direccion, Localidad As String, str2 As String
    Dim rs As New ADODB.Recordset
    
    If optctacte Then Exit Sub
    donde = "Orden Pago principal"
    
    sql = "Select * from movicaja where tipo = 'T' and iddoc = " & val(midDoc)
    rs.Open sql, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then RptOrdenPagoFacContProv.lbltransf = Format(rs!Importe, "#,##0.00")
    Set rs = Nothing
    
    stblChequesOPtmp = TablaTempCrear(tt_ChequeOPtmp)
    With uCheques
        For r = 1 To .rows
            sql = "insert into " & stblChequesOPtmp _
            & " (nroint, banco, cheque, importe, fecha, propio) values( " _
            & .chNroInt(r) & ", '" & .chBancDes(r) & "', '" & .chNumero(r) & "', " & x2s(.chMonto(r)) & ", " & ssFecha(.chFecha(r)) & ", '" & IIf(.chPropio(r), "P", "T") & "')"
            DataEnvironment1.Sistema.Execute sql
        Next r
    End With
    
    
    str = "select * from " & stblChequesOPtmp
    RptOrdenPagoFacContProv.data1.Connection = DataEnvironment1.Sistema
    RptOrdenPagoFacContProv.data1.Source = str
    RptOrdenPagoFacContProv.lblcheques = Format(uCheques.Total, "#,##0.00")
    RptOrdenPagoFacContProv.lblfecha = dtFecha
    RptOrdenPagoFacContProv.lblTitulo = "ORDEN DE PAGO Nº " & Format(VerNroPago(midDoc), "0001 - 00000000")
    RptOrdenPagoFacContProv.TxtFactura = uNumDoc.txtNumero 'Format(txtfact, "00000000")
    RptOrdenPagoFacContProv.lblproveedor = "A la orden de " & UpROV.DESCRIPCION
    RptOrdenPagoFacContProv.lblvalor = "Por la cantidad de pesos " & NroEnLetras(s2n(txtimporte))
    RptOrdenPagoFacContProv.lblefectivo = Format(txtefectivo, "#,##0.00")
    RptOrdenPagoFacContProv.LblRetGanancia = Format(uRetCompras.retgan, "#,##0.00")
    RptOrdenPagoFacContProv.LblretIB = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoFacContProv.lbltotal = Format(txtimporte, "#,#00.00")
    
    donde = "Orden de Pago Constancia RET. IMPUESTO GANANCIA"
        direccion = obtenerDeSQL("select direccion from prov where codigo = " & UpROV.codigo & " ")
        Localidad = obtenerDeSQL("select localidad from prov where codigo = " & UpROV.codigo & " ")

    
    
    sql = "select tipodoc,nrodoc,total as saldo from compras where iddoc =  " & midDoc
    
    If uRetCompras.retgan > 0 Then

        RptOrdenPagoConstRet_IG.DataImp_Ganancia.Connection = DataEnvironment1.Sistema
        RptOrdenPagoConstRet_IG.DataImp_Ganancia.Source = sql '--------
        RptOrdenPagoConstRet_IG.lblfecha = dtFecha
        RptOrdenPagoConstRet_IG.LblRegimen_IG = uRetCompras.IG_Tipo
        RptOrdenPagoConstRet_IG.txtProveedor = UpROV.DESCRIPCION
        RptOrdenPagoConstRet_IG.TxtDomicilioProv = direccion & "    " & Localidad
        RptOrdenPagoConstRet_IG.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
        RptOrdenPagoConstRet_IG.RG_PagosTotalMes = Format(txtimporte, "#,##0.00")
        RptOrdenPagoConstRet_IG.retgan = Format(uRetCompras.retgan, "#,##0.00")
        RptOrdenPagoConstRet_IG.retganEnPesos = enletras(uRetCompras.retgan)
        RptOrdenPagoConstRet_IG.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
        RptOrdenPagoConstRet_IG.Txtop = VerNroPago(midDoc)
        
        RptOrdenPagoConsRet_IG_calculo.lblfecha = dtFecha
        RptOrdenPagoConsRet_IG_calculo.txtProveedor = UpROV.DESCRIPCION
        RptOrdenPagoConsRet_IG_calculo.TxtDomicilioProv = direccion & "    " & Localidad
        RptOrdenPagoConsRet_IG_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
        RptOrdenPagoConsRet_IG_calculo.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.RG_MinimoNoImponible = uRetCompras.RG_MinimoNoImponible
        RptOrdenPagoConsRet_IG_calculo.RG_TxtFormula = uRetCompras.RG_TxtFormula
        RptOrdenPagoConsRet_IG_calculo.retgan = Format(uRetCompras.retgan, "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.RG_PagosAnterioresMes = Format(uRetCompras.RG_PagosAnterioresMes, "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.RG_PagosRetAnteriores = Format(uRetCompras.RG_PagosRetAnteriores, "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
        RptOrdenPagoConsRet_IG_calculo.LblRetGanPesos = enletras(uRetCompras.retgan)
        RptOrdenPagoConsRet_IG_calculo.Pago_Fecha = Format(Abs(CDbl(uRetCompras.RG_PagosAnterioresMes) - CDbl(uRetCompras.RG_PagosTotalMes)), "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.Total_Imponible = Format(CDbl(uRetCompras.RG_PagosRetAnteriores) + CDbl(uRetCompras.retgan), "#,##0.00")
        RptOrdenPagoConsRet_IG_calculo.Printer.Copies = 1
        RptOrdenPagoConstRet_IG.Printer.Copies = 2
        RptOrdenPagoConstRet_IG.Restart
        RptOrdenPagoConsRet_IG_calculo.Restart
        If PREVIEW_IMPRESIONES Then
            RptOrdenPagoConstRet_IG.Show
            RptOrdenPagoConsRet_IG_calculo.Show
        Else
            RptOrdenPagoConstRet_IG.PrintReport False
            RptOrdenPagoConsRet_IG_calculo.PrintReport False
        End If
    End If
    
    donde = "Orden de Pago Constancia RET. INGRESO BRUTOS"
   
   If uRetCompras.retIB > 0 Then
   
        RptOrdenPagoConstRet_IB.DataImp_IB.Connection = DataEnvironment1.Sistema
        RptOrdenPagoConstRet_IB.DataImp_IB.Source = sql '-----------------
        RptOrdenPagoConstRet_IB.lblfecha = dtFecha
        RptOrdenPagoConstRet_IB.LblRegimen_IIBB = uRetCompras.IB_Tipo
        RptOrdenPagoConstRet_IB.txtProveedor = UpROV.DESCRIPCION
        RptOrdenPagoConstRet_IB.TxtDomicilioProv = direccion & "    " & Localidad
        RptOrdenPagoConstRet_IB.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
        RptOrdenPagoConstRet_IB.txtNroIIBB = obtenerDeSQL("select numiibb from prov where codigo = " & UpROV.codigo & " ")
        RptOrdenPagoConstRet_IB.RG_PagosTotalMes = Format(txtimporte, "#,##0.00")
        RptOrdenPagoConstRet_IB.retgan = Format(uRetCompras.retIB, "#,##0.00")
        RptOrdenPagoConstRet_IB.retganEnPesos = enletras(uRetCompras.retIB)
        RptOrdenPagoConstRet_IB.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
        RptOrdenPagoConstRet_IB.Txtop = VerNroPago(midDoc)
        
        RptOrdenPagoConsRet_IB_calculo.lblfecha = dtFecha
        RptOrdenPagoConsRet_IB_calculo.txtProveedor = UpROV.DESCRIPCION
        RptOrdenPagoConsRet_IB_calculo.TxtDomicilioProv = direccion & "    " & Localidad
        RptOrdenPagoConsRet_IB_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
        RptOrdenPagoConsRet_IB_calculo.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
        RptOrdenPagoConsRet_IB_calculo.IB_TxtFormula = uRetCompras.IB_TxtFormula
        RptOrdenPagoConsRet_IB_calculo.retIB = Format(uRetCompras.retIB, "#,##0.00")
        RptOrdenPagoConsRet_IB_calculo.IB_base = Format(uRetCompras.IB_base, "#,##0.00")
        RptOrdenPagoConsRet_IB_calculo.retIB1 = Format(uRetCompras.retIB, "#,##0.00")
        RptOrdenPagoConsRet_IB_calculo.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
        RptOrdenPagoConsRet_IB_calculo.LblRetIIBBPesos = enletras(uRetCompras.retIB)
        RptOrdenPagoConsRet_IB_calculo.Printer.Copies = 1
        RptOrdenPagoConstRet_IB.Printer.Copies = 2
        RptOrdenPagoConstRet_IB.Restart
        RptOrdenPagoConsRet_IB_calculo.Restart
        
        
        If PREVIEW_IMPRESIONES Then
            RptOrdenPagoConstRet_IB.Show
            RptOrdenPagoConsRet_IB_calculo.Show
        Else
            RptOrdenPagoConstRet_IB.PrintReport False
            RptOrdenPagoConsRet_IB_calculo.PrintReport False
        End If
    End If
    
    RptOrdenPagoFacContProv.Printer.Copies = 2
    RptOrdenPagoFacContProv.Restart

    If PREVIEW_IMPRESIONES Then
       RptOrdenPagoFacContProv.Show
    Else
       RptOrdenPagoFacContProv.PrintReport False
    End If
 
 
FinOK:
    Exit Sub
UFAimprimir:
    ufa "Factura Prov grabado, fallo la impresion " & donde, Me.Name & " - " & donde
    Resume FinOK
End Sub


Private Sub cmdImprimir_Click()
ImprimirFacProv
End Sub
Private Function ExistenTerceros() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If Not uCheques.chPropio(x) Then
            ExistenTerceros = True
            Exit Function
        End If
    Next x
End Function
Private Function ExistenPropios() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If uCheques.chPropio(x) = True Then
            ExistenPropios = True
            Exit Function
        End If
    Next x
End Function

Private Sub cmdQuitarProv_Click()
    With gIIBBProvincia
        If .rows > 0 Then
            If .Row >= 0 Then
                .RemoveItem .Row
            End If
        End If
    End With
calGrillaProv
End Sub

Private Sub cmdSalir_Click()
    If confirma(" cerrar formulario?") Then Unload Me
End Sub

Sub LimpioControles()
    UpROV.codigo = 0
    'FrmBorrarTxt Me
    uNumDoc.clear
    lblIDDOC = ""
    txtcodcaja = 1
    midDoc = 0
    'txtfact = ""
    'txtnombre = ""
    CUIT.Text = ""
    dtFecha = Date
'    txttipo = ""
    cmbformapago.ListIndex = -1
    txtfvto = Date
'    optCheques = False
'    optcontado = False
'    optctacte = True
    txtneto = "0"
    txtIva = "0"
    txtper3337 = "0"
    txtimporte = "0"
    txtIva27 = "0"
    txtexento = "0"
    txtreteniva = "0"
    txtimpint = "0"
    txtretengan = "0"
    txtper3431 = "0"
    txtIva10 = "0"
    txtIBcapital = "0"
    txtIBprovincia = "0"
    txtimpcheques = "0"
    txtefectivo = "0"
    txtNoGrabado = "0"
    
    'txtanio = ""
    txtmes.ListIndex = -1
    'txtsuc = ""
    'txtSerie = ""
    txtcotiz = ""
    cmbtipocompra.ListIndex = -1
    cmbMoneda.ListIndex = -1
    'txttransf = "0"
    'txtcodcaja = ""
    'txtcodcuenta = ""
    'txtcaja = ""
    'txtcuenta = ""
    
    'txtsaldo = ""
    cargar = ""
    Ope = ""
     uCheques.Borrar
     uTipoCompra.Borrar
     iniGrillaProv
     TabDetalle.Tab = 0
     
     txtdireccion = ""
     txtLocalidad = ""
     txtpais = ""
     CmbProvincia.ListIndex = 1
End Sub

Private Function iniGrillaProv()
DetalleIIBB "I"
End Function

Private Function calGrillaProv()
Dim iibbSuma As Double, i As Long
iibbSuma = 0
    With gIIBBProvincia
        If .rows > 1 Then
            For i = 1 To .rows - 1
                .TextMatrix(i, 3) = s2n(.TextMatrix(i, 3), 2, True)
                iibbSuma = iibbSuma + s2n(.TextMatrix(i, 3))
            Next
        Else
        End If
    End With
txtIBprovincia = iibbSuma
End Function

Private Sub cmbtipocompra_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Sub HabilitoControles(habilito As Boolean)
    uCheques.enabled = habilito
    uRetCompras.enabled = habilito

    UpROV.enabled = habilito And Not mPorRemito
'    txtcodprov.Enabled = habilito
    
    'txtfact.enabled = habilito
'    txtSuc.enabled = habilito
    uNumDoc.enabled = habilito
    cboIva.enabled = habilito
    
'    txtnombre.Enabled = habilito
    CUIT.enabled = habilito
    dtFecha.enabled = habilito
    cmbformapago.enabled = habilito
    txtfvto.enabled = habilito
'    optcontado.Enabled = habilito
'    optctacte.Enabled = habilito
    
    txtanio.enabled = habilito
    txtmes.enabled = habilito

    txtserie.enabled = habilito
    txtcotiz.enabled = habilito
    cmbtipocompra.enabled = habilito
    cmbMoneda.enabled = habilito
    cmdAgregarProv.enabled = habilito
    
'    txtplan.Enabled = habilito
    
    txtneto.enabled = habilito
    txtIva.enabled = habilito
    txtper3337.enabled = habilito
    txtimporte.enabled = habilito
    txtIva27.enabled = habilito
    txtexento.enabled = habilito
    txtreteniva.enabled = habilito
    txtimpint.enabled = habilito
    txtretengan.enabled = habilito
    txtper3431.enabled = habilito
    txtIva10.enabled = habilito
    txtNoGrabado.enabled = habilito
    txtIBprovincia.enabled = habilito
    txtIBcapital.enabled = habilito
'    cmdprov.Enabled = habilito
    
    txtefectivo.enabled = habilito
    txtcodcaja.enabled = habilito
    cmbcaja.enabled = habilito
    txtimpcheques.enabled = habilito: uCheques.enabled = habilito
    txttransf.enabled = habilito
    cmbcuenta.enabled = habilito
    
    txtdireccion.enabled = habilito
    txtLocalidad.enabled = habilito
    txtpais.enabled = habilito
    CmbProvincia.enabled = habilito
End Sub
Sub HabilitoControles2(habilito As Boolean, habi As Boolean)
    
    txtneto.enabled = habilito
    txtIva.enabled = habilito
    txtper3337.enabled = habilito
    txtimporte.enabled = habi
    txtIva27.enabled = habilito
    txtexento.enabled = habi
    txtreteniva.enabled = habilito
    txtimpint.enabled = habilito
    txtretengan.enabled = habilito
    txtper3431.enabled = habilito
    txtIva10.enabled = habilito
    txtNoGrabado.enabled = habilito
    txtIBprovincia.enabled = habilito
    txtIBcapital.enabled = habilito

End Sub
Sub FormadePago(habilito As Boolean)
    fraContado.enabled = habilito
    txtefectivo.enabled = habilito
    txtimpcheques.enabled = habilito: uCheques.enabled = habilito
    txttransf.enabled = habilito
    txtcodcaja.enabled = habilito
    cmbcaja.enabled = habilito
End Sub


Public Sub CargarDatos()

    Dim rs As New ADODB.Recordset, codigo As String

    If rsmov.State = 1 Then
        rsmov.Close
        Set rsmov = Nothing
    End If
    
    codigo = Trim(Me.Tag)
End Sub

Sub CargoRegistro()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCarga
    Dim z As Double
    z = s2n(rsmov!cotizacion, 4)
    If z = 0 Then z = 1
    lblIDDOC = nSinNull(rsmov!iddoc)
    UpROV.codigo = rsmov!CODPR
    'txtprov = ObtenerDescripcion("Prov", rsmov!codpr)
    UpROV.DESCRIPCION = rsmov!razonsocialprov
    
    'txtfact = rsmov!NroDoc
    uNumDoc.num = rsmov!NroDoc
    
'    txtnombre = rsmov!razonsocialprov
    'cuit.Numero = rsmov!cuitprov
    CUIT.Text = rsmov!cuitprov
    cmbtipocompra = BuscoDato("TipoCompras", rsmov!Tipocompra)
    dtFecha = rsmov!Fecha
    
    '****************************************************************
    If Not IsNull(rsmov!direccion) Then
        txtdireccion = sSinNull(rsmov!direccion)
    Else
        txtdireccion = obtenerDeSQL("select direccion from prov where codigo=" & rsmov!CODPR)
    End If
    If Not IsNull(rsmov!Localidad) Then
        txtLocalidad = sSinNull(rsmov!Localidad)
    Else
        txtLocalidad = obtenerDeSQL("select localidad from prov where codigo=" & rsmov!CODPR)
    End If
    If Not IsNull(rsmov!Provincia) Then
        CmbProvincia.ListIndex = BuscarenComboS(CmbProvincia, ObtenerDescripcionS("provincias", sSinNull(rsmov!Provincia)))
    Else
        CmbProvincia.ListIndex = -1
    End If
    If Not IsNull(rsmov!Pais) Then
        txtpais = sSinNull(rsmov!Pais)
    Else
        txtpais = obtenerDeSQL("select pais from prov where codigo=" & rsmov!CODPR)
    End If
    '****************************************************************
    
    txtcotiz = s2n(rsmov!cotizacion, 4)
    
    txtneto = s2n(rsmov!Neto / z)
    txtIva = s2n(rsmov!IVA_21 / z)
    txtper3337 = s2n(rsmov!percepc / z)
    txtimporte = s2n(rsmov!Total / z)
    txtIva27 = s2n(rsmov!IVA_27 / z)
    txtexento = s2n(rsmov!EXENTO / z)
    txtreteniva = s2n(rsmov!iva_9 / z)
    txtimpint = s2n(rsmov!imp_int / z)
    txtretengan = s2n(rsmov!retgan / z)
    txtper3431 = s2n(rsmov!der_est / z)
    txtIva10 = s2n(rsmov!iva_10 / z)
    txtIBcapital = s2n(rsmov!ibcapital / z)
    txtIBprovincia = s2n(rsmov!ibprovincia / z)
    midDoc = nSinNull(rsmov!iddoc)
    txtNoGrabado = s2n(rsmov!nogravado / z)
    
    txtCodMixto = sSinNull(rsmov!codmixto)
    txtanio = rsmov!anoimp
    txtmes = rsmov!mesimp
    'txtSuc = rsmov!suc
    uNumDoc.suc = rsmov!suc
    
    txtserie = rsmov!Serie
    txtcotiz = rsmov!cotizacion
    cmbMoneda = ObtenerDescripcion("Monedas", rsmov!moneda)
    
    uRetCompras.retgan = nSinNull(rsmov!retganpago)
    uRetCompras.retIB = nSinNull(rsmov!IBPAGO)
    
    If optcontado = True Then
        txttransf = s2n(obtenerDeSQL("select importe from movibanc where operacion='S' and iddoc=" & midDoc))
        txtcodcuenta = s2n(obtenerDeSQL("select cuenta from movibanc where OPERACION='S' and iddoc=" & midDoc))
        txtcuenta = ObtenerDescripcionCuentas("ctasbank", s2n(txtcodcuenta))
        
        'txttransf = ObtenerTransferencia("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        txtcodcaja = ObtenerCaja("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        'txtcodcuenta = ObtenerCuenta("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
        'txtcuenta = ObtenerDescripcionCuentas("Cuentas", Val(txtcodcuenta))
        txtefectivo = ObtenerImporte("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        txtimpcheques = ObtenerTotalCheques("Movicaja", rsmov!NroDoc, rsmov!CODPR)
                
        CargoCheques
    Else
        cmbformapago = BuscoDato("FormasPago", rsmov!FormadePago)
        txtfvto = rsmov!vencim
        txtsaldo = rsmov!saldo
    End If
    
fin:
    Exit Sub
UfaCarga:
    ufa "err cargando registro", ""
    Resume fin
End Sub

Private Sub CargoCheques()
    Dim rs As New ADODB.Recordset
    Dim i As Long
     
    uCheques.Borrar
    i = 1
    rs.Open "select Movicaja.*, Chq_comp.banco from Movicaja inner join Chq_comp on Movicaja.interno = Chq_comp.codigo where Movicaja.codprov = " & UpROV.codigo & " and Movicaja.nrodoc = " & uNumDoc.num & " and Movicaja.tipodoc = 'FAC' and Movicaja.tipo = 'P' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        uCheques.metoCheque i, rs!interno, "P"
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    
    rs.Open "select Movicaja.*, Cheques.banco_nro from Movicaja inner join Cheques on Movicaja.interno = Cheques.nroint where Movicaja.codprov = " & UpROV.codigo & " and Movicaja.nrodoc = " & uNumDoc.num & " and Movicaja.tipodoc = 'FAC' and Movicaja.tipo = 'C'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        uCheques.metoCheque i, rs!interno, "T"
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cuit_GotFocus()
    GotFocusPinto CUIT
End Sub

Private Sub cuit_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub


Private Sub dtFecha_Change()
txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

Private Sub dtfecha_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Sub dtFecha_LostFocus()
'    txtmes = Month(dtfecha)
'    txtanio = Year(dtfecha)
txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

Private Sub Form_Initialize()
    mPorRemito = False
End Sub

Private Sub Form_Load()
    
    comboSql cmbformapago, "select Descripcion, codigo from FormasPago where activo = 1 order by dias"
    comboSql cboIva, "select descripcion, codigo from ivas order by codigo"
    CargaCombo2 cmbtipocompra, "TipoCompras", "descripcion", "codigo", ""
    CargaCombo CmbProvincia, "Provincias", "descripcion", "codigo", ""
    
    CargaCombo3 cmbMoneda, "Monedas", "descripcion", "codigo", ""
    fraMesImputacion.Visible = VerParametro(BS_ComprasConMesImputacion)

    UpROV.ini "select   descripcion   from prov where activo = 1 and codigo = '###'", "select  codigo as [ Codigo         ], cuit as [ CUIT            ], descripcion [ Descripcion                                                             ]  from prov where categ<>2 and activo = 1 order by codigo ", False
    TabDetalle.Tab = 0

'    uRetCompras.Enabled = False
    cmbingresar.Visible = gEMPR_ConSistContable
    revisoCdoCtaCte
    HabilitoControles False
    EsBusqueda = False
    iniGrillaProv
    Set pAsiento = New Collection
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Sub Habilitobotones(buscar As Boolean, agregar As Boolean, eliminar As Boolean, _
                    aceptar As Boolean, Cancelar As Boolean, Imprimir As Boolean)
    cmdbuscar.enabled = buscar
    cmdcancelar.enabled = Cancelar
    cmdeliminar.enabled = eliminar
    cmdnuevo.enabled = agregar
    cmdok.enabled = aceptar
    cmdImprimir.enabled = Imprimir
End Sub

Function ObtenerCuenta(tabla As String, nDoc As Long, prov As Long) As Long
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset

    Dim sqlstrCC As String, Cuenta As Long
    
    sqlstrCC = "Select movbanco from " + tabla + " where nrodoc = " & nDoc & " and codprov = " & prov & " and tipo = 'T' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        Cuenta = rs!MovBanco
        sqlstrCC = "Select cuenta from Movibanc where movbanco = " & Cuenta & " and activo=1"
        rs1.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            ObtenerCuenta = rs1!Cuenta
        Else
            ObtenerCuenta = 0
        End If
        rs1.Close
        Set rs1 = Nothing
    Else
        ObtenerCuenta = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function ObtenerImporte(tabla As String, nDoc As Long, prov As Long) As Double
    Dim rs As New ADODB.Recordset

    Dim sqlstrCC As String
    
    sqlstrCC = "Select importe from " + tabla + " where nrodoc = " & nDoc & " and codprov = " & prov & "  and tipo = 'E' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerImporte = rs!Importe
    Else
        ObtenerImporte = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function


Function ObtenerCaja(tabla As String, nDoc As Long, prov As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim sqlstrCC As String
    
    sqlstrCC = "Select caja from " + tabla + " where nrodoc = " & nDoc & " and codprov = " & prov & " and tipo = 'E' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerCaja = rs!caja
    Else
        ObtenerCaja = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function ObtenerTotalCheques(tabla As String, nDoc As Long, prov As Long) As Double
    Dim rs As New ADODB.Recordset
    
    Dim Total As Double
    Dim sqlstrCC As String
    
    sqlstrCC = "Select * from " + tabla + " where codprov = " & prov & " and nrodoc = " & nDoc & " and (tipo = 'P' or tipo = 'C') and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        Total = Total + rs!Importe
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    ObtenerTotalCheques = Total
    
End Function

Function ObtenerTransferencia(tabla As String, nDoc As Long, prov As Long) As Double
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String
    
    sqlstrCC = "Select importe from " + tabla + " where nrodoc = " & nDoc & " and codprov = " & prov & "  and tipo = 'T' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerTransferencia = rs!Importe
    Else
        ObtenerTransferencia = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Private Sub Form_Unload(cancel As Integer)
    mPorRemito = False
End Sub

Private Sub gIIBBProvincia_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If Col <> 3 Then cancel = True
End Sub

Private Sub gIIBBProvincia_CellChanged(ByVal Row As Long, ByVal Col As Long)
    calGrillaProv
End Sub

Private Sub lblIdDoc_Click()
    frmAsientoManual.mostrar s2n(lblIDDOC)
End Sub

Private Sub optcontado_Click()
    revisoCdoCtaCte
'    CambioOptPago
End Sub
Private Sub optcontado_Validate(cancel As Boolean)
    revisoCdoCtaCte
End Sub
Private Sub optctacte_Click()
    'CambioOptPago
    revisoCdoCtaCte
End Sub
Private Sub optctacte_Validate(cancel As Boolean)
    revisoCdoCtaCte
End Sub

Private Sub revisoCdoCtaCte()
    TabDetalle.TabEnabled(1) = optcontado.Value
    CambioOptPago
End Sub

Private Function RevisoCeldas()
    txtneto = s2n(txtneto)
    txtIva = s2n(txtIva)
    txtper3337 = s2n(txtper3337)
    txtimporte = s2n(txtimporte)
    txtIva27 = s2n(txtIva27)
    txtexento = s2n(txtexento)
    txtreteniva = s2n(txtreteniva)
    txtimpint = s2n(txtimpint)
    txtretengan = s2n(txtretengan)
    txtper3431 = s2n(txtper3431)
    txtIva10 = s2n(txtIva10)
    txtNoGrabado = s2n(txtNoGrabado)
    txtIBprovincia = s2n(txtIBprovincia)
    txtIBcapital = s2n(txtIBcapital)
End Function

Private Sub TabDetalle_Click(PreviousTab As Integer)
    Dim x As String, tiene_c As Long
    Dim auxIDDOC As Long, auxIDAsiento As Long, i As Long
    Dim rsAsiento As New ADODB.Recordset
    Dim z As Double
    
    RevisoCeldas
    
    If TabDetalle.Tab = 1 Then
        uRetCompras.Calcular UpROV.codigo, s2n(txtneto) + s2n(txtexento), s2n(txtneto) + s2n(txtexento), dtFecha
        txtTotalRetPago = uRetCompras.TotalRet
    ElseIf TabDetalle.Tab = 2 Then
        If EsBusqueda Then
            auxIDDOC = s2n(lblIDDOC)
            auxIDAsiento = obtenerDeSQL("select idasiento from asientos where iddoc=" & auxIDDOC)
            rsAsiento.Open "select * from mayor where debe>0 and idasiento= " & auxIDAsiento & " order by idmayor desc", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            With rsAsiento
                If .EOF And .BOF Then
                Else
                    .MoveFirst
                    z = s2n(txtcotiz, 4)
                    If z = 0 Then z = 1
                    For i = 0 To .RecordCount - 1
                       uTipoCompra.agregar !Cuenta, !Debe / z, True
                       .MoveNext
                    Next
                End If
            End With
        Else
            With uTipoCompra
                .Borrar
                'true por false para que se pueda modificar
                .agregar CuentaParam(ID_Cuenta_C_EXENTO), s2n(txtexento), False 'True
                .agregar CuentaParam(ID_Cuenta_C_NOGRABADO), s2n(txtNoGrabado), False ' True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(txtIva), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(txtIva10), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(txtIva27), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IB_CAP), s2n(txtIBcapital), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IB_PROV), s2n(txtIBprovincia), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(txtretengan), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RET_IVA_CPRA), s2n(txtreteniva), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RG3337), s2n(txtper3337), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IMP_INT), s2n(txtimpint), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RG3431), s2n(txtper3431), False 'True
                
                'tiene_c = obtenerDeSQL("select tiene_cuenta from prov where codigo = " & uProv.codigo)
                'If tiene_c = 1 Then
                '    .agregar obtenerDeSQL("select cuenta from prov where codigo = " & uProv.codigo), s2n(txtNeto), False
                'End If
                If pAsiento.Count > 0 Then
                    tiene_c = pAsiento.Count / 3
                    For i = 0 To tiene_c - 1
                        .agregar pAsiento.item("Cuenta" & i), s2n(pAsiento.item("Valor" & i)), False, True
                    Next
                End If
                
            End With
        End If
    End If
End Sub

Private Sub txtanio_GotFocus()
    txtanio = Year(dtFecha)
    PintoFocoActivo
End Sub
Private Sub txtanio_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtcodcaja_GotFocus()
    If Trim$(txtcodcaja) = "" Then txtcodcaja = "1"
    PintoFocoActivo
End Sub
Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtcodcuenta_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

'
'Function estarepetido(Tabla As String, prov As Integer, codigo As Long) As Boolean
'    Dim rs As New ADODB.Recordset
'
'    rs.Open "select * from " & Tabla & " where codpr = " & prov & " and nrodoc = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not rs.EOF Then
'        estarepetido = True
'    Else
'        estarepetido = False
'    End If
'    rs.Close
'    Set rs = Nothing
'
'End Function
Function ObtenerSucursal(tabla As String, codigo As Long) As Long
    Dim rs As New ADODB.Recordset

    rs.Open "select suc from " & tabla & " where codigo = " & Trim(codigo) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        ObtenerSucursal = rs!suc
    Else
        ObtenerSucursal = 0
    End If
    
    rs.Close
    Set rs = Nothing

End Function

Function ObtenerCuit(tabla As String, codigo As Long) As String
If ON_ERROR_HABILITADO Then On Error GoTo fin
Dim rs As New ADODB.Recordset

    rs.Open "select cuit from " & tabla & " where codigo = " & Trim(codigo) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
        ObtenerCuit = rs!CUIT
    Else
        ObtenerCuit = ""
    End If
    
    rs.Close
    
fin:
    Set rs = Nothing
End Function


Private Sub txtcotiz_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtcotiz_KeyPress(KeyAscii As Integer)
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtcotiz_LostFocus()
    txtcotiz = n2r(s2n(txtcotiz, 4), 4)
End Sub

Private Sub txtEfectivo_GotFocus()
''    On Error Resume Next
''    If s2n(txtimporte) <> 0 Then
''        If s2n(txtimporte) <> s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)) Then
''            MsgBox "Los totales no concilian, hay una diferencia de: " & s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)) - s2n(txtimporte)
''        End If
''        txtefectivo = s2n(txtimporte) - (s2n(txtimpcheques) + s2n(txttransf) + uRetCompras.TotalRet)
''    Else
''        MsgBox "Debe ingresar el importe de la factura"
'''        txtimporte.SetFocus
''    End If
''
''    PintoFocoActivo
    If s2n(txtefectivo) = 0 Then txtefectivo = quefaltapagar()
    PintoFocoActivo

End Sub

Private Sub txtefectivo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEfectivo_LostFocus()
    txtefectivo = s2n(txtefectivo)
    
    If s2n(txtefectivo) = 0 Then
        cmbcaja.enabled = False
        txtcodcaja.enabled = False
        txtcaja.enabled = False
        
        txtimpcheques.enabled = True: uCheques.enabled = True
        txttransf.enabled = True
        txtcodcuenta.enabled = True
        cmbcuenta.enabled = True
        
        Exit Sub
    End If
    
    If s2n(txtefectivo) = s2n(txtimporte) Then
        txtimpcheques.enabled = False: uCheques.enabled = False
        txttransf.enabled = False
        txtcodcuenta.enabled = False
        cmbcuenta.enabled = False
        cmbcaja.enabled = True
        txtcodcaja.enabled = True
    Else
        If s2n(txtefectivo) <> "0" Then
            If s2n(s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf)) > s2n(txtimporte) Then
                MsgBox "Con este valor esta superando al importe"
                Exit Sub
            End If
            
            If s2n(txtefectivo) < s2n(txtimporte) Then
                txtimpcheques.enabled = True: uCheques.enabled = True
                txttransf.enabled = True
                txtcodcuenta.enabled = True
                cmbcuenta.enabled = True
                cmbcaja.enabled = True
                txtcodcaja.enabled = True
                txtefectivo = s2n(txtefectivo)
            Else
                MsgBox "El importe en efectivo no puede superar al importe de la factura"
                Exit Sub
            End If
        Else
            txtefectivo = s2n(txtefectivo)
        End If
    End If
End Sub

Private Sub txtexento_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtexento_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtexento_LostFocus()
    txtexento = n2r(s2n(txtexento))
    Importe
End Sub

'Private Sub txtfact_GotFocus()
'    PintoFocoActivo
'End Sub
'Private Sub txtfact_KeyPress(KeyAscii As Integer)
'    If Len(txtfact) < 8 Then
'            If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'                KeyAscii = 0
'            End If
'    Else
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtfact_LostFocus()
'    Dim resu As String
'    resu = ExisteFAC(uProv.codigo, s2n(txtSuc), s2n(txtfact))
'    If resu > "" Then
'        che "Factura existe con fecha " & resu
'    End If
''End Sub
'Private Sub txtfact_LostFocus()
'
'End Sub

Private Sub txtfvto_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtIBcapital_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtIBcapital_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtIBcapital_LostFocus()
    txtIBcapital = n2r(s2n(txtIBcapital))
    Importe
End Sub

Private Sub txtIBprovincia_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtIBprovincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtIBprovincia_LostFocus()
    txtIBprovincia = n2r(s2n(txtIBprovincia))
    Importe True
End Sub

Private Sub txtimpint_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtimpint_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtimpint_LostFocus()
    txtimpint = n2r(s2n(txtimpint))
    Importe
End Sub


Private Function Importe(Optional dFin As Boolean = False) As Double
'    Importe = n2r(s2n(txtneto) + s2n(txtexento) + s2n(txtiva) + s2n(txtiva10) + s2n(txtiva27) + s2n(txtper3337) + s2n(txtper3431) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtreteniva) + s2n(txtIBcapital) + s2n(txtIBprovincia) + s2n(txtNoGrabado))
'    txtimporte = Importe
'    cmbingresar.enabled = s2n(txtimporte) <> 0

'ANTES CALCULABA EL TOTAL DEL DOCUMENTO, AHORA CORROBORA QUE LOS IMPORTES SUMADOS NO SE PASEN DEL TOTAL
'PARA BACIGALUPPI

Dim dTotal As Double, dSuma As Double, dDife As Double

dTotal = s2n(txtimporte)
If dTotal = 0 Then
    MsgBox "Indique total de la factura.", vbInformation
    'txtimporte.SetFocus
    Exit Function
End If

dSuma = s2n(txtneto) + s2n(txtexento) + s2n(txtNoGrabado) + s2n(txtIva) + s2n(txtIva27) + s2n(txtIva10) + s2n(txtper3337) + s2n(txtper3431) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtreteniva) + s2n(txtIBcapital) + s2n(txtIBprovincia)

dDife = s2n(dTotal - dSuma)

If dDife < 0 Then
    MsgBox "Se pasa del total.", vbCritical
End If

If dFin Then
    If dDife > 0 Then
        If MsgBox("Queda " & s2n(dDife, 2, True) & " por asignar." & Chr(13) & "¿Desea sumarlo al exento?", vbYesNo + vbInformation) = vbYes Then
            txtexento = s2n(txtexento + dDife)
        End If
    End If
End If

'uTipoCompra.Total_a_Imputar = s2n(txtimporte)

End Function



Private Sub txtimporte_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtimporte_LostFocus()
txtimporte = n2r(s2n(txtimporte))
Importe
End Sub


Private Sub txtiva_GotFocus()
'    txtiva = s2n(s2n(txtneto) * IVA_21) 's2n(obtenerDeSQL("select Porcentajesiva.Porcentaje from Porcentajesiva  where iva = 2 "))) 'MAGIC 2 = INSCRIPTO
    frmPintoFoco Me
End Sub
Private Sub txtiva_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtiva_LostFocus()
    txtIva = n2r(s2n(txtIva))
    Importe
End Sub
Private Sub txtiva10_GotFocus()
'    If s2n(txtiva) + s2n(txtiva27) = 0 Then txtiva10 = s2n(s2n(txtneto) * IVA_105)
    frmPintoFoco Me
End Sub
Private Sub txtiva10_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtiva10_LostFocus()
    txtIva10 = n2r(s2n(txtIva10))
    Importe
End Sub
Private Sub txtiva27_GotFocus()
'    If s2n(txtiva) = 0 Then txtiva27 = s2n(s2n(txtneto) * IVA_27)
    frmPintoFoco Me
End Sub
Private Sub txtiva27_KeyPress(KeyAscii As Integer)
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtiva27_LostFocus()
        txtIva27 = n2r(s2n(txtIva27))
        Importe
End Sub
Private Sub txtmes_GotFocus()
    txtmes = Month(dtFecha)
End Sub
Private Sub txtmes_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtneto_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtneto_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtneto_LostFocus()
    txtneto = n2r(s2n(txtneto))

'        If uProv.codigo = 0 Then
'            txtiva = s2n(txtneto) * s2n(obtenerDeSQL("select Porcentajesiva.Porcentaje from Porcentajesiva  where iva = 2 ")) 'MAGIC 2 = INSCRIPTO
'        Else
'            If obtenerDeSQL("select letra from Prov inner join ivas on Prov.tipoiva = ivas.codigo where Prov.Codigo = " & uProv.codigo & " and Prov.activo = 1") = "A" Then
'                txtiva = s2n(txtneto) * s2n(obtenerDeSQL("select Porcentajesiva.Porcentaje from Prov inner join Porcentajesiva on Prov.tipoiva = Porcentajesiva.iva where Prov.Codigo = " & uProv.codigo & " and Prov.activo = 1"))
'            Else
'                txtiva = "0"
'            End If
'        End If

    Importe
End Sub


Private Sub txtNoGrabado_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtNoGrabado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtNoGrabado_LostFocus()
txtNoGrabado = n2r(s2n(txtNoGrabado))
Importe
End Sub

Private Sub txtNroIIBB_GotFocus()
    frmPintoFoco Me
End Sub

'Private Sub txtnombre_GotFocus()
'    frmPintoFoco Me
'End Sub
'Private Sub txtnombre_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub
Private Sub txtnumcompra_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtper3337_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtper3337_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtper3337_LostFocus()
   txtper3337 = n2r(s2n(txtper3337))
   Importe
End Sub



Private Sub txtper3431_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtper3431_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtper3431_LostFocus()
        txtper3431 = n2r(s2n(txtper3431))
        Importe
End Sub

Private Sub txtretengan_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtretengan_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtretengan_LostFocus()
    txtretengan = n2r(s2n(txtretengan))
    Importe
End Sub

Private Sub txtreteniva_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtreteniva_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
If KeyAscii = 13 Then
    SendKeys "(tab)"
    KeyAscii = 0
End If
KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtreteniva_LostFocus()
     txtreteniva = n2r(s2n(txtreteniva))
     Importe
End Sub

Private Sub txtsaldo_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtSerie_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtserie_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtSuc_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtsuc_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txttipocompra_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txttransf_GotFocus()
'    If txtimporte <> "" Then
'        If s2n(txtimporte) <> s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia)) Then
'            MsgBox "Los totales no concilian, hay una diferencia de: " & s2n(txtimporte) - s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(txtiva27) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(txtiva10) + s2n(txtIBcapital) + s2n(txtIBprovincia))
'        End If
'    Else
'        MsgBox "Debe ingresar el importe de la factura"
''        txtimporte.SetFocus
'    End If
    
    'txttransf = s2n(s2n(txtimporte) - (s2n(txtefectivo) + s2n(txtimpcheques)))
    If s2n(txttransf) = 0 Then txttransf = quefaltapagar()
    PintoFocoActivo
End Sub
Private Function quefaltapagar() As Double
    quefaltapagar = s2n(Importe() - uRetCompras.TotalRet - s2n(txtefectivo) - s2n(txtimpcheques) - s2n(txttransf))
End Function

Private Sub txttransf_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttransf_LostFocus()
  
    If txttransf <> "" And txttransf <> "0" Then
'        If s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf) > s2n(txtimporte) Then
'            MsgBox "Con este valor esta superando al importe"
'            Exit Sub
'        End If
    
        txtcodcuenta.enabled = True
        cmbcuenta.enabled = True
        txttransf = s2n(txttransf)
    Else
        txttransf = "0"
        txtcodcuenta.enabled = False
        cmbcuenta.enabled = False
    End If
End Sub

Private Sub txtcodcuenta_LostFocus()
    If txtcodcuenta <> "" Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
        
        If txtcuenta = "" Then
            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = "0"
            txtcodcuenta.SetFocus
        Else
            cargar = "CuentasBank"
            CargarDatos
        End If
'    Else
'        If txtcodcuenta <> "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = "0"
'            txtcodcuenta.SetFocus
'        End If
    End If
End Sub

Private Sub txtcodcaja_LostFocus()
    If txtcodcaja <> "" Then
        txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
        If txtcaja = "" Then
            MsgBox "Codigo de caja incorrecto"
        Else
            cargar = "Cajas"
            CargarDatos
        End If
    End If
'    If s2n(txtefectivo) = s2n(txtimporte) Then cmdok.SetFocus
End Sub

'------- POR REMITO -----------------
Public Function PorRemito()
    mPorRemito = True
    Me.Show
End Function
'------- POR REMITO -----------------

Private Sub uCheques_cambio()
    txtimpcheques = uCheques.Total
End Sub


Private Sub uNumDoc_LostFocus()
'    Dim resu As String
'    resu = ExisteFacCompra(uProv.codigo, uNumDoc.suc, uNumDoc.num)
'    If resu > "" Then
'        che "Documento existente " & resu
'    End If
    ExisteFacCompraMSG UpROV.codigo, uNumDoc.suc, uNumDoc.num
End Sub

Private Sub uRetCompras_LostFocus()
    txtTotalRetPago = uRetCompras.TotalRet
End Sub

Private Sub uProv_cambio(codigo As Variant)
    Dim tmp As Variant, tmpTIVA
    If UpROV.codigo = 0 Then
'        txtnombre = ""
        CUIT.Text = ""
        'txttipoiva = ""
        'txtSuc = ""
        uNumDoc.clear
        CUIT.TabStop = True
        txtNroIIBB = ""
        txtNroIIBB.TabStop = True
        uTipoCompra.mProv = 0
    Else
        uTipoCompra.mProv = UpROV.codigo
        tmp = obtenerDeSQL("select cuit, tipoiva, pago, tipoCom, suc, numiibb,provincia,localidad,pais,direccion from Prov where activo = 1 and codigo = " & codigo)
        
        If IsEmpty(tmp) Then
            ufa "No se pudieron cargar datos de proveedor", Me.Name ', Err.Number
        Else
'            txtnombre = uProv.descripcion
            CUIT.Text = sSinNull(tmp(0))

'            txttipoiva = nSinNull(tmp(1))
            cboIva.ListIndex = BuscarEnCombo(cboIva, nSinNull(tmp(1)))
            If Trim(cboIva.Text) = "MONOTRIBUTISTA" Then
                HabilitoControles2 False, True
            Else
                HabilitoControles2 True, True
            End If
            uNumDoc.suc = nSinNull(tmp(4))
            seteoLetra
            
            txtNroIIBB = sSinNull(tmp(5)) '*** *** ***

            If optctacte = True Then cmbformapago.ListIndex = BuscarenComboS(cmbformapago, nSinNull(tmp(2)))
            'cmbtipocompra.ListIndex = tmp(3) 'BuscarenComboS(cmbtipocompra, nSinNull(tmp(3)))
            cmbtipocompra.ListIndex = BuscarenComboS(cmbtipocompra, ObtenerDescripcion("tipocompras", nSinNull(tmp(3))))
            If sSinNull(tmp(6)) <> "" Then CargoProvi sSinNull(tmp(6))
            
            '**************************************************************
            If Not IsNull(tmp(9)) Then
                txtdireccion = sSinNull(tmp(9))
            Else
                txtdireccion = ""
            End If
            If Not IsNull(tmp(7)) Then
                txtLocalidad = sSinNull(tmp(7))
            Else
                txtLocalidad = ""
            End If
            If Not IsNull(tmp(6)) Then
                CmbProvincia.ListIndex = BuscarenComboS(CmbProvincia, ObtenerDescripcionS("provincias", sSinNull(tmp(6))))
            Else
                CmbProvincia.ListIndex = -1
            End If
            If Not IsNull(tmp(8)) Then
                txtpais = sSinNull(tmp(8))
            Else
                txtpais = ""
            End If
            '**************************************************************
            
        End If
        txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
    End If
'    recalcular
End Sub
Private Function CargoProvi(pro As String)
    Dim rRepetido As Boolean
    Dim pRow As Long
    
    rRepetido = False
    For pRow = 1 To gIIBBProvincia.rows - 1
        If gIIBBProvincia.TextMatrix(pRow, 0) = pro Then rRepetido = True
    Next

    If Not rRepetido Then
        With gIIBBProvincia
            .AddItem ""
            pRow = .rows - 1
            .TextMatrix(pRow, 0) = pro
            .TextMatrix(pRow, 1) = obtenerDeSQL("select descripcion from provincias where codigo='" & Trim(pro) & "'")
            .TextMatrix(pRow, 2) = dtFecha
            .TextMatrix(pRow, 3) = "0,00"
        End With
    End If
End Function
Private Function revisarAlProv(COD As Long)
Dim tmp
tmp = obtenerDeSQL("select conperciibbper, conpercganper from clientes where codigo = " & COD)
If tmp(0) = True Then
    If MsgBox("Este cliente tiene percepcion de IIBB personal." & Chr(13) & "¿Desea utilizarlo?", vbYesNo, "Informe") = vbYes Then
        uRetCompras.tieneIIBB = True
    Else
        uRetCompras.tieneIIBB = False
    End If
Else
    uRetCompras.tieneIIBB = False
End If
If tmp(1) = True Then
    If MsgBox("Este cliente tiene percepcion de IIBB personal." & Chr(13) & "¿Desea utilizarlo?", vbYesNo, "Informe") = vbYes Then
        uRetCompras.tieneGAN = True
    Else
        uRetCompras.tieneGAN = False
    End If
Else
    uRetCompras.tieneGAN = False
End If
End Function

Private Sub CambioOptPago()
    'If optcontado = True Then FormadePago (True)
    FormadePago (optcontado = True)
    
    If optcontado = True Then
        cmbformapago.enabled = False
        txtfvto.enabled = False
        If dtFecha.enabled = True Then
            dtFecha.SetFocus
        End If
        UpROV.EditaDescripcion = True
'        TabDetalle.TabEnabled(1) = True
    Else
        cmbformapago.enabled = True
        txtfvto.enabled = True
        UpROV.EditaDescripcion = False
        If UpROV.codigo = 0 Then UpROV.codigo = 0 ' ridiculo?  setea descripcion = ""  ahora puede ser .clear
'        TabDetalle.TabEnabled(1) = False
    End If
End Sub

Private Sub uTipoCompra_GotFocus()
    'uTipoCompra.Total_a_Imputar = s2n(s2n(txtimporte) - sumaTxtIvas() - 0) ' - uRetGan.Total
    
    'uTipoCompra.Total_a_Imputar = Importe() 's2n(txtneto) + s2n(txtexento)
    uTipoCompra.Total_a_Imputar = s2n(txtimporte)
End Sub

Private Function sumaTxtIvas() As Double
    sumaTxtIvas = s2n(txtIva) + s2n(txtIva10) + s2n(txtIva27)
End Function



Private Sub seteoLetra()
    Dim tmpTIVA
    
    tmpTIVA = sSinNull(obtenerDeSQL("select letraprov from ivas where codigo = " & ComboCodigo(cboIva)))     's2n(txttipoiva))
    If IsEmpty(tmpTIVA) Then
        che "Verificar Condición de IVA del proveedor"
    Else
        uNumDoc.letra = tmpTIVA
        tabstopXletra tmpTIVA
    End If
End Sub
Private Sub tabstopXletra(letra)
    Dim esa As Boolean
    esa = (letra <> "C")
    'PARA BACIGALUPPI TIENE QUE TENER TODOS LOS CAMPOS
    'txtexento.TabStop = Not esa
    'txtneto.TabStop = esa
    'txtIva.TabStop = esa
    'txtIva10.TabStop = esa
    'txtIva27.TabStop = esa
End Sub


