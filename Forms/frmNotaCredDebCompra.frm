VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNotaCredDebCompra 
   Caption         =   "Compras: Nota de "
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   Icon            =   "frmNotaCredDebCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodMixto 
      Height          =   315
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   91
      Text            =   "0"
      Top             =   435
      Width           =   2445
   End
   Begin VB.Frame fraBuscar 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   0
      TabIndex        =   78
      Top             =   7065
      Width           =   7080
      Begin VB.OptionButton optCtaCte 
         Caption         =   "Cta Cte"
         Height          =   315
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   45
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optContado 
         Caption         =   "Imputado"
         Height          =   315
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   45
         Width           =   975
      End
      Begin Gestion.ucEntreFechas uBetween 
         Height          =   345
         Left            =   3330
         TabIndex        =   79
         Top             =   30
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
      End
      Begin VB.Label Label27 
         Caption         =   "Opc Busqueda"
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
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   2235
      End
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      Left            =   990
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1380
      Width           =   2190
   End
   Begin VB.TextBox txtNroIIBB 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8430
      TabIndex        =   9
      Top             =   1410
      Width           =   1425
   End
   Begin VB.TextBox txtfechacompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9765
      TabIndex        =   20
      Top             =   900
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttipocompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9525
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtnumcompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9780
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9990
      TabIndex        =   17
      Tag             =   "1"
      Top             =   945
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9480
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtanio 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3945
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   735
   End
   Begin VB.TextBox txtserie 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7350
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Width           =   495
   End
   Begin VB.ComboBox txtmes 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmNotaCredDebCompra.frx":08CA
      Left            =   5205
      List            =   "frmNotaCredDebCompra.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   75
      Width           =   1215
   End
   Begin Gestion.ucBotonera uMenu 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2778
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   1065
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   113770497
      CurrentDate     =   37934
   End
   Begin Gestion.ucCoDe uProv 
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   990
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCuit Cuit 
      Height          =   285
      Left            =   7350
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1005
      Width           =   1335
      _ExtentX        =   2249
      _ExtentY        =   503
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   5295
      Left            =   15
      TabIndex        =   21
      Top             =   1770
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Nota"
      TabPicture(0)   =   "frmNotaCredDebCompra.frx":091D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFactura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Forma de Pago"
      TabPicture(1)   =   "frmNotaCredDebCompra.frx":0939
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Imputaciones Contables"
      TabPicture(2)   =   "frmNotaCredDebCompra.frx":0955
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "uTipoCompra"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraContado 
         BorderStyle     =   0  'None
         Height          =   4650
         Left            =   -74955
         TabIndex        =   58
         Top             =   435
         Width           =   11415
         Begin VB.TextBox txtcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5610
            TabIndex        =   69
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   825
            Width           =   2535
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
            Left            =   4620
            Style           =   1  'Graphical
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   825
            Width           =   855
         End
         Begin VB.TextBox txtcodcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3675
            TabIndex        =   67
            Top             =   840
            Width           =   825
         End
         Begin VB.TextBox txtcodcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3450
            TabIndex        =   66
            Top             =   4260
            Width           =   1215
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
            Left            =   4710
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   4245
            Width           =   855
         End
         Begin VB.TextBox txtcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5700
            TabIndex        =   64
            Tag             =   "2"
            Top             =   4260
            Width           =   3165
         End
         Begin VB.TextBox txttransf 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1335
            TabIndex        =   63
            Top             =   4260
            Width           =   1215
         End
         Begin VB.TextBox txtefectivo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   1320
            TabIndex        =   62
            Top             =   855
            Width           =   1335
         End
         Begin VB.TextBox txtimpcheques 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1305
            TabIndex        =   61
            Top             =   1335
            Width           =   1320
         End
         Begin VB.TextBox txtTotalRetPago 
            Height          =   330
            Left            =   1275
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   210
            Width           =   1365
         End
         Begin Gestion.ucRetCompras uRetCompras 
            Height          =   780
            Left            =   2790
            TabIndex        =   59
            Top             =   15
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   1376
         End
         Begin Gestion.ucCheques uCheques 
            Height          =   2700
            Left            =   2865
            TabIndex        =   70
            Top             =   1335
            Width           =   8520
            _ExtentX        =   14949
            _ExtentY        =   3307
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
            Left            =   2925
            TabIndex        =   76
            Top             =   840
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
            Left            =   2685
            TabIndex        =   75
            Top             =   4260
            Width           =   735
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
            Left            =   600
            TabIndex        =   74
            Top             =   4260
            Width           =   855
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
            Left            =   405
            TabIndex        =   73
            Top             =   900
            Width           =   975
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
            Left            =   285
            TabIndex        =   72
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Retenciones"
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
            TabIndex        =   71
            Top             =   225
            Width           =   1185
         End
      End
      Begin VB.Frame fraFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4830
         Left            =   75
         TabIndex        =   22
         Top             =   375
         Width           =   11475
         Begin VB.CommandButton cmdAgregarProv 
            Caption         =   "Agregar"
            Height          =   300
            Left            =   9180
            TabIndex        =   93
            Top             =   2640
            Width           =   1005
         End
         Begin VB.CommandButton cmdQuitarProv 
            Height          =   300
            Left            =   10215
            Picture         =   "frmNotaCredDebCompra.frx":0971
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   2640
            Width           =   345
         End
         Begin VB.TextBox txtIBcapital 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4365
            TabIndex        =   39
            Top             =   3945
            Width           =   1215
         End
         Begin VB.TextBox txtNoGrabado 
            Height          =   285
            Left            =   6960
            TabIndex        =   30
            Text            =   "0"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox cmbtipocompra 
            Enabled         =   0   'False
            Height          =   315
            Left            =   9540
            Style           =   2  'Dropdown List
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
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
            Left            =   10200
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   4440
            UseMaskColor    =   -1  'True
            Width           =   1065
         End
         Begin VB.TextBox txtIBprovincia 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   40
            Top             =   2490
            Width           =   1215
         End
         Begin VB.ComboBox cmbmoneda 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNotaCredDebCompra.frx":0EFB
            Left            =   1695
            List            =   "frmNotaCredDebCompra.frx":0EFD
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   675
            Width           =   1215
         End
         Begin VB.TextBox txtcotiz 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4305
            TabIndex        =   26
            Top             =   690
            Width           =   1215
         End
         Begin VB.TextBox txtiva10 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            TabIndex        =   33
            Top             =   2940
            Width           =   1215
         End
         Begin VB.TextBox txtper3431 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4335
            TabIndex        =   35
            Top             =   3420
            Width           =   1215
         End
         Begin VB.TextBox txtretengan 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1725
            TabIndex        =   37
            Top             =   3435
            Width           =   1215
         End
         Begin VB.TextBox txtreteniva 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1695
            TabIndex        =   38
            Top             =   3945
            Width           =   1215
         End
         Begin VB.TextBox txtexento 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4335
            TabIndex        =   29
            Top             =   1950
            Width           =   1215
         End
         Begin VB.TextBox txtiva27 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4335
            TabIndex        =   32
            Top             =   2475
            Width           =   1215
         End
         Begin VB.TextBox txtimpint 
            Enabled         =   0   'False
            Height          =   285
            Left            =   9600
            TabIndex        =   36
            Top             =   1890
            Width           =   1215
         End
         Begin VB.TextBox txtneto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1740
            TabIndex        =   28
            Top             =   1935
            Width           =   1215
         End
         Begin VB.TextBox txtimporte 
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1725
            TabIndex        =   27
            Top             =   1425
            Width           =   1215
         End
         Begin VB.TextBox txtiva 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1725
            TabIndex        =   31
            Top             =   2415
            Width           =   1215
         End
         Begin VB.TextBox txtper3337 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1725
            TabIndex        =   34
            Top             =   2940
            Width           =   1215
         End
         Begin VB.ComboBox cmbformapago 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1725
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   225
            Width           =   3840
         End
         Begin MSComCtl2.DTPicker txtfvto 
            Height          =   300
            Left            =   6555
            TabIndex        =   24
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   113770497
            CurrentDate     =   37934
         End
         Begin VSFlex7LCtl.VSFlexGrid gIIBBProvincia 
            Height          =   1290
            Left            =   6360
            TabIndex        =   94
            Top             =   3060
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
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
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
            Left            =   5760
            TabIndex        =   90
            Top             =   1950
            Width           =   1215
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
            TabIndex        =   89
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblImputaciones 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Centro de Costos:"
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
            Left            =   8400
            TabIndex        =   87
            Top             =   4440
            Width           =   1815
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
            Left            =   5625
            TabIndex        =   57
            Top             =   2505
            Width           =   1215
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
            TabIndex        =   56
            Top             =   705
            Width           =   735
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
            TabIndex        =   55
            Top             =   720
            Width           =   780
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
            Left            =   3330
            TabIndex        =   54
            Top             =   2955
            Width           =   975
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
            Left            =   2985
            TabIndex        =   53
            Top             =   3405
            Width           =   1215
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
            Left            =   15
            TabIndex        =   52
            Top             =   3465
            Width           =   1605
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
            Left            =   2655
            TabIndex        =   51
            Top             =   3960
            Width           =   1605
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
            Left            =   345
            TabIndex        =   50
            Top             =   3945
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
            Left            =   2985
            TabIndex        =   49
            Top             =   1980
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
            Left            =   2985
            TabIndex        =   48
            Top             =   2475
            Width           =   1215
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
            Left            =   8835
            TabIndex        =   47
            Top             =   1920
            Width           =   1215
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
            Left            =   15
            TabIndex        =   46
            Top             =   1470
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
            Left            =   15
            TabIndex        =   45
            Top             =   1965
            Width           =   1605
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Iva Gral."
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
            Left            =   15
            TabIndex        =   44
            Top             =   2430
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
            Left            =   15
            TabIndex        =   43
            Top             =   3000
            Width           =   1605
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
            TabIndex        =   42
            Top             =   255
            Width           =   1605
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
            TabIndex        =   41
            Top             =   315
            Width           =   705
         End
      End
      Begin Gestion.ucTipoCompra uTipoCompra 
         Height          =   3870
         Left            =   -74925
         TabIndex        =   77
         Top             =   645
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   6826
      End
   End
   Begin Gestion.uNumDoc uNumDoc 
      Height          =   300
      Left            =   4860
      TabIndex        =   8
      Top             =   1395
      Width           =   2625
      _ExtentX        =   4736
      _ExtentY        =   529
   End
   Begin VB.Label Label32 
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
      Left            =   4080
      TabIndex        =   85
      Top             =   1425
      Width           =   870
   End
   Begin VB.Label Label30 
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
      Left            =   465
      TabIndex        =   84
      Top             =   1380
      Width           =   435
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
      Left            =   7620
      TabIndex        =   83
      Top             =   1425
      Width           =   855
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
      Left            =   345
      TabIndex        =   15
      Top             =   120
      Width           =   615
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
      Left            =   3465
      TabIndex        =   14
      Top             =   90
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
      Left            =   4710
      TabIndex        =   13
      Top             =   105
      Width           =   615
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
      Left            =   6765
      TabIndex        =   12
      Top             =   90
      Width           =   615
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
      Left            =   30
      TabIndex        =   11
      Top             =   990
      Width           =   975
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
      Left            =   6900
      TabIndex        =   10
      Top             =   1005
      Width           =   495
   End
End
Attribute VB_Name = "frmNotaCredDebCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private midDoc  As Long
Private mTipoDoc As String ' N/D N/C
Private mTipo   As tipoNotaCompra
Private rsmov   As New ADODB.Recordset
Private mTabla  As String
Private mCoefIva As Double
Private EsBusqueda As Boolean
Private pAsiento As Collection

Public Enum tipoNotaCompra
    tipoNotaCompraNC
    tipoNotaCompraND
End Enum

Public Sub mostrar(que As tipoNotaCompra)
    If que = tipoNotaCompraNC Then
        mTipoDoc = "N/C"
        Me.caption = "NOTA CREDITO PROVEEDOR"
        mTipo = que
        Me.Show
    ElseIf que = tipoNotaCompraND Then
        mTipoDoc = "N/D"
        Me.caption = "NOTA DEBITO PROVEEDOR"
        mTipo = que
        Me.Show
    Else
        Exit Sub
    End If
End Sub

Private Sub cmbFormaPago_LostFocus()
    revisarFechVto
End Sub

Private Sub cmbingresar_Click()
    If txtimporte <> "" Then
        If s2nt(txtimporte) <> s2n(s2nt(txtneto) + s2nt(txtIva) + s2nt(txtPer3337) + s2nt(txtIva27) + s2nt(txtExento) + s2nt(txtRetenIva) + s2nt(txtImpInt) + s2nt(txtRetenGan) + s2nt(txtPer3431) + s2nt(txtIva10) + s2nt(txtIBcapital) + s2nt(txtIBprovincia) + s2nt(txtNoGrabado)) Then
            MsgBox "Los totales no concilian, hay una diferencia de: " & (s2nt(txtneto) + s2nt(txtIva) + s2nt(txtPer3337) + s2nt(txtIva27) + s2nt(txtExento) + s2nt(txtRetenIva) + s2nt(txtImpInt) + s2nt(txtRetenGan) + s2nt(txtPer3431) + s2nt(txtIva10) + s2nt(txtIBcapital) + s2nt(txtIBprovincia)) - s2nt(txtimporte)
            Exit Sub
        End If
    Else
        MsgBox "Debe ingresar el importe de la factura"
        txtimporte.SetFocus
        Exit Sub
    End If
    
    FrmCostosYContable.Tag = Me.Name
    vieneDE = Me.Name
    FrmCostosYContable.CargarImputacion s2n(txtneto) + s2n(txtExento) + s2n(txtNoGrabado), s2n(txtimporte), 1
    'FrmCostosYContable.txtimptotal = txtimporte
    'FrmCostosYContable.txtimporte = txtNeto
    FrmCostosYContable.Show
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

Private Sub dtFecha_Change()
    revisarFechVto
    ff
    txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

Private Function ff()
    txtanio = Year(Date) 'Year(dtfecha)
    txtmes = Month(Date) 'Month(dtfecha)
End Function

Private Sub dtFecha_Click()
    revisarFechVto
    ff
    txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

Private Sub dtFecha_LostFocus()
    revisarFechVto
    ff
    txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

'Private Function notaDe() As String
'    notaDe = IIf(mTipo = tipoNotaCompraNC, "Credito", "Debito")
'End Function
Private Sub Form_Load()

'    CargaCombo2 cmbformapago, "FormasPago", "descripcion", "codigo", ""
    CargaCombo2 cmbtipocompra, "TipoCompras", "descripcion", "codigo", ""
    comboSql cmbformapago, "select Descripcion, codigo from FormasPago where activo = 1 order by dias"
    CargaCombo3 cmbMoneda, "Monedas", "descripcion", "codigo", ""
    comboSql cboIva, "select descripcion, codigo from ivas order by codigo"

    UpROV.ini "select descripcion from prov where activo = 1 and codigo = '###'", "select codigo as [ Codigo ],cuit as [  Cuit           ], descripcion as [ Descripcion               ] from prov where activo = 1 order by codigo ", False
    TabDetalle.Tab = 0
    TabDetalle.TabVisible(1) = False
    uMenu.init True, True, False, True, True
    
    cmbMoneda.ListIndex = 0
    Set pAsiento = New Collection
    
    iniGrillaProv

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   FrmKeyPress KeyAscii, True, True
End Sub

Private Sub cmbformapago_Change()
    revisarFechVto
End Sub
Private Function revisarFechVto()
    txtfvto = dtFecha + s2n(obtenerDeSQL("select dias from FormasPago where descripcion = '" & cmbformapago & "' and activo = 1"))
End Function
Sub HabilitoControles(habilito As Boolean)

    'txtfact.enabled = habilito
    uNumDoc.enabled = habilito
'    txtnombre.Enabled = habilito
    CUIT.enabled = habilito
    dtFecha.enabled = habilito
    cmbformapago.enabled = habilito
    txtfvto.enabled = habilito
    
'    txtplan.Enabled = habilito
    
    txtneto.enabled = habilito
    txtIva.enabled = habilito
    txtPer3337.enabled = habilito
    txtimporte.enabled = habilito
    txtIva27.enabled = habilito
    txtExento.enabled = habilito
    txtRetenIva.enabled = habilito
    txtImpInt.enabled = habilito
    txtRetenGan.enabled = habilito
    txtPer3431.enabled = habilito
    txtIva10.enabled = habilito
    txtIBcapital.enabled = habilito
    txtIBprovincia.enabled = habilito
    txtNoGrabado.enabled = habilito
'    txtnombre.Enabled = habilito
    cmbMoneda.enabled = habilito
    
    txtanio.enabled = habilito
    txtmes.enabled = habilito
'    txtsuc.enabled = habilito
    txtserie.enabled = habilito
    txtcotiz.enabled = habilito
    
    uTipoCompra.enabled = habilito
    cmbtipocompra.enabled = habilito
End Sub

Sub LimpioControles()
    midDoc = 0
    mCoefIva = 0
    dtFecha = Date
    UpROV.clear
    uTipoCompra.Borrar
    CUIT.Text = ""
    FrmBorrarTxt Me
    TabDetalle.Tab = 0
    cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, "CONTADO", True)
    revisarFechVto
    FrmCostosYContable.LimpioControles
'    FrmCostosYContable.InicioGrilla
    FrmCostosYContable.InicioGrillaCostos
    cmbtipocompra.ListIndex = -1
    Set pAsiento = New Collection
    iniGrillaProv
End Sub

'    txtprov = ""
'    txtfact = ""
 '   txtnombre = ""
'    cmbformapago = ""
'    txtfvto = Date
'    txtneto = "0"
'    txtiva = "0"
'    txtper3337 = "0"
'    txtImporte = "0"
'    txtiva27 = "0"
'    txtexento = "0"
'    txtreteniva = "0"
'    txtimpint = "0"
'    txtretengan = "0"
'    txtper3431 = "0"
'    txtiva10 = "0"
'    txtIBcapital = "0"
'    txtIBprovincia = ""
''    txtplan = "0"
'
'    txtanio = ""
''    txtmes = ""
'    txtsuc = ""
'    txtserie = ""
'    txtcotiz = ""
'
'    txtsaldo = ""
'    cargar = ""
   
'    Ope = ""

'    uTipoCompra.Clear

Sub CargoRegistro()
    Dim z As Double
    
    midDoc = nSinNull(rsmov!iddoc)
    'cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, BuscoDato("FormasPago", nSinNull(rsmov!FormadePago)))
    cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, nSinNull(rsmov!FormadePago))
    If optcontado.Value = True Then
        txtfvto = rsmov!vto_co
    Else
        txtfvto = rsmov!vencim
    End If
    
    txtcotiz = rsmov!cotizacion
    cmbMoneda = ObtenerDescripcion("Monedas", rsmov!moneda)
    
    If Trim(cmbMoneda.Text) <> "Pesos" Then
        z = s2n(txtcotiz, 4)
        If z = 0 Then z = 1
    Else
        z = 1
    End If
    
    UpROV.codigo = rsmov!CODPR
'    txtprov = ObtenerDescripcion("Prov", rsmov!CODPR)
'    txtfact = rsmov!NroDoc
    uNumDoc.num = rsmov!NroDoc
    ' rsmov!razonsocialprov
    CUIT.Text = rsmov!cuitprov
    'uTipoCompra.codigo = rsmov!tipocompra
    dtFecha = rsmov!Fecha
    cmbtipocompra = BuscoDato("TipoCompras", rsmov!Tipocompra)
    
    txtneto = s2n(rsmov!Neto / z)
    txtIva = s2n(rsmov!IVA_21 / z)
    txtPer3337 = s2n(rsmov!percepc / z)
    txtimporte = s2n(rsmov!Total / z)
    txtIva27 = s2n(rsmov!IVA_27 / z)
    txtExento = s2n(rsmov!EXENTO / z)
    txtRetenIva = s2n(rsmov!iva_9 / z)
    txtImpInt = s2n(rsmov!imp_int / z)
    txtRetenGan = s2n(nSinNull(rsmov!retgan / z))
    txtPer3431 = s2n(rsmov!der_est / z)
    txtIva10 = s2n(rsmov!iva_10 / z)
    txtIBcapital = s2n(rsmov!ibcapital / z)
    txtIBprovincia = s2n(rsmov!ibprovincia / z)
    txtNoGrabado = s2n(rsmov!nogravado / z)
    txtCodMixto = rsmov!codmixto
    If optcontado.Value = True Then
        txtsaldo = 0
    Else
        txtsaldo = s2n(rsmov!saldo / z)
    End If
    txtanio = rsmov!anoimp
'    txtmes = rsmov!mesimp
    'txtsuc = rsmov!suc
    uNumDoc.suc = rsmov!suc
    txtserie = rsmov!Serie
    'txtcotiz = rsmov!cotizacion
    'cmbMoneda = ObtenerDescripcion("Monedas", rsmov!moneda)
End Sub

Private Sub Form_Resize()
    Anclar fraBuscar, Me, anclarAbajo + anclarIzquierda
End Sub

Private Sub txtcotiz_LostFocus()
    txtcotiz = s2n(txtcotiz, 4, True)
End Sub

Private Sub txtexento_LostFocus()
    txtExento = s2n(txtExento, 2, True)
    Importe
End Sub
'
'Private Sub txtfact_LostFocus()
'    If uProv.codigo > 0 Then
'        If DocCompraCargado(uProv.codigo, uNumDoc.num) Then '(estarepetido("transcom", uProv.codigo, unumdoc.num) Or estarepetido("compras", uProv.codigo, unumdoc.num)) Then
'            MsgBox "Factura Repetida", 48, "Atencion"
'            txtfact.SetFocus
'        End If
'    End If
'End Sub


Private Sub txtIBcapital_LostFocus()
    txtIBcapital = s2n(txtIBcapital, 2, True)
    Importe
End Sub

Private Sub txtIBprovincia_LostFocus()
    txtIBprovincia = s2n(txtIBprovincia, 2, True)
    Importe True
End Sub

Private Sub txtNoGrabado_LostFocus()
    txtNoGrabado = s2n(txtNoGrabado, 2, True)
    Importe
End Sub

Private Sub uMenu_Nuevo()
    On Error Resume Next
    dtFecha.SetFocus
    FrmCostosYContable.LimpioControles
'    FrmCostosYContable.InicioGrilla
    FrmCostosYContable.InicioGrillaCostos
    EsBusqueda = False
    txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
End Sub

Private Sub uNumDoc_LostFocus()
    'If uProv.codigo > 0 Then
    '    If DocCompraCargado(uProv.codigo, uNumDoc.num) Then  '(estarepetido("transcom", uProv.codigo, unumdoc.num) Or estarepetido("compras", uProv.codigo, unumdoc.num)) Then
    '        MsgBox "Factura Repetida", 48, "Atencion"
    '    End If
    'End If
    
    'Dim resu As String
    
    'resu =
    ExisteFacCompraMSG UpROV.codigo, uNumDoc.suc, uNumDoc.num
    'If resu > "" Then
    '    che "Documento existente " & resu
    'End If
End Sub
'Private Function ExisteFAC(prov, suc, Nro) As String
'    Dim resu
'    If prov = 0 Then Exit Function
'
'    resu = obtenerDeSQL("select fecha from compras where codpr = " & prov & " and suc = " & suc & " and NroDoc = " & Nro)
'    If Not IsEmpty(resu) Then
'        ExisteFAC = CStr(resu)
'        Exit Function
'    End If
'
'    resu = obtenerDeSQL("select fecha from transcom where codpr = " & prov & " and suc = " & suc & " and NroDoc = " & Nro)
'    If Not IsEmpty(resu) Then
'        ExisteFAC = CStr(resu)
'        Exit Function
'    End If
'End Function

Private Sub uProv_cambio(codigo As Variant)
    Dim tmp As Variant, tmpTIVA
    If UpROV.codigo = 0 Then
        CUIT.Text = ""
        'txttipoiva = ""
        uNumDoc.clear
        uTipoCompra.mProv = 0
    Else
        uTipoCompra.mProv = UpROV.codigo
    
        tmp = obtenerDeSQL("select cuit, tipoiva, pago, tipoCom,suc, numiibb,provincia,localidad,pais,direccion from Prov where activo = 1 and codigo = " & codigo)
        
        If IsEmpty(tmp) Then
            ufa "No se pudo cargar dato de proveedor", Me.Name
        Else
            CUIT.Text = sSinNull(tmp(0))
            'txttipoiva = s2n(tmp(1))
            
            mCoefIva = s2n(obtenerDeSQL("select porcentaje from PorcentajesIva where iva = " & ComboCodigo(cboIva) & " and activo = 1 "))
            'txtsuc = sSinNull(tmp(4)) ' ObtenerSucursal("Prov", Val(txtcodprov))
            'tmpTIVA = obtenerDeSQL("select letra from ivas where codigo = " & s2n(txttipoiva))
            'If IsEmpty(tmpTIVA) Then
            '    che "Verificar Condición de IVA del proveedor"
            'End If
            
            cboIva.ListIndex = BuscarEnCombo(cboIva, nSinNull(tmp(1)))
            uNumDoc.suc = nSinNull(tmp(4))
            txtNroIIBB = Prov_NumIIBB(UpROV.codigo)
            seteoLetra
            cmbtipocompra.ListIndex = BuscarenComboS(cmbtipocompra, ObtenerDescripcion("tipocompras", nSinNull(tmp(3))))
            
            If sSinNull(tmp(6)) <> "" Then CargoProvi sSinNull(tmp(6))
            
        End If
        txtCodMixto = NuevoCodigoMixto(UpROV.codigo, dtFecha)
    End If

End Sub

Private Sub seteoLetra()
    Dim tmpTIVA
    
    tmpTIVA = sSinNull(obtenerDeSQL("select letraprov from ivas where codigo = " & ComboCodigo(cboIva)))     's2n(txttipoiva))
    If IsEmpty(tmpTIVA) Then
        che "Verificar Condición de IVA del proveedor"
    Else
        uNumDoc.letra = tmpTIVA
'        tabstopXletra tmpTIVA
    End If
End Sub

'Function estarepetido(Tabla As String, prov As Integer, codigo As Long) As Boolean
'    Dim rs As New ADODB.Recordset
'
'    rs.Open "select * from " & Tabla & " where codpr = " & prov & " and nrodoc = " & codigo & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not rs.EOF Then
'        estarepetido = True
'    Else
'        estarepetido = False
'    End If
'    rs.Close
'    Set rs = Nothing
'End Function

Private Function sumaTxtIvas() As Double
    sumaTxtIvas = s2n(txtIva) + s2n(txtIva10) + s2n(txtIva27)
End Function

Private Function TodoOk() As Boolean
    'If Trim(txtfact) = "" Then
    '    MsgBox "Debe ingresar el número de factura"
    '    Exit Function
    'End If
    Dim resu As String
    
    If Not PuedoCompras(dtFecha.Value) Then
        'msg en la funcion
        Exit Function
    End If
    
    If Not uNumDoc.NumDocValido Then
        che "revisar numero doc"
        Exit Function
    End If

    If UpROV.codigo = 0 Then
        MsgBox "Debe ingresar el código de proveedor"
        Exit Function
    End If
    
    If Trim(cmbMoneda.Text) <> "Pesos" Then
        If txtcotiz = "" Then
            MsgBox "Debe ingresar una cotizacion.", , "ATENCION"
            Exit Function
        End If
    End If
    
'    If DocCompraCargado(uProv.codigo, uNumDoc.num) Then '(estarepetido("transcom", uProv.codigo, unumdoc.num) Or estarepetido("compras", uProv.codigo, unumdoc.num)) Then
'        MsgBox "Factura Repetida", 48, "Atencion"
'        'txtfact.SetFocus
'        Exit Function
'    End If
'    resu = ExisteFacCompra(uProv.codigo, uNumDoc.suc, uNumDoc.num)
'    If resu > "" Then
'        che "Documento Existente " & resu
'        Exit Function
'    End If
    If ExisteFacCompraMSG(UpROV.codigo, uNumDoc.suc, uNumDoc.num) Then
        Exit Function
    End If
    
    If s2n(txtimporte) <> s2n(s2n(txtneto) + s2n(txtIva) + s2n(txtPer3337) + s2n(txtIva27) + s2n(txtExento) + s2n(txtRetenIva) + s2n(txtImpInt) + s2n(txtRetenGan) + s2n(txtPer3431) + s2n(txtIva10) + s2n(txtIBcapital) + s2n(txtIBprovincia) + s2n(txtNoGrabado)) Then
        MsgBox "Los totales no concilian, hay una diferencia de: " & s2n(txtimporte) - (s2n(txtneto) + s2n(txtIva) + s2n(txtPer3337) + s2n(txtIva27) + s2n(txtExento) + s2n(txtRetenIva) + s2n(txtImpInt) + s2n(txtRetenGan) + s2n(txtPer3431) + s2n(txtIva10) + s2n(txtIBcapital) + s2n(txtIBprovincia) + s2n(txtNoGrabado))
        Exit Function
    End If
    
    If cmbtipocompra.ListIndex = -1 Then
        MsgBox "Debe ingresar el tipo de compra"
        Exit Function
    End If

    TodoOk = True
End Function

Private Sub uMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaOK

    If Not TodoOk() Then Exit Sub
    
    If TrabaIva(dtFecha.Value) Then
        MsgBox "La fecha del comprobante esta dentro de las fechas trabadas para emision," & Chr(13) & "verifiquelo con su contadora.", , "ATENCION"
        Exit Sub
    End If
    
    Dim k As Long
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
    
    
    Dim strSql As String, i As Long
    Dim rs As New ADODB.Recordset
    Dim AsientoCompra As New Asiento, iddoc As Long, NroDoc As Long ', TIPODOC As String
    Dim TextoAsientoComprobante As String
    Dim z As Double
    
    NroDoc = uNumDoc.num 's2n(txtfact, 0)
    'TIPODOC = "N/C"
    
    If Trim(cmbMoneda.Text) <> "Pesos" Then
        z = s2n(txtcotiz, 4)
        If z = 0 Then z = 1
    Else
        z = 1
    End If
    
    'BREHM lo quiere en detalle, lo correcto seria ponerlo en cabecera y sacarlo de ahi al imprimir...
    TextoAsientoComprobante = mTipoDoc & " " & NroDoc
    
    With AsientoCompra
        'HEADER ASIENTO
        .nuevo mTipoDoc & UpROV.DESCRIPCION, dtFecha, mTipoDoc
        If mTipo = tipoNotaCompraNC Then
            'HABER
            For i = 1 To uTipoCompra.rows
                .AgregarItem uTipoCompra.imCuenta(i), 0, uTipoCompra.imMonto(i) * z ', TextoAsientoComprobante
                'iva, iibb, perc, retgan
            Next i
            
'            .AgregarItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), 0, s2n(txtiva) + s2n(txtiva10) + s2n(txtiva27), TextoAsientoComprobante
'            .AgregarItem CuentaParam(ID_Cuenta_C_IB_PROV), 0, s2n(txtIBprovincia)
'            .AgregarItem CuentaParam(ID_Cuenta_C_IB_CAP), 0, s2n(txtIBcapital)
'            .AgregarItem uTipoCompra.codigo, 0, s2n(txtneto), TextoAsientoComprobante
            'DEBE
            .AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(txtimporte * z), 0, TextoAsientoComprobante
        Else ' ND
            'debe
            For i = 1 To uTipoCompra.rows
                .AgregarItem uTipoCompra.imCuenta(i), uTipoCompra.imMonto(i) * z, 0 ', TextoAsientoComprobante
                'iva, iibb, perc, retgan
            Next i
            
'            .AgregarItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), 0, s2n(txtiva) + s2n(txtiva10) + s2n(txtiva27), TextoAsientoComprobante
'            .AgregarItem CuentaParam(ID_Cuenta_C_IB_PROV), 0, s2n(txtIBprovincia)
'            .AgregarItem CuentaParam(ID_Cuenta_C_IB_CAP), 0, s2n(txtIBcapital)
'            .AgregarItem uTipoCompra.codigo, 0, s2n(txtneto), TextoAsientoComprobante
            'DEBE
            .AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(txtimporte * z), TextoAsientoComprobante
        
        End If
        If .Diferencia <> 0 Then
            ufa "err prg:No cierra asiento", "grabar NC compra"
            Exit Sub
        End If
    End With
           
        '*****************************************
        DE_BeginTrans
            iddoc = NuevoDocumento(mTipoDoc, NroDoc, UpROV.codigo, 0) ' NO HAY CONTADO
            midDoc = iddoc
        
            Dim porciva As Double
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
            
            DataEnvironment1.dbo_INGCOMPRASCTACTE "A", dtFecha, val(txtanio), val(txtmes) _
                , UpROV.codigo, UpROV.DESCRIPCION, CUIT.Text, ComboCodigo(cboIva), mTipoDoc, uNumDoc.num _
                , uNumDoc.suc, "", s2n(txtimporte * z), s2n(txtneto * z), s2n(txtimporte * z), CDate(txtfvto), s2n(txtIva * z), s2n(txtPer3337 * z), s2n(txtIva27 * z), s2n(txtRetenIva * z) _
                , s2n(txtIva10 * z), s2n(txtImpInt * z), s2n(txtRetenGan * z), s2n(txtPer3431 * z), s2n(txtExento * z), s2n(txtIBcapital * z), s2n(txtIBprovincia * z) _
                , val(txtserie), porciva, ObtenerCodigo("Tipocompras", cmbtipocompra.Text), ObtenerCodigo("Formaspago", cmbformapago.Text), 0, ObtenerCodigo("Monedas", cmbMoneda.Text) _
                , z, 0, Date, UsuarioSistema!codigo, iddoc, 0, 0 _
                , uNumDoc.letra, Prov_NumIIBB(UpROV.codigo), uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo
                

            If txtNoGrabado <> "" Then
                txtNoGrabado = s2n(txtNoGrabado)
            Else
                txtNoGrabado = 0
            End If
            'strSql = "update transcom set codmixto=" & ssTexto(txtCodMixto) & ",nogravado=" & x2s(txtNoGrabado * z) & " where codpr=" & UpROV.codigo & " and tipodoc='" & mTipoDoc & "' and NroDoc=" & NroDoc
            DataEnvironment1.Sistema.Execute "update transcom set codmixto=" & ssTexto(txtCodMixto) & ",nogravado=" & x2s(txtNoGrabado * z) & " where codpr=" & UpROV.codigo & " and tipodoc='" & mTipoDoc & "' and NroDoc=" & NroDoc
            
            If Trim(mTipoDoc) = "N/C" Then
                DetalleIIBB "A", TIPODOC_NC_PROVEEDOR, uNumDoc.num, UpROV.codigo, iddoc, z
            Else
                DetalleIIBB "A", TIPODOC_ND_PROVEEDOR, uNumDoc.num, UpROV.codigo, iddoc, z
            End If
            
            'If
            If siAsiento("AsientosCompras") Then AsientoCompra.Grabar iddoc '= 0 Then
'            DE_RollbackTrans
'            che "fallo al grabar asiento"
'            GoTo fin 'perdon, pero si tratamiento de err no tiene try{}catch{} ...
'         End If

            Dim x As Long
        
            If gEMPR_ConSistContable Then
                If FrmCostosYContable.grillacostos.rows > 1 And FrmCostosYContable.grillacostos.TextMatrix(1, 1) > "" Then
                    
                    'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                    For x = 1 To FrmCostosYContable.grillacostos.rows - 1
                        DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(FrmCostosYContable.grillacostos.TextMatrix(x, 0)), _
                        dtFecha, "NCC", val(uNumDoc.num), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) + s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, "", FrmCostosYContable.grillacostos.TextMatrix(x, 4), UpROV.codigo
                    Next
                End If
            End If
        
        DE_CommitTrans
        '*****************************************
        
        MsgBox "Operación Realizada con éxito", vbOKOnly
        uMenu.AceptarOk
        
        FrmCostosYContable.LimpioControles
'        FrmCostosYContable.InicioGrilla
        FrmCostosYContable.InicioGrillaCostos
        Unload FrmCostosYContable
fin:
    Exit Sub
UfaOK:
    midDoc = 0
    DE_RollbackTrans
    ufa "err al grbar", "ok NC compra"
    Resume fin
End Sub
Private Function Importe(Optional dFin As Boolean = False) As Double
'    Importe = n2r(s2n(txtneto) + s2n(txtexento) + s2n(txtiva) + s2n(txtiva10) + s2n(txtiva27) + s2n(txtper3337) + s2n(txtper3431) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtreteniva) + s2n(txtIBcapital) + s2n(txtIBprovincia))
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

dSuma = s2n(txtneto) + s2n(txtExento) + s2n(txtIva) + s2n(txtIva27) + s2n(txtIva10) + s2n(txtPer3337) + s2n(txtPer3431) + s2n(txtImpInt) + s2n(txtRetenGan) + s2n(txtRetenIva) + s2n(txtIBcapital) + s2n(txtIBprovincia) + s2n(txtNoGrabado)

dDife = s2n(dTotal - dSuma)

If dDife < 0 Then
    MsgBox "Se pasa del total.", vbCritical
End If

If dFin Then
    If dDife > 0 Then
        If MsgBox("Queda " & s2n(dDife, 2, True) & " por asignar." & Chr(13) & "¿Desea sumarlo al exento?", vbYesNo + vbInformation) = vbYes Then
            txtExento = s2n(txtExento + dDife)
        End If
    End If
End If

'uTipoCompra.Total_a_Imputar = s2n(txtimporte)
End Function

Private Sub txtimpint_LostFocus()
    txtImpInt = s2n(txtImpInt, 2, True)
    Importe
End Sub
Private Sub txtiva_LostFocus()
    txtIva = s2n(txtIva, 2, True)
    Importe
End Sub
Private Sub txtiva10_LostFocus()
    txtIva10 = s2n(txtIva10, 2, True)
    Importe
End Sub
Private Sub txtiva27_LostFocus()
    txtIva27 = s2n(txtIva27, 2, True)
    Importe
End Sub
Private Sub txtneto_LostFocus()
    txtneto = s2n(txtneto, 2, True)
    'txtiva = s2n(txtneto) * mCoefIva
    Importe
End Sub
Private Sub txtper3337_LostFocus()
   txtPer3337 = s2n(txtPer3337, 2, True)
   Importe
End Sub
Private Sub txtper3431_LostFocus()
    txtPer3431 = s2n(txtPer3431, 2, True)
    Importe
End Sub
Private Sub txtretengan_LostFocus()
    txtRetenGan = s2n(txtRetenGan, 2, True)
    Importe
End Sub
Private Sub txtreteniva_LostFocus()
     txtRetenIva = s2n(txtRetenIva, 2, True)
     Importe
End Sub

Private Sub uTipoCompra_GotFocus()
    'uTipoCompra.Total_a_Imputar = Importe() 's2n(txtneto) + s2n(txtexento)
    uTipoCompra.Total_a_Imputar = s2n(txtimporte) 'Importe() 's2n(txtneto) + s2n(txtexento)
End Sub
Private Sub txtanio_GotFocus()
    txtanio = Year(dtFecha)
    PintoFocoActivo
End Sub
Private Sub txtmes_GotFocus()
    txtmes = Month(dtFecha)
End Sub

Private Function RevisoCeldas()
    txtneto = s2n(txtneto)
    txtIva = s2n(txtIva)
    txtPer3337 = s2n(txtPer3337)
    txtimporte = s2n(txtimporte)
    txtIva27 = s2n(txtIva27)
    txtExento = s2n(txtExento)
    txtRetenIva = s2n(txtRetenIva)
    txtImpInt = s2n(txtImpInt)
    txtRetenGan = s2n(txtRetenGan)
    txtPer3431 = s2n(txtPer3431)
    txtIva10 = s2n(txtIva10)
    txtNoGrabado = s2n(txtNoGrabado)
    txtIBprovincia = s2n(txtIBprovincia)
    txtIBcapital = s2n(txtIBcapital)
End Function

Private Sub TabDetalle_Click(PreviousTab As Integer)
'    Dim x As String, tiene_c As Long
'
'    If TabDetalle.Tab = 1 Then
'        uRetCompras.Calcular UpROV.codigo, s2n(txtneto) + s2n(txtexento), s2n(txtneto) + s2n(txtexento), dtFecha
'    ElseIf TabDetalle.Tab = 2 Then
'        With uTipoCompra
'            .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA), sumaTxtIvas(), True
'            .agregar CuentaParam(ID_Cuenta_C_IB_CAP), s2n(txtIBcapital), True
'            .agregar CuentaParam(ID_Cuenta_C_IB_PROV), s2n(txtIBprovincia), True
'            .agregar CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(txtretengan), True
'            .agregar CuentaParam(ID_Cuenta_C_RET_IVA_CPRA), s2n(txtreteniva), True
'            tiene_c = obtenerDeSQL("select tiene_cuenta from prov where codigo = " & UpROV.codigo)
'            If tiene_c = 1 Then
'                .agregar obtenerDeSQL("select cuenta from prov where codigo = " & UpROV.codigo), s2n(txtneto), False
'            End If
'        End With
'
'    End If
    
    
    RevisoCeldas
    
    Dim x As String, tiene_c As Long
    Dim auxIDDOC As Long, auxIDAsiento As Long, i As Long
    Dim rsAsiento As New ADODB.Recordset
    Dim z As Double
    If TabDetalle.Tab = 1 Then
        uRetCompras.Calcular UpROV.codigo, s2n(txtneto) + s2n(txtExento), s2n(txtneto) + s2n(txtExento) + s2n(txtNoGrabado), dtFecha
        txtTotalRetPago = uRetCompras.TotalRet
    ElseIf TabDetalle.Tab = 2 Then
        If EsBusqueda Then
            auxIDDOC = s2n(midDoc) 's2n(lblIDDOC)
            auxIDAsiento = obtenerDeSQL("select idasiento from asientos where iddoc=" & auxIDDOC)
            If mTipo = tipoNotaCompraND Then
                rsAsiento.Open "select * from mayor where debe>0 and idasiento= " & auxIDAsiento & " order by idmayor desc", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            Else
                rsAsiento.Open "select * from mayor where haber>0 and idasiento= " & auxIDAsiento & " order by idmayor desc", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            End If
            With rsAsiento
                If .EOF And .BOF Then
                Else
                    .MoveFirst
                    z = 1 's2n(txtcotiz, 4)
                    'If z = 0 Then z = 1
                    For i = 0 To .RecordCount - 1
                        If mTipo = tipoNotaCompraND Then
                            uTipoCompra.agregar !Cuenta, !Debe * z, True
                        Else
                            uTipoCompra.agregar !Cuenta, !haber * z, True
                        End If
                       .MoveNext
                    Next
                End If
            End With
        Else
            With uTipoCompra
                .Borrar
                'true por false para que se pueda modificar
                .agregar CuentaParam(ID_Cuenta_C_EXENTO), s2n(txtExento), False 'True
                .agregar CuentaParam(ID_Cuenta_C_NOGRABADO), s2n(txtNoGrabado), False ' True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(txtIva), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(txtIva10), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(txtIva27), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IB_CAP), s2n(txtIBcapital), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IB_PROV), s2n(txtIBprovincia), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(txtRetenGan), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RET_IVA_CPRA), s2n(txtRetenIva), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RG3337), s2n(txtPer3337), False 'True
                .agregar CuentaParam(ID_Cuenta_C_IMP_INT), s2n(txtImpInt), False 'True
                .agregar CuentaParam(ID_Cuenta_C_RG3431), s2n(txtPer3431), False 'True
                
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

Private Sub uMenu_BorrarControles()
    LimpioControles
End Sub
Private Sub uMenu_Buscar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    Dim ss As String, resu As String
    Dim FECO, tico, nuco
    Dim prco
    mTabla = IIf(optctacte, "Transcom", "Compras")
    
    
    ss = "select Fecha as [ Fecha   ], TipoDoc as [ Tipo   ], NroDoc  as [ Numero     ], Total as [ Importe    ], codpr as [Proveedor ] from " & mTabla & " where activo = 1 and TipoDoc = '" & mTipoDoc & "' and fecha " & uBetween.ssBetween & " "
    With frmBuscar
        LimpioControles
        resu = .MostrarSql(ss)
        If resu > "" Then
'            cmdeliminar.Enabled = (mTabla = "Transcom")
            FECO = .resultado(1)
            tico = .resultado(2)
            nuco = .resultado(3)
            prco = .resultado(5)
            
            txttipocompra = .resultado(2)
            txtnumcompra = .resultado(3)
'
            rsmov.Open "select * from " & mTabla & " where tipodoc = '" & tico & "' and nrodoc = " & ssNum(nuco) & " and activo = 1 AND CODPR =  " & ssNum(prco), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            CargoRegistro
            DetalleIIBB "C", txttipocompra, txtnumcompra, .resultado(5), rsmov!iddoc, s2n(txtcotiz)
            rsmov.Close
            uMenu.BuscarOK
            EsBusqueda = True
        Else
            EsBusqueda = False
        End If
    End With

fin:
    If rsmov.State = 1 Then rsmov.Close
    Exit Sub
ufaChe:
    ufa "err al cargar", " buscar"
    Resume fin
End Sub

Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    Dim mensaje As Long
    
    If Not PuedoCompras(dtFecha) Then
        'msg en funcion
        Exit Sub
    End If
                           
    If s2n(txtsaldo) <> s2n(txtimporte) Then
        che "No se puede anular este coprobante dado que fue parcialmente pagado"
        Exit Sub
    End If
    

    '*****************************************
'    Debug.Print midDoc

'    If Not BorroDocumento(midDoc) Then 'AsientoBaja_idDoc(mIdDoc) Then
'        ufa "err no se pudo borrar dociumento", "middoc " & midDoc
'        DE_RollbackTrans
'        Exit Sub
'    End If

    DE_BeginTrans
        DataEnvironment1.dbo_INGCOMPRASCTACTE "B", 0, 0, 0, UpROV.codigo, "", "", 0, "", uNumDoc.num _
             , 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, val(txtserie), 0, 0, 0, 0, 0, 0, 0, 0 _
             , Date, UsuarioSistema!codigo, midDoc, 0, 0 _
             , "", "", 0, 0
        BorroDocumento midDoc
        DataEnvironment1.dbo_GRABARBITACORA UpROV.codigo, mTabla, UsuarioSistema!codigo, Date, Time, "B"
        
        DetalleIIBB "B", TIPODOC_FAC_PROVEEDOR, uNumDoc.num, UpROV.codigo, midDoc
        
    DE_CommitTrans
    
    MsgBox "Se ha eliminado correctamente.", , "ATENCION"
    uMenu.EliminarOK
    
    '*****************************************

fin:
    Exit Sub
UFAelim:
    ufa "prg: error al eliminar", "NC/ND compra " & UpROV.codigo
    DE_RollbackTrans
    Resume fin
End Sub

Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoControles sino
End Sub

Private Sub uMenu_SALIR()
    Set rsmov = Nothing
    Unload Me
End Sub
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
txtimporte = s2n(txtimporte, 2, True)
Importe
End Sub

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

Private Function iniGrillaProv()
DetalleIIBB "I"
End Function

Private Sub gIIBBProvincia_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If Col <> 3 Then cancel = True
End Sub

Private Sub gIIBBProvincia_CellChanged(ByVal Row As Long, ByVal Col As Long)
    calGrillaProv
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

