VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFactProvBoleta 
   Caption         =   "Carga de Boletas"
   ClientHeight    =   7770
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "FrmFactProvBoleta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucFecha uFeHa 
      Height          =   315
      Left            =   2580
      TabIndex        =   92
      Top             =   7290
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      FechaInit       =   0
   End
   Begin Gestion.ucFecha uFeDe 
      Height          =   300
      Left            =   1725
      TabIndex        =   91
      Top             =   7290
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   529
      FechaInit       =   0
   End
   Begin Gestion.ucCuit CUIT 
      Height          =   315
      Left            =   8520
      TabIndex        =   90
      Top             =   480
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
   End
   Begin Gestion.ucCoDe UpROV 
      Height          =   300
      Left            =   1290
      TabIndex        =   89
      Top             =   495
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   529
      CodigoWidth     =   1000
   End
   Begin Gestion.uNumDoc uNumDoc 
      Height          =   345
      Left            =   5175
      TabIndex        =   88
      Top             =   870
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   609
   End
   Begin Gestion.uCtaBanco uCtaBanco 
      Height          =   345
      Left            =   1725
      TabIndex        =   86
      Top             =   1305
      Visible         =   0   'False
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   609
   End
   Begin VB.TextBox txtNroIIBB 
      Height          =   285
      Left            =   8655
      TabIndex        =   21
      Top             =   900
      Width           =   1425
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      Left            =   1215
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   870
      Width           =   2190
   End
   Begin VB.Frame fraMesImputacion 
      Height          =   495
      Left            =   5145
      TabIndex        =   75
      Top             =   -30
      Width           =   3330
      Begin VB.TextBox txtanio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   555
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   150
         Width           =   735
      End
      Begin VB.ComboBox txtmes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmFactProvBoleta.frx":08CA
         Left            =   1830
         List            =   "FrmFactProvBoleta.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   76
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   150
         Width           =   615
      End
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
      TabIndex        =   71
      Top             =   7350
      Width           =   975
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   5505
      Left            =   45
      TabIndex        =   22
      Top             =   1755
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9710
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Boleta"
      TabPicture(0)   =   "FrmFactProvBoleta.frx":091D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFactura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sin uso"
      TabPicture(1)   =   "FrmFactProvBoleta.frx":0939
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContado"
      Tab(1).Control(1)=   "uRetCompras"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Imputaciones Contables"
      TabPicture(2)   =   "FrmFactProvBoleta.frx":0955
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "uTipoCompra"
      Tab(2).ControlCount=   1
      Begin Gestion.ucRetCompras uRetCompras 
         Height          =   810
         Left            =   -73425
         TabIndex        =   93
         Top             =   375
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   1429
      End
      Begin Gestion.ucTipoCompra uTipoCompra 
         Height          =   4950
         Left            =   -74850
         TabIndex        =   87
         Top             =   450
         Width           =   11400
         _ExtentX        =   20108
         _ExtentY        =   8731
      End
      Begin VB.Frame fraFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4590
         Left            =   45
         TabIndex        =   54
         Top             =   360
         Width           =   11475
         Begin VB.TextBox txtValesAlim 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   13
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtLRT 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   14
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtAportesOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   12
            Top             =   1845
            Width           =   1215
         End
         Begin VB.TextBox txtContSS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   8
            Top             =   510
            Width           =   1215
         End
         Begin VB.TextBox txtVALOR_RESTA 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   1
            Top             =   825
            Width           =   1215
         End
         Begin VB.TextBox txtImporte 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   3690
            Width           =   1215
         End
         Begin VB.TextBox txtIMP_DETERN 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   0
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtIVAVENTAS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   5
            Top             =   2460
            Width           =   1215
         End
         Begin VB.TextBox txtVALOR_SUMA 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   2
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtIVACOMPRAS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   4
            Top             =   2085
            Width           =   1215
         End
         Begin VB.TextBox txtARRASTREMES 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   7
            Top             =   3180
            Width           =   1215
         End
         Begin VB.TextBox txtContOS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   11
            Top             =   1515
            Width           =   1215
         End
         Begin VB.TextBox txtAportesSS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   9
            Top             =   855
            Width           =   1215
         End
         Begin VB.TextBox txtPOSICIONIVA 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   6
            Top             =   2820
            Width           =   1215
         End
         Begin VB.TextBox txtRECARGO_INT 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   3
            Top             =   1545
            Width           =   1215
         End
         Begin VB.TextBox txtContR 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7485
            TabIndex        =   10
            Top             =   1185
            Width           =   1215
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vales alimentarios"
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
            Left            =   4995
            TabIndex        =   84
            Top             =   2205
            Width           =   2385
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "L.R.T."
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
            Left            =   4725
            TabIndex        =   83
            Top             =   2520
            Width           =   2685
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aportes de OS"
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
            Left            =   5055
            TabIndex        =   82
            Top             =   1875
            Width           =   2325
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contribuciones de SS"
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
            Left            =   4455
            TabIndex        =   67
            Top             =   570
            Width           =   2925
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valores Restan"
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
            Left            =   930
            TabIndex        =   66
            Top             =   840
            Width           =   1605
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto Determinado"
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
            TabIndex        =   65
            Top             =   480
            Width           =   2490
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
            Left            =   5445
            TabIndex        =   64
            Top             =   3735
            Width           =   1605
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IVA Ventas"
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
            Left            =   1320
            TabIndex        =   63
            Top             =   2085
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Valores Suman"
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
            Left            =   540
            TabIndex        =   62
            Top             =   1185
            Width           =   1995
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "IVA Compras"
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
            Left            =   990
            TabIndex        =   61
            Top             =   2445
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Arrastre Mes anterior"
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
            Left            =   390
            TabIndex        =   60
            Top             =   3195
            Width           =   2145
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contribuciones de OS"
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
            Left            =   4890
            TabIndex        =   59
            Top             =   1560
            Width           =   2490
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aportes de SS"
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
            Left            =   4920
            TabIndex        =   58
            Top             =   885
            Width           =   2460
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Posicion de IVA"
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
            Left            =   -510
            TabIndex        =   57
            Top             =   2835
            Width           =   3045
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Recargos Internos"
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
            Left            =   735
            TabIndex        =   56
            Top             =   1560
            Width           =   1800
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contribuciones RENATRE"
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
            Left            =   4890
            TabIndex        =   55
            Top             =   1200
            Width           =   2460
         End
      End
      Begin VB.Frame fraContado 
         BorderStyle     =   0  'None
         Height          =   4650
         Left            =   -74955
         TabIndex        =   48
         Top             =   435
         Width           =   11415
         Begin Gestion.ucCheques uCheques 
            Height          =   2850
            Left            =   1500
            TabIndex        =   94
            Top             =   1290
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   5027
         End
         Begin VB.TextBox txtTotalRetPago 
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtimpcheques 
            Enabled         =   0   'False
            Height          =   300
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1605
            Width           =   1320
         End
         Begin VB.TextBox txtefectivo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   45
            TabIndex        =   26
            Top             =   825
            Width           =   1335
         End
         Begin VB.TextBox txttransf 
            Enabled         =   0   'False
            Height          =   285
            Left            =   75
            TabIndex        =   31
            Top             =   4305
            Width           =   1215
         End
         Begin VB.TextBox txtcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4425
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtcodcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2175
            TabIndex        =   32
            Top             =   4335
            Width           =   1215
         End
         Begin VB.TextBox txtcodcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2355
            TabIndex        =   27
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
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   855
            Width           =   855
         End
         Begin VB.TextBox txtcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4290
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   855
            Width           =   2535
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
            TabIndex        =   70
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   870
            Width           =   855
         End
      End
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8775
      TabIndex        =   47
      Top             =   7785
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6255
      TabIndex        =   46
      Tag             =   "1"
      Top             =   7860
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnumcompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6900
      TabIndex        =   45
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttipocompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5355
      TabIndex        =   44
      Top             =   7815
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtfechacompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7875
      TabIndex        =   43
      Top             =   7845
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtserie 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9180
      TabIndex        =   19
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
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
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
      TabIndex        =   17
      Top             =   120
      Value           =   -1  'True
      Visible         =   0   'False
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
      TabIndex        =   30
      Top             =   -60
      Visible         =   0   'False
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
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7350
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
      TabIndex        =   37
      Top             =   7320
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
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7335
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7305
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
      TabIndex        =   36
      Top             =   7335
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
      TabIndex        =   38
      Top             =   7335
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   3570
      TabIndex        =   18
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   132644865
      CurrentDate     =   37934
   End
   Begin VB.Label Label33 
      Caption         =   "Cuenta Bancaria"
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
      TabIndex        =   85
      Top             =   1320
      Width           =   1515
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
      Left            =   7845
      TabIndex        =   81
      Top             =   915
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
      TabIndex        =   80
      Top             =   870
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
      TabIndex        =   74
      Top             =   7320
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
      Left            =   10755
      TabIndex        =   73
      Top             =   105
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
      TabIndex        =   72
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
      TabIndex        =   69
      Top             =   525
      Width           =   570
   End
   Begin VB.Label Label3 
      Caption         =   "Boleta-Prov"
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
      TabIndex        =   68
      Top             =   495
      Width           =   1095
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
      TabIndex        =   42
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Comprobante:"
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
      Left            =   3765
      TabIndex        =   41
      Top             =   915
      Width           =   1365
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
      TabIndex        =   40
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmFactProvBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EsBusqueda As Boolean
Dim Ope As String
Dim rsmov As New ADODB.Recordset
Private midDoc As Long
Private mMovBanc As Long

Private Sub cboIva21_LostFocus()
    seteoLetra
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

Private Sub cmdBuscar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaBuscar
    
    Dim tabla, re, fech As Date, ssql As String, titu As String
    
    'If optcontado = True Then
    '    tabla = "compras"
    '    titu = "Contado y canceladas"
    'Else
        tabla = "GastosBoletas"
        titu = "Boletas"
    'End If
    ssql = " select Fecha as [Fecha ], tipoDoc as [Doc], NroDoc as [ Numero ], total as  [ Importe ], codpr as [ Cod Prov], razonsocial as [ Razon social                           ], iddoc " & _
        " from " & tabla & " left join prov on codpr = prov.codigo  " & _
        " where " & tabla & ".activo = 1 and fecha " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & _
        " order by fecha desc "
    
    re = frmBuscar.MostrarSql(ssql, , titu) ', , , , " " , "Anulada")
    If re > "" Then
'        LimpioControles
        Habilitobotones True, True, True, False, True, True
        fech = CDate(frmBuscar.resultado(1))
        txttipocompra = frmBuscar.resultado(2)
        txtnumcompra = frmBuscar.resultado(3)

        rsmov.Open "select * from " & tabla & " where fecha = " & ssFecha(fech) & " and tipodoc = '" & txttipocompra & "' and iddoc = " & frmBuscar.resultado(7) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        CargoRegistro
        Habilitobotones True, True, True, False, True, True
        EsBusqueda = True
    Else
        EsBusqueda = False
    End If
    
    Set rsmov = Nothing
Exit Sub
UfaBuscar:
    MsgBox "Error al buscar Documento", vbCritical
End Sub


Private Sub cmdCancelar_Click()
    LimpioControles
    HabilitoControles (False)
    FormadePago (False)
    Call Habilitobotones(True, True, False, False, True, False)
    uTipoCompra.Borrar
    EsBusqueda = False
End Sub

Private Sub cmdeliminar_Click()
If ON_ERROR_HABILITADO Then On Error GoTo UFAelimina
Dim cad As String
    
    If Not PuedoCompras(dtFecha) Then
        Exit Sub
    End If
    
    
    If MsgBox("Esta seguro que desea elimnar este registro?", vbYesNo, "Atencion") = vbYes Then
    
        '*************************************************
        DE_BeginTrans
        
        If optcontado = True Then

        Else  ' Cta Cte osea gastos bancarios ja
            cad = "update gastosboletas set activo=0 where iddoc=" & midDoc
            DataEnvironment1.Sistema.Execute cad
            'cad = "update movibanc set activo=0 where iddoc=" & midDoc
            'DataEnvironment1.SISTEMA.Execute cad
            'cad = "update registrodocumentos set activo=0 where iddoc=" & midDoc
            'DataEnvironment1.Sistema.Execute cad
          
        End If

            'Baja Doc y asiento
            If midDoc > 0 Then ' SI fue generado con este sistema, bajo asiento
                If Not BorroDocumento(midDoc) Then
                    MsgBox "Error al borrar documento " & midDoc, vbCritical
                    GoTo UFAelimina:
                End If
            End If
            
        DE_CommitTrans
        '*************************************************

        MsgBox "Los Movimientos se anularon correctamente"
        LimpioControles
        HabilitoControles (True)
        Call Habilitobotones(True, True, False, False, True, False)
    End If
    
Exit Sub
UFAelimina:
    DE_RollbackTrans
    MsgBox "Error al anular", vbCritical
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
    uNumDoc.letra = "N"
    uNumDoc.suc = 1
    uNumDoc.num = NuevaBoleta(UpROV.codigo)
    
    'cmbmoneda.ListIndex = BuscarenComboS(cmbmoneda, Const_PESOS)
    
    HabilitoControles (True)
    FormadePago (False)
    Call Habilitobotones(False, False, False, True, True, False)
    
    uCheques.Borrar
    
    Ope = "A"
    
    optctacte.Value = True ' = vbChecked
    txtanio = Year(dtFecha)
    txtmes = Month(dtFecha)
    
    
    
End Sub

Private Function NuevaBoleta(prov As Long) As Long
NuevaBoleta = nSinNull(obtenerDeSQL("select max(nrodoc) from gastosboletas where codpr=" & prov)) + 1
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
    Dim tiene_c
    Dim a
    
    'If uCtaBanco.codigo = 0 Then
    '    MsgBox "No se indico la cuenta bancaria", vbCritical
    '    Exit Sub
    'End If
    
    If Not PuedoCompras(dtFecha) Then
        Exit Sub
    End If

    EsBusqueda = False
    
    If s2n(txtimporte) = 0 Then
        MsgBox "No se cargo ningun valor", vbCritical
        Exit Sub
    End If
    
    If uNumDoc.num = 0 Then
        MsgBox "Debe ingresar el número de comprobante"
         uNumDoc.SetFocus
        Exit Sub
    End If
    
    If Trim(UpROV.DESCRIPCION) = "" Then
        MsgBox "Debe ingresar el Banco", vbExclamation
        UpROV.SetFocus
        Exit Sub
    End If
    
    If Not optcontado And UpROV.codigo = 0 Then
        MsgBox "Debe ingresar el Banco", vbExclamation
        UpROV.SetFocus
        Exit Sub
    End If

    If ExisteBoletaMSG(UpROV.codigo, uNumDoc.suc, uNumDoc.num) Then
        uNumDoc.SetFocus
        Exit Sub
    End If
    
    'If cmbFormaPago.ListIndex = -1 And optCtaCte = True Then
    '    MsgBox "Debe ingresar la forma de pago"
    '    cmbFormaPago.SetFocus
    '    Exit Sub
    'End If
    
    If s2n(txtefectivo) > 0 And s2n(txtcodcaja) = 0 Then
        MsgBox "Falta Nro de Caja para efectivo", vbExclamation
        Exit Sub
    End If
    
    If s2n(txttransf) <> 0 And Trim(txtcuenta) = "" Then
        MsgBox "Falta cuenta de transferencia", vbExclamation
        Exit Sub
    End If
    
    
    Dim dife As Double
    Dim ladoA As Double, ladoB As Double, ladoC As Double, totalAB As Double
    ladoA = s2n(s2n(txtIMP_DETERN) + s2n(txtVALOR_SUMA) + s2n(txtRECARGO_INT) - s2n(txtVALOR_RESTA))
    
    txtPOSICIONIVA = s2n(txtIVAVENTAS) - s2n(txtIVACOMPRAS)
    ladoB = s2n(s2n(txtPOSICIONIVA) + s2n(txtARRASTREMES))
    
    ladoC = s2n(txtAportesSS) + s2n(txtContSS) + s2n(txtContR) + s2n(txtContOS) + s2n(txtAportesOS) + s2n(txtValesAlim) + s2n(txtLRT)
    
    totalAB = ladoA + ladoB + ladoC
    dife = s2n(s2n(txtimporte) - s2n(totalAB))
    
    '******************sacar
    dife = 0
    
    If dife <> 0 Then
        MsgBox "Los totales no concilian, hay una diferencia de: " & dife
        Exit Sub
    End If
    
    If optcontado = True And (s2n(txtimpcheques) <> 0) And uCheques.Total = 0 Then
         MsgBox "No se ingresaron los cheques correspondientes al pago"
        Exit Sub
    End If
    
    If s2n(txtimpcheques) <> uCheques.Total Then
        MsgBox "No coincide el total de cheques", vbExclamation
        Exit Sub
    End If
    
    If Not uCheques.FechasOk Then Exit Sub
    
    If optcontado = True And s2n(s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf) + uRetCompras.TotalRet - Importe()) <> 0 Then
        MsgBox "El total de la forma de pago no coincide con el importe de la factura"
        Exit Sub
    End If
    
    '*************************sacar
    gEMPR_ConSistContable = False
    
    If gEMPR_ConSistContable Then
        If uTipoCompra.Diferencia <> 0 Or uTipoCompra.Total_a_Imputar = 0 Then
            MsgBox "Falta completar imputaciones contables", vbExclamation
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
     
    TextoAsientoComprobante = "GTO " & NroDoc
    With AsientoCompra
        'HEADER ASIENTO
        .nuevo "Bta. " & UpROV.DESCRIPCION, dtFecha, TIPODOC_FAC_BOLETA
        'DEBE
        For i = 1 To uTipoCompra.rows
            .AgregarItem uTipoCompra.imCuenta(i), uTipoCompra.imMonto(i), 0      ', TextoAsientoComprobante
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

    'If optContado = True Then
    '    NroPago = NuevoNroPago()
    '    If uRetCompras.retgan > 0 Then NroCertifGan = NuevoNroCertifGan()
    '    If uRetCompras.retIB > 0 Then NroCertifIIBB = NuevoNroCertifIIBB()
    'End If
    
    
    iddoc = NuevoDocumento(TIPODOC_FAC_BOLETA, NroDoc, UpROV.codigo, NroPago, NroCertifGan, NroCertifIIBB)
    midDoc = iddoc
    mMovBanc = nuevoCodigo("movibanc", "movBanco")
    
    If Trim(Ope) <> "" Then
        If Ope = "A" Then
                   
            'Dim porciva As Double
            'Dim maximobanc As Long, maximocaja As Long, x As Long
            'Dim valorcuenta As String, valorcuentacon As String, valorcartera As String
       
            'sAssert = "1) %iva "
            'rs.Open "select * from porcentajesiva where iva = " & ComboCodigo(cboIva) & " order by fecha_baja", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            'While Not rs.EOF
            '    If IsNull(rs!fecha_baja) Then
            '        porciva = rs!PORCENTAJE
            '    Else
            '        porciva = 0
            '    End If
            '    rs.MoveNext
            'Wend
            'rs.Close
            'Set rs = Nothing
            
            If optcontado = True Then
            Else            ' *********  alta FC CUENTA CORRIENTE ******* OSEA GASTOS BANCARIOS jeje
                Dim gPeriodo As String
                gPeriodo = qMes(dtFecha)
                
'                If txtIMP_DETERN = "0" Then
'                ElseIf s2n(txtIMP_DETERN) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos " & gPeriodo, _
'                    dtfecha, "G", s2n(txtIMP_DETERN), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtVALOR_RESTA = "0" Then
'                ElseIf s2n(txtVALOR_RESTA) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 21% " & gPeriodo, _
'                    dtfecha, "G", s2n(txtVALOR_RESTA), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtRECARGO_INT = "0" Then
'                ElseIf s2n(txtRECARGO_INT) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 10.5% " & gPeriodo, _
'                    dtfecha, "G", s2n(txtRECARGO_INT), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtVALOR_SUMA = "0" Then
'                ElseIf s2n(txtVALOR_SUMA) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 27% " & gPeriodo, _
'                    dtfecha, "G", s2n(txtVALOR_SUMA), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtcontss = "0" Then
'                ElseIf s2n(txtcontss) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Mant. Cuenta " & gPeriodo, _
'                    dtfecha, "G", s2n(txtcontss), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtaportesss = "0" Then
'                ElseIf s2n(txtaportesss) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Mant. Cuenta Sueldos " & gPeriodo, _
'                    dtfecha, "G", s2n(txtaportesss), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtcontr = "0" Then
'                ElseIf s2n(txtcontr) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos por Chequera " & gPeriodo, _
'                   dtfecha, "G", s2n(txtcontr), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtcontos = "0" Then
'                ElseIf s2n(txtcontos) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos Varios " & gPeriodo, _
'                    dtfecha, "G", s2n(txtcontos), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'*************************************************de aca hasta los puntos de abajo estan pero no hacen nada, no deberian hacer nada
                'Comisiones por Gestion de Cheques
'                If txtIVACOMPRAS = "0" Then
'                ElseIf s2n(txtIVACOMPRAS) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Comisiones por Gestion de Cheq " & gPeriodo, _
'                    dtFecha, "G", s2n(txtIVACOMPRAS), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtaportesos = "0" Then
'                ElseIf s2n(txtaportesos) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Sircreb " & gPeriodo, _
'                    dtFecha, "G", s2n(txtaportesos), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtlrt = "0" Then
'                ElseIf s2n(txtlrt) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Cred " & gPeriodo, _
'                    dtFecha, "G", s2n(txtlrt), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtvalesalim = "0" Then
'                ElseIf s2n(txtvalesalim) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Deb " & gPeriodo, _
'                    dtFecha, "G", s2n(txtvalesalim), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtIVAVENTAS = "0" Then
'                ElseIf s2n(txtIVAVENTAS) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Sobre Giro " & gPeriodo, _
'                    dtFecha, "G", s2n(txtIVAVENTAS), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
                'Comisiones por Transferencias
'                If txtPOSICIONIVA = "0" Then
'                ElseIf s2n(txtPOSICIONIVA) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Comisiones por Transferencias " & gPeriodo, _
'                    dtFecha, "G", s2n(txtPOSICIONIVA), mMovBanc, midDoc, Date, UsuarioActual
'                End If
                
'                If txtARRASTREMES = "0" Then
'                ElseIf s2n(txtARRASTREMES) > 0 Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Percepcion de IIBB " & gPeriodo, _
'                    dtFecha, "G", s2n(txtARRASTREMES), mMovBanc, midDoc, Date, UsuarioActual
'                End If
'***********************************************************************
                
                Dim CuentaBoleta As String, Dat
                Dat = obtenerDeSQL("select tiene_cuenta,cuenta from prov where codigo=" & UpROV.codigo)
                
                If s2n(Dat(0)) = 1 Then
                    CuentaBoleta = Dat(1)
                Else
                    CuentaBoleta = 1 'cambiar por la parametrizada
                End If
                
                AsientoCompra.AgregarItem CuentaBoleta, 0, s2n(txtimporte), TextoAsientoComprobante
                
                a = uNumDoc.num
                AGastos "A", dtFecha, s2n(txtanio), s2n(txtmes), UpROV.codigo, UpROV.DESCRIPCION, CUIT.Text, ComboCodigo(cboIva), TIPODOC_FAC_BOLETA, _
                            uNumDoc.num, Importe, uNumDoc.suc, uCtaBanco.codigo, 1, s2n(txtIMP_DETERN), 0, 0, s2n(txtRECARGO_INT), s2n(txtVALOR_RESTA), s2n(txtVALOR_SUMA), _
                             s2n(txtIVACOMPRAS), s2n(txtValesAlim), s2n(txtLRT), s2n(txtARRASTREMES), s2n(txtAportesOS), s2n(txtIVAVENTAS), 0, 0, s2n(txtContSS), _
                             s2n(txtAportesSS), s2n(txtContR), s2n(txtContOS), s2n(txtPOSICIONIVA), 0, 0, 0, Date, 0, 0, 0, midDoc, UsuarioActual, uNumDoc.letra, txtNroIIBB, 0, 0
            End If
        End If
      
        sAssert = "15) ASIENTOS"
        'AsientoCompra.Grabar iddoc

        DE_CommitTrans

        MsgBox "Operación Realizada con éxito", vbOKOnly
        ImprimirFacProv
        LimpioControles
        HabilitoControles (False)
        Call Habilitobotones(True, True, False, False, True, False)
        uCheques.Borrar
    End If


Exit Sub
UfaOK:
    midDoc = 0
    DE_RollbackTrans
    uCheques.resetNroIntPropios
    MsgBox "Error al grabar factura", vbCritical
End Sub

Public Function qMes(ff As Date) As String
Dim mm As Long
mm = Month(ff)
    Select Case mm
        Case 1: qMes = "Enero"
        Case 2: qMes = "Febrero"
        Case 3: qMes = "Marzo"
        Case 4: qMes = "Abril"
        Case 5: qMes = "Mayo"
        Case 6: qMes = "Junio"
        Case 7: qMes = "Julio"
        Case 8: qMes = "Agosto"
        Case 9: qMes = "Septiembre"
        Case 10: qMes = "Octubre"
        Case 11: qMes = "Noviembre"
        Case 12: qMes = "Diciembre"
    End Select
End Function

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

Private Sub cmdSalir_Click()
    If confirma("¿Cerrar formulario?") Then Unload Me
End Sub

Sub LimpioControles()
    'uProv.codigo = 0
    FrmBorrarTxt Me
    uNumDoc.clear
    lblIDDOC = ""
    txtcodcaja = 1
    midDoc = 0
    CUIT.Text = ""
    dtFecha = Date

    optcontado = False
    optctacte = True

    txtIMP_DETERN = "0"
    txtVALOR_RESTA = "0"
    txtRECARGO_INT = "0"
    txtVALOR_SUMA = "0"
    txtimporte = "0"
    txtIVACOMPRAS = "0"
    txtAportesOS = "0"
    txtLRT = "0"
    txtValesAlim = "0"
    txtContSS = "0"
    txtAportesSS = "0"
    txtContR = "0"
    txtContOS = "0"
    txtIVAVENTAS = "0"
    txtPOSICIONIVA = "0"
    txtARRASTREMES = "0"
    
    cargar = ""
    Ope = ""
     uCheques.Borrar
     uTipoCompra.Borrar
     TabDetalle.Tab = 0
End Sub

Sub HabilitoControles(habilito As Boolean)
    uCheques.enabled = habilito
    uRetCompras.enabled = habilito

    uNumDoc.enabled = habilito
    cboIva.enabled = habilito
    
'    txtnombre.Enabled = habilito
    CUIT.enabled = habilito
    dtFecha.enabled = habilito
    
    txtanio.enabled = habilito
    txtmes.enabled = habilito

    txtserie.enabled = habilito
    
    txtIMP_DETERN.enabled = habilito
    txtVALOR_RESTA.enabled = habilito
    txtRECARGO_INT.enabled = habilito
    txtVALOR_SUMA.enabled = habilito
    txtimporte.enabled = habilito
    txtIVACOMPRAS.enabled = habilito
    txtAportesOS.enabled = habilito
    txtLRT.enabled = habilito
    txtValesAlim.enabled = habilito
    txtContSS.enabled = habilito
    txtAportesSS.enabled = habilito
    txtContR.enabled = habilito
    txtContOS.enabled = habilito
    txtIVAVENTAS.enabled = habilito
    txtPOSICIONIVA.enabled = habilito
    txtARRASTREMES.enabled = habilito
    
    txtefectivo.enabled = habilito
    txtcodcaja.enabled = habilito
    cmbcaja.enabled = habilito
    txtimpcheques.enabled = habilito: uCheques.enabled = habilito
    txttransf.enabled = habilito
    cmbcuenta.enabled = habilito
    
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
    
    lblIDDOC = rsmov!iddoc
    UpROV.codigo = rsmov!CODPR
    UpROV.DESCRIPCION = rsmov!RAZONSOCIAL
    
    uNumDoc.num = rsmov!NroDoc
    
    CUIT.Text = rsmov!CUIT

    dtFecha = rsmov!Fecha
    uCtaBanco.codigo = rsmov!Cuenta
    txtIMP_DETERN = rsmov!imp_detern
    txtVALOR_RESTA = rsmov!valor_resta
    txtRECARGO_INT = rsmov!recargo_int
    txtVALOR_SUMA = rsmov!valor_suma
    txtAportesOS = rsmov!aportesos
    txtimporte = rsmov!Total
    txtIVACOMPRAS = rsmov!IVACOMPRAS
    txtLRT = rsmov!lrt
    txtValesAlim = rsmov!VALESALIM
    txtContSS = rsmov!contribucionss
    txtAportesSS = rsmov!aportesss
    txtContR = rsmov!contribucionr
    txtContOS = rsmov!contribucionos
    txtIVAVENTAS = rsmov!ivaventas
    txtPOSICIONIVA = rsmov!posicioniva
    txtARRASTREMES = rsmov!arrastreiva
    midDoc = rsmov!iddoc
    
    txtanio = rsmov!anoimp
    txtmes = rsmov!mesimp
    'txtSuc = rsmov!suc
    uNumDoc.suc = rsmov!suc
    
    'txtserie = rsmov!Serie
    'txtcotiz = rsmov!COTIZACION
    'cmbmoneda = ObtenerDescripcion("Monedas", rsmov!MONEDA)
    
    uRetCompras.retgan = rsmov!retganpago
    uRetCompras.retIB = rsmov!IBPAGO
    
    If optcontado = True Then 'no se va a usar
        'hay que mejorar esto no es logico que por cada dato alla una funcion
        
        'txttransf = ObtenerTransferencia("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        'txtcodcaja = ObtenerCaja("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        'txtcodcuenta = ObtenerCuenta("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        'txtcaja = ObtenerDescripcionCajas("Cajas", Val(txtcodcaja))
        'txtcuenta = ObtenerDescripcionCuentas("Cuentas", Val(txtcodcuenta))
        'txtefectivo = ObtenerImporte("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        'txtimpcheques = ObtenerTotalCheques("Movicaja", rsmov!NroDoc, rsmov!CODPR)
        
        CargoCheques
    Else
        'cmbFormaPago = BuscoDato("FormasPago", rsmov!FormadePago)
        'txtfvto = rsmov!vencim
        'txtsaldo = rsmov!saldo
    End If
    
Exit Sub
UfaCarga:
    MsgBox "Error cargando registro", vbCritical
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
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub dtFecha_LostFocus()
AnioMes
End Sub

Private Function AnioMes()
    txtmes = Month(dtFecha)
    txtanio = Year(dtFecha)
End Function



Private Sub Form_Load()
    comboSql cboIva, "select descripcion, codigo from ivas order by codigo"
    fraMesImputacion.Visible = VerParametro(BS_ComprasConMesImputacion)
    UpROV.ini "select descripcion from prov where activo = 1 and codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Descripcion               ] from prov where categ=3 and activo = 1 order by codigo ", False
    TabDetalle.Tab = 0

    revisoCdoCtaCte
    HabilitoControles False
    EsBusqueda = False
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


Function ObtenerTotalCheques(tabla As String, nDoc As Long, prov As Long) As Double
    'Dim rs As New ADODB.Recordset
    
    Dim Total As Double
    'Dim sqlstrCC As String
    
    'sqlstrCC = "Select * from " + tabla + " where codprov = " & prov & " and nrodoc = " & nDoc & " and (tipo = 'P' or tipo = 'C') and activo=1"
    'rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    'While Not rs.EOF
    '    Total = Total + rs!importe
    '    rs.MoveNext
    'Wend
    'rs.Close
    'Set rs = Nothing
    
    Total = nSinNull(obtenerDeSQL("Select sum(importe) from " + tabla + " where codprov = " & prov & " and nrodoc = " & nDoc & " and (tipo = 'P' or tipo = 'C') and activo=1"))
    
    ObtenerTotalCheques = Total
    
End Function



Private Sub lblIdDoc_Click()
    frmAsientoManual.mostrar s2n(lblIDDOC)
End Sub

Private Sub optcontado_Click()
    revisoCdoCtaCte
End Sub
Private Sub optcontado_Validate(cancel As Boolean)
    revisoCdoCtaCte
End Sub
Private Sub optctacte_Click()
    revisoCdoCtaCte
End Sub
Private Sub optctacte_Validate(cancel As Boolean)
    revisoCdoCtaCte
End Sub

Private Sub revisoCdoCtaCte()
    TabDetalle.TabEnabled(1) = optcontado.Value
    CambioOptPago
End Sub

Private Sub TabDetalle_Click(PreviousTab As Integer)
Exit Sub
    Dim x As String, tiene_c As Long
    Dim auxIDDOC As Long, auxIDAsiento As Long, i As Long
    Dim rsAsiento As New ADODB.Recordset
    If TabDetalle.Tab = 1 Then
        uRetCompras.Calcular UpROV.codigo, 0, 0, dtFecha
        txtTotalRetPago = 0
    ElseIf TabDetalle.Tab = 2 Then
        If EsBusqueda Then
            auxIDDOC = s2n(lblIDDOC)
            auxIDAsiento = obtenerDeSQL("select idasiento from asientos where iddoc=" & auxIDDOC)
            rsAsiento.Open "select * from mayor where debe>0 and idasiento= " & auxIDAsiento & " order by idmayor desc", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            With rsAsiento
                If .EOF And .BOF Then
                Else
                    .MoveFirst
                    For i = 0 To .RecordCount - 1
                       uTipoCompra.agregar !Cuenta, !Debe, True
                       .MoveNext
                    Next
                End If
            End With
        Else
            With uTipoCompra
                .Borrar
                .Total_a_Imputar = Importe
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(txtVALOR_RESTA), True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(txtRECARGO_INT), True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(txtVALOR_SUMA), True
                
                .agregar CuentaParam(ID_Cuenta_G_ImpCre), s2n(txtLRT), True
                .agregar CuentaParam(ID_Cuenta_G_ImpDeb), s2n(txtValesAlim), True
                .agregar CuentaParam(ID_Cuenta_G_Sircreb), s2n(txtAportesOS), True
                .agregar CuentaParam(ID_Cuenta_G_Sellado), s2n(txtIVACOMPRAS), True
                .agregar CuentaParam(ID_Cuenta_G_MantCta), s2n(txtContSS), True
                .agregar CuentaParam(ID_Cuenta_G_MantCtaSueldos), s2n(txtAportesSS), True
                .agregar CuentaParam(ID_Cuenta_G_Chequera), s2n(txtContR), True
                .agregar CuentaParam(ID_Cuenta_G_Varios), s2n(txtContOS), True
                .agregar CuentaParam(ID_Cuenta_G_ImpPorSobreGiro), s2n(txtIVAVENTAS), True
                .agregar CuentaParam(ID_Cuenta_G_ValNoConformados), s2n(txtPOSICIONIVA), True
                .agregar CuentaParam(ID_Cuenta_G_PercIIBB), s2n(txtARRASTREMES), True
                
                'aca no va la cuenta del prov
                tiene_c = obtenerDeSQL("select tiene_cuenta from prov where codigo = " & UpROV.codigo)
                If tiene_c = 1 Then
                    .agregar obtenerDeSQL("select cuenta from prov where codigo = " & UpROV.codigo), s2n(txtIMP_DETERN), False
                End If
            End With
        End If
    End If
End Sub

Private Sub txtanio_GotFocus()
AnioMes
End Sub
Private Sub txtanio_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtcodcaja_GotFocus()
    If Trim$(txtcodcaja) = "" Then txtcodcaja = "1"
    PintoFocoActivo
End Sub
Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtcodcuenta_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

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

Private Sub txtEfectivo_GotFocus()
''    On Error Resume Next
''    If s2n(txtimporte) <> 0 Then
''        If s2n(txtimporte) <> s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(TXTVALOR_SUMA) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(TXTRECARGO_INT) + s2n(txtIBcapital) + s2n(txtIBprovincia)) Then
''            MsgBox "Los totales no concilian, hay una diferencia de: " & s2n(s2n(txtneto) + s2n(txtiva) + s2n(txtper3337) + s2n(TXTVALOR_SUMA) + s2n(txtexento) + s2n(txtreteniva) + s2n(txtimpint) + s2n(txtretengan) + s2n(txtper3431) + s2n(TXTRECARGO_INT) + s2n(txtIBcapital) + s2n(txtIBprovincia)) - s2n(txtimporte)
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
    KeyAscii = SoloNum(KeyAscii)
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

Private Function Importe() As Double
Dim tot As Double
Dim ladoA As Double
Dim ladoB As Double
Dim ladoC As Double
    ladoA = s2n(s2n(txtIMP_DETERN) + s2n(txtVALOR_SUMA) + s2n(txtRECARGO_INT) - s2n(txtVALOR_RESTA))
    txtPOSICIONIVA = s2n(txtIVAVENTAS) - s2n(txtIVACOMPRAS)
    ladoB = s2n(s2n(txtPOSICIONIVA) + s2n(txtARRASTREMES))
    ladoC = s2n(txtAportesSS) + s2n(txtContSS) + s2n(txtContR) + s2n(txtContOS) + s2n(txtAportesOS) + s2n(txtValesAlim) + s2n(txtLRT)
    tot = ladoA + ladoB + ladoC
    'tot = s2n(txtIMP_DETERN) + s2n(txtVALOR_RESTA) + s2n(txtRECARGO_INT) + s2n(txtVALOR_SUMA) + s2n(txtIVACOMPRAS) + s2n(txtaportesos) + s2n(txtlrt) + s2n(txtvalesalim) + s2n(txtcontss) + s2n(txtaportesss) + s2n(txtcontr) + s2n(txtcontos) + s2n(txtIVAVENTAS) + s2n(txtPOSICIONIVA) + s2n(txtARRASTREMES)
    Importe = n2r(tot)
    txtimporte = n2r(tot)
End Function

Private Sub txtcontr_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtcontos_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtlrt_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtvalesalim_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtIVAVENTAS_GotFocus()
frmPintoFoco Me
End Sub

Private Sub TXTVALOR_RESTA_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub TXTVALOR_RESTA_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtcontss_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtaportesss_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtARRASTREMES_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtIVACOMPRAS_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtIVACOMPRAS_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtlrt_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtvalesalim_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtaportesos_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtaportesos_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtcontss_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtaportesss_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtcontr_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtcontos_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtIVAVENTAS_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtPOSICIONIVA_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtPOSICIONIVA_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtARRASTREMES_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub TXTVALOR_RESTA_LostFocus()
    txtVALOR_RESTA = s2n(txtVALOR_RESTA)
    Importe
End Sub
Private Sub TXTRECARGO_INT_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub TXTRECARGO_INT_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub TXTRECARGO_INT_LostFocus()
    txtRECARGO_INT = s2n(txtRECARGO_INT)
    Importe
End Sub
Private Sub TXTVALOR_SUMA_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub TXTVALOR_SUMA_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub TXTVALOR_SUMA_LostFocus()
        txtVALOR_SUMA = s2n(txtVALOR_SUMA)
        Importe
End Sub
Private Sub txtmes_GotFocus()
AnioMes
End Sub
Private Sub txtmes_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub TXTIMP_DETERN_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub TXTIMP_DETERN_LostFocus()
    txtIMP_DETERN = s2n(txtIMP_DETERN)
    Importe
End Sub

Private Sub txtlrt_LostFocus()
    txtLRT = s2n(txtLRT)
    Importe
End Sub

Private Sub txtvalesalim_LostFocus()
    txtValesAlim = s2n(txtValesAlim)
    Importe
End Sub

Private Sub txtaportesos_LostFocus()
    txtAportesOS = s2n(txtAportesOS)
    Importe
End Sub

Private Sub txtIVACOMPRAS_LostFocus()
    txtIVACOMPRAS = s2n(txtIVACOMPRAS)
    Importe
End Sub

Private Sub txtcontss_LostFocus()
    txtContSS = s2n(txtContSS)
    Importe
End Sub

Private Sub txtaportesss_LostFocus()
    txtAportesSS = s2n(txtAportesSS)
    Importe
End Sub

Private Sub txtcontr_LostFocus()
    txtContR = s2n(txtContR)
    Importe
End Sub

Private Sub txtcontos_LostFocus()
    txtContOS = s2n(txtContOS)
    Importe
End Sub

Private Sub txtIVAVENTAS_LostFocus()
    txtIVAVENTAS = s2n(txtIVAVENTAS)
    Importe
End Sub

Private Sub txtPOSICIONIVA_LostFocus()
    txtPOSICIONIVA = s2n(txtPOSICIONIVA)
    Importe
End Sub

Private Sub txtARRASTREMES_LostFocus()
    txtARRASTREMES = s2n(txtARRASTREMES)
    Importe
End Sub

Private Sub txtNroIIBB_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtnumcompra_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtsaldo_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtSerie_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtserie_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtSuc_GotFocus()
    frmPintoFoco Me
End Sub
Private Sub txtsuc_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txttipocompra_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txttransf_GotFocus()
    If s2n(txttransf) = 0 Then txttransf = quefaltapagar()
    PintoFocoActivo
End Sub
Private Function quefaltapagar() As Double
    quefaltapagar = s2n(Importe() - uRetCompras.TotalRet - s2n(txtefectivo) - s2n(txtimpcheques) - s2n(txttransf))
End Function

Private Sub txttransf_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
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


Private Sub uCheques_cambio()
    txtimpcheques = uCheques.Total
End Sub


Private Sub uNumDoc_LostFocus()
'    Dim resu As String
'    resu = ExisteFacCompra(uProv.codigo, uNumDoc.suc, uNumDoc.num)
'    If resu > "" Then
'        che "Documento existente " & resu
'    End If
    ExisteDocBancoMSG UpROV.codigo, uNumDoc.suc, uNumDoc.num
End Sub

Private Sub uRetCompras_LostFocus()
    txtTotalRetPago = uRetCompras.TotalRet
End Sub


Private Sub uProv_cambio(codigo As Variant)
    Dim tmp As Variant, tmpTIVA
    If UpROV.codigo = 0 Then
        CUIT.Text = ""
        uNumDoc.clear
        CUIT.TabStop = True
        txtNroIIBB = ""
        txtNroIIBB.TabStop = True
    Else
        tmp = obtenerDeSQL("select cuit, tipoiva, pago, tipoCom, suc, numiibb from Prov where activo = 1 and codigo = " & codigo)
        
        If IsEmpty(tmp) Then
            MsgBox "No se pudieron cargar datos de proveedor", vbCritical
        Else
            CUIT.Text = sSinNull(tmp(0))
            cboIva.ListIndex = BuscarEnCombo(cboIva, nSinNull(tmp(1)))
            uNumDoc.suc = nSinNull(tmp(4))
            seteoLetra
            txtNroIIBB = sSinNull(tmp(5)) '*** *** ***

        End If
    End If
    AnioMes
    cmdnuevo_Click
End Sub

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
        If dtFecha.enabled = True Then
            dtFecha.SetFocus
        End If
        UpROV.EditaDescripcion = True
    Else
        UpROV.EditaDescripcion = False
        'If uProv.codigo = 0 Then uProv.codigo = 0 ' ridiculo?  setea descripcion = ""  ahora puede ser .clear
    End If
End Sub

Private Sub uTipoCompra_GotFocus()
    'uTipoCompra.Total_a_Imputar = s2n(s2n(txtimporte) - sumaTxtIvas() - 0) ' - uRetGan.Total
    
    uTipoCompra.Total_a_Imputar = Importe() 's2n(txtneto) + s2n(txtexento)
    
End Sub

Private Function sumaTxtIvas() As Double
    sumaTxtIvas = s2n(txtVALOR_RESTA) + s2n(txtRECARGO_INT) + s2n(txtVALOR_SUMA)
End Function

'Private Function ExisteFAC(prov, suc, Nro) As String
'    Dim resu
'    Dim whe  As String
'
'    If prov = 0 Then Exit Function
'
'    whe = " where  (TIPODOC = 'N/C' OR TIPODOC = 'N/D' OR TIPODOC = 'FAC') and codpr = " & prov & " and suc = " & suc & " and NroDoc = " & Nro
'
'    resu = obtenerDeSQL("select fecha from compras " & whe)
'    If Not IsEmpty(resu) Then
'        ExisteFAC = CStr(resu)
'        Exit Function
'    End If
'
'    resu = obtenerDeSQL("select fecha from transcom " & whe)
'    If Not IsEmpty(resu) Then
'        ExisteFAC = CStr(resu)
'        Exit Function
'    End If
'End Function

Private Sub seteoLetra()
    Dim tmpTIVA
    
    tmpTIVA = sSinNull(obtenerDeSQL("select letraprov from ivas where codigo = " & ComboCodigo(cboIva)))     's2n(txttipoiva))
    If IsNull(tmpTIVA) Or IsEmpty(tmpTIVA) Then
        MsgBox "Verificar Condición de IVA"
    Else
        uNumDoc.letra = tmpTIVA
        tabstopXletra tmpTIVA
    End If
End Sub
Private Sub tabstopXletra(letra)
    Dim esa As Boolean
    esa = (letra <> "C")
    
    txtIMP_DETERN.TabStop = esa
    txtVALOR_RESTA.TabStop = esa
    txtRECARGO_INT.TabStop = esa
    txtVALOR_SUMA.TabStop = esa
End Sub

Public Function AGastos(gOPE As String, gFECHA As Date, gAnio As Long, gMes As Long, gCODBANCO As Long, gRAZON As String, gCUIT As String, gTIPOIVA As Long, gTIPODOC As String, gNRODOC As Long, gTOTAL As Double, gSUC As Long, gCUENTA As String, gCONTADO As Long, gIMPDETERN As Double, _
                            gLIBRE06 As Double, gPORCENIVA As Double, gRECARGOINT As Double, gVALORRESTA As Double, gVALORSUMA As Double, gIVACOMPRAS As Double, gVALESALIM As Double, gLRT As Double, gARRASTRE As Double, gAPORTEOS As Double, gIVAVENTAS As Double, gIBC As Double, gIBP As Double, _
                            gCONTSS As Double, gAPORTESS As Double, gCONTR As Double, gCONTOS As Double, gPOSICIONIVA As Double, gRETGAN As Double, gRETGANP As Double, gIBPAGO As Double, gVTO As Date, gFORMAPAGO As Long, gMONEDA As Long, gCOTIZACION As Double, _
                            gIDDOC As Long, gUSUALTA As Long, gLETRA As String, gNROIIBB As String, gTIPORETGAN As Long, gTIPORETIIBB As Long)

Dim cad As String

cad = "INSERT INTO GASTOSBOLETAS (nrodoc,FECHA,ANOIMP,MESIMP,CODPR,RAZONSOCIAL,CUIT,TIPOIVA,TIPODOC,TOTAL,SUC,CUENTA,CONTADO,IMP_DETERN,SALDO,LIBRE06,PORCENIVA,RECARGO_INT,VALOR_RESTA,VALOR_SUMA,IVACOMPRAS,VALESALIM,LRT,ARRASTREIVA,APORTESOS,IVAVENTAS,IBCAPITAL,IBPROVINCIA,CONTRIBUCIONSS,APORTESSS,CONTRIBUCIONR, " _
        & "CONTRIBUCIONOS,POSICIONIVA,RETGAN,RETGANPAGO,IBPAGO,VTO_CO,FORMADEPAGO,MONEDA,COTIZACION,IDDOC,FECHA_ALTA,USUARIO_ALTA,ACTIVO,LETRA,NROIIBB,TIPORETGAN,TIPORETIIBB) VALUES (" _
        & gNRODOC & "," & ssFecha(gFECHA) & "," & gAnio & "," & gMes & "," & gCODBANCO & "," & ssTexto(gRAZON) & "," & ssTexto(gCUIT) & "," & gTIPOIVA & "," & ssTexto(gTIPODOC) & "," & x2s(gTOTAL) & "," & gSUC & "," & ssTexto(gCUENTA) & "," & gCONTADO & "," & x2s(gIMPDETERN) & "," & x2s(gTOTAL) & "," _
        & x2s(gLIBRE06) & "," & x2s(gPORCENIVA) & "," & x2s(gRECARGOINT) & "," & x2s(gVALORRESTA) & "," & x2s(gVALORSUMA) & "," & x2s(gIVACOMPRAS) & "," & x2s(gVALESALIM) & "," & x2s(gLRT) & "," & x2s(gARRASTRE) & "," & x2s(gAPORTEOS) & "," & x2s(gIVAVENTAS) & "," & x2s(gIBC) & "," & x2s(gIBP) & "," _
        & x2s(gCONTSS) & "," & x2s(gAPORTESS) & "," & x2s(gCONTR) & "," & x2s(gCONTOS) & "," & x2s(gPOSICIONIVA) & "," & x2s(gRETGAN) & "," & x2s(gRETGANP) & "," & x2s(gIBPAGO) & "," & ssFecha(gVTO) & "," & gFORMAPAGO & "," & gMONEDA & "," & x2s(gCOTIZACION) & "," _
        & gIDDOC & "," & ssFecha(Date) & "," & gUSUALTA & ",1," & ssTexto(gLETRA) & "," & ssTexto(gNROIIBB) & "," & gTIPORETGAN & "," & gTIPORETIIBB & ")"
        
DataEnvironment1.Sistema.Execute cad


End Function



