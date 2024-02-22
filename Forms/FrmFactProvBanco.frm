VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmFactProvBanco 
   Caption         =   "Gastos Bancarios"
   ClientHeight    =   7770
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "FrmFactProvBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboEjercicio 
      Height          =   315
      Left            =   9120
      TabIndex        =   95
      Text            =   "Ejercicio"
      Top             =   1320
      Width           =   990
   End
   Begin Gestion.uCtaBanco uCtaBanco 
      Height          =   330
      Left            =   1770
      TabIndex        =   94
      Top             =   1305
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   582
   End
   Begin VB.TextBox txtNroIIBB 
      Height          =   285
      Left            =   8655
      TabIndex        =   24
      Top             =   900
      Width           =   1425
   End
   Begin VB.ComboBox cboIva 
      Height          =   315
      Left            =   1215
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   870
      Width           =   2190
   End
   Begin Gestion.uNumDoc uNumDoc 
      Height          =   300
      Left            =   5085
      TabIndex        =   23
      Top             =   885
      Width           =   2625
      _ExtentX        =   4736
      _ExtentY        =   529
   End
   Begin VB.Frame fraMesImputacion 
      Height          =   495
      Left            =   5145
      TabIndex        =   83
      Top             =   -30
      Width           =   3330
      Begin VB.TextBox txtanio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   555
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   150
         Width           =   735
      End
      Begin VB.ComboBox txtmes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmFactProvBanco.frx":08CA
         Left            =   1830
         List            =   "FrmFactProvBanco.frx":08F2
         Style           =   2  'Dropdown List
         TabIndex        =   84
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
         TabIndex        =   87
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
         TabIndex        =   86
         Top             =   150
         Width           =   615
      End
   End
   Begin Gestion.ucFecha uFeHa 
      Height          =   270
      Left            =   2355
      TabIndex        =   81
      Top             =   7350
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   476
      FechaInit       =   4
   End
   Begin Gestion.ucFecha uFeDe 
      Height          =   270
      Left            =   1515
      TabIndex        =   80
      Top             =   7350
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
      TabIndex        =   77
      Top             =   7350
      Width           =   975
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   5505
      Left            =   45
      TabIndex        =   25
      Top             =   1755
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9710
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Gasto"
      TabPicture(0)   =   "FrmFactProvBanco.frx":091D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraFactura"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sin uso"
      TabPicture(1)   =   "FrmFactProvBanco.frx":0939
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraContado"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Imputaciones Contables"
      TabPicture(2)   =   "FrmFactProvBanco.frx":0955
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "uTipoCompra"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4590
         Left            =   75
         TabIndex        =   60
         Top             =   390
         Width           =   11475
         Begin VB.TextBox txtImpDebitos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7125
            TabIndex        =   14
            Top             =   2580
            Width           =   1215
         End
         Begin VB.TextBox txtImpCreditos 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7125
            TabIndex        =   15
            Top             =   2940
            Width           =   1215
         End
         Begin VB.TextBox txtSircreb 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7110
            TabIndex        =   13
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox txtMantCta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   5
            Top             =   2115
            Width           =   1215
         End
         Begin VB.TextBox txtIva21 
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
            Left            =   7155
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   3690
            Width           =   1215
         End
         Begin VB.TextBox txtGastoNeto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   0
            Top             =   345
            Width           =   1230
         End
         Begin VB.TextBox txtImpxGiro 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7110
            TabIndex        =   10
            Top             =   825
            Width           =   1215
         End
         Begin VB.TextBox txtIva27 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   2
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtSellado 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7110
            TabIndex        =   9
            Top             =   450
            Width           =   1215
         End
         Begin VB.TextBox txtPercIB 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7125
            TabIndex        =   12
            Top             =   1545
            Width           =   1215
         End
         Begin VB.TextBox txtGVarios 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2685
            TabIndex        =   8
            Top             =   3345
            Width           =   1215
         End
         Begin VB.TextBox txtMantCtaS 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2670
            TabIndex        =   6
            Top             =   2460
            Width           =   1215
         End
         Begin VB.TextBox txtValNoConf 
            Enabled         =   0   'False
            Height          =   285
            Left            =   7110
            TabIndex        =   11
            Top             =   1185
            Width           =   1215
         End
         Begin VB.TextBox txtIva10 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2655
            TabIndex        =   4
            Top             =   1545
            Width           =   1215
         End
         Begin VB.TextBox txtGChequera 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2685
            TabIndex        =   7
            Top             =   3015
            Width           =   1215
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Imp Debitos Bancarios"
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
            Left            =   4605
            TabIndex        =   92
            Top             =   2595
            Width           =   2385
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Imp Creditos Bancarios"
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
            Left            =   4845
            TabIndex        =   91
            Top             =   2955
            Width           =   2145
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sircreb Alic B"
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
            Left            =   5790
            TabIndex        =   90
            Top             =   1950
            Width           =   1500
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mant de Cuenta"
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
            Left            =   960
            TabIndex        =   73
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
            Left            =   930
            TabIndex        =   72
            Top             =   840
            Width           =   1605
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Neto Del Gasto"
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
            TabIndex        =   71
            Top             =   345
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
            TabIndex        =   70
            Top             =   3735
            Width           =   1605
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Int Pagados por Acuerdo"
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
            Left            =   4695
            TabIndex        =   69
            Top             =   855
            Width           =   2730
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
            Left            =   1290
            TabIndex        =   68
            Top             =   1185
            Width           =   1215
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comisiones por Gestion de Cheq"
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
            Left            =   4035
            TabIndex        =   67
            Top             =   480
            Width           =   3045
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Percepcion IB"
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
            Left            =   4830
            TabIndex        =   66
            Top             =   1560
            Width           =   2145
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Varios"
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
            Left            =   765
            TabIndex        =   65
            Top             =   3390
            Width           =   1815
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Impuesto a los Sellos"
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
            Left            =   135
            TabIndex        =   64
            Top             =   2490
            Width           =   2430
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Comisiones por Transferencias"
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
            Left            =   3930
            TabIndex        =   63
            Top             =   1200
            Width           =   3045
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
            Left            =   1695
            TabIndex        =   62
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Gastos Chequera"
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
            Left            =   900
            TabIndex        =   61
            Top             =   3030
            Width           =   1650
         End
      End
      Begin VB.Frame fraContado 
         BorderStyle     =   0  'None
         Height          =   4650
         Left            =   -74955
         TabIndex        =   54
         Top             =   435
         Width           =   11415
         Begin Gestion.ucRetCompras uRetCompras 
            Height          =   705
            Left            =   1725
            TabIndex        =   29
            Top             =   90
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   1244
         End
         Begin VB.TextBox txtTotalRetPago 
            Height          =   330
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox txtimpcheques 
            Enabled         =   0   'False
            Height          =   300
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1605
            Width           =   1320
         End
         Begin VB.TextBox txtefectivo 
            Enabled         =   0   'False
            Height          =   330
            Left            =   45
            TabIndex        =   30
            Top             =   825
            Width           =   1335
         End
         Begin VB.TextBox txttransf 
            Enabled         =   0   'False
            Height          =   285
            Left            =   75
            TabIndex        =   36
            Top             =   4305
            Width           =   1215
         End
         Begin VB.TextBox txtcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4425
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtcodcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2175
            TabIndex        =   38
            Top             =   4335
            Width           =   1215
         End
         Begin VB.TextBox txtcodcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2355
            TabIndex        =   31
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
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   855
            Width           =   855
         End
         Begin VB.TextBox txtcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4290
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   855
            Width           =   2535
         End
         Begin Gestion.ucCheques uCheques 
            Height          =   2700
            Left            =   1440
            TabIndex        =   34
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
            TabIndex        =   76
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
            Top             =   870
            Width           =   855
         End
      End
      Begin Gestion.ucTipoCompra uTipoCompra 
         Height          =   3870
         Left            =   -74925
         TabIndex        =   37
         Top             =   645
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   6826
      End
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   8775
      TabIndex        =   53
      Top             =   7785
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6255
      TabIndex        =   52
      Tag             =   "1"
      Top             =   7860
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtnumcompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6900
      TabIndex        =   51
      Top             =   7920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttipocompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5355
      TabIndex        =   50
      Top             =   7815
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtfechacompra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7875
      TabIndex        =   49
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
      TabIndex        =   35
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
      TabIndex        =   41
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
      TabIndex        =   43
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
      TabIndex        =   3
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
      TabIndex        =   45
      Top             =   7290
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
      TabIndex        =   42
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
      TabIndex        =   44
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
      Format          =   186122241
      CurrentDate     =   37934
   End
   Begin Gestion.ucCoDe uProv 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   495
      Width           =   6510
      _ExtentX        =   9657
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCuit Cuit 
      Height          =   285
      Left            =   8685
      TabIndex        =   21
      Top             =   525
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
   End
   Begin VB.Label Label34 
      Caption         =   "Ejercicio"
      Height          =   255
      Left            =   10200
      TabIndex        =   96
      Top             =   1380
      Width           =   735
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
      TabIndex        =   93
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
      TabIndex        =   89
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
      TabIndex        =   88
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
      TabIndex        =   82
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   75
      Top             =   525
      Width           =   570
   End
   Begin VB.Label Label3 
      Caption         =   "Banco Prov"
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
      TabIndex        =   74
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
      TabIndex        =   48
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
      TabIndex        =   47
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
      TabIndex        =   46
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmFactProvBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IVA_21 = 0.21
Private Const IVA_27 = 0.27
Private Const IVA_105 = 0.105
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
        tabla = "GastosBancarios"
        titu = "Gastos Bancarios"
    'End If
    ssql = " select Fecha as [Fecha ], tipoDoc as [Doc], NroDoc as [ Numero ], total as  [ Importe ], codbanco as [ Banco Prov], razonsocialbanco as [ Razon social                           ], iddoc " & _
        " from " & tabla & " left join prov on codbanco = prov.codigo  " & _
        " where " & tabla & ".activo = 1 and fecha " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & _
        " order by fecha desc "
    
    re = frmBuscar.MostrarSql(ssql, , titu) ', , , , " " , "Anulada")
    If re > "" Then
'        LimpioControles
        Call Habilitobotones(True, True, True, False, True, True)
        fech = CDate(frmBuscar.resultado(1))
        txttipocompra = frmBuscar.resultado(2)
        txtnumcompra = frmBuscar.resultado(3)

        rsmov.Open "select * from " & tabla & " where fecha = " & ssFecha(fech) & " and tipodoc = '" & txttipocompra & "' and iddoc = " & frmBuscar.resultado(7) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        CargoRegistro
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
            cad = "update gastosbancarios set activo=0 where iddoc=" & midDoc
            DataEnvironment1.Sistema.Execute cad
            cad = "update movibanc set activo=0 where iddoc=" & midDoc
            DataEnvironment1.Sistema.Execute cad
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
    
    If uCtaBanco.codigo = 0 Then
        MsgBox "No se indico la cuenta bancaria", vbCritical
        Exit Sub
    End If
    
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

    If ExisteDocBancoMSG(UpROV.codigo, uNumDoc.suc, uNumDoc.num) Then
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
    dife = s2n(s2n(txtimporte) - (s2n(s2n(txtGastoNeto) + s2n(txtIva21) + s2n(txtIva10) + s2n(txtIva27) + s2n(txtSellado) + s2n(txtSircreb) + s2n(txtImpCreditos) + s2n(txtImpDebitos) + s2n(txtMantCta) + s2n(txtMantCtaS) + s2n(txtGChequera) + s2n(txtGVarios) + s2n(txtImpxGiro) + s2n(txtValNoConf) + s2n(txtPercIB))))
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
        .nuevo "Gto " & UpROV.DESCRIPCION, dtFecha, TIPODOC_FAC_BANCOGASTO
        'DEBE
        For i = 1 To uTipoCompra.rows
            .AgregarItem uTipoCompra.imCuenta(i), uTipoCompra.imMonto(i), 0     ', TextoAsientoComprobante
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
    
    
    iddoc = NuevoDocumento(TIPODOC_FAC_BANCOGASTO, NroDoc, UpROV.codigo, NroPago, NroCertifGan, NroCertifIIBB)
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
            Else            ' *********  alta FC CUENTA CORRIENTE ******* OSEA GASTOS BANCARIOS
                Dim gPeriodo As String
                gPeriodo = qMes(dtFecha)
                
                If txtGastoNeto = "0" Then
                ElseIf s2n(txtGastoNeto) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos " & gPeriodo, _
                    dtFecha, "G", s2n(txtGastoNeto), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtIva21 = "0" Then
                ElseIf s2n(txtIva21) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 21% " & gPeriodo, _
                    dtFecha, "G", s2n(txtIva21), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtIva10 = "0" Then
                ElseIf s2n(txtIva10) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 10.5% " & gPeriodo, _
                    dtFecha, "G", s2n(txtIva10), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtIva27 = "0" Then
                ElseIf s2n(txtIva27) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Iva 27% " & gPeriodo, _
                    dtFecha, "G", s2n(txtIva27), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                'Comisiones por Gestion de Cheques
                If txtSellado = "0" Then
                ElseIf s2n(txtSellado) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Comisiones por Gestion de Cheq " & gPeriodo, _
                    dtFecha, "G", s2n(txtSellado), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtSircreb = "0" Then
                ElseIf s2n(txtSircreb) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Sircreb " & gPeriodo, _
                    dtFecha, "G", s2n(txtSircreb), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtImpCreditos = "0" Then
                ElseIf s2n(txtImpCreditos) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Cred " & gPeriodo, _
                    dtFecha, "G", s2n(txtImpCreditos), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtImpDebitos = "0" Then
                ElseIf s2n(txtImpDebitos) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Deb " & gPeriodo, _
                    dtFecha, "G", s2n(txtImpDebitos), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtMantCta = "0" Then
                ElseIf s2n(txtMantCta) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Mant. Cuenta " & gPeriodo, _
                    dtFecha, "G", s2n(txtMantCta), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtMantCtaS = "0" Then
                ElseIf s2n(txtMantCtaS) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Mant. Cuenta Sueldos " & gPeriodo, _
                    dtFecha, "G", s2n(txtMantCtaS), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtGChequera = "0" Then
                ElseIf s2n(txtGChequera) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos por Chequera " & gPeriodo, _
                    dtFecha, "G", s2n(txtGChequera), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtGVarios = "0" Then
                ElseIf s2n(txtGVarios) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Gastos Varios " & gPeriodo, _
                    dtFecha, "G", s2n(txtGVarios), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtImpxGiro = "0" Then
                ElseIf s2n(txtImpxGiro) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Imp por Sobre Giro " & gPeriodo, _
                    dtFecha, "G", s2n(txtImpxGiro), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                'Comisiones por Transferencias
                If txtValNoConf = "0" Then
                ElseIf s2n(txtValNoConf) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Comisiones por Transferencias " & gPeriodo, _
                    dtFecha, "G", s2n(txtValNoConf), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                If txtPercIB = "0" Then
                ElseIf s2n(txtPercIB) > 0 Then
                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", "Percepcion de IIBB " & gPeriodo, _
                    dtFecha, "G", s2n(txtPercIB), mMovBanc, midDoc, Date, UsuarioActual, 1
                End If
                
                AsientoCompra.AgregarItem uCtaBanco.CuentaContable, 0, s2n(txtimporte), TextoAsientoComprobante
                
                a = uNumDoc.num
                AGastos "A", dtFecha, s2n(txtanio), s2n(txtmes), UpROV.codigo, UpROV.DESCRIPCION, CUIT.Text, ComboCodigo(cboIva), TIPODOC_FAC_BANCOGASTO, _
                            uNumDoc.num, Importe, uNumDoc.suc, uCtaBanco.codigo, 1, s2n(txtGastoNeto), 0, 0, s2n(txtIva10), s2n(txtIva21), s2n(txtIva27), _
                             s2n(txtSellado), s2n(txtImpDebitos), s2n(txtImpCreditos), s2n(txtPercIB), s2n(txtSircreb), s2n(txtImpxGiro), 0, 0, s2n(txtMantCta), _
                             s2n(txtMantCtaS), s2n(txtGChequera), s2n(txtGVarios), s2n(txtValNoConf), 0, 0, 0, Date, 0, 0, 0, midDoc, UsuarioActual, uNumDoc.letra, txtNroIIBB, 0, 0
            End If
        End If
      
        sAssert = "15) ASIENTOS"
        AsientoCompra.Grabar iddoc, , leerEjercicioId(cboEjercicio)

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
    UpROV.codigo = 0
    FrmBorrarTxt Me
    uNumDoc.clear
    lblIDDOC = ""
    txtcodcaja = 1
    midDoc = 0
    CUIT.Text = ""
    dtFecha = Date

    optcontado = False
    optctacte = True

    txtGastoNeto = "0"
    txtIva21 = "0"
    txtIva10 = "0"
    txtIva27 = "0"
    txtimporte = "0"
    txtSellado = "0"
    txtSircreb = "0"
    txtImpCreditos = "0"
    txtImpDebitos = "0"
    txtMantCta = "0"
    txtMantCtaS = "0"
    txtGChequera = "0"
    txtGVarios = "0"
    txtImpxGiro = "0"
    txtValNoConf = "0"
    txtPercIB = "0"
    
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
    
    txtGastoNeto.enabled = habilito
    txtIva21.enabled = habilito
    txtIva10.enabled = habilito
    txtIva27.enabled = habilito
    txtimporte.enabled = habilito
    txtSellado.enabled = habilito
    txtSircreb.enabled = habilito
    txtImpCreditos.enabled = habilito
    txtImpDebitos.enabled = habilito
    txtMantCta.enabled = habilito
    txtMantCtaS.enabled = habilito
    txtGChequera.enabled = habilito
    txtGVarios.enabled = habilito
    txtImpxGiro.enabled = habilito
    txtValNoConf.enabled = habilito
    txtPercIB.enabled = habilito
    
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
    UpROV.codigo = rsmov!CODBANCO
    UpROV.DESCRIPCION = rsmov!razonsocialbanco
    
    uNumDoc.num = rsmov!NroDoc
    
    CUIT.Text = rsmov!Cuitbanco

    dtFecha = rsmov!Fecha
    uCtaBanco.codigo = rsmov!Cuenta
    txtGastoNeto = rsmov!Neto
    txtIva21 = rsmov!IVA_21
    txtIva10 = rsmov!iva_10
    txtIva27 = rsmov!IVA_27
    txtSircreb = rsmov!SIRCREB
    txtimporte = rsmov!Total
    txtSellado = rsmov!SELLADO
    txtImpCreditos = rsmov!IMPCRE
    txtImpDebitos = rsmov!IMPDEB
    txtMantCta = rsmov!MANTCTA
    txtMantCtaS = rsmov!MANTCTASueldos
    txtGChequera = rsmov!gastoschqra
    txtGVarios = rsmov!gastosvarios
    txtImpxGiro = rsmov!INTXGIRO
    txtValNoConf = rsmov!valnoconfor
    txtPercIB = rsmov!PERIIBB
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
    UpROV.ini "select descripcion from prov where activo = 1 and codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Descripcion               ] from prov where categ=2 and activo = 1 order by codigo ", False
    TabDetalle.Tab = 0

    revisoCdoCtaCte
    HabilitoControles False
    EsBusqueda = False
    
    Dim EjerA As New ADODB.Recordset
    EjerA.Open "select * from ejercicio", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    EjerA.MoveFirst
    While Not EjerA.EOF
        cboEjercicio.AddItem EjerA!denominacion 'EjerA!idejercicio
        EjerA.MoveNext
    Wend
    cboEjercicio = leerEjercicioDenominacion() ' mIdEjercicioActivo
    If UsuarioActual() <> 19 Then
        cboEjercicio.Visible = False
        Label34.Visible = False
    End If

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
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(txtIva21), True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(txtIva10), True
                .agregar CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(txtIva27), True
                
                .agregar CuentaParam(ID_Cuenta_G_ImpCre), s2n(txtImpCreditos), True
                .agregar CuentaParam(ID_Cuenta_G_ImpDeb), s2n(txtImpDebitos), True
                .agregar CuentaParam(ID_Cuenta_G_Sircreb), s2n(txtSircreb), True
                .agregar CuentaParam(ID_Cuenta_G_Sellado), s2n(txtSellado), True
                .agregar CuentaParam(ID_Cuenta_G_MantCta), s2n(txtMantCta), True
                .agregar CuentaParam(ID_Cuenta_G_MantCtaSueldos), s2n(txtMantCtaS), True
                .agregar CuentaParam(ID_Cuenta_G_Chequera), s2n(txtGChequera), True
                .agregar CuentaParam(ID_Cuenta_G_Varios), s2n(txtGVarios), True
                .agregar CuentaParam(ID_Cuenta_G_ImpPorSobreGiro), s2n(txtImpxGiro), True
                .agregar CuentaParam(ID_Cuenta_G_ValNoConformados), s2n(txtValNoConf), True
                .agregar CuentaParam(ID_Cuenta_G_PercIIBB), s2n(txtPercIB), True
                
                'aca no va la cuenta del prov
                tiene_c = obtenerDeSQL("select tiene_cuenta from prov where codigo = " & UpROV.codigo)
                If tiene_c = 1 Then
                    .agregar obtenerDeSQL("select cuenta from prov where codigo = " & UpROV.codigo), s2n(txtGastoNeto), False
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
    tot = s2n(txtGastoNeto) + s2n(txtIva21) + s2n(txtIva10) + s2n(txtIva27) + s2n(txtSellado) + s2n(txtSircreb) + s2n(txtImpCreditos) + s2n(txtImpDebitos) + s2n(txtMantCta) + s2n(txtMantCtaS) + s2n(txtGChequera) + s2n(txtGVarios) + s2n(txtImpxGiro) + s2n(txtValNoConf) + s2n(txtPercIB)
    Importe = n2r(tot)
    txtimporte = n2r(tot)
End Function

Private Sub txtGChequera_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtGVarios_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtImpCreditos_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtImpDebitos_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtImpxGiro_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtiva21_GotFocus()
    txtIva21 = s2n(s2n(txtGastoNeto) * IVA_21)
    frmPintoFoco Me
End Sub
Private Sub txtiva21_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtMantCta_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtMantCtaS_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtPercIB_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtSellado_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtsellado_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtimpcreditos_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtimpdebitos_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtSircreb_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtsircreb_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtmantcta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtmantctas_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtgchequera_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtgvarios_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtimpxgiro_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtValNoConf_GotFocus()
frmPintoFoco Me
End Sub

Private Sub txtvalnoconf_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtpercib_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtiva21_LostFocus()
    txtIva21 = s2n(txtIva21)
    Importe
End Sub
Private Sub txtiva10_GotFocus()
    If s2n(txtIva21) + s2n(txtIva27) = 0 Then txtIva10 = s2n(s2n(txtGastoNeto) * IVA_105)
    frmPintoFoco Me
End Sub
Private Sub txtiva10_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtiva10_LostFocus()
    txtIva10 = s2n(txtIva10)
    Importe
End Sub
Private Sub txtiva27_GotFocus()
    If s2n(txtIva21) = 0 Then txtIva27 = s2n(s2n(txtGastoNeto) * IVA_27)
    frmPintoFoco Me
End Sub
Private Sub txtiva27_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtiva27_LostFocus()
        txtIva27 = s2n(txtIva27)
        Importe
End Sub
Private Sub txtmes_GotFocus()
AnioMes
End Sub
Private Sub txtmes_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtgastoneto_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNum(KeyAscii)
End Sub
Private Sub txtgastoneto_LostFocus()
    txtGastoNeto = s2n(txtGastoNeto)
    Importe
End Sub

Private Sub txtimpcreditos_LostFocus()
    txtImpCreditos = s2n(txtImpCreditos)
    Importe
End Sub

Private Sub txtimpdebitos_LostFocus()
    txtImpDebitos = s2n(txtImpDebitos)
    Importe
End Sub

Private Sub txtsircreb_LostFocus()
    txtSircreb = s2n(txtSircreb)
    Importe
End Sub

Private Sub txtsellado_LostFocus()
    txtSellado = s2n(txtSellado)
    Importe
End Sub

Private Sub txtmantcta_LostFocus()
    txtMantCta = s2n(txtMantCta)
    Importe
End Sub

Private Sub txtmantctas_LostFocus()
    txtMantCtaS = s2n(txtMantCtaS)
    Importe
End Sub

Private Sub txtgchequera_LostFocus()
    txtGChequera = s2n(txtGChequera)
    Importe
End Sub

Private Sub txtgvarios_LostFocus()
    txtGVarios = s2n(txtGVarios)
    Importe
End Sub

Private Sub txtimpxgiro_LostFocus()
    txtImpxGiro = s2n(txtImpxGiro)
    Importe
End Sub

Private Sub txtvalnoconf_LostFocus()
    txtValNoConf = s2n(txtValNoConf)
    Importe
End Sub

Private Sub txtpercib_LostFocus()
    txtPercIB = s2n(txtPercIB)
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
        If UpROV.codigo = 0 Then UpROV.codigo = 0 ' ridiculo?  setea descripcion = ""  ahora puede ser .clear
    End If
End Sub

Private Sub uTipoCompra_GotFocus()
    'uTipoCompra.Total_a_Imputar = s2n(s2n(txtimporte) - sumaTxtIvas() - 0) ' - uRetGan.Total
    
    uTipoCompra.Total_a_Imputar = Importe() 's2n(txtneto) + s2n(txtexento)
    
End Sub

Private Function sumaTxtIvas() As Double
    sumaTxtIvas = s2n(txtIva21) + s2n(txtIva10) + s2n(txtIva27)
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
    
    txtGastoNeto.TabStop = esa
    txtIva21.TabStop = esa
    txtIva10.TabStop = esa
    txtIva27.TabStop = esa
End Sub

Public Function AGastos(gOPE As String, gFECHA As Date, gAnio As Long, gMes As Long, gCODBANCO As Long, gRAZON As String, gCUIT As String, gTIPOIVA As Long, gTIPODOC As String, gNRODOC As Long, gTOTAL As Double, gSUC As Long, gCUENTA As String, gCONTADO As Long, gNETO As Double, _
                            gEXENTO As Double, gPORCENIVA As Double, gIVA10 As Double, gIVA21 As Double, gIVA27 As Double, gSELLADO As Double, gIMPDEB As Double, gIMPCRE As Double, gPERIIBB As Double, gSIRCREB As Double, gINTXGIRO As Double, gIBC As Double, gIBP As Double, _
                            gMANTCTA As Double, gMANTCTAS As Double, gGASTOSCHE As Double, gGASTOSVS As Double, gVALNOCONF As Double, gRETGAN As Double, gRETGANP As Double, gIBPAGO As Double, gVTO As Date, gFORMAPAGO As Long, gMONEDA As Long, gCOTIZACION As Double, _
                            gIDDOC As Long, gUSUALTA As Long, gLETRA As String, gNROIIBB As String, gTIPORETGAN As Long, gTIPORETIIBB As Long)

Dim cad As String

cad = "INSERT INTO GASTOSBANCARIOS (nrodoc,FECHA,ANOIMP,MESIMP,CODBANCO,RAZONSOCIALBANCO,CUITBANCO,TIPOIVA,TIPODOC,TOTAL,SUC,CUENTA,CONTADO,NETO,EXENTO,PORCENIVA,IVA_10,IVA_21,IVA_27,SELLADO,IMPDEB,IMPCRE,PERIIBB,SIRCREB,INTXGIRO,IBCAPITAL,IBPROVINCIA,MANTCTA,MANTCTASUELDOS,GASTOSCHQRA, " _
        & "GASTOSVARIOS,VALNOCONFOR,RETGAN,RETGANPAGO,IBPAGO,VTO_CO,FORMADEPAGO,MONEDA,COTIZACION,IDDOC,FECHA_ALTA,USUARIO_ALTA,ACTIVO,LETRA,NROIIBB,TIPORETGAN,TIPORETIIBB) VALUES (" _
        & gNRODOC & "," & ssFecha(gFECHA) & "," & gAnio & "," & gMes & "," & gCODBANCO & "," & ssTexto(gRAZON) & "," & ssTexto(gCUIT) & "," & gTIPOIVA & "," & ssTexto(gTIPODOC) & "," & x2s(gTOTAL) & "," & gSUC & "," & ssTexto(gCUENTA) & "," & gCONTADO & "," & x2s(gNETO) & "," _
        & x2s(gEXENTO) & "," & x2s(gPORCENIVA) & "," & x2s(gIVA10) & "," & x2s(gIVA21) & "," & x2s(gIVA27) & "," & x2s(gSELLADO) & "," & x2s(gIMPDEB) & "," & x2s(gIMPCRE) & "," & x2s(gPERIIBB) & "," & x2s(gSIRCREB) & "," & x2s(gINTXGIRO) & "," & x2s(gIBC) & "," & x2s(gIBP) & "," _
        & x2s(gMANTCTA) & "," & x2s(gMANTCTAS) & "," & x2s(gGASTOSCHE) & "," & x2s(gGASTOSVS) & "," & x2s(gVALNOCONF) & "," & x2s(gRETGAN) & "," & x2s(gRETGANP) & "," & x2s(gIBPAGO) & "," & ssFecha(gVTO) & "," & gFORMAPAGO & "," & gMONEDA & "," & x2s(gCOTIZACION) & "," _
        & gIDDOC & "," & ssFecha(Date) & "," & gUSUALTA & ",1," & ssTexto(gLETRA) & "," & ssTexto(gNROIIBB) & "," & gTIPORETGAN & "," & gTIPORETIIBB & ")"
        
DataEnvironment1.Sistema.Execute cad


End Function


