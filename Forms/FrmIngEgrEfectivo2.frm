VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmIngEgrEfectivo 
   Caption         =   "Ingreso / Egreso de Cajas y Bancos"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9975
   Icon            =   "FrmIngEgrEfectivo2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboEjercicio 
      Height          =   315
      Left            =   360
      TabIndex        =   49
      Text            =   "Ejercicio"
      Top             =   120
      Width           =   990
   End
   Begin VB.Frame fraMenu 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      TabIndex        =   38
      Top             =   6660
      Width           =   9945
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8865
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Modificar"
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
         Height          =   375
         Left            =   8805
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdbuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Eliminar"
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
         Height          =   375
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Aceptar"
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
         Height          =   375
         Left            =   5745
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdcancelar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancelar"
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
         Height          =   375
         Left            =   7065
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton cmdnuevo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1650
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   405
         Width           =   975
      End
      Begin VB.CommandButton CmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
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
         Height          =   375
         Left            =   2775
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   390
         Width           =   990
      End
      Begin Gestion.ucFecha uFechaDesde 
         Height          =   270
         Left            =   1440
         TabIndex        =   40
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   476
         FechaInit       =   5
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Desde:"
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
         TabIndex        =   48
         Top             =   75
         Width           =   1440
      End
   End
   Begin VB.Frame fraOption 
      Height          =   555
      Left            =   2445
      TabIndex        =   3
      Top             =   -30
      Width           =   5715
      Begin VB.OptionButton optingreso 
         Caption         =   "Ingreso"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   0
         Tag             =   "0"
         Top             =   210
         Width           =   1575
      End
      Begin VB.OptionButton optegreso 
         Caption         =   "Egreso"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1950
         TabIndex        =   1
         Top             =   210
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optdeposito 
         Caption         =   "Depósito"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3750
         TabIndex        =   2
         Top             =   210
         Width           =   1575
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   2700
      Left            =   315
      TabIndex        =   22
      Top             =   3795
      Width           =   5280
      _cx             =   9313
      _cy             =   4762
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.TextBox txtcuentacaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6945
      TabIndex        =   37
      Tag             =   "5"
      Top             =   3825
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcotiz 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   9
      Tag             =   "1"
      Top             =   1020
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtmoneda 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Tag             =   "1"
      Top             =   1020
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbcotizacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cotizaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8640
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6945
      TabIndex        =   34
      Tag             =   "5"
      Top             =   3465
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5865
      TabIndex        =   23
      Tag             =   "8"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   7
      Tag             =   "2"
      Top             =   615
      Width           =   2775
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
      Height          =   315
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtcodcaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Tag             =   "1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtvalor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1905
      TabIndex        =   19
      Top             =   3420
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtconc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1905
      TabIndex        =   18
      Top             =   3030
      Visible         =   0   'False
      Width           =   5700
   End
   Begin VB.CommandButton cmdcargar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
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
      Height          =   495
      Left            =   5865
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Tag             =   "9"
      Top             =   4065
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmbeliminofila 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar Fila"
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
      Height          =   495
      Left            =   5865
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4665
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtcodcli 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4305
      TabIndex        =   14
      Tag             =   "5"
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmbcambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
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
      Height          =   375
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1950
      Width           =   975
   End
   Begin VB.TextBox txtcliente 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Tag             =   "2"
      Top             =   1995
      Width           =   3375
   End
   Begin VB.TextBox txtmovimiento 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "0"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtimporte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1905
      TabIndex        =   13
      Tag             =   "4"
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtconcepto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1905
      TabIndex        =   11
      Tag             =   "2"
      Top             =   1500
      Width           =   5655
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   8385
      TabIndex        =   12
      Tag             =   "3"
      Top             =   1500
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   188743681
      CurrentDate     =   38052
   End
   Begin Gestion.ucCoDe uCuenta 
      Height          =   315
      Left            =   1905
      TabIndex        =   17
      Top             =   2625
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.Label Label9 
      Caption         =   "Ejercicio"
      Height          =   255
      Left            =   1440
      TabIndex        =   50
      Top             =   180
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   9840
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label lblcotiz 
      Caption         =   "Cotización:"
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
      TabIndex        =   36
      Top             =   1020
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblmoneda 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   1020
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "TOTAL"
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
      Left            =   5880
      TabIndex        =   33
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto:"
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
      TabIndex        =   32
      Top             =   3030
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Importe:"
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
      TabIndex        =   31
      Top             =   3420
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta:"
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
      TabIndex        =   30
      Top             =   2625
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   7665
      TabIndex        =   29
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblcambio 
      Caption         =   "Cliente:"
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
      Left            =   3585
      TabIndex        =   28
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Importe:"
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
      TabIndex        =   27
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Nº Movimiento:"
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
      Left            =   360
      TabIndex        =   26
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Concepto/Resp.:"
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
      TabIndex        =   25
      Top             =   1500
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Caja:"
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
      TabIndex        =   24
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "FrmIngEgrEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4  frmcotizacion.cotizacion ?

Dim midDoc As Long
Private Const ASIENTO_ORIGEN = "EFE"

Dim rsefec As New ADODB.Recordset
Dim Ope As String
Dim modifico As Boolean
Dim numero As Long



Private Sub cmbcaja_Click()
    cargar = "Cajas"
    FrmHelp.Show
    CargarHelp "Cajas", "Codigo", "Descripción", "codigo", "responsable"
    FrmHelp.Tag = Me.Name
End Sub

Private Sub cmbcaja_GotFocus()
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub cmbcambio_Click()
    Dim cta
    If optdeposito = True Then
        'FrmHelp.Show
        CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
        'FrmHelp.Tag = Me.Name
        cta = frmBuscar.MostrarSql("select c.codigo,c.Numero as [Numero            ],b.descripcion as [Descripcion                             ] from Ctasbank c inner join bancosgrales b on c.banco=b.codigo  where c.activo=1")
        If cta > "" Then
            txtcodcli.Text = cta
            txtCliente.Text = frmBuscar.resultado(2) & " - " & frmBuscar.resultado(3)
        End If
        cargar = "Deposito"
    End If
    
    If optingreso = True Then
        FrmHelp.Show
        CargarHelp "Clientes", "Codigo", "Descripcion", "codigo", "descripcion"
        FrmHelp.Tag = Me.Name
        cargar = "Clientes"
    End If
    
    If optegreso = True Then
        FrmHelp.Show
        CargarHelp "Prov", "Codigo", "Descripcion", "codigo", "descripcion"
        FrmHelp.Tag = Me.Name
        cargar = "Proveedor"
    End If
        
End Sub

Private Sub cmbcotizacion_Click()
    FrmCotizaciones.cmbMoneda = txtmoneda
    FrmCotizaciones.cmbMoneda.enabled = False
    FrmCotizaciones.Show vbModal
    txtcotiz = FrmCotizaciones.txtCotizacion
End Sub

'Private Sub cmbcuenta_Click()
'    FrmHelp.Show
'    CargarHelpCuentas "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "Cuentas"
'End Sub

Private Sub cmbeliminofila_Click()
    If GRILLA.TextMatrix(GRILLA.Row, GRILLA.Col) <> "" Then
        If GRILLA.rows > 1 Then
            txttotal = s2n(txttotal) - s2n(GRILLA.TextMatrix(GRILLA.Row, 3))
            If GRILLA.rows = 2 Then
                GRILLA.TextMatrix(1, 0) = ""
                GRILLA.TextMatrix(1, 1) = ""
                GRILLA.TextMatrix(1, 2) = ""
                GRILLA.TextMatrix(1, 3) = ""
            Else
                GRILLA.RemoveItem (GRILLA.Row)
            End If
        Else
            MsgBox "No hay items para eliminar o no ha seleccionado ninguno de ellos"
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim Tipo As String
    Dim NroPago As Long  ' RegistroDocumentos

    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba

    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim asie As New Asiento, Importe As Double, Cuenta As String ', iddoc As Long
    Dim nMovBanco As Long
    
    
    If txtcodcaja = "" Then
        MsgBox "Debe ingresar un código de Caja"
        Exit Sub
    End If
    
    If Not optdeposito Then
        If s2n(txttotal) <> s2n(txtimporte) Then
            MsgBox "No coincide el importe ingresado con el importe total"
            Exit Sub
        End If
    End If
    
    If Ope <> "" Then
        
'        If txtmoneda <> "Pesos" Then
'            rs.Open "select * from Cotizaciones where moneda = " & ObtenerCodigo("Monedas", txtmoneda) & " and Fecha = cdate('" & Date & "') and activo = 1", daTaenvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If rs.EOF Then
'               MsgBox "La moneda asociada a la caja ingresada no se encuentra actualizada"
'               End Sub
'            End If
'        End If
        
       
       '***************************************
        DE_BeginTrans
       
        If Ope = "A" Then
        
            If optegreso Then NroPago = NuevoNroPago()
            midDoc = NuevoDocumento(ASIENTO_ORIGEN, s2n(Txtmovimiento), 0, NroPago)

            If optingreso = True Then
                Tipo = "I"
                asie.nuevo "Ingr efectivo caja " & txtcodcaja & " mov " & Txtmovimiento, dtFecha, ASIENTO_ORIGEN
                DataEnvironment1.dbo_MOVICAJAS "A", val(txtcodcaja), val(Txtmovimiento), _
                            val(txtcodcli), "E", "I", s2n(txtimporte), Trim(txtconcepto), dtFecha, Trim(txtcuentacaja), 0, s2n(txtcotiz), _
                            midDoc, Date, UsuarioSistema!codigo
'                cuenta = obtenerdesql("select cuenta from cajas where codigo = "&   txtcodcaja
                asie.AgregarItem txtcuentacaja, s2n(txtimporte), 0
                AsieDet asie, "H"
            
            ElseIf optegreso Then
                Tipo = "E"
                asie.nuevo "Egreso efectivo caja " & txtcodcaja & " mov " & Txtmovimiento, dtFecha, ASIENTO_ORIGEN
                DataEnvironment1.dbo_MOVICAJAS "A", val(txtcodcaja), val(Txtmovimiento), _
                        val(txtcodcli), "E", "E", s2n(txtimporte), Trim(txtconcepto), dtFecha, Trim(txtcuentacaja), 0, s2n(txtcotiz), midDoc, Date, UsuarioSistema!codigo
                asie.AgregarItem txtcuentacaja, 0, s2n(txtimporte)
                AsieDet asie, "D"
                
            ElseIf optdeposito Then
                Tipo = "D"
                asie.nuevo "Deposito  " & txtcodcaja & " mov " & Txtmovimiento, dtFecha, ASIENTO_ORIGEN
                DataEnvironment1.dbo_MOVICAJAS "A", val(txtcodcaja), val(Txtmovimiento), _
                        val(txtcodcli), "E", "E", s2n(txtimporte), Trim(txtconcepto), dtFecha, Trim(txtcuentacaja), 0, s2n(txtcotiz), midDoc, Date, UsuarioSistema!codigo
                nMovBanco = nuevoCodigo("movibanc", "MovBanco")
                DataEnvironment1.dbo_MOVIBANCOS "A", txtcodcli, "E", txtconcepto, dtFecha, "E", s2n(txtimporte), nMovBanco, midDoc, Date, UsuarioActual(), 1
                
                asie.AgregarItem txtcuentacaja, 0, s2n(txtimporte)
                asie.AgregarItem verCuentaContableBanco(val(txtcodcli)), s2n(txtimporte), 0
            End If
            
            asie.Grabar midDoc, , leerEjercicioId(cboEjercicio)
            
        End If
        DE_CommitTrans
        '************************************
        
        MsgBox "La operación fue realizada con éxito"
        ImprimirIngresoCaja Txtmovimiento, Tipo, midDoc
        LimpioControles
        Call Habilitobotones(True, True, True, True, True, True, True)
        Call HabilitoControles(False)
        Call MonedaVisible(False)
        GRILLA.clear
        InicioGrilla
        cargar = ""
        habilitogrillaenable (False)
    Else
        MsgBox "Operación no válida"
    End If
            
'            If optingreso Or optegreso Then
'                For i = 1 To grilla.rows - 1
'                    Importe = s2n(grilla.TextMatrix(i, 3))
'                    cuenta = Trim(grilla.TextMatrix(i, 0))
'
'                    If optingreso Then
'        '                 DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovimiento), _
'        '                    s2n(grilla.TextMatrix(i, 3)), IIf(txtcodcli <> "", Val(txtcodcli), 0), Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "IE"
'                         asie.AcumularItem cuenta, Importe, 0
'                         asie.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), 0, Importe
'                    Else
'                         asie.AcumularItem cuenta, 0, Importe
'                         ' ????????????????????
'                         asie.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), Importe, 0 ' ????????????????????
'                         ' ????????????????????
'                    End If
'                Next

'                if
'                    DE_RollbackTrans
'                    ufa "Err al grabar asiento ", Me.Name & " 1 " '& sAssert
'                    Exit Sub
'                End If
                
'            End If
'        Else
'            If Ope = "M" Then
'                If optingreso = True Then
'                    asie.Nuevo "Ingr efectivo caja " & txtcodcaja & " mov " & txtmovimiento, dtfecha, ASIENTO_ORIGEN
'
'                    DataEnvironment1.dbo_MOVICAJAS "M", Val(txtcodcaja), Val(txtmovimiento), _
'                        Val(txtcodcli), "E", "I", s2n(txtimporte), Trim(txtConcepto), dtfecha, _
'                        Trim(txtcuentacaja), 0, s2n(txtcotiz), 0, 0, 0, 0, 0
'                ElseIf optegreso Then
'                    asie.Nuevo "Egreso efectivo caja " & txtcodcaja & " mov " & txtmovimiento, dtfecha, ASIENTO_ORIGEN
'                    DataEnvironment1.dbo_MOVICAJAS "M", Val(txtcodcaja), Val(txtmovimiento), _
'                        Val(txtcodcli), "E", "E", s2n(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacaja), 0, s2n(txtcotiz), 0, 0, 0, 0, 0
'                End If
'
'                If optingreso = True Or optegreso = True Then
''''''                    DataEnvironment1.Sistema.Execute "delete from DetalleMovCajas where movimiento = " & Val(txtmovimiento) & ""
'                    iddoc = s2n(obtenerDeSQL("select iddoc from RegistroDocumentos where activo = 1 and TipoDoc = '" & ASIENTO_ORIGEN & "' and NroDoc = '" & x2s(txtmovimiento) & "' "))
'                    If iddoc = 0 Then
'                        DE_RollbackTrans
'                        ufa "err al leer documento", ASIENTO_ORIGEN & txtmovimiento
'                        Exit Sub
'                    End If
'                    BorroDocumento iddoc
'
'                    For i = 1 To grilla.rows - 1
'                        Importe = s2n(grilla.TextMatrix(i, 3))
'                        cuenta = Trim(grilla.TextMatrix(i, 0))
'
'                        If optingreso Then
'                             asie.AcumularItem cuenta, Importe, 0
'                             asie.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), 0, Importe
'                        Else
'                             asie.AcumularItem cuenta, 0, Importe
'                             ' ????????????????????
'                             asie.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), Importe, 0 ' ????????????????????
'                             ' ????????????????????
'                        End If
'''''''                        DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovimiento) _
'''''''                          ,s2n(grilla.TextMatrix(i, 3)), IIf(txtcodcli <> "", Val(txtcodcli), 0), Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "IE"
'                    Next
'                End If
'
'                DataEnvironment1.dbo_GRABARBITACORA Val(Trim(txtmovimiento)), "Usuarios", UsuarioSistema!codigo, Date, Time, "M"
'            End If

fin:
    Set asie = Nothing
    Exit Sub
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar ", "aceptar"
    Resume fin
End Sub

Private Sub AsieDet(asie As Asiento, donde As String)
    Dim i As Long, x As Double
    For i = 1 To GRILLA.rows - 1
        x = GRILLA.TextMatrix(i, 3)
        If donde = "D" Then
            asie.AgregarItem GRILLA.TextMatrix(i, 0), x, 0
        Else
            asie.AgregarItem GRILLA.TextMatrix(i, 0), 0, x
        End If
    Next i
End Sub
Private Sub cmdBuscar_Click()
    Dim resu
    
    cargar = "Movicaja"
'    FrmHelp.Show
'    CargarHelp "MOVICAJA", "Movimiento", "Caja", "movimiento", "caja", "movimiento desc", " fecha > " & uFechaDesde.ssFecha
    
    '
    'mIdDoc = s2n(obtenerDeSQL("select iddoc from movicaja where movimiento = " & txtmovimiento))
    
'    FrmHelp.Tag = Me.Name

    resu = frmBuscar.MostrarSql("select [iddoc],[Movimiento],[Caja  ],[Fecha  ],[Concepto                        ],[Importe  ] from MOVICAJA where caja>0 and activo = 1  and  fecha >=  " & ssFecha(uFechaDesde.strFecha) & "  order by movimiento desc")
    If resu <> "" Then
               
        CargarDatos s2n(resu)
    
        Call Habilitobotones(True, False, True, True, True, True, True)
    End If
End Sub

Private Sub cmdCancelar_Click()
    GRILLA.clear
    InicioGrilla
    LimpioControles
    LimpioImputacion
    Call HabilitoControles(False)
    Call Habilitobotones(True, True, True, False, False, False, True)
    Call MonedaVisible(False)
    cargar = ""
    
End Sub
Private Function CargarDatos(Optional CODMOV As Long = 0)
    Dim rs As New ADODB.Recordset
    Dim mon As Long
'    Dim Fecha As String,
    Dim codigo

    If rsefec.State = 1 Then
        rsefec.Close
        Set rsefec = Nothing
    End If
    
'    codigo = Val(Trim(Me.Tag))
    codigo = val(Trim(frmBuscar.resultado(1)))
    
    If cargar = "Cajas" Then
        
        rs.Open "select * from Cajas where codigo = " & val(txtcodcaja) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcaja = rs!codigo
            txtcaja = rs!responsable
            txtcuentacaja = rs!Cuenta
            If Not IsNull(rs!moneda) Then
                txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
            End If
        End If
        rs.Close
        Set rs = Nothing
        
'        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        
        If ObtenerCodigo("Monedas", txtmoneda) <> 1 Then
            rs.Open "select * from Cotizaciones where Fecha =" & ssFecha(dtFecha) & " and moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtmoneda = ObtenerDescripcion("Monedas", ObtenerCodigo("Monedas", txtmoneda))
                txtcotiz = rs!cotizacion
            Else
                MsgBox "Debe ingresar la cotización del día"
            End If
            MonedaVisible (True)
            rs.Close
            Set rs = Nothing
        End If
    End If


    If cargar = "Deposito" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Proveedor" Then
        rs.Open "select * from Prov where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Clientes" Then
        rs.Open "select * from Clientes where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Movicaja" Then
        rsefec.Open "select * from MOVICAJA where activo = 1 and IDDOC = " & CODMOV & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        
        If Not rsefec.EOF Then
            cargodatos
        End If
        rsefec.Close
        Set rsefec = Nothing
    End If

End Function

Private Sub cmdcargar_Click()
    Dim totalgrilla ' compil *******************
    
    If txtvalor <> "" Then
        If modifico = False Then
            If (s2n(txtvalor) <= s2n(txtimporte)) And (s2n(txtvalor) + s2n(txttotal) <= s2n(txtimporte)) Then
                If txttotal <> "" Then
                    If s2n(txttotal) + s2n(txtvalor) <= s2n(txtimporte) Then
                        Cargogrilla
                    Else
                        MsgBox "Con este valor el importe total serìa superado", vbInformation
                    End If
                Else
                    Cargogrilla
                End If
                Limpiotextosgrilla
                If uCuenta.enabled = True Then
                    uCuenta.SetFocus
                End If
            Else
                If optegreso = True Then
                    MsgBox "El valor a egresar debe ser el mismo que el original"
                Else
                    MsgBox "El valor a ingresar no puede superar al importe original"
                End If
                txtvalor.SetFocus
            End If
        Else
            totalgrilla = sumogrilla()
            If totalgrilla - s2n(GRILLA.TextMatrix(GRILLA.Row, 3)) + s2n(txtvalor) <= s2n(txtimporte) Then
                GRILLA.TextMatrix(GRILLA.Row, 0) = uCuenta.codigo 'txtcodcuenta
                GRILLA.TextMatrix(GRILLA.Row, 1) = uCuenta.DESCRIPCION 'txtcuenta
                GRILLA.TextMatrix(GRILLA.Row, 2) = txtconc
                GRILLA.TextMatrix(GRILLA.Row, 3) = txtvalor
                txttotal = sumogrilla()
                LimpioImputacion
                modifico = False
                GRILLA.SetFocus
            Else
                MsgBox "El valor a ingresar no puede superar al total"
                txtvalor.SetFocus
            End If
        End If
    Else
        MsgBox "Debe ingresar un valor"
        txtvalor.SetFocus
    End If
End Sub

Function sumogrilla() As Double
Dim x As Long
Dim Total As Double
    
    For x = 1 To GRILLA.rows - 1
        Total = Total + s2n(GRILLA.TextMatrix(x, 3))
    Next
    sumogrilla = Total
    
End Function

Private Sub LimpioImputacion()
'    txtcodcuenta = ""
'    txtcuenta = ""
    uCuenta.clear
    txtconc = ""
    txtvalor = ""
End Sub

Private Sub MonedaVisible(habilito As Boolean)
    lblmoneda.Visible = habilito
    lblcotiz.Visible = habilito
    txtmoneda.Visible = habilito
    txtcotiz.Visible = habilito
    cmbcotizacion.Visible = habilito
End Sub
Private Sub cmdeliminar_Click()
    'OJO que MOVIBANC borra por iddoc
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
   If confirma("Esta seguro de querer eliminar este registro") Then
        
        '******************************
        DE_BeginTrans

            If midDoc = 0 Then
            '    DE_RollbackTrans
                ufa "", ASIENTO_ORIGEN & " " & Txtmovimiento
                Exit Sub
            Else
                BorroDocumento midDoc
            End If
            
            DataEnvironment1.dbo_MOVICAJAS "B", 0, Trim(Txtmovimiento), 0, "", "", 0, "", 0, "", 0, 0, midDoc, Date, UsuarioSistema!codigo
            DataEnvironment1.dbo_MOVIBANCOS "B", 0, "", "", Date, "", 0, 0, midDoc, Date, UsuarioSistema!codigo, 0
            DataEnvironment1.dbo_GRABARBITACORA val(Trim(Txtmovimiento)), "", UsuarioSistema!codigo, Date, Time, "B"
        
        DE_CommitTrans
        '******************************
        
        
        MsgBox "El registro se ha eliminado"
        Call Habilitobotones(True, True, True, True, False, False, False)
        Call HabilitoControles(False)
        LimpioControles
        InicioGrilla
    End If
fin:
    Exit Sub
UFAelim:
    DE_RollbackTrans
    ufa "err al eliminar", "iddoc " & midDoc & " mov " & Txtmovimiento
    Resume fin
End Sub

Private Sub cmdImprimir_Click()
Dim Tipo As String
If Trim(Txtmovimiento) <> "" Then
   If optingreso = True Then Tipo = "I"
   If optegreso = True Then Tipo = "E"
   If optdeposito = True Then Tipo = "D"
   ImprimirIngresoCaja Txtmovimiento, Tipo, midDoc
End If
End Sub

'Private Sub cmdmodificar_Click()
'    Ope = "M"
'    Call HabilitoControles(True)
'    Call Habilitobotones(True, False, False, True, True, True)
'    habilitogrillaenable (True)
'    Call MonedaVisible(True)
'End Sub

Private Sub cmdnuevo_Click()
    On Error GoTo ufa
    
    Dim rs As New ADODB.Recordset

    Call HabilitoControles(True)
    Call Habilitobotones(False, False, True, False, False, True, True)
    LimpioControles
    
    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(rs!maxcodigo) Then
        Txtmovimiento = rs!maxcodigo + 1
        numero = rs!maxcodigo + 1
    End If
    rs.Close
    Set rs = Nothing
    
    Ope = "A"
    modifico = False
    
    txtcodcaja = "1"
    cargar = "Cajas"
    CargarDatos
    
    optegreso.SetFocus
fin:
    Exit Sub
ufa:
    ufa "", "nuevo, ing egr efectivo"
    Resume fin
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub




Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 1500
    InicioGrilla
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' ", "select cuenta as [ Cuenta           ], Descripcion as [  Descripcion                ] from cuentas where activo = 1 and imputable = 1", True
    uFechaDesde.dtFecha Date - 1
    
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
        Label9.Visible = False
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub grilla_Click()
    modifico = True
    MuestroGrilla
End Sub

Private Sub MuestroGrilla()
'    txtcodcuenta = grilla.TextMatrix(grilla.row, 0)
    uCuenta.codigo = GRILLA.TextMatrix(GRILLA.Row, 0)
'    txtcuenta = grilla.TextMatrix(grilla.row, 1)
    txtconc = GRILLA.TextMatrix(GRILLA.Row, 2)
    txtvalor = GRILLA.TextMatrix(GRILLA.Row, 3)
End Sub

Private Sub optdeposito_Click()
    lblcambio.caption = "Cuenta"
    cmbcambio.caption = "Cuenta"
    habilitogrilla (False)
End Sub

'Private Sub optdeposito_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optegreso_Click()
    lblcambio.caption = "Proveedor"
    cmbcambio.caption = "Proveedor"
    habilitogrilla (True)
End Sub

'Private Sub optegreso_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optingreso_Click()
    lblcambio.caption = "Cliente"
    cmbcambio.caption = "Cliente"
    habilitogrilla (True)
End Sub


'Private Sub optingreso_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub txtcodcaja_GotFocus()
    txtcodcaja.SelStart = 0
    txtcodcaja.SelLength = Len(txtcodcaja.Text)
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcodcaja_LostFocus()
    If IsNumeric(txtcodcaja) Then
        txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
        If txtcaja = "" Then
            MsgBox "Codigo de caja incorrecto"
            txtcodcaja = "0"
            txtcodcaja.SetFocus
        Else
            cargar = "Cajas"
            CargarDatos
        End If
'    Else
'        If txtcodcaja <> "" Then
'            MsgBox "Código de caja incorrecto"
'            txtcodcaja = "0"
'            txtcodcaja.SetFocus
'        End If
    End If
End Sub


Private Sub txtcodcli_GotFocus()
    txtcodcli.SelStart = 0
    txtcodcli.SelLength = Len(txtcodcli.Text)
End Sub

Private Sub txtcodcli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodcli_LostFocus()
    Select Case lblcambio.caption
        Case "Cliente":
                If IsNumeric(txtcodcli) Then
                    txtCliente = ObtenerDescripcion("Clientes", val(txtcodcli))
                    If txtCliente = "" Then
                        MsgBox "Codigo de cliente incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Clientes"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Codigo de cliente incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
                
        Case "Proveedor":
                If IsNumeric(txtcodcli) Then
                    txtCliente = ObtenerDescripcion("Prov", val(txtcodcli))
                    If txtCliente = "" Then
                        MsgBox "Codigo de proveedor incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Proveedor"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Codigo de proveedor incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
                
        Case "Cuenta":
                If IsNumeric(txtcodcli) Then
                    txtCliente = obtenerDeSQL("select banco from CtasBank where codigo =  " & txtcodcli)
                    If txtCliente = "" Then
                        MsgBox "Codigo de deposito incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Deposito"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Codigo de deposito incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
        End Select
End Sub

'Private Sub txtcodcuenta_GotFocus()
'Dim rs As New ADODB.Recordset
'
'    rs.Open "select dato_fijo from datos", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs.EOF Then
'        If rs!DATO_FIJO = 7 Then
'            txtcodcuenta = "1"
'            txtcodcuenta.Enabled = False
'            txtcuenta = "COMPRAS"
'            txtconc = "COMPRAS"
'            txtconc.Enabled = False
'            txtvalor = txtimporte
'            txtvalor.Enabled = False
'            cmbcuenta.Enabled = False
'            cmdcargar.Enabled = False
'            cmbeliminofila.Enabled = False
'            Cargogrilla
'        End If
'    End If
'    rs.Close
'
'End Sub

'Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub

'Private Sub txtcodcuenta_LostFocus()
'    If IsNumeric(txtcodcuenta) Then
'        If Not noestaenlagrilla(txtcodcuenta, grilla) And esimputable(Val(txtcodcuenta)) Then
'            txtcuenta = ObtenerDescripcion("Cuentas", Val(txtcodcuenta))
'            If txtcuenta = "" Then
'                MsgBox "Codigo de cuenta incorrecto"
'                txtcodcuenta = ""
'                txtcodcuenta.SetFocus
'            Else
'                cargar = "Cuentas"
'                CargarDatos
'            End If
'        Else
'            MsgBox "El concepto ya se encuentra cargado o la cuenta no es imputable"
'            txtcodcuenta = ""
'            txtcodcuenta.SetFocus
'        End If
'    Else
'        If txtcodcuenta <> "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = ""
'            txtcodcuenta.SetFocus
'        End If
'    End If
'End Sub

'Private Sub txtconc_Change()
'Dim i As Integer
'    txtconc.Text = UCase(txtconc.Text)
'    i = Len(txtconc.Text)
'    txtconc.SelStart = i
'End Sub

Private Sub txtconc_GotFocus()
'    If uCuenta.codigo = "" Then
'        MsgBox "Debe cargar la cuenta"
'        uCuenta.SetFocus
'    End If
    If Trim(txtconc) = "" Then txtconc = txtconcepto
    PintoFocoActivo
End Sub

'Private Sub txtconc_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub txtconcepto_Change()
Dim i As Long
    txtconcepto.Text = UCase(txtconcepto.Text)
    i = Len(txtconcepto.Text)
    txtconcepto.SelStart = i
End Sub

Private Sub txtConcepto_GotFocus()
    txtconcepto.SelStart = 0
    txtconcepto.SelLength = Len(txtconcepto.Text)
End Sub

'Private Sub txtconcepto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub



Private Sub txtcotiz_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcotiz_LostFocus()
    If Not IsNumeric(txtcotiz) Then
        MsgBox "Cotización incorrecta"
        txtcotiz = "0"
        txtcotiz.SetFocus
    End If
End Sub

Private Sub txtimporte_GotFocus()
    txtimporte.SelStart = 0
    txtimporte.SelLength = Len(txtimporte.Text)
End Sub


Private Sub txtimporte_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

'Private Sub txtmoneda_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub txtmovimiento_GotFocus()
    Txtmovimiento.SelStart = 0
    Txtmovimiento.SelLength = Len(Txtmovimiento.Text)
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub txtmovimiento_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtmovimiento_LostFocus()
    If IsNumeric(Txtmovimiento) Then
        If val(Txtmovimiento) < numero Then
            MsgBox "El código no puede ser menor al último ingresado"
            Txtmovimiento.SetFocus
        End If
    Else
        If Txtmovimiento <> "" Then
            MsgBox "Debe ingresar un código"
            Txtmovimiento = "0"
            Txtmovimiento.SetFocus
        End If
    End If
End Sub

Sub LimpioControles()
    Txtmovimiento = ""
    txtconcepto = ""
    dtFecha = Date
    txtcodcli = ""
    txtCliente = ""
    txttotal = "0"
    txtcotiz = "0"
    txtimporte = ""
    txtcodcaja = ""
    txtcaja = ""
    txtcuentacaja = ""
    optegreso.Value = False
    optingreso.Value = False
    optdeposito.Value = False
    Ope = ""
    
    midDoc = 0
End Sub

Sub cargodatos()
    Dim rs As New ADODB.Recordset

    midDoc = rsefec!iddoc
    
    If rsefec!Ing_egr = "I" Then
        optingreso.Value = True
    Else
        If rsefec!Ing_egr = "E" Then
            optegreso.Value = True
        Else
            optdeposito.Value = True
        End If
    End If
    
    Txtmovimiento = rsefec!movimiento
    txtcodcaja = rsefec!caja
    txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
    txtmoneda = ObtenerDescripcion("Monedas", ObtenerMoneda("Cajas", val(txtcodcaja)))
    If Not IsNull(rsefec!concepto) Then
        txtconcepto = rsefec!concepto
    End If
    
    dtFecha = rsefec!Fecha
    txtimporte = rsefec!Importe
    
    Call MonedaVisible(True)
    
    rs.Open "select * from Cotizaciones where Fecha = " & ssFecha(dtFecha) & " and  moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        MonedaVisible (True)
        txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
        txtcotiz = rs!cotizacion
    Else
        txtcotiz = "0"
    End If
    rs.Close
    Set rs = Nothing
    
    
    If Not IsNull(rsefec!cli_prov) Then
        txtcodcli = rsefec!cli_prov
        txtCliente = ObtenerDescripcion("Clientes", val(txtcodcli))
    End If
        
    InicioGrilla
    txttotal = "0"
    


    Set rs = Nothing
End Sub

Sub HabilitoControles(habilito As Boolean)
    Txtmovimiento.enabled = habilito
    cmbcaja.enabled = habilito
    cmbcambio.enabled = habilito
    txtconcepto.enabled = habilito
    dtFecha.enabled = habilito
    txtcodcli.enabled = habilito
    optegreso.enabled = habilito
    optingreso.enabled = habilito
    txtimporte.enabled = habilito
    optdeposito.enabled = habilito
    txtcodcaja.enabled = habilito
    txtcodcli.enabled = habilito
End Sub

Sub Habilitobotones(busco As Boolean, nuevo As Boolean, Imprimo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdbuscar.enabled = busco
    cmdnuevo.enabled = nuevo
    cmdmodificar.enabled = modifico
    cmdeliminar.enabled = elimino
    cmdAceptar.enabled = acepto
    cmdcancelar.enabled = Cancelo
    cmdImprimir.enabled = Imprimo
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rsefec.State = 1 Then
        rsefec.Close
        Set rsefec = Nothing
    End If
End Sub

'Private Sub txtconcepto_LostFocus()
'    If txtconcepto = "" Then
'        MsgBox "Debe ingresar un concepto"
'        txtconcepto.SetFocus
'    End If
'End Sub

Sub InicioGrilla()
    GRILLA.clear
    'grilla.ColWidth(1) = 1700
    GRILLA.TextMatrix(0, 0) = "Cuenta"
    GRILLA.TextMatrix(0, 1) = "Descripción"
    GRILLA.TextMatrix(0, 2) = "Concepto"
    GRILLA.TextMatrix(0, 3) = "Importe"
    GRILLA.rows = 2
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label2.Visible = habilito

'    txtcodcuenta.Visible = habilito
'    cmbcuenta.Visible = habilito
'    txtcuenta.Visible = habilito
    uCuenta.Visible = habilito  ' ???

    Label6.Visible = habilito
    txtconc.Visible = habilito
    Label3.Visible = habilito
    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
    GRILLA.Visible = habilito
    cmbeliminofila.Visible = habilito
    Label8.Visible = habilito
    txttotal.Visible = habilito
End Sub

Private Sub txtimporte_LostFocus()
    If Not IsNumeric(txtimporte) Then
        MsgBox "Debe ingresar un importe"
        txtimporte = "0"
        txtimporte.SetFocus
    Else
        Call habilitogrillaenable(True)
        txtimporte = s2n(txtimporte)
    End If
End Sub

Private Sub Limpiotextosgrilla()
'    txtcodcuenta = ""
'    txtcuenta = ""
    uCuenta.clear
    
    txtconc = ""
    txtvalor = ""
End Sub


Private Sub Cargogrilla()
    If GRILLA.rows = 2 Then
        GRILLA.Row = 1
        GRILLA.Col = 0
        If Trim(GRILLA.Text) = "" Then
            GRILLA.Row = 1
            GRILLA.Col = 0
            GRILLA.Text = uCuenta.codigo 'txtcodcuenta
            GRILLA.Col = 1
            GRILLA.Text = uCuenta.DESCRIPCION ' txtcuenta
            GRILLA.Col = 2
            GRILLA.Text = txtconc
            GRILLA.Col = 3
            GRILLA.Text = txtvalor
        Else
            GRILLA.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
        End If
    Else
        GRILLA.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
    End If
    If txttotal <> "" Then
        txttotal = s2n(txttotal) + s2n(txtvalor)
    Else
        txttotal = s2n(txtvalor)
    End If
    If txttotal = txtimporte Then
        MsgBox "El detalle ha sido completado"
'        habilitogrillaenable (False)
    End If
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)
    Label2.enabled = habilito
'    txtcodcuenta.Enabled = habilito
'    cmbcuenta.Enabled = habilito
    uCuenta.enabled = habilito
    
    Label6.enabled = habilito
    txtconc.enabled = habilito
    Label3.enabled = habilito
    txtvalor.enabled = habilito
    cmdcargar.enabled = habilito
    GRILLA.enabled = habilito
    cmbeliminofila.enabled = habilito
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtvalor_LostFocus()
    If IsNumeric(txtvalor) Then
'        InicioGrilla
        If GRILLA.Visible = False Then
            habilitogrilla (True)
        End If
        habilitogrillaenable (True)
        txtvalor = s2n(txtvalor)
    Else
        If txtvalor <> "" Then
            MsgBox "Debe ingresar un importe"
            txtvalor = "0"
            txtvalor.SetFocus
        End If
    End If
End Sub

'5/5/5
'   numero long
'

Private Sub uFechaDesde_LostFocus()
    On Error Resume Next
    cmdbuscar.SetFocus
End Sub
