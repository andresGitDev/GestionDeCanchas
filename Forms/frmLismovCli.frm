VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLismovCli 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Listado de composicion de saldo de clientes"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8775
      TabIndex        =   24
      Top             =   120
      Width           =   2535
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   1080
         TabIndex        =   25
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   1080
         TabIndex        =   26
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   39347
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desde"
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
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta"
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
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   135
      TabIndex        =   19
      Top             =   120
      Width           =   8535
      Begin Gestion.ucCoDe uCliH 
         Height          =   330
         Left            =   870
         TabIndex        =   20
         Top             =   750
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   582
         CodigoWidth     =   800
         CodigoInvalido  =   0
      End
      Begin Gestion.ucCoDe uCliD 
         Height          =   315
         Left            =   885
         TabIndex        =   21
         Top             =   285
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   556
         CodigoWidth     =   800
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desde"
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
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hasta"
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
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame fraGri 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   10
      Top             =   2310
      Width           =   11175
      Begin VB.Frame fraSubGri 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2220
         Left            =   195
         TabIndex        =   11
         Top             =   3015
         Width           =   10725
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   1935
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   240
            Width           =   5415
            _cx             =   9551
            _cy             =   3413
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
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
         Begin VSFlex7LCtl.VSFlexGrid GrillaMoviCaja 
            Height          =   855
            Left            =   5520
            TabIndex        =   13
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   240
            Width           =   5415
            _cx             =   9551
            _cy             =   1508
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
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
         Begin VSFlex7LCtl.VSFlexGrid GrillaEfectivo 
            Height          =   855
            Left            =   5520
            TabIndex        =   14
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   1320
            Width           =   5415
            _cx             =   9551
            _cy             =   1508
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento en Efectivo"
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
            Left            =   5520
            TabIndex        =   17
            Top             =   1080
            Width           =   2070
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento de Caja"
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
            Left            =   5520
            TabIndex        =   16
            Top             =   0
            Width           =   1785
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comprobantes que imputa:"
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
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   2415
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2775
         Left            =   195
         TabIndex        =   18
         ToolTipText     =   "Haga Click para ver el Detalle de la Orden de Compra"
         Top             =   240
         Width           =   10860
         _cx             =   19156
         _cy             =   4895
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
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
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   420
      Left            =   135
      TabIndex        =   3
      Top             =   7650
      Width           =   11145
      Begin VB.CommandButton cmdexcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Enviar a Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1275
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
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
         Height          =   360
         Left            =   10095
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdcancelar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   975
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   360
         Left            =   3150
         TabIndex        =   9
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   3495
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Con Saldo"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Saldo"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmLismovCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'sacar iva dela grilla como locaire
' puaj las constantes son un asco   x like 'FA%%'
'


Option Explicit

Private Const CONST_AJUSTE_CLI_DEBITO = "ACD"
Private Const CONST_AJUSTE_CLI_CREDITO = "ACC"
Private Const CONST_FACTURAS_A = "FAA"
Private Const CONST_FACTURAS_B = "FAB"
Private Const CONST_FACTURAS_E = "FAE"
Private Const CONST_NOTAS_DEBITOS_A = "NDA"
Private Const CONST_NOTAS_DEBITOS_B = "NDB"
Private Const CONST_NOTAS_CREDITOS_A = "NCA"
Private Const CONST_NOTAS_CREDITOS_B = "NCB"
Private Const CONST_NOTAS_CREDITOS_E = "NCE"
Private Const CONST_RECIBOS = "RAA"
Private Const CONST_RECIBOS_IMPUTADOS = "REC"
Private TablaTemp As String
Private Const CONST_CONTADO = True
'Private msRsCli As String
Private mfiltro As String

Private Function VaEnElDebe(TipoDocumento As String) As Boolean
'funcion que devuelve TRUE si el tipo de comprobante va en el DEBE o en el HABER
    If (x2s(TipoDocumento) = CONST_AJUSTE_CLI_DEBITO) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_A) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_E) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_B) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_A) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_B) Or _
                (x2s(TipoDocumento) = "FAAV") Or _
                (x2s(TipoDocumento) = "FABV") Or _
                (x2s(TipoDocumento) = "NDAV") Then
        VaEnElDebe = True
    Else
        VaEnElDebe = False
    End If
End Function


Private Function CalcularSaldoAnterior(CodigoCliente As Long, fechahasta As Date) As Double

    Dim Debe As Double
    Dim haber As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim rsaux As New ADODB.Recordset
    Dim sal As Double
    Dim tot As Double
    Dim Consulta As String

    Debe = 0
    haber = 0
    
    'TABLA FACTURAVENTA
'    Consulta = "Select TIPODOC, FORMAPAGO, CONTADO, Sum(TOTAL) as Total,codigo From FACTURAVENTA " & _
'        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & _
'        " Group By TIPODOC, FORMAPAGO, CONTADO,codigo"
    Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
                & " From FACTURAVENTA as F " _
                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCliente & " And F.FECHA<" & ssFecha(fechahasta) _
                & " Order By F.FECHA, F.CODIGO"
                
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween(dtfechad.Value, dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCliente, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsaux.EOF = True And rsaux.BOF = True Then
            If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                'pregunto si la forma de pago es contado, porque con esta no hago nada _
                '(ya que debe sumar en el DEBE y restar en el HABER)
                If Not ( _
                    (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                    And s2n(rsCuenta!contado) = CONST_CONTADO) Then
                    Debe = Debe + s2n(rsCuenta!Total)
                End If
            Else
                haber = haber + s2n(rsCuenta!Total)
            End If
            rsCuenta.MoveNext
        Else
            sal = 0
            While Not rsaux.EOF
                sal = sal + rsaux!Importe
                rsaux.MoveNext
            Wend
            If sal = rsCuenta!Total Then
                rsCuenta.MoveNext
            Else
                sal = rsCuenta!Total - sal
                
                If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                    'pregunto si la forma de pago es contado, porque con esta no hago nada _
                    '(ya que debe sumar en el DEBE y restar en el HABER)
                    If Not ( _
                        (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                        And s2n(rsCuenta!contado) = CONST_CONTADO) Then
                        Debe = Debe + s2n(sal)
                    End If
                Else
                    haber = haber + s2n(sal)
                End If
                rsCuenta.MoveNext
            End If

        End If
        Set rsaux = Nothing
    Wend
    rsCuenta.Close
    Set rsCuenta = Nothing
    
    'TABLA RECIBOS
    Consulta = "Select CLIENTE, SUM(TOTAL) AS TOTAL From RECIBOS " & _
        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & " Group By CLIENTE"
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        haber = haber + s2n(rsCuenta!Total)
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close
    Set rsCuenta = Nothing
    
    CalcularSaldoAnterior = Debe - haber
End Function

Private Sub CalcularSaldo()
    Dim rsaux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Double
    Dim CodigoCli As Long
    Dim CodigoCliActual As Long
    
    With rsaux
        Consulta = "Select * From " & TablaTemp & " Order By CODIGO_CLI, FECHA, ID"
        .Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        'If Not rsAux.EOF Then rsAux.MoveFirst
        While Not .EOF
            CodigoCli = !CODIGO_CLI
            CodigoCliActual = CodigoCli
            saldo = 0
            While CodigoCli = CodigoCliActual
                
                'If Not IsNull(!debe) And Not IsNull(rsAux!haber) Then saldo = saldo + s2n(rsAux!debe) - s2n(rsAux!haber)
                saldo = saldo + s2n(!Debe) - s2n(!haber)
                
                If !TIPO_DOCUMENTO <> "" Then
                    'Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & rsAux!ID
                   !saldo = CStr(Round(saldo, 2))
                Else
                    'Consulta = "Update " & TablaTemp & " Set SALDO = ' ' Where ID = " & rsAux!ID
                    !saldo = " "
                End If

                'DataEnvironment1.sistema.Execute Consulta
                .Update
                .MoveNext
                
                If rsaux.EOF Then
                    CodigoCliActual = 0
                Else
                    CodigoCliActual = !CODIGO_CLI
                End If
            Wend
        Wend
    End With
End Sub

Private Sub CrearConsulta(ConDetalle As Boolean)
    Dim SaldoCuenta As Double
    Dim CodigoCli As Long
    Dim Consulta As String
    Dim rsCli As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsFac As New ADODB.Recordset
    Dim rsCHQ As New ADODB.Recordset
    Dim rsaux As New ADODB.Recordset
    Dim sal As Double
    Dim NroRem As String
    
    Dim tempoCli As Variant, cliDes As String

    DataEnvironment1.Sistema.Execute "delete from " & TablaTemp
    
    
    rsCli.Open "select * from clientes where activo = 1 and codigo between " & uCliD.codigo & " and " & uCliH.codigo & mfiltro, _
        DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    While Not rsCli.EOF
'    For CodigoCli = CLng(txtcodclid) To (txtcodclih)
    
    
      'ARREGLAR PERFORMANCE !! Aca va parche, pero hay que hacer un rs clientes y barrer el rs
      'tempoCli = (obtenerDeSQL("select codigo, descripcion  from clientes where codigo = '" & CodigoCli & "' and activo = 1"))
      'If Not IsEmpty(tempoCli) Then
        cliDes = rsCli!DESCRIPCION ' tempoCli(1)
        CodigoCli = rsCli!codigo
        'ARREGLAR PERFORMANCE !! Aca va parche, pero hay que hacer un rs clientes y barrer el rs
        
        SaldoCuenta = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value))
        
         If Option3.Value = True Then ' sin saldo =0
            If SaldoCuenta = 0 Then
            Else
                If SaldoCuenta < 0 Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                            ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, 2))) & "', '" & (s2n(SaldoCuenta, 2)) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, 2)) & "', '0',  '" & (s2n(SaldoCuenta, 2)) & "')"
                End If
                DataEnvironment1.Sistema.Execute Consulta
            End If
        End If
        If Option4.Value = True Then
            If SaldoCuenta < 0 Then
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, 2))) & "', '" & (s2n(SaldoCuenta, 2)) & "')"
            Else
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, 2)) & "', '0',  '" & (s2n(SaldoCuenta, 2)) & "')"
            End If
            DataEnvironment1.Sistema.Execute Consulta
        End If
        
        SaldoCuenta = 0
        'If VerParametro(BS_NOMBRE_EMPRESA) = "nimisan swartz" Then
        '    Consulta = "Select F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
        '        & " From FACTURAVENTA as F " _
        '        & " Where ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
        '        & " Order By F.FECHA, F.CODIGO"
        'Else
            Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
                & " From FACTURAVENTA as F " _
                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and total>0" _
                & " Order By F.FECHA, F.CODIGO"
        'End If
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            Consulta = "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<=" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli
            'rs.MovePrevious
            
            rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<=" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            'rsaux.MoveFirst
            'If rs!codigo = 35411 Then
            '    Stop
            'End If
            If rsaux.EOF = True And rsaux.BOF = True Then
            
                If s2n(rs!Remito) = 0 Then
                    NroRem = " "
                Else
                    NroRem = s2n(rs!Remito)
                End If
                If VaEnElDebe(x2s(rs!TIPODOC)) Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(rs!Total, 2) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                End If
                DataEnvironment1.Sistema.Execute Consulta
                    
                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                        "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                    DataEnvironment1.sistema.Execute Consulta
'                End If
                rs.MoveNext
            Else 'esto es por si tengo algun saldo para mostrar
                sal = 0
                While Not rsaux.EOF
                    'sal = rsaux!codrecibo
                    sal = sal + rsaux!Importe
                    rsaux.MoveNext
                Wend
                If sal = rs!Total Then
                    rs.MoveNext
                Else
                    If s2n(rs!Remito) = 0 Then
                        NroRem = " "
                    Else
                        NroRem = s2n(rs!Remito)
                    End If
                    sal = rs!Total - sal
                    If VaEnElDebe(x2s(rs!TIPODOC)) Then
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(sal, 2) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    Else
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(sal, 2) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    End If
                    DataEnvironment1.Sistema.Execute Consulta
                        
                    'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                    If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                    "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                            "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                    ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                        DataEnvironment1.sistema.Execute Consulta
'                    End If
                    rs.MoveNext
                End If

            End If
            Set rsaux = Nothing
        Wend
        rs.Close
        Set rs = Nothing
        
        'TABLA RECIBOS
        
'        Select R.*,
'            D.FACTURAVENTA , D.importe
'        From RECIBOS AS R
'            INNER JOIN RECIBOSDETALLE AS D ON R.CODIGO = D.CODRECIBO
'        Where R.ACTIVO = 1 And R.CLIENTE = 1000 AND
'            R.FECHA  between convert(datetime , '06-17-04', 1)  AND convert(datetime , '12-17-04', 1)
        
'        Consulta = "Select R.* From RECIBOS AS R " _
'            & "Where R.ACTIVO = 1 And R.CLIENTE = " & CodigoCli & " AND R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
'            & "Order By R.FECHA, R.CODIGO"
'
'        rs.Open Consulta, DataEnvironment1.sistema, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then rs.MoveFirst
'        While Not rs.EOF
''            Consulta = "Insert Into " & TablaTemp & _
''                " (CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
''                " VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!Fecha) & _
''                ", '" & CONST_RECIBOS_IMPUTADOS & "', '" & x2s(rs!numero) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'            Consulta = "Insert Into " & TablaTemp & _
'                " (CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                " VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!fecha) & _
'                ", '" & rs!TIPODOC & "', '" & x2s(rs!numero) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"

'            DataEnvironment1.sistema.Execute Consulta
'            Consulta = "Select FACTURAVENTA, IMPORTE From RECIBOSDETALLE " & _
'                        "Where CODRECIBO = " & ObtenerDatoDB("RECIBOS", "NUMERO", rs!numero, "CODIGO") & " Order By CODIGO"
'            rsFac.Open Consulta, DataEnvironment1.sistema, adOpenDynamic, adLockOptimistic
'            Consulta = "Select NRO, IMPORTE From CHEQUES Where ACTIVO = 1 And TDOC = '" & CONST_RECIBOS & "' And NDOC = " & s2n(rs!numero)
'            rsCHQ.Open Consulta, DataEnvironment1.sistema, adOpenDynamic, adLockOptimistic
'            While Not rsFac.EOF Or Not rsCHQ.EOF
'            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                        "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, FACTURAS, CHEQUES) " & _
'                                "VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!fecha) & _
'                                        ", '', '', '', '', '', "
'                If Not rsFac.EOF Then
'                    Consulta = Consulta & "'FAC NRO " & x2s(rsFac!FACTURAVENTA) & " - " & Format(x2s(rsFac!importe), "###0.00") & "',"
'                    rsFac.MoveNext
'                Else
'                    Consulta = Consulta & "'',"
'                End If
'
'                If Not rsCHQ.EOF Then
 '                   Consulta = Consulta & "'CHQ NRO " & x2s(rsCHQ!Nro) & " - " & Format(x2s(rsCHQ!importe), "###0.00") & "')"
'                    rsCHQ.MoveNext
'                Else
'                    Consulta = Consulta & "'')"
'                End If
'
'                DataEnvironment1.sistema.Execute Consulta
'            Wend


'            rsFac.Close
'            Set rsFac = Nothing'

'            rsCHQ.Close
'            Set rsCHQ = Nothing
            
'            rs.MoveNext
'        Wend
'        rs.Close
'        Set rs = Nothing
'      End If ' del if existe cliente
'    Next CodigoCli
        rsCli.MoveNext
    Wend
    
    CalcularSaldo

    Set rs = Nothing
    Set rsCli = Nothing
End Sub

Private Sub cmdaceptar_Click()
    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If dtfechad.Value < CDate("01/04/2006") Then
        MsgBox "Debe ingresar una fecha posterior al 01/04/2006."
        dtfechad.Value = "01/04/2006"
        Exit Sub
    End If
    
    If rangoOk Then
                
        relojito
        
        CrearConsulta False
        LimpiarGrilla grilla
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaEfectivo
        LimpiarGrilla GrillaMoviCaja
        
'        LlenarGrilla grilla, _
'            "Select CODIGO_CLI AS CODIGO, C.DESCRIPCION AS DESCRIPCION, L.FECHA, L.TIPO_DOCUMENTO, " & _
'            "L.NRO_DOCUMENTO, L.REMITO, L.IVA, L.DEBE, L.HABER, L.SALDO " & _
'            "From " & TablaTemp & " AS L INNER JOIN CLIENTES AS C ON C.CODIGO = L.CODIGO_CLI " & _
'            "Where l.facturas = ' ' and l.cheques = ' ' " & _
'            "Order By CODIGO_CLI, FECHA, ID", True
        LlenarGrilla grilla, _
            " Select CODIGO_CLI AS CODIGO,  DESCRIPCION_CLI as DESCRIPCION, FECHA, TIPO_DOCUMENTO, " & _
            " NRO_DOCUMENTO, REMITO, DEBE, HABER, SALDO, '' as [Saldo final] " & _
            " From " & TablaTemp & _
            " Where facturas = ' ' and cheques = ' ' " & _
            " Order By CODIGO_CLI, FECHA, ID", True
        grillaMarcoSaldosFinales grilla, 0, 9, 8
        
        limpioGrilla 9
        relojito False
    
'        MsgBox "" & Grilla.rows
        
    End If
End Sub

Private Function limpioGrilla(Col As Long) 'limpio la grilla y tabla temporal de los importes con cero, incluyendo si el total es cero borro historial de cliente
    Dim i As Long
    Dim j As Long
    Dim cli As Long
    Dim Borrar As String
    
    i = 1
    While i < grilla.rows
        If grilla.TextMatrix(i, Col) = "0" Then
            cli = grilla.TextMatrix(i, 0)
            j = 1
            While j < grilla.rows
                If grilla.TextMatrix(j, 0) = CStr(cli) Then
                    grilla.TextMatrix(j, 0) = ""
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
    
    j = 1
    While j < grilla.rows
        If grilla.TextMatrix(j, 0) = "" Then
            Borrar = "delete from " & TablaTemp & " where descripcion_cli='" & grilla.TextMatrix(j, 1) & "'"
            DataEnvironment1.Sistema.Execute Borrar
            grilla.RemoveItem (j)
        Else
            j = j + 1
        End If
        'j = j + 1
    Wend
End Function


'Private Sub cmdAyudaCliH_Click()
'
'End Sub

'Private Sub cmdAyudaCliD_Click()
'
'End Sub

Private Sub cmdexcel_Click()
Dim rs As New ADODB.Recordset
Dim Consulta As String

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then
'        CrearConsulta False

        If MsgBox("¿Desea imprimir los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
                    " DEBE, HABER, SALDO " & _
                    " From  " & TablaTemp & _
                    " Order By CODIGO_CLI, FECHA, ID"
                                
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO " & _
                    " From " & TablaTemp & _
                    " Where TIPO_DOCUMENTO <> '' " & _
                    " Order By CODIGO_CLI, FECHA, ID"
        End If
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
'        LlenarGrilla grilla, Consulta, False
        
        
        VinculoXl "C:\ComposicionCuentaCli.xls", "Composicion cuenta de Cliente/s", , , rs '"C:\ComposicionCuentaCli", "Composicion cuenta cliente"
        rs.Close
        Set rs = Nothing
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim Consulta As String
Dim rsempresa As New ADODB.Recordset

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then
'        CrearConsulta False

        If MsgBox("¿Desea imprimir los detalles de los comprobantes?", vbYesNo, "Atencion") <> vbYes Then
            Consulta = "Select CODIGO_CLI, DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO " & _
                        "From " & TablaTemp & _
                        " Where TIPO_DOCUMENTO <> '' " & _
                        "Order By CODIGO_CLI, FECHA, ID"
        Else
            Consulta = "Select CODIGO_CLI, DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                        "TIPO_DOCUMENTO, NRO_DOCUMENTO, FACTURAS, CHEQUES, " & _
                        "DEBE, HABER, SALDO " & _
                        "From " & TablaTemp & _
                        " Order By CODIGO_CLI, FECHA, ID"
        End If
        RptLisMovCtaCli.data1.Connection = DataEnvironment1.Sistema
        RptLisMovCtaCli.data1.Source = Consulta
        RptLisMovCtaCli.lblfecha = Date
        RptLisMovCtaCli.LBLFECHAD = dtfechad.Value
        RptLisMovCtaCli.LBLFECHAH = dtfechah.Value
        rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        'If Not IsNull(rsempresa!nombrelogo) Then
            RptLisMovCtaCli.ImageLOGO.Picture = FrmPrincipal.imgLogoSimple ' LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
        'End If
        rsempresa.Close
        Set rsempresa = Nothing
        RptLisMovCtaCli.Show
        
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
    End If
End Sub

'Private Sub cmdAyudaCliD_Click()
'Dim Res As String
'
'    Res = frmBuscar.MostrarSql("Select CODIGO, DESCRIPCION As [Cliente                           ] From CLIENTES Where ACTIVO = 1")
'    If Res > "" Then
'        txtcodclid = frmBuscar.resultado
'        txtcliented = frmBuscar.resultado(2)
'    End If
'End Sub
'
'Private Sub cmdAyudaCliH_Click()
'Dim Res As String
'
'    Res = frmBuscar.MostrarSql("Select CODIGO, DESCRIPCION AS [Cliente                            ] From CLIENTES Where ACTIVO = 1")
'    If Res > "" Then
'        txtcodclih = frmBuscar.resultado
'        txtclienteh = frmBuscar.resultado(2)
'    End If
'End Sub

Private Sub cmdcancelar_Click()
    Dim tempo
    tempo = obtenerDeSQL("select min(codigo) as mini, max(codigo) as maxi from clientes")
    'txtcodclih = tempo(1) ' "9999999"
    uCliH.codigo = tempo(1)
'    txtclienteh = ""
    'txtcodclid = tempo(0)
    uCliD.codigo = tempo(0)
'    txtcliented = "" 'tempo(1)
    dtfechad.Value = "01/04/2007" 'Date
    dtfechah.Value = Date
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
'    cmdCancelar_Click
    
    Dim s1 As String, s2 As String, sCli As String, tempo As Variant
    
    TablaTemp = TablaTempCrear("([ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_CLI] [numeric](18, 0) NULL ," _
        & "[DESCRIPCION_CLI] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FECHA] [datetime] NULL ," _
        & "[TIPO_DOCUMENTO] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[NRO_DOCUMENTO] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[IVA] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DEBE] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[HABER] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[SALDO] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FACTURAS] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[CHEQUES] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[REMITO] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" _
        & ") ON [PRIMARY]")
    
    DataEnvironment1.Sistema.Execute " ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [DF_" & TablaTemp & "] DEFAULT (N' ') FOR [FACTURAS]," _
        & " CONSTRAINT [DF_" & TablaTemp & "1] DEFAULT (N' ') FOR [CHEQUES]"
    
    ucXls1.ini grilla, "C:\ComposicionCuentaCli.xls", "Composicion cuenta cliente"
    
    mfiltro = ""
    
    
'    tempo = VerParametro(BS_MOV_CTACTE_SOLO_MAYORISTAS)
'    If Not IsEmpty(tempo) Then
'        If tempo = 1 Then
'            mfiltro = " and mayorista = 1 "
'        End If
'    End If
    s1 = "select descripcion from clientes where codigo = ### and activo = 1 "
    s2 = "Select codigo, descripcion as [ Descripcion                                           ] from clientes where activo = 1 " '& mfiltro

    uCliD.ini s1, s2, False, True
    uCliH.ini s1, s2, False, True
    
    cmdcancelar_Click
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
    Anclar fraGri, Me, anclarLadosTodos
    Anclar grilla, fraGri, anclarLadosTodos
    Anclar fraSubGri, fraGri, anclarIzquierda + anclarAbajo
End Sub

Private Sub grilla_Click()
Dim TIPODOC As String
Dim NroDoc As Long
Dim CodInt As Long

    With grilla
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(.TextMatrix(.Row, 3))
            If .TextMatrix(.Row, 4) <> "" Then NroDoc = CLng(.TextMatrix(.Row, 4))
            
            LimpiarGrilla GrillaDetalle
            LimpiarGrilla GrillaMoviCaja
            LimpiarGrilla GrillaEfectivo
            
            Select Case TIPODOC
                Case CONST_FACTURAS_A, CONST_FACTURAS_B, CONST_FACTURAS_E
'                    LlenarGrilla GrillaDetalle, "Select D.CANTIDAD, D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO AS 'PRECIO UNITARIO', " & _
'                                                                    "D.PRECIOTOTAL AS 'PRECIO TOTAL', S.SERIE " & _
'                                         "From FACTURAVENTADETALLE AS D " & _
'                                            "Left Join SERIES as S ON S.PRODUCTO=D.PRODUCTO " & _
'                                         "Where D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & _
'                                            " AND S.NROCOMPROBANTE = " & NroDoc, False
                    LlenarGrilla GrillaDetalle, "Select D.CANTIDAD, D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO AS 'PRECIO UNITARIO', " & _
                                                                    "D.PRECIOTOTAL AS 'PRECIO TOTAL', S.SERIE " & _
                                         "From FACTURAVENTADETALLE AS D " & _
                                            "left Join SERIES as S ON S.PRODUCTO=D.PRODUCTO AND S.NROCOMPROBANTE=D.NROFACTURA " & _
                                         "Where D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & _
                                         " Order By ID", False



                Case CONST_RECIBOS, CONST_RECIBOS_IMPUTADOS
                    CodInt = ObtenerDatoDB("RECIBOS", "NUMERO", NroDoc, "CODIGO")
                    LlenarGrilla GrillaDetalle, "Select F.TIPODOC AS 'Tipo', F.NROFACTURA AS 'Numero', Importe " & _
                                                "From RECIBOSDETALLE AS R " & _
                                                    "Inner Join FACTURAVENTA AS F on F.CODIGO = R.FACTURAVENTA " & _
                                                "Where R.CODRECIBO = " & CodInt & " Order By R.CODIGO", True
          
                    LlenarGrilla GrillaMoviCaja, "Select Fecha, nro as 'Numero Cheque', Importe From CHEQUES " & _
                                        "Where ACTIVO = 1 And TDOC = '" & TIPODOC & "' AND NDOC = " & NroDoc, True
                    LlenarGrilla GrillaEfectivo, "Select Fecha, Importe From MOVICAJA " & _
                                        "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                            "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                
                
                Case CONST_NOTAS_DEBITOS_A, CONST_NOTAS_CREDITOS_A, CONST_NOTAS_CREDITOS_B
                
                Case CONST_AJUSTE_CLI_DEBITO, CONST_AJUSTE_CLI_CREDITO
                
            End Select
        End If
    End With
End Sub

'Private Sub txtCodCliH_Change()
'
'End Sub

'Private Sub txtCodCliD_Change()
'
'End Sub

'Private Sub txtClienteD_Change()
'
'End Sub

'Private Sub txtClienteH_Change()
'
'End Sub

'Private Sub txtcliented_Change()
'
'End Sub
'
'Private Sub txtCodCliD_GotFocus()
'    txtCodCliD.SelStart = 0
'    txtCodCliD.SelLength = Len(txtCodCliD.Text)
'End Sub
'
'Private Sub txtcodclid_LostFocus()
'    If Trim(txtCodCliD) <> "" Then
'        txtClienteD = ObtenerDescripcion("CLIENTES", Val(txtCodCliD))
'    End If
'End Sub
'
'Private Sub txtCodCliH_GotFocus()
'    txtCodCliH.SelStart = 0
'    txtCodCliH.SelLength = Len(txtCodCliH.Text)
'End Sub
'
'Private Sub txtcodclih_LostFocus()
'    If Trim(txtCodCliH) <> "" Then
'        txtClienteH = ObtenerDescripcion("CLIENTES", Val(txtCodCliH))
'    End If
'End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    Dim Consulta As String

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then

        If MsgBox("¿Desea los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
                    " DEBE, HABER, SALDO, '' as [Saldo Final] " & _
                    " From  " & TablaTemp & _
                    " Order By CODIGO_CLI, FECHA, ID"
            CrearConsulta True
            LlenarGrilla grilla, Consulta, False
            grillaMarcoSaldosFinales grilla, 0, 10, 9
            limpioGrilla 10
                    
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO, '' as  Final " & _
                    " From " & TablaTemp & _
                    " Where TIPO_DOCUMENTO <> '' " & _
                    " Order By CODIGO_CLI, FECHA, ID"
            CrearConsulta True
            LlenarGrilla grilla, Consulta, False
            grillaMarcoSaldosFinales grilla, 0, 8, 7
            limpioGrilla 8
        End If
        
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
        Cancel = True
    End If
     
End Sub
Private Function rangoOk() As Boolean
    rangoOk = (uCliD.codigo > 0 And uCliH.codigo > 0)
End Function


