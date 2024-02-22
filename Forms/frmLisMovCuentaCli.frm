VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLisMovCuentaCli 
   Caption         =   "Listado de Movimientos de Cuenta de Clientes"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "frmLisMovCuentaCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMenu 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   780
      Left            =   120
      TabIndex        =   10
      Top             =   7050
      Width           =   11145
      Begin VB.CommandButton cmdResumir 
         Caption         =   "Resumir"
         Height          =   360
         Left            =   5610
         TabIndex        =   26
         Top             =   75
         Width           =   1605
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   45
         Width           =   975
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
         TabIndex        =   13
         Top             =   45
         Width           =   975
      End
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
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Visible         =   0   'False
         Width           =   1395
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   735
         Left            =   3150
         TabIndex        =   11
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1191
      End
   End
   Begin VB.Frame fraGri 
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
      Left            =   105
      TabIndex        =   5
      Top             =   1710
      Width           =   11175
      Begin VB.Frame fraSubGri 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2220
         Left            =   195
         TabIndex        =   17
         Top             =   3015
         Width           =   10725
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   1935
            Left            =   0
            TabIndex        =   18
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
            TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   23
            Top             =   0
            Width           =   2415
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
            TabIndex        =   22
            Top             =   0
            Width           =   1785
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
            TabIndex        =   21
            Top             =   1080
            Width           =   2070
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2775
         Left            =   195
         TabIndex        =   2
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
   Begin VB.Frame Frame1 
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
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8535
      Begin VB.CheckBox chkVer 
         Caption         =   "Ver resaltado"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin Gestion.ucCoDe uCliH 
         Height          =   330
         Left            =   870
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   285
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   556
         CodigoWidth     =   800
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
         TabIndex        =   9
         Top             =   840
         Width           =   1095
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
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1575
      Left            =   8760
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147585
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   142147585
         CurrentDate     =   38252
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
         TabIndex        =   7
         Top             =   840
         Width           =   540
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
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLisMovCuentaCli"
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

Private Const CONST_FACTURAS_FEA = "FEA"
Private Const CONST_FACTURAS_FEB = "FEB"
Private Const CONST_FACTURAS_FEC = "FEC"
Private Const CONST_NOTA_DEA = "DEA"
Private Const CONST_NOTA_DEB = "DEB"
Private Const CONST_NOTA_DEC = "DEC"
Private Const CONST_NOTA_CEA = "CEA"
Private Const CONST_NOTA_CEB = "CEB"
Private Const CONST_NOTA_CEC = "CEC"

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
                (x2s(TipoDocumento) = CONST_FACTURAS_FEA) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_FEB) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_FEC) Or _
                (x2s(TipoDocumento) = CONST_NOTA_DEA) Or _
                (x2s(TipoDocumento) = CONST_NOTA_DEB) Or _
                (x2s(TipoDocumento) = CONST_NOTA_DEC) Then
        VaEnElDebe = True
    Else
        VaEnElDebe = False
    End If
End Function


Private Function CalcularSaldoAnterior(CodigoCliente As Long, fechahasta As Date) As Double

    Dim Debe As Double
    Dim haber As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim Consulta As String

    Debe = 0
    haber = 0
    
    'TABLA FACTURAVENTA
    Consulta = "Select TIPODOC, FORMAPAGO, CONTADO, Sum(TOTAL) as Total From FACTURAVENTA " & _
        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & _
        " Group By TIPODOC, FORMAPAGO, CONTADO"
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
            'pregunto si la forma de pago es contado, porque con esta no hago nada _
            '(ya que debe sumar en el DEBE y restar en el HABER)
            If Not ((x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_FEA Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_FEB Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_FEC) And s2n(rsCuenta!contado) = CONST_CONTADO) Then
                Debe = Debe + s2n(rsCuenta!Total)
                Debug.Print "F+ " & s2n(rsCuenta!Total, 8)
            End If
        Else
            haber = haber + s2n(rsCuenta!Total, 8)
            Debug.Print "F- " & s2n(rsCuenta!Total)
        End If
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close
    Set rsCuenta = Nothing
    
    'TABLA RECIBOS
    Consulta = "Select CLIENTE, SUM(TOTAL) AS TOTAL From RECIBOS " & _
        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & " Group By CLIENTE"
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        haber = haber + s2n(rsCuenta!Total, 8)
        Debug.Print "R- " & s2n(rsCuenta!Total)
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
                saldo = saldo + s2n(!Debe, 4) - s2n(!haber, 4)
                
                If !TIPO_DOCUMENTO <> "" Then
                    'Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & rsAux!ID
                   !saldo = CStr(Round(saldo, 4))
                Else
                    'Consulta = "Update " & TablaTemp & " Set SALDO = ' ' Where ID = " & rsAux!ID
                    !saldo = " "
                End If

                'DataEnvironment1.Sistema.Execute Consulta
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
    Dim NroRem As String
    Dim i As Long
    Dim Res As String
    
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
        
        SaldoCuenta = CalcularSaldoAnterior(CodigoCli, dtfechad.Value)
        
        If SaldoCuenta < 0 Then
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                            "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, 4))) & "', '" & (s2n(SaldoCuenta, 4)) & "')"
        Else
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                            "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                            ", 'SI', '" & (s2n(SaldoCuenta, 4)) & "', '0',  '" & (s2n(SaldoCuenta, 4)) & "')"
        End If
        DataEnvironment1.Sistema.Execute Consulta
        
        SaldoCuenta = 0
        
        
        Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO,resaltar " _
            & " From FACTURAVENTA as F " _
            & " Where ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
            & " Order By F.FECHA, F.CODIGO"
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            If s2n(rs!Remito) = 0 Then
                NroRem = " "
            Else
                NroRem = s2n(rs!Remito)
            End If
            If rs!resaltar Then
                If chkVer.Value = 1 Then
                    Res = "* "
                Else
                    Res = ""
                End If
            Else
                Res = ""
            End If
            If VaEnElDebe(x2s(rs!TIPODOC)) Then
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                        " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                        " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                        ", '" & Res & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(rs!Total, 4) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
            Else
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                        " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                        " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                        ", '" & Res & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(rs!Total, 4) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
            End If
            DataEnvironment1.Sistema.Execute Consulta
                
            'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
            If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E Or x2s(rs!TIPODOC) = CONST_FACTURAS_FEA Or x2s(rs!TIPODOC) = CONST_FACTURAS_FEB Or x2s(rs!TIPODOC) = CONST_FACTURAS_FEC) And rs!contado = CONST_CONTADO Then
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                                            ", 'CON', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(rs!Total, 4) & "', '" & SaldoCuenta & "')"
                DataEnvironment1.Sistema.Execute Consulta
            End If
            rs.MoveNext
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
        
        Consulta = "Select R.* From RECIBOS AS R " _
            & "Where R.ACTIVO = 1 And R.CLIENTE = " & CodigoCli & " AND R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
            & "Order By R.FECHA, R.CODIGO"
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
'            Consulta = "Insert Into " & TablaTemp & _
'                " (CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                " VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!Fecha) & _
'                ", '" & CONST_RECIBOS_IMPUTADOS & "', '" & x2s(rs!numero) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
            
            Consulta = "Insert Into " & TablaTemp & _
                " (CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                " VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!Fecha) & _
                ", '" & rs!TIPODOC & "', '" & x2s(rs!numero) & "', '0', '" & s2n(rs!Total, 4) & "', '" & SaldoCuenta & "')"

            DataEnvironment1.Sistema.Execute Consulta
            Consulta = "Select FACTURAVENTA, IMPORTE From RECIBOSDETALLE " & _
                        "Where CODRECIBO = " & ObtenerDatoDB("RECIBOS", "NUMERO", rs!numero, "CODIGO") & " Order By CODIGO"
            rsFac.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Consulta = "Select NRO, IMPORTE From CHEQUES Where ACTIVO = 1 And TDOC = '" & CONST_RECIBOS & "' And NDOC = " & s2n(rs!numero)
            rsCHQ.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            While Not rsFac.EOF Or Not rsCHQ.EOF
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                                        "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, FACTURAS, CHEQUES) " & _
                                "VALUES (" & CodigoCli & ", '" & cliDes & "', " & ssFecha(rs!Fecha) & _
                                        ", '', '', '', '', '', "
                If Not rsFac.EOF Then
                    rsaux.Open "select * from facturaventa where codigo=" & rsFac!FACTURAVENTA, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    Consulta = Consulta & "'FAC NRO " & x2s(rsaux!NroFactura) & " - " & s2n(rsFac!Importe, 4) & "',"
                    rsFac.MoveNext
                    Set rsaux = Nothing
                    'rsFac!facturaventa
                Else '###0.00 'Format$(x2s(rsFac!importe), "standard")
                    Consulta = Consulta & "'',"
                End If
                
                If Not rsCHQ.EOF Then
                    Consulta = Consulta & "'CHQ NRO " & x2s(rsCHQ!Nro) & " - " & Format(x2s(rsCHQ!Importe), "###0.00") & "')"
                    rsCHQ.MoveNext
                Else
                    Consulta = Consulta & "'')"
                End If
                
                DataEnvironment1.Sistema.Execute Consulta
            Wend


            rsFac.Close
            Set rsFac = Nothing

            rsCHQ.Close
            Set rsCHQ = Nothing
            
            rs.MoveNext
        Wend
        rs.Close
'        Set rs = Nothing
'      End If ' del if existe cliente
'    Next CodigoCli
        rsCli.MoveNext
    Wend
    
    CalcularSaldo
    

    

    Set rs = Nothing
    Set rsCli = Nothing
End Sub

Private Sub cmdAceptar_Click()
Dim i As Long
    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then
                
        relojito
        
        CrearConsulta False
        LimpiarGrilla GRILLA
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaEfectivo
        LimpiarGrilla GrillaMoviCaja
        
'        LlenarGrilla grilla, _
'            "Select CODIGO_CLI AS CODIGO, C.DESCRIPCION AS DESCRIPCION, L.FECHA, L.TIPO_DOCUMENTO, " & _
'            "L.NRO_DOCUMENTO, L.REMITO, L.IVA, L.DEBE, L.HABER, L.SALDO " & _
'            "From " & TablaTemp & " AS L INNER JOIN CLIENTES AS C ON C.CODIGO = L.CODIGO_CLI " & _
'            "Where l.facturas = ' ' and l.cheques = ' ' " & _
'            "Order By CODIGO_CLI, FECHA, ID", True
        LlenarGrilla GRILLA, _
            " Select CODIGO_CLI AS CODIGO,  DESCRIPCION_CLI as DESCRIPCION, FECHA, TIPO_DOCUMENTO, " & _
            " NRO_DOCUMENTO, REMITO, DEBE, HABER, SALDO, '' as [Saldo final] " & _
            " From " & TablaTemp & _
            " Where facturas = ' ' and cheques = ' ' " & _
            " Order By CODIGO_CLI, FECHA, ID", True
        grillaMarcoSaldosFinales GRILLA, 0, 9, 8
        
        For i = 1 To GRILLA.rows - 1
            GRILLA.TextMatrix(i, 6) = s2n(GRILLA.TextMatrix(i, 6), 2)
            GRILLA.TextMatrix(i, 7) = s2n(GRILLA.TextMatrix(i, 7), 2)
            GRILLA.TextMatrix(i, 8) = s2n(GRILLA.TextMatrix(i, 8), 2)
        Next
        
        
        relojito False
        
    End If
End Sub



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
        
        
        VinculoXl "C:\MovCuentaCli.xls", "Saldo a cuenta de Cliente/s", , , rs
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

Private Sub cmdCancelar_Click()
    Dim tempo
    tempo = obtenerDeSQL("select min(codigo) as mini, max(codigo) as maxi from clientes where activo =1")
    'txtcodclih = tempo(1) ' "9999999"
    uCliH.codigo = tempo(1)
'    txtclienteh = ""
    'txtcodclid = tempo(0)
    uCliD.codigo = tempo(0)
'    txtcliented = "" 'tempo(1)
    dtfechad.Value = "01/01/" & Year(Date)
    dtfechah.Value = Date
End Sub

Private Sub cmdResumir_Click()
Dim i As Long, esto As Long, saldofinal As Double
With GRILLA
    If .rows > 0 Then
    esto = 1
    saldofinal = 0
devuelta:
        For i = esto To .rows - 1
            If s2n(.TextMatrix(i, .cols - 1)) <> 0 Then
                .TextMatrix(i, 2) = dtfechah
                .TextMatrix(i, 3) = "RESUMEN"
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 5) = "saldo"
                .TextMatrix(i, 6) = 0
                .TextMatrix(i, 7) = 0
                .TextMatrix(i, 8) = 0
                saldofinal = saldofinal + s2n(.TextMatrix(i, .cols - 1))
            Else
                esto = i
                .RemoveItem i
                GoTo devuelta
            End If
        Next
        .AddItem ""
        .TextMatrix(.rows - 1, .cols - 1) = s2n(saldofinal)
    End If
End With
End Sub

Private Sub cmdSalir_Click()
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
    
    ucXls1.ini GRILLA, "C:\MovCuentaCli", "Mov cuenta cliente"
    
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
    
    cmdCancelar_Click
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
    Anclar fraGri, Me, anclarLadosTodos
    Anclar GRILLA, fraGri, anclarLadosTodos
    Anclar fraSubGri, fraGri, anclarIzquierda + anclarAbajo
End Sub

Private Sub grilla_Click()
Dim TIPODOC As String
Dim NroDoc As Long
Dim CodInt As Long
Dim cli As Long

    With GRILLA
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(Replace(Trim(.TextMatrix(.Row, 3)), "*", ""))
            cli = .TextMatrix(.Row, 0)
            If .TextMatrix(.Row, 4) <> "" Then NroDoc = CLng(.TextMatrix(.Row, 4))
            
            LimpiarGrilla GrillaDetalle
            LimpiarGrilla GrillaMoviCaja
            LimpiarGrilla GrillaEfectivo
            
            Select Case TIPODOC
                Case CONST_FACTURAS_A, CONST_FACTURAS_B, CONST_FACTURAS_E, CONST_FACTURAS_FEA, CONST_FACTURAS_FEB, CONST_FACTURAS_FEC
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
                                            " left join facturaventa as f on f.codigo=d.codigofactura " & _
                                         "Where D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & " and f.cliente=" & cli & _
                                         " Order By ID", False



                Case CONST_RECIBOS, CONST_RECIBOS_IMPUTADOS
                    CodInt = ObtenerDatoDB2("RECIBOS", "NUMERO", NroDoc, "CODIGO")
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

Private Sub ucXls1_ClicQUITADO(cancel As Boolean)
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
            LlenarGrilla GRILLA, Consulta, False
            grillaMarcoSaldosFinales GRILLA, 0, 10, 9
                    
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO, '' as  Final " & _
                    " From " & TablaTemp & _
                    " Where TIPO_DOCUMENTO <> '' " & _
                    " Order By CODIGO_CLI, FECHA, ID"
            CrearConsulta True
            LlenarGrilla GRILLA, Consulta, False
            grillaMarcoSaldosFinales GRILLA, 0, 8, 7
        End If
        
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
        cancel = True
    End If
     
End Sub
Private Function rangoOk() As Boolean
    rangoOk = (uCliD.codigo > 0 And uCliH.codigo > 0)
End Function
