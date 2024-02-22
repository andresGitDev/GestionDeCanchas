VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSaldoCuentaCli 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Saldo de Cuenta de Clientes"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12780
   Icon            =   "frmSaldoCuentaCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar por diferencias"
      Height          =   195
      Left            =   6720
      TabIndex        =   31
      Top             =   640
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar centavos"
      Height          =   195
      Left            =   5280
      TabIndex        =   30
      Top             =   640
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenar por Descripcion de cliente"
      Height          =   195
      Left            =   2400
      TabIndex        =   25
      Top             =   620
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenar por nro de cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   620
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Frame Framedevol 
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
      ForeColor       =   &H00400000&
      Height          =   600
      Left            =   60
      TabIndex        =   6
      Top             =   -45
      Width           =   4005
      Begin VB.OptionButton opttodos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos los Clientes"
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
         Left            =   1845
         TabIndex        =   8
         Tag             =   "1"
         Top             =   255
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton optuno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Elegir Cliente"
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
         Left            =   135
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraGri 
      BackColor       =   &H00E0E0E0&
      Height          =   6060
      Left            =   60
      TabIndex        =   4
      Top             =   825
      Width           =   12645
      Begin VB.Frame fraSubGrillas 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2580
         Left            =   120
         TabIndex        =   16
         Top             =   3435
         Width           =   11730
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   2310
            Left            =   15
            TabIndex        =   17
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   195
            Width           =   6090
            _cx             =   10742
            _cy             =   4075
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
            Height          =   1035
            Left            =   6300
            TabIndex        =   18
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   180
            Width           =   5745
            _cx             =   10134
            _cy             =   1826
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
            Height          =   1125
            Left            =   6330
            TabIndex        =   19
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   1395
            Width           =   5700
            _cx             =   10054
            _cy             =   1984
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento de Cheques"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6315
            TabIndex        =   22
            Top             =   -15
            Width           =   1710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comprobantes que imputa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   15
            TabIndex        =   21
            Top             =   -15
            Width           =   1890
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento en Efectivo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   6345
            TabIndex        =   20
            Top             =   1200
            Width           =   1665
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   3075
         Left            =   135
         TabIndex        =   5
         Top             =   330
         Width           =   12405
         _cx             =   21881
         _cy             =   5424
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
      Begin VB.Label lblTel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefonos:"
         Height          =   315
         Left            =   1110
         TabIndex        =   23
         Top             =   105
         Width           =   5010
      End
   End
   Begin VB.Frame fraCliente 
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
      ForeColor       =   &H00800000&
      Height          =   645
      Left            =   4200
      TabIndex        =   0
      Top             =   -45
      Visible         =   0   'False
      Width           =   8445
      Begin VB.TextBox txtDescCliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   5475
      End
      Begin VB.CommandButton cmdAyudaCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cliente"
         Height          =   285
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "2"
         Top             =   255
         Width           =   750
      End
      Begin VB.TextBox txtCodCliente 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   8760
      TabIndex        =   26
      Top             =   640
      Width           =   3855
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todo"
         Height          =   195
         Left            =   3120
         TabIndex        =   29
         Top             =   0
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Positivos mayores a $10"
         Height          =   195
         Left            =   1080
         TabIndex        =   28
         Top             =   0
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Negativos"
         Height          =   195
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   60
      TabIndex        =   9
      Top             =   6900
      Width           =   11820
      Begin VB.CheckBox chkVer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Resaltado"
         Height          =   315
         Left            =   4080
         TabIndex        =   33
         Top             =   120
         Width           =   2415
      End
      Begin VB.CheckBox chkNoMostrarSaldos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No mostrar saldo total"
         Height          =   375
         Left            =   4080
         TabIndex        =   32
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
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
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Visible         =   0   'False
         Width           =   1455
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
         Height          =   375
         Left            =   10755
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   60
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
         Height          =   375
         Left            =   8580
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
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
         Height          =   375
         Left            =   9570
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   975
      End
      Begin Gestion.ucXls uXls 
         Height          =   825
         Left            =   2640
         TabIndex        =   10
         Top             =   30
         Width           =   870
         _extentx        =   1535
         _extenty        =   1455
      End
      Begin VB.Label lblSaldoTotal 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   34
         Top             =   480
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmSaldoCuentaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Const CONST_NOTAS_DEBITOS_E = "NDE"
Private Const CONST_NOTAS_CREDITOS_E = "NCE"

Private Const CONST_FACTURAS_FEA = "FEA"
Private Const CONST_FACTURAS_FEB = "FEB"
Private Const CONST_FACTURAS_FEC = "FEC"
Private Const CONST_NOTA_DEA = "DEA"
Private Const CONST_NOTA_DEB = "DEB"
Private Const CONST_NOTA_DEC = "DEC"
Private Const CONST_NOTA_CEA = "CEA"
Private Const CONST_NOTA_CEB = "CEB"
Private Const CONST_NOTA_CEC = "CEC"

Private Const CONST_RECIBOS = "RAA" 'RECIBOS A CUENTA

Private Const CONST_CONTADO = 1

Private Enum gri_lla
    griCLCO
    griCLDE
    griFECH
    griTDOC
    griPUNTO
    griNDOC
    griBSAL
    griVENC
    griDEBE
    griHABE
    griSALD
End Enum


Dim SaldoCero As Boolean
Dim SaldoTotal As Double

Private Sub CalcularSaldo()
    Dim rsaux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Double
    Dim CodigoCli As Long
    Dim CodigoCliActual As Long
    Dim ide As Long

    Consulta = "Select * From LIST_SALDO_CLI Order By CLIENTE, FECHA_DOC, ID"
    rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rsaux.EOF Then
        rsaux.MoveFirst
        SaldoCero = False
    Else
        SaldoCero = optUno.Value
    End If
    SaldoTotal = 0
    While Not rsaux.EOF
        CodigoCli = rsaux!cliente
        CodigoCliActual = CodigoCli
        saldo = 0
        While CodigoCli = CodigoCliActual
            If Not IsNull(rsaux!Debe) And Not IsNull(rsaux!haber) Then saldo = s2n(saldo + s2n(rsaux!Debe) - s2n(rsaux!haber))
            Consulta = "Update LIST_SALDO_CLI Set SALDO = " & x2s(saldo) & " Where ID = " & rsaux!ID
            DataEnvironment1.Sistema.Execute Consulta
            ide = rsaux!ID
            rsaux.MoveNext
            If rsaux.EOF Then
                CodigoCliActual = 0
            Else
                CodigoCliActual = rsaux!cliente
            End If
        Wend
        
        Consulta = "Update LIST_SALDO_CLI Set SALDO2 = '" & Format(saldo, "#,##0.00") & "' Where ID = " & ide
        DataEnvironment1.Sistema.Execute Consulta
        
        SaldoTotal = SaldoTotal + saldo
        lblSaldoTotal.caption = "El Saldo Total es: $ " & Format(SaldoTotal, "#,##0.00")
    Wend
End Sub

Private Sub CrearReporte()
Dim rsSaldo As New ADODB.Recordset
Dim Consulta As String
Dim moneda As String
Dim Tipo As String
Dim str As String
Dim Res As String
Dim Obs As String

    DataEnvironment1.Sistema.Execute "Delete From LIST_SALDO_CLI"
    
    If Check2.Value = 1 Then
        str = " (saldo<0 or saldo>0.10) "
    Else
        str = " Saldo <> 0 "
    End If
    If Check1.Value = 1 Then
        str = str & " and not(total>=0 and total<=1) "
    End If
    
    Consulta = "Select CODIGO, TIPODOC,PUNTOVENTA, NROFACTURA, FECHA, VENCIMIENTO, CLIENTE, SALDO, total,moneda,cotizacion,resaltar " & _
        " From FACTURAVENTA " & _
        " Where " & str & " and Activo = 1"
    If optUno.Value Then
        If txtCodCliente.Text <> "" Then
            Consulta = Consulta & " and CLIENTE = " & s2n(txtCodCliente.Text)
        Else
            MsgBox "Debe ingresar un Cliente", vbOKOnly, "Atencion"
            Exit Sub
        End If
    End If
    
    Consulta = Consulta & " Order By CLIENTE, FECHA, CODIGO"
            
    rsSaldo.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rsSaldo.EOF Then rsSaldo.MoveFirst
    While Not rsSaldo.EOF
        moneda = Trim(IIf((rsSaldo!moneda) > 1, "U$S Cotiz. $" & rsSaldo!cotizacion, ""))
        Tipo = IIf(Trim(x2s(rsSaldo!TIPODOC)) = "RAA", "RCA", Trim(x2s(rsSaldo!TIPODOC)))
        If rsSaldo!resaltar Then
            If chkVer.Value = 1 Then
                Res = " *** A *** "
            Else
                Res = ""
            End If
        Else
            Res = ""
        End If
        Obs = ""
        Obs = IIf(rsSaldo!Total <> rsSaldo!saldo, "'" & Res & "(saldo)" & moneda & "'", "' " & Res & moneda & "'")
        
        Consulta = "Insert Into LIST_SALDO_CLI (CLIENTE, CODIGO_DOC, FECHA_DOC, TIPO_DOC, " & _
                                                "NRO_DOC, VENCIMIENTO, obs, DEBE, HABER, SALDO,PUNTOVENTA) " & _
                        "Values (" & s2n(rsSaldo!cliente) & ", " & s2n(rsSaldo!codigo) & _
                            ", " & ssFecha(rsSaldo!Fecha) & ", '" & Tipo & _
                            "', '" & x2s(rsSaldo!NroFactura) & "', " & IIf(Trim(rsSaldo!TIPODOC) = "RAA", ssFecha(rsSaldo!Fecha), ssFecha(rsSaldo!Vencimiento)) & _
                            ", " & Obs
        If VaEnElDebe(x2s(rsSaldo!TIPODOC)) Then
            Consulta = Consulta & ", '" & s2n(rsSaldo!saldo, 2, True) & "', '0', '0'," & ssTexto(rsSaldo!PuntoVenta) & ")"
        Else
            Consulta = Consulta & ", '0', '" & s2n(rsSaldo!saldo, 2, True) & "', '0'," & ssTexto(rsSaldo!PuntoVenta) & ")"
        End If
        
        DataEnvironment1.Sistema.Execute Consulta
        
        rsSaldo.MoveNext
    Wend
    
    CalcularSaldo
    
    rsSaldo.Close
    Set rsSaldo = Nothing
End Sub

Private Sub HabilitarControles(habilitar As Boolean)
    fraCliente.Visible = habilitar
'    lblCliente.Visible = habilitar
'    txtCodCliente.Visible = habilitar
'    cmdAyudaCliente.Visible = habilitar
'    txtDescCliente.Visible = habilitar
End Sub

Private Sub LimpiarControles()
    optUno.Value = True
    txtCodCliente.Text = ""
    txtDescCliente.Text = ""
    HabilitarControles True
End Sub

Private Function VaEnElDebe(TipoDocumento As String) As Boolean
    'funcion que devuelve TRUE si el tipo de comprobante va en el DEBE o en el HABER
    If (x2s(TipoDocumento) = CONST_AJUSTE_CLI_DEBITO) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_A) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_B) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_E) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_E) Or _
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

Private Sub cmdAceptar_Click()
    Dim Orden As String, ssaldo As String
    Dim str As String
    Dim i As Long
    If Option1.Value = True Then
        Orden = " CLIENTE, "
    Else
        Orden = " descripcion, "
    End If
    
    CrearReporte
    
    If Option4.Value = True Then
        str = "delete from list_saldo_cli where cliente in (select cliente FROM LIST_SALDO_CLI where id in (SELECT max(id) FROM LIST_SALDO_CLI group by cliente) and saldo<=10)"
        DataEnvironment1.Sistema.Execute str
    ElseIf Option3.Value = True Then
        str = "delete from list_saldo_cli where cliente in (select cliente FROM LIST_SALDO_CLI where id in (SELECT max(id) FROM LIST_SALDO_CLI group by cliente) and saldo>10)"
        DataEnvironment1.Sistema.Execute str
    End If
    'elimino los comprobantes con centavos por pedido de claudio
    'STR = "delete from list_saldo_cli where cliente in (select cliente FROM LIST_SALDO_CLI where id in (SELECT max(id) FROM LIST_SALDO_CLI group by cliente) and (saldo>=0 and saldo<=1))"
    'DataEnvironment1.Sistema.Execute STR
    
    If chkNoMostrarSaldos Then
        lblSaldoTotal.Visible = False
    Else
        lblSaldoTotal.Visible = True
    End If
    ssaldo = " SALDO "
    
    
    LimpiarGrilla GRILLA
    LlenarGrilla GRILLA, " Select CLIENTE, DESCRIPCION, FECHA_DOC as 'FECHA', TIPO_DOC as 'DOCUMENTO',PUNTOVENTA,NRO_DOC AS 'NUMERO DOC.', " & _
        " LIST_SALDO_CLI.obs as obs, VENCIMIENTO, DEBE, HABER, " & ssaldo & ", '' as [Saldo Final] " & _
        " From LIST_SALDO_CLI Inner Join CLIENTES On CLIENTES.CODIGO = LIST_SALDO_CLI.CLIENTE " & _
        " Order By " & Orden & " FECHA_DOC, ID", False ', 1
    grillaMarcoSaldosFinales GRILLA, 0, 11, 10
    grillaWidth GRILLA, Array(780, 2040, 1125, 550, 950, 950, 1750, 1110, 1200, 1200, 1200, 1200)
    GRILLA.ColAlignment(0) = flexAlignCenterCenter
    If GRILLA.rows > 1 Then GRILLA.ColHidden(10) = True
    i = 1
    While i < GRILLA.rows
        If InStr(1, Trim(GRILLA.TextMatrix(i, 6)), "*** A ***") > 0 Then
            GRILLA.Col = 6
            GRILLA.Row = i
            GRILLA.CellFontBold = True
        End If
        i = i + 1
    Wend
End Sub

Private Sub cmdAyudaCliente_Click()
    frmBuscar.MostrarSql "Select CODIGO, DESCRIPCION From CLIENTES Order By CODIGO", , , , , , True
    If frmBuscar.resultado <> "" Then
        txtCodCliente.Text = frmBuscar.resultado
        txtDescCliente.Text = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Consulta As String
    Dim str As String
        
    CrearReporte
    
    If Option4.Value = True Then
        str = "delete from list_saldo_cli where cliente in (select cliente FROM LIST_SALDO_CLI where id in (SELECT max(id) FROM LIST_SALDO_CLI group by cliente) and saldo<=10)"
        DataEnvironment1.Sistema.Execute str
    ElseIf Option3.Value = True Then
        str = "delete from list_saldo_cli where cliente in (select cliente FROM LIST_SALDO_CLI where id in (SELECT max(id) FROM LIST_SALDO_CLI group by cliente) and saldo>10)"
        DataEnvironment1.Sistema.Execute str
    End If
    
    If Not SaldoCero Then
        If Option2.Value = True Then
            rptListSaldoCliente.DataMember = "List_Saldo_Cli2"
            rptListSaldoCliente.Sections("section0").Controls("text1").DataMember = "List_Saldo_Cli2"
            rptListSaldoCliente.Sections("section0").Controls("text2").DataMember = "List_Saldo_Cli2"
            rptListSaldoCliente.Sections("section0").Controls("texto3").DataMember = "List_Saldo_Cli2"
            rptListSaldoCliente.Sections("section1").Controls("txtFecha").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtTipoDoc").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtNroDoc").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtVencimiento").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtDebe").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtHaber").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("txtSaldo").DataMember = "List_Saldo_Cli_detalle2"
            If chkNoMostrarSaldos Then
                rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").Visible = False
            Else
                rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").Visible = True
            End If
            
            rptListSaldoCliente.Sections("section1").Controls("texto1").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section1").Controls("texto2").DataMember = "List_Saldo_Cli_detalle2"
            'rptListSaldoCliente.Sections("section1").Controls("texto3").DataMember = "List_Saldo_Cli_detalle2"
            rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").caption = lblSaldoTotal.caption
            Printer.PaperSize = vbPRPSA4 'hoja A4=9
            DataEnvironment1.rsList_Saldo_Cli2.Open
            rptListSaldoCliente.Sections("Section4").Controls("Label13").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
            rptListSaldoCliente.Show vbModal
            DataEnvironment1.rsList_Saldo_Cli2.Close
        Else
            rptListSaldoCliente.DataMember = "List_Saldo_Cli"
            
            Printer.PaperSize = vbPRPSA4 'hoja A4=9
            DataEnvironment1.rsList_Saldo_Cli.Open
            If chkNoMostrarSaldos Then
                rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").Visible = False
            Else
                rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").Visible = True
            End If
            
            rptListSaldoCliente.Sections("Section4").Controls("Label13").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
            rptListSaldoCliente.Sections("section5").Controls("Etiqueta12").caption = lblSaldoTotal.caption
            rptListSaldoCliente.Show vbModal
            DataEnvironment1.rsList_Saldo_Cli.Close
        End If
    Else
        MsgBox "El Cliente elegido no tiene saldo.", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
End Sub

Private Sub cmdexcel_Click()
    Dim rs As New ADODB.Recordset
    Dim Consulta As String

    CrearReporte

    Consulta = "Select CLIENTE, FECHA_DOC as 'FECHA', TIPO_DOC as 'DOCUMENTO', NRO_DOC as 'NUMERO DOC.', " & _
                        "obs, VENCIMIENTO, DEBE, HABER, SALDO From LIST_SALDO_CLI "

    Consulta = Consulta & " Order By CLIENTE, FECHA_DOC, ID"


    rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    VinculoXl "C:\SaldoCli.xls", "Saldo a cuenta de Cliente/s", , , rs
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uXls.ini GRILLA, "C:\SaldoCliente", "Saldo Clientes   " & Date
    uXls.caption = "Grilla a XLS"
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraGri, Me, anclarLadosTodos
    Anclar fraSubGrillas, Me, anclarAbajo + anclarLadosAncho
    Anclar GRILLA, fraGri, anclarLadosTodos
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
End Sub

'Private Sub Form_Load()
'    CentrarMe Me
'End Sub

Private Sub grilla_Click()
    Dim TIPODOC As String
    Dim PuntoVta As String
    Dim NroDoc As Long
    Dim CodInt As Long
    
    Dim r As Long, C As Long
    Dim clicod As Long
    
    With GRILLA
    
        r = .Row
        C = .Col
        clicod = s2n(.TextMatrix(r, griCLCO))
        ' sin clinte
        If clicod = 0 Then Exit Sub
        
        
        'telefono
        lblTel = sSinNull(obtenerDeSQL("select telefono from clientes where codigo = " & clicod))
        TIPODOC = Trim(Replace(Trim(.TextMatrix(r, griTDOC)), "*", ""))
        PuntoVta = Trim(.TextMatrix(r, griPUNTO))
        NroDoc = s2n(.TextMatrix(r, griNDOC))  '  ??? If .TextMatrix(.Row, 5) <> "" Then
            
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaMoviCaja
        LimpiarGrilla GrillaEfectivo
            
        Select Case TIPODOC
         Case CONST_FACTURAS_A, CONST_FACTURAS_B, CONST_FACTURAS_E, CONST_FACTURAS_FEA, CONST_FACTURAS_FEB, CONST_FACTURAS_FEC
            LlenarGrilla GrillaDetalle, _
                " Select D.CANTIDAD, D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO AS 'PRECIO UNITARIO', " & _
                " D.PRECIOTOTAL AS 'PRECIO TOTAL', S.SERIE " & _
                " From (facturaventa as f inner join FACTURAVENTADETALLE AS D on f.codigo=d.codigofactura) " & _
                " left Join SERIES as S ON S.PRODUCTO=D.PRODUCTO AND S.NROCOMPROBANTE=D.NROFACTURA " & _
                " Where f.puntoventa=" & ssTexto(PuntoVta) & " AND D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & _
                " Order By ID", False

         Case CONST_RECIBOS  ' el recibo no va a aparecer aca...!!
            CodInt = ObtenerDatoDB("RECIBOS", "NUMERO", NroDoc, "CODIGO")
            LlenarGrilla GrillaDetalle, _
                " Select FACTURAVENTA AS 'Factura que imputa', Importe From RECIBOSDETALLE " & _
                " Where CODRECIBO = " & CodInt & " Order By CODIGO", False
            LlenarGrilla GrillaMoviCaja, "Select Fecha, nro as 'Numero Cheque', Importe From CHEQUES " & _
                " Where ACTIVO = 1 And TDOC = '" & TIPODOC & "' AND NDOC = " & NroDoc, True
            LlenarGrilla GrillaEfectivo, "Select Fecha, Importe From MOVICAJA " & _
                " Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                                    
         Case CONST_NOTAS_DEBITOS_A, CONST_NOTAS_CREDITOS_A, CONST_NOTAS_CREDITOS_B
                
         Case CONST_AJUSTE_CLI_DEBITO, CONST_AJUSTE_CLI_CREDITO
                
        End Select
    End With
End Sub

Private Sub opttodos_Click()
    HabilitarControles False
End Sub

Private Sub optuno_Click()
    HabilitarControles True
End Sub

Private Sub txtCodCliente_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtCodCliente_LostFocus()
    If Trim(txtCodCliente) <> "" Then
        txtDescCliente = ObtenerDescripcion("CLIENTES", val(txtCodCliente))
    Else
        txtDescCliente = ""
    End If
End Sub

