VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLismovCli3 
   Caption         =   "Composicion de clientes por periodo"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   3495
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Saldo"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Con Saldo"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   420
      Left            =   135
      TabIndex        =   19
      Top             =   6930
      Width           =   11145
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Height          =   360
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   45
         Width           =   1275
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   360
         Left            =   3150
         TabIndex        =   25
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
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
      Top             =   1590
      Width           =   12510
      Begin VB.Frame fraSubGri 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   660
         Left            =   195
         TabIndex        =   11
         Top             =   4560
         Visible         =   0   'False
         Width           =   10725
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   255
            Left            =   0
            TabIndex        =   12
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   240
            Visible         =   0   'False
            Width           =   5415
            _cx             =   9551
            _cy             =   450
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
            Height          =   255
            Left            =   5520
            TabIndex        =   13
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   240
            Visible         =   0   'False
            Width           =   5415
            _cx             =   9551
            _cy             =   450
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
            Height          =   255
            Left            =   5520
            TabIndex        =   14
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   720
            Visible         =   0   'False
            Width           =   5415
            _cx             =   9551
            _cy             =   450
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
            TabIndex        =   17
            Top             =   0
            Visible         =   0   'False
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
            TabIndex        =   16
            Top             =   0
            Visible         =   0   'False
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
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   2070
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   885
         Left            =   195
         TabIndex        =   18
         ToolTipText     =   "Haga Click para ver el Detalle de la Orden de Compra"
         Top             =   4155
         Visible         =   0   'False
         Width           =   10860
         _cx             =   19156
         _cy             =   1561
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
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   4515
         Left            =   105
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   615
         Width           =   12330
         _cx             =   21749
         _cy             =   7964
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLismovCli3.frx":0000
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
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VENCIDO"
         Height          =   255
         Left            =   3435
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   3840
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A VENCER"
         Height          =   255
         Left            =   7275
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   3870
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
      TabIndex        =   5
      Top             =   0
      Width           =   8535
      Begin Gestion.ucCoDe uCliH 
         Height          =   330
         Left            =   870
         TabIndex        =   6
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
         TabIndex        =   7
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
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin MSComCtl2.DTPicker dtFechaD 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72482817
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   72482817
         CurrentDate     =   39347
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
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "a la fecha:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmLismovCli3"
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

Private Function VaEnElDebe2(TipoDocumento As String) As Boolean
'funcion que devuelve TRUE si el tipo de comprobante va en el DEBE o en el HABER
    If (Trim(TipoDocumento) = CONST_AJUSTE_CLI_DEBITO) Or _
                (Trim(TipoDocumento) = CONST_FACTURAS_A) Or _
                (Trim(TipoDocumento) = CONST_FACTURAS_E) Or _
                (Trim(TipoDocumento) = CONST_FACTURAS_B) Or _
                (Trim(TipoDocumento) = CONST_NOTAS_DEBITOS_A) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_B) Or _
                (x2s(TipoDocumento) = "FAAV") Or _
                (x2s(TipoDocumento) = "FABV") Or _
                (x2s(TipoDocumento) = "NDAV") Then
        VaEnElDebe2 = True
    Else
        VaEnElDebe2 = False
    End If
End Function

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
    'Set rsCuenta = Nothing
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween(fechahasta, dtfechad.Value) & " and activo=1 and r.cliente=" & CodigoCliente, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsaux.EOF = True And rsaux.BOF = True Then
            If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                'pregunto si la forma de pago es contado, porque con esta no hago nada _
                '(ya que debe sumar en el DEBE y restar en el HABER)
                If Not ( _
                    (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                    And s2n(rsCuenta!contado) = CONST_CONTADO) Then
                    Debe = s2n(Debe + s2n(rsCuenta!Total))
                    'If debe = 2577655.64 Then
                    'Stop
                    'End If
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
                        'If debe > 280000 Then
                        'Stop
                        'End If
                        
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


Private Function CalcularSaldoAnterior2(CodigoCliente As Long, fechahasta As Date, Optional fec As Date, Optional recibo As Boolean) As Double

'    Dim debe As Double
'    Dim haber As Double
'    Dim rsCuenta As New ADODB.Recordset
'    Dim rsaux As New ADODB.Recordset
'    Dim sal As Double
'    Dim tot As Double
'    Dim Consulta As String
'    Dim we As String
'    Dim we2 As String
'    Dim saldoViej As Double

'    debe = 0
'    haber = 0
'
'    'If fec = "01/01/1900" Then
'    '    we = " and f.fecha>=" & ssFecha("22/04/2006")
'    '    we2 = " and fecha>=" & ssFecha("22/04/2006")
'    'Else
'    '    we = " and vencimiento>=" & ssFecha(fec)
'    '    we2 = " and fecha>=" & ssFecha(fec)
'    'End If
'
'    'TABLA FACTURAVENTA
''    Consulta = "Select TIPODOC, FORMAPAGO, CONTADO, Sum(TOTAL) as Total,codigo From FACTURAVENTA " & _
''        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & _
''        " Group By TIPODOC, FORMAPAGO, CONTADO,codigo"
'    Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
'                & " From FACTURAVENTA as F " _
'                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCliente & " And F.fecha<" & ssFecha(fechahasta) & we _
'                & " Order By F.FECHA, F.CODIGO"
'
'    rsCuenta.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
'    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
'   While Not rsCuenta.EOF
'        'consulta = "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween(dtfechad.Value, dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCliente
'        rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween("01/01/1900", dtfechad.Value) & " and activo=1 and r.cliente=" & CodigoCliente, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
'        If rsaux.EOF = True And rsaux.BOF = True Then
'            If VaEnElDebe(Trim(rsCuenta!TIPODOC)) Then
'                'pregunto si la forma de pago es contado, porque con esta no hago nada _
'                '(ya que debe sumar en el DEBE y restar en el HABER)
'                If Not ( _
'                    (Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
'                    And (rsCuenta!contado) = CONST_CONTADO) Then debe = debe + s2n(rsCuenta!Total)
'            Else
'                haber = haber + s2n(rsCuenta!Total)
'            End If
'            rsCuenta.MoveNext
'        Else
'            sal = 0
'            While Not rsaux.EOF
'                sal = sal + rsaux!importe
'                rsaux.MoveNext
'            Wend
'            If s2n(sal) = s2n(rsCuenta!Total) Then
'                rsCuenta.MoveNext
'            Else
'                sal = rsCuenta!Total - sal
'
'                If VaEnElDebe(Trim(rsCuenta!TIPODOC)) Then
'                    'pregunto si la forma de pago es contado, porque con esta no hago nada _
'                    '(ya que debe sumar en el DEBE y restar en el HABER)
'                    If Not ( _
'                        (Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or Trim(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
'                        And (rsCuenta!contado) = CONST_CONTADO) Then debe = debe + s2n(sal)
'                Else
'                    haber = haber + s2n(sal)
'                End If
'                rsCuenta.MoveNext
'            End If '

'        End If
'        Set rsaux = Nothing
'    Wend
'    rsCuenta.Close
'    Set rsCuenta = Nothing
    
'    'TABLA RECIBOS
'    If recibo Then
'        saldoViej = s2n(SaldoViejo(CodigoCliente, "22/04/2006"))  'es la fecha de migracion de sistema
'        CalcularSaldoAnterior = debe - haber
'        CalcularSaldoAnterior = CalcularSaldoAnterior + saldoViej
        
'        Consulta = "Select CLIENTE, SUM(TOTAL) AS TOTAL From RECIBOS " & _
'            " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & we2 & " Group By CLIENTE"
'        rsCuenta.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
'        If Not rsCuenta.EOF Then rsCuenta.MoveFirst
'        While Not rsCuenta.EOF
'            haber = haber + s2n(rsCuenta!Total)
'             rsCuenta.MoveNext
'        Wend
'        rsCuenta.Close
'        Set rsCuenta = Nothing
'    Else
'        CalcularSaldoAnterior = debe - haber
'    End If
'        CalcularSaldoAnterior = debe - haber
        
        
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
                    And s2n(rsCuenta!contado) = CONST_CONTADO) Then Debe = Debe + s2n(rsCuenta!Total)
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
                        And s2n(rsCuenta!contado) = CONST_CONTADO) Then Debe = Debe + s2n(sal)
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
    
    CalcularSaldoAnterior2 = Debe - haber
        
        
        
End Function

Private Function SaldoViejo(CodigoCliente As Long, fechahasta As Date) As Double

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
        rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween(fechahasta, dtfechad.Value) & " and activo=1 and r.cliente=" & CodigoCliente, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsaux.EOF = True And rsaux.BOF = True Then
            If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                'pregunto si la forma de pago es contado, porque con esta no hago nada _
                '(ya que debe sumar en el DEBE y restar en el HABER)
                If Not ( _
                    (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                    And s2n(rsCuenta!contado) = CONST_CONTADO) Then Debe = Debe + s2n(rsCuenta!Total)
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
                        And s2n(rsCuenta!contado) = CONST_CONTADO) Then Debe = Debe + s2n(sal)
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
    
    SaldoViejo = Debe - haber
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

                'DataEnvironment1.AMR.Execute Consulta
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
Private Function ssaldo(SaldoCuenta As Double, CodigoCli As Long, cliDes As String, fec As Date)
    Dim Consulta As String
    
    If Option3.Value = True Then ' sin saldo =0
            If SaldoCuenta = 0 Then
            Else
                If SaldoCuenta < 0 Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(fec) & _
                                            ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, 2))) & "', '" & (s2n(SaldoCuenta, 2)) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(fec) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, 2)) & "', '0',  '" & (s2n(SaldoCuenta, 2)) & "')"
                End If
                DataEnvironment1.Sistema.Execute Consulta
            End If
        End If
        If Option4.Value = True Then
            If SaldoCuenta < 0 Then
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(fec) & _
                                    ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, 2))) & "', '" & (s2n(SaldoCuenta, 2)) & "')"
            Else
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(fec) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, 2)) & "', '0',  '" & (s2n(SaldoCuenta, 2)) & "')"
            End If
            DataEnvironment1.Sistema.Execute Consulta
        End If
End Function

Private Sub CrearConsulta(ConDetalle As Boolean)
    Dim SaldoCuenta As Double
    Dim SaldoCuenta2 As Double
    Dim SaldoCuenta3 As Double
    Dim SaldoCuenta4 As Double
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
        
        'SaldoCuenta = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value, dtfechad.Value - 90))           '0 a 3 meses
        'SaldoCuenta2 = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value - 90, dtfechad.Value - 180))  '3 a 6 meses
        'SaldoCuenta3 = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value - 180, dtfechad.Value - 365)) '6 a 12 meses
        'SaldoCuenta4 = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value - 365, "01/01/1900", True))  'mas de 12 meses
        
        'ssaldo SaldoCuenta, CodigoCli, cliDes, dtfechad.Value - 50
        'ssaldo SaldoCuenta2, CodigoCli, cliDes, dtfechad.Value - 100
        'ssaldo SaldoCuenta3, CodigoCli, cliDes, dtfechad.Value - 200
        'ssaldo SaldoCuenta4, CodigoCli, cliDes, dtfechad.Value - 400
        
        
        SaldoCuenta = 0
        'If VerParametro(BS_NOMBRE_EMPRESA) = "nimisan swartz" Then
        '    Consulta = "Select F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
        '        & " From FACTURAVENTA as F " _
        '        & " Where ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
        '        & " Order By F.FECHA, F.CODIGO"
        'Else
            Consulta = "Select DISTINCT F.CODIGO, F.vencimiento, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
                & " From FACTURAVENTA as F " _
                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.vencimiento " & ssBetween(dtfechad.Value, dtfechah.Value) & " and total>0" _
                & " Order By F.vencimiento, F.CODIGO"
        'End If
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            'consulta = "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli
            rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<=" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If rsaux.EOF = True And rsaux.BOF = True Then
            
                If s2n(rs!Remito) = 0 Then
                    NroRem = " "
                Else
                    NroRem = s2n(rs!Remito)
                End If
                If VaEnElDebe(x2s(rs!TIPODOC)) Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Vencimiento) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(rs!Total, 2) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Vencimiento) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                End If
                DataEnvironment1.Sistema.Execute Consulta
                    
                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                        "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                    DataEnvironment1.AMR.Execute Consulta
'                End If
                rs.MoveNext
            Else 'esto es por si tengo algun saldo para mostrar
                sal = 0
                While Not rsaux.EOF
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
                    If VaEnElDebe(rs!TIPODOC) Then
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & Trim(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Vencimiento) & _
                                ", '" & Trim(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(sal, 2) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    Else
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & Trim(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Vencimiento) & _
                                ", '" & Trim(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(sal, 2) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    End If
                    DataEnvironment1.Sistema.Execute Consulta
                        
                    'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                    If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                    "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                            "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                    ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                        DataEnvironment1.AMR.Execute Consulta
'                    End If
                    rs.MoveNext
                End If

            End If
            Set rsaux = Nothing
        Wend
        rs.Close
        Set rs = Nothing
        
        rsCli.MoveNext
    Wend
    
    CalcularSaldo

    Set rs = Nothing
    Set rsCli = Nothing
End Sub

Private Sub cmdAceptar2_Click()
    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If dtfechad.Value < CDate("01/04/2006") Then
        MsgBox "Debe ingresar una fecha posterior al 01/04/2006."
        dtfechad.Value = "01/04/2006"
        Exit Sub
    End If
    
    If rangoOk Then
                
        relojito
        
        
        
        CrearConsulta False
        LimpiarGrilla Grilla
        grilla2.rows = 1
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaEfectivo
        LimpiarGrilla GrillaMoviCaja
        
'        LlenarGrilla grilla, _
'            "Select CODIGO_CLI AS CODIGO, C.DESCRIPCION AS DESCRIPCION, L.FECHA, L.TIPO_DOCUMENTO, " & _
'            "L.NRO_DOCUMENTO, L.REMITO, L.IVA, L.DEBE, L.HABER, L.SALDO " & _
'            "From " & TablaTemp & " AS L INNER JOIN CLIENTES AS C ON C.CODIGO = L.CODIGO_CLI " & _
'            "Where l.facturas = ' ' and l.cheques = ' ' " & _
'            "Order By CODIGO_CLI, FECHA, ID", True
        LlenarGrilla Grilla, _
            " Select CODIGO_CLI AS CODIGO,  DESCRIPCION_CLI as DESCRIPCION, FECHA, TIPO_DOCUMENTO, " & _
            " NRO_DOCUMENTO, REMITO, DEBE, HABER, SALDO, '' as [Saldo final] " & _
            " From " & TablaTemp & _
            " Where facturas = ' ' and cheques = ' ' " & _
            " Order By CODIGO_CLI, FECHA, ID", True
            
        grillaMarcoSaldosFinales Grilla, 0, 9, 8
        '20080331
        limpioGrilla 9
        
        lleno2
        relojito False
        
'        MsgBox "" & Grilla.rows
        
    End If
End Sub

Private Function lleno2()
    Dim i As Long
    Dim CUIT As String
    Dim razon As String
    Dim COD As Long
    Dim nro1 As Double
    Dim nro2 As Double
    Dim nro3 As Double
    Dim nro4 As Double
    Dim nro5 As Double
    Dim nro6 As Double
    Dim nro7 As Double
    Dim nro8 As Double
    Dim a As Date
    Dim b As Long
    Dim C As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim k As Long
    
    
    If Grilla.rows < 2 Then
        Exit Function
    End If
    grilla2.rows = 1
    
    nro1 = 0: nro2 = 0: nro3 = 0: nro4 = 0: nro5 = 0: nro6 = 0: nro7 = 0: nro8 = 0
    i = 1
    COD = Grilla.TextMatrix(1, 0)
    CUIT = Trim(obtenerDeSQL("select cuit from clientes where codigo=" & Grilla.TextMatrix(1, 0)))
    razon = Trim(Grilla.TextMatrix(1, 1))
    While i < Grilla.rows
        If COD = Grilla.TextMatrix(i, 0) Then
            'If Trim(Grilla.TextMatrix(i, 3)) <> "SI" Then
                b = Year(Grilla.TextMatrix(i, 2)) & Format(Month(Grilla.TextMatrix(i, 2)), "00") & Format(Day(Grilla.TextMatrix(i, 2)), "00")
                C = Year(dtfechad.Value) & Format(Month(dtfechad.Value), "00") & Format(Day(dtfechad.Value), "00")
                d = Year(dtfechad.Value - 90) & Format(Month(dtfechad.Value - 90), "00") & Format(Day(dtfechad.Value - 90), "00")
                e = Year(dtfechad.Value - 180) & Format(Month(dtfechad.Value - 180), "00") & Format(Day(dtfechad.Value - 180), "00")
                f = Year(dtfechad.Value - 365) & Format(Month(dtfechad.Value - 365), "00") & Format(Day(dtfechad.Value - 365), "00")
                g = Year(dtfechad.Value + 90) & Format(Month(dtfechad.Value + 90), "00") & Format(Day(dtfechad.Value + 90), "00")
                h = Year(dtfechad.Value + 180) & Format(Month(dtfechad.Value + 180), "00") & Format(Day(dtfechad.Value + 180), "00")
                k = Year(dtfechad.Value + 365) & Format(Month(dtfechad.Value + 365), "00") & Format(Day(dtfechad.Value + 365), "00")
                If b < C And b >= d Then 'vencido 0 a 3
                    nro1 = nro1 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b < d And b >= e Then 'vencido 3 a 6
                    nro2 = nro2 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b < e And b >= f Then 'vencido 6 a 12
                    nro3 = nro3 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b < f Then 'vencido mas de 12
                    nro4 = nro4 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b >= C And b <= g Then 'por vencer 0 a 3
                    nro5 = nro5 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b > g And b <= h Then 'por vencer 3 a 6
                    nro6 = nro6 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b > h And b <= k Then 'por vencer 6 a 12
                    nro7 = nro7 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                ElseIf b > k Then 'por vencer mas de 12
                    nro8 = nro8 + s2n(Grilla.TextMatrix(i, 6)) - s2n(Grilla.TextMatrix(i, 7))
                End If
            'End If
            i = i + 1
        Else
            grilla2.AddItem CUIT & Chr(9) & razon & Chr(9) & "" & Chr(9) & nro1 & Chr(9) & nro2 & Chr(9) & nro3 & Chr(9) & nro4 & Chr(9) & nro5 & Chr(9) & nro6 & Chr(9) & nro7 & Chr(9) & nro8
            CUIT = Trim(obtenerDeSQL("select cuit from clientes where codigo=" & Grilla.TextMatrix(1, 0)))
            COD = Grilla.TextMatrix(i, 0)
            razon = Trim(Grilla.TextMatrix(i, 1))
            nro1 = 0: nro2 = 0: nro3 = 0: nro4 = 0: nro5 = 0: nro6 = 0: nro7 = 0: nro8 = 0
        End If
        'i = i + 1
    Wend
    If i > 1 Then
        grilla2.AddItem CUIT & Chr(9) & razon & Chr(9) & "" & Chr(9) & nro1 & Chr(9) & nro2 & Chr(9) & nro3 & Chr(9) & nro4 & Chr(9) & nro5 & Chr(9) & nro6 & Chr(9) & nro7 & Chr(9) & nro8
    End If
    
End Function

Private Function limpioGrilla(Col As Long) 'limpio la grilla y tabla temporal de los importes con cero, incluyendo si el total es cero borro historial de cliente
    Dim i As Long
    Dim j As Long
    Dim cli As Long
    Dim Borrar As String
    
    i = 1
    While i < Grilla.rows
        If Grilla.TextMatrix(i, Col) = "0" Then
            cli = Grilla.TextMatrix(i, 0)
            j = 1
            While j < Grilla.rows
                If Grilla.TextMatrix(j, 0) = CStr(cli) Then
                    Grilla.TextMatrix(j, 0) = ""
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
    
    j = 1
    While j < Grilla.rows
        If Grilla.TextMatrix(j, 0) = "" Then
            Borrar = "delete from " & TablaTemp & " where descripcion_cli='" & Grilla.TextMatrix(j, 1) & "'"
            DataEnvironment1.Sistema.Execute Borrar
            Grilla.RemoveItem (j)
        Else
            j = j + 1
        End If
        'j = j + 1
    Wend
End Function


Private Sub cmdAceptar_Click()
Dim rsTodoF As New ADODB.Recordset, i As Long
Dim rsClientes As New ADODB.Recordset, C As Long

If dtfechad.Value < CDate("01/01/2007") Then
    dtfechad.Value = "01/01/2007"
End If

Dim Corte As Date
Dim VMas03 As Double, VMas36 As Double, VMas612 As Double, VMas12 As Double
Dim VMenos03 As Double, VMenos36 As Double, VMenos612 As Double, VMenos12 As Double
Dim VAux As Double, VAux2 As Double, VAux3 As Double, VAnterior As Double
Dim Mas03 As Date, Mas36 As Date, Mas612 As Date
Dim Menos03 As Date, Menos36 As Date, Menos612 As Date
Corte = dtfechad
Mas03 = Corte + 90
Mas36 = Corte + 180
Mas612 = Corte + 360
Menos03 = Corte - 90
Menos36 = Corte - 180
Menos612 = Corte - 360

Dim tt As String
    tt = TablaTempCrear("([ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_CLI] [numeric](18, 0) NULL ," _
        & "[CUIT_CLI] [nvarchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DESCRIPCION_CLI] [nvarchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMenos12] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMenos612] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMenos36] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMenos03] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMas03] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMas36] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMas612] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[VMas12] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & ") ON [PRIMARY]")
        

'<Menos12>= -- <Menos36>= -- <Menos03>= -- <CORTE>= -- <Mas03>= -- <Mas36>= -- <Mas12>=


rsTodoF.Open "SELECT     TipoDoc, NroFactura,Codigo, Total, Cliente, RazonSocial, Vencimiento,Fecha FROM         FacturaVenta WHERE   (CLIENTE>=" & uCliD.codigo & " AND CLIENTE<= " & uCliH.codigo & " ) AND   (Fecha >='01/01/07' and  Fecha<=" & ssFecha(Corte) & ") AND (Activo = 1) ORDER BY fecha", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
rsClientes.Open "select * from clientes where    (CODIGO>=" & uCliD.codigo & " AND CODIGO<= " & uCliH.codigo & " ) AND  activo=1 order by codigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

If rsClientes.EOF And rsClientes.BOF Then Exit Sub
rsClientes.MoveFirst

With rsTodoF
    If .EOF And .BOF Then
    Else
        Grilla.clear
        Grilla.cols = 10
        Grilla.rows = 1
        Grilla.TextMatrix(0, 0) = "TipoDoc"
        Grilla.TextMatrix(0, 1) = "NroFac"
        Grilla.TextMatrix(0, 2) = "Codigo"
        Grilla.TextMatrix(0, 3) = "Total"
        Grilla.TextMatrix(0, 4) = "Suma"
        Grilla.TextMatrix(0, 5) = "Resta"
        Grilla.TextMatrix(0, 6) = "Cli"
        Grilla.TextMatrix(0, 7) = "Razon"
        Grilla.TextMatrix(0, 8) = "Vencimiento"
        Grilla.TextMatrix(0, 9) = "Fecha"
        Grilla.Visible = False
    
        For C = 0 To rsClientes.RecordCount - 1
            .MoveFirst
                If .EOF And .BOF Then GoTo salto
            
                VMas03 = 0
                VMas36 = 0
                VMas612 = 0
                VMas12 = 0
                VMenos03 = 0
                VMenos36 = 0
                VMenos612 = 0
                VMenos12 = 0
                
                'If rsClientes!codigo = 14 Then
                'Stop
                'End If
                
                VAnterior = CalcularSaldoAnterior(rsClientes!codigo, CDate("01/04/07")) ', 0, True)
                
                VMenos12 = s2n(VAnterior)
                If VMenos12 = -0.01 Then VMenos12 = 0
                
                For i = 0 To .RecordCount - 1
                    If !cliente <> rsClientes!codigo Then GoTo OTRO
                    
                    'If Trim(!TIPODOC) = "NCA" Then
                        'Stop
                    'End If
                    
                    'If !cliente = 43 Then '' Or !cliente = 2 Then
                    '    If VaEnElDebe(!TIPODOC) Then
                    '        grilla.AddItem !TIPODOC & Chr(9) & !NroFactura & Chr(9) & !codigo & Chr(9) & !Total & Chr(9) & !Total & Chr(9) & Chr(9) & !cliente & Chr(9) & !RAZONSOCIAL & Chr(9) & !Vencimiento & Chr(9) & !fecha
                    '    Else
                    '        grilla.AddItem !TIPODOC & Chr(9) & !NroFactura & Chr(9) & !codigo & Chr(9) & !Total & Chr(9) & Chr(9) & !Total & Chr(9) & !cliente & Chr(9) & !RAZONSOCIAL & Chr(9) & !Vencimiento & Chr(9) & !fecha
                    '    End If
                    'Else
                        'Stop
                    'End If

                    If CDate(!Vencimiento) < Menos612 Then
                        'VAux2 = s2n(ObtenerPago(!codigo, CDate(!fecha), Corte, !cliente))
                        'If VAux2 = 0 Then
                            VAux2 = s2n(ObtenerPago(!codigo, CDate(!Fecha) - 365, Corte, !cliente))
                        'End If
                        VAux = s2n(!Total) - s2n(VAux2)
                        If VaEnElDebe(!TIPODOC) Then
                            VMenos12 = s2n(VMenos12 + s2n(VAux))
                        Else
                            VMenos12 = s2n(VMenos12 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Menos612 And CDate(!Vencimiento) < Menos36 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMenos612 = s2n(VMenos612 + s2n(VAux))
                        Else
                            VMenos612 = s2n(VMenos612 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Menos36 And CDate(!Vencimiento) < Menos03 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        'If VAux > 0 Then
                        '    Stop
                        'End If
                        If VaEnElDebe(!TIPODOC) Then
                            VMenos36 = s2n(VMenos36 + s2n(VAux))
                        Else
                            VMenos36 = s2n(VMenos36 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Menos03 And CDate(!Vencimiento) < Corte Then
                        'If !codigo = 35411 Then
                        '    Stop
                        'End If
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMenos03 = s2n(VMenos03 + s2n(VAux))
                        Else
                            VMenos03 = s2n(VMenos03 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Corte And CDate(!Vencimiento) < Mas03 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMas03 = s2n(VMas03 + s2n(VAux))
                        Else
                            VMas03 = s2n(VMas03 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Mas03 And CDate(!Vencimiento) < Mas36 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMas36 = s2n(VMas36 + s2n(VAux))
                        Else
                            VMas36 = s2n(VMas36 + s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Mas36 And CDate(!Vencimiento) < Mas612 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMas612 = s2n(VMas612 + s2n(VAux))
                        Else
                            VMas612 = s2n(VMas612 - s2n(VAux))
                        End If
                    ElseIf CDate(!Vencimiento) >= Mas612 Then
                        VAux = s2n(!Total) - s2n(ObtenerPago(!codigo, CDate(!Fecha), Corte, !cliente))
                        If VaEnElDebe(!TIPODOC) Then
                            VMas12 = s2n(VMas12 + s2n(VAux))
                        Else
                            VMas12 = s2n(VMas12 - s2n(VAux))
                        End If
                    End If
OTRO:
                    .MoveNext
                Next
                '!Razonsocial
                If s2n(s2n(VMenos12) + s2n(VMenos612) + s2n(VMenos36) + s2n(VMenos03) + s2n(VMas03) + s2n(VMas36) + s2n(VMas612) + s2n(VMas12)) > 0 Then
                    DataEnvironment1.Sistema.Execute "insert into " & tt & " (CODIGO_CLI,CUIT_CLI,DESCRIPCION_CLI,VMENOS12,VMENOS612,VMENOS36,VMENOS03,VMAS03,VMAS36,VMAS612,VMAS12) Values " _
                    & "(" & rsClientes!codigo & ",'" & rsClientes!CUIT & "','" & rsClientes!DESCRIPCION & "'," & x2s(VMenos12) & "," & x2s(VMenos612) & "," & x2s(VMenos36) & "," & x2s(VMenos03) & "," & x2s(VMas03) & "," & x2s(VMas36) & "," & x2s(VMas612) & "," & x2s(VMas12) & "  )"
                End If
salto:
            rsClientes.MoveNext
        Next
        Set rsClientes = Nothing
        Set rsTodoF = Nothing
        LlenarGrilla grilla2, "SELECT CUIT_CLI AS CUIT,DESCRIPCION_CLI AS RAZONSOCIAL,VMENOS12 AS [MAS DE 12],VMENOS612 AS [6 A 12],VMENOS36 AS [3 A 6],VMENOS03 AS [0 A 3 M],VMAS03 AS [0 A 3 M],VMAS36 AS [3 A 6],VMAS612 AS [6 A 12],VMAS12 as [MAS DE 12],'' as TOTAL FROM " & tt & " WHERE CODIGO_CLI>=" & uCliD.codigo & " AND CODIGO_CLI<=" & uCliH.codigo, False
        
        
        With grilla2
            .ColWidth(0) = 1200
            .ColWidth(1) = 2100
            .ColWidth(2) = 1100
            .ColWidth(3) = 950
            .ColWidth(4) = 950
            .ColWidth(5) = 950
            .ColWidth(6) = 950
            .ColWidth(7) = 950
            .ColWidth(8) = 950
            .ColWidth(9) = 1100
            For C = 1 To .rows - 1
                .TextMatrix(C, 10) = s2n(s2n(.TextMatrix(C, 2)) + s2n(.TextMatrix(C, 3)) + s2n(.TextMatrix(C, 4)) + s2n(.TextMatrix(C, 5)) + s2n(.TextMatrix(C, 6)) + s2n(.TextMatrix(C, 7)) + s2n(.TextMatrix(C, 8)) + s2n(.TextMatrix(C, 9)))
            Next
            
        End With
        ini2
        
    End If
End With

End Sub


Private Function ObtenerPago(NroFactura As String, FechaD As Date, FechaH As Date, Optional cliente As Long = 0) As Double
Dim rsTodoP As New ADODB.Recordset, Valor As Double, i As Long, C As String, Aux, we As String
'Debug.Print NroFactura
'If NroFactura = "34406" Then
'    Stop
'End If


If cliente > 0 Then
    we = " and r.cliente=" & cliente
Else
    we = ""
End If
If FechaD < CDate("01/01/07") Then FechaD = "01/01/07"
C = "SELECT  r.cliente,r.codigo,r.numero,r.fecha,r.activo,d.* FROM  Recibos  as R inner join RecibosDetalle as D ON R.CODIGO=D.CODRECIBO WHERE d.facturaVENTA='" & Trim(NroFactura) & "' and (r.ACTIVO = 1) AND (r.FECHA >= " & ssFecha(FechaD) & " and r.fecha<=" & ssFecha(FechaH) & " ) " & we
rsTodoP.Open C, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsTodoP
    If .EOF And .BOF Then
    Valor = 0
    Else
        .MoveFirst
        Valor = 0
        For i = 0 To .RecordCount - 1
            Aux = !numero
            Valor = Valor + s2n(!Importe)
            .MoveNext
        Next
    End If
    Set rsTodoP = Nothing
    If Valor > 0 Then
        Grilla.AddItem "REC" & Chr(9) & Aux & Chr(9) & NroFactura & Chr(9) & s2n(Valor) & Chr(9) & Chr(9) & s2n(Valor) & Chr(9) & "1" & Chr(9) & "C" & Chr(9) & FechaH & Chr(9) & FechaH
    End If
End With
ObtenerPago = Valor
End Function



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
    '        x As VSPrinter8LibCtl.OrientationSettings
    If grilla2.rows < 2 Then Exit Sub
    
    grilla2.GridLines = flexGridNone
    grilla2.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla2.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 12
'    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    
    FrmImpresiones.VSPrinter.StartDoc
    'FrmImpresiones.VSPrinter.Paragraph = "Listado Mayor al " & Format$(Date, "dd / mm / yyyy")
    FrmImpresiones.VSPrinter.Paragraph = "Listado de composicion"
    FrmImpresiones.VSPrinter.Paragraph = "" '"Entre fechas : " & uFechaD.dtFecha & " - " & uFechaH.dtFecha  '& "     Rango de Cuentas : " & CmbCtaD & "  -  " & CmbCtaH
    FrmImpresiones.VSPrinter.Paragraph = " "
    
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    FrmImpresiones.VSPrinter.RenderControl = grilla2.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla2.GridLines = flexGridFlat
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
    tempo = obtenerDeSQL("select min(codigo) as mini, max(codigo) as maxi from clientes")
    'txtcodclih = tempo(1) ' "9999999"
    uCliH.codigo = tempo(1)
'    txtclienteh = ""
    'txtcodclid = tempo(0)
    uCliD.codigo = tempo(0)
'    txtcliented = "" 'tempo(1)
    dtfechad.Value = "01/04/2007" 'Date
    dtfechah.Value = "01/01/2500" 'Date
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
        & "[DESCRIPCION_CLI] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
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
    
    ucXls1.ini grilla2, "C:\ComposicionCuentaCli_PorPeriodo", "Composicion cuenta cliente por periodo"
    
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
    ini
    dtfechad = Date
End Sub

Private Sub Form_Resize()
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
    Anclar fraGri, Me, anclarLadosTodos
    Anclar Grilla, fraGri, anclarLadosTodos
    Anclar fraSubGri, fraGri, anclarIzquierda + anclarAbajo
End Sub

Private Sub grilla_Click()
Dim TIPODOC As String
Dim NroDoc As Long
Dim CodInt As Long

    With Grilla
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
'    Dim consulta As String

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
'    If rangoOk Then

'        If MsgBox("¿Desea los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            
'            consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
'                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
'                    " DEBE, HABER, SALDO, '' as [Saldo Final] " & _
'                    " From  " & TablaTemp & _
'                    " Order By CODIGO_CLI, FECHA, ID"
'            CrearConsulta True
'            LlenarGrilla grilla, consulta, False
'            grillaMarcoSaldosFinales grilla, 0, 10, 9
'            limpioGrilla 10
'
'        Else
'            consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
'                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO, '' as  Final " & _
'                    " From " & TablaTemp & _
'                    " Where TIPO_DOCUMENTO <> '' " & _
'                    " Order By CODIGO_CLI, FECHA, ID"
'            CrearConsulta True
'            LlenarGrilla grilla, consulta, False
'            grillaMarcoSaldosFinales grilla, 0, 8, 7
'            limpioGrilla 8
'        End If
        
'    Else
'        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
'        Cancel = True
'    End If
     
End Sub
Private Function rangoOk() As Boolean
    rangoOk = (uCliD.codigo > 0 And uCliH.codigo > 0)
End Function

Private Function ini2()
    With grilla2
        .AddItem "", 0
        .TextMatrix(0, 2) = "VENCIDO"
        .TextMatrix(0, 6) = "A VENCER"
        '.TextMatrix(1, 0) = "CUIT"
        '.TextMatrix(1, 1) = "Razon Social"
        '.TextMatrix(1, 2) = "Doc"
        '.TextMatrix(1, 3) = "0 a 3 meses"
        '.TextMatrix(1, 4) = "3 a 6"
        '.TextMatrix(1, 5) = "6 a 12"
        '.TextMatrix(1, 6) = "mas de 12"
        '.TextMatrix(1, 7) = "0 a 3 meses"
        '.TextMatrix(1, 8) = "3 a 6"
        '.TextMatrix(1, 9) = "6 a 12"
        '.TextMatrix(1, 10) = "mas de 12"
        .cell(flexcpFontBold, 1, 0, 1, 10) = True
    End With
End Function

Private Function ini()
    With grilla2
        .rows = 2
        .TextMatrix(0, 2) = "VENCIDO"
        .TextMatrix(0, 6) = "A VENCER"
        .TextMatrix(1, 0) = "CUIT"
        .TextMatrix(1, 1) = "Razon Social"
        '.TextMatrix(1, 2) = "Doc"
        .TextMatrix(1, 2) = "0 a 3 meses"
        .TextMatrix(1, 3) = "3 a 6"
        .TextMatrix(1, 4) = "6 a 12"
        .TextMatrix(1, 5) = "mas de 12"
        .TextMatrix(1, 6) = "0 a 3 meses"
        .TextMatrix(1, 7) = "3 a 6"
        .TextMatrix(1, 8) = "6 a 12"
        .TextMatrix(1, 9) = "mas de 12"
        .TextMatrix(1, 10) = "Total"
        .cell(flexcpFontBold, 1, 0, 1, 10) = True
    End With
End Function



