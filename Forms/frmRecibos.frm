VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmRecibos 
   Caption         =   "Recibos e Imputaciones"
   ClientHeight    =   9225
   ClientLeft      =   225
   ClientTop       =   480
   ClientWidth     =   11565
   Icon            =   "frmRecibos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   11565
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabRecibo 
      Height          =   6690
      Left            =   15
      TabIndex        =   15
      Top             =   915
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11800
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Documentos"
      TabPicture(0)   =   "frmRecibos.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Forma de Pago"
      TabPicture(1)   =   "frmRecibos.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBorraItemRet"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraMoney"
      Tab(1).Control(2)=   "gRetenciones"
      Tab(1).Control(3)=   "Label3(1)"
      Tab(1).Control(4)=   "lblSumaRet"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdBorraItemRet 
         Height          =   435
         Left            =   -74040
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmRecibos.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Borrar Item"
         Top             =   5010
         Width           =   435
      End
      Begin VB.Frame fraMoney 
         BorderStyle     =   0  'None
         Height          =   3825
         Left            =   -74895
         TabIndex        =   31
         Top             =   360
         Width           =   11235
         Begin Gestion.uCtaBanco uCtaBanco 
            Height          =   375
            Left            =   1395
            TabIndex        =   46
            Top             =   3405
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   661
         End
         Begin VB.TextBox txtTransferencia 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   15
            TabIndex        =   45
            Top             =   3405
            Width           =   1275
         End
         Begin VB.CommandButton cmdBorraItem 
            Height          =   435
            Left            =   840
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmRecibos.frx":0C0C
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Borrar Item"
            Top             =   1620
            Width           =   435
         End
         Begin VB.TextBox txtEfectivo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            TabIndex        =   35
            Top             =   360
            Width           =   1275
         End
         Begin VB.TextBox txtTotalRecibo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7485
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   45
            Width           =   1275
         End
         Begin VB.TextBox txtCaja 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1590
            TabIndex        =   33
            Text            =   "1"
            Top             =   375
            Width           =   675
         End
         Begin VB.TextBox txtCuentaEfectivo 
            Height          =   315
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   375
            Width           =   1650
         End
         Begin VSFlex7LCtl.VSFlexGrid gCheques 
            Height          =   2385
            Left            =   1380
            TabIndex        =   37
            Top             =   885
            Width           =   8715
            _cx             =   15372
            _cy             =   4207
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
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia:"
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
            Index           =   3
            Left            =   0
            TabIndex        =   48
            Top             =   3105
            Width           =   1275
         End
         Begin VB.Label lblSumaCheques 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   60
            TabIndex        =   44
            Top             =   1200
            Width           =   1275
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Falta :"
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
            Index           =   1
            Left            =   6855
            TabIndex        =   43
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblFaltaPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   7485
            TabIndex        =   42
            Top             =   375
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "en Cheques :"
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
            Index           =   0
            Left            =   60
            TabIndex        =   41
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "en Efectivo :"
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
            Left            =   120
            TabIndex        =   40
            Top             =   90
            Width           =   1155
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Index           =   0
            Left            =   6840
            TabIndex        =   39
            Top             =   45
            Width           =   855
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caja :"
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
            Left            =   1680
            TabIndex        =   38
            Top             =   90
            Width           =   735
         End
      End
      Begin VB.Frame fra4 
         BorderStyle     =   0  'None
         Height          =   5355
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   11295
         Begin VB.TextBox txtFyD 
            Height          =   285
            Left            =   10395
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1125
            Width           =   855
         End
         Begin VB.TextBox txtTotalDoc 
            Height          =   285
            Left            =   10395
            Locked          =   -1  'True
            TabIndex        =   22
            Tag             =   "8"
            Top             =   2595
            Width           =   855
         End
         Begin VB.TextBox txtCyR 
            Height          =   285
            Left            =   10380
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1860
            Width           =   855
         End
         Begin VSFlex7LCtl.VSFlexGrid gReci 
            Height          =   2400
            Left            =   5280
            TabIndex        =   17
            Top             =   2955
            Width           =   5040
            _cx             =   8890
            _cy             =   4233
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
         Begin VSFlex7LCtl.VSFlexGrid gCred 
            Height          =   2400
            Left            =   0
            TabIndex        =   18
            Top             =   2940
            Width           =   5145
            _cx             =   9075
            _cy             =   4233
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
         Begin VSFlex7LCtl.VSFlexGrid gDebi 
            Height          =   2400
            Left            =   5280
            TabIndex        =   19
            Top             =   255
            Width           =   5040
            _cx             =   8890
            _cy             =   4233
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
         Begin VSFlex7LCtl.VSFlexGrid gFact 
            Height          =   2400
            Left            =   0
            TabIndex        =   20
            Top             =   240
            Width           =   5145
            _cx             =   9075
            _cy             =   4233
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
         Begin VB.Line Line4 
            BorderColor     =   &H00400000&
            Visible         =   0   'False
            X1              =   10320
            X2              =   11220
            Y1              =   2220
            Y2              =   2220
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   10395
            TabIndex        =   30
            Top             =   2325
            Width           =   675
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "FAC - ND"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   10380
            TabIndex        =   29
            Top             =   915
            Width           =   855
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "NC - REC"
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   10365
            TabIndex        =   28
            Top             =   1605
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Facturas"
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
            Left            =   60
            TabIndex        =   27
            Top             =   0
            Width           =   915
         End
         Begin VB.Label lbl7 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "N Debito"
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
            Index           =   1
            Left            =   5685
            TabIndex        =   26
            Top             =   15
            Width           =   915
         End
         Begin VB.Label lbl7 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "N Credito"
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
            Index           =   2
            Left            =   60
            TabIndex        =   25
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label lbl7 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Recibos a Cuenta:"
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
            Index           =   3
            Left            =   5685
            TabIndex        =   24
            Top             =   2655
            Width           =   1755
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid gRetenciones 
         Height          =   2250
         Left            =   -73470
         TabIndex        =   51
         Top             =   4245
         Width           =   8715
         _cx             =   15372
         _cy             =   3969
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Retenciones :"
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
         Index           =   1
         Left            =   -74805
         TabIndex        =   52
         Top             =   4290
         Width           =   1275
      End
      Begin VB.Label lblSumaRet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -74835
         TabIndex        =   49
         Top             =   4590
         Width           =   1275
      End
   End
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton cmdCliente 
         Caption         =   "?"
         Height          =   315
         Left            =   2820
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1680
         TabIndex        =   3
         Top             =   465
         Width           =   1095
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtNumero 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   6060
         TabIndex        =   2
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   77463553
         CurrentDate     =   38252
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha :"
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
         Left            =   5280
         TabIndex        =   11
         Top             =   60
         Width           =   735
      End
      Begin VB.Label lblNroRecibo 
         Caption         =   "Nro :"
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
         Left            =   180
         TabIndex        =   10
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lblCodigo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Left            =   9900
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lbl7 
         Caption         =   "Codigo:"
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
         Index           =   0
         Left            =   8220
         TabIndex        =   8
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblTipo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RAA"
         Height          =   315
         Left            =   9120
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Height          =   1515
      Left            =   30
      TabIndex        =   6
      Top             =   7650
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   2672
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucEntreFechas ucFechas 
         Height          =   315
         Left            =   4020
         TabIndex        =   13
         Top             =   0
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buscar Entre"
         Height          =   195
         Left            =   2820
         TabIndex        =   14
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Retenciones :"
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
      Index           =   2
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   1275
   End
End
Attribute VB_Name = "frmRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private stblRecibosTmp As String
Private stblChequestmp As String
Private stblRetencionestmp As String
Private Const tt_Chequetmp = "( [nroint] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [banco] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[cheque] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [importe] [float] NULL ,  [fecha] [datetime] NULL, [propio] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL )"
Private Const tt_Retenciontmp = "( [nroRet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [fecha] [datetime] NULL,[codRet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoRet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [importe] [float] NULL)"
Private Const tt_ReciboCobroTemp = " ([TIPODOC] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NRODOC] [numeric](18, 0) NULL , [FECHA] [datetime] NULL , [SALDO] [float] NULL , [COBRADO] [float] NULL , [ULTIMOSALDO] [float] NULL) "


Private midDoc As Long

Public Enum TipoRecibo_REC_IMP
    trRECIBO
    trIMPUTACION
End Enum

Private Const TipoReciboREC = "REC"
Private Const TipoReciboIMC = "IMC"

Private mTipoRecibo As TipoRecibo_REC_IMP

Private WithEvents cliente As LiCodigo
Attribute cliente.VB_VarHelpID = -1
Private caja As LiCodigo


Private WithEvents g    As LiGrilla '   cheques
Attribute g.VB_VarHelpID = -1
Private WithEvents g4F  As LiGrilla '   facturas
Attribute g4F.VB_VarHelpID = -1
Private WithEvents g4C  As LiGrilla '   N Credito
Attribute g4C.VB_VarHelpID = -1
Private WithEvents g4D  As LiGrilla '   N Debito
Attribute g4D.VB_VarHelpID = -1
Private WithEvents g4R  As LiGrilla '   Recibos a Cuenta
Attribute g4R.VB_VarHelpID = -1
Private WithEvents gRet As LiGrilla '   Retenciones
Attribute gRet.VB_VarHelpID = -1

'cuadro
Private g4TIPO  As Long
Private g4NUME  As Long
Private g4MONT  As Long
Private g4SALD  As Long
Private g4PAGA  As Long
Private g4CODI  As Long ' hidd

'cheques
Private gBANCC  As Long
Private gBANCD  As Long
Private gNROCH  As Long
Private gMONTO  As Long ' monto imputable al recibo,  <= al total.
Private gTOTAL  As Long ' nueva columna
Private gFECHA  As Long
Private gPT     As Long
Private gCODCH  As Long

'retenciones
Private gR_NRO As Long  'nro retencion
Private gR_FEC As Long  'fech
Private gR_TIP As Long  'tipo: descripcion RIB RGA
Private gR_IMP As Long  'importe
Private gR_CUE As Long  'cuenta
Private gR_TIC As Long  'codigo AHORA LO VE EL USUARIO
Private gR_IdC As Long  'idCuentasParam, el unico que se graba
Private gR_FAC As Long  'numero factura ?
'


Public Sub mostrar(que As TipoRecibo_REC_IMP)
    mTipoRecibo = que
    limpiar
    Me.Show
    fraMoney.Visible = (que = trRECIBO)
End Sub

Private Sub cliente_cambio(codigo) ' As Integer)
    'pregunto estado p q no joda cdo cargo uno viejo
    If ucMenu.estado = ucbEditando Then carga4 False
End Sub

'-------------------------------------------------------------
'
Private Sub cmdBorraItem_Click()
    If g.Row > 0 Then g.delRow (g.Row)
    If g.rows = 1 Then g.rows = 2
    recalcFalta
End Sub

Private Sub cmdBorraItemRet_Click()
    If gRet.Row > 1 Then gRet.delRow (gRet.Row)
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_Load()
'    CentrarMe Me
    Me.KeyPreview = True
    
    inigrilla
    iniCliente
    inimenu
    tabRecibo.Tab = 0

    limpiar
    txtcaja = 1
    verCajaEfectivo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub carga4(buscar As Boolean)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim gt As LiGrilla, i As Long, s As String
    Dim fnrodoc As String, ftotal As Double
    Limpiar4
    If buscar Then
        's = "SELECT FacturaVenta.codigo, FacturaVenta.TipoDoc, FacturaVenta.NroFactura, FacturaVenta.total, '?' as saldo , RecibosDetalle.Importe as paga " _
                & " FROM FacturaVenta RIGHT outer JOIN RecibosDetalle ON FacturaVenta.Codigo = RecibosDetalle.FacturaVenta " _
                & " where  RecibosDetalle.CodRecibo = " & lblCodigo _
                & " order by FacturaVenta.NroFactura  "
        s = "SELECT RecibosDetalle.Facturaventa as codigo, FacturaVenta.TipoDoc, FacturaVenta.NroFactura, FacturaVenta.total, RecibosDetalle.Saldo , RecibosDetalle.Importe as paga " _
                & " FROM FacturaVenta RIGHT outer JOIN RecibosDetalle ON FacturaVenta.Codigo = RecibosDetalle.FacturaVenta " _
                & " where  RecibosDetalle.CodRecibo = " & lblCodigo _
                & " order by FacturaVenta.NroFactura  "
    Else
        If cliente.DESCRIPCION = "" Then Exit Sub
        s = "select codigo, tipodoc, nrofactura, total, saldo, '' as paga " _
                & " from FacturaVenta where activo = 1 and saldo >0 and cliente = " & cliente.codigo _
                & " order by FacturaVenta.NroFactura  "
    End If
    
    With rs
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Select Case Trim(!TIPODOC)
            Case "FAA", "FAB", "FAE", "FEA", "FEB", "FEC":                                       Set gt = g4F
            Case "NCA", "NCB", "NCE", "CEA", "CEB", "CEC", "NCT", "NCN", "ACC":                  Set gt = g4C
            Case "NDA", "NDB", "NDE", "DEA", "DEB", "DEC", "ACD":                                Set gt = g4D
            Case "RAA", "FAC":                                              Set gt = g4R
            Case IsNull(Trim(!TIPODOC)), IsEmpty(Trim(!TIPODOC)), "Nulo":   Set gt = g4R
            Case Else
                If IsNull(Trim(!TIPODOC)) Or IsEmpty(Trim(!TIPODOC)) Or Trim(!TIPODOC) = "Nulo" Then
                    Set gt = g4R
                Else
                    Set gt = Nothing
                End If
            End Select
            
            If Not gt Is Nothing Then
                i = gt.addRow
                gt.tx i, g4CODI, !codigo
                gt.tx i, g4TIPO, IIf(IsNull(!TIPODOC), "FAC", !TIPODOC)
                If IsNull(Trim(!NroFactura)) Or IsEmpty(Trim(!NroFactura)) Or Trim(!NroFactura) = "Nulo" Then
                    fnrodoc = obtenerDeSQL("select c.nrodoc from compras c where c.activo=1 and c.id=" & !codigo & " Union select c.nrodoc from transcom c where c.activo=1 and c.id=" & !codigo)
                Else
                    fnrodoc = !NroFactura
                End If
                gt.tx i, g4NUME, fnrodoc
                
                If IsNull(Trim(!Total)) Or IsEmpty(Trim(!Total)) Or Trim(!Total) = "Nulo" Then
                    ftotal = obtenerDeSQL("select c.total from compras c where c.activo=1 and c.id=" & !codigo & " Union select c.total from transcom c where c.activo=1 and c.id=" & !codigo)
                Else
                    ftotal = !Total
                End If
                gt.tx i, g4MONT, ftotal
                gt.tx i, g4SALD, !saldo
                gt.tx i, g4PAGA, !paga
            End If
            .MoveNext
        Wend
    End With
    Set rs = Nothing
    Set gt = g4R
    Dim z As Long, xCuit As String
    xCuit = obtenerDeSQL("select cuit from clientes where codigo=" & s2n(txtCodCliente))
    
    If buscar Then
        rs.Open "SELECT * FROM COMPRAS WHERE CUITPROV='" & Trim(xCuit) & "' AND [PLAN]=" & lblCodigo, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    Else
        rs.Open "SELECT * FROM TRANSCOM WHERE CUITPROV='" & Trim(xCuit) & "' AND FORMADEPAGO=-1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    End If
    If rs.EOF And rs.BOF Then
    Else
        rs.MoveFirst
        For z = 0 To rs.RecordCount - 1
            i = gt.addRow
            gt.tx i, g4CODI, rs!ID
            gt.tx i, g4TIPO, rs!TIPODOC
            gt.tx i, g4NUME, rs!NroDoc
            gt.tx i, g4MONT, rs!Total
            If buscar Then
                gt.tx i, g4SALD, "?"
                gt.tx i, g4PAGA, rs!Total
            Else
                gt.tx i, g4SALD, rs!saldo
                gt.tx i, g4PAGA, ""
            End If
            rs.MoveNext
        Next
    End If
    
    

    Set rs = Nothing
Exit Sub
ufaErr:
    ufa "err llenando grillas", Me.Name ', Err
    Set rs = Nothing
End Sub


' ini ----------------------------
Private Sub iniCliente()
    Set cliente = New LiCodigo
    cliente.init cmbCliente, txtCodCliente, "clientes", False, False, cmdCliente, "activo = 1", True
    cliente.EditaDescripcion = False
    
End Sub
Private Sub inimenu()
    
        
    ucMenu.init True, True, False, True, True, "select * from recibos where activo = 1 and tipodoc = '" & TipoRecibo() & "' order by codigo ", DataEnvironment1.Sistema
    ucMenu.MsgConfirmaSalir = "Cerrar Form ?"
    ucMenu.MsgConfirmaEliminar = "Elimina Recibo ?"
End Sub
Private Sub inigrilla()
    Dim sr As String, nUso As ID_UsoCuenta
    nUso = ID_UsoCuenta_RETVTA
    
    Set gRet = New LiGrilla
    gRet.init gRetenciones

    Set g = New LiGrilla
    With g
        .init gCheques
        gBANCC = .AddCol(" Banco ", "N", 0)
        gBANCD = .AddCol("  Banco                        ")
        gNROCH = .AddCol("  Nro Cheque      ", "S")
        gTOTAL = .AddCol("  Total      ", "N")
        gMONTO = .AddCol(" ..a imputar ", "H")
        gFECHA = .AddCol("  Fecha     ", "D")
        gPT = .AddCol(" P/T ", "S")
        gCODCH = .AddCol("Cod Interno")
    End With
    
    sr = strComboGrilla("select descripcion from cuentasparam where usocuenta = '" & nUso & "' and activo = 1 ")
    With gRet
        gR_NRO = .AddCol(" Nro Ret     ", "N", 0)
        gR_FEC = .AddCol(" Fecha Ret   ", "D")
        gR_TIC = .AddCol(" Cod Ret ", "S") ', "H")
        gR_TIP = .AddCol(" Tipo Ret                     ", "B", sr)
        gR_IMP = .AddCol(" Importe Ret  ", "N", 2)
        gR_CUE = .AddCol(" Cuenta        ")
        gR_IdC = .AddCol(" idCuentaParam", "H")
        gR_FAC = .AddCol(" Nro Factura ", "H") ' --------------
    End With
    
    Set g4F = New LiGrilla
    Set g4R = New LiGrilla
    Set g4C = New LiGrilla
    Set g4D = New LiGrilla
    
    iniG g4F, gFact
    iniG g4R, gReci
    iniG g4C, gCred
    iniG g4D, gDebi
End Sub
Private Sub iniG(gg As LiGrilla, gO As VSFlexGrid)
    With gg
        .init gO
        g4TIPO = .AddCol(" Tipo ")
        g4NUME = .AddCol(" Numero ")
        g4MONT = .AddCol(" Total         ", "9")
        g4SALD = .AddCol(" Saldo         ", "9")
        g4PAGA = .AddCol(" A imputar     ", "N")
        g4CODI = .AddCol("codigo", "H")
    End With
End Sub


' grilla Cheques -----------------------------------
Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim deba As String
    Select Case Col
    Case gBANCC ': g.tx Row, gBANCD, ObtenerDescripcion("BancosGrales", s2n(txt))
        deba = verDescBanco(s2n(txt))
        g.tx Row, gBANCD, deba
        If deba = "" Then g.tx Row, Col, ""
    Case gMONTO ', gTOTAL
'        recalcFalta
        If s2n(g.tx(Row, gMONTO)) > s2n(g.tx(Row, gTOTAL)) Then g.tx Row, gTOTAL, txt
        recalcFalta
    Case gTOTAL
'        If s2n(g.tx(row, gMONTO)) = 0 Then
        g.tx Row, gMONTO, txt
        If s2n(g.tx(Row, gMONTO)) > s2n(g.tx(Row, gTOTAL)) Then g.tx Row, gMONTO, txt
        recalcFalta
    End Select
    
    If g.rows < Row + 2 Then
        g.rows = g.rows + 1
        'g.tx g.rows - 1, gPRoTE, "P"
    End If
    
End Sub


Private Sub g_DblClick()
    Dim resu
    If g.Col = gBANCC Then
        resu = frmBuscar.MostrarCodigoDescripcionActivo("BancosGrales")
        If resu > "" Then g.tx g.Row, g.Col, resu
    End If
End Sub
Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    Select Case Col
    Case gPT: cancel = (g.EditText <> "P" And g.EditText <> "T")
    '
    End Select
End Sub
'----------------------------------------------------

Private Sub recalcFalta()
    lblSumaCheques = n2r(g.suma(gMONTO))
    lblSumaRet = n2r(gRet.suma(gR_IMP))
    lblFaltaPagar = n2r(s2n(txtTotalRecibo) _
        - s2n(txtefectivo) - s2n(lblSumaCheques) - s2n(lblSumaRet) - s2n(txtTransferencia))
End Sub


'ValidarImporte
Private Sub g4C_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = malPaga(Row, Col, g4C)
End Sub
Private Sub g4D_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = malPaga(Row, Col, g4D)
End Sub
Private Sub g4F_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = malPaga(Row, Col, g4F)
End Sub
Private Sub g4R_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = malPaga(Row, Col, g4R)
End Sub
Private Function malPaga(Row As Long, Col As Long, gg As LiGrilla) As Boolean ' como las minas
    Dim tmp
    If Col = g4PAGA Then
        tmp = s2n(gg.EditText)
        malPaga = tmp > s2n(gg.tx(Row, g4SALD)) Or tmp < 0
    End If
End Function
'-------------------

'-------------------
'relleno auto
Private Sub g4C_DblClick()
    IniSaldoG4 g4C ', Row, Col
End Sub
Private Sub g4D_DblClick()
    IniSaldoG4 g4D ', Row, Col
End Sub
Private Sub g4F_DblClick()
    IniSaldoG4 g4F ' Row, Col
End Sub
Private Sub g4R_DblClick()
    IniSaldoG4 g4R ', Row, Col
End Sub
Private Sub IniSaldoG4(gg As LiGrilla) ', Row As Long, Col As Long)
    Dim Col As Long, Row As Long
    Col = gg.Col
    Row = gg.Row
    If Col = g4PAGA And Trim(gg.tx(Row, g4PAGA)) = "" Then gg.tx Row, Col, gg.tx(Row, g4SALD)
End Sub
'-------------------'


'-------------------
'actualizar txtbox
Private Sub g4D_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    ReverTotal
End Sub
Private Sub g4F_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    ReverTotal
End Sub
Private Sub g4C_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    ReverTotal
End Sub
Private Sub g4R_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    ReverTotal
End Sub
Private Sub ReverTotal()
    txtFyD = s2n(g4F.suma(g4PAGA) + g4D.suma(g4PAGA))
    txtCyR = s2n(g4C.suma(g4PAGA) + g4R.suma(g4PAGA))
    txtTotalDoc = s2n(txtFyD - txtCyR)
    txtTotalRecibo = s2n(txtFyD - txtCyR)
'    txtefectivo = s2n(txtFyD - txtCyR)
End Sub
'-------------------


Private Function RevisarReten(cual, Row As Long) As Boolean
    Dim tempo
    With gRet
        tempo = obtenerDeSQL("select codigo, cuenta, id from cuentasparam where descripcion = '" & cual & "' and activo = 1 and usoCuenta = '" & ID_UsoCuenta_RETVTA & "'")
        If IsEmpty(tempo) Then
            .tx Row, gR_CUE, ""
            .tx Row, gR_TIC, 0
            .tx Row, gR_IdC, 0
            RevisarReten = False
        Else
            .tx Row, gR_CUE, tempo(1)
            .tx Row, gR_TIC, tempo(0)
            .tx Row, gR_IdC, tempo(2)
            RevisarReten = True
        End If
    End With
End Function
Private Function RevisarRetenTIC(cual, Row As Long) As Boolean
    Dim tempo
    With gRet
        tempo = obtenerDeSQL("select codigo, cuenta, id, descripcion  from cuentasparam where codigo = '" & cual & "' and activo = 1 and usoCuenta = '" & ID_UsoCuenta_RETVTA & "'")
        If IsEmpty(tempo) Then
            .tx Row, gR_CUE, ""
            '.tx row, gR_TIC, 0
            .tx Row, gR_TIP, ""
            .tx Row, gR_IdC, 0
            RevisarRetenTIC = False
        Else
            .tx Row, gR_CUE, tempo(1)
            '.tx row, gR_TIC, tempo(0)
            .tx Row, gR_IdC, tempo(2)
            .tx Row, gR_TIP, tempo(3)
            RevisarRetenTIC = True
        End If
    End With
End Function

Private Sub gRet_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    With gRet
        If Row = .rows - 1 Then .rows = .rows + 1
    
        If Col = gR_NRO Then
            If s2n(.tx(Row, Col)) = 0 Then
                'borro la parte de retencion
                .tx Row, gR_NRO, ""
                .tx Row, gR_CUE, ""
                .tx Row, gR_FEC, dtFecha
                .tx Row, gR_IMP, ""
                .tx Row, gR_TIC, ""
                .tx Row, gR_TIP, ""
                .tx Row, gR_IdC, ""
            End If
        End If
        If Col = gR_IMP Then
            'lblSumaRet = s2n(gRet.suma(gR_IMP))
            recalcFalta
        End If
    End With
End Sub

Private Sub gRet_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    Select Case Col
'     Case gR_NRO
     Case gR_TIP
        If Not RevisarReten(gRet.EditText, Row) Then cancel = True
'     Case gR_IMP
'        If s2n(gRet.EditText) > s2n(gRet.tx(row, gF_SAL)) Then cancel = True
    Case gR_TIC
        If Not RevisarRetenTIC(gRet.EditText, Row) Then cancel = True
    End Select
    If Trim(gRet.tx(Row, gR_FEC)) = "" Then gRet.tx Row, gR_FEC, dtFecha
End Sub

Private Sub tabRecibo_Click(PreviousTab As Integer)
    recalcFalta
End Sub

Private Sub txtcaja_GotFocus()
    If Trim$(txtcaja) = "" Then txtcaja = "1"
    PintoFocoActivo
End Sub

Private Sub txtCaja_Validate(cancel As Boolean)
    cancel = Not verCajaEfectivo()
End Sub

Private Sub txtEfectivo_GotFocus()
    If s2n(txtefectivo) = 0 Then txtefectivo = lblFaltaPagar 's2n(txtFyD) - s2n(txtCyR)
    PintoFocoActivo
End Sub

Private Sub txtEfectivo_LostFocus()
    recalcFalta
End Sub

'------------------------------
Private Sub txtEfectivo_Validate(cancel As Boolean)
    If Not IsNumeric(txtefectivo) Then cancel = True
    txtefectivo = s2n(txtefectivo)
End Sub


Private Sub txtNumero_LostFocus()
    If Trim$(txtNumero) > "" Then
        If mTipoRecibo = trRECIBO Then
            If YaEstaRecibo(txtNumero) Then
                che "Recibo ya cargado"
                'txtNumero.SetFocus
                'Exit Function
            End If
        Else
            If Not IsEmpty(obtenerDeSQL("select codigo from recibos where tipodoc = 'IMC' and activo = 1 and numero = " & s2n(txtNumero))) Then
                che " Numero imputacion ya cargado"
    '            Exit Function
            End If
        End If
    End If
End Sub

Private Sub txtTransferencia_LostFocus()
    recalcFalta
End Sub

'No edito, sale de g4
'Private Sub txtTotalRecibo_Validate(Cancel As Boolean) ' aca no lo edito
'    If Not IsNumeric(txtTotalRecibo) Then Cancel = True
'    txtTotal = s2n(txtTotalRecibo)
'End Sub
'Private Sub txtTotalReci_GotFocus()
'    txtTotalReci = s2n(txtTotalDoc)
'End Sub

' menu ------------------------------
Private Sub ucMenu_Aceptar()
    If Falta() Then Exit Sub
    'If yasta Then Exit Sub
    If GrabaRecibo Then
        MsgBox "Operacion concluida"
        ucMenu.AceptarOk
    End If
End Sub

Private Sub ucMenu_BorrarControles()
    limpiar
    Limpiar4
End Sub

Private Sub ucMenu_Buscar()
    If BuscaRecibo() Then ucMenu.BuscarOK
End Sub
Private Sub ucMenu_eliminar()
    'esto es para dejar eliminar o no!!anda!!
    If Not permiteEliminar Then 'son los permisos ESPECIALES de los usuarios
        MsgBox "No tiene permiso para poder Eliminar."
        Exit Sub
    End If
    
    If EliminaRecibo() Then ucMenu.EliminarOK
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    fraControl.enabled = sino
    fraMoney.enabled = sino 'And (mTipoRecibo = trRECIBO) ' o sea si imputacion no mueve guita, solo papeles
    'fraMoney.Visible = SiNo And (mTipoRecibo = trRECIBO) ' o sea si imputacion no mueve guita, solo papeles
End Sub

Private Sub ucMenu_Imprimir()
ImprimirReciboConImputacion
End Sub

Private Sub ucMenu_Nuevo()
    On Error Resume Next
    limpiar
    If mTipoRecibo = trRECIBO Then
        txtNumero.SetFocus
    Else
        dtFecha.SetFocus
    End If
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
'------------------------------------

Private Sub limpiar()
    FrmBorrarTxt Me
    lblCodigo = ""
    g.Borrar
    g.rows = 2
    dtFecha = Date
    ucFechas.ini , , ucefHorizontal, ucefFormatoSqlServer
    cliente.codigo = 0
    txtcaja = 1
    verCajaEfectivo
    Limpiar4
    gRet.Borrar
    gRet.rows = 20
    
    tabRecibo.Tab = 0
    midDoc = 0
    
    uCtaBanco.codigo = 0
'reset
    lblCodigo = nuevoCodigo("recibos")
    If mTipoRecibo = trRECIBO Then
        Me.caption = "ingreso de RECIBOS EMITIDOS"
        lblNroRecibo.caption = "Nro Recibo:"
        lblTipo = TipoReciboREC '"REC"
        txtNumero = ""
        txtNumero.enabled = True
        fraMoney.Visible = True
    ElseIf mTipoRecibo = trIMPUTACION Then
        Me.caption = "ingreso de IMPUTACIONES - cancelacion mutua de documentos venta"
        lblTipo = TipoReciboIMC ' "IMC"
        lblNroRecibo.caption = "Nro Imputacion:"
        txtNumero = nuevoCodigo("recibos", "numero", "TipoDoc = 'IMC'")
        txtNumero.enabled = False
        fraMoney.Visible = False
    End If
End Sub

Private Sub Limpiar4()
    g.Borrar
    g.rows = 2
    g4F.Borrar
    g4R.Borrar
    g4C.Borrar
    g4D.Borrar
End Sub

Private Function Falta() As Boolean
    Dim i As Long
    Dim dife As Double
    
    Falta = True
    
    'cabecera
    If s2n(txtNumero) = 0 Then 'And s2n(txtTotalRecibo) <> 0 Then ''  ahora siempre debe tener numero, REC e IMC
        che "Falta ingresar numero de recibo"
        txtNumero.SetFocus
        Exit Function
    End If
    If cliente.DESCRIPCION = "" Then ' Or s2n(txtTotalRecibo) = 0 Then
        che "faltan cliente"
        Exit Function
    End If
    If s2n(txtTotalRecibo) = 0 And mTipoRecibo = trRECIBO Then
        che "Recibo sin Monto"
        Exit Function
    End If
   


    'Nro Repetido
    If mTipoRecibo = trRECIBO Then
        If YaEstaRecibo(txtNumero) Then
            che "Recibo ya cargado"
            txtNumero.SetFocus
            Exit Function
        End If
    Else
        If Not IsEmpty(obtenerDeSQL("select codigo from recibos where tipodoc = 'IMC' and activo = 1 and numero = " & s2n(txtNumero))) Then
            che " Numero imputacion ya cargado"
            Exit Function
        End If
    End If
    
    'grilla cheque
    i = g.PrimerVacio(gBANCD)
    If i <> g.PrimerVacio(gNROCH) Or i <> g.PrimerVacio(gMONTO) Or i <> g.PrimerVacio(gFECHA) Or i <> g.PrimerVacio(gPT) Then
        che "revisar datos en grilla"
        Exit Function
    End If


    '-----------------------
    'montos
    dife = s2n(sumaPagos() - s2n(txtTotalRecibo))
    
    If dife < 0 Then
        che "No coinciden los montos" & vbCrLf & vbCrLf _
            & "Efectivo = " & s2n(txtefectivo) & vbCrLf _
            & "Cheques = " & g.suma(gMONTO) & vbCrLf _
            & " Retenc = " & gRet.suma(gR_IMP) & vbCrLf _
            & " Transf = " & s2n(txtTransferencia) & vbCrLf & vbCrLf _
            & " Total = " & s2n(txtTotalRecibo) & vbCrLf _
            & " Faltan: " & dife
        Exit Function
    End If


    If dife > 0 Then
        
        che "No coinciden los montos" & vbCrLf & vbCrLf _
            & "Efectivo = " & s2n(txtefectivo) & vbCrLf _
            & "Cheques = " & g.suma(gTOTAL) & vbCrLf _
            & " Retenc = " & gRet.suma(gR_IMP) & vbCrLf _
            & " Transf = " & s2n(txtTransferencia) & vbCrLf & vbCrLf _
            & " Total = " & s2n(txtTotalRecibo) & vbCrLf _
            & " Diferencia: " & dife
    
        If mTipoRecibo = trIMPUTACION Then
            Exit Function
        Else
            If Not confirma("Pasar " & dife & " a cuenta ? ") Then
                Exit Function
            End If
        End If
                    
        If g.suma(gTOTAL) > 0 Then ' Hay cheques
                'si hay diferencia, veo si es un cheque q puedo pasar a cuenta
                'que carajo quiere decir con esto es cualquiera
                'lo sig hace esto
                'toma el valor del primer cheque y lo compara con la diferencia ???? este si que se fumaba la vida
                'despues si es mayor lo descuenta del total de la factura que esta en el foco total de cheques ?? mamita
                'si es menor muestra un mensaje ?? asi funciona gustavo
                'raul
                If s2n(g.tx(1, gTOTAL)) > dife Then
                    g.tx 1, gMONTO, s2n(g.tx(1, gTOTAL)) - dife
                Else
                    che "el 1er cheque tiene monto menor al sobrante"
                    Exit Function
                End If
        End If
    End If
    '-----------------------
    
    
    If s2n(txtefectivo) <> 0 And Trim(txtcaja) = "" Then
        che "revisar caja efectivo"
        Exit Function
    End If
    
    If s2n(txtTransferencia) <> 0 And uCtaBanco.codigo = 0 Then
        che "Falta cuenta banco"
        Exit Function
    End If
    
    If s2n(txtTransferencia) <> 0 And uCtaBanco.codigo > 0 Then
        Dim pSaldo As Double
        pSaldo = dife - s2n(txtTransferencia)
        If pSaldo > 0 Then
            If MsgBox("La cantidad a transferir no es necesaria, con los otros valores alcanza el total a pagar." & Chr(13) & "¿Desea continuar con la operacion?", vbExclamation + vbYesNo, "Operacion no necesaria") = vbNo Then
                Exit Function
            End If
        End If
    End If

    'retenciones
    Dim hay As Boolean
    With gRet
        hay = False
        For i = 1 To .rows - 1
            If s2n(.tx(i, gR_NRO)) > 0 Then
                hay = True
                If .tx(i, gR_FEC) = "" Or s2n(.tx(i, gR_IMP)) = 0 Or Trim(.tx(i, gR_TIP)) = "" Then
                    che "item no completado"
                    Exit Function
                End If
            Else
                If s2n(.tx(i, gR_IMP)) > 0 Then ' unico caso grave
                    che "Falta Nro Retencion"
                    Exit Function
                End If
            End If
        Next i
    End With
    

    Falta = False
End Function

Private Function HayFAE() As Boolean ' FAE NCE NDE
    If g4F.buscar(g4TIPO, "FAE") > 0 Then HayFAE = True
    If g4C.buscar(g4TIPO, "NCE") > 0 Then HayFAE = True
    If g4D.buscar(g4TIPO, "NDE") > 0 Then HayFAE = True
End Function
 
Private Function sumaPagos()
    sumaPagos = s2n(g.suma(gTOTAL) + s2n(txtefectivo) + gRet.suma(gR_IMP) + s2n(txtTransferencia))
End Function
 
Private Function GrabaRecibo() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim CantCheques, i As Long, nMoviC, MoviConcepto, DetConcepto, nMoviB
    Dim chNro, chMonto As Double, chCod, chFech As Date, chPT, chCta, bco, chTotal As Double
    Dim arrayMsgCheques As Variant, iddoc As Long, AsiVta As New Asiento
    Dim asse As String
    Dim TextoAsientoComprobante As String
    Dim cuentaExt As Boolean
    Dim movbanc As Long, tra As Double ' transferencia
    Dim z As Double
   
    z = 1

    asse = "Numero de cuenta"
    
    CantCheques = g.PrimerVacio(gMONTO) - 1
    MoviConcepto = "Recibo " & Format(txtNumero, "00000000") & "Cliente " & Format(cliente.codigo, "0000") & "   " & cliente.DESCRIPCION
    DetConcepto = "RecCta " & Format(txtNumero, "00000000") & "  Cliente " & Format(cliente.codigo, "0000")
    movbanc = nuevoCodigo("movibanc", "movBanco")
    
    Dim DifTot As Double, DifChe As Double ', difEfe As Double  ' Para pasar excedente como Pago a Cuenta
   

    DE_BeginTrans

    asse = "1 doc y cabecera asiento"
    iddoc = NuevoDocumento(lblTipo, numero(), 0, 0)
    
    TextoAsientoComprobante = "Rec " & numero()
       
    AsiVta.nuevo "Rec " & cmbCliente, dtFecha, lblTipo

    asse = "SP cabecera"
    DataEnvironment1.dbo_abmRecibos "A", s2n(lblCodigo), lblTipo, numero(), dtFecha, cliente.codigo, s2n(txtTotalRecibo), 0, Date, UsuarioActual(), iddoc, z, 1
    
    DifTot = Round(sumaPagos() - s2n(txtTotalRecibo), 2)
    If DifTot > 0 Then
        asse = "recibo a cuenta "
        DataEnvironment1.dbo_abmFacturaVenta "A", nuevoCodigo("FacturaVenta", "codigo"), "RAA", numero(), dtFecha, 0, 0, 0, cliente.codigo, cliente.DESCRIPCION, "", "", 0, 0, DifTot, 0, 0, DifTot, DifTot, 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, 0, 0, iddoc         ' s2n(txtcotizacion),
    End If
    
    
    
    asse = "grabadet grilla"
    'detalle
    GrabaDet g4F, iddoc ', AsiVta    '*****************************???????????
    GrabaDet g4C, iddoc ', AsiVta
    GrabaDet g4D, iddoc ', AsiVta
    GrabaDet g4R, iddoc ', AsiVta
    

    cuentaExt = HayFAE()
    'cuenta debe p' detalle
    asse = "cuenta anticipo"
    If cuentaExt Then
        AsiVta.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE_EXT), g4R.suma(g4PAGA), 0, TextoAsientoComprobante
    Else
        AsiVta.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), g4R.suma(g4PAGA), 0, TextoAsientoComprobante
    End If
    
    asse = "cuenta contable cliente"
    Dim tiene_c, CUENTA_DEUDxVENTAS As String
    tiene_c = obtenerDeSQL("select tiene_cuenta from clientes where codigo = " & txtCodCliente)
        If tiene_c = 1 Then
            CUENTA_DEUDxVENTAS = obtenerDeSQL("select cuenta from clientes where codigo = " & txtCodCliente)
        Else
            CUENTA_DEUDxVENTAS = CuentaParam(ID_Cuenta_V_DEUDxVENTAS)
        End If
    
    AsiVta.AcumularItem CUENTA_DEUDxVENTAS, 0, g4R.suma(g4PAGA), TextoAsientoComprobante
    
    asse = "efectivo"
    'Efectivo
    If s2n(txtefectivo) > 0 Then
        'MoviCaja
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        DataEnvironment1.dbo_MOVICAJASdoc "A", s2n(txtcaja), nMoviC, cliente.codigo _
            , "E", "I", s2n(txtefectivo), MoviConcepto, dtFecha.Value, Trim(txtCuentaEfectivo) _
            , 0, 1, lblTipo, numero(), Date, UsuarioActual(), 0, iddoc
            
        AsiVta.AcumularItem txtCuentaEfectivo, s2n(txtefectivo), 0, TextoAsientoComprobante
        AsiVta.AcumularItem CUENTA_DEUDxVENTAS, 0, s2n(txtefectivo), TextoAsientoComprobante
    End If
    
    tra = s2n(txtTransferencia)
    If tra > 0 Then
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, "T", "I", tra, MoviConcepto, dtFecha, "", movbanc, 1, "REC", numero(), Date, UsuarioActual(), 0, iddoc
        DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "E", Left(MoviConcepto, 50), dtFecha, "E", tra, movbanc, iddoc, Date, UsuarioActual(), z
        AsiVta.AcumularItem uCtaBanco.CuentaContable, tra, 0, TextoAsientoComprobante
        AsiVta.AcumularItem CUENTA_DEUDxVENTAS, 0, tra, TextoAsientoComprobante
    End If


    asse = "cheques "
    'CHEQUES
    ReDim arrayMsgCheques(CantCheques) ' msg
    For i = 1 To CantCheques
        chNro = g.tx(i, gNROCH)
        chMonto = s2n(g.tx(i, gMONTO))
        chTotal = s2n(g.tx(i, gTOTAL))
        chCta = CuentaParam(ID_Cuenta_M_CH_CARTERA) 'obtenerParametro("Cta_Caja")
        chFech = CDate(g.tx(i, gFECHA))
        chPT = g.tx(i, gPT)
        bco = s2n(g.tx(i, gBANCC))
        '
        chCod = nuevoCodigo("cheques", "NroInt")
        g.tx i, gCODCH, chCod

        asse = "sp cheques "
        'Cheques
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "A", chCod, chFech, chNro, chTotal, s2n(txtNumero) _
            , lblTipo, dtFecha, Date, "C", bco, chPT, cliente.codigo, Date, UsuarioActual(), iddoc, 0
'        DataEnvironment1.dbo_INGCHEQUESTERCEROS "A", chCod, chFech, chNro, chMonto, s2n(txtNumero)_
'            , lblTipo, Date, Date, "C", bco, chPT, cliente.codigo, Date, UsuarioActual(), 0, 0

        'movi
        nMoviB = nuevoCodigo("movibanc", "movBanco")
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        '    'moviBanc Ver Nota
        DataEnvironment1.dbo_MOVIBANCOS2 "A", "0", "I", MoviConcepto, Date, "C", chMonto, nMoviB, iddoc, Date, UsuarioActual(), chCod
        asse = " sp MC"
        'MoviCaja
        DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, "C", "I", chMonto, MoviConcepto, dtFecha, chCta, 0, 1, lblTipo, s2n(txtNumero), Date, UsuarioActual, chCod, iddoc
        
'''        'DetMovCaja
'''        DataEnvironment1.dbo_DETMOVCAJAS "A", nMoviC, chMonto, cliente.codigo, 0, DetConcepto, "RA"
        
        
        arrayMsgCheques(i) = "Codigo " & chCod & ", Bco " & g.tx(i, gBANCD) & ", Nro " & chNro
        AsiVta.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), chMonto, 0, "ch " & chNro
        AsiVta.AcumularItem CUENTA_DEUDxVENTAS, 0, chMonto, TextoAsientoComprobante
        
        

        DifChe = Round(chTotal - chMonto, 2)


        If DifChe > 0 Then
            If HayFAE Then
                AsiVta.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE_EXT), 0, DifChe
            Else
                AsiVta.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), 0, DifChe
            End If
        
        
            'movi
            nMoviC = nuevoCodigo("movicaja", "movimiento")
            'MoviCaja
            DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, "C", "I", DifChe, MoviConcepto, dtFecha, CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, 1, "RAA", txtNumero, Date, UsuarioActual, chCod, iddoc
           
            AsiVta.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), DifChe, 0 ', TextoAsientoComprobante

        
        ' ************************************************
        End If
    Next i
    
    
    asse = "retenciones"
    'Retenciones
    Dim refe As Date, reim As Double, reti As String, reco As String, recu As String, reip As Long, renu As Long, refa As Long
    With gRet
    
        For i = 1 To .rows - 1
            
            renu = s2n(.tx(i, gR_NRO))
            If renu > 0 Then
                refe = CDate(.tx(i, gR_FEC))
                reim = s2n(.tx(i, gR_IMP))
                reti = .tx(i, gR_TIP)
                recu = .tx(i, gR_CUE)
                reco = .tx(i, gR_TIC)
                reip = s2n(.tx(i, gR_IdC), 0)
                refa = s2n(.tx(i, gR_FAC), 0)
                    
                AsiVta.AcumularItem recu, reim, 0
                AsiVta.AcumularItem CUENTA_DEUDxVENTAS, 0, reim, TextoAsientoComprobante
                DataEnvironment1.Sistema.Execute "insert into RecibosRetenciones " _
                    & " (iddoc, idCuentasParam, Numero,  Fecha , Importe, nroFactura, cuenta) values " _
                    & " ('" & iddoc & "', " & reip & ", '" & renu & "', " & ssFecha(refe) & ", " & x2s(reim) & ",  " & x2s(refa) & ", '" & recu & "' ) "
            End If
        Next i
    End With
    
    asse = "asiento"
    If siAsiento("AsientosRecibos") Then AsiVta.Grabar iddoc
    
    asse = " commit"
    DE_CommitTrans
    midDoc = iddoc
    ' --- Transaccion aqui -----------  fin

    GrabaRecibo = True
    If CantCheques > 0 Then
        che "Operacion concluida" & vbCrLf & "Puede anotar los codigos internos de los cheques"
        lMsg arrayMsgCheques
    End If
    asse = " impresion"
    ImprimirReciboConImputacion
fin:
    Exit Function
ufaErr:
    DE_RollbackTrans
    ufa "error grabando Recibo :" + asse, Me.Name, Err + " " + asse
    Resume fin
End Function



Private Sub ImprimirReciboConImputacion()
Dim r, i As Integer
Dim rsTemp As New ADODB.Recordset
Dim sql, str, str1 As String
Dim Fecha As Date
'On Error GoTo UFAimprimir
                               
stblRecibosTmp = TablaTempCrear(tt_ReciboCobroTemp)
stblChequestmp = TablaTempCrear(tt_Chequetmp)
stblRetencionestmp = TablaTempCrear(tt_Retenciontmp)

'FACTURAS
For r = 1 To gFact.rows - 1
    gFact.Row = r
    If Trim(gFact.TextMatrix(r, 4)) <> "" And Trim(gFact.TextMatrix(r, 4)) <> 0 Then
        sql = "Select fecha from facturaventa  where  codigo=" & gFact.TextMatrix(r, g4CODI)
        rsTemp.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
         If Not rsTemp.EOF Then
            Fecha = rsTemp!Fecha
            sql = "insert into " & stblRecibosTmp & " (tipodoc, nrodoc, fecha, saldo,cobrado,ultimosaldo) " _
            & " values (" & ssTexto(gFact.TextMatrix(r, 0)) & ", " & gFact.TextMatrix(r, 1) & ", " & ssFecha(Fecha) & ", " & x2s(gFact.TextMatrix(r, 3)) & ", " & x2s(s2n(gFact.TextMatrix(r, 4))) & "," & x2s(s2n(s2n(gFact.TextMatrix(r, 3)) - s2n(gFact.TextMatrix(r, 4)))) & ")"
            DataEnvironment1.Sistema.Execute sql
            rsTemp.Close
            
         End If
         Set rsTemp = Nothing
    End If
Next r
    
'DEBITOS
For r = 1 To gDebi.rows - 1
    gDebi.Row = r
    If Trim(gDebi.TextMatrix(r, 4)) <> "" And Trim(gDebi.TextMatrix(r, 4)) <> 0 Then
        sql = "Select fecha,tipodoc from facturaVenta where codigo=" & gDebi.TextMatrix(r, g4CODI)
        rsTemp.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      If Not rsTemp.EOF Then
        Fecha = rsTemp!Fecha
        sql = "insert into " & stblRecibosTmp & " (tipodoc,nrodoc,fecha,saldo,cobrado,ultimosaldo) " _
        & " values (" & ssTexto(gDebi.TextMatrix(r, 0)) & "," & gDebi.TextMatrix(r, 1) & ", " & ssFecha(Fecha) & ",  " & x2s(gDebi.TextMatrix(r, 3)) & "," & x2s(s2n(gDebi.TextMatrix(r, 4))) & "," & x2s(s2n(s2n(gDebi.TextMatrix(r, 3)) - s2n(gDebi.TextMatrix(r, 4)))) & ")"
        DataEnvironment1.Sistema.Execute sql
        rsTemp.Close
        
      End If
      Set rsTemp = Nothing
    End If
Next r

'CREDITOS
For r = 1 To gCred.rows - 1
    gCred.Row = r
    If Trim(gCred.TextMatrix(r, 4)) <> "" And Trim(gCred.TextMatrix(r, 4)) <> 0 Then
        sql = "Select fecha,tipodoc from facturaVenta where codigo=" & gCred.TextMatrix(r, g4CODI)
        rsTemp.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      If rsTemp.EOF And rsTemp.BOF Then
      Else
        Fecha = rsTemp!Fecha
        sql = "insert into " & stblRecibosTmp & " (tipodoc,nrodoc,fecha,saldo,cobrado,ultimosaldo) " _
        & " values (" & ssTexto(gCred.TextMatrix(r, 0)) & "," & gCred.TextMatrix(r, 1) & ", " & ssFecha(Fecha) & ",'" & x2s(gCred.TextMatrix(r, 3)) & "', " & x2s(s2n(gCred.TextMatrix(r, 4))) & "," & x2s(s2n(s2n(gCred.TextMatrix(r, 3)) - s2n(gCred.TextMatrix(r, 4)))) & ")"
        DataEnvironment1.Sistema.Execute sql
        rsTemp.Close
        
      End If
      Set rsTemp = Nothing
    End If
Next r

'RECIBOS A CUENTA
For r = 1 To gReci.rows - 1
    gReci.Row = r
    If Trim(gReci.TextMatrix(r, 4)) <> "" And Trim(gReci.TextMatrix(r, 4)) <> 0 Then
        sql = "Select fecha,tipodoc from facturaVenta where codigo=" & gReci.TextMatrix(r, g4CODI)
        rsTemp.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      If Not rsTemp.EOF Then
        Fecha = rsTemp!Fecha
        sql = "insert into " & stblRecibosTmp & " (tipodoc,nrodoc,fecha,saldo,cobrado,ultimosaldo) " _
        & " values (" & ssTexto(gReci.TextMatrix(r, 0)) & "," & gReci.TextMatrix(r, 1) & ", " & ssFecha(Fecha) & "," & x2s(gReci.TextMatrix(r, 3)) & " ," & x2s(s2n(gReci.TextMatrix(r, 4))) & "," & x2s(s2n(s2n(gReci.TextMatrix(r, 3)) - s2n(gReci.TextMatrix(r, 4)))) & ")"
        DataEnvironment1.Sistema.Execute sql
        rsTemp.Close
        
      End If
      Set rsTemp = Nothing
    End If
Next r

Dim rptRecibo As New RptReciboImputacion
Dim rptImputacion As New RptReciboConImputacion
Dim rptValores As New RptDetalledeValores

If txtTotalDoc = 0 Then
   With rptRecibo
        sql = "select * from " & stblRecibosTmp
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = sql
        .lblfecha = dtFecha
        .lblTitulo = "IMPUTACION Nº: " & Format(txtNumero, "00000000")
        .lblcliente = cmbCliente.Text
        .Restart
        
        If PREVIEW_IMPRESIONES Then
            .Show
        Else
            .PrintReport False
        End If
        
   End With
 Else
    'SUBREPORTE VALORES
    With gCheques
        For r = 1 To .rows - 1
          If .TextMatrix(r, 0) <> "" Or .TextMatrix(r, 1) <> "" Then
           str = "insert into " & stblChequestmp & "(nroint, banco, cheque, importe, fecha, propio) values( " & _
           " " & sSinNull(.TextMatrix(r, 7)) & ", '" & sSinNull(.TextMatrix(r, 1)) & "', '" & sSinNull(.TextMatrix(r, 2)) & "'," & x2s(.TextMatrix(r, 3)) & ", " & ssFecha(.TextMatrix(r, 5)) & ", '" & sSinNull(.TextMatrix(r, 6)) & "')"
           DataEnvironment1.Sistema.Execute str
          End If
        Next r
    End With
        str = "select * from " & stblChequestmp
        rptValores.data1.Connection = DataEnvironment1.Sistema
        rptValores.data1.Source = str
        rptImputacion.SubReport1.object = rptValores
        
        'SUBREPORTE RETENCIONES
    With gRetenciones
        For r = 1 To .rows - 1
          If .TextMatrix(r, 0) <> "" Or .TextMatrix(r, 1) <> "" Then
            str1 = "insert into " & stblRetencionestmp & "(nroret, fecha, codret, tiporet,importe) values( " & _
            " " & sSinNull(.TextMatrix(r, 0)) & ", " & ssFecha(.TextMatrix(r, 1)) & ", '" & sSinNull(.TextMatrix(r, 2)) & "','" & x2s(.TextMatrix(r, 3)) & "', " & x2s(.TextMatrix(r, 4)) & ")"
            DataEnvironment1.Sistema.Execute str1
          End If
        Next r
    End With
    
    
    
        str1 = "select * from " & stblRetencionestmp
        RptDetalledeRetenciones.dataRet.Connection = DataEnvironment1.Sistema
        RptDetalledeRetenciones.dataRet.Source = str1
        rptImputacion.SubReport2.object = RptDetalledeRetenciones
    
        'REPORTE PRINCIPAL
    With rptImputacion
        sql = "select * from " & stblRecibosTmp
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = sql
        .lblfecha = dtFecha
        .lblTitulo = "RECIBO CON IMPUTACION Nº: " & Format(txtNumero, "00000000")
        .lblcliente = "RECIBIMOS DE  " & cmbCliente.Text
        .lblvalor = "LA CANTIDAD DE PESOS  " & NroEnLetras(txtTotalRecibo)
        .lblefectivo = Format(txtefectivo, "#,##0.00")
        .lblcheques = Format(lblSumaCheques, "#,##0.00")
        
        If midDoc = 0 Then
            .lbltransf = 0
           Else
            .lbltransf = obtenerDeSQL("select importe from movicaja where iddoc = " & midDoc & " and tipo = 'T' ")
        End If
        
        .lblRetenciones = Format(lblSumaRet, "#,##0.00")
        .lbltotal = Format(txtTotalRecibo, "#,##0.00")
    End With
        
        rptImputacion.Restart
        If PREVIEW_IMPRESIONES Then
            rptImputacion.Show
        Else
            rptImputacion.PrintReport False
        End If
End If
fin:
    Exit Sub
'UFAimprimir:
    ufa "error en la impresión", Me.Name
    Resume fin:
End Sub


Private Sub GrabaDet(gg As LiGrilla, iddo As Long) ', asi As Asiento)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim i As Long, codRec, codFac, paga, saldVie, saldNue, codTipo, CodProv, codCompras
    
    For i = 1 To gg.rows - 1
        codTipo = gg.tx(i, g4TIPO)
        codRec = s2n(lblCodigo)
        saldVie = s2n(gg.tx(i, g4SALD))
        paga = s2n(gg.tx(i, g4PAGA))
        saldNue = s2n(saldVie - paga)
        codFac = gg.tx(i, g4CODI)
        If paga <> 0 Then
            'DataEnvironment1.dbo_abmRecibosDetalle codRec, codFac, paga, iddo
            If codTipo = "FAC" Then
                DataEnvironment1.dbo_abmRecibosDetalle codRec, codFac, paga, saldVie, iddo, codTipo
                CodProv = obtenerDeSQL("select codpr from transcom where id=" & gg.tx(i, g4CODI))
                If saldVie - paga = 0 Then
                    DataEnvironment1.dbo_MODIFICOSALDOYPASOREG CodProv, codTipo, gg.tx(i, g4NUME)
                    DE_CommitTrans
                    DE_BeginTrans
                    codCompras = obtenerDeSQL("select id from compras where nrodoc=" & gg.tx(i, g4NUME))
                    DataEnvironment1.Sistema.Execute "update compras set [plan]=" & codRec & " where id=" & codCompras
                Else
                    DataEnvironment1.dbo_MODIFICOSALDOTRANS CodProv, codTipo, gg.tx(i, g4NUME), saldNue
                End If
                
            Else
                DataEnvironment1.dbo_abmRecibosDetalle codRec, codFac, paga, saldVie, iddo, codTipo
                DataEnvironment1.Sistema.Execute "update FacturaVenta set saldo =" & ssNum(saldNue) & " where codigo = " & codFac  'saldo = saldo - " & x2s(paga) & " where codigo = " & codFac
            End If
        End If
    Next i
Exit Sub
ufaErr:
    MsgBox "Error durante la grabacion del detalle.", vbCritical, "Alvertencia"
End Sub

Private Sub cargaRecibo(codigoRecibo)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim i As Long, rs As New ADODB.Recordset, s As String
    Dim Movi As Long, temp As Variant ', tmpTipo As String

    'tmpTipo = IIf(mTipoRecibo = trIMPUTACION, "IMC", "REC")
    
    s = "select codigo as [interno], TipoDoc as [ TIPO DOC ] ,Numero as [ NRO RECIBO ], Fecha, Cliente, Total, iddoc " _
        & " from Recibos where activo = 1 and codigo = " & codigoRecibo
    temp = obtenerDeSQL(s)

    limpiar
    'cabecera
    lblCodigo = codigoRecibo
    lblTipo = sSinNull(temp(1))
    txtNumero = s2n(temp(2))
    dtFecha = CDate(temp(3))
    cliente.codigo = s2n(temp(4))
    txtTotalRecibo = s2n(temp(5))
    midDoc = s2n(temp(6))
            
    'cargar efectivo y cheques y grillas
    'efectivo egreso TipDoc NroDoc
    s = "select caja, movimiento, importe, cuenta, cotizacion, tipo " _
        & " from movicaja where TipoDoc = '" & lblTipo & "' and ing_egr = 'I' and tipo = 'E' and NroDoc = " & numero()
    temp = obtenerDeSQL(s)
    If IsEmpty(temp) Then
        txtcaja = ""
        txtefectivo = 0
        txtCuentaEfectivo = 0
    Else
        txtcaja = s2n(temp(0))
        Movi = temp(1)
        txtefectivo = s2n(temp(2))
        txtCuentaEfectivo = s2n(temp(3))
        ' cotizacion ?  = s2n(temp(4))
    End If
    
    'grillas
    carga4 True
    
    'cheques
    With rs
        s = "select NroInt, fecha, nro, importe, banco_Nro, procedencia " _
            & " from cheques where tdoc = '" & lblTipo & "' and nDoc = " & numero() & " and activo = 1 "
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        g.rows = 1
        While Not .EOF
            i = g.addRow()
            g.tx i, gCODCH, s2n(!NroInt, 0)
            g.tx i, gFECHA, !Fecha
            g.tx i, gNROCH, !Nro
            g.tx i, gBANCC, !BANCO_NRO
            'g.tx i, gBANCD,
            g.tx i, gPT, !procedencia
            g.tx i, gMONTO, s2n(!Importe)
            .MoveNext
        Wend
        .Close
        gCheques.TopRow = 1
    End With
    
    'retenciones
    Dim tempo
    With rs
        s = "select * from RecibosRetenciones where iddoc = '" & midDoc & "' "
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        gRet.rows = 1
        While Not .EOF
            i = gRet.addRow()
            
            gRet.tx i, gR_NRO, !numero
            gRet.tx i, gR_FEC, !Fecha
            gRet.tx i, gR_TIC, sSinNull(obtenerDeSQL("select codigo from CuentasParam where id = '" & !idCuentasParam & "' "))
            gRet.tx i, gR_TIP, sSinNull(obtenerDeSQL("select Descripcion from CuentasParam where id = '" & !idCuentasParam & "' "))
            gRet.tx i, gR_IMP, !Importe
            gRet.tx i, gR_CUE, !Cuenta
            gRet.tx i, gR_IdC, !idCuentasParam
            gRet.tx i, gR_FAC, !NroFactura
            .MoveNext
        Wend
        .Close
        gRetenciones.TopRow = 1
    End With
    
    'transf
    tempo = obtenerDeSQL("select cuenta, importe from movibanc where cuenta <>0 and iddoc = " & midDoc)
    If Not IsEmpty(tempo) Then
        txtTransferencia = tempo(1)
        uCtaBanco.codigo = tempo(0)
    End If
    
    GoTo fin
ufaErr:
    ufa "err cargando recibo", Me.Name & " " & codigoRecibo ', Err
fin:
    Set rs = Nothing
End Sub

Private Function BuscaRecibo() As Boolean
    'busca en tabla recibos
    ' y recibosDetalle

    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim i As Long, s As String ', rs As New ADODB.Recordset
    Dim Movi As Long, temp As Variant ', tmpTipo As String

'    tmpTipo = IIf(mTipoRecibo = trIMPUTACION, "IMC", "REC")
    
    s = "select codigo as [interno], TipoDoc as [ TIPO DOC ] ,Numero as [ NRO RECIBO ], Fecha, Cliente, Total " _
        & " from Recibos where activo = 1 and fecha " & ucFechas.ssBetween() & " and TipoDoc = '" & TipoRecibo() & "' " _
        & " order by codigo desc "
        
        
    If frmBuscar.MostrarSql(s) > "" Then
        cargaRecibo s2n(frmBuscar.resultado(1))
        BuscaRecibo = True
    End If
    
'    With frmBuscar
'        Limpiar
'        'cabecera
'        lblCodigo = s2n(.resultado(1))
'        lblTipo = sSinNull(.resultado(2))
'        txtNumero = s2n(.resultado(3))
'        dtFecha = CDate(.resultado(4))
'        cliente.codigo = s2n(.resultado(5))
'        txtTotalRecibo = s2n(.resultado(6))
'    End With
'    'cargar efectivo y cheques y grillas
'    'efectivo
'    s = "select caja, movimiento, importe, cuenta, cotizacion " _
'        & " from movicaja where TipoDoc = '" & lblTipo & "' and NroDoc = " & Numero()
'    temp = obtenerDeSQL(s)
'    If Not IsEmpty(temp) Then
'        txtCaja = s2n(temp(0))
'        movi = temp(1)
'        txtEfectivo = s2n(temp(2))
'        txtCuentaEfectivo = s2n(temp(3))
'        ' cotizacion ?  = s2n(temp(4))
'    End If
'    'cheques
'    With rs
'        s = "select NroInt, fecha, nro, importe, banco_Nro, procedencia " _
'            & " from cheques where tdoc = '" & lblTipo & "' and nDoc = " & Numero() & " and activo = 1 "
'        .Open s, daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'        While Not .EOF
'            i = g.addRow()
'            g.tx i, gCODCH, s2n(!nroint, 0)
'            g.tx i, gFECHA, !fecha
'            g.tx i, gNROCH, !Nro
'            g.tx i, gBANCC, !Banco_Nro
'            'g.tx i, gBANCD,
'            g.tx i, gPT, !procedencia
'            g.tx i, gMONTO, s2n(!importe)
'            .MoveNext
'        Wend
'        .Close
'        'grillas
''        s = "select FacturaVenta, Importe " _
'            & " from ReciboDetalle where activo = 1 and CodRecibo = " & s2n(lblCodigo)
''        .Open s, daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'        carga4 True
'    End With
'    BuscaRecibo = True

fin:
    'Set rs = Nothing
    Exit Function
ufaErr:
    ufa "err cargando recibo", Me.Name & " " & frmBuscar.resultado(1) ', Err
    Resume fin
End Function

Private Function EliminaRecibo() As Boolean


'    'BUSCO EN RAC
'    asse = "13 Rac - "
'    If rac.TextMatrix(1, 0) <> "" Then
'        For x = 1 To rac.rows - 1
'            DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, "RAC", rac.TextMatrix(x, 0)
'            Importe = ObtenerImporte("RAC", s2n(rac.TextMatrix(x, 0)))
'            If Importe > 0 Then ' sumo al saldo lo q hay en relfnr_c
'                DataEnvironment1.dbo_sumoSALDOTRANS CodProv, "RAC", rac.TextMatrix(x, 0), Importe
'                'DataEnvironment1.dbo_DOYDEBAJARELFNRC CodProv, "RAC", Val(rac.TextMatrix(x, 0)), NumOPAGO
'            End If
'        Next
'    End If



    '--anula recibo---
    '   borra cheques
    '   borra movicaja
    
    'lo distinto a ReciboCuenta:
    '   barre reciboDetalle y ReActualiza  FacturaVenta.saldo
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset, tmp As Variant
    Dim CodProv, NroDocProv
    
'   excepto txtNumero, que SI puede ser 0
'    If txtNumero = "" Or lblTipo = "" Or s2n(lblCodigo) = 0 Then
    If lblTipo = "" Or s2n(lblCodigo) = 0 Then
        che "no puedo eliminar"
        Exit Function
    End If
    
    tmp = obtenerDeSQL("select activo from recibos where codigo = " & lblCodigo)
    
    If IsNull(tmp) Then
        ufa "err: No encuentro recibo cod " & lblCodigo, Me.Name ', Err
        Exit Function
    End If
    If tmp = 0 Then
        che "Recibo ya eliminado"
        Exit Function
    End If
    
    
'----------------------------------------------------
    DE_BeginTrans
    
'= a ReciboCuenta

    'cheques
    'daTaenvironment1.dbo_INGRESOCHTERCEROS "B",
    DataEnvironment1.Sistema.Execute "update cheques set activo = 0, fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & UsuarioActual() & " where ndoc = '" & txtNumero & "' and tdoc = '" & lblTipo & "'"
    'movicaja
    DataEnvironment1.dbo_MOVICAJASdoc "B", 0, 0, 0, "", "", 0, "", Date, "", 0, 0, lblTipo, txtNumero, Date, UsuarioActual(), 0, midDoc
    DataEnvironment1.dbo_MOVIBANCOS "B", 0, "", "", Date, "", 0, 0, midDoc, Date, UsuarioActual, 0
    'DataEnvironment1.dbo_MOVIBANCOS "B", "", "", "", Date, "", 0, "", midDoc, Date, UsuarioActual, 0
    
'propio de Recibo
    'borra Recibo
    DataEnvironment1.Sistema.Execute "update recibos set activo=0 where codigo = " & s2n(lblCodigo)
    
    'si fue recibo con pago a cuenta
    DataEnvironment1.Sistema.Execute "update FacturaVenta set activo = 0 where iddoc = " & midDoc
    
    'actualiza saldo con recibodetalle
    With rs
        .Open "select facturaventa, importe,_tfac as tipo from RecibosDetalle where CodRecibo = " & s2n(lblCodigo), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            If Trim(!Tipo) = "FAC" Then
                CodProv = obtenerDeSQL("select C.codpr from transcom C where C.id=" & !FACTURAVENTA & " UNION select C.codpr from COMPRAS C where C.id=" & !FACTURAVENTA)
                NroDocProv = obtenerDeSQL("select C.NRODOC from transcom C where C.id=" & !FACTURAVENTA & " UNION select C.NRODOC from COMPRAS C where C.id=" & !FACTURAVENTA)
                DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, "FAC", NroDocProv
                DataEnvironment1.dbo_sumoSALDOTRANS CodProv, "FAC", NroDocProv, !Importe
            Else
                DataEnvironment1.Sistema.Execute "update facturaventa set saldo = saldo + " & x2s(!Importe) & " where codigo = " & !FACTURAVENTA
            End If
            
            .MoveNext
        Wend
    End With
    
'comun a ambos
        'Baja Doc y asiento
        If Not BorroDocumento(midDoc) Then
            ufa "err al borrar documento", " middoc = " & midDoc
            DE_RollbackTrans
            GoTo fin: ' sorry no way, es VB, no java
        End If
        
        DataEnvironment1.Sistema.Execute "delete from RecibosRetenciones where iddoc = '" & midDoc & "' "
        grabaBitacora "B", midDoc, "recibos"
    
    DE_CommitTrans
'----------------------------------------------------


    EliminaRecibo = True
    che "Eliminado"
    limpiar
    GoTo fin

ufaErr:
    DE_RollbackTrans
    ufa "error en la baja", Me.Name & lblCodigo & " " & txtNumero ', Err
fin:
End Function

Private Function verCajaEfectivo() As Boolean
    Dim tmp 'As String
   
    tmp = obtenerDeSQL("select cuenta from cajas where codigo = " & s2n(txtcaja))
    If Not IsEmpty(tmp) Then 'tmp > "" Then
        verCajaEfectivo = True
        txtCuentaEfectivo = tmp
    Else
        che "No existe la caja"
        verCajaEfectivo = False
    End If
End Function

Private Function numero() As Double
    numero = s2n(txtNumero)
End Function
Private Function TipoRecibo() As String
    TipoRecibo = IIf(mTipoRecibo = trRECIBO, "REC", "IMC")
End Function

Private Sub ucMenu_SeMovio()
    'lblCodigo = ucMenu.rs!codigo
    cargaRecibo ucMenu.rs!codigo
End Sub


'19/11/4    adapt licodigo, +cmd, +where
'2/12/4     msg de ok
'           control de duplicado
'3/1/5      fix no existe caja
'21/1/5     muestro cod cheques aparte
'16/3/5     pone saldo con hago dbl clic, no al pasar por encima
'15/4/5    SubimeSi800x600
'12/7/5     s2n txtefectivo gotfocus
