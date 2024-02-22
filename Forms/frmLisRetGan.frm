VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisRetIva 
   Caption         =   "Listado de Retencion de Iva"
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   10965
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   285
      Left            =   165
      TabIndex        =   3
      Top             =   105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   96665601
      CurrentDate     =   40263
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   840
      Left            =   7440
      TabIndex        =   2
      Top             =   105
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1482
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   825
      Left            =   8430
      TabIndex        =   1
      Top             =   90
      Width           =   840
   End
   Begin VSFlex7Ctl.VSFlexGrid gGan 
      Height          =   9885
      Left            =   135
      TabIndex        =   0
      Top             =   975
      Width           =   10200
      _cx             =   17992
      _cy             =   17436
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Top             =   90
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   81788929
      CurrentDate     =   40263
   End
End
Attribute VB_Name = "frmLisRetIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

