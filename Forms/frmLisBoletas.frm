VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisBoletas 
   Caption         =   "Informe de Boletas"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "frmLisBoletas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   390
      Left            =   4635
      TabIndex        =   3
      Top             =   75
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   300
      Left            =   2175
      TabIndex        =   2
      Top             =   105
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   39895
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   39895
   End
   Begin VSFlex7LCtl.VSFlexGrid gBoletas 
      Height          =   5250
      Left            =   120
      TabIndex        =   0
      Top             =   585
      Width           =   11055
      _cx             =   19500
      _cy             =   9260
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
End
Attribute VB_Name = "frmLisBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVer_Click()
Dim s As String, i As Long
s = "select FECHA,CODPR AS PROV,RAZONSOCIAL,TIPODOC,NRODOC,TOTAL from GastosBoletas where activo=1 and fecha>=" & ssFecha(dtDesde) & " and fecha<=" & ssFecha(dtHasta) & " order by codpr"
LlenarGrilla gBoletas, s, False
With gBoletas
    'For i = 1 To .rows - 1
    'Next
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
End With
End Sub

Private Sub Form_Load()
dtDesde = CDate("01/01/" & Year(Date))
dtHasta = Date
End Sub
