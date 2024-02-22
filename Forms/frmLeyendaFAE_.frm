VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLeyendaFAE_ 
   Caption         =   "Leyenda FAE"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   390
      Left            =   5415
      TabIndex        =   2
      Top             =   5430
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "ok"
      Height          =   390
      Left            =   6915
      TabIndex        =   1
      Top             =   5430
      Width           =   1140
   End
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Align           =   1  'Align Top
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8370
      _cx             =   14764
      _cy             =   9340
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
      Rows            =   25
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLeyendaFAE_.frx":0000
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
      Editable        =   2
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
Attribute VB_Name = "frmLeyendaFAE_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Renglones()
    Renglones = 10
End Function




Private Sub cmdGrabar_Click()
    Dim i As Long
    
    With DataEnvironment1.AMR
    .Execute "delete from FacturaVentaLeyendaFAE"
    For i = 1 To Renglones()
        .Execute "insert into FacturaVentaLeyendaFAE (renglon) values ('" & ssStr(gri.TextMatrix(i, 0)) & "') "
    Next i
    che "grabado"
    End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LlenarGrilla gri, "select top " & Renglones() & " renglon from FacturaVentaLeyendaFae order by id", False
    gri.rows = Renglones() + 1
End Sub
