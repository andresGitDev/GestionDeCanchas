VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLeyendaFAE 
   Caption         =   "Leyenda FAE"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "frmLeyendaFAE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   390
      Left            =   3345
      TabIndex        =   2
      Top             =   5400
      Width           =   1140
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "cerrar"
      Height          =   390
      Left            =   4845
      TabIndex        =   1
      Top             =   5400
      Width           =   1140
   End
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Align           =   1  'Align Top
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6195
      _cx             =   10927
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
      FormatString    =   $"frmLeyendaFAE.frx":08CA
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
Attribute VB_Name = "frmLeyendaFAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Function Renglones()
    Renglones = 20
End Function

Private Sub cmdGrabar_Click()
    Dim i As Long
    
    With DataEnvironment1.Sistema
    .Execute "delete from FacturaVentaLeyendaFAE"
    For i = 1 To Renglones()
        .Execute "insert into FacturaVentaLeyendaFAE (renglon) values ('" & ssStr(Left(gri.TextMatrix(i, 0), 50)) & "') "
    Next i
    che "grabado"
    End With
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LlenarGrilla gri, "select top " & Renglones() & " renglon from FacturaVentaLeyendaFae order by id", False
    gri.rows = Renglones() + 1
End Sub

Private Sub gri_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    gri.EditMaxLength = 50
End Sub

