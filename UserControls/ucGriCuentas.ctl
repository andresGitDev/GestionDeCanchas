VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.UserControl ucGriCuentas 
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   ScaleHeight     =   4650
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Quitar"
      Height          =   810
      Left            =   7830
      Picture         =   "ucGriCuentas.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   465
      Width           =   720
   End
   Begin VB.TextBox txtTotal 
      Height          =   285
      Left            =   7815
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   30
      Width           =   1290
   End
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Align           =   3  'Align Left
      Height          =   4650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _cx             =   13652
      _cy             =   8202
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
      Rows            =   50
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
End
Attribute VB_Name = "ucGriCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private Enum eColu
'    coluCuenta
'    coluImporte
'    coluConcepto
'End Enum
'mcols = 3
'
'Public Event ClicImporte(row As Long, col As Long, ImporteActual As Double)
'
'Public Property Get rows() As Long
'    rows = gri.rows
'End Property
'
'Private Sub cmdBorrar_Click()
'    If gri.row > 0 Then gri.RemoveItem gri.row
'    recalcular
'End Sub
'
'Private Sub gri_DblClick()
'    Dim resu As String
'
'    If gri.row = 0 Then Exit Sub
'    If Not UserControl.Enabled Then Exit Sub
'
'
'
'    Select Case gri.col
'    Case colCuenta
'        resu = BuscarCuenta(False, False)
'        If resu > "" Then
'            gri.TextMatrix(row, col) = resu
'    Case coluImporte
'        RaiseEvent ClicImporte(gri.row, gri.col, nSinNull(gri.Text))
'    Case coluConcepto
'
'    End Select
'
'    recalcular
'End Sub
'
'Private Sub UserControl_Initialize()
'    gri.cols = mcols
'    gri.rows = 1
'    gri.rows = 2
'
'End Sub
'
'Private Sub UserControl_Resize()
'    Dim x As Long
'    x = UserControl.Width - txtTotal.Width + 20
'    gri.Width = x
'    txtTotal.Left = x
'    cmdBorrar.Left = x
'End Sub
'
'Private Function recalcular()
'    Dim i As Long, total As Double
'    For i = 1 To gri.rows
'        total = total + nSinNull(gri.TextMatrix(i, coluImporte))
'    Next i
'    txtTotal = Format(total, "standard")
'    recalcular = total
'End Function
Private Sub cmdBorrar_Click()

End Sub
