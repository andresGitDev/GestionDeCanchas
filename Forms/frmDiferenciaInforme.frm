VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiferenciaInforme 
   Caption         =   "Informe de Ajustes"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   Icon            =   "frmDiferenciaInforme.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   150
      Picture         =   "frmDiferenciaInforme.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   765
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   795
      Left            =   4365
      TabIndex        =   0
      Top             =   105
      Width           =   780
      _extentx        =   1376
      _extenty        =   1402
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4350
      Left            =   120
      TabIndex        =   1
      Top             =   5865
      Width           =   10125
      _cx             =   17859
      _cy             =   7673
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
   Begin VSFlex7LCtl.VSFlexGrid gAjustes 
      Height          =   4770
      Left            =   105
      TabIndex        =   2
      Top             =   990
      Width           =   10110
      _cx             =   17833
      _cy             =   8414
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
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   555
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   17235969
      CurrentDate     =   39594
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   2835
      TabIndex        =   5
      Top             =   555
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   17235969
      CurrentDate     =   39594
   End
End
Attribute VB_Name = "frmDiferenciaInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVer_Click()
    Ver
End Sub

Private Sub Form_Load()
    iFechas
    Ver
End Sub

Sub iFechas()
    dtDesde = "01/" & Month(Date) & "/" & Year(Date)
    dtHasta = Date
End Sub

Private Function Ver()
Dim cver As String, dProd As Long
    gAjustes.rows = 1
    cver = "Select R.MOVIMIENTOINTERNO AS NroAjuste, R.Fecha, C.DESCRIPCION from RemitoDiferenciaStock R INNER JOIN CONCEPTOS C ON R.CONCEPTO=C.CODIGO where R.activo=1 and (R.fecha>= " & ssFecha(dtDesde) & " and R.fecha<=" & ssFecha(dtHasta) & " )"
    LlenarGrilla gAjustes, cver, True
    gDetalle.rows = 1
    'For dProd = 1 To gPartes.rows - 1
    '    d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gPartes.TextMatrix(dProd, 2))))
    '    If d = 0 Then d = 1
    '    gPartes.TextMatrix(dProd, 4) = s2n(s2n(gPartes.TextMatrix(dProd, 4)) / d)
    'Next
End Function

Private Function VerD()
Dim dVer As String, dAjustes As Long, dComp As Long, d
    gDetalle.rows = 1
    If gAjustes.Row = 0 Then
        dAjustes = 0
    Else
        dAjustes = gAjustes.TextMatrix(gAjustes.Row, 0)
    End If
    dVer = "Select I.NUMERO AS NroRemito,I.PRODUCTO, p.DESCRIPCION,I.CANTIDAD AS VOLUMEN,0 AS CANTIDAD from ITEMREMITODIFERENCIASTOCK I INNER JOIN PRODUCTO P ON I.PRODUCTO=P.CODIGO where I.NUMERO=" & dAjustes
    LlenarGrilla gDetalle, dVer, True
    For dComp = 1 To gDetalle.rows - 1
        d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gDetalle.TextMatrix(dComp, 1))))
        If d = 0 Then d = 1
        gDetalle.TextMatrix(dComp, 4) = s2n(s2n(gDetalle.TextMatrix(dComp, 3)) / d)
    Next
    gDetalle.ColHidden(4) = True
End Function

Private Sub gAjustes_ChangeEdit()
VerD
End Sub

Private Sub gAjustes_Click()
VerD
End Sub

Private Sub gAjustes_SelChange()
VerD
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    ucXls1.ini gAjustes, "C:\AJUSTES_DE_PRODUCTOS.XLS"
End Sub

