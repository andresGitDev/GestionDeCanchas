VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPartesInforme 
   Caption         =   "Informe de Partes de Produccion"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmPartesInforme.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   795
      Left            =   5400
      Picture         =   "frmPartesInforme.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   855
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   795
      Left            =   4350
      TabIndex        =   5
      Top             =   90
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1402
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4350
      Left            =   75
      TabIndex        =   4
      Top             =   5835
      Width           =   12495
      _cx             =   22040
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
   Begin VSFlex7LCtl.VSFlexGrid gPartes 
      Height          =   4770
      Left            =   75
      TabIndex        =   3
      Top             =   960
      Width           =   12510
      _cx             =   22066
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
      Left            =   1185
      TabIndex        =   1
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   39594
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   135
      Picture         =   "frmPartesInforme.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   75
      Width           =   765
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   2820
      TabIndex        =   2
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58458113
      CurrentDate     =   39594
   End
End
Attribute VB_Name = "frmPartesInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdVer_Click()
    Ver
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    iFechas
    Ver
    ucXls1.ini gPartes, "C:\PARTEPRODUCCION.XLS"
End Sub

Sub iFechas()
    dtDesde = "01/" & Month(Date) & "/" & Year(Date)
    dtHasta = Date
End Sub

Private Function Ver()
Dim cver As String, dProd As Long
    gPartes.rows = 1
    cver = "Select PA.NRO AS NroParte, PA.Fecha, PA.Producido as [PRODUCTO PRODUCIDO],P.DESCRIPCION,PA.CANTIDAD from PartesProduccion PA INNER JOIN PRODUCTO P ON PA.PRODUCIDO=P.CODIGO where PA.activo=1 and (pa.fecha>= " & ssFecha(dtDesde) & " and pa.fecha<=" & ssFecha(dtHasta) & " )"
    LlenarGrilla gPartes, cver, True
    gDetalle.rows = 1
    For dProd = 1 To gPartes.rows - 1
        d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gPartes.TextMatrix(dProd, 2))))
        If d = 0 Then d = 1
        gPartes.TextMatrix(dProd, 4) = s2n(s2n(gPartes.TextMatrix(dProd, 4)) / d)
    Next
End Function

Private Function VerD()
Dim dVer As String, dParte As Long, dComp As Long, d
    gDetalle.rows = 1
    If gPartes.Row = 0 Then
        dParte = 0
    Else
        dParte = gPartes.TextMatrix(gPartes.Row, 0)
    End If
    dVer = "Select F.COMPONENTE, P.DESCRIPCION,F.CANTIDAD AS VOLUMEN,0 AS CANTIDAD from FormulasDetalle F INNER JOIN PRODUCTO P ON F.COMPONENTE=P.CODIGO where F.activo=1 AND CODIGO_PARTE=" & dParte
    LlenarGrilla gDetalle, dVer, True
    For dComp = 1 To gDetalle.rows - 1
        d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gDetalle.TextMatrix(dComp, 0))))
        If d = 0 Then d = 1
        gDetalle.TextMatrix(dComp, 3) = s2n(s2n(gDetalle.TextMatrix(dComp, 2)) / d)
    Next
    
End Function

Private Sub gPartes_ChangeEdit()
VerD
End Sub

Private Sub gPartes_Click()
VerD
End Sub

Private Sub gPartes_SelChange()
VerD
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    ucXls1.ini gPartes, "C:\PARTEPRODUCCION.XLS"
End Sub
