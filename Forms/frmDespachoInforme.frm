VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDespachoInforme 
   Caption         =   "Listado de Despachos"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmDespachoInforme.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   165
      Picture         =   "frmDespachoInforme.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   765
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   795
      Left            =   4380
      TabIndex        =   0
      Top             =   75
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1402
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4350
      Left            =   135
      TabIndex        =   1
      Top             =   5835
      Width           =   10305
      _cx             =   18177
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
   Begin VSFlex7LCtl.VSFlexGrid gDespacho 
      Height          =   4770
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   10320
      _cx             =   18203
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
      Left            =   1215
      TabIndex        =   3
      Top             =   525
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94502913
      CurrentDate     =   39594
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   2850
      TabIndex        =   5
      Top             =   525
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94502913
      CurrentDate     =   39594
   End
End
Attribute VB_Name = "frmDespachoInforme"
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
    ucXls1.ini gDespacho, "C:\Despachos.XLS"
End Sub

Sub iFechas()
    dtDesde = "01/" & Month(Date) & "/" & Year(Date)
    dtHasta = Date
End Sub

Private Function Ver()
Dim cver As String
    gDespacho.rows = 1
    cver = "Select D.codigoDespacho AS Cod,D.numeroDespacho as Despacho, D.Fecha, D.Observacion1 [Comprobantes],d.observacion2 [Observacion] from Despacho D where D.activo=1 and (D.fecha>= " & ssFecha(dtDesde) & " and D.fecha<=" & ssFecha(dtHasta) & " )"
    LlenarGrilla gDespacho, cver, True
    gDetalle.rows = 1
End Function

Private Function VerD()
Dim dVer As String, dDesp As String, i As Long
    gDetalle.rows = 1
    If gDespacho.Row = 0 Then
        dDesp = 0
    Else
        dDesp = gDespacho.TextMatrix(gDespacho.Row, 1)
    End If
    dVer = "Select D.DOCUMENTO, D.NUMERO ,'' as Cliente from DespachoDetalle D where NumeroDespacho=" & sstexto(dDesp)
    LlenarGrilla gDetalle, dVer, True
    
    With gDetalle
        For i = 1 To .rows - 1
            If InStr(.TextMatrix(i, 0), "Factura") Then
                .TextMatrix(i, 2) = obtenerDeSQL("select razonsocial from facturaventa where nrofactura=" & .TextMatrix(i, 1))
            ElseIf InStr(.TextMatrix(i, 0), "Remito") Then
                .TextMatrix(i, 2) = obtenerDeSQL("select c.descripcion from remitoventa r inner join clientes c on r.cliente=c.codigo where r.numero=" & .TextMatrix(i, 1))
            ElseIf InStr(.TextMatrix(i, 0), "Pedido") Then
                .TextMatrix(i, 2) = obtenerDeSQL("select c.descripcion from pedidos_clientes p inner join clientes c on p.cliente=c.codigo where p.numero=" & .TextMatrix(i, 1))
            End If
        Next
    End With
End Function

Private Sub gDespacho_ChangeEdit()
VerD
End Sub

Private Sub gDespacho_Click()
VerD
End Sub

Private Sub gDespacho_SelChange()
VerD
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    ucXls1.ini gDespacho, "C:\Despachos.XLS"
End Sub

