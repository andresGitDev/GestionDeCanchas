VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPartesInformeMovi 
   Caption         =   "Informe de Movimientos de Partes de Produccion"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12660
   Icon            =   "frmPartesInformeMovi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   12660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   795
      Left            =   5400
      Picture         =   "frmPartesInformeMovi.frx":08CA
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
      _extentx        =   1376
      _extenty        =   1402
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4350
      Left            =   105
      TabIndex        =   4
      Top             =   6855
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
      Left            =   90
      TabIndex        =   3
      Top             =   1980
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
      Format          =   58261505
      CurrentDate     =   39594
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   120
      Picture         =   "frmPartesInformeMovi.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
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
      Format          =   58261505
      CurrentDate     =   39594
   End
   Begin Gestion.ucCoDe uProdD 
      Height          =   360
      Left            =   150
      TabIndex        =   7
      Top             =   975
      Width           =   9075
      _extentx        =   16007
      _extenty        =   635
      codigowidth     =   1000
   End
   Begin Gestion.ucCoDe uProdH 
      Height          =   390
      Left            =   150
      TabIndex        =   8
      Top             =   1455
      Width           =   9090
      _extentx        =   16034
      _extenty        =   688
      codigowidth     =   1000
   End
End
Attribute VB_Name = "frmPartesInformeMovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const tt_Partes_Movi_Temp = "( [NROPARTE] [VARCHAR] (4000) NULL,[FECHA] [datetime] NULL, [CODPRO] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL  , [DESCRIPCION] [varchar] (4000) NULL , [PRODUCIDO] [float] NULL , [CANTIDAD] [float] NULL, [SALDO] [float] NULL )"

Private Sub cmdVer_Click()
    Ver
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    iFechas
    Ver
    ucXls1.ini gPartes, "C:\PARTEPRODUCCION_MOVIMIENTO.XLS"
    uProd_ini
End Sub

Private Sub uProd_ini()
Dim sqlbuscar As String, sqldesc As String

sqldesc = "select descripcion from producto where codigo = '###' "
sqlbuscar = "select p.codigo as [ Codigo                 ],  p.descripcion as [ Descripcion                                                 ],m.abreviatura as [  Medida  ] from producto as p inner join unidadesmedida as m on p.umedida=m.umcodigo where p.activo = 1 and formula=1 order by p.codigo "

uProdD.ini sqldesc, sqlbuscar, True
uProdH.ini sqldesc, sqlbuscar, True

Dim d
d = obtenerDeSQL("select max(codigo),min(codigo) from producto where activo=1")
uProdD.codigo = d(1)
uProdH.codigo = d(0)
End Sub

Sub iFechas()
    dtDesde = "01/" & Month(Date) & "/" & Year(Date)
    dtHasta = Date
End Sub

Private Function Ver()
Dim cver As String, dProd As Long
Dim rsPro As New ADODB.Recordset, i As Long, x As Long
Dim rsParte As New ADODB.Recordset, d
Dim tMovi As String, mSaldo As Double, ultProd As String
tMovi = TablaTempCrear(tt_Partes_Movi_Temp)
If uProdD.codigo = 0 And uProdH.codigo = 0 Then Exit Function
With rsPro
    .Open "select * from producto where activo=1 and codigo >=" & sstexto(uProdD.codigo) & " and codigo<=" & sstexto(uProdH.codigo) & " order by codigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            mSaldo = 0
            rsParte.Open "Select PA.* from PartesProduccion PA where PA.activo=1 and (pa.fecha>= " & ssFecha(dtDesde) & " and pa.fecha<=" & ssFecha(dtHasta) & " )  and PA.producido=" & sstexto(!codigo) & " order by pa.fecha,pa.nro", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            If rsParte.EOF And rsParte.BOF Then
            Else
                rsParte.MoveFirst
                For x = 0 To rsParte.RecordCount - 1
                    mSaldo = mSaldo + s2n(rsParte!cantidad)
                    insertTabla tMovi, rsParte!Nro, rsParte!fecha, !codigo, !DESCRIPCION, rsParte!cantidad, rsParte!cantidad, s2n(mSaldo)
                    rsParte.MoveNext
                Next
            End If
            Set rsParte = Nothing
            
            rsParte.Open "Select PA.nro,PA.fecha,F.* from PartesProduccion PA inner join formulasdetalle F ON pa.nro=f.codigo_parte where f.activo=1 and PA.activo=1 and (pa.fecha>= " & ssFecha(dtDesde) & " and pa.fecha<=" & ssFecha(dtHasta) & " )  and F.COMPONENTE=" & sstexto(!codigo) & " order by pa.fecha,pa.nro", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            If rsParte.EOF And rsParte.BOF Then
            Else
                rsParte.MoveFirst
                For x = 0 To rsParte.RecordCount - 1
                    mSaldo = mSaldo - s2n(rsParte!cantidad)
                    insertTabla tMovi, rsParte!Nro, rsParte!fecha, !codigo, !DESCRIPCION, -rsParte!cantidad, -rsParte!cantidad, s2n(mSaldo)
                    rsParte.MoveNext
                Next
            End If
            Set rsParte = Nothing
            .MoveNext
            
        Next
    End If
End With


    gPartes.rows = 1
    cver = "Select NROPARTE as [NRO PARTE],FECHA,CODPRO AS CODIGO,DESCRIPCION,PRODUCIDO,CANTIDAD,SALDO from " & tMovi & " "
    LlenarGrilla gPartes, cver, True
    gDetalle.rows = 1
    ultProd = ""
    For dProd = 1 To gPartes.rows - 1
        d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gPartes.TextMatrix(dProd, 2))))
        If d = 0 Then d = 1
        gPartes.TextMatrix(dProd, 4) = s2n(s2n(gPartes.TextMatrix(dProd, 5)) / d)
        If ultProd <> gPartes.TextMatrix(dProd, 2) Then
            gPartes.cell(flexcpFontBold, dProd, 0, dProd, 6) = True
            ultProd = gPartes.TextMatrix(dProd, 2)
        End If
    Next

    With gPartes
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .ColWidth(2) = 1000
        .ColWidth(3) = 3500
        .ColWidth(4) = 1200
        .ColWidth(5) = 1000
        .ColWidth(6) = 1200
    End With
    
End Function

Private Function insertTabla(iTabla As String, iNROPARTE As String, iFECHA As Date, iCODPRO As String, iDESCRIPCION As String, iPRODUCIDO As Double, iCANTIDAD As Double, iSALDO As Double)
Dim cad As String
If Trim(iTabla) = "" Then Exit Function

cad = " insert into " & iTabla & " ( NROPARTE,FECHA,CODPRO,DESCRIPCION,PRODUCIDO,CANTIDAD,SALDO ) VALUES (" & _
        " " & sstexto(iNROPARTE) & "," & ssFecha(iFECHA) & "," & sstexto(iCODPRO) & "," & sstexto(iDESCRIPCION) & "," & x2s(iPRODUCIDO) & "," & x2s(iCANTIDAD) & "," & x2s(iSALDO) & ")"

DataEnvironment1.Sistema.Execute cad
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
