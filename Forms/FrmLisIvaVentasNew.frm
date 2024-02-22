VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLisIvaVentasNew 
   Caption         =   "Listado de Iva Ventas"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Mostrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1620
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mostrar Todo"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   300
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Mostrar Exportacion"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   660
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Mostrar Comprobantes A y B"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   1020
      Width           =   2415
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   930
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1640
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4245
      Left            =   120
      TabIndex        =   4
      Top             =   2115
      Width           =   9075
      _cx             =   16007
      _cy             =   7488
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
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   315
      Left            =   1545
      TabIndex        =   8
      Top             =   525
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   394133505
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   394133505
      CurrentDate     =   38252
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1245
      Left            =   120
      Top             =   180
      Width           =   8160
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   540
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLisIvaVentasNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '10/3/5

Private Const tt_Iva_Ventas_Temp = "([FECHA] [datetime] NULL , [RAZONSOCIAL] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NROCUIT] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [TIPOYNRO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NETO21] [float] NULL,[NETO10] [float] NULL , [NoGrav] [float] NULL , [EXENTO] [float] NULL , [IVARNI] [float] NULL , [IVARI] [float] NULL , [IVACF] [float] NULL , [IVABC] [float] NULL , [RETIVA] [float] NULL , [iibb] [float] NULL , [IMPTOTAL] [float] NULL )"


Private Const IVA_CONS_FINAL = 1
Private Const IVA_INSCRIPTO = 2
Private Const IVA_NO_INSCRIPTO = 3
Private Const IVA_MONO = 7
Private Const IVA_EXENTO = 4
Private Const IVA_EXENTOLEY = 8
Private Const IVA_EXTERIOR = 9


Private sTmpIvaVentas As String
'

Private Sub cmdAceptar_Click()
Dim Opp As Long
If Option1 Then
    Opp = 1
ElseIf Option2 Then
    Opp = 2
ElseIf Option3 Then
    Opp = 3
Else
    Opp = 1
End If

CargoGV GRILLA, dtfechad, dtfechah, Opp

With GRILLA
otraves:
    For Opp = 1 To .rows - 1
        If CORTO(.TextMatrix(Opp, 3), 0, Len(.TextMatrix(Opp, 3)) - 3) = "Ret" Then
            .RemoveItem Opp
            GoTo otraves
        End If
    Next
    .RemoveItem .rows - 1
    .ColWidth(12) = 0
    sumarizo GRILLA, Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
End With

PrintG GRILLA, pHorizontal, "SUBDIARIO DE VENTA", Date, "SUBDIARIO DE VENTA DE " & dtfechad & " AL " & dtfechah, 5


End Sub

Public Function CargoGV(ByRef gg As Object, ByVal FechaD As Date, ByVal FechaH As Date, Optional Opcion As Long = 1)
    Dim rptV As New RptIvaVentas

    Dim str As String, rs As New ADODB.Recordset, Consulta As String, FechaTemp As String
    Dim IVARNI As Double, IVARI As Double, IVACF As Double, IVABC As Double, IVAEXEN As Double, EXENTO As Double, NoGrav As Double, signo As Variant, retIva As Double, IIBB As Double
    Dim Neto As Double, Neto10 As Double, tdoc As String, tota As Double, razon As String
    Dim scampos As String, PIVA As Double, sWhere As String, i As Long
    Dim sLetra1 As String, sLetra2 As String, sLetra3 As String
    
    Dim IVA_INSCRIPTO_GE As Integer, doc As New FacturaElectronica
    IVA_INSCRIPTO_GE = doc.CodigoIvaGranEmpresa()
    
    ucXls1.ini gg, "c:\SubdiarioIva ", "Subdiario Ventas " & FechaD & "  -  " & FechaH
    
    Select Case Opcion
        Case 1:
            sWhere = " (tipodoc like 'D%' OR tipodoc like 'C%' OR  tipodoc like 'F%' OR TIPODOC like 'N%' ) "
        Case 2:
            sWhere = " (tipodoc like '%E' ) "
        Case 3:
            sWhere = " (tipodoc like '%A' or tipodoc like '%B'  ) "
    End Select
    sTmpIvaVentas = TablaTempCrear(tt_Iva_Ventas_Temp)
    Consulta = "SELECT FECHA, RAZONSOCIAL, CUIT, TIPODOC, NROFACTURA, NETO, TIPOIVA, IVA, TOTAL, NoGrav , ND_xChequeRechazado, IIBB,  activo,PORCENTAJEIVA, iddoc,porciva,variasfac,codigo,puntoventa " _
        & " FROM FACTURAVENTA " _
        & " WHERE FECHA " & ssBetween(FechaD, FechaH) _
        & " and " & sWhere
    
    rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF And Not rs.BOF Then
        rs.MoveFirst
        While Not rs.EOF
        
            
            signo = ""
            EXENTO = 0
            IVARI = 0
            IVARNI = 0
            IVACF = 0
            IVABC = 0
            NoGrav = 0
            retIva = 0
            IIBB = rs!IIBB
            Neto = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=21"), 2), "standard")
            Neto10 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=10.5"), 2), "standard")
            If Neto = 0 And Neto10 = 0 Then NoGrav = rs!Neto
            tota = rs!Total
            razon = rs!RAZONSOCIAL
            
            tdoc = Trim(rs!TIPODOC)
            sLetra1 = CORTO(tdoc, 0, 2)
            sLetra2 = CORTO(tdoc, 1, 1)
            sLetra3 = CORTO(tdoc, 2, 0)
            Select Case sLetra3
            Case "A"
                IVARI = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=21"), 2) * s2n(0.21, 4), "standard")
                IVABC = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=10.5"), 2) * s2n(0.105, 4), "standard")
                If IVARI = 0 And IVABC = 0 Then IVARI = nSinNull(rs!Iva)
            Case "B", "C"
                IVARI = 0
                IVABC = 0
                
                IVARI = Format$(s2n(obtenerDeSQL("select (sum(preciototal) * 0.173553719008264) as i21 from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=21"), 2), "standard")
                IVABC = Format$(s2n(obtenerDeSQL("select (sum(preciototal) * 0.095022624434389) as i10 from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=10.5"), 2), "standard")
                
                Neto = Neto - IVARI
                Neto10 = Neto10 - IVABC
                
            Case "E"
                NoGrav = rs!Total
                Neto = 0
                Neto10 = 0
                EXENTO = 0
                IVARNI = 0
                IVARI = 0
                IVACF = 0
                IVABC = 0
            End Select
            
            If tdoc Like "CE*" Or tdoc Like "NC*" Then signo = "-"
            
            If rs!ND_xChequeRechazado Then
                Neto = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=21"), 2), "standard")
                Neto10 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=10.5"), 2), "standard")
                IVARI = Format$(s2n(Neto, 2) * s2n(0.21), "standard")
                IVABC = Format$(s2n(Neto10, 2) * s2n(0.105), "standard")
                NoGrav = rs!Neto
            End If
            
           
            Consulta = "Insert Into " & sTmpIvaVentas & " (FECHA, RAZONSOCIAL, NROCUIT, TIPOYNRO, NETO21,NETO10, EXENTO, " & _
                    " IVARNI, IVARI, IVACF, IVABC, IMPTOTAL, NoGrav, RetIva, IIBB) " & _
                    "Values ( " & ssFecha(rs!Fecha) & " , '" & razon & "', '" & rs!CUIT & "', '" & _
                    (Trim(rs!TIPODOC) & "  " & rs!NroFactura) & "', " & signo & ssNum(Neto) & "," & signo & ssNum(Neto10) & ", " & signo & ssNum(EXENTO) & ", " & _
                    signo & ssNum(IVARNI) & ", " & signo & ssNum(IVARI) & ", " & signo & ssNum(IVACF) & ", " _
                    & signo & ssNum(IVABC) & ", " _
                    & signo & ssNum(tota) & ", " _
                    & signo & ssNum(NoGrav) & ", " & ssNum(retIva) & ", '" & ssNum(IIBB) & "' )"
            DataEnvironment1.Sistema.Execute Consulta
            rs.MoveNext
        Wend
    End If
    
    rs.Close

    'Retenciones, hechas con el sistema VB
    rs.Open "select ret.*, c.descripcion, c.cuit,  recibos.* " & _
           " FROM Clientes as c INNER JOIN (Recibos INNER JOIN RecibosRetenciones as ret ON Recibos.idDoc = ret.iddoc) ON c.codigo = Recibos.Cliente " & _
            " where ret.fecha " & ssBetween(FechaD, FechaH) & " and idCuentasParam = '" & ID_Cuenta_R_RET_IVA_RG3125 & "' " _
            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        Consulta = "Insert Into " & sTmpIvaVentas & _
                        " (FECHA, RAZONSOCIAL, NROCUIT, TIPOYNRO, NETO21,NETO10, EXENTO, " & _
                        "IVARNI, IVARI, IVACF, IVABC, IMPTOTAL, NoGrav, RetIva, iibb) " & _
                    "Values ( " & _
                        ssFecha(rs!Fecha) & " , '" & ssStr(rs!DESCRIPCION) & "', '" & ssStr(rs!CUIT) & "', 'Ret " & _
                        (rs!numero) & "', 0 ,0, 0, 0, 0, 0 ,0, '" & ssNum(-rs!Importe) & "' , 0, '" & ssNum(-rs!Importe) & "' ,0 )"
        DataEnvironment1.Sistema.Execute Consulta
        rs.MoveNext
    Wend
    
    Set rs = Nothing

    scampos = " [FECHA] as [ Fecha], [RAZONSOCIAL] as [Razon Social], [NROCUIT]  as [CUIT]," & _
        " [TIPOYNRO] as [Documento], [NETO21] , [NETO10] , [NoGrav] as [No grav], [EXENTO] as [Exento ]," & _
        " [IVARNI] as [IVA RNI], [IVARI] as [IVA 21%], [IVACF] as [IVA CF], [IVABC] as [IVA 10.5%],[RETIVA] AS [RET IVA], " & _
        " [IMPTOTAL]  as [Total]"
    str = "select " & scampos & " from " & sTmpIvaVentas & " order by fecha,tipoynro"
    
    LlenarGrilla gg, str, True
    sumarizo gg, Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13)
    
    If gg.rows > 1 Then
        gg.ColWidth(2) = 1300
        gg.ColWidth(3) = 1200
        gg.ColWidth(4) = 1300
        gg.ColWidth(5) = 1300
        gg.ColWidth(6) = 1300
        gg.ColWidth(7) = 1300
        gg.ColWidth(8) = 1300
        gg.ColWidth(9) = 1300
        gg.ColWidth(10) = 1300
        gg.ColWidth(11) = 1300
        gg.ColWidth(12) = 1300
        gg.ColWidth(13) = 1300
    End If
    

End Function

Private Sub cmdCancelar_Click()
    dtfechad = Date - Day(Date) + 1
    dtfechah = Date
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtfechad_LostFocus()
    dtfechah = CDate(DateSerial(Year(dtfechad), Month(dtfechad) + 1, 0))
End Sub

Private Sub Form_Load()
    dtfechad = Date - Day(Date) + 1
    dtfechah = Date
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar GRILLA, Me, anclarLadosTodos
End Sub

Private Sub ucXls1_Clic(cancel As Boolean)
    ucXls1.aTitulo = "Subdiario Ventas " & dtfechad & "  -  " & dtfechah
End Sub

Private Sub sumarizo(GRILLA, a)
    Dim i As Long
    If GRILLA = "" Then Exit Sub
    With GRILLA
        .SubtotalPosition = flexSTBelow
        For i = 0 To UBound(a):        .subtotal flexSTSum, -1, a(i), , , , True, , , True: Next
    End With
End Sub
