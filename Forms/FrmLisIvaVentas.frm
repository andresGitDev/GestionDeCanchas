VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisIvaVentas 
   Caption         =   "Listado de Iva Ventas"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   Icon            =   "FrmLisIvaVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Mostrar Comprobantes A y B"
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   960
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Mostrar Exportacion"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mostrar Todo"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   930
      Left            =   8400
      TabIndex        =   4
      Top             =   60
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   1640
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4245
      Left            =   120
      TabIndex        =   8
      Top             =   2055
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
      TabIndex        =   3
      Top             =   1560
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
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
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
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   38252
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
      TabIndex        =   7
      Top             =   480
      Width           =   1335
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
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1245
      Left            =   120
      Top             =   120
      Width           =   8160
   End
End
Attribute VB_Name = "FrmLisIvaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '10/3/5

Private Const tt_Iva_Ventas_Temp = "([FECHA] [datetime] NULL , [RAZONSOCIAL] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NROCUIT] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [TIPOYNRO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NETO] [float] NULL , [NoGrav] [float] NULL , [EXENTO] [float] NULL , [IVARNI] [float] NULL , [IVARI] [float] NULL , [IVACF] [float] NULL , [IVABC] [float] NULL , [RETIVA] [float] NULL , [iibb] [float] NULL , [IMPTOTAL] [float] NULL )"


Private Const IVA_CONS_FINAL = 1
Private Const IVA_INSCRIPTO = 2
Private Const IVA_NO_INSCRIPTO = 3
Private Const IVA_MONO = 7
Private Const IVA_EXENTO = 4

Private sTmpIvaVentas As String
'

Private Sub cmdaceptar_Click()
    Dim rptV As New RptIvaVentas

    Dim STR As String
    Dim rs As New ADODB.Recordset
    'VARIABLES QUE USO PARA GRABAR EN LA TABLA TEMPORAL
    Dim Consulta As String
    Dim FechaTemp As String
    Dim IVARNI As Double
    Dim IVARI As Double
    Dim IVACF As Double
    Dim IVABC As Double
    Dim IVAEXEN As Double
    Dim EXENTO As Double
    Dim NoGrav As Double
    Dim signo As Variant
    Dim retIva As Double
    Dim IIBB As Double
    Dim Neto As Double
    Dim tdoc As String
    Dim tota As Double
    Dim razon As String
    Dim scampos As String
    Dim Aux As Double
    Dim whe As String
    Dim i As Long
    Dim Arrai
    Dim j As Long
    Dim cant As Long

    
    ucXls1.ini grilla, "c:\SubdiarioIva ", "Subdiario Ventas " & dtfechad & "  -  " & dtfechah

    If Option1.Value = True Then
        whe = "tipodoc like 'FA%%' OR TIPODOC like 'NC%%' OR TIPODOC like 'ND%%'"
    ElseIf Option2.Value = True Then
        whe = "tipodoc='FAE' OR TIPODOC='NCE' OR TIPODOC='NDE'"
    ElseIf Option3.Value = True Then
        whe = "tipodoc='FAA' OR TIPODOC='NCA' OR TIPODOC='NDA' or tipodoc='FAB' OR TIPODOC='NCB' OR TIPODOC='NDB'"
    End If
    
'    If sTmpIvaVentas = "" Then
    sTmpIvaVentas = TablaTempCrear(tt_Iva_Ventas_Temp)
'    Else
'        daTaenvironment1.Sistema.Execute "Delete From " & sTmpIvaVentas
'    End If
        
'    Consulta = "SELECT FECHA, RAZONSOCIAL, CUIT, TIPODOC, NROFACTURA, NETO, TIPOIVA, IVA, TOTAL, NoGrav , ND_xChequeRechazado" _
        & " FROM FACTURAVENTA " _
        & " WHERE FECHA " & ssBetween(dtfechad, dtfechah) _
        & " and " _
        & " (tipodoc='FAA' OR TIPODOC='FAB' OR TIPODOC='NCA' OR TIPODOC='NDA' OR TIPODOC='NCB' OR TIPODOC='NDB' OR tipodoc = 'FAE'or tipodoc = 'NCE' ) " _
        & " and ACTIVO = 1 "
    Consulta = "SELECT FECHA, RAZONSOCIAL, CUIT, TIPODOC, NROFACTURA, NETO, TIPOIVA, IVA, TOTAL, NoGrav , ND_xChequeRechazado, IIBB,  activo,PORCENTAJEIVA, iddoc,porciva " _
        & " FROM FACTURAVENTA " _
        & " WHERE FECHA " & ssBetween(dtfechad, dtfechah) _
        & " and (" _
        & "" & whe & " OR TIPODOC = 'RET')  and ACTIVO = 1 "
'        & " (tipodoc like 'FA%%' OR TIPODOC like 'NC%%' OR TIPODOC like 'ND%%' OR TIPODOC = 'RET' ) " '        & " and ACTIVO = 1 "
    
    rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
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
            Neto = rs!Neto '- (rs!neto * s2n(rs!descuento)) '- rs!NoGrav
            tota = rs!Total
            razon = rs!RAZONSOCIAL
            If rs!NroFactura = 954 Then
                MsgBox ""
            End If
            
            tdoc = Trim(rs!TIPODOC)
'            If tdoc = "RET" Then Stop
            
            Select Case rs!tipoiva
            Case IVA_INSCRIPTO
                'If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then 'si son notas de credito debo poner valores negativos
                If tdoc Like "NC*" Then  'si son notas de credito debo poner valores negativos
                    IVARI = nSinNull(rs!Iva) ' - (rs!iva * 2)
                Else
                    IVARI = nSinNull(rs!Iva)
                End If
                
            Case IVA_NO_INSCRIPTO
                'If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then 'si son notas de credito debo poner valores negativos
                If tdoc Like "NC*" Then  'si son notas de credito debo poner valores negativos
                    IVARNI = nSinNull(rs!Iva) '- (rs!iva * 2)
                Else
                    IVARNI = nSinNull(rs!Iva)
                End If
                
            Case IVA_CONS_FINAL, IVA_MONO
                'If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then 'si son notas de credito debo poner valores negativos
'                If tdoc Like "NC*" Then  'si son notas de credito debo poner valores negativos
                    'IVACF = nSinNull(rs!Iva) '- (rs!iva * 2)
                    Aux = s2n(rs!porciva, 4)
                    If Aux = 0 Then
                        IVACF = s2n((Neto / 121) * 21) 's2n(s2n(Neto * 1.21) - Neto)
                        Neto = Neto - IVACF
                    Else
                        IVACF = s2n((Neto / (100 + (Aux * 100))) * (Aux * 100)) 's2n(s2n(Neto * 1.21) - Neto)
                        Neto = Neto - IVACF 'IVAEXEN
                    End If
'                Else
'                    IVACF = nSinNull(rs!iva)
'                End If
'            Case IVA_EXENTO
'                If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then 'si son notas de credito debo poner valores negativos
'                    EXENTO = nSinNull(rs!Iva) '- (rs!iva * 2)
'                Else
'                    EXENTO = nSinNull(rs!Iva)
'                End If
            Case IVA_EXENTO
                    Aux = s2n(rs!porciva, 4)
                    If Aux = 0 Then
                        EXENTO = s2n((Neto / 121) * 21) 's2n(s2n(Neto * 1.21) - Neto)
                        Neto = Neto - EXENTO 'IVAEXEN
                    Else
                        EXENTO = s2n((Neto / (100 + (Aux * 100))) * (Aux * 100)) 's2n(s2n(Neto * 1.21) - Neto)
                        Neto = Neto - EXENTO 'IVAEXEN
                    End If
            Case Else
                'If s2n(rs!Iva) <> 0 Then ufa "Err iva listado ", "LisIvaVentas iddoc " & rs!iddoc
            End Select
            
            If rs!PorcentajeIva = 0.105 Then
                'Debug.Print rs!razonsocial & " " & rs!TIPODOC & " " & rs!nrofactura
                IVARNI = 0
                IVARI = 0
                IVACF = 0
                IVABC = nSinNull(rs!Iva)
            End If
            'If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Or Trim(rs!TIPODOC) = "NCE" Then Stop
            If tdoc Like "NC*" Then  'si son notas de credito debo poner valores negativos
                signo = "-"
            End If
            
'            If tdoc = "FAE" Or tdoc = "NCN" Then
'                EXENTO = rs!Total '= rs!neto
'                Neto = 0
'            End If
            
            If tdoc = "RET" Then ' PARA dATOS MIGRADOS SOLAMENTE
                retIva = rs!Total
                tota = rs!Total '0
            End If
            
            If rs!ND_xChequeRechazado Then
                tota = 0
            End If
            
            If Not rs!activo Then
                'porlasdudas
                razon = "    *ANULADA*"
                Neto = 0
                EXENTO = 0
                IVARNI = 0
                IVARI = 0
                IVACF = 0
                IVABC = 0
                tota = 0
                retIva = 0
                IIBB = 0
            End If
           
            Consulta = "Insert Into " & sTmpIvaVentas & " (FECHA, RAZONSOCIAL, NROCUIT, TIPOYNRO, NETO, EXENTO, " & _
                    " IVARNI, IVARI, IVACF, IVABC, IMPTOTAL, NoGrav, RetIva, IIBB) " & _
                    "Values ( " & ssFecha(rs!Fecha) & " , '" & razon & "', '" & rs!CUIT & "', '" & _
                    (rs!TIPODOC & rs!NroFactura) & "', " & signo & ssNum(Neto) & ", " & signo & ssNum(EXENTO) & ", " & _
                    signo & ssNum(IVARNI) & ", " & signo & ssNum(IVARI) & ", " & signo & ssNum(IVACF) & ", " _
                    & signo & ssNum(IVABC) & ", " _
                    & signo & ssNum(tota) & ", " _
                    & signo & ssNum(rs!NoGrav) & ", " & ssNum(retIva) & ", '" & ssNum(IIBB) & "' )"
            DataEnvironment1.Sistema.Execute Consulta
            rs.MoveNext
        Wend
    End If
    
    rs.Close

' rs.Open "select ret.*, clientes.* from RecibosRetenciones  as ret left join clientes on ret.cliente = clientes.codigo where fecha " & ssBetween(dtfechad, dtfechah) & " and idCuentasParam = '" & ID_Cuenta_R_RET_IVA_RG3125 & "' "
'" from RecibosRetenciones as ret left join clientes on ret.cliente = clientes.codigo " &
    
    
    'Retenciones, hechas con el sistema VB
    rs.Open "select ret.*, c.descripcion, c.cuit,  recibos.* " & _
           " FROM Clientes as c INNER JOIN (Recibos INNER JOIN RecibosRetenciones as ret ON Recibos.idDoc = ret.iddoc) ON c.codigo = Recibos.Cliente " & _
            " where ret.fecha " & ssBetween(dtfechad, dtfechah) & " and idCuentasParam = '" & ID_Cuenta_R_RET_IVA_RG3125 & "' " _
            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            '       recibos.fecha ???
    While Not rs.EOF
        Consulta = "Insert Into " & sTmpIvaVentas & _
                        " (FECHA, RAZONSOCIAL, NROCUIT, TIPOYNRO, NETO, EXENTO, " & _
                        "IVARNI, IVARI, IVACF, IVABC, IMPTOTAL, NoGrav, RetIva, iibb) " & _
                    "Values ( " & _
                        ssFecha(rs!Fecha) & " , '" & ssStr(rs!DESCRIPCION) & "', '" & ssStr(rs!CUIT) & "', 'Ret " & _
                        (rs!numero) & "', 0 , 0, 0, 0, 0 ,0, '" & ssNum(-rs!Importe) & "', 0, '" & ssNum(-rs!Importe) & "' ,0 )"
        'DataEnvironment1.Sistema.Execute Consulta
        rs.MoveNext
    Wend
    
    Set rs = Nothing
                                                                                
                                                                            
    'str = "select * from IVA_VENTAS_TEMP order by fecha,tipoynro"
    scampos = " [FECHA] as [ Fecha], [RAZONSOCIAL] as [Razon Social], [NROCUIT]  as [CUIT]," & _
        " [TIPOYNRO] as [Documento], [NETO]  , [NoGrav] as [No grav], [EXENTO] as [Exento ]," & _
        " [IVARNI] as [IVA RNI], [IVARI] as [IVA RI 21%], [IVACF] as [IVA CF], [IVABC] as [IVA RI 10.5%], " & _
        " [iibb] as [IIBB ], [IMPTOTAL]  as [Total]"
    STR = "select " & scampos & " from " & sTmpIvaVentas & " order by fecha,tipoynro"
    
    LlenarGrilla grilla, STR, True
    sumarizo grilla, Array(4, 5, 6, 7, 8, 9, 10, 11, 12)
    
    grilla.ColWidth(2) = 1300
    grilla.ColWidth(3) = 1200
    grilla.ColWidth(4) = 1300
    grilla.ColWidth(5) = 1300
    grilla.ColWidth(6) = 1300
    grilla.ColWidth(7) = 1300
    grilla.ColWidth(8) = 1300
    grilla.ColWidth(9) = 1300
    grilla.ColWidth(10) = 1300
    grilla.ColWidth(11) = 1300
    grilla.ColWidth(12) = 1300
    
    
    i = 1
    cant = 36
    ReDim Arrai(9)
    While i < grilla.rows
        If i = cant Then
            grilla.AddItem "" & Chr(9) & "SUBTOTAL DE TRANSPORTE" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Arrai(0) & Chr(9) & Arrai(1) & Chr(9) & Arrai(2) & Chr(9) & _
                        Arrai(3) & Chr(9) & Arrai(4) & Chr(9) & Arrai(5) & Chr(9) & Arrai(6) & Chr(9) & Arrai(7) & Chr(9) & Arrai(8) & Chr(9) & Arrai(9), i
            grilla.AddItem "" & Chr(9) & "SUBTOTAL DE TRANSPORTE" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Arrai(0) & Chr(9) & Arrai(1) & Chr(9) & Arrai(2) & Chr(9) & _
                        Arrai(3) & Chr(9) & Arrai(4) & Chr(9) & Arrai(5) & Chr(9) & Arrai(6) & Chr(9) & Arrai(7) & Chr(9) & Arrai(8) & Chr(9) & Arrai(9), i + 1
            cant = cant + 38
            i = i + 2
        End If
        j = 0
        While j < 9
            Arrai(j) = s2n(Arrai(j)) + s2n(grilla.TextMatrix(i, 4 + j))
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    i = 1
    While i < grilla.rows
        j = 4
        While j < grilla.cols
            grilla.TextMatrix(i, j) = s2n(grilla.TextMatrix(i, j), 2, True)
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    PrintG grilla, pHorizontal, "SUBDIARIO DE VENTA", Date, "SUBDIARIO DE VENTA DE " & dtfechad & " AL " & dtfechah, 5
    
    i = 1
    While i < grilla.rows
        If Trim(grilla.TextMatrix(i, 1)) = "SUBTOTAL DE TRANSPORTE" Then
            grilla.RemoveItem i
            i = i - 1
        End If
        i = i + 1
    Wend
    
    'hago esto porque el listado esta preparado para los campos originales... sorry
'    STR = "select  * from " & sTmpIvaVentas & " order by fecha,tipoynro"
    
'    rptV.Data.Connection = DataEnvironment1.Sistema
'    rptV.Data.Source = STR
'    rptV.lblfecha = Date
'    rptV.lblTitulo = "Subdiario Ventas"
    
    'DETALLE DEL REPORTE
'    With rptV
'        .fieFecha.DataField = "FECHA"
'        .fieRazonSocial.DataField = "RAZONSOCIAL"
'        .fieCUIT.DataField = "NROCUIT"
'        .fieTipoYNro.DataField = "TIPOYNRO"
'        .fieNeto.DataField = "NETO"
'        .fieExento.DataField = "EXENTO"
'        .fieIVARNI.DataField = "IVARNI"
'        .fieIVARI.DataField = "IVARI"
'        .fieIVACF.DataField = "IVACF"
'        .fieIVABC.DataField = "IVABC"
'        .fieImpTotal.DataField = "IMPTOTAL"
'        .fieNoGrav.DataField = "NoGrav"
'        .fieRetIva.DataField = "retiva"
'        .fieRetIva.Visible = False
'        .Label18.Visible = False
'        .Field12.Visible = False
'        .fieIIBB.DataField = "iibb"
    
    'COLA DEL REPORTE
        
'        .fieTotalNeto.DataField = "NETO"
'        .fieTotalExento.DataField = "EXENTO"
'        .fieTotalIVARNI.DataField = "IVARNI"
'        .fieTotalIVARI.DataField = "IVARI"
'        .fieTotalIVACF.DataField = "IVACF"
'        .fieTotalIVABC.DataField = "IVABC"
'        .fieTotalImporte.DataField = "IMPTOTAL"
'        .fieTotalNoGrav.DataField = "NoGrav"
'        .fieTotalIIBB.DataField = "iibb"
'        .fieTotalRetIva.DataField = "RetIva"
'        .fieTotalRetIva.Visible = False
'        .Field12.Visible = False
'    End With
    
'    Dim bb As Boolean
'    bb = confirma("imprime fecha de emision")
'    rptV.Label1.Visible = bb
'    rptV.lblfecha.Visible = bb
'    rptV.lblfecha = " entre " & dtfechad & " y " & dtfechah
'    rptV.lblfecha.Width = 5000
    
'    rptV.Show
'    Set rptV = Nothing
End Sub

Private Sub cmdcancelar_Click()
    dtfechad = Date - Day(Date) + 1
    dtfechah = Date
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub dtfechad_LostFocus()
    dtfechah = CDate(DateSerial(Year(dtfechad), Month(dtfechad) + 1, 0))
End Sub

Private Sub Form_Load()
    dtfechad = Date - Day(Date) + 1
    dtfechah = Date
'    sTmpIvaVentas = TablaTempCrear(tt_Iva_Ventas_Temp)
    Form_Resize
End Sub

''Private Sub Form_Unload(cancel As Integer)
''    TablaTempBorrar sTmpIvaVentas
''    sTmpIvaVentas = ""
''End Sub

'10/3/5 tabla temp
Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
'    ucXls1.ini grilla, "c:\", "Subdiario Ventas "
    ucXls1.aTitulo = "Subdiario Ventas " & dtfechad & "  -  " & dtfechah
End Sub

Private Sub sumarizo(grilla, a)
    Dim i As Long
    If grilla = "" Then Exit Sub
    With grilla
        .SubtotalPosition = flexSTBelow
        For i = 0 To UBound(a):        .subtotal flexSTSum, -1, a(i), , , , True, , , True: Next
'        .TextMatrix(.rows - 1, 0) = " Totales"
    End With
End Sub

