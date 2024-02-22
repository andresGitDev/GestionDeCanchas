VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListadoChequesTerceros2 
   Caption         =   "Historico de cheques de tercero"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ver "
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
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "8"
      Top             =   1650
      Width           =   1185
   End
   Begin VB.Frame fraEstados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1875
      Left            =   270
      TabIndex        =   8
      Tag             =   "0"
      Top             =   165
      Width           =   2985
      Begin VB.OptionButton optIngresados 
         Caption         =   "Cheques Ingresados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optAcreditarse 
         Caption         =   "Cheques Por Acreditarse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Tag             =   "1"
         Top             =   210
         Width           =   2655
      End
      Begin VB.OptionButton optTransferidos 
         Caption         =   "Cheques Transferidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Tag             =   "2"
         Top             =   510
         Width           =   2655
      End
      Begin VB.OptionButton optCartera 
         Caption         =   "Cheques En Cartera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Tag             =   "3"
         Top             =   825
         Width           =   2655
      End
      Begin VB.OptionButton optRechazados 
         Caption         =   "Cheques Rechazados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Tag             =   "4"
         Top             =   1125
         Width           =   2655
      End
   End
   Begin VB.Frame fraFechas 
      Height          =   915
      Left            =   3360
      TabIndex        =   1
      Top             =   165
      Width           =   4575
      Begin VB.OptionButton optFeChe 
         Caption         =   "Fecha Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   255
         TabIndex        =   3
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton optFeIng 
         Caption         =   "Fecha Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   255
         TabIndex        =   2
         Top             =   570
         Value           =   -1  'True
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtFDesde 
         Height          =   300
         Left            =   2865
         TabIndex        =   4
         Tag             =   "6"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtFHasta 
         Height          =   300
         Left            =   2865
         TabIndex        =   5
         Tag             =   "7"
         Top             =   510
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62324737
         CurrentDate     =   39147
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   7
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   6
         Top             =   225
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1650
      Width           =   1185
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4080
      Left            =   135
      TabIndex        =   14
      Top             =   2235
      Width           =   9990
      _cx             =   17621
      _cy             =   7197
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
      FormatString    =   $"FrmListadoChequesTerceros2.frx":0000
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
   Begin Gestion.ucCoDe uCliente 
      Height          =   345
      Left            =   4125
      TabIndex        =   15
      Top             =   1185
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   609
      CodigoWidth     =   1000
   End
   Begin Gestion.ucXls uXls 
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   1650
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
   End
   Begin VB.Label lblcliente 
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   1230
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   2025
      Left            =   120
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "FrmListadoChequesTerceros2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 16/9/4
' 14/7/6 lo Rehice.
    
    
Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
   
    Dim rs As New ADODB.Recordset, s As String
    Dim i As Long, sDato As String, sDato2
    'pal corte
    Dim Mes As Long, Anio As Long, rsMes  As Long, rsAnio As Long
    Dim Total As Double, TotalTotal As Double
    
    'pal where
    Dim campoFecha As String, queRangoFecha As String, queCliente As String, queEstado As String
    Dim fec As String
    
    relojito
    InicioGrilla
    
    campoFecha = IIf(optFeIng, "cc.FECHA_INGR", "cc.FECHA")
    queRangoFecha = " and ( " & campoFecha & ssBetween(dtFDesde, dtFHasta) & " ) "
    queCliente = IIf(uCliente.codigo > 0, "and cc.cliente = " & uCliente.codigo, "")
    fec = " and (bb.fecha " & ssBetween(dtFDesde, dtFHasta) & " ) "
    
    queEstado = ""
    
    '****************** parte a modificar **********************
    
    'If optCartera = True Then queEstado = " and estado = 'C' "
    'If optAcreditarse = True Then queEstado = " and estado = 'D' "
    'If optRechazados = True Then queEstado = " and estado = 'R' "
    'If optTransferidos Then queEstado = " and estado = 'T' "
    
    If optCartera = True Then 'muestro los que estan en cartera
        s = "select distinct(cc.nroint),cc.*, BancosGrales.descripcion as NombreBanco,  clientes.descripcion as NombreCliente,  prov.descripcion as NombreProv " & _
            " from cheques cc " & _
            " left outer join movicaja mm  on mm.interno=cc.nroint " & _
            " LEFT OUTER JOIN BancosGrales ON cc.BANCO_NRO = BancosGrales.codigo " & _
            " LEFT OUTER JOIN  Clientes ON cc.CLIENTE = Clientes.codigo " & _
            " left outer join prov on cc.prov = prov.codigo " & _
            " where mm.tipo='C' " & queCliente & " and mm.ing_egr='I' and cc.activo=1 " & queRangoFecha & " " & _
            " and not exists (select * from movibanc bb where bb.interno=cc.nroint  and (bb.operacion='V' or bb.operacion='R' or bb.operacion='J' or bb.operacion='T' or bb.operacion='A' or bb.operacion='D') and bb.activo=1 " & fec & ") order by " & campoFecha
            
            ' and cc.estado='C'
                
    ElseIf optAcreditarse = True Then ' muestro los depositados por acreditarse
        s = " select distinct(cc.nroint),cc.*, BancosGrales.descripcion as NombreBanco," & _
            " clientes.descripcion as NombreCliente,  prov.descripcion as NombreProv " & _
            " FROM CHEQUES cc " & _
            " left outer join movibanc b on b.interno=cc.nroint " & _
            " LEFT OUTER JOIN BancosGrales ON Cc.BANCO_NRO = BancosGrales.codigo  " & _
            " LEFT OUTER JOIN  Clientes ON Cc.CLIENTE = Clientes.codigo " & _
            " left outer join prov on cc.prov = prov.codigo " & _
            " where b.operacion='D' and b.documento='C' and cc.activo = 1 " & queRangoFecha & queEstado & queCliente & _
            " and not exists (select * from movibanc bb where bb.interno=cc.nroint and (bb.operacion='V' or bb.operacion='R' or bb.operacion='J' or bb.operacion='T' or bb.operacion='A' or bb.operacion='D') and bb.activo=1 " & fec & " )  order by  " & campoFecha

    ElseIf optRechazados = True Then
        s = " select distinct(cc.nroint),cc.*, BancosGrales.descripcion as NombreBanco," & _
            " clientes.descripcion as NombreCliente,  prov.descripcion as NombreProv " & _
            " FROM CHEQUES cc " & _
            " left outer join movibanc b on b.interno=cc.nroint " & _
            " LEFT OUTER JOIN BancosGrales ON Cc.BANCO_NRO = BancosGrales.codigo  " & _
            " LEFT OUTER JOIN  Clientes ON Cc.CLIENTE = Clientes.codigo " & _
            " left outer join prov on cc.prov = prov.codigo " & _
            " where (b.operacion='R' or  b.operacion='V') and b.documento='C' and cc.activo = 1 " & queRangoFecha & queEstado & queCliente & _
            " order by  " & campoFecha

    ElseIf optTransferidos = True Then
        s = " select distinct(cc.nroint),cc.*, BancosGrales.descripcion as NombreBanco, " & _
            " clientes.descripcion as NombreCliente,  prov.descripcion as NombreProv " & _
            " FROM CHEQUES cc " & _
            " left outer join movibanc b on b.interno=cc.nroint " & _
            " LEFT OUTER JOIN BancosGrales ON Cc.BANCO_NRO = BancosGrales.codigo  " & _
            " LEFT OUTER JOIN  Clientes ON Cc.CLIENTE = Clientes.codigo " & _
            " left outer join prov on cc.prov = prov.codigo " & _
            " where b.operacion='T' and b.documento='C' and cc.activo = 1 " & _
            " " & queRangoFecha & queEstado & queCliente & _
            " order by  " & campoFecha
            
    ElseIf optIngresados = True Then
        s = " select distinct(cc.nroint),cc.*, BancosGrales.descripcion as NombreBanco," & _
            " clientes.descripcion as NombreCliente,  prov.descripcion as NombreProv " & _
            " FROM CHEQUES cc " & _
            " left outer join movicaja mm on mm.interno=cc.nroint " & _
            " LEFT OUTER JOIN BancosGrales ON Cc.BANCO_NRO = BancosGrales.codigo " & _
            " LEFT OUTER JOIN  Clientes ON Cc.CLIENTE = Clientes.codigo  " & _
            " left outer join prov on cc.prov = prov.codigo " & _
            " where mm.tipo='C' and mm.ing_egr='I' and cc.activo = 1 " & queRangoFecha & queEstado & queCliente & _
            " and not exists (select * from movibanc bb where bb.interno=cc.nroint and (bb.operacion='V' or bb.operacion='R' or bb.operacion='J' or bb.operacion='D' or bb.operacion='T' or bb.operacion='A') and bb.activo=1 " & fec & " )" & _
            " order by  " & campoFecha
    End If
        
    '***************************************************************************************************
    
'    Debug.Print s
    With rs
    .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                  
    If Not .EOF Then
        Mes = Month(.Fields(Replace(campoFecha, "cc.", "")))
        Anio = Year(.Fields(Replace(campoFecha, "cc.", "")))
        
        While Not .EOF
            rsMes = Month(.Fields(Replace(campoFecha, "cc.", "")))
            rsAnio = Year(.Fields(Replace(campoFecha, "cc.", "")))
            
            If Not (rsMes = Mes) Or (rsMes = Mes And rsAnio <> Anio) Then
                grilla.AddItem "Total " & ObtenerMes(Mes) & " " & Anio & ": " & Total
                Total = 0
            End If
            
            grilla.AddItem vbTab & _
                !Fecha & vbTab & _
                !NroInt & vbTab & _
                !nombreBanco & vbTab & _
                !Nro & vbTab & _
                !Importe & vbTab & _
                !nombreCliente & vbTab & _
                IIf(!procedencia = "T", "TERCEROS", "PROPIO") & vbTab & _
                !tdoc & vbTab & _
                !nDoc & vbTab & _
                !fecha_ingr & vbTab & _
                verEstado(!estado) & vbTab & _
                !NombreProv & vbTab & _
                Left(!tdocProv & "    ", 4) & IIf(s2n(!ndocProv) = 0, "", !ndocProv)

            Total = Total + s2n(!Importe)
            TotalTotal = TotalTotal + s2n(!Importe)
            
            Mes = rsMes
            Anio = rsAnio
            
            .MoveNext
        Wend
        
        grilla.AddItem "Total " & ObtenerMes(Mes) & " " & Anio & ": " & Total
        grilla.AddItem "Total Gral:         " & Format(TotalTotal, "standard")
    End If
    End With
    
    With grilla
        If .rows > 1 Then
            For i = 1 To .rows - 1
                sDato = .TextMatrix(i, 13)
                If sDato > "" Then
                    sDato2 = Split(sDato, " ")
                    If sDato2(0) = "O/P" Then
                        .TextMatrix(i, 14) = sSinNull(obtenerDeSQL("SELECT FECHA FROM REC_COMP WHERE ACTIVO=1 AND NRO=" & s2n(sDato2(1))))
                    ElseIf sDato2(0) = "RAC" Then
                        .TextMatrix(i, 14) = sSinNull(obtenerDeSQL("SELECT FECHA FROM COMPRAS WHERE TIPODOC='RAC' AND ACTIVO=1 AND NRODOC=" & s2n(sDato2(1))))
                        If .TextMatrix(i, 14) = "" Then
                            .TextMatrix(i, 14) = sSinNull(obtenerDeSQL("SELECT FECHA FROM TRANSCOM WHERE TIPODOC='RAC' AND ACTIVO=1 AND NRODOC=" & s2n(sDato2(1))))
                        End If
                    End If
                End If
            Next
        End If
    End With
    
fin:
    relojito False
    Set rs = Nothing
    Exit Sub
ufa:
    ufa "err listado", ""
    Resume fin
End Sub
  
Private Function verEstado(cual As String) As String
    Dim s
    s = "C-Cartera       T-Transferido   R-Rechazado    A-Acreditado   D-Depositado    J-Canjeado  V-Vencido   "
    verEstado = Mid(s, InStr(s, cual & "-") + 2, 12)
End Function


Function ObtenerMes(codigo As Long) As String
    Dim a
    a = Split("X ENE FEB MAR ABR MAY JUN JUL AGO SEP OCT NOV DIC")
    ObtenerMes = a(codigo)
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtFDesde = CDate("1/4/" & Year(Date))
    dtFHasta = Date
    InicioGrilla
    uCliente.ini "select descripcion from clientes where codigo = ### ", "select codigo as [ Codigo    ], descripcion as [ Cliemte                                              ] from clientes order by descripcion"
    uXls.ini grilla, "C:\ListadoCheques", "Listado de Cheques"
    Form_Resize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub InicioGrilla()
    With grilla
        .clear
        .cols = 15
        .rows = 1
        .FormatString = " | Fecha | NroInt | Banco | Nro Cheque | Importe| Cliente | Procedencia | Tipo | NroDoc| FechaIngreso | Estado | Destino | Documento | FechaDoc "
        .ColFormat(5) = "#.00"
        .ColAlignment(5) = flexAlignRightCenter
    End With
    grillaWidth grilla, Array(900, 1050, 750, 2370, 1080, 945, 1560, 1080, 570, 450, 1290)
End Sub

