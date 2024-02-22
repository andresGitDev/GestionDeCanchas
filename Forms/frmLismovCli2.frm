VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLismovCli2 
   Caption         =   "Listado de composicion de saldo de clientes sin detalle"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   3495
      TabIndex        =   19
      Top             =   1440
      Width           =   3615
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Saldo"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Con Saldo"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   420
      Left            =   135
      TabIndex        =   12
      Top             =   6090
      Width           =   11145
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
         Height          =   360
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   45
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
         Height          =   360
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
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
         Height          =   360
         Left            =   10095
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdexcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Enviar a Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
         Width           =   1275
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   360
         Left            =   3150
         TabIndex        =   18
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
      End
   End
   Begin VB.Frame fraGri 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   2070
      Width           =   11175
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2100
         Left            =   390
         TabIndex        =   11
         ToolTipText     =   "Haga Click para ver el Detalle de la Orden de Compra"
         Top             =   1425
         Visible         =   0   'False
         Width           =   10860
         _cx             =   19156
         _cy             =   3704
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
         Rows            =   2
         Cols            =   2
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
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   3495
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Haga Click para ver el Detalle de la Orden de Compra"
         Top             =   255
         Width           =   10980
         _cx             =   19368
         _cy             =   6165
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLismovCli2.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   135
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin Gestion.ucCoDe uCliH 
         Height          =   330
         Left            =   870
         TabIndex        =   6
         Top             =   750
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   582
         CodigoWidth     =   800
         CodigoInvalido  =   0
      End
      Begin Gestion.ucCoDe uCliD 
         Height          =   315
         Left            =   885
         TabIndex        =   7
         Top             =   285
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   556
         CodigoWidth     =   800
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8775
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   107872257
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   107872257
         CurrentDate     =   39347
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmLismovCli2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'sacar iva dela grilla como locaire
' puaj las constantes son un asco   x like 'FA%%'
'


Option Explicit

Private Const CONST_AJUSTE_CLI_DEBITO = "ACD"
Private Const CONST_AJUSTE_CLI_CREDITO = "ACC"
Private Const CONST_FACTURAS_A = "FAA"
Private Const CONST_FACTURAS_B = "FAB"
Private Const CONST_FACTURAS_E = "FAE"
Private Const CONST_NOTAS_DEBITOS_A = "NDA"
Private Const CONST_NOTAS_DEBITOS_B = "NDB"
Private Const CONST_NOTAS_CREDITOS_A = "NCA"
Private Const CONST_NOTAS_CREDITOS_B = "NCB"
Private Const CONST_NOTAS_CREDITOS_E = "NCE"
Private Const CONST_RECIBOS = "RAA"
Private Const CONST_RECIBOS_IMPUTADOS = "REC"
Private TablaTemp As String
Private Const CONST_CONTADO = True
'Private msRsCli As String
Private mfiltro As String

Private Function VaEnElDebe(TipoDocumento As String) As Boolean
'funcion que devuelve TRUE si el tipo de comprobante va en el DEBE o en el HABER
    If (x2s(TipoDocumento) = CONST_AJUSTE_CLI_DEBITO) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_A) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_E) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_B) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_A) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_B) Or _
                (x2s(TipoDocumento) = "FAAV") Or _
                (x2s(TipoDocumento) = "FABV") Or _
                (x2s(TipoDocumento) = "NDAV") Then
        VaEnElDebe = True
    Else
        VaEnElDebe = False
    End If
End Function


Private Function CalcularSaldoAnterior(CodigoCliente As Long, fechahasta As Date) As Double

    Dim Debe As Double
    Dim haber As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim rsaux As New ADODB.Recordset
    Dim sal As Double
    Dim tot As Double
    Dim Consulta As String

    Debe = 0
    haber = 0
    
    'TABLA FACTURAVENTA
'    Consulta = "Select TIPODOC, FORMAPAGO, CONTADO, Sum(TOTAL) as Total,codigo From FACTURAVENTA " & _
'        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & _
'        " Group By TIPODOC, FORMAPAGO, CONTADO,codigo"
    Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
                & " From FACTURAVENTA as F " _
                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCliente & " And F.FECHA<" & ssFecha(fechahasta) _
                & " Order By F.FECHA, F.CODIGO"
                
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rsCuenta!codigo & " and r.fecha " & ssBetween(dtfechad.Value, dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCliente, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsaux.EOF = True And rsaux.BOF = True Then
            If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                'pregunto si la forma de pago es contado, porque con esta no hago nada _
                '(ya que debe sumar en el DEBE y restar en el HABER)
                If Not ( _
                    (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                    And s2n(rsCuenta!contado) = CONST_CONTADO) Then Debe = Debe + s2n(rsCuenta!Total)
            Else
                haber = haber + s2n(rsCuenta!Total)
            End If
            rsCuenta.MoveNext
        Else
            sal = 0
            While Not rsaux.EOF
                sal = sal + rsaux!Importe
                rsaux.MoveNext
            Wend
            If sal = rsCuenta!Total Then
                rsCuenta.MoveNext
            Else
                sal = rsCuenta!Total - sal
                
                If VaEnElDebe(x2s(rsCuenta!TIPODOC)) Then
                    'pregunto si la forma de pago es contado, porque con esta no hago nada _
                    '(ya que debe sumar en el DEBE y restar en el HABER)
                    If Not ( _
                        (x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_A Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_B Or x2s(rsCuenta!TIPODOC) = CONST_FACTURAS_E) _
                        And s2n(rsCuenta!contado) = CONST_CONTADO) Then
                        Debe = Debe + s2n(sal)
                    End If
                Else
                    haber = haber + s2n(sal)
                End If
                rsCuenta.MoveNext
            End If

        End If
        Set rsaux = Nothing
    Wend
    rsCuenta.Close
    Set rsCuenta = Nothing
    
    'TABLA RECIBOS
    Consulta = "Select CLIENTE, SUM(TOTAL) AS TOTAL From RECIBOS " & _
        " Where ACTIVO = 1 And CLIENTE = " & CodigoCliente & " And FECHA < " & ssFecha(fechahasta) & " Group By CLIENTE"
    rsCuenta.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        haber = haber + s2n(rsCuenta!Total)
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close
    Set rsCuenta = Nothing
    
    CalcularSaldoAnterior = Debe - haber
End Function

Private Sub CalcularSaldo()
    Dim rsaux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Double
    Dim CodigoCli As Long
    Dim CodigoCliActual As Long
    Dim sDecimal As Long
    sDecimal = 4
    With rsaux
        Consulta = "Select * From " & TablaTemp & " Order By CODIGO_CLI, FECHA, ID"
        .Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        'If Not rsAux.EOF Then rsAux.MoveFirst
        While Not .EOF
            CodigoCli = !CODIGO_CLI
            CodigoCliActual = CodigoCli
            saldo = 0
            While CodigoCli = CodigoCliActual
                
                'If Not IsNull(!debe) And Not IsNull(rsAux!haber) Then saldo = saldo + s2n(rsAux!debe) - s2n(rsAux!haber)
                saldo = saldo + s2n(!Debe, sDecimal) - s2n(!haber, sDecimal)
                
                If !TIPO_DOCUMENTO <> "" Then
                    'Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & rsAux!ID
                   !saldo = CStr(Round(saldo, sDecimal))
                Else
                    'Consulta = "Update " & TablaTemp & " Set SALDO = ' ' Where ID = " & rsAux!ID
                    !saldo = " "
                End If

                'DataEnvironment1.sistema.Execute Consulta
                .Update
                .MoveNext
                
                If rsaux.EOF Then
                    CodigoCliActual = 0
                Else
                    CodigoCliActual = !CODIGO_CLI
                End If
            Wend
        Wend
    End With
End Sub

Private Sub CrearConsulta(ConDetalle As Boolean)
    Dim SaldoCuenta As Double
    Dim CodigoCli As Long
    Dim Consulta As String
    Dim rsCli As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsFac As New ADODB.Recordset
    Dim rsCHQ As New ADODB.Recordset
    Dim rsaux As New ADODB.Recordset
    Dim sal As Double
    Dim NroRem As String
    Dim sDecimal As Long
    Dim tempoCli As Variant, cliDes As String
    sDecimal = 4
    
    DataEnvironment1.Sistema.Execute "delete from " & TablaTemp
    
    
    rsCli.Open "select * from clientes where activo = 1 and codigo between " & uCliD.codigo & " and " & uCliH.codigo & mfiltro, _
        DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    While Not rsCli.EOF
'    For CodigoCli = CLng(txtcodclid) To (txtcodclih)
    
    
      'ARREGLAR PERFORMANCE !! Aca va parche, pero hay que hacer un rs clientes y barrer el rs
      'tempoCli = (obtenerDeSQL("select codigo, descripcion  from clientes where codigo = '" & CodigoCli & "' and activo = 1"))
      'If Not IsEmpty(tempoCli) Then
        cliDes = rsCli!DESCRIPCION ' tempoCli(1)
        CodigoCli = rsCli!codigo
        'ARREGLAR PERFORMANCE !! Aca va parche, pero hay que hacer un rs clientes y barrer el rs
        
        SaldoCuenta = s2n(CalcularSaldoAnterior(CodigoCli, dtfechad.Value), sDecimal, True)
        'SaldoCuenta = s2n(CalcularSaldoAnterior(46, dtfechad.Value))
        
        If Option3.Value = True Then ' sin saldo =0
            If SaldoCuenta = 0 Then
            Else
                If SaldoCuenta < 0 Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                            ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, sDecimal))) & "', '" & (s2n(SaldoCuenta, sDecimal)) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, sDecimal)) & "', '0',  '" & (s2n(SaldoCuenta, sDecimal)) & "')"
                End If
                Debug.Print Consulta
                DataEnvironment1.Sistema.Execute Consulta
            End If
        End If
        If Option4.Value = True Then
            If SaldoCuenta < 0 Then
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '0', '" & (Abs(s2n(SaldoCuenta, sDecimal))) & "', '" & (s2n(SaldoCuenta, sDecimal)) & "')"
            Else
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoCli & ",'" & ssStr(cliDes) & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '" & (s2n(SaldoCuenta, sDecimal)) & "', '0',  '" & (s2n(SaldoCuenta, sDecimal)) & "')"
            End If
            Debug.Print Consulta
            DataEnvironment1.Sistema.Execute Consulta
        End If
        
        SaldoCuenta = 0
        'If VerParametro(BS_NOMBRE_EMPRESA) = "nimisan swartz" Then
        '    Consulta = "Select F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
        '        & " From FACTURAVENTA as F " _
        '        & " Where ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
        '        & " Order By F.FECHA, F.CODIGO"
        'Else
            Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO " _
                & " From FACTURAVENTA as F " _
                & " Where contado<>1 and ACTIVO = 1 And F.CLIENTE = " & CodigoCli & " And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) _
                & " Order By F.FECHA, F.CODIGO"
        'End If
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            'Consulta = "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli
            rsaux.Open "select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where d.facturaventa=" & rs!codigo & " and r.fecha<=" & ssFecha(dtfechah.Value) & " and activo=1 and r.cliente=" & CodigoCli, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If rsaux.EOF = True And rsaux.BOF = True Then
            
                If s2n(rs!Remito) = 0 Then
                    NroRem = " "
                Else
                    NroRem = s2n(rs!Remito)
                End If
                If VaEnElDebe(x2s(rs!TIPODOC)) Then
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(rs!Total, sDecimal) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva, sDecimal) & "')"
                Else
                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                            " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(rs!Total, sDecimal) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva, sDecimal) & "')"
                End If
                Debug.Print Consulta
                DataEnvironment1.Sistema.Execute Consulta
                    
                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                    Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                        "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                    DataEnvironment1.sistema.Execute Consulta
'                End If
                rs.MoveNext
            Else 'esto es por si tengo algun saldo para mostrar
                sal = 0
                While Not rsaux.EOF
                    sal = sal + rsaux!Importe
                    rsaux.MoveNext
                Wend
                If sal = rs!Total Then
                    rs.MoveNext
                Else
                    If s2n(rs!Remito) = 0 Then
                        NroRem = " "
                    Else
                        NroRem = s2n(rs!Remito)
                    End If
                    sal = rs!Total - sal
                    If VaEnElDebe(x2s(rs!TIPODOC)) Then
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '" & s2n(sal, sDecimal) & "', '0', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    Else
                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO, REMITO, IVA) " & _
                                " VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!Fecha) & _
                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroFactura) & "', '0', '" & s2n(sal, sDecimal) & "', '" & SaldoCuenta & "', '" & NroRem & "', '" & s2n(rs!Iva) & "')"
                    End If
                    Debug.Print Consulta
                    DataEnvironment1.Sistema.Execute Consulta
                        
                    'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                    If (x2s(rs!TIPODOC) = CONST_FACTURAS_A Or x2s(rs!TIPODOC) = CONST_FACTURAS_B Or x2s(rs!TIPODOC) = CONST_FACTURAS_E) And rs!contado = CONST_CONTADO Then
'                        Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                                    "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                            "VALUES (" & CodigoCli & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                                                    ", 'CON', '" & x2s(rs!nrofactura) & "', '0', '" & s2n(rs!Total, 2) & "', '" & SaldoCuenta & "')"
'                        DataEnvironment1.sistema.Execute Consulta
'                    End If
                    rs.MoveNext
                End If

            End If
            Set rsaux = Nothing
        Wend
        rs.Close
        Set rs = Nothing
        
        rsCli.MoveNext
    Wend
    
    CalcularSaldo

    Set rs = Nothing
    Set rsCli = Nothing
End Sub

Private Sub cmdAceptar_Click()
    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If dtfechad.Value < CDate("01/04/2006") Then
        MsgBox "Debe ingresar una fecha posterior al 01/04/2006."
        dtfechad.Value = "01/04/2006"
        Exit Sub
    End If
    
    If rangoOk Then
                
        relojito
        
        CrearConsulta False
        LimpiarGrilla GRILLA
        
        LlenarGrilla GRILLA, _
            " Select CODIGO_CLI AS CODIGO,  DESCRIPCION_CLI as DESCRIPCION, FECHA, TIPO_DOCUMENTO, " & _
            " NRO_DOCUMENTO, REMITO, DEBE, HABER, SALDO, '' as [Saldo final] " & _
            " From " & TablaTemp & _
            " Where facturas = ' ' and cheques = ' ' " & _
            " Order By CODIGO_CLI, FECHA, ID", True
        grillaMarcoSaldosFinales GRILLA, 0, 9, 8
        limpioGrilla 9
        relojito False
        
'        MsgBox "" & Grilla.rows
        filtragrilla
        
    End If
End Sub

Private Function limpioGrilla(Col As Long) 'limpio la grilla y tabla temporal de los importes con cero, incluyendo si el total es cero borro historial de cliente
    Dim i As Long
    Dim j As Long
    Dim cli As Long
    Dim Borrar As String
    
    i = 1
    While i < GRILLA.rows
        If GRILLA.TextMatrix(i, Col) = "0" Then
            cli = GRILLA.TextMatrix(i, 0)
            j = 1
            While j < GRILLA.rows
                If GRILLA.TextMatrix(j, 0) = CStr(cli) Then
                    GRILLA.TextMatrix(j, 0) = ""
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
    
    j = 1
    While j < GRILLA.rows
        If GRILLA.TextMatrix(j, 0) = "" Then
            Borrar = "delete from " & TablaTemp & " where descripcion_cli='" & GRILLA.TextMatrix(j, 1) & "'"
            DataEnvironment1.Sistema.Execute Borrar
            GRILLA.RemoveItem (j)
        Else
            j = j + 1
        End If
        'j = j + 1
    Wend
End Function

Private Function filtragrilla()
    Dim i As Long
    Dim cli As Long
    Dim Valor As Double
    Dim CUIT As String
    Dim nom As String
    
    grilla2.rows = 1
    i = 1
    Valor = 0
    If GRILLA.rows > 1 Then
        If s2n(GRILLA.TextMatrix(i, 6)) > 0 Then
            Valor = s2n(GRILLA.TextMatrix(i, 6))
        Else
            Valor = -s2n(GRILLA.TextMatrix(i, 7))
        End If
        cli = GRILLA.TextMatrix(i, 0)
        i = i + 1
        While i < GRILLA.rows
            If GRILLA.TextMatrix(i, 0) <> "" Then
                If cli = GRILLA.TextMatrix(i, 0) Then
                    If s2n(GRILLA.TextMatrix(i, 6)) > 0 Then
                        Valor = Valor + s2n(GRILLA.TextMatrix(i, 6))
                    Else
                        Valor = Valor - s2n(GRILLA.TextMatrix(i, 7))
                    End If
                Else
                    CUIT = obtenerDeSQL("select cuit from clientes where codigo=" & cli)
                    nom = obtenerDeSQL("select descripcion from clientes where codigo=" & cli)
                    grilla2.AddItem CUIT & Chr(9) & nom & Chr(9) & s2n(Valor)
                    Valor = 0
                    cli = GRILLA.TextMatrix(i, 0)
                    If s2n(GRILLA.TextMatrix(i, 6)) > 0 Then
                        Valor = Valor + s2n(GRILLA.TextMatrix(i, 6))
                    Else
                        Valor = Valor - s2n(GRILLA.TextMatrix(i, 7))
                    End If
                End If
            End If
            i = i + 1
        Wend
        CUIT = obtenerDeSQL("select cuit from clientes where codigo=" & cli)
        nom = obtenerDeSQL("select descripcion from clientes where codigo=" & cli)
        grilla2.AddItem CUIT & Chr(9) & nom & Chr(9) & s2n(Valor)
    Else
        MsgBox "No se han encontrado datos para la busqueda.", , "ATENCION"
    End If
End Function

Private Sub cmdexcel_Click()
Dim rs As New ADODB.Recordset
Dim Consulta As String

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then
'        CrearConsulta False

        If MsgBox("¿Desea imprimir los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
                    " DEBE, HABER, SALDO " & _
                    " From  " & TablaTemp & _
                    " Order By CODIGO_CLI, FECHA, ID"
                                
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO " & _
                    " From " & TablaTemp & _
                    " Where TIPO_DOCUMENTO <> '' " & _
                    " Order By CODIGO_CLI, FECHA, ID"
        End If
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
        
        VinculoXl "C:\ComposicionCuentaCli.xls", "Composicion cuenta de Cliente/s", , , rs '"C:\ComposicionCuentaCli", "Composicion cuenta cliente"
        rs.Close
        Set rs = Nothing
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Consulta As String
    Dim rsempresa As New ADODB.Recordset
    Dim i As Long

    
    
    'rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.sistema, adOpenStatic, adLockReadOnly
    'RptLisMovCtaProv2.ImageLOGO.Picture = FrmPrincipal.imgLogoSimple 'LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
    
    If grilla2.rows < 2 Then Exit Sub
    
    grilla2.GridLines = flexGridNone
    grilla2.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.MarginLeft = 2000
    FrmImpresiones.VSPrinter.Orientation = orPortrait ' orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla2.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 20
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 16
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Listado de composicion de clientes sin detalle"
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & dtfechad.Value & " - " & dtfechah.Value  '& "     Rango de Cuentas : " & CmbCtaD & "  -  " & CmbCtaH
    FrmImpresiones.VSPrinter.Paragraph = " "
    
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    FrmImpresiones.VSPrinter.RenderControl = grilla2.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla2.GridLines = flexGridFlat
    
    'Set rsempresa = Nothing
End Sub


Private Sub cmdCancelar_Click()
    Dim tempo
    tempo = obtenerDeSQL("select min(codigo) as mini, max(codigo) as maxi from clientes")
    'txtcodclih = tempo(1) ' "9999999"
    uCliH.codigo = tempo(1)
'    txtclienteh = ""
    'txtcodclid = tempo(0)
    uCliD.codigo = tempo(0)
'    txtcliented = "" 'tempo(1)
    dtfechad.Value = "01/04/2007" 'Date
    dtfechah.Value = Date
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
'    cmdCancelar_Click
    
    Dim s1 As String, s2 As String, sCli As String, tempo As Variant
    
    TablaTemp = TablaTempCrear("([ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_CLI] [numeric](18, 0) NULL ," _
        & "[DESCRIPCION_CLI] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FECHA] [datetime] NULL ," _
        & "[TIPO_DOCUMENTO] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[NRO_DOCUMENTO] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[IVA] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DEBE] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[HABER] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[SALDO] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FACTURAS] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[CHEQUES] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[REMITO] [nvarchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" _
        & ") ON [PRIMARY]")
    
    DataEnvironment1.Sistema.Execute " ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [DF_" & TablaTemp & "] DEFAULT (N' ') FOR [FACTURAS]," _
        & " CONSTRAINT [DF_" & TablaTemp & "1] DEFAULT (N' ') FOR [CHEQUES]"
    
    ucXls1.ini grilla2, "C:\ComposicionCliSinDet.xls", "Composicion cuenta cliente sin detalle"
    
    mfiltro = ""
    
    
    s1 = "select descripcion from clientes where codigo = ### and activo = 1 "
    s2 = "Select codigo, descripcion as [ Descripcion                                           ] from clientes where activo = 1 " '& mfiltro

    uCliD.ini s1, s2, False, True
    uCliH.ini s1, s2, False, True
    
    cmdCancelar_Click
    Form_Resize
    ini
End Sub
Private Sub ini()
    grilla2.rows = 1
    grilla2.TextMatrix(0, 0) = "CUIT"
    grilla2.TextMatrix(0, 1) = "NOMBRE"
    grilla2.TextMatrix(0, 2) = "SALDO"
End Sub
Private Sub Form_Resize()
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
    Anclar fraGri, Me, anclarLadosTodos
    Anclar GRILLA, fraGri, anclarLadosTodos
End Sub


Private Sub ucXls1_Clic(cancel As Boolean)
    Dim Consulta As String

    'If Trim(txtcodclid) <> "" And Trim(txtcodclih) <> "" Then
    If rangoOk Then

        If MsgBox("¿Desea los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
                    " DEBE, HABER, SALDO, '' as [Saldo Final] " & _
                    " From  " & TablaTemp & _
                    " Order By CODIGO_CLI, FECHA, ID"
            CrearConsulta True
            LlenarGrilla GRILLA, Consulta, False
            grillaMarcoSaldosFinales GRILLA, 0, 10, 9
            limpioGrilla 10
                    
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', DEBE, HABER, SALDO, '' as  Final " & _
                    " From " & TablaTemp & _
                    " Where TIPO_DOCUMENTO <> '' " & _
                    " Order By CODIGO_CLI, FECHA, ID"
            CrearConsulta True
            LlenarGrilla GRILLA, Consulta, False
            grillaMarcoSaldosFinales GRILLA, 0, 8, 7
            limpioGrilla 8
        End If
        
    Else
        MsgBox "Debe seleccionar un cliente donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
        cancel = True
    End If
     
End Sub
Private Function rangoOk() As Boolean
    rangoOk = (uCliD.codigo > 0 And uCliH.codigo > 0)
End Function




