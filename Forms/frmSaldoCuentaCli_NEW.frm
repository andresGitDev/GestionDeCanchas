VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSaldoCuentaCli_NEW 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Saldo de Cuenta de Clientes"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   Icon            =   "frmSaldoCuentaCli_NEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucCoDe uClie 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   5445
      _extentx        =   9604
      _extenty        =   582
      codigowidth     =   1000
   End
   Begin VB.Frame fraMenu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   60
      TabIndex        =   4
      Top             =   6660
      Width           =   11820
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
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   1455
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
         Height          =   375
         Left            =   10755
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
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
         Left            =   8595
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
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
         Left            =   9570
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   975
      End
      Begin Gestion.ucXls uXls 
         Height          =   765
         Left            =   2655
         TabIndex        =   5
         Top             =   30
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   1349
      End
   End
   Begin VB.Frame fraGri 
      BackColor       =   &H00E0E0E0&
      Height          =   6060
      Left            =   60
      TabIndex        =   1
      Top             =   585
      Width           =   11925
      Begin VB.Frame fraSubGrillas 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2580
         Left            =   120
         TabIndex        =   11
         Top             =   3435
         Width           =   11730
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   2310
            Left            =   15
            TabIndex        =   12
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   195
            Width           =   5850
            _cx             =   10319
            _cy             =   4075
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
         Begin VSFlex7LCtl.VSFlexGrid GrillaMoviCaja 
            Height          =   1035
            Left            =   5940
            TabIndex        =   13
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   180
            Width           =   5745
            _cx             =   10134
            _cy             =   1826
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
         Begin VSFlex7LCtl.VSFlexGrid GrillaEfectivo 
            Height          =   1125
            Left            =   5970
            TabIndex        =   14
            ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
            Top             =   1395
            Width           =   5700
            _cx             =   10054
            _cy             =   1984
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento de Cheques"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5955
            TabIndex        =   17
            Top             =   -15
            Width           =   1710
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Comprobantes que imputa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   15
            TabIndex        =   16
            Top             =   -15
            Width           =   1890
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Movimiento en Efectivo"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5985
            TabIndex        =   15
            Top             =   1200
            Width           =   1665
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   3075
         Left            =   135
         TabIndex        =   2
         Top             =   330
         Width           =   11685
         _cx             =   20611
         _cy             =   5424
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
      Begin VB.Label lblTel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Telefonos:"
         Height          =   315
         Left            =   1110
         TabIndex        =   18
         Top             =   105
         Width           =   5010
      End
   End
   Begin VB.Label lblSaldoTotal 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6345
      TabIndex        =   3
      Top             =   90
      Width           =   5430
   End
End
Attribute VB_Name = "frmSaldoCuentaCli_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_AJUSTE_CLI_DEBITO = "ACD"
Private Const CONST_AJUSTE_CLI_CREDITO = "ACC"
Private Const CONST_FACTURAS_A = "FAA"
Private Const CONST_FACTURAS_B = "FAB"
Private Const CONST_FACTURAS_E = "FAE"
Private Const CONST_NOTAS_DEBITOS_A = "NDA"
Private Const CONST_NOTAS_CREDITOS_A = "NCA"
Private Const CONST_NOTAS_CREDITOS_B = "NCB"
Private Const CONST_NOTAS_DEBITOS_E = "NDE"
Private Const CONST_NOTAS_CREDITOS_E = "NCE"

Private Const CONST_RECIBOS = "RAA" 'RECIBOS A CUENTA

Private Const CONST_CONTADO = 1

Private Enum gri_lla
    griCLCO
    griCLDE
    griFECH
    griTDOC
    griNDOC
    griBSAL
    griVENC
    griDEBE
    griHABE
    griSALD
End Enum

Private mPorUno As Boolean

Dim SaldoCero As Boolean
Dim SaldoTotal As Double
'


Private Sub CalcularSaldo()
    Dim rsAux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Double
    Dim CodigoCli As Long
    Dim CodigoCliActual As Long

    Consulta = "Select * From LIST_SALDO_CLI Order By CLIENTE, FECHA_DOC, ID"
    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rsAux.EOF Then
        rsAux.MoveFirst
        SaldoCero = False
    Else
        SaldoCero = mPorUno
    End If
    SaldoTotal = 0
    While Not rsAux.EOF
        CodigoCli = rsAux!cliente
        CodigoCliActual = CodigoCli
        saldo = 0
        While CodigoCli = CodigoCliActual
            If Not IsNull(rsAux!Debe) And Not IsNull(rsAux!haber) Then saldo = saldo + s2n(rsAux!Debe) - s2n(rsAux!haber)
            Consulta = "Update LIST_SALDO_CLI Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & rsAux!ID
            DataEnvironment1.Sistema.Execute Consulta
            rsAux.MoveNext
            If rsAux.EOF Then
                CodigoCliActual = 0
            Else
                CodigoCliActual = rsAux!cliente
            End If
        Wend
        SaldoTotal = SaldoTotal + saldo
        lblSaldoTotal.caption = "El Saldo Total es: $ " & Format(SaldoTotal, "#,##0.00")
    Wend
End Sub

Private Sub CrearReporte()
Dim rsSaldo As New ADODB.Recordset
Dim Consulta As String

    DataEnvironment1.Sistema.Execute "Delete From LIST_SALDO_CLI"
    
    
    Consulta = "Select CODIGO, TIPODOC, NROFACTURA, FECHA, VENCIMIENTO, CLIENTE, SALDO, total " & _
        " From FACTURAVENTA " & _
        " Where Saldo <> 0 and Activo = 1"
    
    If mPorUno Then Consulta = Consulta & " and CLIENTE = " & uClie.codigo
    
    Consulta = Consulta & " Order By CLIENTE, FECHA, CODIGO"
            
            
    rsSaldo.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rsSaldo.EOF Then rsSaldo.MoveFirst
    While Not rsSaldo.EOF
        Consulta = "Insert Into LIST_SALDO_CLI (CLIENTE, CODIGO_DOC, FECHA_DOC, TIPO_DOC, " & _
                                                "NRO_DOC, VENCIMIENTO, obs, DEBE, HABER, SALDO) " & _
                        "Values (" & s2n(rsSaldo!cliente) & ", " & s2n(rsSaldo!codigo) & _
                            ", " & ssFecha(rsSaldo!fecha) & ", '" & x2s(rsSaldo!TIPODOC) & _
                            "', '" & x2s(rsSaldo!nrofactura) & "', " & ssFecha(rsSaldo!Vencimiento) & _
                            ", " & IIf(rsSaldo!Total <> rsSaldo!saldo, "'(saldo)'", "' '")
        If VaEnElDebe(x2s(rsSaldo!TIPODOC)) Then
            Consulta = Consulta & ", '" & s2n(rsSaldo!saldo, 2) & "', '0', '0')"
        Else
            Consulta = Consulta & ", '0', '" & s2n(rsSaldo!saldo, 2) & "', '0')"
        End If
        
        DataEnvironment1.Sistema.Execute Consulta
        
        rsSaldo.MoveNext
    Wend
    
    CalcularSaldo
    
    rsSaldo.Close
    Set rsSaldo = Nothing
End Sub

Private Sub HabilitarControles(habilitar As Boolean)
    'fraCliente.Visible = habilitar
End Sub

Private Sub LimpiarControles()
    'optuno.Value = True
    mPorUno = False
'    txtCodCliente.Text = ""
'    txtDescCliente.Text = ""
    uClie.codigo = 0
    HabilitarControles True
End Sub

Private Function VaEnElDebe(TipoDocumento As String) As Boolean
    'funcion que devuelve TRUE si el tipo de comprobante va en el DEBE o en el HABER
    If (x2s(TipoDocumento) = CONST_AJUSTE_CLI_DEBITO) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_A) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_B) Or _
                (x2s(TipoDocumento) = CONST_FACTURAS_E) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_E) Or _
                (x2s(TipoDocumento) = CONST_NOTAS_DEBITOS_A) Then
        VaEnElDebe = True
    Else
        VaEnElDebe = False
    End If
End Function

Private Sub cmdAceptar_Click()
    CrearReporte
    LimpiarGrilla Grilla
    LlenarGrilla Grilla, "Select CLIENTE, DESCRIPCION, FECHA_DOC as 'FECHA', TIPO_DOC as 'DOCUMENTO', NRO_DOC AS 'NUMERO DOC.', " & _
                         " obs, VENCIMIENTO, DEBE, HABER, SALDO " & _
                         "From LIST_SALDO_CLI Inner Join CLIENTES On CLIENTES.CODIGO = LIST_SALDO_CLI.CLIENTE " & _
                         "Order By CLIENTE, FECHA_DOC, ID", False, 1
    grillaWidth Grilla, Array(780, 2040, 1125, 550, 850, 750, 1110, 1200, 1200, 1200)
End Sub

Private Sub cmdAyudaCliente_Click()
    frmBuscar.MostrarSql "Select CODIGO, DESCRIPCION From CLIENTES Order By CODIGO"
    If frmBuscar.resultado <> "" Then
        'txtCodCliente.Text = frmBuscar.resultado
        'txtDescCliente.Text = frmBuscar.resultado(2)
        uClie.codigo = frmBuscar.resultado
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Consulta As String
        
    CrearReporte
    
    If Not SaldoCero Then
        DataEnvironment1.rsList_Saldo_Cli.Open
        rptListSaldoCliente.Sections("Encabezado").Controls("lblempresa").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        rptListSaldoCliente.Show vbModal
        DataEnvironment1.rsList_Saldo_Cli.Close
    Else
        MsgBox "El Cliente elegido no tiene saldo.", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
End Sub

Private Sub cmdexcel_Click()
    Dim rs As New ADODB.Recordset
    Dim Consulta As String

    CrearReporte

    Consulta = "Select CLIENTE, FECHA_DOC as 'FECHA', TIPO_DOC as 'DOCUMENTO', NRO_DOC as 'NUMERO DOC.', " & _
                        "obs, VENCIMIENTO, DEBE, HABER, SALDO From LIST_SALDO_CLI "

    Consulta = Consulta & " Order By CLIENTE, FECHA_DOC, ID"


    rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    VinculoXl "C:\SaldoCli.xls", "Saldo a cuenta de Cliente/s", , , rs
    rs.Close
    Set rs = Nothing

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uXls.ini Grilla, "C:\SaldoCliente", "Saldo Clientes   " & Date
    uXls.caption = "Grilla a XLS"
    uClie.ini "Select descripcion from clientes where codigo = '###' and activo = 1", "select Codigo, Descripcion as [Nombre                                        ] from clientes where activo =  1", False
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraGri, Me, anclarLadosTodos
    Anclar fraSubGrillas, Me, anclarAbajo + anclarLadosAncho
    Anclar Grilla, fraGri, anclarLadosTodos
    Anclar fraMenu, Me, anclarAbajo + anclarIzquierda
End Sub

'Private Sub Form_Load()
'    CentrarMe Me
'End Sub

Private Sub grilla_Click()
    Dim TIPODOC As String
    Dim NroDoc As Long
    Dim CodInt As Long
    
    Dim r As Long, c As Long
    Dim clicod As Long
    
    With Grilla
    
        r = .Row
        c = .Col
        clicod = s2n(.TextMatrix(r, griCLCO))
        ' sin clinte
        If clicod = 0 Then Exit Sub
        
        
        'telefono
        lblTel = obtenerDeSQL("select telefono from clientes where codigo = " & clicod)
        TIPODOC = Trim(.TextMatrix(r, griTDOC))
        NroDoc = s2n(.TextMatrix(r, griNDOC))  '  ??? If .TextMatrix(.Row, 5) <> "" Then
            
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaMoviCaja
        LimpiarGrilla GrillaEfectivo
            
        Select Case TIPODOC
         Case CONST_FACTURAS_A, CONST_FACTURAS_B, CONST_FACTURAS_E
            LlenarGrilla GrillaDetalle, _
                " Select D.CANTIDAD, D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO AS 'PRECIO UNITARIO', " & _
                " D.PRECIOTOTAL AS 'PRECIO TOTAL', S.SERIE " & _
                " From FACTURAVENTADETALLE AS D " & _
                " left Join SERIES as S ON S.PRODUCTO=D.PRODUCTO AND S.NROCOMPROBANTE=D.NROFACTURA " & _
                " Where D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & _
                " Order By ID", False

         Case CONST_RECIBOS  ' el recibo no va a aparecer aca...!!
            CodInt = ObtenerDatoDB("RECIBOS", "NUMERO", NroDoc, "CODIGO")
            LlenarGrilla GrillaDetalle, _
                " Select FACTURAVENTA AS 'Factura que imputa', Importe From RECIBOSDETALLE " & _
                " Where CODRECIBO = " & CodInt & " Order By CODIGO", False
            LlenarGrilla GrillaMoviCaja, "Select Fecha, nro as 'Numero Cheque', Importe From CHEQUES " & _
                " Where ACTIVO = 1 And TDOC = '" & TIPODOC & "' AND NDOC = " & NroDoc, True
            LlenarGrilla GrillaEfectivo, "Select Fecha, Importe From MOVICAJA " & _
                " Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                                    
         Case CONST_NOTAS_DEBITOS_A, CONST_NOTAS_CREDITOS_A, CONST_NOTAS_CREDITOS_B
                
         Case CONST_AJUSTE_CLI_DEBITO, CONST_AJUSTE_CLI_CREDITO
                
        End Select
    End With
End Sub
'
'Private Sub opttodos_Click()
'    HabilitarControles False
'End Sub
'
'Private Sub optuno_Click()
'    HabilitarControles True
'End Sub
'
'Private Sub txtCodCliente_GotFocus()
'    PintoFocoActivo
'End Sub
'
'Private Sub txtCodCliente_LostFocus()
'    If Trim(txtCodCliente) <> "" Then
'        txtDescCliente = ObtenerDescripcion("CLIENTES", Val(txtCodCliente))
'    Else
'        txtDescCliente = ""
'    End If
'End Sub

Private Sub uClie_cambio(codigo As Variant)
    mPorUno = (uClie.codigo <> 0)
End Sub
