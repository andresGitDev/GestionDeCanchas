VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecibosACuenta 
   Caption         =   "Recibos a Cuenta"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9390
   Icon            =   "frmRecibosACuenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   720
      Left            =   15
      TabIndex        =   25
      Top             =   5070
      Width           =   9150
      Begin VB.TextBox txtTransferencia 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   27
         Top             =   285
         Width           =   1275
      End
      Begin Gestion.uCtaBanco uCtaBanco 
         Height          =   345
         Left            =   1395
         TabIndex        =   26
         Top             =   285
         Width           =   7590
         _extentx        =   13388
         _extenty        =   609
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   3
         Left            =   15
         TabIndex        =   28
         Top             =   30
         Width           =   1275
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Height          =   1695
      Left            =   -15
      TabIndex        =   20
      Top             =   5910
      Width           =   9390
      _extentx        =   16563
      _extenty        =   2990
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      Begin Gestion.ucEntreFechas ucFechas 
         Height          =   315
         Left            =   1650
         TabIndex        =   23
         Top             =   60
         Width           =   2595
         _extentx        =   4577
         _extenty        =   556
      End
      Begin VB.Label Label12 
         Caption         =   "Buscar Entre"
         Height          =   195
         Left            =   330
         TabIndex        =   24
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   9075
      Begin VB.CommandButton cmdCliente 
         Caption         =   "?"
         Height          =   315
         Left            =   3045
         TabIndex        =   21
         Top             =   645
         Width           =   375
      End
      Begin VB.TextBox txtCuentaEfectivo 
         Height          =   315
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1395
      End
      Begin VB.TextBox txtCaja 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   4050
         TabIndex        =   7
         Text            =   "1"
         Top             =   1890
         Width           =   675
      End
      Begin VB.TextBox txtNumero 
         Height          =   375
         Left            =   1710
         TabIndex        =   1
         Top             =   195
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Height          =   255
         Left            =   1770
         TabIndex        =   5
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtEfectivo 
         Height          =   300
         Left            =   1770
         TabIndex        =   6
         Top             =   1890
         Width           =   1275
      End
      Begin VB.CommandButton cmdBorraItem 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   8190
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmRecibosACuenta.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Borrar Item"
         Top             =   2010
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   375
         Left            =   1770
         TabIndex        =   4
         Top             =   1005
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   77266945
         CurrentDate     =   38252
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   3525
         TabIndex        =   3
         Top             =   645
         Width           =   4095
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1725
         TabIndex        =   2
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Caja :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3210
         TabIndex        =   18
         Top             =   1890
         Width           =   735
      End
      Begin VB.Label lblTipo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RAA"
         Height          =   315
         Left            =   6960
         TabIndex        =   17
         Top             =   180
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   240
         X2              =   8880
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lbl7 
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6060
         TabIndex        =   16
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lblCodigo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Left            =   7740
         TabIndex        =   15
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   855
         TabIndex        =   13
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "en Efectivo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   540
         TabIndex        =   12
         Top             =   1935
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "en Cheques :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   525
         TabIndex        =   11
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Nro Recibo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   390
         TabIndex        =   10
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   810
         TabIndex        =   9
         Top             =   1065
         Width           =   735
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   765
         TabIndex        =   8
         Top             =   645
         Width           =   915
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid gCheques 
      Height          =   2475
      Left            =   360
      TabIndex        =   22
      Top             =   2640
      Width           =   8475
      _cx             =   14949
      _cy             =   4366
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
End
Attribute VB_Name = "frmRecibosACuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'TO DO:
'  VERIFICAR Q NRO RECIBO NO ESTE !!
'

Option Explicit '2/12/4
' Lito Explicit 22/8/4

'movibanc: en DOS se guardaba, pero no vemos q sea util, porq no representa un mov de cuenta bancaria hasta q se deposita
' deberia unificar SP con los de AMIR, aunq no estoy seguro xq se usan parametros diferentes

Private midDoc As Long
Private cliente As LiCodigo
Private caja As LiCodigo
Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1

Private gBANCC  As Long
Private gBANCD  As Long
Private gNROCH  As Long
Private gMONTO  As Long
Private gFECHA  As Long
Private gPT     As Long
Private gCODCH  As Long
Private Const tt_ChequeRaCuentaTmp = _
"( [nroint] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
"[banco] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
"[cheque] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[importe] [float] NULL ,  [fecha] [datetime] NULL, " & _
"[propio] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL )"
'

'

'-------------------------------------------------------------
'
Private Sub cmdBorraItem_Click()
    If g.Row > 1 Then g.delRow (g.Row)
End Sub


Private Sub Form_Load()
'    CentrarMe Me
    Me.KeyPreview = True
    
    inigrilla
    iniCliente
    inimenu

    limpiar
    txtCaja = 1
    verCajaEfectivo
    Form_Resize
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub


' ini ----------------------------
Private Sub iniCliente()
    Set cliente = New LiCodigo
    cliente.init cmbCliente, txtCodCliente, "clientes", False, False, cmdCliente, "activo = 1", True
End Sub
Private Sub inimenu()
    ucMenu.init True, True, False, True, True
    ucMenu.MsgConfirmaSalir = "Cerrar Ventana ?"
    ucMenu.MsgConfirmaEliminar = "Elimina Recibo ?"
End Sub
Private Sub inigrilla()
    Set g = New LiGrilla
    With g
        .init gCheques
        gBANCC = .AddCol(" Banco ", "N", 0)
        gBANCD = .AddCol("  Banco                        ")
        gNROCH = .AddCol("  Nro Cheque      ", "S")
        gMONTO = .AddCol("  Monto     ", "N")
        gFECHA = .AddCol("  Fecha     ", "D")
        gPT = .AddCol(" P/T ", "S")
        gCODCH = .AddCol("Cod Interno")
    End With
End Sub


Private Sub Form_Resize()
    Anclar fraTra, Me, anclarAbajo + anclarIzquierda
    Anclar gCheques, Me, anclarLadosTodos
End Sub

' grilla -----------------------------------
Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim deba As String

    Select Case Col
    Case gBANCC ': g.tx Row, gBANCD, ObtenerDescripcion("BancosGrales", s2n(txt))
        deba = verDescBanco(s2n(txt))
        g.tx Row, gBANCD, deba
        If deba = "" Then g.tx Row, Col, ""
    End Select
    
    If g.Row = g.rows - 1 Then g.addRow
  
End Sub

Private Sub g_DblClick()
    If g.Col = gBANCC Then g.tx g.Row, g.Col, frmBuscar.MostrarCodigoDescripcionActivo("BancosGrales")
End Sub

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    Select Case Col
    Case gPT: cancel = (g.EditText <> "P" And g.EditText <> "T")
        
    End Select
End Sub



Private Sub txtcaja_GotFocus()
    If Trim$(txtCaja) = "" Then txtCaja = "1"
    PintoFocoActivo
End Sub

Private Sub txtCaja_Validate(cancel As Boolean)
    cancel = Not verCajaEfectivo()
End Sub

'------------------------------
Private Sub txtEfectivo_Validate(cancel As Boolean)
    If Not IsNumeric(txtEfectivo) Then cancel = True
    txtEfectivo = s2n(txtEfectivo)
End Sub
Private Sub txtTotal_Validate(cancel As Boolean)
    If Not IsNumeric(txtTotal) Then cancel = True
    txtTotal = s2n(txtTotal)
End Sub


Private Sub txtTransferencia_LostFocus()
    txtTransferencia = s2n(txtTransferencia)
End Sub

' menu ------------------------------
Private Sub ucMenu_Aceptar()
    If Falta Then Exit Sub
    If GrabaRecibo Then
        MsgBox "Operacion completa"
        ImprimirReciboAcuenta
        ucMenu.AceptarOk
    End If
End Sub

Private Sub ucMenu_BorrarControles()
    limpiar
End Sub

Private Sub ucMenu_Buscar()
    If BuscaRecibo() Then ucMenu.BuscarOK
End Sub
Private Sub ucMenu_eliminar()
    If EliminaRecibo() Then ucMenu.EliminarOK
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    fraControl.enabled = sino
    'g.esEditable = sino
    gCheques.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
End Sub

Private Sub ucMenu_Imprimir()
ImprimirReciboAcuenta
End Sub

Private Sub ImprimirReciboAcuenta()
Dim sTablaTempReciboAcuenta As String
Dim rs As New ADODB.Recordset
Dim sql, sqlTemp As String
Dim i As Integer
    
sTablaTempReciboAcuenta = TablaTempCrear(tt_ChequeRaCuentaTmp)
With g
If .rows > 1 Then
   For i = 1 To .rows - 1
     If .TextMatrix(i, 0) = "" Or .TextMatrix(i, 1) = "" Then Exit For
     sql = "insert into " & sTablaTempReciboAcuenta & "" & _
     " (nroint,banco,cheque,importe,fecha,propio)" & _
     " values( '" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 1) & "'," & _
     " '" & .TextMatrix(i, 2) & "'," & x2s(.TextMatrix(i, 3)) & "," & ssFecha(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 5) & "') "
     DataEnvironment1.Sistema.Execute sql
   Next i
End If
End With
RptReciboAcuenta.lblfecha = dtFecha
RptReciboAcuenta.LblCodCliente = txtCodCliente
RptReciboAcuenta.TxtNroRecibo = Format(txtNumero, "00000000")
RptReciboAcuenta.lblcliente = cmbCliente.Text
RptReciboAcuenta.LblImporte = enletras(txtTotal)
RptReciboAcuenta.txtEfectivo = txtEfectivo
RptReciboAcuenta.TxtCheques = s2n(txtTotal) - s2n(s2n(txtEfectivo) + s2n(txtTransferencia))
RptReciboAcuenta.txtTotal = s2n(txtTotal)
RptReciboAcuenta.txtTransferencia = s2n(txtTransferencia)
sqlTemp = "select * from " & sTablaTempReciboAcuenta & ""
rs.Open sqlTemp, DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
RptReciboAcuenta.Data.Connection = DataEnvironment1.Sistema
RptReciboAcuenta.Data.Source = sqlTemp
RptReciboAcuenta.Restart

If PREVIEW_IMPRESIONES Then
    RptReciboAcuenta.Show
Else
    RptReciboAcuenta.PrintReport False
End If

End Sub


Private Sub ucMenu_Nuevo()
    limpiar
    'txtNumero.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
'------------------------------------

Private Sub limpiar()
    FrmBorrarTxt Me
    g.Borrar
    g.rows = 40
    dtFecha = Date
    ucFechas.ini
    txtCaja = 1
    verCajaEfectivo
    cliente.codigo = 0
    midDoc = 0
    uCtaBanco.codigo = 0
End Sub

Private Function Falta() As Boolean ' o si sobra
    Dim i As Long, tmp
    Falta = True
    
    'cabecera
    If txtNumero = "" Or cliente.DESCRIPCION = "" Or s2n(txtTotal) = 0 Then
        che "faltan datos en cabecera: Numero, Cliente, Total"
        Exit Function
    End If
    
    'Nro Repetido
    If YaEstaRecibo(txtNumero) Then
        che "Recibo ya cargado"
        Exit Function
    End If
    
    'grilla
    i = g.PrimerVacio(gBANCD)
    If i <> g.PrimerVacio(gNROCH) Or i <> g.PrimerVacio(gMONTO) Or i <> g.PrimerVacio(gFECHA) Or i <> g.PrimerVacio(gPT) Then
        che "revisar datos en grilla (informacion incoherente)"
        Exit Function
    End If
        
    'transf
    If s2n(txtTransferencia) > 0 And uCtaBanco.codigo = 0 Then
        che "Falta cuenta bancaria"
        Exit Function
    End If
    
    'efectivo
    If s2n(txtEfectivo) <> 0 And Trim(txtCuentaEfectivo) = "" Then
        che "revisar cuenta caja efectivo"
        Exit Function
    End If
            
    'montos
    If s2n(g.suma(gMONTO) + s2n(txtEfectivo) + s2n(txtTransferencia) - s2n(txtTotal)) <> 0 Then
        che "No coinciden los montos"
        Exit Function
    End If
    
   
    Falta = False
End Function


'Private Sub che(que)
'    MsgBox que, vbExclamation, "Aviso"
'End Sub

 
Private Function GrabaRecibo() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim CantCheques, i As Long, nMoviC, MoviConcepto, DetConcepto ', nMoviB
    Dim chNro, chMonto As Double, chCod, chFech As Date, chPT, chCta, bco, iddoc As Long, AsiVta As New Asiento
    Dim TextoAsientoComprobante As String
    Dim tra As Double
    Dim nMoviB As Long ', nMoviC As Long
    Dim z As Double
    
    z = 0
    
    lblCodigo = nuevoCodigo("FacturaVenta", "codigo")

    
    CantCheques = g.PrimerVacio(gMONTO) - 1
   'MoviConcepto = "Recibo 00000973  Cliente 2028   NASTIQUE OSCAR ALF"
    MoviConcepto = "Recibo " & Format(txtNumero, "00000000") & "Cliente " & Format(cliente.codigo, "0000") & "   " & cliente.DESCRIPCION
    DetConcepto = "RecCta " & Format(txtNumero, "00000000") & "  Cliente " & Format(cliente.codigo, "0000")
    
 
 
    ' --- Transaccion aqui -----------  ini
    DE_BeginTrans
    
    iddoc = NuevoDocumento(lblTipo, txtNumero, 0, 0)
    TextoAsientoComprobante = "RAC " & txtNumero
    AsiVta.nuevo "Rec " & txtCodCliente, dtFecha, "RECV"
    AsiVta.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), 0, s2n(txtTotal)
    
    'ComprobanteRecibo
    DataEnvironment1.dbo_abmFacturaVenta "A", s2n(lblCodigo), lblTipo, s2n(txtNumero), dtFecha, dtFecha, 0, 0, cliente.codigo, cliente.DESCRIPCION, "", "", 0, 0, s2n(txtTotal), 0, 0, s2n(txtTotal), s2n(txtTotal), 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, 0, 0, iddoc         ' s2n(txtcotizacion),
    
    'Efectivo
    If s2n(txtEfectivo) > 0 Then
        'MoviCaja
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        DataEnvironment1.dbo_MOVICAJASdoc "A", s2n(txtCaja), nMoviC, cliente.codigo, "E", "I", s2n(txtEfectivo), MoviConcepto, dtFecha.Value, Trim(txtCuentaEfectivo), 0, 1, lblTipo, txtNumero, Date, UsuarioActual(), 0, iddoc
'''''        'detMovCaja
'''''        DataEnvironment1.dbo_DETMOVCAJAS "A", nMoviC, s2n(txtefectivo), cliente.codigo, Trim(txtCuentaEfectivo), DetConcepto, "RA"
        AsiVta.AgregarItem (txtCuentaEfectivo), s2n(txtEfectivo), 0, TextoAsientoComprobante
    End If
   
    'CHEQUES
    For i = 1 To CantCheques
        chNro = g.tx(i, gNROCH)
        chMonto = s2n(g.tx(i, gMONTO))
        chCta = CuentaParam(ID_Cuenta_M_CH_CARTERA) 'obtenerParametro("Cta_Caja")
        chFech = CDate(g.tx(i, gFECHA))
        chPT = g.tx(i, gPT)
        bco = s2n(g.tx(i, gBANCC))
        '
        chCod = nuevoCodigo("cheques", "NroInt")
        g.tx i, gCODCH, chCod
          
        'Cheques
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "A", chCod, chFech, chNro, chMonto, s2n(txtNumero), lblTipo, dtFecha, Date, "C", bco, chPT, cliente.codigo, Date, UsuarioActual(), iddoc, 0
        
        'movi
        'nMoviB = nuevoCodigo("movibanc", "movBanco")
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        '    'moviBanc Ver Nota
        '    'daTaenvironment1.dbo_MOVIbancos "A", , "I", MoviConcepto, Date, "C", chMonto, nMoviB, Date, UsuarioActual(), 0, 0, 1
        'MoviCaja
        DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, "C", "I", chMonto, MoviConcepto, dtFecha.Value, chCta, 0, 1, lblTipo, txtNumero, Date, UsuarioActual, chCod, iddoc
        'DetMovCaja
'        DataEnvironment1.dbo_DETMOVCAJAS "A", nMoviC, chMonto, cliente.codigo, 0, DetConcepto, "RA"
        
        AsiVta.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), chMonto, 0, TextoAsientoComprobante
    Next i
    
    'transferencia
    tra = s2n(txtTransferencia)
    If tra > 0 Then
        nMoviC = nuevoCodigo("movicaja", "movimiento")
        nMoviB = nuevoCodigo("movibanc", "movBanco")
        DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, "T", "I", tra, MoviConcepto, dtFecha, "", nMoviB, 1, "REC", txtNumero, Date, UsuarioActual(), 0, iddoc
        DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "E", Left(MoviConcepto, 50), dtFecha, "E", tra, nMoviB, iddoc, Date, UsuarioActual(), z
        AsiVta.AcumularItem uCtaBanco.CuentaContable, tra, 0, TextoAsientoComprobante
'        AsiVta.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), 0, tra, TextoAsientoComprobante
    End If
    

    



    If siAsiento("AsientosRecibos") Then AsiVta.Grabar iddoc
        'If AsiVta.Grabar(iddoc) = 0 Then
'            DE_RollbackTrans
'            ufa "Err al grabar asiento ", ""
'            GoTo fin
'        End If
    
    
    DE_CommitTrans
    midDoc = iddoc
    ' --- Transaccion aqui -----------  fin
    
    GrabaRecibo = True
    If CantCheques > 0 Then che "Operacion concluida" & vbCrLf & "Puede anotar los codigos internos de los cheques"
    
fin:
    Exit Function
    
ufaErr:
    DE_RollbackTrans
    ufa "error grabando Recibo", Me.Name ', Err
    Resume fin
End Function


Private Function BuscaRecibo() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim i As Long, rs As New ADODB.Recordset
    Dim tempo
    
    With frmBuscar
'        If .mostrarSql("select codigo, NroFactura, Fecha, Cliente, Total from FacturaVenta where activo = 1 and TipoDoc = 'RAA' and fecha " & ssBetween(dtDesde, dtHasta)) > "" Then
        If .MostrarSql("select codigo, NroFactura as [ Recibo    ], Fecha, Cliente, Total , iddoc from FacturaVenta where activo = 1 and TipoDoc = '" & TipoDoc_RECIBO & "' and fecha " & ucFechas.ssBetween() & " order by codigo desc ") > "" Then
            limpiar
            'cabecera
            lblCodigo = s2n(.resultado(1))
            txtNumero = s2n(.resultado(2))
            dtFecha = CDate(.resultado(3))
            cliente.codigo = s2n(.resultado(4))
            txtTotal = s2n(.resultado(5))
            midDoc = s2n(.resultado(6))
            
            'cargar efectivo y cheques
            With rs
'                Dim tempo ', rs As New ADODB.Recordset ', i As Long
                tempo = obtenerDeSQL("select caja, importe from movicaja where activo = 1 and tipodoc = 'RAA' and NroDoc = " & txtNumero & " and tipo = 'E' ")
                If Not IsEmpty(tempo) Then
                    txtEfectivo = s2n(tempo(1))
                    txtCaja = s2n(tempo(0), 0)
                End If
'                .Close
                
                g.Borrar
                .Open "select  NroInt, Fecha, Nro, Importe, Banco_Nro, procedencia from cheques where tdoc = 'RAA' and nDoc = " & txtNumero, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                While Not .EOF
                    i = g.addRow()
                    g.tx i, gBANCC, !BANCO_NRO
                    g.tx i, gCODCH, !NroInt
                    g.tx i, gFECHA, !Fecha
                    g.tx i, gMONTO, !Importe
                    g.tx i, gNROCH, !Nro
                    g.tx i, gPT, !procedencia
                    .MoveNext
                Wend
                .Close
                
                'transf
                tempo = obtenerDeSQL("select cuenta, importe from movibanc where iddoc = " & midDoc)
                If Not IsEmpty(tempo) Then
                    txtTransferencia = tempo(1)
                    uCtaBanco.codigo = tempo(0)
                End If
            End With
        End If
    End With
    BuscaRecibo = True
    GoTo fin
    
ufaErr:
    ufa "err cargando recibo", Me.Name & " " & frmBuscar.resultado(1) ', Err
fin:
    Set rs = Nothing
End Function

Private Function EliminaRecibo() As Boolean
    Dim CantCheques, i As Long
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    If Trim$(txtNumero) = "" Or Trim$(lblTipo) = "" Then
        che "no puedo eliminar, inf insuficiente: numero, codigo, tipo"
        Exit Function
    End If
    
    Dim sp, COD
    COD = obtenerDeSQL("select codigo from facturaventa where tipodoc='RAA' and nrofactura=" & s2n(txtNumero))
    sp = obtenerDeSQL("select * from recibosdetalle d inner join recibos r on d.codrecibo=r.codigo where r.activo=1 and d.facturaventa=" & s2n(COD))
    If IsNull(sp) Or IsEmpty(sp) Then
    Else
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Function
    End If
    
    CantCheques = g.PrimerVacio(gMONTO) - 1
    For i = 1 To CantCheques
        'If Trim(g.tx(i, gPT)) = "T" Then
            If Trim(obtenerDeSQL("select estado from cheques where nroint=" & g.tx(i, gCODCH))) <> "C" Then
                MsgBox "No se puede eliminar el recibo debido a que el cheque interno " & g.tx(i, gCODCH) & " esta siendo utilizado en algun comprobante.", vbInformation, "ATENCION"
                GoTo ufaErr
            End If
        'End If
    Next i
    
    
    DE_BeginTrans
    'ComprobanteRecibo
    DataEnvironment1.dbo_abmFacturaVenta "B", s2n(lblCodigo), "", 0, 0, 0, 0, 0, 0, "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0     ' s2n(txtcotizacion),

    'cheques
    'daTaenvironment1.dbo_INGRESOCHTERCEROS "B"
    DataEnvironment1.Sistema.Execute "update cheques set activo = 0, fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & UsuarioActual() & " where ndoc = '" & txtNumero & "' and tdoc = '" & lblTipo & "'"
    'movicaja
    DataEnvironment1.dbo_MOVICAJASdoc "B", 0, 0, 0, "", "", 0, "", Date, "", 0, 0, lblTipo, txtNumero, Date, UsuarioActual(), 0, midDoc
''     'DetMovCaja
''      daTaenvironment1.dbo_DETMOVCAJAS "A", nMoviC, chMonto, cliente.Codigo, 0, DetConcepto, "RA"

            'Baja Doc y asiento
            If Not BorroDocumento(midDoc) Then
                ufa "err al borrar documento", " middoc = " & midDoc
                DE_RollbackTrans
                GoTo fin
            End If
    DE_CommitTrans
    
    EliminaRecibo = True
    che "Eliminado"
    
fin:
    Exit Function
ufaErr:
    DE_RollbackTrans
    ufa "error en la baja", Me.Name & txtNumero ', Err
    Resume fin
End Function

Private Function verCajaEfectivo() As Boolean
    Dim tmp 'As String
    
    tmp = obtenerDeSQL("select cuenta from cajas where codigo = " & s2n(txtCaja))
    If Not IsEmpty(tmp) Then ' > "" Then
        verCajaEfectivo = True
        txtCuentaEfectivo = tmp
    Else
        che "No existe la caja"
        verCajaEfectivo = False
    End If
End Function

'        'Cheques
'        'daTaenvironment1.Sistema.Execute " insert into cheques " & _
'          "(nroint, fecha, nro, importe, ndoc, tdoc, dep_cuenta, fecha_ingr, estado, procedencia, Banco_Nro, Cliente, fecha_alta, usuario_alta, activo ) " & _
'          " values ( " & chCod & "," & ssFecha(chFech) & ", " & chNro & ", " & chMonto & ", " & s2n(txtNumero) & ", '" & lblTipo & "', " & chCta & ", " & _
'          ssFecha(dtFecha) & ", 'C', '" & chPT & "', " & bco & ", " & cliente.codigo & ", " & ssFecha(Date) & ", " & UsuarioActual() & ",1)"
'


'19/11/4 adapt licodigo, +cmd, +where
'2/12/4     control nro duplicado
'3/1/5      fix cajas
'

Private Sub uCtaBanco_cambio(codigo As Variant)

End Sub
