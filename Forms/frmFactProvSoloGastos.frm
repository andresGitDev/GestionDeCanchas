VERSION 5.00
Begin VB.Form frmFactProvSoloGastos 
   Caption         =   "Ingreso de Gastos"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11340
   Icon            =   "frmFactProvSoloGastos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComprobante 
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   1140
      Width           =   1995
   End
   Begin VB.TextBox txtConcepto 
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   5835
   End
   Begin VB.Frame fraDetalle 
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   0
      TabIndex        =   15
      Top             =   1620
      Width           =   11295
      Begin Gestion.ucCheques uCheques 
         Height          =   2415
         Left            =   2760
         TabIndex        =   10
         Top             =   480
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   4260
      End
      Begin VB.TextBox txtCuenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5100
         TabIndex        =   14
         Tag             =   "2"
         Top             =   2520
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton cmdCuenta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta"
         Enabled         =   0   'False
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtCodCuenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtimpcheques 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtefectivo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Top             =   60
         Width           =   1215
      End
      Begin VB.TextBox txttransf 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   795
         TabIndex        =   11
         Top             =   2925
         Visible         =   0   'False
         Width           =   1215
      End
      Begin Gestion.ucCoDe uCajaEfectivo 
         Height          =   315
         Left            =   2760
         TabIndex        =   8
         Top             =   60
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Caja"
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
         Index           =   1
         Left            =   2220
         TabIndex        =   20
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
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
         Left            =   0
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo"
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
         Left            =   60
         TabIndex        =   18
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Transf."
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
         Left            =   60
         TabIndex        =   17
         Top             =   2925
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Index           =   0
         Left            =   2010
         TabIndex        =   16
         Top             =   2925
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   5070
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   2831
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin Gestion.ucFecha uFecha 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      FechaInit       =   4
   End
   Begin Gestion.ucCoDe uTipoCompra 
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Comprobante"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   24
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   23
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8940
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Total:"
      Height          =   255
      Index           =   2
      Left            =   8460
      TabIndex        =   22
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmFactProvSoloGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCuenta_Click()
    Dim re
    re = frmBuscar.MostrarSql("SELECT CTASBANK.CODIGO AS [ Codigo   ], CTASBANK.BANCO AS [Cod Banco ], BancosGrales.descripcion AS [ Nombre Banco                             ], CTASBANK.NUMERO, Monedas.descripcion AS [ Moneda        ] FROM CTASBANK LEFT OUTER JOIN                     BancosGrales ON CTASBANK.BANCO = BancosGrales.codigo LEFT OUTER JOIN                      Monedas ON Monedas.codigo = CTASBANK.MONEDA Where (CTASBANK.ACTIVO = 1)")
    If re > "" Then
        txtcodcuenta = frmBuscar.resultado(1)
        txtcuenta = frmBuscar.resultado(3) & " - " & frmBuscar.resultado(4)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    'uTipoCompra.ini "Select descripcion from CuentasParam where codigo = '###' and sistema = 0 and activo = 1", "select codigo, descripcion , cuenta from CuentasParam where sistema = 0 and activo = 1", True
    uTipoCompra.ini "select descripcion from Cuentas where cuenta = '###' and activo = 1", "select cuenta  as [ Cuenta       ], descripcion as [ Descripcion               ] from Cuentas where activo = 1 ", True
    uMenu.init True, True, False, False, True
    uCajaEfectivo.ini "select Responsable from cajas where activo = 1 and codigo = ###", "select Codigo as [ Codigo   ], responsable as [  Responsable      ]  from cajas where activo = 1"
End Sub

Private Function TaTodo() As Boolean

    If Not PuedoCompras(uFecha.dtFecha) Then
        'msg dentro de las funcion
        Exit Function
    End If
    
    If uTipoCompra.DESCRIPCION = "" Then
        che "Falta tipo compra"
        Exit Function
    End If
    
    If s2n(lbltotal) = 0 Then
        che "falta importe de gastos"
        Exit Function
    End If
    
    If s2n(txtefectivo) > 0 And uCajaEfectivo.DESCRIPCION = "" Then
        che "Falta caja efectivo"
        Exit Function
    End If
    
    If s2n(txttransf) > 0 And txtcuenta = "" Then
        che "falta cuenta transferencia "
        Exit Function
    End If
    
    TaTodo = True
End Function

Private Sub recalcular()
    txtimpcheques = s2n(uCheques.Total)
    lbltotal = s2n(s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf), 2)
End Sub

Private Sub txtEfectivo_LostFocus()
    recalcular
End Sub

Private Function GrabaOk() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim asiCom As New Asiento, iddoc As Long
    Dim efec As Double, cuentaefec As String
    Dim tran As Double, cuentatran As String
    Dim i As Long, Total As Double, NroDoc As Long, maxCaja As Long
    

   
    
    Total = s2n(lbltotal)
    'NroDoc = nuevoCodigo("Compras", "nrodoc", "tipodoc = 'FCG'") '"codpr = 0 ")
    NroDoc = nuevoCodigo("RegistroDocumentos", "nrodoc", "tipodoc = '" & TIPODOC_FAC_PROVGASTO & "' ") '"codpr = 0 ")
    efec = s2n(txtefectivo)
    cuentaefec = sSinNull(obtenerDeSQL("select cuenta from Cajas where codigo = '" & (uCajaEfectivo.codigo) & "' "))
    tran = s2n(txttransf)
    cuentatran = sSinNull(obtenerDeSQL("select cuenta_con from Ctasbank where codigo = " & Val(txtcodcuenta) & " "))
    
    DE_BeginTrans
        iddoc = NuevoDocumento(TIPODOC_FAC_PROVGASTO, NroDoc, 0, NuevoNroPago()) '   si hay ret:  NuevoNroCertifGan(), NuevoNroCertifIIBB()  no hecho aca
        'NroDoc = iddoc
'        If iddoc = 0 Then
        
        
        DataEnvironment1.dbo_INGCOMPRAS "A", uFecha.dtFecha, 0, 0, _
                0, "", "", 1, TIPODOC_FAC_PROVGASTO, NroDoc, _
                  0, txtcodcuenta, Total, Total, 0, 0, 0, 0, _
                0, 0, 0, 0, 0, 0, 0, _
                1, 0, 0, 0, 0, 0, _
                0, Date, UsuarioSistema!codigo, iddoc, 0, 0, _
                "", "", 0, 0
        'utipocompra.codigo no lo pude grabar, se pierde
   
        asiCom.nuevo Trim(txtconcepto), uFecha.dtFecha, TIPODOC_FAC_PROVGASTO
        
        'asiCom.AgregarItem CuentaParamxCodigo(uTipoCompra.codigo), Total, 0
        asiCom.AgregarItem (uTipoCompra.codigo), Total, 0, txtComprobante
        
        'EFECTIVO
        If efec > 0 Then
            'asiento
            asiCom.AcumularItem cuentaefec, 0, efec
            'movicaja
            maxCaja = nuevoCodigo("movicaja", "movimiento")
            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", uCajaEfectivo.codigo, maxCaja, "E", "E", efec, "Gasto " & iddoc _
                , uFecha.dtFecha, 0, 0, TIPODOC_FAC_PROVGASTO, NroDoc, cuentaefec, 0 _
                , iddoc, Date, UsuarioSistema!codigo, 1
        End If
        
'        'TRANSFERENCIA
'        If tran > 0 Then
'            'Asiento
'            asiCom.AcumularItem cuentatran, 0, tran
'            'movicaja
'            maxCaja = nuevoCodigo("movicaja", "movimiento")
'            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", Val(txtcodcaja), maximocaja, "T", "E", s2n(txttransf), "Fac. " & txtfact & "Prov. " & uProv.codigo, _
'                dtfecha, 0, uProv.codigo, TIPODOC_FAC_PROVGASTO, Val(txtfact), valorcuentacon, maximobanc, _
'                Date, UsuarioSistema!codigo, 0, 0, 1
'        End If
        
        Dim x As Long, cuent As String
        'CHQ PROPIO
        If ExistenPropios() Then
            'movicaja
            maxCaja = nuevoCodigo("movicaja", "movimiento")
            
            For x = 1 To uCheques.rows
                If uCheques.chPropio(x) Then
                
                    'cuent = sSinNull(obtenerDeSQL("select cuenta from ctasBank where  codigo = '" & uCheques.chBancCod(x) & "' and activo = 1")) ', 0, uCheques.chMonto(x)
                    cuent = uCheques.chCuenta(x)
                    DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maxCaja, "P", "E", uCheques.chMonto(x), "Gasto  " & NroDoc _
                        , uFecha.dtFecha, uCheques.chNroInt(x), 0, TIPODOC_FAC_PROVGASTO, NroDoc, cuent, 0 _
                        , iddoc, Date, UsuarioSistema!codigo, 1
'
                    
                    DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "A", uCheques.chNroInt(x), uCheques.chFecha(x), uCheques.chMonto(x) _
                        , NroDoc, TIPODOC_FAC_PROVGASTO, 0, "T", uCheques.chFecha(x), uFecha.dtFecha, Date, UsuarioSistema!codigo, 0, 0, 1, 1, 1
'
                    'haber CHEQUE PROPIO
                    asiCom.AcumularItem cuent, 0, uCheques.chMonto(x), "ch " & uCheques.chNumero(x)
                End If
            Next
        End If
        
        'CHQ 3ROS
        If ExistenTerceros() Then
            'movicaja
            maxCaja = nuevoCodigo("movicaja", "movimiento")
            'cuent = CuentaParam(ID_CuentasParam_CH_CARTERA)
            cuent = uCheques.chCuenta(x)
            For x = 1 To uCheques.rows
                If Not uCheques.chPropio(x) Then
                    DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maxCaja, "C", "E", uCheques.chMonto(x), "Gasto " & NroDoc _
                        , uFecha.dtFecha, uCheques.chNroInt(x), 0, TIPODOC_FAC_PROVGASTO, NroDoc, cuent, 0 _
                        , iddoc, Date, UsuarioSistema!codigo, 1

                    DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "A", uCheques.chNroInt(x), 0, "", 0, NroDoc, 0 _
                        , uFecha.dtFecha, "T", "FDC", Date, UsuarioSistema!codigo, 0, 0, 1, 1, 1
'
'                                'INCREMENTO EL AUTOMATICO DE MOVIBANC
'                                maximobanc = maximobanc + 1
'
'                                sAssert = "9b) dbo_INGCOMPRAMOVIBANC - Asiento"
                    'Haber Cheques 3ros
                    asiCom.AcumularItem cuent, 0, uCheques.chMonto(x), "ch3 " & uCheques.chNumero(x)
                End If
            Next
        End If
        
        If asiCom.Grabar(iddoc) = 0 Then
            DE_RollbackTrans
            GoTo fin ' sorry
        End If
        
    DE_CommitTrans
    GrabaOk = True
    
fin:
    Exit Function
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar ", ""
    Resume fin
End Function


'--------------------------------------

Private Function ExistenTerceros() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If Not uCheques.chPropio(x) Then
            ExistenTerceros = True
            Exit Function
        End If
    Next x
End Function
Private Function ExistenPropios() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If uCheques.chPropio(x) = True Then
            ExistenPropios = True
            Exit Function
        End If
    Next x
End Function

Private Sub uCheques_LostFocus()
    recalcular
End Sub

'**************************************
Private Sub uMenu_AceptarAlta()
    If Not TaTodo() Then Exit Sub
    If GrabaOk() Then
        che "Operacion ok"
        uMenu.AceptarOk
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    uFecha.dtFecha Date
    uTipoCompra.clear
    uCajaEfectivo.clear
    uCheques.Borrar
End Sub
Private Sub uMenu_Buscar()
    Dim resu
    Dim ssql
    'sSql = "select Fecha as [Fecha ], tipoDoc as [Doc], NroDoc as [ Numero ], total as  [ Importe ], codPr as [ Prov ], descripcion as [ Razon social                           ] from compras inner join prov on codpr = prov.codigo  where compras.activo = 1 and tipodoc = '" & TIPODOC_FAC_PROVGASTO & "' and contado = 1 and codpr = 0 order by fecha desc "
    ssql = "select Fecha as [Fecha ], tipoDoc as [Doc], NroDoc as [ Numero ], total as  [ Importe ] from compras  where compras.activo = 1 and tipodoc = '" & TIPODOC_FAC_PROVGASTO & "' order by fecha desc "
    resu = frmBuscar.MostrarSql(ssql)
'    If resu <> "" Then cargagasto resu(3)
'    End If
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    uFecha.enabled = sino
    fraDetalle.enabled = sino
    txtconcepto.enabled = sino
    'uFecha.enabled = sino
    uTipoCompra.enabled = sino
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
'*************************************
