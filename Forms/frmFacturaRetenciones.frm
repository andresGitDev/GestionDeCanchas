VERSION 5.00
Begin VB.Form frmFacturaRetenciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Factura - Retenciones "
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8415
      Begin VB.TextBox txtImporteRetencion 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         TabIndex        =   6
         Top             =   2940
         Width           =   1755
      End
      Begin GestionWin.ucFecha uFecha 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   2940
         Width           =   1320
         _ExtentX        =   2884
         _ExtentY        =   661
         FechaInit       =   4
      End
      Begin VB.TextBox txtNroRetencion 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   1875
      End
      Begin GestionWin.ucCoDe Factura 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   900
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.ComboBox cboRetencion 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1980
         Width           =   1935
      End
      Begin GestionWin.ucCoDe cliente 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   300
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
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
         Left            =   6480
         TabIndex        =   20
         Top             =   0
         Width           =   915
      End
      Begin VB.Label txtCodigo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   19
         Top             =   0
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Retencion:"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   18
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Retencion:"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   17
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Retencion:"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   16
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Retencion:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sobre Factura:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   6180
         TabIndex        =   12
         Top             =   1380
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Saldo :"
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe :"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblImporteFactura 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   6180
         TabIndex        =   9
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente: "
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin GestionWin.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   4110
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   1508
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin GestionWin.ucEntreFechas ucFechas 
         Height          =   300
         Left            =   2520
         TabIndex        =   21
         Top             =   0
         Width           =   2580
         _ExtentX        =   4683
         _ExtentY        =   529
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entre:"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   22
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Retencion:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmFacturaRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 28/12/04
' Falta verificacion numeroRetencion vs cliente repetido
'
' ojo factura.codigo es Factura!NroFactura en tabla, no Factura!codigo
'     ERGO no alcanza nro para definir cual es
'

Private Const mTablaFV = "facturaventa"
Private Const INI_FAC_SEL = "SELECT TipoDoc FROM FacturaVenta  WHERE NroFactura= ###  AND Activo=1 AND CLIENTE = "
Private Const INI_FAC_BUS = "select NroFactura, TipoDoc from FacturaVenta where activo = 1 AND saldo > 0 AND cliente = "
Private TIPODOC As String
'



Private Sub cliente_cambio(codigo As Variant)
    If cliente.codigo = 0 Then Factura.codigo = 0
    Factura.ini INI_FAC_SEL & cliente.codigo, INI_FAC_BUS & cliente.codigo, False
End Sub


Private Sub Factura_cambio(codigo As Variant)
    Dim tmp
    lblImporteFactura = 0
    lblSaldo = 0
    If codigo > 0 Then
        tmp = obtenerDeSQL("select total, saldo from FacturaVenta where activo = 1 and cliente = " & cliente.codigo & " and NroFactura = " & Factura.codigo)
        If IsEmpty(tmp) Then
            ufa "", Me.Name ', Err
        Else
            lblImporteFactura = s2n(tmp(0))
            lblSaldo = s2n(tmp(1))
        End If
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True ', True
End Sub
Private Sub Form_Load()
    capFrm Me
 
    cliente.ini "select descripcion from clientes where codigo = ###", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    Factura.ini INI_FAC_SEL & cliente.codigo, INI_FAC_BUS & cliente.codigo, False
    
    'comboArray cboRetencion, Array("Ingresos brutos", "Iva", "Ganacias", "Bonos"), Array(1, 2, 3, 4)
    comboSql cboRetencion, "select Descripcion, id from TipoRetenciones "
    uMenu.init True, True, False, False, True
End Sub
Private Function TipoRet() As String
    TipoRet = obtenerDeSQL("select codigo from TipoRetenciones where id = " & ComboCodigo(cboRetencion))
'    Dim s As String
'
'    Select Case ComboCodigo(cboRetencion)
'    Case 1: s = "RIB"
'    Case 2: s = "RET"
'    Case 3: s = "RGA"
'    Case 4: s = "RBO"
'    End Select
'    TipoRet = s
End Function


Private Function TodoOk() As Boolean
    Dim tmp
    
    If s2n(txtImporteRetencion) > s2n(lblSaldo) Then
        che "importe supera saldo"
        Exit Function
    End If
    If cliente.codigo = 0 Then
        che "Falta cliente"
        Exit Function
    End If
    If Factura.codigo = 0 Then
        che "Falta Factura"
        Exit Function
    End If
    If s2n(txtImporteRetencion) = 0 Then
        che "falta importe"
        Exit Function
    End If
    If s2n(txtNroRetencion) = 0 Then
        che "falta nro retencion"
        Exit Function
    End If
    
    If TipoRet() = "" Then
        che "falta Tipo Retencion "
        Exit Function
    End If
    
    'Verificar si esta,
    'luego
    


    'If Not confirma("Ya existe ese numero para este cliente" & vbCrLf & " Continua?") Then Exit Sub
    
    TodoOk = True
End Function

Private Function GrabaRet() As Boolean
    If MODO_ON_ERROR_ABM_ON Then On Error GoTo UfaGraba
    
    Dim codRet, cli, provi, Cuit, TIVA, tmp
    Dim codFac
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    cli = cliente.codigo
    
    tmp = obtenerDeSQL("select provincia, cuit, iva from clientes where activo = 1 and codigo = " & cli)
    provi = sSinNull(tmp(0)): If provi = "*" Then provi = " "
    Cuit = sSinNull(tmp(1))
    TIVA = nSinNull(tmp(2))
    
    codFac = obtenerDeSQL("select codigo from facturaventa where activo = 1 and NroFactura = " & Factura.codigo & " and cliente = " & cli)
    
    rs2.Open "select * from Facturaventa where nrofactura=" & Factura.codigo & " and tipodoc='" & Trim(Factura.descripcion) & "'", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    rs.Open "select * from FacturaRetencion where tdoc_ret='" & TipoRet() & "' and codfactura=" & rs2!codigo, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    If rs.EOF = True And rs.BOF = True Then
    Else
        If MsgBox("La retencion para la factura seleccionada ya esta cargada." & Chr(13) & "¿Desea continuar de todas formas?", vbQuestion + vbYesNo, "ATENCION") = vbNo Then
            Set rs = Nothing
            Set rs2 = Nothing
            Exit Function
        End If
        
        'Factura.SetFocus
        'Exit Function
    End If
    Set rs = Nothing
    Set rs2 = Nothing

'transaccion aqui -------------------INI
    DE_BeginTrans
    
    codRet = nuevoCodigo("FacturaVenta", "codigo")
    
    DataEnvironment1.dbo_abmFacturaVenta "A", codRet, TipoRet(), Trim(txtNroRetencion), uFecha.dtfecha, Date, 0, 0, cli, cliente.descripcion, provi, Cuit, TIVA, 0, 0, 0, 0, s2n(txtImporteRetencion), 0, 0, 0, 0, UsuarioActual(), Date, 1, 1, 0, 0, 0, 0, 0, 0, 0
    DataEnvironment1.AMR.Execute "insert into FacturaRetencion (CodFactura, tDoc_ret, nDoc_Ret)  values ( " & codFac & " ,'" & TipoRet() & "', " & s2n(txtNroRetencion) & " ) "
    DataEnvironment1.AMR.Execute "update FacturaVenta set saldo = saldo - " & x2s(txtImporteRetencion) & " where codigo = " & codFac
    rs.Open "select max(id) as max from FacturaRetencion", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    rs2.Open "select max(codigo) as max2 from Facturaventa", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    DataEnvironment1.AMR.Execute "update FacturaRetencion set _codretencion = " & rs2!max2 & " where id = " & rs!Max
    DE_CommitTrans
'transaccion aqui -------------------FIN
    
    Set rs = Nothing
    Set rs2 = Nothing
    GrabaRet = True
fin:
    Exit Function
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar", Me.Name & " grabadet " ', Err
    Resume fin
End Function


'///////////////////////////////// MENU //////////////
Private Sub uMenu_AceptarAlta()
    If Not TodoOk() Then Exit Sub
    If GrabaRet() Then
        che "operacion concluidda"
        uMenu.AceptarOk
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarCbo Me
    FrmBorrarTxt Me
    cliente.codigo = 0
    
    
    cboRetencion.ListIndex = -1
End Sub
Private Sub uMenu_Buscar()
    '    On Error GoTo fin
    Dim re As Variant, WhereTipo As String, WhereFecha As String, tmp, tmp1
    Dim rs As New ADODB.Recordset
    're = frmBuscar.mostrarSql("select codigo, NroFactura, Cliente, Fecha  from " & mTablaFV & " where fecha " & ssBetween(dtDesde, dtHasta) & " order by fecha desc ")
    
     WhereTipo = " (TipoDoc = 'RIB' or TipoDoc = 'RGA' or TipoDoc = 'RET' or TipoDoc = 'RBO' or TipoDoc = 'RCP') "
    'WhereTipo = IIf(optBuscarTipo.Item(0).Value, " (TipoDoc = 'FAA' or TipoDoc = 'NCA' or TipoDoc = 'NDA') ", "(TipoDoc = 'FAB')")
    
    WhereFecha = "fecha " & ucFechas.ssBetween()
    
    With frmBuscar
        re = .MostrarSql("select f.Codigo as Codigo, TipoDoc, NroFactura, Cliente, c.descripcion as [ Nombre                        ], Fecha as [Fecha ], f.activo as Activa, f.total as [Importe  ]  from " & mTablaFV & " as f left join clientes as c on c.codigo = f.cliente where " & WhereTipo & " and " & WhereFecha & " order by NroFactura desc ", , , , "", "Anulada")
        If re = "" Then Exit Sub
        
        rs.Open "select * from FacturaRetencion where _codretencion=" & .resultado(1), DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If rs.EOF = True And rs.BOF = True Then 'para mostrar lo anterior a la modificacion y que estaba muy mal!!
            txtcodigo = .resultado(1)
            TIPODOC = .resultado(2)
            txtNroRetencion = .resultado(3)
            cliente.codigo = .resultado(4)
            uFecha.dtfecha CDate(.resultado(6))
            txtImporteRetencion = .resultado(8)
                
            cboRetencion.ListIndex = BuscarenComboS(cboRetencion, ObtenerDescripcionS("TipoRetenciones", TIPODOC))
            tmp = obtenerDeSQL("select codfactura from FacturaRetencion where TDoc_Ret = '" & TIPODOC & "' and Ndoc_Ret = " & s2n(txtNroRetencion))
                
            lblImporteFactura = 0
            lblSaldo = 0
            tmp1 = obtenerDeSQL("select total, saldo, nrofactura from FacturaVenta where activo = 1 and cliente = " & cliente.codigo & " and Codigo = " & tmp)
            If Not IsEmpty(tmp1) Then
                Factura.codigo = s2n(tmp1(2))
                lblImporteFactura = s2n(tmp1(0))
                lblSaldo = s2n(tmp1(1))
            End If
        Else 'esto es para mostrar bien los datos a partir del 22/02/08
            txtcodigo = .resultado(1)
            TIPODOC = .resultado(2)
            txtNroRetencion = .resultado(3)
            cliente.codigo = .resultado(4)
            uFecha.dtfecha CDate(.resultado(6))
            txtImporteRetencion = .resultado(8)
                
            cboRetencion.ListIndex = BuscarenComboS(cboRetencion, ObtenerDescripcionS("TipoRetenciones", TIPODOC))
            tmp = obtenerDeSQL("select codfactura from FacturaRetencion where TDoc_Ret = '" & Trim(TIPODOC) & "' and Ndoc_Ret = " & s2n(txtNroRetencion) & " and _codretencion=" & txtcodigo)
                
            lblImporteFactura = 0
            lblSaldo = 0
            'tmp1 = obtenerDeSQL("select total, saldo, nrofactura from FacturaVenta where activo = 0 and cliente = " & cliente.codigo & " and Codigo = " & tmp)
            tmp1 = obtenerDeSQL("select total, saldo, nrofactura from FacturaVenta where cliente = " & cliente.codigo & " and Codigo = " & tmp)
            If Not IsEmpty(tmp1) Then
                Factura.codigo = s2n(tmp1(2))
                lblImporteFactura = s2n(tmp1(0))
                lblSaldo = s2n(tmp1(1))
            End If

        End If
        Set rs = Nothing
    ''End If
        'mFAE = (.resultado(2) = "FAE")
        'lblExterior.Visible = mFAE
            
    End With
    'CargaDatos
    
    
    'gO.Borrar
    uMenu.BuscarOK
fin:
End Sub
Private Sub uMenu_eliminar()
    'If MODO_ON_ERROR_ABM_ON Then On Error GoTo UfaGraba
    
    Dim codRet, cli, provi, Cuit, TIVA, tmp
    Dim codFac
    
    cli = cliente.codigo
    
    'tmp = obtenerDeSQL("select provincia, cuit, iva from clientes where activo = 1 and codigo = " & cli)
    'provi = sSinNull(tmp(0)): If provi = "*" Then provi = " "
    'Cuit = sSinNull(tmp(1))
    'TIVA = nSinNull(tmp(2))
    
    codFac = obtenerDeSQL("select codigo from facturaventa where activo = 1 and NroFactura = " & Factura.codigo & " and cliente = " & cli)

'transaccion aqui -------------------INI
    DE_BeginTrans
    
        codRet = nuevoCodigo("FacturaVenta", "codigo")
        
        'DataEnvironment1.dbo_abmFacturaVenta "A", codRet, TipoRet(), Trim(txtNroRetencion), uFecha.dtfecha, Date, 0, 0, cli, cliente.descripcion, provi, Cuit, TIVA, 0, 0, 0, 0, s2n(txtImporteRetencion), 0, 0, 0, 0, UsuarioActual(), Date, 1, 1, 0, 0, 0, 0, 0, 0, 0
        'DataEnvironment1.AMR.Execute "insert into FacturaRetencion (CodFactura, tDoc_ret, nDoc_Ret)  values ( " & codFac & " ,'" & TipoRet() & "', " & s2n(txtNroRetencion) & " ) "
        DataEnvironment1.AMR.Execute "update FacturaVenta set saldo = saldo + " & x2s(txtImporteRetencion) & " where codigo = " & codFac
        DataEnvironment1.AMR.Execute "delete from FacturaRetencion where codfactura = " & codFac & " and tdoc_ret = '" & Trim(TIPODOC) & "' "
        DataEnvironment1.AMR.Execute "delete from FacturaVenta where tipodoc = '" & Trim(TIPODOC) & "' and  Nrofactura = " & x2s(txtNroRetencion)
    
    DE_CommitTrans
'transaccion aqui -------------------FIN
    che "eliminada"
    uMenu.EliminarOK
    'GrabaRet = True
fin:
    Exit Sub
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar", Me.Name & " grabadet " ', Err
    Resume fin
End Sub

Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fra.Enabled = sino
End Sub
Private Sub uMenu_Nuevo()
    cliente.SetFocus
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
'///////////////////////////////// MENU //////////////


'28/12/4 start & end
'7/4/5  x2s() en execute, que verguenza
'7/6/5 Cambio grosito: codigos pasan a tabla de programador TipoRetenciones
'


