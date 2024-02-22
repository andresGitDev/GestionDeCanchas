VERSION 5.00
Begin VB.Form FrmRelacionProductoCliente 
   Caption         =   "Relacion Producto Cliente"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "FrmRelacionProductoCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAbm 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   2910
      Left            =   15
      TabIndex        =   8
      Top             =   15
      Width           =   8475
      Begin Gestion.ucCoDe uProd 
         Height          =   345
         Left            =   1545
         TabIndex        =   1
         Top             =   795
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   609
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uClie 
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   315
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   582
         CodigoWidth     =   1000
      End
      Begin VB.TextBox txtcodigocliente 
         Height          =   285
         Left            =   1650
         TabIndex        =   3
         Top             =   1530
         Width           =   1575
      End
      Begin VB.TextBox txtprecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6735
         TabIndex        =   6
         Top             =   1545
         Width           =   1455
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6735
         TabIndex        =   7
         Top             =   2025
         Width           =   1455
      End
      Begin VB.TextBox txtDestino 
         Height          =   285
         Left            =   1650
         TabIndex        =   4
         Top             =   1935
         Width           =   1575
      End
      Begin VB.TextBox txtLetra 
         Height          =   285
         Left            =   1665
         TabIndex        =   5
         Top             =   2370
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente :"
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
         Left            =   105
         TabIndex        =   15
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Producto :"
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
         Left            =   105
         TabIndex        =   14
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo Cliente :"
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
         Left            =   180
         TabIndex        =   13
         Top             =   1530
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Precio :"
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
         Left            =   5895
         TabIndex        =   12
         Top             =   1545
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   2685
         Left            =   75
         Top             =   165
         Width           =   8325
      End
      Begin VB.Label Label5 
         Caption         =   "Unidades x Bulto :"
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
         Left            =   5055
         TabIndex        =   11
         Top             =   2025
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Destino :"
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
         Left            =   225
         TabIndex        =   10
         Top             =   1950
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Letra :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2370
         Width           =   1305
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1560
      Left            =   0
      TabIndex        =   2
      Top             =   3090
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2752
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "FrmRelacionProductoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 12/8/4
'17/8/4 eliminar fix

'Dim rs As New ADODB.Recordset


'Private Sub cmdAceptar_Click()
''Dim Fecha As Variant
'
'   'Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'    If Ope <> "M" Then
'        DataEnvironment1.dbo_PRODUCTOCLIENTE "A", Trim(txtcodprod), Trim(txtcodcli), Trim(txtcodigocliente), s2n(txtprecio), Val(txtcantidad), Trim(txtDestino), Date, UsuarioSistema!codigo, 0, 0
'    Else
'        DataEnvironment1.dbo_PRODUCTOCLIENTE "M", Trim(txtcodprod), Trim(txtcodcli), Trim(txtcodigocliente), s2n(txtprecio), Val(txtcantidad), Trim(txtDestino), Date, 0, 0, 0
'        DataEnvironment1.dbo_GRABARBITACORA Val(txtcodcli), "Relacion_Producto_Cliente", Val(UsuarioSistema!codigo), Date, Time, "M"
'    End If
'    LimpioTxt
'    HabilitoTxt (True)
'End Sub


'Private Sub cmdayudacli_Click()
''    FrmHelp.Show
''    CargarHelp "Clientes", "Codigo", "Descripcion", "codigo", "descripcion"
''    FrmHelp.Tag = "CliRel"
'
'    Dim resu As String
'    resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
'    If resu > "" Then
''                grillahelp.Col = 0
'        txtcodcli = resu
''                grillahelp.Col = 1
'        TxtCliente = frmBuscar.resultado(2)
''        txtCodigo = resu
''        CargarDatos
''        Call HabilitoControles(
'    End If
'End Sub

'Private Sub cmdayudaprod_Click()
''    FrmHelp.Show
''    CargarHelp "Producto", "Codigo", "Descripcion", "codigo", "descripcion"
''    FrmHelp.Tag = "ProdRel"
'
'    Dim resu As String
'    resu = frmBuscar.MostrarCodigoDescripcionActivo("Producto")
'    If resu > "" Then
''                grillahelp.Col = 0
'        txtcodprod = resu
''                grillahelp.Col = 1
'        txtproducto = frmBuscar.resultado(2)
'        txtcodprod.SetFocus
''        txtCodigo = resu
''        CargarDatos
''        Call HabilitoControles(
'    End If
'
'End Sub

'Private Sub cmdBuscar_Click()
'    Dim resu As String, s As String
'    s = "SELECT r.PRODUCTO, r.CLIENTE, r.PRODUCTOCLIENTE, r.PRECIO, r.LETRA, r.UNIDADESXCAJA, r.DESTINO, p.descripcion, c.descripcion " _
'        & " FROM (Relacion_Producto_Cliente AS r LEFT JOIN Producto AS p ON r.PRODUCTO = p.codigo) LEFT JOIN Clientes AS c ON r.CLIENTE = c.codigo " _
'        & " where r.activo = 1 and c.activo = 1 and p.activo = 1 " _
'        & " order by r.cliente, r.producto "
'
'    With frmBuscar
'        resu = frmBuscar.MostrarSql(s)
'        If resu = "" Then Exit Sub
'
'        txtcodcli = .resultado(2)
'        txtcodprod = .resultado(1)
'        txtproducto = .resultado(8)
'        txtcliente = .resultado(9)
'        txtprecio = .resultado(4)
'        txtcodigocliente = .resultado(3)
'        txtcantidad = .resultado(6)
'        txtDestino = .resultado(7)
'    End With
'
'End Sub


'Private Sub cmdCancelar_Click()
'    LimpioTxt
'End Sub

'Private Sub cmdeliminar_Click()
'Dim RegistroEliminado As Integer ', Fecha
'
'    'Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'    If MsgBox("Esta seguro que desea borrar el registro actual ? ", vbYesNo + vbDefaultButton2 + vbQuestion, "Ventana de eliminacion") = vbYes Then
'
'        DataEnvironment1.dbo_PRODUCTOCLIENTE "B", Trim(txtcodprod), Val(Trim(txtcodcli)), "", 0, 0, "", 0, 0, Val(UsuarioSistema!codigo), Date
'        DataEnvironment1.dbo_GRABARBITACORA Val(txtcodcli), "Relacion_Producto_Cliente", Val(UsuarioSistema!codigo), Date, Time, "B"
'        MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
'        LimpioTxt
''Fallaba
'''        Call HabilitoControles(False, False, False, True, False, True)
''falla no existe
'        HabilitoTxt (True)
'    End If
'End Sub

'Private Sub cmdmodificar_Click()
'    Ope = "M"
'    HabilitoTxt (False)
'    txtprecio.SetFocus
'End Sub

Private Sub Form_Load()
'    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
'    LimpioTxt
    uClie.ini "select descripcion from clientes where codigo = '###' and activo = 1 ", "select codigo, descripcion as [ Cliente                            ] from clientes where activo = 1 "
    uProd.ini "select descripcion from producto where codigo = '###' and activo = 1", "select codigo as [codigo          ], descripcion as [Descripcion                                               ] from producto where activo = 1 order by codigo ", True
    uMenu.init True, True, True, False, True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub


Private Sub LimpioTxt()
'    txtcodcli = ""
'    txtcodprod = ""
'    txtproducto = ""
'    txtcliente = ""
'    txtprecio = "0.00"
'    txtcodigocliente = ""
'    txtcantidad = "0"
'    txtDestino = ""
    FrmBorrarTxt Me
    uClie.clear
    uProd.clear
End Sub
'Private Sub HabilitoTxt(habilito As Boolean) ' OJO BOLeano al reves
'    txtprecio.Locked = habilito
'    txtcantidad.Locked = habilito
'    txtcodigocliente.Locked = habilito
'    txtDestino.Locked = habilito
'End Sub
'
'Private Sub txtcodigocliente_Change()
'Dim i As Long
'
'    txtcodigocliente.Text = UCase(txtcodigocliente.Text)
'    i = Len(txtcodigocliente.Text)
'    txtcodigocliente.SelStart = i
'End Sub
''
'Private Sub txtcodigocliente_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtcodigocliente_GotFocus()
'
'    txtcodigocliente.SelStart = 0
'    txtcodigocliente.SelLength = Len(txtcodigocliente.Text)
'
'End Sub
Private Sub txtcantidad_GotFocus()

    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)

End Sub


'Private Sub CargoRegistro()
'    txtcodprod = rs!producto
'    txtproducto = ObtenerDescripcionS("producto", rs!producto)
'    txtcodcli = rs!cliente
'    txtcliente = ObtenerDescripcion("clientes", rs!cliente)
'    txtprecio = rs!precio
'
'    If Not IsNull(rs!productocliente) Then
'        txtcodigocliente = rs!productocliente
'    Else
'        txtcodigocliente = ""
'    End If
'End Sub
'Private Sub txtcodprod_LostFocus()
'
'    If Trim(txtcodprod) <> "" And Trim(txtcodcli) <> "" Then
'        rs.Open "select * from Relacion_Producto_Cliente where producto='" & Trim(txtcodprod) & "' and cliente=" & Val(txtcodcli), DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
'            MsgBox "Esa relacion ya esta establecida" ',solo puede modificarla", 48, "Atencion"
''            CargoRegistro
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'End Sub

Private Sub txtprecio_GotFocus()

    txtprecio.SelStart = 0
    txtprecio.SelLength = Len(txtprecio.Text)

End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub
Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub




Private Sub uClie_cambio(codigo As Variant)
    'UProd.Clear
    'FrmBorrarTxt Me
    CargoRegistro
End Sub


' *********** MENU LITO ************************
Private Sub uMenu_Aceptar()
    Dim tempo
    tempo = obtenerDeSQL("select * from relacion_Producto_Cliente where producto = '" & uProd.codigo & "'  and cliente = " & uClie.codigo)
    
    
    If IsEmpty(tempo) Then
        DataEnvironment1.dbo_PRODUCTOCLIENTE "A", uProd.codigo, uClie.codigo, Trim(txtcodigocliente), s2n(txtprecio, 4), Val(txtCantidad), Trim(txtDestino), Trim(txtLetra), Date, UsuarioSistema!codigo, 0, 0
        uMenu.AceptarOk
    Else
        DataEnvironment1.dbo_PRODUCTOCLIENTE "M", uProd.codigo, uClie.codigo, Trim(txtcodigocliente), s2n(txtprecio, 4), Val(txtCantidad), Trim(txtDestino), Trim(txtLetra), Date, 0, 0, 0
        uMenu.AceptarOk
    End If
    
End Sub

Private Sub uMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub uMenu_Buscar()
    Dim resu As String, s As String
    s = "SELECT r.PRODUCTO, r.CLIENTE, r.PRODUCTOCLIENTE, r.PRECIO, r.LETRA, r.UNIDADESXCAJA as [U x Caja], r.DESTINO, p.descripcion, c.descripcion " _
        & " FROM (Relacion_Producto_Cliente AS r LEFT JOIN Producto AS p ON r.PRODUCTO = p.codigo) LEFT JOIN Clientes AS c ON r.CLIENTE = c.codigo " _
        & " where r.activo = 1 and c.activo = 1 and p.activo = 1 " _
        & " order by r.cliente, r.producto "

    With frmBuscar
        resu = frmBuscar.MostrarSql(s, , , "")
        If resu = "" Then Exit Sub
        
        uClie.codigo = .resultado(2)
        uProd.codigo = .resultado(1)
        
        CargoRegistro
'        'txtcodprod = .resultado(1)
'        'txtcodcli = .resultado(2)

'        txtcodigocliente = .resultado(3)
'        txtprecio = .resultado(4)
'        txtLetra = .resultado(5)
'        txtcantidad = .resultado(6)
'        txtDestino = Trim(.resultado(7))
'        'txtproducto = .resultado(8)
'        'TxtCliente = .resultado(9)
        
        uMenu.BuscarOK
    End With
End Sub

Private Sub CargoRegistro()
    Dim tempo
    
    FrmBorrarTxt Me
    
    If uClie.codigo = 0 Or uProd.codigo = 0 Then Exit Sub
    
    tempo = obtenerDeSQL("select producto, cliente, productocliente, precio, letra, unidadesxcaja, destino, activo  from relacion_Producto_Cliente where producto = '" & uProd.codigo & "'  and cliente = " & uClie.codigo)
    If IsEmpty(tempo) Then Exit Sub
    
    ' entonces esta...
     
    txtcodigocliente = tempo(2)
    txtprecio = s2n(tempo(3), 4)
    txtLetra = sSinNull(tempo(4))
    txtCantidad = s2n(tempo(5))
    txtDestino = sSinNull(tempo(6))
    If Not tempo(7) Then che "Se recupera relacion eliminada"
End Sub

Private Sub uMenu_eliminar()
    DataEnvironment1.dbo_PRODUCTOCLIENTE "B", uProd.codigo, uClie.codigo, "", 0, 0, "", "", 0, 0, Val(UsuarioSistema!codigo), Date
    'DataEnvironment1.dbo_GRABARBITACORA Val(txtcodcli), "Relacion_Producto_Cliente", Val(UsuarioSistema!codigo), Date, Time, "B"
    grabaBitacora "B", uClie.codigo, "rel prod clie"
    MsgBox "La relacion fue borrada", 48, "Atencion"
    uMenu.EliminarOK
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    'HabilitoTxt Not sino  ' thats crazy
    fraAbm.enabled = sino
End Sub


Private Sub uMenu_Modificar()
    uClie.enabled = False
    uProd.enabled = False
End Sub
Private Sub uMenu_Nuevo()
    uClie.enabled = True
    uProd.enabled = True
End Sub


Private Sub uMenu_SALIR()
    Unload Me
End Sub

Private Sub uProd_cambio(codigo As Variant)
    Dim tempo
    tempo = obtenerDeSQL("select * from Relacion_Producto_Cliente where producto = '" & uProd.codigo & "' and cliente=" & uClie.codigo)
    If Not IsEmpty(tempo) Then
        CargoRegistro
    End If
End Sub

' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
' 14/2/6
'   ahora anda, FALTA control errores grabacion
