VERSION 5.00
Begin VB.Form FrmRelacionProductoProveedor 
   Caption         =   "Relacion Producto Proveedor"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "FrmRelacionProductoProveedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      Height          =   3645
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   8745
      Begin VB.TextBox txtprecio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6150
         TabIndex        =   9
         Top             =   1740
         Width           =   1455
      End
      Begin VB.TextBox txtcodigoprov 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2985
         TabIndex        =   8
         Top             =   1740
         Width           =   1815
      End
      Begin VB.TextBox txtporcentaje 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1425
         TabIndex        =   7
         Top             =   2175
         Width           =   855
      End
      Begin VB.TextBox txtcaja 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3945
         TabIndex        =   6
         Top             =   2175
         Width           =   855
      End
      Begin VB.TextBox txtpedmin 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6705
         TabIndex        =   5
         Top             =   2175
         Width           =   855
      End
      Begin VB.TextBox txtcotizacion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1425
         TabIndex        =   4
         Top             =   2655
         Width           =   855
      End
      Begin Gestion.ucCoDe uProd 
         Height          =   330
         Left            =   1215
         TabIndex        =   2
         Top             =   1200
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   582
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uProv 
         Height          =   345
         Left            =   1215
         TabIndex        =   3
         Top             =   765
         Width           =   7155
         _ExtentX        =   11192
         _ExtentY        =   609
         CodigoWidth     =   1000
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
         Left            =   5310
         TabIndex        =   18
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo Producto Proveedor :"
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
         TabIndex        =   17
         Top             =   1740
         Width           =   2655
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
         Left            =   150
         TabIndex        =   16
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Proveedor :"
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
         Left            =   150
         TabIndex        =   15
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje :"
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
         TabIndex        =   14
         Top             =   2175
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "CantidadxCaja :"
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
         Left            =   2505
         TabIndex        =   13
         Top             =   2175
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Pedido Minimo :"
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
         Left            =   5145
         TabIndex        =   12
         Top             =   2175
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Cotizacion :"
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
         TabIndex        =   11
         Top             =   2655
         Width           =   1095
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7065
         TabIndex        =   10
         Top             =   255
         Width           =   1230
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   3900
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   2540
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "FrmRelacionProductoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 12/8/4
'17/8/4 eliminar fix



'Private Sub cmdAceptar_Click()
'Dim fecha As Variant
'If Trim(txtProveedor) <> "" And Trim(txtproducto) <> "" Then
'    fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'    If Ope <> "M" Then
'        DataEnvironment1.dbo_PRODUCTOPROVEEDOR "A", Trim(txtcodprod), Trim(TxtCodProv), Trim(txtcodigoprov), Replace(s2n(txtprecio), ".", ","), Replace(s2n(txtporcentaje), ".", ","), Replace(s2n(txtcaja), ".", ","), Replace(s2n(txtpedmin), ".", ","), Replace(s2n(txtcotizacion), ".", ","), fecha
'    Else
        
'        DataEnvironment1.dbo_PRODUCTOPROVEEDOR "M", Trim(txtcodprod), Trim(TxtCodProv), Trim(txtcodigoprov), Replace(s2n(txtprecio), ".", ","), Replace(s2n(txtporcentaje), ".", ","), Replace(s2n(txtcaja), ".", ","), Replace(s2n(txtpedmin), ".", ","), Replace(s2n(txtcotizacion), ".", ","), fecha
'        DataEnvironment1.dbo_GRABARBITACORA Val(TxtCodProv), "Relacion_Producto_Proveedor", Val(UsuarioSistema!codigo), fecha, Time, "M"
'    End If
'    LimpioTxt
'Else
'    MsgBox "Debe cargar datos correctos para aceptar", 48, "Atencion"
'End If
'    HabilitoTxt (True)
'End Sub

Private Sub Form_Load()

    UpROV.ini " select descripcion from prov where codigo = '###' and activo = 1", "select codigo, descripcion as [ Nombre proveedor            ] from prov where activo = 1"
    uProd.ini " select descripcion from producto where codigo = '###' and activo = 1", "select codigo as [ Codigo            ] , descripcion as [ Descripcion             ] from producto  where activo = 1 ", True
    uMenu.MsgConfirmaEliminar = "Elimina relacion"
    uMenu.init True, True, True, False, True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub txtcodigoprov_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtporcentaje_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtpedmin_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcotizacion_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcaja_GotFocus()
    PintoFocoActivo
End Sub

Private Sub CargoRegistro()
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "select * from relacion_producto_proveedor where producto = '" & uProd.codigo & "' and proveedor = " & UpROV.codigo, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        If Not .EOF Then
            lblId = s2n(!ID)
            txtprecio = nSinNull(!precio)
            txtcaja = nSinNull(!caja)
            txtporcentaje = nSinNull(!PORCENTAJE)
            txtpedmin = nSinNull(!pedidominimo)
            txtCotizacion = nSinNull(!cotizacion)
            txtcodigoprov = sSinNull(!CodigoProveedor)
        End If
        uMenu.BuscarOK
    End With
    Set rs = Nothing
        
'    txtcodprod = rs!producto
'    txtproducto = ObtenerDescripcionS("producto", rs!producto)
'    TxtCodProv = rs!Proveedor
'    txtProveedor = ObtenerDescripcion("prov", rs!Proveedor)
        '>
'    If Not IsNull(rs!caja) Then
'        txtcaja = rs!caja
'    Else
'        txtcaja = "0"
'    End If
        
    
    'If Not IsNull(rs!PORCENTAJE) Then
        
    'Else
    '    txtporcentaje = "0"
    'End If
    'if Not IsNull(rs!pedidominimo) Then
        
    'Else
    '    txtpedmin = "0"
    'End If
    'If Not IsNull(rs!cotizacion) Then

   ' Else
    '    txtcotizacion = "0"
   ' End If
    'If Not IsNull(rs!CodigoProveedor) Then

    'Else
    '    txtcodigoprov = ""
    'End If
    
End Sub


Private Sub uProd_cambio(codigo As Variant)
    Dim tempo
    tempo = obtenerDeSQL("select * from Relacion_Producto_Proveedor where producto = '" & uProd.codigo & "' and proveedor =" & UpROV.codigo)
    If Not IsEmpty(tempo) Then
        If uMenu.estado = ucbEditando Then MsgBox "relacion existente"
        CargoRegistro
    End If
End Sub
Private Sub uProv_cambio(codigo As Variant)
    Dim tempo
    tempo = obtenerDeSQL("select * from Relacion_Producto_Proveedor where producto = '" & uProd.codigo & "' and proveedor =" & UpROV.codigo)
    If Not IsEmpty(tempo) Then
        If uMenu.estado = ucbEditando Then MsgBox "relacion existente"
        CargoRegistro
    End If
End Sub


'Private Sub txtcodprod_LostFocus()
'
'    If Trim(txtcodprod) <> "" And Trim(TxtCodProv) <> "" Then
'        rs.Open "select * from Relacion_Producto_proveedor where producto='" & Trim(txtcodprod) & "' and proveedor=" & Val(TxtCodProv), DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
'            MsgBox "Esa relacion ya esta establecida,solo puede modificarla", 48, "Atencion"
'            CargoRegistro
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'End Sub


Private Sub uMenu_Aceptar()
    Dim s As String
    
    If uProd.codigo = "" Or UpROV.codigo = 0 Then
        che "faltan datos"
        Exit Sub
    End If
    
    If s2n(lblId) = 0 Then
        DataEnvironment1.dbo_PRODUCTOPROVEEDOR "A", _
            uProd.codigo, UpROV.codigo, Trim(txtcodigoprov), s2n(txtprecio, 4), _
            s2n(txtporcentaje), s2n(txtcaja), s2n(txtpedmin), s2n(txtCotizacion), Date
            
        che "Nueva relacion grabada"
        
    Else
        s = "update relacion_producto_proveedor set " & _
           " producto = '" & uProd.codigo & "' , codigoproveedor = '" & Trim(txtcodigoprov) & "', " & _
           " porcentaje = " & ssNum(txtporcentaje) & ", " & _
           " precio = " & ssNum(txtprecio) & ", caja = " & ssNum(txtcaja) & ", " & _
           " pedidominimo = " & ssNum(txtpedmin) & ", " & _
           " fechacarga = " & ssFecha(Date) & ", cotizacion = " & ssNum(txtCotizacion) & _
           " where id = " & ssNum(lblId)
        DataEnvironment1.Sistema.Execute s
        grabaBitacora "M", s2n(lblId), "Rel_prod_prov"
        
        che "modificado"
    End If
    
    uMenu.AceptarOk
End Sub

Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    uProd.clear
    UpROV.clear
    lblId = ""
End Sub
Private Sub uMenu_Buscar()
    Dim s
    
    s = "select Producto as [ Producto        ], proveedor as [ Prov ], CodigoProveedor  as [ Producto Proveedor ] from relacion_producto_proveedor"
    If UpROV.codigo > 0 Then
        s = s & " where proveedor = " & UpROV.codigo
    End If
    If frmBuscar.MostrarSql(s) = "" Then Exit Sub
    
    UpROV.codigo = frmBuscar.resultado(2)
    uProd.codigo = frmBuscar.resultado(1)
    
    uMenu.BuscarOK
End Sub

Private Sub uMenu_eliminar()
    Dim s
    If s2n(lblId) = 0 Then
        ufa "prg: sin registro a eliminar", "rel prod prov  id=0"
        Exit Sub
    End If
    s = "delete from relacion_producto_proveedor where id = " & s2n(lblId)
    DataEnvironment1.Sistema.Execute s
    che "eliminado"
    uMenu.EliminarOK
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fra.enabled = sino
End Sub
Private Sub uMenu_Modificar()
    UpROV.enabled = False
    uProd.enabled = False
End Sub
Private Sub uMenu_Nuevo()
    UpROV.enabled = True
    uProd.enabled = True
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub

' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR

