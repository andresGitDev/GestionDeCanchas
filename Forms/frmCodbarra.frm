VERSION 5.00
Begin VB.Form frmCodbarra 
   Caption         =   "Asignacion de Codigos de barra"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   Icon            =   "frmCodbarra.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin Gestion.ucCoDe uProd 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   6495
      _extentx        =   11456
      _extenty        =   503
      codigoinvalido  =   0
      codigowidth     =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Producto :"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo de barra :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCodbarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim sql As String
    
    If uProd.codigo = "" Then
        MsgBox "Debe ingresar el producto para asignarle el codigo de barra."
        Exit Sub
    End If
    
    sql = "update producto set codigobarra='" & Text1.Text & "' where codigo='" & uProd.codigo & "'"
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se ha realizado con exito la carga."
    Text1.Text = ""
    
    Unload Me
    frmColector.cmdGuardar.Value = True
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = datobarra
    set_uProd
End Sub

Private Sub set_uProd() ' lo copie de pedido cliente
    Dim sqlbuscar As String, sqldesc As String

'    If Propio() Then    'propio
        sqldesc = "select descripcion from producto where codigo = '###' "
        sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
'    Else    'relCliente
'        sqldesc = "select descripcion from producto  " _
'            & " inner join relacion_Producto_Cliente " _
'            & " on producto.codigo = relacion_Producto_cliente.producto " _
'            & " where cliente = " & cliente.codigo & " and productoCliente = '###'"
'        sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
'            & " from producto  " _
'            & " inner join relacion_Producto_Cliente " _
'            & " on producto.codigo = relacion_Producto_cliente.producto " _
'            & " where cliente = " & cliente.codigo _
'            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
'            & " order by producto"
'    End If
    uProd.ini sqldesc, sqlbuscar, True
    uProd.EditaDescripcion = True
End Sub

