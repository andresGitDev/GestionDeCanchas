VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmListadoPreciosHistorico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Precios Historicos"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "FrmListadoPreciosHistorico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEntreFechas 
      Caption         =   "Elija la Fecha de Comienzo y Fin del Informe "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   4335
      Begin Gestion.ucEntreFechas ucEntreFechas 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3855
         _ExtentX        =   4895
         _ExtentY        =   661
      End
   End
   Begin VB.CommandButton cmdcancelar 
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
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
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
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
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
      Height          =   495
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
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
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   7575
      Begin VSFlex7LCtl.VSFlexGrid GrillaPrecios 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   7335
         _cx             =   12938
         _cy             =   4683
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
   End
   Begin VB.Frame FraProductos 
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   7575
      Begin Gestion.ucCoDe ucCoDeProductos 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
      Begin VB.OptionButton optTodos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optUno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Uno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame FraPrecios 
      Caption         =   "Listado de Precios de "
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin Gestion.ucCoDe ucCoDeProveedores 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9551
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe ucCoDeClientes 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9551
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
      Begin VB.OptionButton optProveedores 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton optCliente 
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmListadoPreciosHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function CrearConsulta() As String
'CONSULTA DE CLIENTE

'select f.fecha, f.cliente, f.razonsocial, d.producto, d.descripcion, d.preciounitario
'from facturaventadetalle as d inner join facturaventa as f on f.codigo = d.codigofactura
'where f.cliente = 2  and d.producto = 'I  EM 032'
'order by d.producto, f.fecha
'
'CONSULTA DE PROVEEDOR

'select R.FECHA, R.PROVEEDOR, P.DESCRIPCION, D.PRODUCTO, PR.DESCRIPCION, D.COSTO
'from REMITOCOMPRADETALLE as D inner join REMITOCOMPRA as R on R.CODIGO = D.CODIGOREMITO
'    INNER JOIN PROV AS P ON P.CODIGO = R.PROVEEDOR
'    LEFT JOIN PRODUCTO AS PR ON PR.CODIGO = D.PRODUCTO
'Where r.Proveedor = 10
'order by D.PRODUCTO, R.FECHA

Dim Consulta As String

    If optCliente = False And optProveedores = False Then
        MsgBox "Debe Seleccionar si Quiere un Listado de Precios por Cliente o por Proveedor.", vbOKOnly, "Atencion"
        Exit Function
    End If
    
    If optCliente.Value = True And ucCoDeClientes.codigo = 0 Then
        MsgBox "Debe seleccionar un cliente.", vbOKCancel, "Atencion"
        ucCoDeClientes.SetFocus
        Exit Function
    End If
        
    If optProveedores.Value = True And ucCoDeProveedores.codigo = 0 Then
        MsgBox "Debe seleccionar un proveedor.", vbOKOnly, "Atencion"
        ucCoDeProveedores.SetFocus
        Exit Function
    End If

    If opttodos.Value = False And optuno = False Then
        MsgBox "Debe seleccionar si quiere listar todos los productos o uno en particular.", vbOKOnly, "Atencion"
        Exit Function
    End If
    
    If optuno.Value = True And ucCoDeProductos.codigo = "" Then
        MsgBox "Debe seleccionar un producto.", vbOKOnly, "Atencion"
        ucCoDeProductos.SetFocus
        Exit Function
    End If

    If optCliente.Value Then
        Consulta = "Select F.FECHA, F.CLIENTE, F.RAZONSOCIAL AS 'RAZON SOCIAL', D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO " & _
                    "From FACTURAVENTADETALLE as d inner join FACTURAVENTA as f on f.codigo = d.codigofactura " & _
                    "Where F.FECHA " & ucEntreFechas.ssBetween & " and f.cliente = " & ucCoDeClientes.codigo
        If optuno.Value = True Then Consulta = Consulta & "and d.producto = '" & ucCoDeProductos.codigo & "' "
        Consulta = Consulta & "Order by D.PRODUCTO, F.FECHA"
        
'        LlenarGrilla GrillaPrecios, Consulta, True
    End If
    
    If optProveedores.Value Then
        Consulta = "Select R.FECHA, R.PROVEEDOR, P.DESCRIPCION AS 'RAZON SOCIAL', D.PRODUCTO, PR.DESCRIPCION, D.COSTO " & _
                    "From REMITOCOMPRADETALLE as D Inner Join REMITOCOMPRA as R on R.CODIGO = D.CODIGOREMITO " & _
                                                "Inner Join PROV as P on P.CODIGO = R.PROVEEDOR " & _
                                                "Left Join PRODUCTO as PR on PR.CODIGO = D.PRODUCTO " & _
                    "Where r.fecha " & ucEntreFechas.ssBetween & " And R.PROVEEDOR = " & ucCoDeProveedores.codigo
        If optuno.Value = True Then Consulta = Consulta & " and d.producto = '" & ucCoDeProductos.codigo & "' "
        Consulta = Consulta & " Order By D.PRODUCTO, R.FECHA "
    
'        LlenarGrilla GrillaPrecios, Consulta, True
    End If

    CrearConsulta = Consulta
End Function

Private Sub cmdAceptar_Click()
Dim ConsultaAux As String
    ConsultaAux = CrearConsulta
    If ConsultaAux <> "" Then
        LlenarGrilla GrillaPrecios, ConsultaAux, True
    Else
        MsgBox "No Existe Resultado a la Busqueda.", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdCancelar_Click()
    optCliente.Value = False
    optProveedores.Value = False
    opttodos.Value = False
    optuno.Value = False
    LimpiarGrilla GrillaPrecios
    
    ucCoDeClientes.codigo = ""
    ucCoDeClientes.Visible = False
    
    ucCoDeProveedores.codigo = ""
    ucCoDeProveedores.Visible = False
    
    ucCoDeProductos.codigo = ""
    ucCoDeProductos.Visible = False
    
    ucEntreFechas.ini Date, Date
    

End Sub

Private Sub cmdImprimir_Click()
Dim ConsultaAux As String
    ConsultaAux = CrearConsulta
    If ConsultaAux <> "" Then
        With arListHistoricoPrecios
            
            If optCliente.Value Then
                .lblTitulo.caption = "Listado Historico de Precios de Producto/s De Cliente"
                .fieCodigo.Text = ucCoDeClientes.codigo
                .fieDescripcion.Text = ucCoDeClientes.DESCRIPCION
            Else
                .lblTitulo.caption = "Listado Historico de Precios de Producto/s De Proveedor"
                .fieCodigo.Text = ucCoDeProveedores.codigo
                .fieDescripcion.Text = ucCoDeProveedores.DESCRIPCION
            End If
            .lblFecha.caption = Date
            .Data.Connection = DataEnvironment1.Sistema
            .Data.Source = ConsultaAux
            
            .fieFecha.DataField = "FECHA"
            .fieCodigoProd.DataField = "PRODUCTO"
            .fieDescripcionProd.DataField = "DESCRIPCION"
            If optCliente.Value Then
                .fiePU.DataField = "PRECIOUNITARIO"
            Else
                .fiePU.DataField = "COSTO"
            End If
            
            .Show vbModal
        End With
    Else
        MsgBox "No Existe Resultado a la Busqueda.", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdsalir_Click()
     Unload Me
End Sub

Private Sub Form_Load()
    ucCoDeClientes.ini "Select DESCRIPCION from CLIENTES Where CODIGO = '###'", _
                        "Select CODIGO, DESCRIPCION From CLIENTES Where ACTIVO = 1", False
    ucCoDeProveedores.ini "Select DESCRIPCION from PROV Where CODIGO = '###'", _
                        "Select CODIGO, DESCRIPCION From PROV Where ACTIVO = 1", False
    ucCoDeProductos.ini "Select DESCRIPCION from PRODUCTO Where CODIGO = '###'", _
                        "Select CODIGO, DESCRIPCION From PRODUCTO Where ACTIVO = 1", True
    cmdCancelar_Click

End Sub

Private Sub optcliente_Click()
    ucCoDeClientes.Visible = True
    ucCoDeProveedores.Visible = False
End Sub

Private Sub optProveedores_Click()
    ucCoDeClientes.Visible = False
    ucCoDeProveedores.Visible = True
End Sub

Private Sub opttodos_Click()
    ucCoDeProductos.Visible = False
End Sub

Private Sub optuno_Click()
    ucCoDeProductos.Visible = True
End Sub

