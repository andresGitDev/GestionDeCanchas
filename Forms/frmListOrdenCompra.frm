VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListOrdenCompra 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ordenes de Compra"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   Icon            =   "frmListOrdenCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   5640
      Width           =   3375
      Begin VB.CheckBox chkPendiente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar solo las OC Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4095
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   9495
      Begin VSFlex7LCtl.VSFlexGrid GrillaOC 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
         Top             =   240
         Width           =   9255
         _cx             =   16325
         _cy             =   6588
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   975
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame2 
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
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Width           =   2655
      Begin MSComCtl2.DTPicker dtFechaD 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   284229633
         CurrentDate     =   38334
      End
      Begin MSComCtl2.DTPicker dtFechaH 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   284229633
         CurrentDate     =   38334
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proveedores "
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
      Width           =   6735
      Begin VB.Frame fraOrden 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ordenar Por "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   15
         Top             =   0
         Width           =   3375
         Begin VB.OptionButton Proveedor 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   17
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optFecha 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton optUno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Elegir Uno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
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
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin Gestion.ucCoDe ucProveedor 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
   End
End
Attribute VB_Name = "frmListOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function LimpiarGrilla()
    GrillaOC.clear
    GrillaOC.rows = 2
    GrillaOC.cols = 2
End Function

Private Sub cmdAceptar_Click()
'Select O.CODIGO AS NUMERO, P.CODIGO, P.DESCRIPCION, O.FECHA, O.FECHAENTREGA AS 'FECHA ENTREGA',
'        O.FECHAPAGO AS 'FECHA PAGO', FP.DESCRIPCION AS 'FORMA PAGO',
'        M.DESCRIPCION AS MONEDA, O.IMPORTE
'From ORDENESDECOMPRAS AS O
'    INNER JOIN PROV AS P ON P.CODIGO = O.PROVEEDOR
'    INNER JOIN MONEDAS AS M ON O.MONEDA = M.CODIGO
'    INNER JOIN FORMASPAGO AS FP ON O.FORMAPAGO = FP.CODIGO
'Where P.ACTIVO = 1 And O.FECHA  between convert(datetime , '12-13-03', 1)  AND convert(datetime , '12-13-04', 1)

Dim Consulta As String
Dim rsOC As New ADODB.Recordset

    LimpiarGrilla
    Consulta = CrearConsulta(False)
    
    If optFecha Then
        Consulta = Consulta & " Order By O.FECHA, P.CODIGO"
    Else
        Consulta = Consulta & " Order By P.CODIGO, O.FECHA"
    End If
        
    LlenarGrilla GrillaOC, Consulta, False
End Sub

Private Sub cmdCancelar_Click()
    dtfechad.Value = Date
    dtfechah.Value = Date
    opttodos_Click
    optFecha.Value = True
    fraOrden.Visible = False
    chkPendiente.Value = 0
    LimpiarGrilla
    ucProveedor.ini "Select DESCRIPCION From PROV Where CODIGO = ###", _
                    "Select CODIGO, DESCRIPCION from PROV Order By CODIGO", False
End Sub
Private Function CrearConsulta(ConDetalle As Boolean) As String
'EL PARAMETRO CON DETALLE DETERMINA SI LA CONSULTA DEVUELVE O NO LOS ITEMS DE UNA ORDEN DE COMPRA
Dim Consulta As String
    
    If Not ConDetalle Then
         Consulta = " Select Distinct O.CODIGO As NUMERO , P.CODIGO, P.DESCRIPCION, O.FECHA," & _
                    " O.FECHAENTREGA AS 'FECHA ENTREGA', " & _
                    " O.FECHAPAGO AS 'FECHA PAGO', FP.DESCRIPCION AS 'FORMA PAGO', " & _
                    " M.DESCRIPCION AS MONEDA, O.IMPORTE "
    Else
         Consulta = " Select O.CODIGO As NUMERO , P.CODIGO, P.DESCRIPCION, O.FECHA," & _
                    " O.FECHAENTREGA AS 'FECHA ENTREGA', I.CANTIDAD, I.PRODUCTO, " & _
                    " I.FECHAENTREGA AS FECHAENTREGAITEM,I.SALDO, I.COSTO , " & _
                    " O.FECHAPAGO AS 'FECHA PAGO', FP.DESCRIPCION AS 'FORMA PAGO', " & _
                    " M.DESCRIPCION AS MONEDA, O.IMPORTE, Pd.DESCRIPCION AS DESC_PROD "
    End If
    
    If chkPendiente.Value = 1 Then Consulta = Consulta & ", I.SALDO "
                                           
    'pregunto si quiere ver los detalles de las OC
    
    If ConDetalle Or chkPendiente = 1 Then
      Consulta = Consulta & "From ORDENESDECOMPRAS AS O " & _
                 "left JOIN PROV AS P ON P.CODIGO = O.PROVEEDOR " & _
                 "left JOIN MONEDAS AS M ON O.MONEDA = M.CODIGO " & _
                 "left JOIN FORMASPAGO AS FP ON O.FORMAPAGO = FP.CODIGO " & _
                 "INNER JOIN ITEMORDENCOMPRA AS I ON O.CODIGO = I.ORDENCOMPRA " & _
                 "INNER JOIN Producto Pd ON I.producto = Pd.codigo "
       Else
           Consulta = Consulta & "From ORDENESDECOMPRAS AS O " & _
                 "left JOIN PROV AS P ON P.CODIGO = O.PROVEEDOR " & _
                 "left JOIN MONEDAS AS M ON O.MONEDA = M.CODIGO " & _
                 "left JOIN FORMASPAGO AS FP ON O.FORMAPAGO = FP.CODIGO "
    End If
               
                                            
    Consulta = Consulta & "Where P.ACTIVO = 1"
    
    'pregunto si quiere ver un proveedor en particular
    If optuno And ucProveedor.codigo <> 0 Then Consulta = Consulta & " And P.CODIGO = " & ucProveedor.codigo
    
    'pregunto si quiere mostrar las OC Pendientes
    If chkPendiente = 1 Then
        Consulta = Consulta & " And I.SALDO > 0 "
    Else
        Consulta = Consulta & " And O.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
    End If
    
    CrearConsulta = Consulta
End Function
Private Sub cmdImprimir_Click()
Dim Consulta As String
Dim rs As New ADODB.Recordset
Dim Detalle As Boolean

    Detalle = MsgBox("¿Desea imprimir los detalles de la/s orden/es de compra/s?", vbYesNo, "Atencion") = vbYes
    
    Consulta = CrearConsulta(Detalle) & " Order By P.CODIGO, O.FECHA"
    
    If Detalle Then
        With rptListadoOCDetalle
        
        .lblTitulo.caption = "Listado de Ordenes de Compras"
        .lblFecha.caption = Date
        .DataControl1.Connection = DataEnvironment1.Sistema
        .DataControl1.Source = Consulta
        
        .fieCodigoProv.DataField = "CODIGO"
        .fieDescripcionProv.DataField = "DESCRIPCION"
        .fieDes_Prod.DataField = "DESC_PROD"
        .fieFecha.DataField = "FECHA"
        .fieNroOC.DataField = "NUMERO"
        .fieFechaEntrega.DataField = "FECHA ENTREGA"
        .fieFechaPago.DataField = "FECHA PAGO"
        .fieFP.DataField = "FORMA PAGO"
        .fieMoneda.DataField = "MONEDA"
        
        .fieCantidad.DataField = "CANTIDAD"
        .fieProducto.DataField = "PRODUCTO"
        .fieFechaEntregaDetalle.DataField = "FECHAENTREGAITEM"
        .fieCosto.DataField = "COSTO"
        .fieSaldo.DataField = "SALDO"
        
        .GroupHeader1.DataField = "NUMERO"
    
        .Show
    
        End With
    Else
        With rptListadoOC
    
        .lblTitulo.caption = "Listado de Ordenes de Compras"
        .lblFecha.caption = Date
        .DataControl1.Connection = DataEnvironment1.Sistema
        .DataControl1.Source = Consulta
        
        .fieCodigoProv.DataField = "CODIGO"
        .fieDescripcionProv.DataField = "DESCRIPCION"
        .fieFechaOC.DataField = "FECHA"
        .fieNroOC.DataField = "NUMERO"
        .fieFechaEntrega.DataField = "FECHA ENTREGA"
        .fieFechaPago.DataField = "FECHA PAGO"
        .fieFP.DataField = "FORMA PAGO"
        .fieMoneda.DataField = "MONEDA"
        .fieImporte.DataField = "IMPORTE"
        
        .GroupHeader1.DataField = "CODIGO"
    
        .Show
    
        End With
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    cmdCancelar_Click
End Sub

Private Sub GrillaOC_DblClick()
'SELECT I.ORDENCOMPRA, I.PRODUCTO, I.CANTIDAD, I.SALDO, I.COSTO, I.FECHAENTREGA AS 'FECHA ENTREGA'
'FROM ITEMORDENCOMPRA AS I
'Where i.ORDENCOMPRA = 902

    With GrillaOC
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            frmBuscar.MostrarSql "SELECT I.ORDENCOMPRA, I.PRODUCTO, I.CANTIDAD, I.SALDO, " & _
                                        "I.COSTO, I.FECHAENTREGA AS 'FECHA ENTREGA' " & _
                                 "FROM ITEMORDENCOMPRA AS I " & _
                                 "Where i.ORDENCOMPRA = " & .TextMatrix(.Row, 0), , _
                                            "Detalle de la Orden de Compra Nro " & .TextMatrix(.Row, 0)
        End If
    End With
End Sub

Private Sub opttodos_Click()
    ucProveedor.Visible = False
    fraOrden.Visible = True
End Sub

Private Sub optuno_Click()
    ucProveedor.Visible = True
    fraOrden.Visible = False
End Sub

