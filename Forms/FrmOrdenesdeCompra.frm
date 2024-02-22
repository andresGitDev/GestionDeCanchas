VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmOrdenesdeCompra 
   Caption         =   "Orden de Compra"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "FrmOrdenesdeCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucCoDe uProv 
      Height          =   330
      Left            =   1725
      TabIndex        =   1
      Top             =   615
      Width           =   6540
      _ExtentX        =   14102
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   315
      Left            =   1455
      TabIndex        =   10
      Top             =   2850
      Width           =   6825
      _ExtentX        =   14552
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VSFlex7LCtl.VSFlexGrid grillaProductos 
      Height          =   2550
      Left            =   150
      TabIndex        =   36
      Top             =   4170
      Width           =   8565
      _cx             =   15108
      _cy             =   4498
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
   Begin Gestion.ucBotonera uMenu 
      Cancel          =   -1  'True
      Height          =   2100
      Left            =   -30
      TabIndex        =   15
      Top             =   6810
      Width           =   8910
      _ExtentX        =   18680
      _ExtentY        =   3122
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucEntreFechas uBuscaFechas 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   1530
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total OC:"
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
         Left            =   6600
         TabIndex        =   38
         Top             =   75
         Width           =   975
      End
      Begin VB.Label lblTotalOc 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7740
         TabIndex        =   37
         Top             =   75
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Entre:"
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   1590
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Height          =   570
      Left            =   6900
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmOrdenesdeCompra.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3570
      Width           =   555
   End
   Begin VB.TextBox txtcotizacion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5685
      TabIndex        =   9
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtrecibio 
      Height          =   300
      Left            =   2130
      TabIndex        =   7
      Top             =   2355
      Width           =   2175
   End
   Begin VB.TextBox txtordenprov 
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtcosto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6300
      TabIndex        =   13
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtelaboro 
      Height          =   285
      Left            =   2130
      TabIndex        =   5
      Top             =   1725
      Width           =   2175
   End
   Begin VB.ComboBox cmbmoneda 
      Height          =   315
      Left            =   5685
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1575
      Width           =   2415
   End
   Begin VB.TextBox txtdescuento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2130
      TabIndex        =   4
      Top             =   1410
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtfechapago 
      Height          =   300
      Left            =   6150
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   81330177
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   300
      Left            =   2940
      TabIndex        =   0
      Top             =   225
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   81330177
      CurrentDate     =   38098
   End
   Begin VB.ComboBox cmbformapago 
      Height          =   315
      Left            =   2130
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1065
      Width           =   2355
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1455
      TabIndex        =   11
      Top             =   3225
      Width           =   855
   End
   Begin VB.CommandButton cmdotro 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "FrmOrdenesdeCompra.frx":1194
      Height          =   570
      Left            =   6300
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmOrdenesdeCompra.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3570
      Width           =   585
   End
   Begin MSComCtl2.DTPicker Dtfechaentrega 
      Height          =   300
      Left            =   4185
      TabIndex        =   12
      Top             =   3225
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      Format          =   81330177
      CurrentDate     =   38098
   End
   Begin VB.Label lblEntregado 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2325
      TabIndex        =   35
      Top             =   3225
      Width           =   915
   End
   Begin VB.Label Label16 
      Caption         =   "Entrega:"
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
      Left            =   3435
      TabIndex        =   32
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label15 
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
      Left            =   4650
      TabIndex        =   31
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00400000&
      Height          =   1275
      Left            =   4500
      Top             =   1425
      Width           =   3735
   End
   Begin VB.Label Label14 
      Caption         =   "Recibio :"
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
      Left            =   1215
      TabIndex        =   30
      Top             =   2355
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "O/C Proveedor :"
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
      Left            =   1605
      TabIndex        =   29
      Top             =   2055
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Costo :"
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
      Left            =   5670
      TabIndex        =   28
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Elaboro :"
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
      Left            =   1260
      TabIndex        =   27
      Top             =   1755
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Moneda :"
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
      Left            =   4770
      TabIndex        =   26
      Top             =   1590
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   6735
      Left            =   75
      Top             =   75
      Width           =   8760
   End
   Begin VB.Label Label9 
      Caption         =   "Fecha de Pago :"
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
      Left            =   4635
      TabIndex        =   25
      Top             =   1095
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "% Descuento :"
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
      Left            =   795
      TabIndex        =   24
      Top             =   1410
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Nro :"
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
      Left            =   675
      TabIndex        =   22
      Top             =   240
      Width           =   495
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
      Left            =   615
      TabIndex        =   21
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Forma de Pago :"
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
      Left            =   585
      TabIndex        =   20
      Top             =   1065
      Width           =   1455
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2250
      TabIndex        =   19
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      Left            =   495
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   225
      X2              =   8145
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad :"
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
      Left            =   495
      TabIndex        =   17
      Top             =   3210
      Width           =   1095
   End
End
Attribute VB_Name = "FrmOrdenesdeCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const gCANT = 2
Private Const gPREC = 3

Private Const tt_OrdenCompraTemp = "([CODIGO] [varchar] (50), [ProdProv] [varchar] (50), [DESCRIPCION] [varchar] (900), [FECHA_ENTREGA] [varchar] (50), [CANTIDAD] [varchar] (50), [PU] [varchar] (50), [PT] [varchar] (50), [Letra] char (1) )"

Sub CargoRegistro()
    Dim rsdetalle As New ADODB.Recordset
    
    With uMenu.rs
        txtCodigo = !codigo
        dtFecha = !Fecha
        UpROV.codigo = s2n(!Proveedor)
        
        txtCotizacion = s2n(!cotizacion)
        
        If Not IsNull(!formapago) Then
            cmbformapago.Text = ObtenerDescripcion("formaspago", !formapago)
        Else
            cmbformapago.Text = ""
        End If
        
        If Not IsNull(!moneda) Then
            cmbMoneda.Text = ObtenerDescripcion("monedas", !moneda)
        Else
            cmbMoneda.Text = ""
        End If
        
        If Not IsNull(!ELABORO) Then
            txtelaboro = !ELABORO
        Else
            txtelaboro = ""
        End If
        
        If Not IsNull(!ordenproveedor) Then
            txtordenprov = !ordenproveedor
        Else
            txtordenprov = ""
        End If
        If Not IsNull(!RECIBIO) Then
            txtrecibio = !RECIBIO
        Else
            txtrecibio = ""
        End If
        If Not IsNull(!Descuento) Then
            txtDescuento = !Descuento
        Else
            txtDescuento = ""
        End If
        If Not IsNull(!fechapago) Then
            dtfechapago.Value = !fechapago
        Else
            dtfechapago.Value = Date
        End If
        InicioGrilla
        rsdetalle.Open "Select * from itemordencompra where ordencompra=" & s2n(txtCodigo), DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rsdetalle.EOF Then
            Do While Not rsdetalle.EOF
                If grillaproductos.rows = 2 Then
                    grillaproductos.Row = 1
                    grillaproductos.Col = 0
                    
                    If Trim(grillaproductos.Text) = "" Then
                        grillaproductos.Row = 1
                        grillaproductos.Col = 0
                        grillaproductos.Text = rsdetalle!Producto
                        grillaproductos.Col = 1
                        grillaproductos.Text = ObtenerDescripcionS("producto", Trim(rsdetalle!Producto))
                        grillaproductos.Col = 2
                        grillaproductos.Text = rsdetalle!Cantidad
                        grillaproductos.Col = 3
                        grillaproductos.Text = rsdetalle!costo
                        grillaproductos.Col = 4
                        'grillaproductos.Text = rsdetalle!Estado
                        grillaproductos.Text = rsdetalle!saldo
                        grillaproductos.Col = 5
                        grillaproductos.Text = rsdetalle!fechaEntrega
                    Else
                        grillaproductos.AddItem rsdetalle!Producto & Chr(9) & ObtenerDescripcionS("producto", rsdetalle!Producto) & Chr(9) & rsdetalle!Cantidad & Chr(9) & rsdetalle!costo & Chr(9) & rsdetalle!saldo & Chr(9) & rsdetalle!fechaEntrega
                    End If
                Else
                    grillaproductos.AddItem rsdetalle!Producto & Chr(9) & ObtenerDescripcionS("producto", rsdetalle!Producto) & Chr(9) & rsdetalle!Cantidad & Chr(9) & rsdetalle!costo & Chr(9) & rsdetalle!saldo & Chr(9) & rsdetalle!fechaEntrega
                End If
                rsdetalle.MoveNext
            Loop
        End If
    End With
    CalculaTotal
fin:
    Set rsdetalle = Nothing
    Exit Sub
ufaErr:
    ufa "Err cargando registro", "" ', Err
    Resume fin
End Sub

Private Function ObtenerCostoProv(COD As String) As Variant
    Dim t As Double
   
    t = s2n(obtenerDeSQL("select precio from RELACION_PRODUCTO_proveedor where producto='" & COD & "' and proveedor=" & UpROV.codigo), 4)
    
    If t = 0 Then
        t = s2n(obtenerDeSQL("select costobase from producto where codigo ='" & COD & "' "), 4)
    End If
    ObtenerCostoProv = t
    
End Function

Private Function ObtenerFormaPagoProveedor(COD As Long) As String
Dim rs As New ADODB.Recordset

    rs.Open "Select pago from prov where codigo=" & COD, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        If Not IsNull(rs!pago) Then
            ObtenerFormaPagoProveedor = rs!pago
        Else
            ObtenerFormaPagoProveedor = ""
        End If
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub cmdBorrar_Click()
    borraLineaEdicion
End Sub

Private Sub borraLineaEdicion()
    uProd.clear
    txtcosto = ""
    txtcantidad = ""
    lblEntregado = ""
End Sub

Private Sub cmdImprimir_Click()
    Dim rsDatosEmpresa As New ADODB.Recordset
    Dim rsDatosProvee As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Consulta As String
    Dim i As Long
    Dim Total As Double
    
    Dim codi As String, letra As String, desc As String, ProdProv As String
    
    Dim sTablaTempOC As String
    
    sTablaTempOC = TablaTempCrear(tt_OrdenCompraTemp)

    With rptOrdenCompra
        'DATOS DEL ENCABEZADO DE LA EMPRESA
        .fieNroOC.Text = txtCodigo
        .fieFecha.Text = dtFecha.Value

        rsDatosEmpresa.Open "Select * From DATOSEMPRESA", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            rsDatosEmpresa.MoveFirst
            .lblTitulo.caption = sSinNull(obtenerDeSQL("select NombreCortoParaListados from datosempresa where idempresa=" & rsDatosEmpresa!idempresa)) 'sSinNull(rsDatosEmpresa!NombreCortoParaListados)
            .fieTelefono.Text = sSinNull(rsDatosEmpresa!Telefono)
            .fieCUIT.Text = sSinNull(rsDatosEmpresa!CUITEMPRESA)
            .fieIVA.Text = sSinNull(rsDatosEmpresa!Iva)
            .fieIIBB.Text = sSinNull(rsDatosEmpresa!IBrutos)
            .fieCajaPrev.Text = sSinNull(rsDatosEmpresa!cajaprev)
        Set rsDatosEmpresa = Nothing

        'DATOS DEL PROVEEDOR
        rsDatosProvee.Open "Select * From PROV Where CODIGO = " & UpROV.codigo, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rsDatosProvee.EOF Then
            rsDatosProvee.MoveFirst
            If Not IsNull(rsDatosProvee!DESCRIPCION) Then .fieDescripcionProve.Text = rsDatosProvee!DESCRIPCION
            If Not IsNull(rsDatosProvee!codigo) Then .fieCodigoProve.Text = rsDatosProvee!codigo
            If Not IsNull(rsDatosProvee!direccion) Then .fieDomicilioProve.Text = rsDatosProvee!direccion
            If Not IsNull(rsDatosProvee!codigopostal) Then .fieCPProve.Text = rsDatosProvee!codigopostal
            If Not IsNull(rsDatosProvee!Localidad) Then .fieLocalidadProve.Text = rsDatosProvee!Localidad
            If Not IsNull(rsDatosProvee!Telefono) Then .fieTelefonoProve.Text = rsDatosProvee!Telefono
            If Not IsNull(rsDatosProvee!CUIT) Then .fieCuitProve.Text = rsDatosProvee!CUIT
            If Not IsNull(rsDatosProvee!tipoiva) Then .fieIvaProve.Text = ObtenerDescripcion("IVAS", rsDatosProvee!tipoiva)
        End If
        Set rsDatosProvee = Nothing

        .fieFormaPago.Text = cmbformapago.Text
        .fieFechaPago.Text = dtfechapago.Value
        .fieElaboro.Text = txtelaboro.Text
        .fieMoneda.Text = cmbMoneda.Text

        
        Total = 0
        For i = 1 To grillaproductos.rows - 1
            With grillaproductos
                codi = .TextMatrix(i, 0)
                letra = sSinNull(obtenerDeSQL("select letra from producto where codigo = '" & codi & "' "))
                desc = .TextMatrix(i, 1)
                ProdProv = "Cod Prov:  " & VerProdProv(codi, UpROV.codigo)
                
                Consulta = "Insert Into " & sTablaTempOC & " (CODIGO, DESCRIPCION, FECHA_ENTREGA, CANTIDAD, PU, PT, ProdProv, Letra ) " & _
                            "Values ('" & codi & "', '" & desc & "', '" & _
                                    .TextMatrix(i, 5) & "', '" & .TextMatrix(i, 2) & "', '" & _
                                    .TextMatrix(i, 3) & "', '" & (s2n(.TextMatrix(i, 2)) * s2n(.TextMatrix(i, 3))) & "', " & _
                                    " '" & ProdProv & "', '" & letra & "' )"
                
                DataEnvironment1.Sistema.Execute Consulta
                Total = Total + (s2n(.TextMatrix(i, 2)) * s2n(.TextMatrix(i, 3)))
            End With
        Next
        Consulta = "Select * From " & sTablaTempOC
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = Consulta

        'DETALLE DEL REPORTE
        .fieCodigo.DataField = "CODIGO"
        .FieDescripcion.DataField = "DESCRIPCION"
        .fieFechaEntrega.DataField = "FECHA_ENTREGA"
        .fieCantidad.DataField = "CANTIDAD"
        .fiePU.DataField = "PU"
        .fiePT.DataField = "PT"
        .fieCodProv.DataField = "ProdProv"
        .fieLetra.DataField = "Letra"
        .fieLetraP.DataField = "Letra"

        'COLA DEL REPORTE
        .fieTotal.Text = Total
        .lblleyenda.caption = leyenda()

        If PREVIEW_IMPRESIONES Then
            .Show
        Else
            .Restart
            .PrintReport False
        End If
    End With
End Sub

Private Function leyenda()
    leyenda = "Modalidad de entrega  a confirmar."
    'leyenda = "Lugar de entrega: Santiago del Estero y J. Hernandez - Garin - Pcia. de Bs. As. " & vbCrLf & _
            "    Tel.: (03488) 471-630 - Atn.: Sr. Ruben (Confirma fecha de entrega y coordinar horario) "
    'leyenda = "Sr. Proveedor: " & vbCrLf & _
    '        "    * Si el producto tiene Letra de actualización, verifique que coincida con la letra del plano que Ud. tiene. De no ser así, comuníquese con Tonka a la brevedad, al 4725-1566, para solicitar el plano correspondiente." & vbCrLf & _
    '        "      En el remito que el proveedor envía TIENE que figurar esa letra. " & vbCrLf & _
    '        "    * Le recordamos la necesidad de mencionar el número de ORDEN de COMPRA cuando nos envía mercadería."
End Function

Private Sub cmdotro_Click()
    Dim i As Long, CodProd As String
    
    CodProd = uProd.codigo
    
    If s2n(txtcantidad) <= 0 Then
        che "Falta especificar cantidad"
        txtcantidad.SetFocus
        Exit Sub
    End If
    If uProd.DESCRIPCION = "" Then
        che "Falta especificar producto"
        uProd.SetFocus
        Exit Sub
    End If
    If s2n(txtcantidad) < s2n(lblEntregado) Then
        che "Cantidad menor a lo ya entregado"
        txtcantidad.SetFocus
        Exit Sub
    End If
    For i = 1 To grillaproductos.rows - 1
        If CodProd > "" And CodProd = grillaproductos.TextMatrix(i, 0) Then
            che "ya existe ese producto en la grilla"
            Exit Sub
        End If
    Next i
    
    Dim rs As New ADODB.Recordset, ssql As String
    ssql = "select codigo, componente, cantidad from formulas where activo = 1 and codigo = '" & CodProd & "'"
    rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
   
    With grillaproductos
        .AddItem uProd.codigo & vbTab & uProd.DESCRIPCION & vbTab & s2n(txtcantidad) & vbTab & s2n(txtcosto, 4) & vbTab & (s2n(txtcantidad) - s2n(lblEntregado)) & vbTab & dtfechaentrega.Value
        
        lblEntregado = ""
        txtcantidad.Text = s2n(0)
        txtcosto.Text = s2n(0, 4)
        
        uProd.clear
        uProd.SetFocus
        
    End With
    CalculaTotal
    
    Set rs = Nothing
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
    CalculaTotal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub
Sub InicioGrilla()
    With grillaproductos
        .clear
        .FormatString = "Producto             |<Descripcion                                     |>Cantidad     |>Costo     | Saldo  |Entrega         |Formula           "
        .rows = 1
        .cols = 7
    End With
    CalculaTotal
End Sub


Private Sub grillaProductos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = gCANT Or Col = gPREC Then CalculaTotal
End Sub

Private Sub grillaproductos_DblClick()
    On Error Resume Next
    Dim i As Long
    
    If uMenu.estado <> ucbEditando Then
        Exit Sub
    End If
    If grillaproductos.rows < 2 Then
        Exit Sub
    End If
    
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Sub
    End If
    
    With grillaproductos
    
       
        .Col = 0:   uProd.codigo = .Text
        .Col = 1:   If uProd.codigo = "" Then uProd.DESCRIPCION = .Text
        .Col = 2:   txtcantidad = .Text
        .Col = 5:   dtfechaentrega.Value = .Text
        .Col = 3:   txtcosto = .Text
        .Col = 4:   lblEntregado = s2n(txtcantidad) - s2n(.Text)
        
        .RemoveItem (.Row)
    
    End With
    
    CalculaTotal
End Sub

Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub

Private Sub CalculaTotal()
    Dim i As Long, tot As Double
    With grillaproductos
        For i = 1 To .rows - 1
            tot = tot + s2n(.TextMatrix(i, gCANT), 4) * s2n(.TextMatrix(i, gPREC), 4)
        Next i
    End With
    lblTotalOc = s2n(tot)
End Sub

Private Sub txtcosto_GotFocus()
    PintoFocoActivo
End Sub

Sub LimpioTxt()
    uProd.clear
    UpROV.clear
    txtCodigo = ""
    txtDescuento = "0"
    txtelaboro = ""
    txtrecibio = ""
    txtordenprov = ""
    txtcantidad = "0,00"
    txtcosto = "0,0000"
    txtCotizacion = "0,00"
    cmbMoneda.Text = ""
    cmbformapago.Text = ""
    dtFecha.Value = Date
    dtfechapago.Value = Date
    lblEntregado = ""
    InicioGrilla
End Sub
Sub HabilitoTxt(habilito As Boolean) 'OJO lo puso al reves, HABILITA CON FALSE
    uProd.enabled = Not habilito
    UpROV.enabled = Not habilito
    txtCodigo.Locked = habilito
    txtDescuento.Locked = habilito
    txtelaboro.Locked = habilito
    txtrecibio.Locked = habilito
    txtordenprov.Locked = habilito
    txtcantidad.Locked = habilito
    txtcosto.Locked = habilito
    txtCotizacion.Locked = habilito
    cmbMoneda.Locked = habilito
    cmbformapago.Locked = habilito
    dtfechaentrega.enabled = Not habilito
    dtfechapago.enabled = Not habilito
    cmdotro.enabled = Not habilito
End Sub

Private Sub Form_Load()
    CargaCombo cmbMoneda, "Monedas", "descripcion", "codigo", ""
    CargaCombo cmbformapago, "formasPago", "descripcion", "codigo", ""
    
    uProd.ini "select descripcion from producto where codigo = '###' and activo = 1", "select codigo as [codigo          ], descripcion as [Descripcion                                               ] from producto where activo = 1 order by codigo ", True
    uProd.EditaDescripcion = True
    
    UpROV.ini "select descripcion from prov where activo = 1 and codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Descripcion               ] from prov where activo = 1 order by codigo ", False
    UpROV.EditaDescripcion = True
    
    uMenu.init True, True, True, True, True, "select * from ordenesdecompras where activo=1", DataEnvironment1.Sistema
End Sub

Private Sub uProd_cambio(codigo As Variant)
    If uMenu.estado = ucbEditando Then txtcosto = ObtenerCostoProv(CStr(codigo))
End Sub


Private Sub Txtdescuento_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtrecibio_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtelaboro_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtordenprov_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcotizacion_GotFocus()
    PintoFocoActivo
End Sub

Private Sub uMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAerrALTA
    
    Dim fEntrega As Date
    Dim codigo As Long
    Dim i As Long
    Dim COD As String
    Dim cant As Double
    Dim costo As Double
    Dim Total As Variant
    
    If UpROV.DESCRIPCION = "" Then
        MsgBox "No hay proveedor cargado.", vbInformation, "Falta proveedor."
        UpROV.SetFocus
        Exit Sub
    End If
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Sub
    End If
    
    codigo = nuevoCodigo("OrdenesDeCompras")
    
    DE_BeginTrans

    'DataEnvironment1.dbo_ORDENCOMPRA "A", s2n(txtCodigo), uProv.CODIGO, Trim(txtdescuento), dtfechapago, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), Trim(txtelaboro), ObtenerCodigo("monedas", Trim(cmbmoneda.Text)), s2n(txtcotizacion), 0, Trim(txtrecibio), Trim(txtordenprov), Date, UsuarioSistema!CODIGO, 0, 0
    If ABMOrdenCompra("A", s2n(txtCodigo), UpROV.codigo, Trim(txtDescuento), dtfechapago, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), Trim(txtelaboro), ObtenerCodigo("monedas", Trim(cmbMoneda.Text)), s2n(txtCotizacion), 0, Trim(txtrecibio), Trim(txtordenprov), UsuarioSistema!codigo, 0) = False Then GoTo UFAerrALTA
    DataEnvironment1.Sistema.Execute "DELETE FROM ITEMORDENCOMPRA WHERE ORDENCOMPRA=" & s2n(txtCodigo)
    
    Total = 0
    With grillaproductos
        For i = 1 To .rows - 1
            COD = .TextMatrix(i, 0)
            cant = .TextMatrix(i, 2)
            costo = .TextMatrix(i, 3)
            If .TextMatrix(i, 5) > "" Then
                fEntrega = CDate(.TextMatrix(i, 5))
            Else
                fEntrega = 0
            End If
            Total = Total + (cant * costo)
            'DataEnvironment1.dbo_ITEMORDENESCOMPRA cod, s2n(txtCodigo), s2n(cant), s2n(cant), "PE", s2n(COSTO), fEntrega
            If AItemOC(COD, s2n(txtCodigo), s2n(cant), s2n(cant), "PE", s2n(costo), fEntrega) = False Then GoTo UFAerrALTA
        Next i
        DataEnvironment1.Sistema.Execute "UPDATE OrdenesdeCompras set importe=" & x2s(Total) & " where codigo=" & s2n(txtCodigo)
    End With
    DE_CommitTrans
    
    MsgBox "La Operación se ha realizado con éxito.", vbInformation, "Fin de OC"
    cmdImprimir_Click
    uMenu.AceptarOk

    Exit Sub
UFAerrALTA:
    DE_RollbackTrans
    MsgBox "Error en OC", vbCritical, "No se guardo la OC"
End Sub
Private Sub uMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufamodi
    
    Dim fEntrega As Date
    Dim codigo As Long
    Dim i As Long
    Dim COD As String
    Dim cant As Double
    Dim costo As Double
    Dim Total As Variant
    Dim saldo As Double


    If UpROV.DESCRIPCION = "" Then
        MsgBox "No hay proveedor cargado.", vbInformation, "Falta proveedor."
        UpROV.SetFocus
        Exit Sub
    End If
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Sub
    End If

    DE_BeginTrans
    
    'DataEnvironment1.dbo_ORDENCOMPRA "M", Val(Trim(txtCodigo)), uProv.CODIGO, Trim(txtdescuento), dtfechapago, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), Trim(txtelaboro), ObtenerCodigo("monedas", Trim(cmbmoneda.Text)), s2n(txtcotizacion), 0, Trim(txtrecibio), Trim(txtordenprov), Date, UsuarioSistema!CODIGO, 0, 0
    ABMOrdenCompra "M", s2n(txtCodigo), UpROV.codigo, Trim(txtDescuento), dtfechapago, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), Trim(txtelaboro), ObtenerCodigo("monedas", Trim(cmbMoneda.Text)), s2n(txtCotizacion), 0, Trim(txtrecibio), Trim(txtordenprov), UsuarioSistema!codigo, 0
    DataEnvironment1.Sistema.Execute "DELETE FROM ITEMORDENCOMPRA WHERE ORDENCOMPRA=" & s2n(txtCodigo)
    
    
    Total = 0
    With grillaproductos
        For i = 1 To .rows - 1
            COD = .TextMatrix(i, 0)
            cant = .TextMatrix(i, 2)
            costo = .TextMatrix(i, 3)
            If .TextMatrix(i, 5) > "" Then
                fEntrega = CDate(.TextMatrix(i, 5))
            Else
                fEntrega = 0
            End If
        
        Total = Total + (cant * costo)
        
        'DataEnvironment1.dbo_ITEMORDENESCOMPRA cod, s2n(txtCodigo), s2n(cant), SALDO, "PE", s2n(COSTO), fEntrega
        AItemOC COD, s2n(txtCodigo), s2n(cant), s2n(cant), "PE", s2n(costo), fEntrega
        Next i
    End With
    DataEnvironment1.Sistema.Execute "UPDATE ordenesdecompras set importe=" & x2s(Total) & " where codigo=" & s2n(txtCodigo)
    DataEnvironment1.dbo_GRABARBITACORA s2n(txtCodigo), "OrdenesdeCompras", UsuarioSistema!codigo, Date, Time, "M"
    uMenu.AceptarOk
    
    DE_CommitTrans
    
    MsgBox "La Operación se ha realizado con éxito.", vbInformation, "OC Modificada"
Exit Sub
ufamodi:
    MsgBox "Error en OC", vbCritical, "No se guardo la OC"
    DE_RollbackTrans
End Sub
Private Sub uMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub uMenu_Buscar()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select ordenesdecompras.codigo as [ Nro  ], Prov.Codigo as [Prov], prov.descripcion  as [ Nombre Proveedor ],  ordenproveedor as [ Nro OC Prov] from OrdenesDeCompras inner join prov  on OrdenesDeCompras.proveedor = prov.codigo where OrdenesDeCompras.activo = 1 and fecha " & uBuscaFechas.ssBetween & " order by ordenesdecompras.codigo desc")
    If resu > "" Then
        txtCodigo = resu
        uMenu.BuscarOK "codigo = " & resu
        CargoRegistro
    End If
End Sub
Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim asse As String
    
    Dim sp
    sp = obtenerDeSQL("select * from remitocompradetalle d inner join remitocompra r on d.codigoremito=r.codigo where r.activo=1 and d.ordencompra=" & s2n(txtCodigo))
    If IsNull(sp) Or IsEmpty(sp) Then
    Else
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If
       

    DE_BeginTrans
    
    asse = "OCE"
    'DataEnvironment1.dbo_ORDENCOMPRA "B", Trim(txtCodigo), 0, "", 0, 0, "", 0, 0, 0, "", "", 0, 0, UsuarioSistema!CODIGO, Date
    ABMOrdenCompra "B", s2n(txtCodigo), 0, "", Date, 0, "", 0, 0, 0, "", "", 0, UsuarioSistema!codigo
    asse = "ITEM"
    DataEnvironment1.Sistema.Execute "DELETE FROM ITEMORDENCOMPRA WHERE ORDENCOMPRA=" & s2n(txtCodigo)
    'DataEnvironment1.dbo_BORRARITEMORDENESCOMPRA s2n(txtCodigo)
    asse = "bita"
    DataEnvironment1.dbo_GRABARBITACORA s2n(txtCodigo), "ordenesdecompras", UsuarioSistema!codigo, Date, Time, "B"
    uMenu.EliminarOK
    
    DE_CommitTrans

fin:
    Exit Sub
ufaErr:
    DE_RollbackTrans
    ufa "error al eliminar", "OC eliminar, cod: " & txtCodigo & " ,asse =  " & asse ', Err
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt Not sino

End Sub
Private Sub uMenu_Imprimir()
    cmdImprimir_Click
End Sub
Private Sub uMenu_Modificar()

    UpROV.SetFocus
End Sub
Private Sub uMenu_Nuevo()
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    txtCodigo = nuevoCodigo("OrdenesDeCompras")

    dtfechaentrega = Date
    UpROV.SetFocus
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
Private Sub uMenu_SeMovio()
    txtCodigo = s2n(uMenu.rs!codigo)
    CargoRegistro
End Sub

Private Sub uProv_cambio(codigo As Variant)
    If uMenu.estado <> ucbEditando Then Exit Sub
    
    cmbformapago.Text = ObtenerFormaPagoProveedor(UpROV.codigo)
    cmbformapago.SetFocus
End Sub

Public Function ABMOrdenCompra(oOPE As String, oCODIGO As Long, oPROVEEDOR As Long, oDESCUENTO As String, oPAGO As Date, oFORMAPAGO As Long, oELABORO As String, oMONEDA As Long, oCOTIZA As Double, oIMPORTE As Double, oRECIBIO As String, oORDENPROV As String, oUSUALTA As Long, oUSUBAJA As Long) As Boolean
On Error GoTo omal
Dim iud As String
ABMOrdenCompra = True
Select Case oOPE
    Case "A":
        iud = "INSERT INTO ORDENESDECOMPRAS (CODIGO,PROVEEDOR,FECHA,DESCUENTO,FECHAPAGO,FORMAPAGO,ELABORO,MONEDA,COTIZACION,IMPORTE,ORDENPROVEEDOR,RECIBIO, FECHA_ALTA, USUARIO_ALTA,  ACTIVO) " _
                                & " VALUES (" & oCODIGO & "," & oPROVEEDOR & "," & ssFecha(Date) & "," & ssTexto(oDESCUENTO) & "," & ssFecha(oPAGO) & "," & oFORMAPAGO & "," & ssTexto(oELABORO) _
                                & "," & oMONEDA & "," & x2s(oCOTIZA) & "," & x2s(oIMPORTE) & "," & ssTexto(oORDENPROV) & "," & ssTexto(oRECIBIO) & "," & ssFecha(Date) & "," & oUSUALTA & ", 1)"
        DataEnvironment1.Sistema.Execute iud
    Case "B":
        iud = "UPDATE ORDENESDECOMPRAS  SET ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA= " & oUSUBAJA & " WHERE CODIGO=" & oCODIGO
        DataEnvironment1.Sistema.Execute iud
    Case "M":
        iud = "UPDATE ORDENESDECOMPRAS SET PROVEEDOR=" & oPROVEEDOR & ",DESCUENTO=" & ssTexto(oDESCUENTO) & ",FECHAPAGO=" & ssFecha(oPAGO) & ",FORMAPAGO=" & oFORMAPAGO & ",ELABORO=" & ssTexto(oELABORO) _
                                & ",MONEDA=" & oMONEDA & ",COTIZACION=" & x2s(oCOTIZA) & ",IMPORTE=" & x2s(oIMPORTE) & ",ORDENPROVEEDOR=" & ssTexto(oORDENPROV) & ",RECIBIO=" & ssTexto(oRECIBIO) _
        & " WHERE CODIGO=" & oCODIGO
        DataEnvironment1.Sistema.Execute iud
End Select
Exit Function
omal:
ABMOrdenCompra = False
End Function

Public Function AItemOC(iPRODUCTO As String, iORDEN As Long, iCANTIDAD As Double, iSALDO As Double, iESTADO As String, iCOSTO As Double, iENTREGA As Date) As Boolean
On Error GoTo imal
Dim ca  As String
AItemOC = True
ca = " INSERT INTO ITEMORDENCOMPRA (PRODUCTO,ORDENCOMPRA,CANTIDAD,SALDO,ESTADO,COSTO,FECHAENTREGA) " _
                        & " VALUES (" & ssTexto(iPRODUCTO) & "," & iORDEN & "," & x2s(iCANTIDAD) & "," & x2s(iSALDO) & "," & ssTexto(iESTADO) & "," & x2s(iCOSTO) & "," & ssFecha(iENTREGA) & ")"
DataEnvironment1.Sistema.Execute ca
Exit Function
imal:
AItemOC = False
End Function

