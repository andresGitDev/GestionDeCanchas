VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPedidosClientes 
   Caption         =   "Carga de pedidos de Clientes"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11655
   Icon            =   "FrmPedidosClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11655
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmPedidosClientes.frx":08CA
      Left            =   5040
      List            =   "FrmPedidosClientes.frx":08DA
      TabIndex        =   47
      Top             =   3840
      Width           =   1575
   End
   Begin VSFlex7LCtl.VSFlexGrid grillaproductos 
      Height          =   2235
      Left            =   180
      TabIndex        =   43
      Top             =   4260
      Width           =   11295
      _cx             =   19923
      _cy             =   3942
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmPedidosClientes.frx":0902
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
   Begin Gestion.ucCuit uCuit 
      Height          =   315
      Left            =   7980
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1275
      _extentx        =   2249
      _extenty        =   556
   End
   Begin Gestion.ucCoDe uCliente 
      Height          =   315
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   6075
      _extentx        =   10716
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   3420
      Width           =   7515
      _extentx        =   13256
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin VB.CommandButton cmdBorraItem 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10980
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmPedidosClientes.frx":09D7
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "Borrar Item"
      Top             =   3660
      Width           =   495
   End
   Begin VB.CheckBox chkPropio 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo Propio"
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1755
   End
   Begin VB.ComboBox cmbvendedor 
      Height          =   315
      Left            =   9960
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1575
   End
   Begin VB.TextBox txtprecio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtobs 
      Height          =   555
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2100
      Width           =   6255
   End
   Begin VB.TextBox txtnropedidocli 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9300
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1740
      Width           =   1035
   End
   Begin VB.ComboBox cmbTransporte 
      Height          =   315
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   1380
      Width           =   1815
   End
   Begin VB.TextBox Txtcontacto 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1380
      Width           =   3255
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Width           =   4635
   End
   Begin VB.CommandButton cmdotro 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "FrmPedidosClientes.frx":0CE1
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3660
      Width           =   495
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3900
      Width           =   975
   End
   Begin VB.ComboBox cmbformapago 
      Height          =   315
      Left            =   9300
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   1380
      Width           =   2235
   End
   Begin VB.TextBox txtNro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   180
      Width           =   855
   End
   Begin VB.TextBox txttel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   1575
   End
   Begin VB.ComboBox cmbMoneda 
      Height          =   315
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1575
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2535
   End
   Begin VB.TextBox txtdireccionentrega 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1740
      Width           =   6255
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   300
      Left            =   9780
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58064897
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker dtfechaentrega 
      Height          =   300
      Left            =   8940
      TabIndex        =   17
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58064897
      CurrentDate     =   38098
   End
   Begin Gestion.ucBotonera ucMenu 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   1545
      Left            =   0
      TabIndex        =   42
      Top             =   6750
      Width           =   11655
      _extentx        =   20558
      _extenty        =   2725
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      Begin VB.Label lblTotalPedi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10500
         TabIndex        =   46
         Top             =   60
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Pedido:"
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
         Left            =   9000
         TabIndex        =   45
         Top             =   60
         Width           =   1395
      End
   End
   Begin VB.Label lblEntrgado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   44
      Top             =   3900
      Width           =   1155
   End
   Begin VB.Label Label15 
      Caption         =   "Vendedor :"
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
      Left            =   9360
      TabIndex        =   39
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label lblPrecio 
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
      Left            =   6900
      TabIndex        =   38
      Top             =   3900
      Width           =   735
   End
   Begin VB.Label lblunidad 
      Height          =   255
      Left            =   2700
      TabIndex        =   37
      Top             =   4500
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Obs:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "NroPedidoCliente :"
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
      Left            =   7500
      TabIndex        =   35
      Top             =   1740
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Transporte :"
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
      Left            =   4320
      TabIndex        =   34
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label Label17 
      Caption         =   "Contacto :"
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
      Left            =   120
      TabIndex        =   33
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio :"
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
      Left            =   120
      TabIndex        =   32
      Top             =   1020
      Width           =   1095
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
      Left            =   180
      TabIndex        =   31
      Top             =   3900
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   180
      X2              =   11460
      Y1              =   2940
      Y2              =   2940
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
      Left            =   180
      TabIndex        =   30
      Top             =   3420
      Width           =   975
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
      Left            =   9000
      TabIndex        =   29
      Top             =   120
      Width           =   675
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
      Left            =   7740
      TabIndex        =   28
      Top             =   1380
      Width           =   1455
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
      Left            =   120
      TabIndex        =   27
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Nro Pedido:"
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
      Left            =   120
      TabIndex        =   26
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Tel:"
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
      Left            =   9420
      TabIndex        =   25
      Top             =   660
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   6615
      Left            =   60
      Top             =   60
      Width           =   11535
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
      Left            =   7860
      TabIndex        =   24
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Localidad:"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label13 
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
      Left            =   120
      TabIndex        =   22
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Cuit :"
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
      Left            =   7500
      TabIndex        =   21
      Top             =   660
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   735
      Left            =   7740
      Top             =   2100
      Width           =   2775
   End
   Begin VB.Label Label16 
      Caption         =   "Fecha Entrega :"
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
      Left            =   8940
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "FrmPedidosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const gPRODUCTO = 0
Const gDESCRIPCION = 1
Const gCANTIDAD = 2
Const gPRECIO = 3
Const gFENTREGA = 4
Const gESTADO = 5
Const gSaldo = 6
Const gSALDOFAC = 7
  
Private sCampoExistencia As String

Private Sub chkPropio_Click()
    set_uProd
End Sub

Private Sub cmdBorraItem_Click()
    On Error Resume Next
    Dim i As Long
    i = grillaproductos.Row
    If i > 0 Then
        If s2n(grillaproductos.TextMatrix(i, gSALDOFAC)) < s2n(grillaproductos.TextMatrix(i, gCANTIDAD)) Then
            If confirma("producto parcialmente entregado, elimina ?") Then
                grillaproductos.RemoveItem (grillaproductos.Row)
            End If
        Else
            grillaproductos.RemoveItem (grillaproductos.Row)
        End If
    End If
    
    chkPropioEnabled True
    CalculaTotal
End Sub

Private Sub cmdotro_Click()
    Dim i As Long
    Dim AuxAlias As String
    Dim a, aFactor As Double, aCargar As Double
    Dim rs As New ADODB.Recordset, codigomio As String, ssql As String, nuevoSaldo As Double, tengo As Double, queveo As Double, ssmsg As String, mjStock As Long
    
    If s2n(txtcantidad) <= 0 Then
        che "Especificar cantidad Mayor o igual a 0"
        Exit Sub
    End If
    If s2n(txtcantidad) - s2n(lblEntrgado) < 0 Then
        che "No puede poner cantidad menor a lo ya entregado"
        Exit Sub
    End If
    If Trim(uProd.DESCRIPCION) = "" Then
        che "Falta producto"
        Exit Sub
    End If
    
    AuxAlias = (uProd.codigo)
    mjStock = ManejaStock(AuxAlias)
    
    ssql = "select codigo, componente, cantidad from formulas where activo = 1 and codigo = '" & AuxAlias & "'"
    rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    nuevoSaldo = (s2n(txtcantidad) - s2n(lblEntrgado))
    queveo = QueHay(AuxAlias, True)
    
    With grillaproductos
        If .rows = 1 Then .rows = 2
        If .rows > 2 Or Trim$(.TextMatrix(1, 0)) > "" Then .rows = .rows + 1
        i = .rows - 1
        
        'Set a = Nothing
        'a = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(uProd.codigo))
        'If IsNull(a) Or IsEmpty(a) Then
            aFactor = 1
        'Else
        '    aFactor = a
        'End If
        aCargar = aFactor * s2n(txtcantidad)
        
        .TextMatrix(i, gPRODUCTO) = uProd.codigo
        .TextMatrix(i, gDESCRIPCION) = uProd.DESCRIPCION
        .TextMatrix(i, gCANTIDAD) = s2n(aCargar)
        .TextMatrix(i, gPRECIO) = s2n(txtprecio, 4)
        .TextMatrix(i, gFENTREGA) = dtfechaentrega.Value
        .TextMatrix(i, gESTADO) = ESTADO_ADEUDADO
        .TextMatrix(i, gSaldo) = nuevoSaldo
        .TextMatrix(i, gSALDOFAC) = nuevoSaldo
    End With
    uProd.clear
    txtcantidad = "0"
    lblEntrgado = s2n(0)
    txtprecio = "0"
    dtfechaentrega.Value = Date
    uProd.SetFocus
    chkPropioEnabled True
    CalculaTotal
End Sub

Private Sub Combo1_Click()
    If Combo1.Text = "Lista 1" Then
        txtprecio = obtenerDeSQL("select precio from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 2" Then
        txtprecio = obtenerDeSQL("select precio2 from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 3" Then
        txtprecio = obtenerDeSQL("select precio3 from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 4" Then
        txtprecio = obtenerDeSQL("select precio4 from producto where codigo='" & uProd.codigo & "'")
    End If
End Sub
Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_Load()
   
    CargaCombo cmbMoneda, "monedas", "descripcion", "codigo", ""
    CargaCombo cmbTransporte, "transportes", "descripcion", "codigo", ""
    CargaCombo cmbformapago, "formaspago", "descripcion", "codigo", ""
    CargaCombo cmbvendedor, "usuarios", "descripcion", "codigo", ""
    InicioGrilla
    dtFecha = Date
    
    HabilitoTxt False

    ucMenu.init True, True, True, True, True, "select * from pedidos_clientes where activo = 1 order by numero", DataEnvironment1.Sistema, True
    ucMenu.MsgConfirmaEliminar = "Desea Eliminar este pedido?"
    ucMenu.MsgConfirmaSalir = "Cerrar formulario ?"

    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    set_uProd
    
    GeneraExistenciaCalculada
    If gEMPR_FormulaEsVirtual Then
        sCampoExistencia = " ExistenciaCalculada "
    Else
        sCampoExistencia = " Existencia "
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub


Sub InicioGrilla()
    With grillaproductos
        .clear
        .FormatString = "^Codigo                                    |<Descripcion                                                                        |>Cantidad |    Precio      |^Fecha Entrega  | Estado | Saldo   | Saldo-prueba- | Formula  "
        .rows = 2
        .cols = 8 '7 '9 ????
        .ColHidden(gESTADO) = True
    End With
End Sub


Private Sub grillaProductos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = gCANTIDAD Or Col = gPRECIO Then CalculaTotal
End Sub

Private Sub grillaproductos_DblClick()
    On Error Resume Next
    Dim i As Long
    
    If ucMenu.estado <> ucbEditando Then Exit Sub
    
    With grillaproductos
        i = .Row
        If i = 0 Then Exit Sub
            
        If Trim$(uProd.DESCRIPCION) > "" Then
            If s2n(lblEntrgado) > 0 Then
                che "Hay un item cargado en la linea de edicion parcialmente entregado, no puedo sobreescribirlo"
                Exit Sub
            End If
            If Not confirma("Hay un item cargado en la linea de edicion, lo sobreescribe ?") Then
                Exit Sub
            End If
        End If
        
        uProd.codigo = Trim(.TextMatrix(i, gPRODUCTO))
        
        txtcantidad = .TextMatrix(i, gCANTIDAD)
        txtprecio = .TextMatrix(i, gPRECIO)
        dtfechaentrega = .TextMatrix(i, gFENTREGA)
        lblEntrgado = s2n(.TextMatrix(i, gCANTIDAD) - .TextMatrix(i, gSALDOFAC))
        .RemoveItem (grillaproductos.Row)
    
    End With
    CalculaTotal
End Sub

Private Sub CalculaTotal()
    Dim i As Long, tot As Double
    With grillaproductos
        For i = 1 To .rows - 1
            tot = tot + s2n(.TextMatrix(i, gCANTIDAD)) * s2n(.TextMatrix(i, gPRECIO), 4)
        Next i
    End With
    lblTotalPedi = s2n(tot)
End Sub


Private Sub txtCantidad_LostFocus()
    If Val(txtcantidad) <= 0 Then
        MsgBox "La cantidad debe ser mayor a 0", 48, "Atencion"
    End If
End Sub


Private Sub LimpioTxt()
    On Error Resume Next
    
    lblEntrgado = s2n(0)
    FrmBorrarTxt Me
    uCliente.codigo = 0
    uCuit.Text = ""
    uProd.clear
    
    txtNro = ""
    txttel = ""
    txtdireccion = ""
    txtlocalidad = ""
    Txtcontacto = ""
    dtFecha.Value = Date
    dtfechaentrega.Value = Date
    cmbTransporte.ListIndex = 0
    cmbformapago.ListIndex = 0
    cmbvendedor.ListIndex = 0
    txtdireccionentrega = ""
    txtnropedidocli = ""
    txtObs = ""
    cmbMoneda.ListIndex = 1
    chkPropio.Value = vbChecked
    txtcantidad = "0"
    txtprecio = "0.00"
    'grillaproductos.rows = 1
    
    InicioGrilla
    CalculaTotal
End Sub

Private Sub HabilitoTxt(habilito As Boolean)
    Dim bloqueo
    bloqueo = Not habilito
    
    txtNro.Locked = bloqueo
    uCliente.enabled = habilito
    uProd.enabled = habilito
    Txtcontacto.Locked = bloqueo
    dtFecha.enabled = Not bloqueo
    dtfechaentrega.enabled = Not bloqueo
    cmbTransporte.Locked = bloqueo
    cmbformapago.Locked = bloqueo
    cmbvendedor.Locked = bloqueo
    chkPropioEnabled (Not bloqueo)
    txtdireccionentrega.Locked = bloqueo
    txtnropedidocli.Locked = bloqueo
    txtObs.Locked = bloqueo
    cmbMoneda.Locked = bloqueo
    txtcantidad.Locked = bloqueo
    txtprecio.Locked = bloqueo
    cmbMoneda.enabled = habilito
    cmdBorraItem.enabled = habilito
    cmdotro.enabled = habilito
End Sub

Private Sub CargoDatosCliente()
    On Error Resume Next
    Dim tmp
    
    tmp = obtenerDeSQL("select codigo, direccion, localidad, contacto, direccion_comercial, telefono_comercial, formapago, transporte, vendedor , cuit from clientes where codigo = " & uCliente.codigo)
    
    txtdireccion = sSinNull(tmp(1))
    txtlocalidad = sSinNull(tmp(2))
    Txtcontacto = sSinNull(tmp(3))
    txtdireccionentrega = sSinNull(tmp(4))
    txttel = sSinNull(tmp(5))
    cmbformapago.ListIndex = BuscarenComboS(cmbformapago, ObtenerDescripcion("formaspago", nSinNull(tmp(6))))
    cmbTransporte.ListIndex = BuscarenComboS(cmbTransporte, ObtenerDescripcion("transportes", nSinNull(tmp(7))))
    cmbvendedor.ListIndex = BuscarenComboS(cmbvendedor, ObtenerDescripcion("usuarios", nSinNull(tmp(8))))
    uCuit.Text = sSinNull(tmp(9))
    
End Sub
Private Sub txtprecio_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txttel_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtDireccion_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtlocalidad_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcontacto_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtdireccionentrega_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtnropedidocli_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtobs_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcodprod_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtDescripcion_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub
Private Sub txtprecio_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub


Private Function ConCodigoPropio() As Boolean
    ConCodigoPropio = (chkPropio.Value = vbChecked)
End Function


Private Function altaItems() As Boolean
On Error GoTo aimal
altaItems = True
    Dim fechaEntrega As Date, estado As String, i, COD As String, cant As Double, costo As Double, saldo As Double, formula As String
    Dim InsItem As String
    'SE ELIMINA TODOS LOS ITEMS
    InsItem = "delete from ItemPedidoCliente where pedido = " & x2s(txtNro)
    DataEnvironment1.Sistema.Execute InsItem
    
    'SE CARGA DE VUELTA TODOS LOS ITEM
    For i = 1 To grillaproductos.rows - 1
        grillaproductos.Row = i
        grillaproductos.Col = 0
        If ConCodigoPropio() Then
            COD = grillaproductos.Text
        Else
            COD = obtenerDeSQL("select producto from relacion_producto_cliente where productocliente = '" & grillaproductos.Text & "'")
        End If
        grillaproductos.Col = 2:        cant = CDbl(grillaproductos.Text)
        grillaproductos.Col = 3:        costo = CDbl(grillaproductos.Text)
        grillaproductos.Col = 4:        fechaEntrega = grillaproductos.Text
        grillaproductos.Col = gESTADO:  estado = grillaproductos.Text
        grillaproductos.Col = gSaldo: saldo = grillaproductos.Text
'        grillaProductos.col = gFORMULA: formula = grillaProductos.Text
        
        InsItem = "INSERT INTO ITEMPEDIDOCLIENTE (PEDIDO,PRODUCTO, CANTIDAD, FACTURAR,SALDO,ESTADO,PRECIO,FORMULA, FECHAENTREGA) " _
                & " VALUES (" & x2s(txtNro) & "," & ssTexto(COD) & "," & x2s(cant) & "," & x2s(cant) & ", " & x2s(saldo) & "," & ssTexto(estado) & "," & x2s(costo) & "," & ssTexto(formula) & "," & ssFecha(fechaEntrega) & ")"
        DataEnvironment1.Sistema.Execute InsItem
    Next i
Exit Function
aimal:
altaItems = False
End Function

Private Sub leerPedido(cual)
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset, ssql As String, tra As String
    
    Dim i As Long
    
    LimpioTxt
    
    ssql = "select * from pedidos_clientes where numero = " & cual
    
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        txtNro = !numero
        uCliente.codigo = !cliente
        'txtcliente = ObtenerDescripcion("clientes", !cliente)
        dtFecha = !Fecha
        cmbformapago = ObtenerDescripcion("formasPago", !pago)
        cmbvendedor = ObtenerDescripcion("usuarios", !Vendedor)
        txtnropedidocli = !pedido_cli
        chkPropio.Value = IIf(!CODIGOPROPIO, vbChecked, vbUnchecked)
        tra = sSinNull(ObtenerDescripcion("transportes", !Transporte))
        
        txtObs = !observaciones
    
        .Close
        CargoDatosCliente
        ssql = "select * from itemPedidoCliente where pedido = " & Val(txtNro) & " order by codigo"
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        i = 1
        While Not .EOF
            
            grillaproductos.rows = i + 1
            grillaproductos.TextMatrix(i, 0) = IIf(ConCodigoPropio(), !Producto, obtenerDeSQL("select productoCliente from Relacion_Producto_Cliente where producto = '" & !Producto & "' and cliente = " & uCliente.codigo))
            grillaproductos.TextMatrix(i, 1) = ObtenerDescripcionS("producto", !Producto)
            grillaproductos.TextMatrix(i, 2) = !Cantidad
            grillaproductos.TextMatrix(i, 3) = !precio
            grillaproductos.TextMatrix(i, 4) = !fechaEntrega
            grillaproductos.TextMatrix(i, gESTADO) = !estado
            grillaproductos.TextMatrix(i, gSaldo) = !saldo
            grillaproductos.TextMatrix(i, gSALDOFAC) = !facturar
'            grillaproductos.TextMatrix(i, gEXISTENCIA) = QueHay(!producto, True) ' EnExistencia(!producto)
'            grillaproductos.TextMatrix(i, gFORMULA) = !formula
            
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
    If Trim(tra) > "" Then cmbTransporte = tra
    Set rs = Nothing
    
    HabilitoTxt False
    CalculaTotal
End Sub

Private Sub chkPropioEnabled(que As Boolean)
    chkPropio.enabled = grillaproductos.rows < 2 And que
End Sub

Private Function buscaRelProdClie(cliente As String)
    Dim ssql As String
    ssql = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
        & " from producto  " _
        & " inner join relacion_Producto_Cliente " _
        & " on producto.codigo = relacion_Producto_cliente.producto " _
        & " where cliente = " & cliente _
        & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
        & " order by producto"
    frmBuscar.MostrarSql (ssql)
End Function

Private Function FaltanCosas() As Boolean
    Dim i As Long
    
    FaltanCosas = True
        
    If uCliente.codigo = 0 Then
        MsgBox "Falta cargar el cliente"
        uCliente.SetFocus
        Exit Function
    End If
        
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Function
    End If
    
    With grillaproductos
        If .rows < 2 Then
            che "faltan datos en grilla"
            Exit Function
        End If
        
        For i = 1 To .rows - 1
            If s2n(.TextMatrix(i, gCANTIDAD)) = 0 Then    ' no debe ser posible, pero...
                che "falta especificar cantidad en la grilla"
                Exit Function
            End If
            If Trim(.TextMatrix(i, gPRODUCTO)) = "" Then   ' no puede ser!!
                che "falta especificar producto en la grilla"
                Exit Function
            End If
        Next i
    End With

    FaltanCosas = False
End Function

Private Sub set_uProd()
    Dim sqlbuscar As String, sqldesc As String

    If ConCodigoPropio() Then    'propio
        sqldesc = "select descripcion from producto where codigo = '###' "
       sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    Else    'relCliente
        sqldesc = "select descripcion from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & uCliente.codigo & " and productoCliente = '###'"
        sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
            & " from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & uCliente.codigo _
            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
            & " order by producto"
    End If
    uProd.ini sqldesc, sqlbuscar, True
End Sub


Private Sub ucMenu_Imprimir()
    ImprimirPedido s2n(txtNro, 0)
End Sub

Private Sub uProd_cambio(codigo As Variant)
    txtprecio = precioProducto(CStr(codigo), ConCodigoPropio(), uCliente.codigo)
    Combo1.ListIndex = 0
End Sub
Private Sub uCliente_cambio(codigo As Variant)
    CargoDatosCliente
    set_uProd
End Sub

Private Function QueHay(quePro As String, Optional RestoReservado As Boolean)
    Dim hay As Double, reserva As Double, kk
    quePro = VerProductoMio(quePro, ConCodigoPropio())
    
    If RestoReservado Then
        QueHay = s2n(obtenerDeSQL("select (" & sCampoExistencia & " - ReservaCalculada) as quequeda from producto  where activo = 1 and codigo = '" & quePro & "' "))
    Else
        QueHay = s2n(obtenerDeSQL("select " & sCampoExistencia & " from producto  where activo = 1 and codigo = '" & quePro & "' "))
    End If
   
End Function

Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim np As Long
    
    If FaltanCosas() Then Exit Sub
    
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    
    np = nuevoCodigo("Pedidos_Clientes", "Numero")
    txtNro = np
    
    '**************************************
    DE_BeginTrans
    
        If ABMPedidoCliente("A", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtObs, 0, dtFecha, UsuarioSistema!codigo) = False Then GoTo ufaErr
        If altaItems = False Then GoTo ufaErr
    
    DE_CommitTrans
    '**************************************
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ImprimirPedido CDbl(np)
    ucMenu.AceptarOk '"Numero = " & Trim(txtNro)

Exit Sub
ufaErr:
    DE_RollbackTrans
    MsgBox "Error al guardar pedido", vbCritical
End Sub

 Private Sub ucMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    If FaltanCosas() Then Exit Sub

    DE_BeginTrans
    If ABMPedidoCliente("M", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtObs, 0, dtFecha, UsuarioSistema!codigo) = False Then GoTo ufaErr
    If altaItems = False Then GoTo ufaErr
    grabaBitacora "M", s2n(txtNro), "PedidosClientes"
    DE_CommitTrans
    
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk "Numero = " & Trim(txtNro)
Exit Sub
ufaErr:
    DE_RollbackTrans
    MsgBox "Error al guardar el pedido.", vbCritical
End Sub
Private Sub ucMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub ucMenu_Buscar()
    Dim s As String
    s = " select p.numero, p.Pedido_cli as PedidoClie, c.descripcion as [ Cliente                              ], p.fecha as [ Fecha      ], max(i.saldo) as pendiente, max(i.facturar) as [PendientePrueba] " & _
        " from pedidos_clientes as p inner join clientes as c on c.codigo = p.cliente inner join itempedidocliente as i on i.pedido = p.numero " & _
        " where p.activo = 1 " & _
        " group by numero, Pedido_cli, c.descripcion, fecha " & _
        " order by numero desc "
    
    If frmBuscar.MostrarSql(s) = "" Then Exit Sub
    leerPedido Val(frmBuscar.resultado(1))
    ucMenu.BuscarOK "numero = " & txtNro
End Sub
Private Sub ucMenu_BuscarYa(que As Variant)
    If Not IsEmpty(obtenerDeSQL("select numero from pedidos_clientes where pedidos_clientes.activo = 1 and numero = " & s2n(que))) Then
        leerPedido s2n(que)
        ucMenu.BuscarOK "numero = " & txtNro
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim sp
    sp = obtenerDeSQL("select * from remitoventadetalle d inner join remitoventa r on d.numero=r.numero where r.cancelado=0 and r.anulado=0 and d.pedido=" & s2n(txtNro))
    If IsNull(sp) Or IsEmpty(sp) Then
        Set sp = Nothing
        sp = obtenerDeSQL("select * from facturaventadetalle d inner join facturaventa f on d.nrofactura=f.nrofactura where d.nropedido=" & s2n(txtNro))
        If IsNull(sp) Or IsEmpty(sp) Then
        Else
            GoTo o
        End If
    Else
o:
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If
    
    ABMPedidoCliente "B", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtObs, 0, Date, UsuarioSistema!codigo
    grabaBitacora "B", s2n(txtNro), "Pedidos_Clientes"
    ucMenu.EliminarOK
fin:
    Exit Sub

ufaErr:
    ufa "error al eliminar", Me.Name & " " & txtNro ', Err
    Resume fin
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt sino
End Sub
Private Sub ucMenu_Nuevo()
    txtNro = nuevoCodigo("Pedidos_Clientes", "Numero")
    chkPropio.Value = vbChecked
    chkPropio.enabled = True
    uCliente.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Private Sub ucMenu_SeMovio()
    On Error Resume Next
    leerPedido ucMenu.rs!numero
End Sub

Public Function ABMPedidoCliente(pOPE As String, pNRO As Long, pCLIENTE As Long, pPAGO As Long, pVENDEDOR As Long, pPEDCLI As String, pPROPIO As Long, pTRANSPORTE As Long, pObs As String, pCANC As Long, pFecha As Date, pUsu As Long) As Boolean
On Error GoTo pcmal
Dim iud As String
ABMPedidoCliente = True
Select Case pOPE
    Case "A"
        iud = "INSERT INTO PEDIDOS_CLIENTES ( NUMERO,CLIENTE,FECHA,PAGO,VENDEDOR,PEDIDO_CLI,CODIGOPROPIO,TRANSPORTE,OBSERVACIONES,CANCELADO, FECHA_ALTA, USUARIO_ALTA,  ACTIVO) " _
            & " VALUES ( " & pNRO & "," & pCLIENTE & "," & ssFecha(pFecha) & "," & pPAGO & "," & pVENDEDOR & "," & ssTexto(pPEDCLI) & "," & pPROPIO & "," & pTRANSPORTE & "," & ssTexto(pObs) & "," & pCANC & "," & ssFecha(Date) & "," & pUsu & ", 1)"
        DataEnvironment1.Sistema.Execute iud
    Case "M"
        iud = " UPDATE PEDIDOS_CLIENTES " _
        & " SET CLIENTE=" & pCLIENTE & ",PAGO=" & pPAGO & ",VENDEDOR=" & pVENDEDOR & ",PEDIDO_CLI=" & ssTexto(pPEDCLI) & ",CODIGOPROPIO=" & pPROPIO & ",TRANSPORTE=" & pTRANSPORTE & ",OBSERVACIONES=" & ssTexto(pObs) & ",CANCELADO=" & pCANC & " WHERE NUMERO=" & pNRO
        DataEnvironment1.Sistema.Execute iud
        

    Case "B":
        iud = "UPDATE PEDIDOS_CLIENTES  SET ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA= " & pUsu & " WHERE NUMERO=" & pNRO
        DataEnvironment1.Sistema.Execute iud
End Select
Exit Function
pcmal:
ABMPedidoCliente = False
End Function

