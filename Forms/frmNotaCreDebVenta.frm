VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaCreDebVenta 
   Caption         =   "Emision Nota Credito"
   ClientHeight    =   8385
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   9960
   Icon            =   "frmNotaCreDebVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1635
      Left            =   0
      TabIndex        =   40
      Top             =   6750
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   2884
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   30
      TabIndex        =   9
      Top             =   -15
      Width           =   10080
      Begin VB.TextBox txtCotizacion 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   6900
         TabIndex        =   42
         Top             =   1785
         Width           =   1680
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   6900
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtPIVA10 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "10,5"
         Top             =   4545
         Width           =   615
      End
      Begin VB.TextBox txtIVA21 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8460
         TabIndex        =   37
         Top             =   4170
         Width           =   945
      End
      Begin VB.TextBox txtIVA10 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   8460
         TabIndex        =   36
         Top             =   4545
         Width           =   945
      End
      Begin VB.CommandButton cmbingresar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingresar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8490
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6225
         UseMaskColor    =   -1  'True
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Gastos Administrativos"
         Height          =   465
         Left            =   7500
         TabIndex        =   32
         ToolTipText     =   "Genera Resibo a Cuenta."
         Top             =   3105
         Width           =   1890
      End
      Begin Gestion.ucFecha uFechaBuscaCheque 
         Height          =   330
         Left            =   8295
         TabIndex        =   6
         Top             =   2220
         Width           =   975
         _ExtentX        =   2990
         _ExtentY        =   582
         FechaInit       =   0
      End
      Begin VB.TextBox txtIIBB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   23
         Top             =   5295
         Width           =   945
      End
      Begin Gestion.ucCoDe uCuenta 
         Height          =   315
         Left            =   1485
         TabIndex        =   7
         Top             =   2685
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5670
         Width           =   945
      End
      Begin Gestion.ucCoDe uCheques 
         Height          =   315
         Left            =   1515
         TabIndex        =   5
         Top             =   2235
         Width           =   4605
         _ExtentX        =   6800
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   1020
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2910
         Left            =   1500
         TabIndex        =   8
         Top             =   3270
         Width           =   5775
         _cx             =   10186
         _cy             =   5133
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
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   22
         Top             =   4920
         Width           =   945
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3825
         Width           =   945
      End
      Begin VB.TextBox txtPIVA21 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "21,0"
         Top             =   4170
         Width           =   615
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   1785
         Width           =   1455
      End
      Begin VB.ComboBox cmbTipoIva 
         Height          =   315
         Left            =   6900
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2415
      End
      Begin VB.TextBox TxtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   2580
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtCodigo 
         Height          =   320
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1035
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1380
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   5460
         TabIndex        =   0
         Top             =   540
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   38126
      End
      Begin VB.Label Label17 
         Caption         =   "Iva:"
         Height          =   255
         Index           =   4
         Left            =   7365
         TabIndex        =   45
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizacion:"
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
         Left            =   5760
         TabIndex        =   44
         Top             =   1815
         Width           =   975
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda:"
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
         Left            =   5925
         TabIndex        =   43
         Top             =   1515
         Width           =   900
      End
      Begin VB.Label Label17 
         Caption         =   "Total:"
         Height          =   255
         Index           =   3
         Left            =   7905
         TabIndex        =   39
         Top             =   4980
         Width           =   495
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Centro de Costos"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6975
         TabIndex        =   35
         Top             =   6270
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblChequeBusca 
         Caption         =   "Busca cheque desde:"
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
         Height          =   315
         Left            =   6135
         TabIndex        =   31
         Top             =   2250
         Width           =   2070
      End
      Begin VB.Label Label17 
         Caption         =   "IIBB:"
         Height          =   255
         Index           =   2
         Left            =   7365
         TabIndex        =   30
         Top             =   5340
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta:"
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
         Index           =   1
         Left            =   720
         TabIndex        =   29
         Top             =   2685
         Width           =   795
      End
      Begin VB.Label Label17 
         Caption         =   "Total:"
         Height          =   255
         Index           =   1
         Left            =   7290
         TabIndex        =   28
         Top             =   5670
         Width           =   495
      End
      Begin VB.Label lblCheque 
         Caption         =   "Nro Int Cheque"
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
         Height          =   315
         Left            =   165
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Iva:"
         Height          =   255
         Index           =   0
         Left            =   7365
         TabIndex        =   25
         Top             =   4215
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "Neto:"
         Height          =   255
         Left            =   7320
         TabIndex        =   24
         Top             =   3855
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Neto:"
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
         Left            =   960
         TabIndex        =   19
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo IVA:"
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
         Left            =   6900
         TabIndex        =   18
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1215
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
         Left            =   4620
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label9 
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
         Left            =   720
         TabIndex        =   15
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label14 
         Caption         =   "FormaPago:"
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
         Left            =   300
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
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
         Index           =   0
         Left            =   780
         TabIndex        =   10
         Top             =   1020
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmNotaCreDebVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '19/11/4

'Private WithEvents cliente As LiCodigo

Private mEXT As Boolean

Public Enum FTipoNota
    Tipo_NotaCredito
    Tipo_NotaDebito
    Tipo_NotaDebitoChRechazado
End Enum
Private mTipoNota As FTipoNota

Private gDESC As Long
Private g As LiGrilla
Private Const CANT_RENGLONES = 20

Private Sub cmbingresar_Click()
    If txttotal = "" Then
        MsgBox "Debe ingresar algun producto en la factura"
        Exit Sub
    End If
    'If txtDescuento = "" Then
    '    txtDescuento = 0
    'End If
    
'    FrmCostosYContable.txtimporte.Enabled = False
    FrmCostosYContable.cargar.enabled = False
'    FrmCostosYContable.txtcuentacod.Enabled = False
'    FrmCostosYContable.cmbcuenta.Enabled = False
'    FrmCostosYContable.txtcuenta.Enabled = False
    FrmCostosYContable.txtconc.enabled = False
    FrmCostosYContable.txtvalor.enabled = False
    FrmCostosYContable.cmdcargar.enabled = False
    FrmCostosYContable.cmbeliminofila.enabled = False
'    FrmCostosYContable.txtTotal.Enabled = False
'    FrmCostosYContable.Grilla.Enabled = False
    
    FrmCostosYContable.CargarImputacion s2n(txtneto) + s2n(txtIva), s2n(txttotal), 2
    'FrmCostosYContable.txtimptotal = txtimporte
    'FrmCostosYContable.txtimporte = txtNeto
    FrmCostosYContable.txtimporte = ""
    FrmCostosYContable.cargar = ""
    FrmCostosYContable.txtcuentacod = ""
    FrmCostosYContable.txtcuenta = ""
    FrmCostosYContable.txtconc = ""
    FrmCostosYContable.txtvalor = ""
    FrmCostosYContable.txttotal = ""
'    FrmCostosYContable.grilla.Clear
    
    FrmCostosYContable.Tag = Me.Name
    vieneDE = Me.Name
    FrmCostosYContable.Show
    
End Sub

Private Sub Command1_Click()
    frmAsientoManual.Show
End Sub

'Private Sub cmdAceptar_Click()
'    Dim tmpfec As Date, tipoForm As String
'
'    If s2n(txtCodigo) > 0 Then
'        MsgBox "Ya fue Grabada"
'        Exit Sub
'    End If
'
'    If s2n(txtMonto) = 0 Then
'        MsgBox "Monto no Valido"
'        Exit Sub
'    End If
'
''    tipoformu = obtenerDeSQL("select letra from ")
''    'alta
'    txtCodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
'    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then    'If ComboCodigo(cmbTipoIva) = IVA_ConsumidorFinal Then
'        txtTipoDoc = TipoDoc_NCREDITO_B
'        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_B) + 1
'        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_B)
'    Else
'        txtTipoDoc = TipoDoc_NCREDITO_A
'        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_A) + 1
'        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_A)
'    End If
'
'
'    If tmpfec > dtFecha Then
'        MsgBox "Fecha menor que la ultima factura cargada"
'        Exit Sub
'    End If
'
'
'    If confirma("Nro Factura: " & TxtNroFactura) Then
'        If GrabaFactura() Then
''            ucBoton.AceptarOk
'            MsgBox "Nro Factura: " & TxtNroFactura & vbCrLf & "Grabado"
'            ImprimirComprobante s2n(txtCodigo)
'        End If
'    Else
'        txtCodigo = ""
'        TxtNroFactura = ""
'    End If
'
'End Sub

'Private Sub cmdReImprimir_Click()
''    If s2n(txtCodigo) = 0 Then Exit Sub
''    ImprimirComprobante (s2n(txtCodigo))
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
'    rever
End Sub

Private Sub Form_Load()
    Set g = New LiGrilla

'    CentrarMe Me
    fra.Top = 0
    fra.Left = 0
    fra.Height = Me.ScaleHeight
    fra.Width = Me.ScaleWidth
    
    dtFecha = Date

    comboSql cmbformapago, "select descripcion, codigo from formaspago where activo = 1"
    comboSql cmbTipoIva, "select descripcion, codigo from ivas"
    
    uCliente.ini "select descripcion from clientes where codigo = ### ", "select codigo as [ Codigo    ], descripcion as [ Cliemte                                              ] from clientes order by descripcion"
    g.init Grilla
    gDESC = g.AddCol("Descripcion" & Space(90), "S")
    g.rows = CANT_RENGLONES
    Grilla.EditMaxLength = 74
    
    'le saque el boton imprimir Hasta q haya busqueda
    uMenu.init False, True, False, False, False
    'uCheques.ini "select importe from cheques where NroInt = ### and estado = 'R' and activo = 1", "select NroInt, fecha, cliente, Importe from cheques where activo = 1 and estado = 'R' order by NroInt desc", False
    uCheques.ini "select importe from cheques where NroInt = ### and (estado = 'T' or estado = 'C' or estado = 'R' ) and activo = 1", "select NroInt, fecha, cliente, Importe, estado from cheques where activo = 1 and (estado = 'T' or estado = 'C') and fecha > " & uFechaBuscaCheque.ssFecha & " order by NroInt desc", False
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    comboSql cboMoneda, "select descripcion, codigo from monedas order by codigo"
    cboMoneda.ListIndex = 0
    txtPIVA21 = "21"
    txtIIBB = 0
End Sub


Private Function GrabaFactura() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    

    Dim asse As String ' assert
    Dim ac As Variant, i As Long, tota As Double, iddoc As Long
    Dim AsientoVenta As New Asiento, tipo As String
    Dim TextoAsientoComprobante As String
    Dim totND As Double
    Dim Neto As Double
    Dim cant As Double
    Dim Valor As Integer, P21 As Double, P10 As Double
    Dim z As Double
    
    GrabaFactura = False
    'tota = s2n(txtTotal)
    z = s2n(txtCotizacion, 4)
    If z = 0 Then z = 1
    tota = s2n(txttotal)
    
    If mEXT = True Then
        tota = s2n(txttotal * s2n(txtCotizacion.Text, 4))
    Else
        tota = s2n(txttotal)
    End If
 
 
'*** una transaccion aqui ..... *********************
    DE_BeginTrans
    
    asse = " Actualizo tabla parametros BS "
    
    
'''    AumentarParametroN CAMPO_BS_CodFactura_VENTA, s2n(txtcodigo)
''    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "A" Then
''        AumentarParametroD CAMPO_BS_FecFACTURA_A, dtfecha
''        AumentarParametroN CAMPO_BS_NroFACTURA_A, s2n(TxtNroFactura)
''    ElseIf TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then
''        AumentarParametroD CAMPO_BS_FecFACTURA_B, dtfecha
''        AumentarParametroN CAMPO_BS_NroFACTURA_B, s2n(TxtNroFactura)
''    Else
''        ufa "PrgErr: TipoDoc No reconocido", Me.Name ', Err
''    End If
    'tipo = IIf(Left(txtTipoDoc, 2) = "NC", "N.Credito venta", "N.Debito venta")
    iddoc = NuevoDocumento(txtTipoDoc, TxtNroFactura, 0, 0)
    
    'DEBERIA ESTAR EN CABECERA ASIENTO
    
    If mEXT = True Then
        tipo = IIf(Left(txtTipoDoc, 2) = "NC", "N.Credito venta exterior", "N.Debito venta exterior")
    Else
        tipo = IIf(Left(txtTipoDoc, 2) = "NC", "N.Credito venta", "N.Debito venta")
    End If
    
    TextoAsientoComprobante = tipo & TxtNroFactura
    
    
    AsientoVenta.nuevo tipo & " " & uCliente.DESCRIPCION, dtFecha, txtTipoDoc
    
    
    
    asse = "Graba Cabecera "
    ac = obtenerDeSQL("select provincia, cuit from clientes where codigo = " & uCliente.codigo)
    

    
    asse = "Graba detalle ND ch rech"

    If mTipoNota <> Tipo_NotaDebitoChRechazado Then
        
        If s2n(txtPIVA21) > 0 Then
            P21 = s2n(txtPIVA21 / 100, 4)
        Else
            P21 = 0
        End If
        If s2n(txtPIVA10) > 0 Then
            P10 = s2n(txtPIVA10 / 100, 4)
        Else
            P10 = 0
        End If
        
        DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtFecha, dtFecha, ComboCodigo(cmbformapago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n((txtMonto) * z), s2n(P21 + P10, 4), s2n(txtIva * z), tota * z, tota * z, 0, 0, 0, UsuarioActual(), Date, z, ComboCodigo(cboMoneda), s2n(txtIIBB * z), 0, 0, 0, 0, 0, 0, iddoc
        'DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtfecha, dtfecha, ComboCodigo(cmbFormaPago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n(txtMonto), s2n(txtPIVA21) - Valor, s2n(txtIva), tota, tota, 0, 0, 0, UsuarioActual(), Date, 0, 0, s2n(txtIIBB), 0, 0, 0, 0, 0, 0, iddoc
    Else
        'If s2n(txtPIVA) = 0 Then
        '    Valor = 0
        'Else
        '    Valor = 1
        'End If
        If s2n(txtPIVA21) > 0 Then
            P21 = s2n(txtPIVA21 / 100, 4)
        Else
            P21 = 0
        End If
        If s2n(txtPIVA10) > 0 Then
            P10 = s2n(txtPIVA10 / 100, 4)
        Else
            P10 = 0
        End If
        
        'DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtfecha, dtfecha, ComboCodigo(cmbFormaPago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n(txtMonto), s2n(txtPIVA) - Valor, s2n(txtIva), tota, tota, 0, 0, 0, UsuarioActual(), Date, 0, 0, s2n(txtIIBB), 0, 0, 0, 1, 0, s2n(txtTotal), iddoc
        DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtFecha, dtFecha, ComboCodigo(cmbformapago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n(txtMonto), s2n(P21 + P10, 4), s2n(txtIva), tota, tota, 0, 0, 0, UsuarioActual(), Date, 0, 0, s2n(txtIIBB), 0, 0, 0, 1, 0, s2n(txttotal), iddoc
        
     
  
        
        DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, 0, True, "", DatosCheque(), "", 0, 0, 0, 0, 0, 0, iddoc
        If s2n(txttotal) > 0 Then
            DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, 0, True, "", "Gastos gravados  " & x2s(txttotal) & "", "", 0, 0, 0, 0, 0, 0, iddoc
        End If
        
'        DataEnvironment1.Sistema.Execute _
'            "update cheques set estado = 'R' where nroint = " & uCheques.codigo
                    
            
            
        Dim d_q_cuenta
        d_q_cuenta = obtenerDeSQL("select dep_cuenta from cheques where nroint =" & uCheques.codigo)
        If d_q_cuenta = 0 Then
            If MsgBox("El cheque no tiene cuenta de deposito." & Chr(13) & "Por favor indique una a continuacion, gracias.", vbInformation + vbYesNo) = vbYes Then
            d_q_cuenta = frmBuscar.MostrarSql("select c.codigo as [CODIGO], c.banco as [BANCO - Nº],b.descripcion as  [NOMBRE  ],c.numero as [CUENTA - Nº] from ctasbank c inner join bancosgrales b on c.banco=b.codigo where c.activo=1", , "Cuentas bancarias", " - ")
            End If
        End If
            
        DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", d_q_cuenta, "R", "Rechazo de Cheque", dtFecha, "C" _
          , uCheques.codigo, x2s(txttotal), nuevoCodigo("movibanc", "movbanco"), iddoc, Date, UsuarioSistema!codigo
        
        ' con iddoc tengo un problema: son 3 operaciones: entra, quizas sale, rechazo.
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", uCheques.codigo, 0, "", 0, 0, "", 0, dtFecha, "R", 0, "", 0, Date, 0, 0, iddoc
    End If
        
    asse = "Graba detalle text"
    For i = 1 To g.rows - 1
        If Trim(g.tx(i, gDESC)) > "" Then
            If Trim(g.tx(i + 1, gDESC)) > "" Then
                Neto = 0
                cant = 0
            Else
                Neto = s2n(txtMonto)
                cant = 1
            End If
            DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, cant, True, uCheques.codigo, Trim(g.tx(i, gDESC)), "", Neto * z, Neto * z, 0, 0, 0, 0, iddoc
        End If
    Next i
    
        Dim tiene_c, CUENTA_C As String
        tiene_c = obtenerDeSQL("select tiene_cuenta from clientes where codigo = " & uCliente.codigo)
        If tiene_c = 1 Then
            CUENTA_C = obtenerDeSQL("select cuenta from clientes where codigo = " & uCliente.codigo)
        Else
            CUENTA_C = CuentaParam(ID_Cuenta_V_DEUDxVENTAS)
        End If
    
    If mEXT Then
        If mTipoNota = Tipo_NotaCredito Then 'credito
            AsientoVenta.AgregarItem uCuenta.codigo, s2n(txtneto * s2n(z, 4)), 0, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(txtIva * s2n(z, 4)), 0, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(txtIIBB * s2n(z, 4)), 0
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), 0, s2n(txttotal * s2n(z, 4)), TextoAsientoComprobante
        ElseIf mTipoNota = Tipo_NotaDebito Then 'debito
            totND = s2n(txtMonto * s2n(txtCotizacion, 4))
                        
            AsientoVenta.AgregarItem uCuenta.codigo, 0, totND, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva * s2n(z, 4)), TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(txtIIBB * s2n(z, 4))
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), s2n(txttotal * s2n(z, 4)), 0, TextoAsientoComprobante
        Else
            totND = s2n(txtMonto)
            If mTipoNota = Tipo_NotaDebitoChRechazado Then totND = totND + uCheques.DESCRIPCION
            
            AsientoVenta.AgregarItem uCuenta.codigo, 0, totND, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva), TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(txtIIBB)
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), s2n(txttotal), 0, TextoAsientoComprobante
        End If
    Else
        If mTipoNota = Tipo_NotaCredito Then 'credito
            AsientoVenta.AgregarItem uCuenta.codigo, s2n(txtneto * z, 4), 0, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(txtIva * z, 4), 0, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(txtIIBB * z, 4), 0
            AsientoVenta.AgregarItem CUENTA_C, 0, s2n(txttotal * z, 4), TextoAsientoComprobante
        Else 'debito
            ' = s2n(TxtTotal)
    
            'If mTipoNota = Tipo_NotaDebitoChRechazado Then
                totND = s2n(totND + txtMonto)
            'Else
            '    totND = s2n(totND + txtMonto)
            'End If
            
            AsientoVenta.AgregarItem uCuenta.codigo, 0, totND * z, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva * z, 4), TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(txtIIBB * z, 4)
            AsientoVenta.AgregarItem CUENTA_C, s2n(txttotal * z, 4), 0, TextoAsientoComprobante
        End If
    End If
    'If
    If siAsiento("AsientosVentas") Then AsientoVenta.Grabar iddoc
    '= 0 Then
'        DE_RollbackTrans
'        ufa "Err al grabar asiento", " " & iddoc
'        GoTo fin
'    End If
   
    DE_CommitTrans
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura = True
    GoTo fin
    
UfaGraba:
    DE_RollbackTrans
    ufa "Err al grabar ", Me.Name & " - grabaFactura() - " & asse ', Err
fin:

End Function

Private Function DatosCheque() As String
    On Error Resume Next
    Dim tmp
    tmp = obtenerDeSQL("select nro, fecha, importe, descripcion from cheques inner join bancosGrales on cheques.banco_nro = BancosGrales.codigo where cheques.activo = 1 and cheques.NroInt = " & uCheques.codigo)
    'LO SIG LO HAGO POR EN TMP(0) SE GUARDA EL NUMERO DE CHEQUE, PERO NECESITO QUE TENGA EL TOTAL
'    tmp(0) = TxtTotal
    DatosCheque = "Cheque Numero " & tmp(0) & Format(tmp(1), "dd/mm/yy") & " " & tmp(3) & "  " & x2s(tmp(2))
End Function

Private Sub txtIIBB_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtIIBB_LostFocus()
    rever
    RefrescoDeValores
    If Label5.caption = "A" Then
        txtIva21.Text = 21
    Else
        'txtPIVA.Text = 0
        'txtTotal = txtTotal - txtIva
        'txtIva = 0
    End If
End Sub

Private Sub txtiva_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtiva_LostFocus()
    txtIva = s2n(txtIva)
    rever
End Sub

Private Sub RefrescoDeValores() '*********modificado el 16/5/7***raul
    Dim nrocheque As Variant
    Dim montoCheque As Variant
    Dim montoCHsinIva As Variant
    
    'EN UCHEQUES.DESCRIPCION SE COLOCABA EL MONTO PERO EL MONTO YA ESTABA EN OTRO FOCO
    'POR ESO UTILISE UCHEQUES.DESCRIPCION PARA QUE TENGA EL NUMERO DE CHEQUE
    'COMO EL CONTROL DEVUELTE EL MONTO Y SI LO MODIFICO CAMBIO A TODOS LO QUE LO UTILISEN
    'CON LO SIG HAGO COMO UN REFRESCO DE LO QUE QUIERO QUE TENGA
    If mTipoNota = Tipo_NotaDebitoChRechazado Then
        If (uCheques.codigo > 0) Then
            nrocheque = obtenerDeSQL("select nro from cheques where nroint = " & uCheques.codigo)
            uCheques.EditaDescripcion = True
            uCheques.DESCRIPCION = (nrocheque)
            montoCheque = s2n(obtenerDeSQL("select importe from cheques where nroint = " & uCheques.codigo), 2)
            If s2n(txtIva21) > 0 Then
                montoCHsinIva = montoCheque / (1 + (s2n(txtIva) / 100))
            Else
                montoCHsinIva = montoCheque
            End If
            txtMonto = s2n(montoCHsinIva)
            txtneto = s2n(montoCHsinIva)
        End If
    End If
End Sub




Private Sub txtiva21_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtiva10_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtMonto_GotFocus()
    frmPintoFoco Me
    rever
    RefrescoDeValores
    If Label5.caption = "A" Then
        'txtPIVA21.Text = 21
        'txtPIVA10.Text = 0
    Else
        'txtPIVA21.Text = 0
        'txtPIVA10.Text = 0
    End If
End Sub

Private Sub txtMonto_LostFocus()
    rever
    RefrescoDeValores
    If Label5.caption = "A" Then
        'txtPIVA21.Text = 21
        'txtPIVA10.Text = 0
    Else
        'txtIva.Text = 0
        'txttotal = txttotal - txtIva
        'txtIva = 0
    End If
End Sub

Private Sub rever() '******modificado el 16/5/7***raul
    Dim valorCiva As Double
    Dim valorSiva As Double
    Dim valorIva21 As Double, valorIva10 As Double
    Dim mCh As Double
    
    If mTipoNota = Tipo_NotaDebitoChRechazado Then
        'COPIO EL TOTAL DEL CHEQUE, ESTE VALOR ES SACADO DE LA BD
        'valorCiva = s2n(obtenerDeSQL("select importe from cheques where nroint = " & uCheques.codigo), 2) 'IIf(uCheques.codigo = 0, 0, s2n(txtMonto, 4)) ESTO ESTABA ANTES
        valorSiva = s2n(obtenerDeSQL("select importe from cheques where nroint = " & uCheques.codigo), 2) 'IIf(uCheques.codigo = 0, 0, s2n(txtMonto, 4)) ESTO ESTABA ANTES
        
        'VALIDO EL VALOR DEL CHEQUE PARA PONERLE LA COMA
        valorSiva = s2n(valorSiva, 2)
        
        'CALCULO EL VALOR DEL CHEQUE SIN IVA
        If s2n(txtPIVA21) > 0 Then
            valorIva21 = s2n(valorSiva * (s2n(txtPIVA21) / 100))
            txtIva21 = s2n(valorIva21)
        End If
        If s2n(txtPIVA10) > 0 Then
            valorIva10 = s2n(valorSiva * (s2n(txtPIVA10) / 100))
            txtIva10 = s2n(valorIva10)
        End If
                
        'If txtPIVA = "" Then txtPIVA = 1
        'If txtPIVA = 0 Then txtPIVA = 1
        'valorSiva = valorCiva / s2n(txtPIVA, 2) ' permito mods manuales'If s2n(TxtIVA) = 0 Then txtIVA = valorCiva / s2n(txtPIVA, 4) ' permito mods manuales
        'txtPIVA = "1,21"
        
        'CON ESTO SIEMPRE CALCULO EL VALOR DEL IVA
        txtIva = s2n(txtIva21) + s2n(txtIva10) 's2n(valorCiva - valorSiva) 'If s2n(TxtIVA) = 0 Then TxtIVA = valorCiva - valorSiva ESTO ESTABA ANTES
        
        'ACA CALCULO EL TOTAL SIN EL VALOR DEL IVA
        txtneto = s2n(valorSiva) ' - txtIva, 2) ' s2n(txtMonto, 4) + mCh
        
        'Y POR ULTIMO RECALCULO EL TOTAL SUMANDO EL VALOR SIN IVA + EL IVA + LOS INGRESOS BRUTOS
        'txttotal = s2n(valorSiva) + s2n(txtIva, 2) + s2n(txtIIBB) 's2n(s2n(txtMonto, 2) + s2n(TxtIVA, 2) + mCh + s2n(txtIIBB)) ESTO ESTABA ANTES
        txttotal = s2n(valorSiva) + s2n(txtIva, 2) + s2n(txtIIBB)
    End If
    
    If mTipoNota = Tipo_NotaDebito Then
        valorSiva = s2n(txtMonto)
        txtMonto = s2n(txtMonto)
        'txtPIVA = "1,21"
        If s2n(txtPIVA21) > 0 Then
            valorIva21 = s2n(valorSiva * (s2n(txtPIVA21) / 100))
            txtIva21 = s2n(valorIva21)
        End If
        If s2n(txtPIVA10) > 0 Then
            valorIva10 = s2n(valorSiva * (s2n(txtPIVA10) / 100))
            txtIva10 = s2n(valorIva10)
        End If
        txtIva = s2n(valorIva21 + valorIva10)
        valorCiva = s2n(valorSiva + txtIva)
        txtneto = s2n(valorSiva)
        txttotal = s2n(valorSiva) + s2n(txtIva) + s2n(txtIIBB)
    End If
            
    If mTipoNota = Tipo_NotaCredito Then '******modificado el 16/5/7***raul
        valorSiva = s2n(txtMonto)
        txtMonto = s2n(txtMonto)
        If s2n(txtPIVA21) > 0 Then
            valorIva21 = s2n(valorSiva * (s2n(txtPIVA21) / 100))
            txtIva21 = s2n(valorIva21)
        End If
        If s2n(txtPIVA10) > 0 Then
            valorIva10 = s2n(valorSiva * (s2n(txtPIVA10) / 100))
            txtIva10 = s2n(valorIva10)
        End If
        
        
        'txtPIVA = "1,21"
        'valorCiva = s2n(valorSiva * txtPIVA)
        txtIva = s2n(valorIva21 + valorIva10)
        'valorSiva = s2n(valorSiva) - s2n(txtIva)
        txtneto = s2n(valorSiva) '
        txttotal = s2n(valorSiva) + s2n(txtIva) + s2n(txtIIBB)
    End If

    
    
    
End Sub


Private Sub txtPIVA_LostFocus()
    rever
End Sub


Private Sub txtneto_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub





Private Sub txtPIVA10_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtPIVA10_LostFocus()
txtPIVA10 = s2n(txtPIVA10)
rever
End Sub

Private Sub txtPIVA21_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub


Private Sub txtPIVA21_LostFocus()
txtPIVA21 = s2n(txtPIVA21)
rever
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub uCheques_Buscar()
    Dim ss As String
'    ss = "select NroInt, Fecha, Cliente, Importe from cheques where estado = 'R' and activo = 1 "
'    ESTADO T=TRANSFERIDO  C=CARTERA
    ss = "select NroInt, fecha, cliente, Importe, estado from cheques where (activo = 1) and (estado = 'T' or estado = 'C')  and  fecha > " & uFechaBuscaCheque.ssFecha
    If uCliente.codigo > 0 Then ss = ss & " and cliente = " & uCliente.codigo
    ss = ss & " order by NroInt desc"
    uCheques.strSqlBuscar = ss
End Sub

Private Sub uCheques_cambio(codigo As Variant)
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    Dim tclie As Long
    Dim nrocheque As Long
    Dim montoCheque As Variant

    'If uCheques.codigo > 0 Then
        'busca el numero de cliente que tenga ese numero de cheque
        tclie = obtenerDeSQL("select cliente from cheques where nroint = " & uCheques.codigo)
        'busca el numero de cheque(este es el numero con el que viene el cheque) correspondiente a ese cheque(ucheques.codigo es igual al numero interno que se le da)
        nrocheque = obtenerDeSQL("select nro from cheques where nroint = " & uCheques.codigo)
        'busca el monto de ese cheque
        montoCheque = obtenerDeSQL("select importe from cheques where nro = " & nrocheque)


        If tclie > 0 Then
            If tclie <> uCliente.codigo Then
                If uCliente.codigo > 0 Then che "el cheque corresponde a otro cliente"
                uCliente.codigo = tclie
            End If
        End If
    'End If

    RefrescoDeValores
fin:
End Sub

Private Sub uCheques_LostFocus()
    rever
    'RefrescoDeValores
End Sub

Private Sub uCliente_cambio(codigo As Variant)
    On Error Resume Next
    Dim ac As Variant
    Dim sCtas As String, sWhe As String, ss As String
    ac = obtenerDeSQL("select iva, FormaPago from clientes where codigo = " & codigo)
   
    cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, ac(1))
    cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, ac(0))
    'txtPIVA = s2n(obtenerDeSQL("select porcentaje from porcentajesiva where activo = 1 and iva =  " & ComboCodigo(cmbTipoIva)))
    'EL VALOR ANTERIOR LO SACABA DE LA BASE DE DATOS PERO NO ERA CORRECTO EL VALOR, ENTONCES SE LO PASO POR CODIGO
    'txtPIVA = "1,21"
    txtIIBB = 0
    
    
    rever
    
    If mEXT = True Then
        Select Case mTipoNota
        Case Tipo_NotaCredito
            txtTipoDoc = "NCE"
        Case Tipo_NotaDebito
            txtTipoDoc = "NDE"
        Case Tipo_NotaDebitoChRechazado
            txtTipoDoc = "ND"
        End Select
    Else
'        Select Case mTipoNota
'        Case Tipo_NotaCredito
'            txtTipoDoc = "NC"
'        Case Tipo_NotaDebito
'            txtTipoDoc = "ND"
'        Case Tipo_NotaDebitoChRechazado
'            txtTipoDoc = "ND"
'        End Select
    End If
    
    Label5.caption = TipoFormVenta(ComboCodigo(cmbTipoIva))
    If Label5.caption = "A" Then
        txtPIVA21.Text = 21
        txtPIVA10.Text = 0
    Else
        txtPIVA21.Text = 0
        txtPIVA10.Text = 0
    End If
    sWhe = ""
    If s2n(uCliente.codigo) > 0 Then
        sCtas = sSinNull(obtenerDeSQL("select cuentasventas from clientes where codigo=" & uCliente.codigo))
        If sCtas = "" Then
        Else
            sWhe = " and cuenta in (" & Replace(sCtas, "#", "'") & ")"
        End If
        'ss = ss & sWhe
    End If
    ss = "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 " & sWhe & "order by cuenta "
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", ss, True
    
    BuscoNroYTipo
End Sub

Private Function BuscoNroYTipo() As Boolean
    Dim tmpfec, letra As String
    Dim a As String
    BuscoNroYTipo = True
    Dim ss As String, andtipo As String, tmp
    If mEXT = True Then
        letra = "E"
    Else
        letra = TipoFormVenta(ComboCodigo(cmbTipoIva))
    End If

    If letra = "B" Then
        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_B & "' or TipoDoc = '" & TipoDoc_FACTURA_B & "'  or TipoDoc = '" & TipoDoc_NDEBITO_B & "' ) "
    ElseIf letra = "A" Then
        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_A & "' or TipoDoc = '" & TipoDoc_FACTURA_A & "'  or TipoDoc = '" & TipoDoc_NDEBITO_A & "' ) "
    ElseIf letra = "E" Then
        If CORTO(txtTipoDoc, 1, 1) = "D" Then
            'andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_E & "' or TipoDoc = '" & TipoDoc_FACTURA_E & "'  or TipoDoc = '" & TipoDoc_NDEBITO_E & "' ) "
            andtipo = " ( TipoDoc = '" & TipoDoc_NDEBITO_E & "' ) "
        Else
            andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_E & "'   ) "
        End If
            
    Else
        ufa "prg: No se encontro letra doc para Tipo Iva :" & cmbTipoIva, Me.Name ', 0
        BuscoNroYTipo = False
        Exit Function
    End If
    
    a = "select max(NroFactura) from FacturaVenta where " & andtipo
    ' si vacio, lo lleno
    If Trim(TxtNroFactura) = "" Then TxtNroFactura = nSinNull(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo)) + 1

    BuscoNroYTipo = RevisaNroYFechaOk("FacturaVenta", "NroFactura", "Fecha", s2n(TxtNroFactura, 0), dtFecha, andtipo) ' Then Exit Sub
'    'existe factura
'    ss = "select codigo from facturaVenta where  NroFactura = " & TxtNroFactura & andtipo
'    tmp = obtenerDeSQL(ss)
'    If Not IsEmpty(tmp) Then
'        che "Factura Existente con el codigo interno  " & tmp
'        uMenu.SetFocus
'        BuscoNroYTipo = False
'        Exit Function
'    End If
'
'    Dim maxfac, maxfe, minfe, minfac
'    'fecha factura menor mas alta
'    ss = "select max(NroFactura) from FacturaVenta where activo = 1 and NroFactura < " & TxtNroFactura & andtipo
'    maxfac = obtenerDeSQL(ss)
'    ss = "select Fecha from FacturaVenta where NroFactura = " & maxfac & andtipo
'    maxfe = CDate(obtenerDeSQL(ss))
'    If dtFecha < maxfe Then
'        che " Fecha Factura " & dtFecha & " menor que de factura " & maxfac & " " & maxfe
'        BuscoNroYTipo = False
'        Exit Function
'    End If
'
'    'fecha factura mayor mas baja
'    ss = "select min(NroFactura) from FacturaVenta where activo = 1 and NroFactura > " & TxtNroFactura & andtipo
'    minfac = obtenerDeSQL(ss)
'    If IsNull(minfac) Then Exit Function
'
'    ss = "select Fecha from FacturaVenta where NroFactura = " & minfac & andtipo
'    minfe = (obtenerDeSQL(ss))
'
'    minfe = CDate(minfe)
'    If dtFecha > minfe Then
'        che " Fecha Factura " & dtFecha & " mayor que de factura " & minfac & " " & minfe
'        BuscoNroYTipo = False
'        Exit Function
'    End If
End Function



Public Sub mostrar(que As FTipoNota, Optional Exterior As Boolean = False)
    mTipoNota = que
    mEXT = Exterior
    If mEXT Then
        Select Case que
        Case Tipo_NotaCredito
            Me.caption = "Emision Nota Credito Exterior"
            txtTipoDoc = "NC"
            'fraCheque.Visible = False
            'uCheques.Visible = False
            'lblCheque.Visible = False
            txtMonto.Visible = True
            cboMoneda.Visible = True
            Label21.Visible = True
            Label8.Visible = True
            txtCotizacion.Visible = True
            txtCotizacion = ""
            vercheqes False
        Case Tipo_NotaDebito 'ver no modificado para exterior
            Me.caption = "Emision Nota Debito Exterior"
            txtTipoDoc = "ND"
            'fraCheque.Visible = False
            'uCheques.Visible = False
            'lblCheque.Visible = False
            cboMoneda.Visible = True
            txtCotizacion.Visible = True
            txtCotizacion = ""
            vercheqes False
        Case Tipo_NotaDebitoChRechazado 'ver no modificado para exterior
            Me.caption = "Emision Nota Debito Cheque Rechazado"
            txtTipoDoc = "ND"
            'fraCheque.Visible = True
            'uCheques.Visible = True
            'lblCheque.Visible = True
            cboMoneda.Visible = False
            txtCotizacion.Visible = False
            txtCotizacion = "1"
            vercheqes True
        End Select
    Else
        Select Case que
        Case Tipo_NotaCredito
            Me.caption = "Emision Nota Credito"
            txtTipoDoc = "NC"
            'fraCheque.Visible = False
            'uCheques.Visible = False
            'lblCheque.Visible = False
            cboMoneda.Visible = True
            txtCotizacion.Visible = True
            txtCotizacion = ""
            vercheqes False
        Case Tipo_NotaDebito
            Me.caption = "Emision Nota Debito"
            txtTipoDoc = "ND"
            'fraCheque.Visible = False
            'uCheques.Visible = False
            'lblCheque.Visible = False
            cboMoneda.Visible = True
            txtCotizacion.Visible = True
            txtCotizacion = ""
            vercheqes False
        Case Tipo_NotaDebitoChRechazado
            Me.caption = "Emision Nota Debito Cheque Rechazado"
            txtTipoDoc = "ND"
            'fraCheque.Visible = True
            'uCheques.Visible = True
            'lblCheque.Visible = True
            cboMoneda.Visible = False
            txtCotizacion.Visible = False
            txtCotizacion = "1"
            vercheqes True
        End Select
    End If
    Me.Show
End Sub

Private Sub vercheqes(sino As Boolean)
        uCheques.Visible = sino
        lblCheque.Visible = sino
        lblChequeBusca.Visible = sino
        uFechaBuscaCheque.Visible = sino
End Sub




'----------------------MENU -----------------------
Private Sub uMenu_AceptarAlta()
    Dim tmpfec As Date, tipoForm As String, andtipo  As String
    Dim sAssert As String
    Dim x As Long

    If Not PuedoVentas(dtFecha) Then
        'msg en funcion
        Exit Sub
    End If

    If s2n(txttotal) = 0 Then
        che "Monto no Valido"
        txtMonto.SetFocus
        Exit Sub
    End If
    If mTipoNota = Tipo_NotaDebitoChRechazado And uCheques.codigo = 0 Then
        che "Falta cheque"
        uCheques.SetFocus
        Exit Sub
    End If
    If uCuenta.codigo = "" Then
        che "falta cuenta contable"
        Exit Sub
    End If
    'If txtCotizacion = "" And cboMoneda.Text <> "Pesos" Then
    '    MsgBox "Debe ingresar una cotizacion.", , "ATENCION"
    '    Exit Sub
    'End If
    
    If mEXT = True Then
        If txtCotizacion = "" And cboMoneda.Text <> "Pesos" Then
            MsgBox "Debe ingresar una cotizacion.", , "ATENCION"
            Exit Sub
        End If
    End If
    
'    'alta
    'txtcodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
    txtCodigo = nuevoCodigo("FacturaVenta", "codigo")
    
    If mEXT = True Then
        If mTipoNota = Tipo_NotaCredito Then
            txtTipoDoc = TipoDoc_NCREDITO_E
        Else
            txtTipoDoc = TipoDoc_NDEBITO_E
        End If
    Else
        If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then
            If mTipoNota = Tipo_NotaCredito Then
                txtTipoDoc = TipoDoc_NCREDITO_B
            Else
                txtTipoDoc = TipoDoc_NDEBITO_B
            End If
    '        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_B & "' or TipoDoc = '" & TipoDoc_FACTURA_B & "'  or TipoDoc = '" & TipoDoc_NDEBITO_B & "' ) "
        Else
            If mTipoNota = Tipo_NotaCredito Then
                txtTipoDoc = TipoDoc_NCREDITO_A
            Else
                txtTipoDoc = TipoDoc_NDEBITO_A
            End If
    '        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_A & "' or TipoDoc = '" & TipoDoc_FACTURA_A & "'  or TipoDoc = '" & TipoDoc_NDEBITO_A & "' ) "
        End If
    End If
    'If Not BuscoNroYTipo() Then Exit Sub
'    If Not RevisaNroYFechaOk("FacturaVenta", "NroFactura", "Fecha", s2n(TxtNroFactura, 0), dtFecha, andtipo) Then Exit Sub
    If Not BuscoNroYTipo() Then
        MsgBox "Se actualizara el numero de comprobante, para que pueda grabar.", , "ATENCION"
        TxtNroFactura = TxtNroFactura + 1
        Exit Sub   ' de nuevo, por las dudas (si fuera multiusuario habria q meter mas control aun)
    End If
    
    If confirma("Nro de Nota: " & TxtNroFactura) Then
        If GrabaFactura() Then
            
            If gEMPR_ConSistContable Then
                If FrmCostosYContable.grillacostos.rows > 1 And FrmCostosYContable.grillacostos.TextMatrix(1, 1) > "" Then
                    sAssert = " dbo_INGCENTROCOSTOS "
                    
                    'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                    For x = 1 To FrmCostosYContable.grillacostos.rows - 1
                        DataEnvironment1.dbo_INGCENTROCOSTOS "A", Val(FrmCostosYContable.grillacostos.TextMatrix(x, 0)), _
                        dtFecha, txtTipoDoc, Val(TxtNroFactura), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) + s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, "", FrmCostosYContable.grillacostos.TextMatrix(x, 4), 0
                        
                    Next
                    FrmCostosYContable.LimpioControles
                    FrmCostosYContable.InicioGrillaCostos
                    Unload FrmCostosYContable
                End If
            End If
            
            MsgBox "Nro Factura: " & TxtNroFactura & vbCrLf & "Grabado"
'            If mTipoNota = Tipo_NotaCredito Then
'                ImprimirComprobante s2n(TxtCodigo) 'si es nota de credito
'            Else
'                ImprimirComprobante s2n(TxtCodigo) 'si es nota de debito
'            End If
            
            If gEMPR_idEmpresa = 11 Then
                ImprimirComprobThor (s2n(txtCodigo))
            ElseIf gEMPR_idEmpresa = 6 And (Trim(txtTipoDoc) = "FAE" Or Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A) Then
                'imprimo para amr alvarez thomas
                ImprimirAMRAT (s2n(txtCodigo))  '******** esto esta listo solo tienen q avisar para usarlo
            ElseIf gEMPR_idEmpresa = 4 And (Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A) Then
                
                ImprimirAMRAT s2n(txtCodigo), True
'                If MsgBox("Desea imprimir el triplicado?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
'                    ImprimirAMRAT (s2n(txtCodigo)), True
'                End If
            ElseIf gEMPR_idEmpresa = 1 Then
                ImprimirComprobanteLOC (s2n(txtCodigo))
            Else
                ImprimirComprobante (s2n(txtCodigo))
            End If
        
            uMenu.AceptarOk
        End If
'    Else
'        txtCodigo = ""
'        TxtNroFactura = ""
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    FrmBorrarCbo Me
    uCliente.codigo = 0
    uCheques.clear
    uCuenta.clear
    g.Borrar
    g.rows = CANT_RENGLONES
    If cboMoneda.ListCount > 0 Then
        cboMoneda.ListIndex = 0
    End If
End Sub
Private Sub uMenu_Buscar()
'
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fra.enabled = sino
End Sub

Private Sub uMenu_Imprimir()
    If s2n(txtCodigo) = 0 Then Exit Sub
'    ImprimirComprobante (s2n(txtCodigo))
    
    If gEMPR_idEmpresa = 11 Then
        ImprimirComprobThor (s2n(txtCodigo))
    ElseIf gEMPR_idEmpresa = 6 And (Trim(txtTipoDoc) = "FAE" Or Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A) Then
        'imprimo para amr alvarez thomas
        ImprimirAMRAT (s2n(txtCodigo))  '******** esto esta listo solo tienen q avisar para usarlo
    ElseIf gEMPR_idEmpresa = 4 And (Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A) Then
        
        ImprimirAMRAT s2n(txtCodigo), True
'        If MsgBox("Desea imprimir el triplicado?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
'            ImprimirAMRAT (s2n(txtCodigo)), True
'        End If
    ElseIf gEMPR_idEmpresa = 1 Then
        ImprimirComprobanteLOC (s2n(txtCodigo))
    Else
        ImprimirComprobante (s2n(txtCodigo))
    End If
End Sub
Private Sub uMenu_Nuevo()
    'dtFecha.SetFocus
    'uCliente.SetFocus
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
'----------------------MENU -----------------------


'19/11/4
'    adapt licodigo, +cmd, +where
'16/12/4
'   ucLiCode cliente,
'   add grilla descripcion
'   unifico frm debito/ credito,
'17/12/4
'   foco en cliente
'28/2/5
'   null
'18/4/5
'   cheque rechazado
'20/5/5
'   nro factura manual, verifica fechas y correlatividad
'26/5/5
'   codigo FV desde tabla FacturaVenta no BS
'31/5/5
'   fix montos neto, tot,  (y cheque)
'   iva editable
'   pone nro al principio
'
