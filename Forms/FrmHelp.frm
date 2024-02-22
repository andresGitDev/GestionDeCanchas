VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmHelp 
   Caption         =   "Ayuda"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6105
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   4270
      Width           =   3495
      Begin VB.TextBox txtbuscar 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         Picture         =   "FrmHelp.frx":08CA
         ScaleHeight     =   615
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtbuscdesc 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "xCodigo"
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
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "xNombre"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Picture         =   "FrmHelp.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Picture         =   "FrmHelp.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillahelp 
      Bindings        =   "FrmHelp.frx":17A8
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12632256
      BackColorBkg    =   14737632
      SelectionMode   =   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdaceptar_Click()

    Select Case Me.Tag
    Case "FrmPagosACuenta"
                grillahelp.Col = 0
                Select Case FrmPagosACuenta.cargar
                    Case "C": FrmPagosACuenta.TxtCodProv = grillahelp.Text
                    Case "BU":  FrmPagosACuenta.dtFecha = grillahelp.Text
                                grillahelp.Col = 1
                                FrmPagosACuenta.txttipocompra = grillahelp.Text
                                grillahelp.Col = 2
                                FrmPagosACuenta.txtopago = grillahelp.Text
                                grillahelp.Col = 3
                                FrmPagosACuenta.txtimporte = grillahelp.Text
                    Case "Cajas": FrmPagosACuenta.txtcodcaja = grillahelp.Text
                End Select
                FrmPagosACuenta.CargarDatos
        
        
'    Case "FrmSaldoCuentaProv"
'                grillahelp.Col = 0
'                FrmSaldoCuentaProv.TxtCodProv = grillahelp.Text
'                FrmSaldoCuentaProv.CargarDatos
                
    Case "FrmOrdenPago"
                grillahelp.Col = 0
                Select Case FrmOrdenPago.cargar
'                    Case "C": FrmOrdenPago.txtcodprov = grillahelp.Text
'                    Case "C": FrmOrdenPago.uProv.Codigo = s2n(grillahelp.Text)
                    Case "CuentasBank": FrmOrdenPago.txtcodcuenta = grillahelp.Text
                    Case "Cajas": FrmOrdenPago.txtcodcaja = grillahelp.Text
                End Select
                FrmOrdenPago.CargarDatos
                
    Case "FrmAjustes"
                grillahelp.Col = 0
                Select Case FrmAjustes.cargar
                    Case "Prov": FrmAjustes.TxtCodProv = grillahelp.Text
                    Case "Motivo": FrmAjustes.txtcodmotivo = grillahelp.Text
'                    Case "Cuentas": FrmAjustes.grillahelp.Text
                    Case "BU":  grillahelp.Col = 2
                                FrmAjustes.txtnrodoc = grillahelp.Text
                End Select
                FrmAjustes.CargarDatos

       
    Case "FrmExtractoBanc"
                grillahelp.Col = 0
                If FrmExtractoBanc.cargo = "Ctasbank1" Then
                    FrmExtractoBanc.txtcodcta1 = grillahelp.Text
                Else
                    FrmExtractoBanc.txtcodcta2 = grillahelp.Text
                End If
                FrmExtractoBanc.CargarDatos

'    Case "FrmListadoChequesTerceros"
'                grillahelp.Col = 0
'                FrmListadoChequesTerceros.txtcodcli = grillahelp.Text
'                FrmListadoChequesTerceros.CargarDatos
'
'    Case "FrmFactProv"
'                grillahelp.Col = 0
'                Select Case FrmFactProv.cargar
''                Case "C": FrmFactProv.txtcodprov = grillahelp.text
''                Case "CuentasBank": FrmFactProv.txtcodcuenta = grillahelp.text
''                Case "Cajas": FrmFactProv.txtcodcaja = grillahelp.text
''                Case "BU":  FrmFactProv.txtfechacompra = grillahelp.text
''                            grillahelp.col = 1
''                            FrmFactProv.txttipocompra = grillahelp.text
''                            grillahelp.col = 2
''                            FrmFactProv.txtnumcompra = grillahelp.text
'                End Select
'                FrmFactProv.CargarDatos

    Case "FrmListadoCheques"
                grillahelp.Col = 0
                FrmListadoCheques.txtcodcuenta = grillahelp.Text
                FrmListadoCheques.CargarDatos
    Case "FrmProcesoChTerceros"
                grillahelp.Col = 0
                FrmProcesoChTerceros.txtcodcuenta = grillahelp.Text
                FrmProcesoChTerceros.CargarDatos
    Case "FrmFormasPagos"
                grillahelp.Col = 0
                FrmFormasPagos.txtCodigo = grillahelp.Text
                grillahelp.Col = 1
                FrmFormasPagos.txtDescripcion = grillahelp.Text
                FrmFormasPagos.CargarDatos
    Case "FrmFactProv"
                grillahelp.Col = 0
                FrmFactProv.txtfechacompra = grillahelp.Text
                grillahelp.Col = 1
                FrmFactProv.txttipocompra = grillahelp.Text
                grillahelp.Col = 2
                FrmFactProv.txtnumcompra = grillahelp.Text
                FrmFactProv.CargarDatos
    Case "FrmIngChequesTerceros"
                grillahelp.Col = 0
                Select Case FrmIngChequesTerceros.cargar
                    'Case "Bancos": FrmIngChequesTerceros.txtcodbanco = grillahelp.Text
                    'Case "Clientes": FrmIngChequesTerceros.txtcodcli = grillahelp.Text
                End Select
                FrmIngChequesTerceros.Tag = grillahelp.Text
                FrmIngChequesTerceros.CargarDatos
    Case "FrmTransfBanc"
                grillahelp.Col = 0
                Select Case FrmTransfBanc.cargar
                    Case "Cuentao": FrmTransfBanc.txtcodctao = grillahelp.Text
                    Case "Cuentad": FrmTransfBanc.txtcodctad = grillahelp.Text
                End Select
                FrmTransfBanc.Tag = grillahelp.Text
                FrmTransfBanc.CargarDatos
    Case "FrmCtasBancarias"
                grillahelp.Col = 0
                Select Case FrmCtasBancarias.cargar
'                    Case "Cuentas": FrmCtasBancarias.txtCodCuenta = grillahelp.Text
                    Case "Tipocuentas": FrmCtasBancarias.txtTipo = grillahelp.Text
                    Case "Bancos": FrmCtasBancarias.txtcodbanco = grillahelp.Text
'                    Case "cuentas": FrmCtasBancarias.txtCodCuenta = grillahelp.Text
                    Case Else: FrmCtasBancarias.txtCodigo = grillahelp.Text
                End Select
                FrmCtasBancarias.Tag = grillahelp.Text
                FrmCtasBancarias.CargarDatos
'    Case "FrmGastosBancarios"
'                grillahelp.col = 0
'                Select Case FrmGastosBancarios.cargar
''                    Case "CuentasBank": FrmGastosBancarios.txtcodcta = grillahelp.Text
''                    Case "Cuentas": FrmGastosBancarios.txtcodcuenta = grillahelp.Text
'                    Case Else: 'FrmGastosBancarios.txtcodcta = grillahelp.Text
'                               grillahelp.col = 2
'                               FrmGastosBancarios.txtmovbank = grillahelp.Text
'                End Select
'                FrmGastosBancarios.Tag = grillahelp.Text
'                FrmGastosBancarios.CargarDatos
    Case "FrmIngEgrEfectivo"
                'grillahelp.Col = 0
                'Select Case FrmIngEgrEfectivo.cargar
                '    Case "Cajas": FrmIngEgrEfectivo.txtcodcaja = grillahelp.Text
                '    Case "Clientes": FrmIngEgrEfectivo.txtcodcli = grillahelp.Text
                '    Case "Proveedor": FrmIngEgrEfectivo.txtcodcli = grillahelp.Text
                '    Case "Deposito": FrmIngEgrEfectivo.txtcodcli = grillahelp.Text
'               '     Case "Cuentas": FrmIngEgrEfectivo.txtcodcuenta = grillahelp.Text
                '    Case Else: FrmIngEgrEfectivo.txtmovimiento = grillahelp.Text
                'End Select
                'FrmIngEgrEfectivo.Tag = grillahelp.Text
                'FrmIngEgrEfectivo.CargarDatos
    Case "FrmLisProveedores"
                grillahelp.Col = 0
                Select Case FrmLisProveedores.cargar
                    Case "ProvDesde": FrmLisProveedores.txtdesde = grillahelp.Text
                    Case "ProvHasta": FrmLisProveedores.txthasta = grillahelp.Text
                End Select
                FrmLisProveedores.Tag = grillahelp.Text
                FrmLisProveedores.CargarDatos
''    Case "FrmLibracionCheques"
''                grillahelp.col = 0
''                Select Case FrmLibracionCheques.cargar
''                    Case "Cheques": FrmLibracionCheques.txtint = grillahelp.Text
'''                    Case "Prov": FrmLibracionCheques.txtcod = grillahelp.Text
'''                    Case "Cajas": FrmLibracionCheques.txtcod = grillahelp.Text
'''                    Case "Cuentas": FrmLibracionCheques.txtcuentacod = grillahelp.Text
''                    Case "Movimientos": FrmLibracionCheques.txtint = grillahelp.Text
''                End Select
''                FrmLisProveedores.Tag = grillahelp.Text
''                FrmLibracionCheques.CargarDatos
    Case "FrmIngresoChequera"
                grillahelp.Col = 0
                FrmIngresoChequera.txtcodcta = grillahelp.Text
                FrmIngresoChequera.CargarDatos
    Case "FrmProveedor"
                grillahelp.Col = 0
                FrmProveedor.txtCodigo = grillahelp.Text
                FrmProveedor.CargarDatos
'    Case "FrmClientes"
'                grillahelp.Col = 0
'                FrmClientes.txtCodigo = grillahelp.Text
'                FrmFormasPagos.CargarDatos
    Case "FrmAbmTiposUsuarios"
                grillahelp.Col = 0
                FrmAbmTiposUsuarios.txtCodigo = grillahelp.Text
                FrmAbmTiposUsuarios.CargarDatos
    Case "FrmCategorias"
                grillahelp.Col = 0
                FrmCategorias.txtCodigo = grillahelp.Text
                FrmCategorias.CargarDatos
    Case "FrmAbmUsuarios"
                grillahelp.Col = 0
                FrmAbmUsuarios.txtCodigo = grillahelp.Text
                FrmAbmUsuarios.CargarDatos
    Case "FrmBancosGenerales"
                grillahelp.Col = 0
                FrmBancosGenerales.txtCodigo = grillahelp.Text
                FrmBancosGenerales.CargarDatos
    Case "FrmTransportes"
                grillahelp.Col = 0
                FrmTransportes.txtCodigo = grillahelp.Text
                FrmTransportes.CargarDatos
    Case "FrmZonas"
                grillahelp.Col = 0
                FrmZonas.txtCodigo = grillahelp.Text
                FrmZonas.CargarDatos
    Case "FrmMotivosAjuste"
                grillahelp.Col = 0
                FrmMotivosAjuste.txtCodigo = grillahelp.Text
                FrmMotivosAjuste.CargarDatos
    Case "frmAjusteDeCosto"
                grillahelp.Col = 0
                If frmAjusteDeCosto.cargar = "Cuentas" Then
                    frmAjusteDeCosto.txtcuentacod = grillahelp.Text
                Else
                    frmAjusteDeCosto.txtCodigo = grillahelp.Text
                End If
                frmAjusteDeCosto.Tag = grillahelp.Text
                frmAjusteDeCosto.CargarDatos
    Case "FrmCostosYContable"
                grillahelp.Col = 0
                If FrmCostosYContable.cargar = "Cuentas" Then
                    FrmCostosYContable.txtcuentacod = grillahelp.Text
                Else
                    FrmCostosYContable.txtCodigo = grillahelp.Text
                End If
                FrmCostosYContable.Tag = grillahelp.Text
                FrmCostosYContable.CargarDatos
    End Select
    Unload Me
End Sub

Private Sub cmdcancelar_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call cmdcancelar_Click
End Sub
Private Sub grillahelp_DblClick()
    cmdaceptar_Click
End Sub

Private Sub txtbuscar_Change()
    Dim r As Long
    Dim codigo As String
    Dim desc As String
    Dim i As Long
    
    r = 1
    For i = 1 To grillahelp.rows - 1
        grillahelp.Col = 0
        grillahelp.Row = i
        grillahelp.CellBackColor = vbWhite
        grillahelp.Col = 1
        grillahelp.CellBackColor = vbWhite
    Next i
    If txtbuscar <> "" Then
        grillahelp.Col = 0
        grillahelp.Row = r
        
        While LCase(Mid(grillahelp.Text, 1, Len(txtbuscar))) <> LCase(txtbuscar) And r < grillahelp.rows
            r = r + 1
            grillahelp.Row = r - 1
        Wend
        
        If r <= grillahelp.rows Then
             grillahelp.CellBackColor = vbRed
            
             If r = 1 Then
                grillahelp.TopRow = r
             Else
                grillahelp.TopRow = r - 1
             End If
         End If
    End If
End Sub

Private Sub txtbuscdesc_Change()
    Dim r As Long
    Dim codigo As String
    Dim desc As String
    Dim i As Long
    
    r = 1
    For i = 1 To grillahelp.rows - 1
            grillahelp.Col = 1
            grillahelp.Row = i
            grillahelp.CellBackColor = vbWhite
            grillahelp.Col = 0
            grillahelp.CellBackColor = vbWhite
    Next i
    If txtbuscdesc <> "" Then
        
        grillahelp.Col = 1
        grillahelp.Row = r
        
        While LCase(Mid(grillahelp.Text, 1, Len(txtbuscdesc))) <> LCase(txtbuscdesc) And r < grillahelp.rows
            r = r + 1
            grillahelp.Row = r - 1
        Wend
        If r <= grillahelp.rows Then
            
             grillahelp.CellBackColor = vbRed
             If r = 1 Then
                grillahelp.TopRow = r
             Else
                grillahelp.TopRow = r - 1
             End If
                      
        End If
    End If
End Sub


