VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReiniciarSist 
   Caption         =   "Reiniciar Sistema"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   Icon            =   "frmReiniciarSist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBase 
      Height          =   285
      Left            =   5610
      TabIndex        =   13
      Top             =   2505
      Width           =   1755
   End
   Begin VB.CheckBox chkAchique 
      Caption         =   "Achicar base"
      Height          =   315
      Left            =   5595
      TabIndex        =   12
      Top             =   2205
      Value           =   1  'Checked
      Width           =   1740
   End
   Begin VB.CheckBox chkProveedores 
      Caption         =   "Reiniciar Proveedores"
      Height          =   315
      Left            =   3525
      TabIndex        =   11
      Top             =   2220
      Width           =   2010
   End
   Begin VB.CheckBox chkClientes 
      Caption         =   "Reiniciar Clientes"
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2220
      Width           =   1620
   End
   Begin VB.TextBox txtDireccion 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   1035
      Width           =   5535
   End
   Begin VB.TextBox txtCuit 
      Height          =   345
      Left            =   5250
      TabIndex        =   4
      Top             =   600
      Width           =   2085
   End
   Begin VB.TextBox txtEmpresaCorto 
      Height          =   345
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2085
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   360
      Left            =   1800
      TabIndex        =   2
      Top             =   165
      Width           =   5535
   End
   Begin MSComctlLib.ProgressBar barProgreso 
      Height          =   630
      Left            =   1800
      TabIndex        =   1
      Top             =   1500
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   1111
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   2e-4
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   1065
      Left            =   180
      Picture         =   "frmReiniciarSist.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1455
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Cuit"
      Height          =   240
      Left            =   4860
      TabIndex        =   9
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      Height          =   240
      Left            =   45
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Empresa (Nombre Corto)"
      Height          =   240
      Left            =   30
      TabIndex        =   7
      Top             =   645
      Width           =   2070
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de Empresa"
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   225
      Width           =   1575
   End
End
Attribute VB_Name = "frmReiniciarSist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcesar_Click()
Dim nBase As String
    Exec "delete from MAYOR"
    Exec "DBCC CHECKIDENT(mayor, RESEED, 0)"
    Exec "delete from Asientos"
    Exec "DBCC CHECKIDENT(Asientos, RESEED, 0)"
    Exec "delete from Bitacora"
    Exec "delete from BSuLog"
    Exec "delete from CHEQUES"
    Exec "delete from CHQ_COMP"
    Exec "delete from Clientesaa"
    Exec "delete from ClientesFER"
    Exec "delete from CODIGO"
    Exec "delete from COMPRAS"
    Exec "delete from ComprasDetalle"
    Exec "delete from ComprasRetenciones"
    Exec "delete from ConceptoOP"
    Exec "delete from ConceptoOPdetalle"
    Exec "delete from ConceptoPAGO"
    Exec "delete from Conceptos"
    Exec "delete from Consignatarios"
    Exec "delete from ConsultaStock"
    Exec "delete from Deposito"
    Exec "delete from exAsientos"
    Exec "delete from exMAYOR"
    Exec "delete from FacturaCompraRemito"
    Exec "delete from FacturaRetencion"
    Exec "delete from FacturasSucursal"
    Exec "delete from FacturasSucursalDetalle"
    Exec "delete from FacturaVenta"
    Exec "delete from FacturaVentaDetalle"
    Exec "delete from Formulas"
    Exec "delete from GastoBankTemp"
    Exec "delete from GruposProducto"
    Exec "delete from Hoja1$"
    Exec "delete from IMPPRO"
    Exec "delete from Impr_Factura"
    Exec "delete from Imputaciones"
    Exec "delete from INGRESOCHTEMP"
    Exec "delete from ItemOrdenCompra"
    Exec "delete from ItemPartesProduccion"
    Exec "delete from ItemPedidoCliente"
    Exec "delete from ItemRemitoDiferenciaStock"
    Exec "delete from LIST_MOV_CUENTA_CLI"
    Exec "delete from LIST_MOV_CUENTA_PROV"
    Exec "delete from LIST_MOV_CUENTA_PROV_DET"
    Exec "delete from LIST_SALDO_CLI"
    Exec "delete from LISTADOMOVIMIENTO"
    Exec "delete from Listas"
    Exec "delete from MoviBanc"
    Exec "delete from MoviCaja"
    Exec "delete from MOVIMIENTO_STOCK_TEMP"
    Exec "delete from OBRAS"
    Exec "delete from OrdenesdeCompras"
    Exec "delete from ORDENPAGOTEMP"
    Exec "delete from PartesProduccion"
    Exec "delete from Pedidos_Clientes"
    Exec "delete from REC_COMP"
    Exec "delete from Recibos"
    Exec "delete from RecibosDetalle"
    Exec "delete from RecibosRetenciones"
    Exec "delete from RegistroDocumentos"
    Exec "DBCC CHECKIDENT(RegistroDocumentos, RESEED, 0)"
    Exec "delete from RELFNR_C"
    Exec "delete from RemitoCompra"
    Exec "delete from RemitoCompraDetalle"
    Exec "delete from RemitoDiferenciaStock"
    Exec "delete from RemitoVenta"
    Exec "delete from RemitoVentaDetalle"
    Exec "delete from RemitoVentaDetalleSuc"
    Exec "delete from RemitoVentaSuc"
    Exec "delete from Resultados"
    Exec "delete from SaldoProvTMP"
    Exec "delete from Series"
    Exec "delete from Tabla1"
    Exec "delete from TEMP_CONTROL_CALIDAD"
    Exec "delete from TEMP_Etiquetas"
    Exec "delete from TmpEnvioARTIC"
    Exec "delete from TmpEnvioBAJAS"
    Exec "delete from TmpEnvioMAE_VEND"
    Exec "delete from TmpEnvioRE0000"
    Exec "delete from TRANSCOM"
    If chkClientes Then
        Exec "delete from Clientes"
    Else
        barProgreso.Max = barProgreso.Max - 1
    End If
    If chkProveedores Then
        Exec "delete from Prov"
    Else
        barProgreso.Max = barProgreso.Max - 1
    End If
    Exec "update DatosEmpresa set nombre='" & Trim(txtEmpresa) & "'"
    Exec "update DatosEmpresa set nombrecortoparalistados='" & Trim(txtEmpresaCorto) & "'"
    Exec "update DatosEmpresa set direccion='" & Trim(txtDireccion) & "'"
    Exec "update DatosEmpresa set cuitempresa='" & Trim(txtCuit) & "'"
    If chkAchique And Trim(txtBase) <> "" Then
        nBase = Trim(txtBase)
        Exec "use " & nBase & ""
        Exec "checkpoint"
        Exec "Exec sp_addumpdevice 'disk','" & nBase & "2','c:\logAchique.bak'"
        Exec "backup database " & nBase & " to " & nBase & "2"
        Exec "backup log " & nBase & " with truncate_only"
        Exec "dbcc shrinkfile(" & nBase & "_log, 100)"
    Else
        barProgreso.Max = barProgreso.Max - 6
    End If
    MsgBox "Terminado.", vbInformation
End Sub

Private Function Exec(comando As String)
On Error Resume Next
If Trim(comando) = "" Then
Else
    DataEnvironment1.Sistema.Execute comando
End If
If barProgreso.Value = barProgreso.Max Then barProgreso.Value = 0.0001: barProgreso.Max = 91.0001
barProgreso.Value = barProgreso.Value + 1
End Function

Private Sub Form_Load()
barProgreso.Value = 0.0001
barProgreso.Max = 91.0001
End Sub

