VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmIngEgrEfectivo 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ingreso / Egreso de Cajas y Bancos"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcuentacaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   45
      Tag             =   "5"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcotiz 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   43
      Tag             =   "1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtmoneda 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   41
      Tag             =   "1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbcotizacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cotizaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   39
      Tag             =   "5"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   37
      Tag             =   "8"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   36
      Tag             =   "2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmbcaja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caja"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtcodcaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Tag             =   "1"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtvalor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtconc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox txtcuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   31
      Tag             =   "2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmbcuenta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuenta"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtcodcuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdcargar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "9"
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmbeliminofila 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar Fila"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton optdeposito 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Depósito"
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6240
      TabIndex        =   27
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton optegreso 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Egreso"
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4440
      TabIndex        =   26
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton optingreso 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingreso"
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Tag             =   "0"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   975
   End
   Begin VB.TextBox txtcodcli 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Tag             =   "5"
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmbcambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtcliente 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   16
      Tag             =   "2"
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtmovimiento 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "0"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtimporte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "4"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtconcepto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1560
      Width           =   5655
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1560
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   95092737
      CurrentDate     =   38052
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmIngEgrEfectivo.frx":0000
      Height          =   2535
      Left            =   360
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      Enabled         =   0   'False
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      MergeCells      =   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).BandIndent=   10
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   9840
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblcotiz 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cotización:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblmoneda 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   6765
      Left            =   120
      Top             =   120
      Width           =   9840
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6120
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Concepto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblcambio 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº Movimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Concepto/Resp.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº Caja:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "FrmIngEgrEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4  frmcotizacion.cotizacion ?


Dim rsefec As New ADODB.Recordset
Dim Ope As String
Dim modifico As Boolean
Dim numero As Long

Private Sub cmbcaja_Click()
    cargar = "Cajas"
    FrmHelp.Show
    CargarHelp "Cajas", "Codigo", "Descripción", "codigo", "responsable"
    FrmHelp.Tag = Me.Name
End Sub

Private Sub cmbcaja_GotFocus()
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub cmbcambio_Click()
    If optdeposito = True Then
        FrmHelp.Show
        CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
        FrmHelp.Tag = Me.Name
        cargar = "Deposito"
    End If
    
    If optingreso = True Then
        FrmHelp.Show
        CargarHelp "Clientes", "Codigo", "Descripcion", "codigo", "descripcion"
        FrmHelp.Tag = Me.Name
        cargar = "Clientes"
    End If
    
    If optegreso = True Then
        FrmHelp.Show
        CargarHelp "Prov", "Codigo", "Descripcion", "codigo", "descripcion"
        FrmHelp.Tag = Me.Name
        cargar = "Proveedor"
    End If
        
End Sub

Private Sub cmbcotizacion_Click()
    FrmCotizaciones.cmbmoneda = txtmoneda
    FrmCotizaciones.cmbmoneda.Enabled = False
    FrmCotizaciones.Show vbModal
    txtcotiz = FrmCotizaciones.txtcotizacion
End Sub

Private Sub cmbcuenta_Click()
    FrmHelp.Show
    CargarHelpCuentas "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
    FrmHelp.Tag = Me.Name
    cargar = "Cuentas"
End Sub

Private Sub cmbeliminofila_Click()
    If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
        If grilla.rows > 1 Then
            txttotal = s2nt(txttotal) - s2nt(grilla.TextMatrix(grilla.Row, 3))
            If grilla.rows = 2 Then
                grilla.TextMatrix(1, 0) = ""
                grilla.TextMatrix(1, 1) = ""
                grilla.TextMatrix(1, 2) = ""
                grilla.TextMatrix(1, 3) = ""
            Else
                grilla.RemoveItem (grilla.Row)
            End If
        Else
            MsgBox "No hay items para eliminar o no ha seleccionado ninguno de ellos"
        End If
    End If
End Sub

Private Sub cmdaceptar_Click()
Dim rs As New ADODB.Recordset
Dim fecha As Variant
Dim i As Long

If txtcodcaja = "" Then
    MsgBox "Debe ingresar un código de Caja"
    Exit Sub
End If

If txttotal <> txtimporte Then
    MsgBox "No coincide el importe ingresado con el importe total"
    Exit Sub
End If

If Ope <> "" Then
        
'        If txtmoneda <> "Pesos" Then
'            rs.Open "select * from Cotizaciones where moneda = " & ObtenerCodigo("Monedas", txtmoneda) & " and Fecha = cdate('" & Date & "') and activo = 1", daTaenvironment1.amr, adOpenStatic, adLockOptimistic
'            If rs.EOF Then
'               MsgBox "La moneda asociada a la caja ingresada no se encuentra actualizada"
'               End Sub
'            End If
'        End If
        
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
       
        If Ope = "A" Then
            If optingreso = True Then
                DataEnvironment1.dbo_MOVICAJAS "A", val(txtcodcaja), val(txtmovimiento), _
                val(txtcodcli), "E", "I", s2nt(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacaja), 0, s2nt(txtcotiz), fecha, UsuarioSistema!codigo, 0, 0, 1
            Else
                DataEnvironment1.dbo_MOVICAJAS "A", val(txtcodcaja), val(txtmovimiento), _
                val(txtcodcli), "E", "E", s2nt(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacaja), 0, s2nt(txtcotiz), fecha, UsuarioSistema!codigo, 0, 0, 1
            End If
            
            If optingreso = True Or optegreso = True Then
                For i = 1 To grilla.rows - 1
                    DataEnvironment1.dbo_DETMOVCAJAS "A", val(txtmovimiento), _
                    s2nt(grilla.TextMatrix(i, 3)), IIf(txtcodcli <> "", val(txtcodcli), 0), Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "IE"
                Next
            End If
        Else
            If Ope = "M" Then
                If optingreso = True Then
                    DataEnvironment1.dbo_MOVICAJAS "M", val(txtcodcaja), val(txtmovimiento), _
                    val(txtcodcli), "E", "I", s2nt(txtimporte), Trim(txtConcepto), dtfecha, _
                    Trim(txtcuentacaja), 0, s2nt(txtcotiz), 0, 0, 0, 0, 0
                Else
                    DataEnvironment1.dbo_MOVICAJAS "M", val(txtcodcaja), val(txtmovimiento), _
                    val(txtcodcli), "E", "E", s2nt(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacaja), 0, s2nt(txtcotiz), 0, 0, 0, 0, 0
                End If
                
                If optingreso = True Or optegreso = True Then
                    DataEnvironment1.AMR.Execute "delete from DetalleMovCajas where movimiento = " & val(txtmovimiento) & ""
                    
                    For i = 1 To grilla.rows - 1
                        DataEnvironment1.dbo_DETMOVCAJAS "A", val(txtmovimiento), _
                        s2nt(grilla.TextMatrix(i, 3)), IIf(txtcodcli <> "", val(txtcodcli), 0), Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "IE"
                    Next
                End If
                
                DataEnvironment1.dbo_GRABARBITACORA val(Trim(txtmovimiento)), "Usuarios", UsuarioSistema!codigo, fecha, Time, "M"
            End If
        End If
        MsgBox "La operación fue realizada con éxito"
        LimpioControles
        Call Habilitobotones(True, True, True, True, True, True)
        Call HabilitoControles(False)
        Call MonedaVisible(False)
        grilla.Clear
        InicioGrilla
        cargar = ""
        habilitogrillaenable (False)
Else
    MsgBox "Operación no válida"
End If

End Sub

Private Sub cmdBuscar_Click()
    cargar = "Movicaja"
    FrmHelp.Show
    CargarHelp "MOVICAJA", "Movimiento", "Caja", "movimiento", "caja", "movimiento"
    FrmHelp.Tag = Me.Name
    Call Habilitobotones(True, False, True, True, True, True)
End Sub

Private Sub cmdcancelar_Click()
    grilla.Clear
    InicioGrilla
    LimpioControles
    LimpioImputacion
    Call HabilitoControles(False)
    Call Habilitobotones(True, True, False, False, False, True)
    Call MonedaVisible(False)
    cargar = ""
End Sub
Public Sub CargarDatos()
Dim rs As New ADODB.Recordset
Dim mon As Long
Dim fecha As String, codigo

    If rsefec.State = 1 Then
        rsefec.Close
        Set rsefec = Nothing
    End If
    
    codigo = val(Trim(Me.Tag))
    
    If cargar = "Cajas" Then
        
        rs.Open "select * from Cajas where codigo = " & val(txtcodcaja) & " and activo = 1", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcaja = rs!codigo
            txtcaja = rs!responsable
            txtcuentacaja = rs!Cuenta
            If Not IsNull(rs!moneda) Then
                txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
            End If
        End If
        rs.Close
        Set rs = Nothing
        
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        
        If ObtenerCodigo("Monedas", txtmoneda) <> 1 Then
            rs.Open "select * from Cotizaciones where Fecha =" & ssFecha(dtfecha) & " and moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtmoneda = ObtenerDescripcion("Monedas", ObtenerCodigo("Monedas", txtmoneda))
                txtcotiz = rs!cotizacion
            Else
                MsgBox "Debe ingresar la cotización del día"
            End If
            MonedaVisible (True)
            rs.Close
            Set rs = Nothing
        End If
    End If

    If cargar = "Cuentas" Then
        If txtcodcuenta = "" Then
            txtcodcuenta = Trim(STR(codigo))
        End If
        If Not noestaenlagrilla(txtcodcuenta, grilla) And esimputable(txtcodcuenta) Then
            rs.Open "select * from Cuentas where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtcodcuenta = rs!codigo
                txtcuenta = rs!descripcion
                txtconc.SetFocus
            End If
            rs.Close
            Set rs = Nothing
        Else
            MsgBox "El concepto ya se encuentra cargado"
            txtcodcuenta = ""
            txtcodcuenta.SetFocus
        End If
    End If


    If cargar = "Deposito" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Proveedor" Then
        rs.Open "select * from Prov where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = rs!descripcion
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Clientes" Then
        rs.Open "select * from Clientes where codigo = " & val(txtcodcli) & " and activo = 1", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcli = rs!codigo
            txtCliente = rs!descripcion
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Movicaja" Then
        rsefec.Open "select * from MOVICAJA where activo = 1 and movimiento = " & codigo & "", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
        If Not rsefec.EOF Then
            cargodatos
        End If
        rsefec.Close
        Set rsefec = Nothing
    End If

End Sub

Private Sub cmdcargar_Click()
    Dim totalgrilla ' compil *******************
    
    If txtvalor <> "" Then
        If modifico = False Then
            If (s2nt(txtvalor) <= s2nt(txtimporte)) And (s2nt(txtvalor) + s2nt(txttotal) <= s2nt(txtimporte)) Then
                If txttotal <> "" Then
                    If s2nt(txttotal) + s2nt(txtvalor) <= s2nt(txtimporte) Then
                        Cargogrilla
                    Else
                        MsgBox "Con este valor el importe total serìa superado", vbInformation
                    End If
                Else
                    Cargogrilla
                End If
                Limpiotextosgrilla
                If txtcodcuenta.Enabled = True Then
                    txtcodcuenta.SetFocus
                End If
            Else
                If optegreso = True Then
                    MsgBox "El valor a egresar debe ser el mismo que el original"
                Else
                    MsgBox "El valor a ingresar no puede superar al importe original"
                End If
                txtvalor.SetFocus
            End If
        Else
            totalgrilla = sumogrilla()
            If totalgrilla - s2nt(grilla.TextMatrix(grilla.Row, 3)) + s2nt(txtvalor) <= s2nt(txtimporte) Then
                grilla.TextMatrix(grilla.Row, 0) = txtcodcuenta
                grilla.TextMatrix(grilla.Row, 1) = txtcuenta
                grilla.TextMatrix(grilla.Row, 2) = txtconc
                grilla.TextMatrix(grilla.Row, 3) = txtvalor
                txttotal = sumogrilla()
                LimpioImputacion
                modifico = False
                grilla.SetFocus
            Else
                MsgBox "El valor a ingresar no puede superar al total"
                txtvalor.SetFocus
            End If
        End If
    Else
        MsgBox "Debe ingresar un valor"
        txtvalor.SetFocus
    End If
End Sub

Function sumogrilla() As Double
Dim x As Long
Dim Total As Double
    
    For x = 1 To grilla.rows - 1
        Total = Total + s2nt(grilla.TextMatrix(x, 3))
    Next
    sumogrilla = Total
    
End Function

Private Sub LimpioImputacion()
    txtcodcuenta = ""
    txtcuenta = ""
    txtconc = ""
    txtvalor = ""
End Sub

Private Sub MonedaVisible(habilito As Boolean)
    lblmoneda.Visible = habilito
    lblcotiz.Visible = habilito
    txtmoneda.Visible = habilito
    txtcotiz.Visible = habilito
    cmbcotizacion.Visible = habilito
End Sub
Private Sub cmdeliminar_Click()

Dim fecha As Variant
Dim mensaje As String

    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_MOVICAJAS "B", 0, Trim(txtmovimiento), 0, "", "", 0, "", 0, "", 0, 0, 0, 0, UsuarioSistema!codigo, fecha, 0
        DataEnvironment1.dbo_GRABARBITACORA val(Trim(txtmovimiento)), "", UsuarioSistema!codigo, fecha, Time, "B"
        
        MsgBox "El registro se ha eliminado"
        Call Habilitobotones(True, True, False, False, False, False)
        Call HabilitoControles(False)
        LimpioControles
        InicioGrilla
    End If

End Sub

Private Sub cmdmodificar_Click()
    Ope = "M"
    Call HabilitoControles(True)
    Call Habilitobotones(True, False, False, True, True, True)
    habilitogrillaenable (True)
    Call MonedaVisible(True)
End Sub

Private Sub cmdnuevo_Click()
Dim rs As New ADODB.Recordset

    Call HabilitoControles(True)
    Call Habilitobotones(False, False, False, False, True, True)
    LimpioControles
    
    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
    If Not IsNull(rs!maxcodigo) Then
        txtmovimiento = rs!maxcodigo + 1
        numero = rs!maxcodigo + 1
    End If
    rs.Close
    Set rs = Nothing
    
    Ope = "A"
    modifico = False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 1500
    InicioGrilla
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub grilla_Click()
    modifico = True
    MuestroGrilla
End Sub

Private Sub MuestroGrilla()
    txtcodcuenta = grilla.TextMatrix(grilla.Row, 0)
    txtcuenta = grilla.TextMatrix(grilla.Row, 1)
    txtconc = grilla.TextMatrix(grilla.Row, 2)
    txtvalor = grilla.TextMatrix(grilla.Row, 3)
End Sub
Private Sub optdeposito_Click()
    lblcambio.caption = "Cuenta"
    cmbcambio.caption = "Cuenta"
    habilitogrilla (False)
End Sub

Private Sub optdeposito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optegreso_Click()
    lblcambio.caption = "Proveedor"
    cmbcambio.caption = "Proveedor"
    habilitogrilla (True)
End Sub

Private Sub optegreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optingreso_Click()
    lblcambio.caption = "Cliente"
    cmbcambio.caption = "Cliente"
    habilitogrilla (True)
End Sub


Private Sub optingreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtcodcaja_GotFocus()
    txtcodcaja.SelStart = 0
    txtcodcaja.SelLength = Len(txtcodcaja.Text)
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodcaja_LostFocus()
    If IsNumeric(txtcodcaja) Then
        txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
        If txtcaja = "" Then
            MsgBox "Còdigo de caja incorrecto"
            txtcodcaja = "0"
            txtcodcaja.SetFocus
        Else
            cargar = "Cajas"
            CargarDatos
        End If
    Else
        If txtcodcaja <> "" Then
            MsgBox "Còdigo de caja incorrecto"
            txtcodcaja = "0"
            txtcodcaja.SetFocus
        End If
    End If
End Sub


Private Sub txtcodcli_GotFocus()
    txtcodcli.SelStart = 0
    txtcodcli.SelLength = Len(txtcodcli.Text)
End Sub

Private Sub txtcodcli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodcli_LostFocus()
    Select Case lblcambio.caption
        Case "Cliente":
                If IsNumeric(txtcodcli) Then
                    txtCliente = ObtenerDescripcion("Clientes", val(txtcodcli))
                    If txtCliente = "" Then
                        MsgBox "Còdigo de cliente incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Clientes"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Còdigo de cliente incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
                
        Case "Proveedor":
                If IsNumeric(txtcodcli) Then
                    txtCliente = ObtenerDescripcion("Prov", val(txtcodcli))
                    If txtCliente = "" Then
                        MsgBox "Còdigo de proveedor incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Proveedor"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Còdigo de proveedor incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
                
        Case "Deposito":
                If IsNumeric(txtcodcli) Then
                    txtCliente = ObtenerDescripcion("Cajas", val(txtcodcli))
                    If txtCliente = "" Then
                        MsgBox "Còdigo de depósito incorrecto"
                        txtcodcli = "0"
                        txtcodcli.SetFocus
                    Else
                        cargar = "Deposito"
                        CargarDatos
                    End If
                Else
                    If txtcodcli <> "" Then
                        MsgBox "Còdigo de depósito incorrecto"
                        'txtcodcli = "0"
                        txtcodcli.SetFocus
                    End If
                End If
        End Select
End Sub

Private Sub txtcodcuenta_GotFocus()
Dim rs As New ADODB.Recordset

    rs.Open "select dato_fijo from datos", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        If rs!DATO_FIJO = 7 Then
            txtcodcuenta = "1"
            txtcodcuenta.Enabled = False
            txtcuenta = "COMPRAS"
            txtconc = "COMPRAS"
            txtconc.Enabled = False
            txtvalor = txtimporte
            txtvalor.Enabled = False
            cmbcuenta.Enabled = False
            cmdcargar.Enabled = False
            cmbeliminofila.Enabled = False
            Cargogrilla
        End If
    End If
    rs.Close

End Sub

Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodcuenta_LostFocus()
    If IsNumeric(txtcodcuenta) Then
        If Not noestaenlagrilla(txtcodcuenta, grilla) And esimputable(val(txtcodcuenta)) Then
            txtcuenta = ObtenerDescripcion("Cuentas", val(txtcodcuenta))
            If txtcuenta = "" Then
                MsgBox "Còdigo de cuenta incorrecto"
                txtcodcuenta = ""
                txtcodcuenta.SetFocus
            Else
                cargar = "Cuentas"
                CargarDatos
            End If
        Else
            MsgBox "El concepto ya se encuentra cargado o la cuenta no es imputable"
            txtcodcuenta = ""
            txtcodcuenta.SetFocus
        End If
    Else
        If txtcodcuenta <> "" Then
            MsgBox "Còdigo de cuenta incorrecto"
            txtcodcuenta = ""
            txtcodcuenta.SetFocus
        End If
    End If
End Sub

Private Sub txtconc_Change()
Dim i As Long
    txtconc.Text = UCase(txtconc.Text)
    i = Len(txtconc.Text)
    txtconc.SelStart = i
End Sub

Private Sub txtconc_GotFocus()
    If txtcodcuenta = "" Then
        MsgBox "Debe cargar la cuenta"
        txtcodcuenta.SetFocus
    End If
End Sub

Private Sub txtconc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtconcepto_Change()
Dim i As Long
    txtConcepto.Text = UCase(txtConcepto.Text)
    i = Len(txtConcepto.Text)
    txtConcepto.SelStart = i
End Sub

Private Sub txtConcepto_GotFocus()
    txtConcepto.SelStart = 0
    txtConcepto.SelLength = Len(txtConcepto.Text)
End Sub

Private Sub txtconcepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub



Private Sub txtcotiz_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcotiz_LostFocus()
    If Not IsNumeric(txtcotiz) Then
        MsgBox "Cotización incorrecta"
        txtcotiz = "0"
        txtcotiz.SetFocus
    End If
End Sub

Private Sub txtimporte_GotFocus()
    txtimporte.SelStart = 0
    txtimporte.SelLength = Len(txtimporte.Text)
End Sub


Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtmoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtmovimiento_GotFocus()
    txtmovimiento.SelStart = 0
    txtmovimiento.SelLength = Len(txtmovimiento.Text)
    If optdeposito = False And optegreso = False And optingreso = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If
End Sub

Private Sub txtmovimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtmovimiento_LostFocus()
    If IsNumeric(txtmovimiento) Then
        If val(txtmovimiento) < numero Then
            MsgBox "El código no puede ser menor al último ingresado"
            txtmovimiento.SetFocus
        End If
    Else
        If txtmovimiento <> "" Then
            MsgBox "Debe ingresar un código"
            txtmovimiento = "0"
            txtmovimiento.SetFocus
        End If
    End If
End Sub

Sub LimpioControles()
    txtmovimiento = ""
    txtConcepto = ""
    dtfecha = Date
    txtcodcli = ""
    txtCliente = ""
    txttotal = "0"
    txtcotiz = "0"
    txtimporte = ""
    txtcodcaja = ""
    txtcaja = ""
    txtcuentacaja = ""
    optegreso.Value = False
    optingreso.Value = False
    optdeposito.Value = False
    Ope = ""
End Sub

Sub cargodatos()
Dim rs As New ADODB.Recordset

    If rsefec!ing_egr = "I" Then
        optingreso.Value = True
    Else
        If rsefec!ing_egr = "E" Then
            optegreso.Value = True
        Else
            optdeposito.Value = True
        End If
    End If
    
    txtmovimiento = rsefec!movimiento
    txtcodcaja = rsefec!Caja
    txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
    txtmoneda = ObtenerDescripcion("Monedas", ObtenerMoneda("Cajas", val(txtcodcaja)))
    If Not IsNull(rsefec!concepto) Then
        txtConcepto = rsefec!concepto
    End If
    
    dtfecha = rsefec!fecha
    txtimporte = rsefec!Importe
    
    Call MonedaVisible(True)
    
    rs.Open "select * from Cotizaciones where Fecha = " & ssFecha(dtfecha) & " and  moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        MonedaVisible (True)
        txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
        txtcotiz = rs!cotizacion
    Else
        txtcotiz = "0"
    End If
    rs.Close
    Set rs = Nothing
    
    
    If Not IsNull(rsefec!cli_prov) Then
        txtcodcli = rsefec!cli_prov
        txtCliente = ObtenerDescripcion("Clientes", val(txtcodcli))
    End If
        
    InicioGrilla
    txttotal = "0"
    
    rs.Open "select * from DetalleMovcajas where movimiento = " & val(txtmovimiento) & "", DataEnvironment1.AMR, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        habilitogrilla (True)
        grilla.rows = 2
        grilla.Row = 0
        While Not rs.EOF
            grilla.Row = grilla.Row + 1
            grilla.TextMatrix(grilla.Row, 0) = rs!Cuenta
            grilla.TextMatrix(grilla.Row, 1) = ObtenerDescripcion("Cuentas", val(rs!Cuenta))
            grilla.TextMatrix(grilla.Row, 2) = rs!concepto
            grilla.TextMatrix(grilla.Row, 3) = rs!Importe
            If txttotal <> "" Then
                txttotal = s2nt(txttotal) + s2nt(rs!Importe)
            Else
                txttotal = s2nt(rs!Importe)
            End If
            rs.MoveNext
            If Not rs.EOF Then
                grilla.rows = grilla.rows + 1
            End If
        Wend
    End If
    rs.Close
    Set rs = Nothing
    
    
End Sub

Sub HabilitoControles(habilito As Boolean)
    txtmovimiento.Enabled = habilito
    cmbcaja.Enabled = habilito
    cmbcambio.Enabled = habilito
    txtConcepto.Enabled = habilito
    dtfecha.Enabled = habilito
    txtcodcli.Enabled = habilito
    optegreso.Enabled = habilito
    optingreso.Enabled = habilito
    txtimporte.Enabled = habilito
    optdeposito.Enabled = habilito
    txtcodcaja.Enabled = habilito
    txtcodcli.Enabled = habilito
End Sub

Sub Habilitobotones(busco As Boolean, Nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdBuscar.Enabled = busco
    cmdnuevo.Enabled = Nuevo
    cmdModificar.Enabled = modifico
    cmdeliminar.Enabled = elimino
    cmdaceptar.Enabled = acepto
    cmdcancelar.Enabled = Cancelo
End Sub

'**************************************** compil
'Private Sub cmdPrimero_Click()
'    rsefec.MoveFirst
'    txtCodigo = rsefec!codigo
'    txtDescripcion = rsefec!Descripcion
'    Call HabilitoBotonesMoverse(False, False, True, True)
'End Sub
'
'Private Sub cmdsiguiente_Click()
'    rsefec.MoveNext
'    If Not rsefec.EOF Then
'        txtCodigo = rsefec!codigo
'        txtDescripcion = rsefec!Descripcion
'        Call HabilitoBotonesMoverse(True, True, True, True)
'    Else
'        Call HabilitoBotonesMoverse(True, True, False, False)
'    End If
'End Sub
'
'Private Sub cmdUltimo_Click()
'    rsefec.MoveLast
'    txtCodigo = rsefec!codigo
'    txtDescripcion = rsefec!Descripcion
'    Call HabilitoBotonesMoverse(True, True, False, False)
'End Sub
'
'Private Sub cmdanterior_Click()
'    rsefec.MovePrevious
'    If Not rsefec.BOF Then
'        txtCodigo = rsefec!codigo
'        txtDescripcion = rsefec!Descripcion
'        Call HabilitoBotonesMoverse(True, True, True, True)
'    Else
'        Call HabilitoBotonesMoverse(False, False, True, True)
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsefec.State = 1 Then
        rsefec.Close
        Set rsefec = Nothing
    End If
End Sub

'Private Sub txtconcepto_LostFocus()
'    If txtconcepto = "" Then
'        MsgBox "Debe ingresar un concepto"
'        txtconcepto.SetFocus
'    End If
'End Sub

Sub InicioGrilla()
    grilla.Clear
    'grilla.ColWidth(1) = 1700
    grilla.TextMatrix(0, 0) = "Cuenta"
    grilla.TextMatrix(0, 1) = "Descripción"
    grilla.TextMatrix(0, 2) = "Concepto"
    grilla.TextMatrix(0, 3) = "Importe"
    grilla.rows = 2
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label2.Visible = habilito
    txtcodcuenta.Visible = habilito
    cmbcuenta.Visible = habilito
    txtcuenta.Visible = habilito
    Label6.Visible = habilito
    txtconc.Visible = habilito
    Label3.Visible = habilito
    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
    grilla.Visible = habilito
    cmbeliminofila.Visible = habilito
    Label8.Visible = habilito
    txttotal.Visible = habilito
End Sub

Private Sub txtimporte_LostFocus()
    If Not IsNumeric(txtimporte) Then
        MsgBox "Debe ingresar un importe"
        txtimporte = "0"
        txtimporte.SetFocus
    Else
        Call habilitogrillaenable(True)
        txtimporte = s2nt(txtimporte)
    End If
End Sub

Private Sub Limpiotextosgrilla()
    txtcodcuenta = ""
    txtcuenta = ""
    txtconc = ""
    txtvalor = ""
End Sub


Private Sub Cargogrilla()
    If grilla.rows = 2 Then
        grilla.Row = 1
        grilla.Col = 0
        If Trim(grilla.Text) = "" Then
            grilla.Row = 1
            grilla.Col = 0
            grilla.Text = txtcodcuenta
            grilla.Col = 1
            grilla.Text = txtcuenta
            grilla.Col = 2
            grilla.Text = txtconc
            grilla.Col = 3
            grilla.Text = txtvalor
        Else
            grilla.AddItem txtcodcuenta & Chr(9) & txtcuenta & Chr(9) & txtconc & Chr(9) & txtvalor
        End If
    Else
        grilla.AddItem txtcodcuenta & Chr(9) & txtcuenta & Chr(9) & txtconc & Chr(9) & txtvalor
    End If
    If txttotal <> "" Then
        txttotal = s2nt(txttotal) + s2nt(txtvalor)
    Else
        txttotal = s2nt(txtvalor)
    End If
    If txttotal = txtimporte Then
        MsgBox "El detalle ha sido completado"
'        habilitogrillaenable (False)
    End If
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)
    Label2.Enabled = habilito
    txtcodcuenta.Enabled = habilito
    cmbcuenta.Enabled = habilito
    Label6.Enabled = habilito
    txtconc.Enabled = habilito
    Label3.Enabled = habilito
    txtvalor.Enabled = habilito
    cmdcargar.Enabled = habilito
    grilla.Enabled = habilito
    cmbeliminofila.Enabled = habilito
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtvalor_LostFocus()
    If IsNumeric(txtvalor) Then
'        InicioGrilla
        If grilla.Visible = False Then
            habilitogrilla (True)
        End If
        habilitogrillaenable (True)
        txtvalor = s2nt(txtvalor)
    Else
        If txtvalor <> "" Then
            MsgBox "Debe ingresar un importe"
            txtvalor = "0"
            txtvalor.SetFocus
        End If
    End If
End Sub

'5/5/5
'   numero long
'

