VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPagosACuenta 
   Caption         =   "Pagos a Cuenta"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   Icon            =   "FrmPagosACuenta3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFaltaPagar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1365
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
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
      Height          =   375
      Left            =   3930
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6795
      Width           =   975
   End
   Begin Gestion.ucCuit cuit 
      Height          =   255
      Left            =   8160
      TabIndex        =   47
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
   End
   Begin Gestion.ucCheques uCheques 
      Height          =   2865
      Left            =   1215
      TabIndex        =   18
      Top             =   3375
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   5054
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   390
      TabIndex        =   8
      Top             =   7590
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcotiz 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cmbmoneda 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmPagosACuenta3.frx":08CA
      Left            =   6120
      List            =   "FrmPagosACuenta3.frx":08CC
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtserie 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   180
      Width           =   495
   End
   Begin VB.TextBox txtimporte 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
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
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6795
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
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
      Height          =   375
      Left            =   6675
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6795
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6795
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
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
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6795
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Salir"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6795
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
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
      Height          =   375
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6795
      Width           =   975
   End
   Begin VB.TextBox txtopago 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8160
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox txtnombre 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.TextBox txtcodprov 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdprov 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedor"
      Height          =   375
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "2"
      Top             =   495
      Width           =   975
   End
   Begin VB.TextBox txtprov 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4080
      TabIndex        =   6
      Top             =   540
      Width           =   5295
   End
   Begin VB.TextBox txtimpcheques 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3675
      Width           =   1035
   End
   Begin VB.TextBox txtefectivo 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   13
      Top             =   2820
      Width           =   1020
   End
   Begin VB.TextBox txttransf 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   75
      TabIndex        =   19
      Top             =   6300
      Width           =   1095
   End
   Begin VB.TextBox txtcuenta 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4050
      TabIndex        =   22
      Tag             =   "2"
      Top             =   6315
      Width           =   3795
   End
   Begin VB.CommandButton cmbcuenta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6315
      Width           =   855
   End
   Begin VB.TextBox txtcodcuenta 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2190
      TabIndex        =   20
      Top             =   6315
      Width           =   915
   End
   Begin VB.TextBox txtcodcaja 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2355
      TabIndex        =   14
      Top             =   2820
      Width           =   915
   End
   Begin VB.CommandButton cmbcaja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2805
      Width           =   855
   End
   Begin VB.TextBox txtcaja 
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4275
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   2805
      Width           =   3795
   End
   Begin VB.TextBox txttipoiva 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7830
      TabIndex        =   28
      Tag             =   "1"
      Top             =   7590
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtfechacompra 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   30
      TabIndex        =   27
      Top             =   7470
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txttipocompra 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1710
      TabIndex        =   26
      Top             =   7590
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtnumcompra 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3510
      TabIndex        =   25
      Top             =   7590
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6510
      TabIndex        =   24
      Tag             =   "1"
      Top             =   7590
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   186843137
      CurrentDate     =   37934
   End
   Begin Gestion.ucRetCompras uRetCompras 
      Height          =   720
      Left            =   60
      TabIndex        =   12
      Top             =   1815
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1270
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   9240
      Y1              =   6735
      Y2              =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   0
      X2              =   12120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblcotiz 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cotiz."
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
      Left            =   7380
      TabIndex        =   46
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblmoneda 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   45
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Serie"
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
      Left            =   5520
      TabIndex        =   44
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Total"
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
      Left            =   240
      TabIndex        =   42
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Documento"
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
      Left            =   6660
      TabIndex        =   40
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lblnombre 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Left            =   240
      TabIndex        =   39
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblcuit 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuit"
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
      Left            =   7560
      TabIndex        =   38
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cheques"
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
      Left            =   135
      TabIndex        =   37
      Top             =   3405
      Width           =   870
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Efectivo"
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
      TabIndex        =   36
      Top             =   2580
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00400000&
      X1              =   240
      X2              =   10560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transf."
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
      Left            =   60
      TabIndex        =   35
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Left            =   1470
      TabIndex        =   34
      Top             =   6315
      Width           =   735
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Caja"
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
      Left            =   1635
      TabIndex        =   33
      Top             =   2805
      Width           =   855
   End
End
Attribute VB_Name = "FrmPagosACuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit '23/5/4
 ' 30/3/5
 '********* cont

Dim midDoc  As Long, ssDonde As String
Dim Ope As String
Dim rsmov As New ADODB.Recordset

Private Const RECIBOS_A_CUENTA_MOVICAJA = "O/P"
Private Const RECIBOS_A_CUENTA_TRANSCOM = "RAC"
'

Private Sub cmbcaja_Click()
    FrmHelp.Show
    CargarHelp "Cajas", "Codigo", "Descripción", "codigo", "responsable"
    FrmHelp.Tag = Me.Name
    cargar = "Cajas"
End Sub

Private Sub cmbcuenta_Click()
    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo,banco,numero,moneda from ctasbank where activo = 1 order by codigo")
    If resu > "" Then
        txtcodcuenta = frmBuscar.resultado
        txtcuenta = ObtenerDescripcion("BancosGrales", frmBuscar.resultado(2))
    End If
    cargar = "CuentasBank"
End Sub

Private Sub cmdBuscar_Click()
        
    Dim z As Double
    
    Call Habilitobotones(True, True, True, True, False, True)
    cargar = "BU"

    Dim Consulta As String

    If MsgBox("¿Buscar usados?", vbYesNo) = vbYes Then
        ssDonde = " COMPRAS "
    Else
        ssDonde = " TRANSCOM "
    End If

    Consulta = "Select Fecha, Tipodoc as 'Tipo Doc.', nrodoc as 'Nro Doc.', Total, RetGan as [_H_rg],iddoc,cotizacion " & _
                "From  " & ssDonde & _
                "Where TIPODOC = '" & RECIBOS_A_CUENTA_TRANSCOM & "' and activo = 1 "
    
    If TxtCodProv <> "" Then Consulta = Consulta & " and CODPR = " & s2n(TxtCodProv.Text)

    Consulta = Consulta & " Order By FECHA desc, NRODOC desc"
    
    frmBuscar.MostrarSql Consulta
    If frmBuscar.resultado <> "" Then
        With frmBuscar
            z = s2n(nSinNull(.resultado(7)), 4)
            If z = 0 Then z = 1
            dtFecha.Value = .resultado
            txttipocompra.Text = .resultado(2)
            txtopago.Text = .resultado(3)
            txtimporte.Text = .resultado(4) / z
            uRetCompras.retgan = nSinNull(.resultado(5)) / z '    txtRetGan = .resultado(5)
            midDoc = nSinNull(.resultado(6))
            CargarDatos
        End With
    End If

End Sub
Private Sub cmdCancelar_Click()
    LimpioControles
    HabilitoControles (False)
    Call Habilitobotones(True, True, False, False, False, True)
    uCheques.Borrar
    FrmCostosYContable.LimpioControles
'    FrmCostosYContable.InicioGrilla
End Sub

Private Sub cmdeliminar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    
    Dim mensaje As Long, cont As Long
    Dim rs As New ADODB.Recordset
    
    Dim conttot, contdif
    Dim sp
    sp = obtenerDeSQL("select * from relfnr_c d inner join IMPPRO o on d.ndoc=o.nro where o.activo=1 and prov=" & s2n(TxtCodProv) & " and d.tfac='RAC' and d.fact=" & s2n(txtopago))
    If IsNull(sp) Or IsEmpty(sp) Then
    Else
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If

    mensaje = MsgBox("Esta seguro que desea elimnar este registro?", vbYesNo, "Atencion")
    If mensaje = vbYes Then
        
        If txtsaldo = txtimporte Then
            conttot = 0
            contdif = 0
            
            rs.Open "select * from Chq_comp where tipodoc = 'RAC' and nrodoc = " & ssNum(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                conttot = conttot + 1
                If rs!estado <> "T" Then
                    cont = cont + 1
                End If
                rs.MoveNext
            Wend
            rs.Close
            Set rs = Nothing
            
            '****************************************************************
            DE_BeginTrans
'
'           If conttot = contdif Then ' NO LE CREO A ESTA PREGUNTA, ademas contdif SIEMPRE esta en 0
            If cont = 0 Then           ' si  = 0 TODOS LOS CHEQUES SON ESTADO 'T'      'lito   --------  creo q es asi -------
'
                conttot = 0
                contdif = 0
                rs.Open "select * from cheques where tdocprov = 'RAC' and ndocprov = " & n2s(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                While Not rs.EOF
                    conttot = conttot + 1
                    If rs!estado <> "T" Then
                        cont = cont + 1
                    End If
                    rs.MoveNext
                Wend
                rs.Close
                Set rs = Nothing
'
'               If conttot = contdif Then
                If cont = 0 Then           ' si  = 0 TODOS LOS CHEQUES SON ESTADO 'T'      'lito   --------  creo q es asi -------
'
                    rs.Open "select * from Movibanc where fecha = " & ssFecha(dtFecha) & " and tipdoc = 'O/P' and nrodoc = " & n2s(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    While Not rs.EOF
                        DataEnvironment1.dbo_INGCOMPRAMOVIBANC "B", 0, "", "", _
                          0, "", 0, 0, "", 0, rs!MovBanco, midDoc, Date, UsuarioSistema!codigo, 1
                        rs.MoveNext
                    Wend
                    rs.Close
                    Set rs = Nothing
                
                    rs.Open "select * from Movicaja where fecha = " & ssFecha(dtFecha) & " and tipo = 'O/P' and nrodoc = " & n2s(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    While Not rs.EOF
                        DataEnvironment1.dbo_INGCOMPRAMOVICAJA "B", 0, rs!movimiento, "", "", 0, "", _
                          0, 0, 0, "", 0, "", 0, midDoc, Date, UsuarioSistema!codigo, 1
                        rs.MoveNext
                    Wend
                    rs.Close
                    Set rs = Nothing
                    
                    DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "B", 0, 0, 0, s2n(txtopago), "RAC", 0, "", 0, 0, 0, 0, Date, UsuarioSistema!codigo, 1, 1, 0
                    DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "B", 0, 0, "", s2n(TxtCodProv), s2n(txtopago), 0 _
                      , 0, "", "RAC", 0, 0, Date, UsuarioSistema!codigo, 1, 1, 1
                Else
                    MsgBox "No se puede dar de baja dado que uno o mas cheques de terceros ya fueron acreditados"
                    Exit Sub
                End If
            Else
                MsgBox "No se puede dar de baja dado que uno o mas cheques propios ya fueron debitados"
                Exit Sub
            End If
                                        
            DataEnvironment1.dbo_INGCOMPRASCTACTE "B", 0, 0, 0, s2n(TxtCodProv), "", "", 0, "", s2n(txtopago), _
              0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, s2n(txtserie), 0, 0, 0, 0, 0, 0, 0, 0, Date, UsuarioSistema!codigo, 0, 0, 0, _
              "", "", 0, 0
            
'            DataEnvironment1.dbo_INGCOMPRASDETALLE "B", 0, 0, 0, "", "", "", s2n(txtcodprov), "O/P", s2n(txtopago), 0
            
            DataEnvironment1.dbo_GRABARBITACORA s2n(TxtCodProv), "Transcom", UsuarioSistema!codigo, Date, Time, "B"
            
            'borr doc y asiento
            
            
            
            If midDoc > 0 Then  ' TEMPORARIO   ASUMO QUE SI ES 0 ES DATO MIGRADO
            
                If Not BorroDocumento(midDoc) Then 'AsientoBaja_idDoc(mIdDoc) Then
                    ufa "err no se pudo borrar doc PCTA", "middoc " & midDoc
                    DE_RollbackTrans
                    Exit Sub
                End If
                
            End If
            
            DE_CommitTrans
            '**********************************************
            
            LimpioControles
            HabilitoControles (True)
            Call Habilitobotones(True, True, False, False, False, True)
    '        cmbingresar.Enabled = False
        Else
            MsgBox "No se puede anular este coprobante dado que fue parcialmente pagado"
        End If
    End If
    
fin:
    Exit Sub
UFAelim:
    DE_RollbackTrans
    ufa "err al anular", ""
    Resume fin
End Sub

Private Sub cmdImprimir_Click()
ImprimirPagoCuenta
End Sub

Private Sub cmdnuevo_Click()
    LimpioControles
    HabilitoControles (True)
    
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    
    Call Habilitobotones(False, False, False, False, True, True)
    txtopago = nuevoCodigoOP()
    uCheques.Borrar
    TxtCodProv.SetFocus
    
    Ope = "A"
End Sub

Private Sub cmdOk_Click() 'ALTA SOLAMENTE
    If ON_ERROR_HABILITADO Then On Error GoTo UfaOK
    
    Dim NroCertifGan As Long, NroCertifIIBB As Long
    Dim numIntP As Long
    Dim cueche As Long ' cuenta banc che pro
    Dim z As Double
    Dim x As Long
    
    z = s2n(txtcotiz, 4)
    If z = 0 Then z = 1
    
    If txtopago = "" Then
        MsgBox "Debe ingresar el número del pago"
        Exit Sub
    End If
    
    'ALTA
    If existeOP(s2n(txtopago, 0)) Then
        che "numero ya existe "
        Exit Sub
    End If
    
    If TxtCodProv = "" Then
        MsgBox "Debe ingresar el código de proveedor"
        Exit Sub
    End If
    
    If txtimporte = "" Then
        MsgBox "Debe ingresar el importe del pago"
        Exit Sub
    End If
    If s2n(txttransf) <> 0 Then
        If Trim(txtcodcuenta) = "" Or Trim(txtcuenta) = "" Then
            MsgBox "Debe Cargar la Cuenta para la transferencia"
            Exit Sub
        End If
    End If
    
    'If s2n(s2n(txtefectivo) + s2n(txtimpcheques) + s2n(txttransf)) <> s2n(txtimporte) Then
    If s2n(SumoPagos() - s2n(txtimporte)) <> 0 Then
        MsgBox "El total de la forma de pago no coincide con el importe"
        Exit Sub
    End If

    If s2n(txtefectivo) > 0 And s2n(txtcodcaja) = 0 Then
        che "falta caja"
        txtcodcaja.SetFocus
        Exit Sub
    End If
    
    If ChequeaChq = False Then
        If MsgBox("Desea chequear los cheques antes de continuar?" & Chr(13) & "Tenga en cuenta que si continua pueden duplicarse.", vbQuestion + vbYesNo) = vbYes Then
            Exit Sub
        End If
    End If
    
    For x = 1 To uCheques.rows
        If Not uCheques.chPropio(x) Then 'cheque de tercero
            If uCheques.chNroInt(x) = 0 Then
                MsgBox "El cheque numero " & uCheques.chNumero(x) & " debe ser seleccionado y no cargado." & Chr(13) & "Haga doble clic para buscarlo.", , "ATENCION"
                Exit Sub
            End If
        End If
    Next


    If Not uCheques.FechasOk Then Exit Sub
    
    
    


    Dim cheque As Boolean, contado As Boolean
    Dim Total As Double
    Dim sumo As Double
    Dim rs As New ADODB.Recordset

If Trim(Ope) <> "" Then
    If Ope = "A" Then
       

         Dim fechapropio As Date
         Dim valcartera As String
         Dim porciva As Double, maximobanc As Long, maximocaja As Long ', x As Long
         Dim valorcuenta As String, valorcuentacon As String, valorcartera As String, Neto As Double, Iva21 As Double
         Dim sucursal As Long
        Dim TextoAsientoComprobante As String
'        Dim retgan As Double
        Dim iddoc As Long, AsientoCompra As New Asiento
'        Dim RetIb As Double
    
'        retgan = s2n(txtRetGan)
'        RetIb = s2n(txtIBpago)
       
        rs.Open "select * from porcentajesiva where iva = " & n2s(ObtenerIvaProv("Prov", s2n(TxtCodProv))) & " order by fecha_baja", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not rs.EOF
            If IsNull(rs!fecha_baja) Then
                porciva = rs!PORCENTAJE
            Else
                porciva = 0
            End If
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        rs.Open "select sucursal from datos", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            sucursal = nSinNull(rs!sucursal)
        Else
            sucursal = 0
        End If
        rs.Close
        Set rs = Nothing
                                    
'        If porciva <> 0 Then
            Neto = s2n((s2n(txtimporte) / (1 + porciva)) * z)
'        Else
'            neto = s2n(txtimporte)
'        End If

        'truchisimo, pero hay que REHACER TODO LO QUE TENGA QUE VER CON IVA
        If s2n(porciva) = 0.21 Then Iva21 = s2n((Neto * porciva) * z) ' 21 / 100
        
        '*********************************************
        '*********************************************
        DE_BeginTrans
        
        If uRetCompras.retgan > 0 Then NroCertifGan = NuevoNroCertifGan()
        If uRetCompras.retIB > 0 Then NroCertifIIBB = NuevoNroCertifIIBB()

        
        iddoc = NuevoDocumento("RAC", s2n(txtopago, 0), s2n(TxtCodProv), s2n(txtopago, 0), NroCertifGan, NroCertifIIBB)
        
        midDoc = iddoc
        'DEBERIA IR EN CABECERA ASIENTO
        TextoAsientoComprobante = "RAC " & txtopago
        'AsientoCompra.Nuevo "Recibo " & txtopago & " " & txtprov, dtfecha, "PAC"
        AsientoCompra.nuevo "Pago " & txtprov, dtFecha, "PAC"
        AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_ANTICIP_A_PROV), s2n(txtimporte) * z, 0, TextoAsientoComprobante
        
        'INGRESO A TRANSCOM
        'If cmbmoneda.ListIndex = -1 Then cmbmoneda.ListIndex = 0
        
        
        'asiento retenciones  gan ib
        AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_GAN_3ros), 0, uRetCompras.retgan * z ', sComprobante
        AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_IB_3ros), 0, uRetCompras.retIB * z 'RetIb
'        AsientoCompra.AcumularItem CuentaParam(ID_CuentasParam_DEUD_A_PROV), retgan, 0 ', sComprobante
        
        
        DataEnvironment1.dbo_INGTRANSCOM "A", dtFecha, s2n(TxtCodProv), txtNombre, CUIT.Text, "RAC", s2n(txtopago) _
          , s2n(txtimporte) * z, s2n(txtimporte) * z, Neto, Iva21, sucursal, s2n(txtserie), ObtenerCodigo("Monedas", cmbMoneda.Text), z, Date, UsuarioSistema!codigo, 0, 0, 1, iddoc, 0, 0, 0, uRetCompras.retgan, uRetCompras.retIB, _
          "", Prov_NumIIBB(s2n(TxtCodProv)), uRetCompras.IB_CodTipo, uRetCompras.IG_CodTipo

        'SI REALIZO UNA TRANSFERENCIA
        'If txttransf <> "" And txttransf <> "0" Then
        If s2n(txttransf) <> 0 Then
'            rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            If Not IsNull(rs!maxcodigo) Then
'                maximobanc = rs!maxcodigo + 1
'            Else
'                maximobanc = 1
'            End If
'            rs.Close
'            Set rs = Nothing
            maximobanc = NuevoMovibanc()
                                
            DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", s2n(txtcodcuenta), "S", "Transf. " & "Prov. " & ObtenerDescripcion("Prov", s2n(TxtCodProv)), _
              dtFecha, "E", 0, s2n(txttransf) * z, "O/P", s2n(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z

        End If
                    
        'SI PAGO CON CHEQUES PROPIOS
'        If txtimpcheques <> "" And txtimpcheques <> "0" Then
        If s2n(txtimpcheques) <> 0 Then
            If ExistenPropios Then
                
'                rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximobanc = rs!maxcodigo + 1
'                Else
'                    maximobanc = 1
'                End If
'                rs.Close
'                Set rs = Nothing
                maximobanc = NuevoMovibanc()
                            
                For x = 1 To uCheques.rows
                    'INGCOMPRAMOVIBANC es igual al STORE del INGCHEQUEMOVIBANC
                    If uCheques.chPropio(x) Then
                        
                        If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                            If uCheques.chNroInt(x) = 0 Then
                                numIntP = nuevoCodigo("chq_Comp")
                                ' cargo por 1ra vez
                                DataEnvironment1.dbo_INGRESOCHEQUERA numIntP, 0, uCheques.chNumero(x), uCheques.chBancCod(x), uCheques.chBancCod(x), _
                                         0, 0, "", 0, "C", 0, 0, Date, UsuarioSistema!codigo, 0, 0, 1
                                uCheques.chSetearNroInt x, numIntP
                            End If
                        End If
                        
                        
                        'mod lito 20/7/6  cuenta = la del cheque
                        cueche = s2n(obtenerDeSQL("select cuentabancaria from chq_comp where codigo = " & uCheques.chNroInt(x)))
                        DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", cueche, "L", "O/P " & txtopago & "Prov. " & ObtenerDescripcion("Prov", s2n(TxtCodProv)), _
                            dtFecha, "P", uCheques.chNroInt(x), uCheques.chMonto(x) * z, "O/P", s2n(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                        'INCREMENTO EL AUTOMATICO DE MOVIBANC
                        maximobanc = maximobanc + 1
                    End If
                Next

            End If
        End If
                    
        'SI PAGO CON CHEQUES DE TERCEROS
        'If txtimpcheques <> "" And txtimpcheques <> "0" Then
        If s2n(txtimpcheques) <> 0 Then
            If ExistenTerceros Then
                
'                rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximobanc = rs!maxcodigo + 1
'                Else
'                    maximobanc = 1
'                End If
'                rs.Close
'                Set rs = Nothing
                maximobanc = NuevoMovibanc()
                            
                For x = 1 To uCheques.rows
                    If Not uCheques.chPropio(x) Then
                        'INGCOMPRAMOVIBANC es igual al STORE del INGCHEQUEMOVIBANC
                        'mod lito 20/7/6  cuenta = 0, no debe aparecer como mov bancario
                        
                        DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", 0, "T", "O/P " & txtopago & "Prov. " & ObtenerDescripcion("Prov", s2n(TxtCodProv)) _
                            , dtFecha, "C", uCheques.chNroInt(x), uCheques.chMonto(x) * z, "O/P", s2n(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                        'INCREMENTO EL AUTOMATICO DE MOVIBANC
                        maximobanc = maximobanc + 1
                    End If
                Next

            End If
        End If
                                
        
        'ACA EMPIEZA LAS ALTAS A MOVICAJA
                                           
        'SI PAGO EN EFECTIVO
'        If txtefectivo <> "" And txtefectivo <> "0" Then
         If s2n(txtefectivo) <> 0 Then
            
'            rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            If Not IsNull(rs!maxcodigo) Then
'                maximocaja = rs!maxcodigo + 1
'            Else
'                maximocaja = 1
'            End If
'            rs.Close
''            Set rs = Nothing
            maximocaja = NuevoMoviCaja()
            
'            rs.Open "select cuenta from Cajas where codigo = " & n2s(txtcodcaja) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If Not rs.EOF Then
'                valorcuenta = rs!cuenta
'            Else
'                valorcuenta = ""
'            End If
'            rs.Close
''            Set rs = Nothing
            valorcuenta = verCuentaContableCaja(s2n(txtcodcaja))
            
            
            'haber EFECTIVO
            AsientoCompra.AgregarItem valorcuenta, 0, s2n(txtefectivo) * z, TextoAsientoComprobante
                                
            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", s2n(txtcodcaja), maximocaja, "E", "E", s2n(txtefectivo) * z, "O/P " & txtopago & "Prov. " & s2n(TxtCodProv), _
              dtFecha, 0, s2n(TxtCodProv), "O/P", s2n(txtopago), valorcuenta, 0, _
              iddoc, Date, UsuarioSistema!codigo, z
        End If
                    
        'SI REALIZO UNA TRANSFERENCIA
        'If txttransf <> "" And txttransf <> "0" Then
        If s2n(txttransf) <> 0 Then
        
'            rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            If Not IsNull(rs!maxcodigo) Then
'                maximocaja = rs!maxcodigo + 1
'            Else
'                maximocaja = 1
'            End If
'            rs.Close
''            Set rs = Nothing
            maximocaja = NuevoMoviCaja()
        
'            rs.Open "select cuenta_con from Ctasbank where codigo = " & n2s(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If Not rs.EOF Then
'                valorcuentacon = rs!cuenta_con
'            Else
'                valorcuentacon = ""
'            End If
'            rs.Close
''            Set rs = Nothing
            valorcuentacon = verCuentaContableBanco(s2n(txtcodcuenta))
                            
'            rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            If Not IsNull(rs!maxcodigo) Then
'                maximobanc = rs!maxcodigo + 1
'            Else
'                maximobanc = 1
'            End If
'            rs.Close
''            Set rs = Nothing
            maximobanc = NuevoMovibanc()
                            
            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "T", "E", s2n(txttransf) * z, "O/P " & txtopago & "Prov. " & s2n(TxtCodProv), _
              dtFecha, 0, s2n(TxtCodProv), "O/P", s2n(txtopago), valorcuentacon, maximobanc, _
              iddoc, Date, UsuarioSistema!codigo, z
            
            'haber  TRANSFERENCIA
            AsientoCompra.AgregarItem obtenerDeSQL("select cuenta_con from ctasbank where activo = 1 and codigo = '" & x2s(s2n(txtcodcuenta)) & "' "), 0, s2n(txttransf) * z, TextoAsientoComprobante
        End If
                   
        'SI PAGO CON CHEQUES PROPIOS
        If s2n(txtimpcheques) <> 0 Then '"" And txtimpcheques <> "0" Then
            If ExistenPropios Then
                                                                        
'                rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximocaja = rs!maxcodigo + 1
'                Else
'                    maximocaja = 1
'                End If
'                rs.Close
                maximocaja = NuevoMoviCaja()
'               'Set rs = Nothing
    
                
'                rs.Open "select cuenta_con from Ctasbank where codigo = " & s2n(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'                If Not rs.EOF Then
'                    valorcuentacon = rs!cuenta_con
'                Else
'                    valorcuentacon = ""
'                End If
'                rs.Close
''                Set rs = Nothing
                valorcuentacon = verCuentaContableBanco(s2n(txtcodcuenta))
                
'                rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximobanc = rs!maxcodigo + 1
'                Else
'                    maximobanc = 1
'                End If
'                rs.Close
''                Set rs = Nothing
                maximobanc = NuevoMovibanc()
                
                For x = 1 To uCheques.rows
                    If uCheques.chPropio(x) Then
                        DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "P", "E", uCheques.chMonto(x) * z, "O/P " & txtopago & "Prov. " & s2n(TxtCodProv), _
                          dtFecha, uCheques.chNroInt(x), s2n(TxtCodProv), "O/P", s2n(txtopago), valorcuentacon, maximobanc, _
                          iddoc, Date, UsuarioSistema!codigo, z
                        DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "A", uCheques.chNroInt(x), uCheques.chFecha(x), uCheques.chMonto(x) * z _
                            , s2n(txtopago), "RAC", s2n(TxtCodProv), "T", uCheques.chFecha(x), dtFecha, Date, UsuarioSistema!codigo, 0, 0, 1, z, ObtenerCodigo("Monedas", cmbMoneda.Text)
                        
                        'INCREMENTO EL AUTOMATICO DE MOVIBANC
                        maximobanc = maximobanc + 1
                        
                        'haber CHEQUE PROPIO
                        'AsientoCompra.AcumularItem sSinNull(obtenerDeSQL("select cuenta from ctasBank where  codigo = '" & uCheques.chBancCod(x) & "' and activo = 1")), 0, uCheques.chMonto(x)
                        AsientoCompra.AcumularItem uCheques.chCuenta(x), 0, uCheques.chMonto(x) * z
                    End If
                Next
            End If
        End If
                   
                   
        'SI PAGO CON CHEQUES TERCEROS
        'If txtimpcheques <> "" And txtimpcheques <> "0" Then
        If s2n(txtimpcheques) <> 0 Then
            If ExistenTerceros Then
                                                
'                rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximocaja = rs!maxcodigo + 1
'                Else
'                    maximocaja = 1
'                End If
'                rs.Close
''                Set rs = Nothing
                maximocaja = NuevoMoviCaja()


'                rs.Open "select valores_cartera from Imputaciones", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not rs.EOF Then
'                    valcartera = rs!valores_cartera
'                Else
'                    valcartera = ""
'                End If
'                rs.Close
'                Set rs = Nothing
                valcartera = CuentaParam(ID_Cuenta_M_CH_CARTERA)
                
'                rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not IsNull(rs!maxcodigo) Then
'                    maximobanc = rs!maxcodigo + 1
'                Else
'                    maximobanc = 1
'                End If
'                rs.Close
''                Set rs = Nothing
                maximobanc = NuevoMovibanc()
                
                For x = 1 To uCheques.rows 'FrmCheques.grillaterceros.rows - 1
                    If Not uCheques.chPropio(x) Then
                        DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "C", "E", uCheques.chMonto(x) * z, "O/P " & txtopago & "Prov. " & s2n(TxtCodProv), _
                          dtFecha, uCheques.chNroInt(x), s2n(TxtCodProv), "O/P", s2n(txtopago), valcartera, maximobanc, _
                          iddoc, Date, UsuarioSistema!codigo, z
                        

                        DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "A", uCheques.chNroInt(x), 0, "", s2n(TxtCodProv), s2n(txtopago), 0, _
                          dtFecha, "T", "RAC", Date, UsuarioSistema!codigo, 0, 0, 1, 1, z
'                         dtfecha, "T", "FDC", Date, UsuarioSistema!Codigo, 0, 0, 1

                        'INCREMENTO EL AUTOMATICO DE MOVIBANC
                        maximobanc = maximobanc + 1
                                
                                
                        'Haber Cheques 3ros
                        AsientoCompra.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, uCheques.chMonto(x) * z
                        
                    End If
                Next
            End If
        End If
        
       'if
       If siAsiento("AsientosPagos") Then AsientoCompra.Grabar iddoc
       '= 0 Then
       '     DE_RollbackTrans
       '     ufa "", " asiento  - "
       '     Exit Sub
       ' End If
        
        
        DE_CommitTrans
        '*********************************************
        '*********************************************
        
        
        If ON_ERROR_HABILITADO Then On Error GoTo UFAimpresion
        
        'INCREMENTO NUM_OPAGO DE LA TABLA BS
'        DataEnvironment1.dbo_INCREMENTONUMOPAGO s2n(txtopago)
        
        MsgBox "Operación Realizada con éxito", vbOKOnly
        HabilitoControles (False)
        Call Habilitobotones(True, True, False, False, False, True)
        
        ImprimirPagoCuenta
        LimpioControles

'        FormadePago (False)
        uCheques.Borrar
        
        If gEMPR_ConSistContable Then
            FrmCostosYContable.LimpioControles
'            FrmCostosYContable.InicioGrilla
        End If
    End If
End If

fin:
    Set rs = Nothing
    Exit Sub
UFAimpresion:
    che "Pago a cuenta grabado, fallo la impresion"
    Resume fin
UfaOK:
    DE_RollbackTrans
    uCheques.resetNroIntPropios
    ufa "err en el alta", "iddoc " & midDoc & " op " & txtopago
    midDoc = 0
    Resume fin
End Sub


Private Sub ImprimirPagoCuenta()
On Error GoTo UFAimprimir

Dim stblChequesOPtmp As String
Dim Localidad, direccion As String
Dim r As Long
Dim str, sql, donde, sqlTemp As String
Dim rs As New ADODB.Recordset

donde = " Pago a cuenta principal"

stblChequesOPtmp = TablaTempCrear(tt_ChequeOPtmp)
With uCheques
    For r = 1 To .rows
        sqlTemp = "insert into " & stblChequesOPtmp _
        & " (nroint, banco, cheque, importe, fecha, propio) values( " _
        & .chNroInt(r) & ", '" & .chBancDes(r) & "', '" & .chNumero(r) & "', " & x2s(.chMonto(r)) & ", " & ssFecha(.chFecha(r)) & ", '" & IIf(.chPropio(r), "P", "T") & "')"
        DataEnvironment1.Sistema.Execute sqlTemp
    Next
    
   rs.Open "select * from " & stblChequesOPtmp, DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    While Not rs.EOF
      Debug.Print rs!NroInt
      rs.MoveNext
    Wend
    Set rs = Nothing
    
    Localidad = obtenerDeSQL("select localidad from prov where codigo = " & TxtCodProv & " ")
    direccion = obtenerDeSQL("select direccion from prov where codigo = " & TxtCodProv & " ")
End With
    str = "select * from " & stblChequesOPtmp
    RptOrdenPagoAcuenta.Data.Connection = DataEnvironment1.Sistema
    RptOrdenPagoAcuenta.Data.Source = str
    RptOrdenPagoAcuenta.lblfecha = dtFecha
    RptOrdenPagoAcuenta.NroCertificado = Format(VerNroPago(midDoc), "0001-00000000")
    RptOrdenPagoAcuenta.txtProveedor = txtprov
    RptOrdenPagoAcuenta.TxtDomicilioProv = direccion
    RptOrdenPagoAcuenta.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoAcuenta.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoAcuenta.lblefectivo = Format(txtefectivo, "#,##0.00")
    RptOrdenPagoAcuenta.lblcheques = Format(txtimpcheques, "#,##0.00")
    RptOrdenPagoAcuenta.lbltransf = Format(txttransf, "#,##0.00")
    RptOrdenPagoAcuenta.LblRetGan = Format(uRetCompras.retgan, "#,##0.00")
    RptOrdenPagoAcuenta.LblretIB = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoAcuenta.lbltotal = Format(txtimporte, "#,##0.00")
    
    
    
donde = "Pago a cuenta Constancia RET. IMPUESTO GANANCIA"
sql = "select tipodoc,nrodoc,total as saldo from transcom where iddoc =  " & midDoc
If uRetCompras.retgan > 0 Then
    
    RptOrdenPagoConstRet_IG.DataImp_Ganancia.Connection = DataEnvironment1.Sistema
    RptOrdenPagoConstRet_IG.DataImp_Ganancia.Source = sql 'esto no sirve pero lo hago para q no se rompa para las OP
    RptOrdenPagoConstRet_IG.lblfecha = dtFecha
    RptOrdenPagoConstRet_IG.LblRegimen_IG = uRetCompras.IG_Tipo
    RptOrdenPagoConstRet_IG.txtProveedor = txtprov 'uProv.descripcion
    RptOrdenPagoConstRet_IG.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConstRet_IG.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & TxtCodProv & " ")
    Localidad = obtenerDeSQL("select localidad from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoConstRet_IG.RG_PagosTotalMes = Format(txtimporte, "#,##0.00")
    RptOrdenPagoConstRet_IG.retgan = Format(uRetCompras.retgan, "#,##0.00")
    RptOrdenPagoConstRet_IG.retganEnPesos = enletras(uRetCompras.retgan)
    RptOrdenPagoConstRet_IG.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
    RptOrdenPagoConstRet_IG.Txtop.Text = Format(txtopago, "00000000")
    RptOrdenPagoConstRet_IG.Label9.Visible = False
    RptOrdenPagoConstRet_IG.Label15.Visible = False
    RptOrdenPagoConstRet_IG.Label16.Visible = False
    RptOrdenPagoConstRet_IG.Label18.Visible = False
    RptOrdenPagoConstRet_IG.fieTipoDoc.Visible = False
    RptOrdenPagoConstRet_IG.fieNroDoc.Visible = False
    RptOrdenPagoConstRet_IG.fieSaldo.Visible = False
    
    RptOrdenPagoConsRet_IG_calculo.lblfecha = dtFecha
    RptOrdenPagoConsRet_IG_calculo.txtProveedor = txtprov
    RptOrdenPagoConsRet_IG_calculo.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConsRet_IG_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoConsRet_IG_calculo.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.RG_MinimoNoImponible = uRetCompras.RG_MinimoNoImponible
    RptOrdenPagoConsRet_IG_calculo.RG_TxtFormula = uRetCompras.RG_TxtFormula
    RptOrdenPagoConsRet_IG_calculo.retgan = Format(uRetCompras.retgan, "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.RG_PagosAnterioresMes = Format(uRetCompras.RG_PagosAnterioresMes, "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.RG_PagosRetAnteriores = Format(uRetCompras.RG_PagosRetAnteriores, "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
    RptOrdenPagoConsRet_IG_calculo.LblRetGanPesos = enletras(uRetCompras.retgan)
    RptOrdenPagoConsRet_IG_calculo.Pago_Fecha = Format(Abs(CDbl(uRetCompras.RG_PagosAnterioresMes) - CDbl(uRetCompras.RG_PagosTotalMes)), "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.Total_Imponible = Format(CDbl(uRetCompras.RG_PagosRetAnteriores) + CDbl(uRetCompras.retgan), "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.Printer.Copies = 1
    RptOrdenPagoConstRet_IG.Printer.Copies = 2
    
    RptOrdenPagoConstRet_IG.Restart
    RptOrdenPagoConsRet_IG_calculo.Restart
    
    If PREVIEW_IMPRESIONES Then
        RptOrdenPagoConstRet_IG.Show
        RptOrdenPagoConsRet_IG_calculo.Show
    Else
        RptOrdenPagoConstRet_IG.PrintReport False
        RptOrdenPagoConsRet_IG_calculo.PrintReport False
    End If
   
   End If
donde = "Pago a cunta Constancia RET. INGRESO BRUTOS"
   If uRetCompras.retIB > 0 Then
    RptOrdenPagoConstRet_IB.DataImp_IB.Connection = DataEnvironment1.Sistema
    RptOrdenPagoConstRet_IB.DataImp_IB.Source = sql '-----------------
    RptOrdenPagoConstRet_IB.lblfecha = dtFecha
    RptOrdenPagoConstRet_IB.LblRegimen_IIBB = uRetCompras.IB_Tipo
    RptOrdenPagoConstRet_IB.txtProveedor = txtprov
    RptOrdenPagoConstRet_IB.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConstRet_IB.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoConstRet_IB.txtNroIIBB = obtenerDeSQL("select numiibb from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoConstRet_IB.RG_PagosTotalMes = Format(txtimporte, "#,##0.00")
    RptOrdenPagoConstRet_IB.retgan = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConstRet_IB.retganEnPesos = enletras(uRetCompras.retIB)
    RptOrdenPagoConstRet_IB.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
    RptOrdenPagoConstRet_IB.Txtop = Format(txtopago, "00000000")
    RptOrdenPagoConstRet_IB.Label9.Visible = False
    RptOrdenPagoConstRet_IB.Label15.Visible = False
    RptOrdenPagoConstRet_IB.Label16.Visible = False
    RptOrdenPagoConstRet_IB.Label18.Visible = False
    RptOrdenPagoConstRet_IB.fieTipoDoc.Visible = False
    RptOrdenPagoConstRet_IB.fieNroDoc.Visible = False
    RptOrdenPagoConstRet_IB.fieSaldo.Visible = False
    
    RptOrdenPagoConsRet_IB_calculo.lblfecha = dtFecha
    RptOrdenPagoConsRet_IB_calculo.txtProveedor = txtprov
    RptOrdenPagoConsRet_IB_calculo.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConsRet_IB_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & TxtCodProv & " ")
    RptOrdenPagoConsRet_IB_calculo.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.IB_TxtFormula = uRetCompras.IB_TxtFormula
    RptOrdenPagoConsRet_IB_calculo.retIB = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.IB_base = Format(uRetCompras.IB_base, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.retIB1 = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
    RptOrdenPagoConsRet_IB_calculo.LblRetIIBBPesos = enletras(uRetCompras.retIB)
    RptOrdenPagoConsRet_IB_calculo.Printer.Copies = 1
    RptOrdenPagoConstRet_IB.Printer.Copies = 2
    
    RptOrdenPagoConstRet_IB.Restart
    RptOrdenPagoConsRet_IB_calculo.Restart
    
    If PREVIEW_IMPRESIONES Then
        RptOrdenPagoConstRet_IB.Show
        RptOrdenPagoConsRet_IB_calculo.Show
    Else
        RptOrdenPagoConstRet_IB.PrintReport False
        RptOrdenPagoConsRet_IB_calculo.PrintReport False
    End If
    
End If

If PREVIEW_IMPRESIONES Then
   RptOrdenPagoAcuenta.Show
  Else
   RptOrdenPagoAcuenta.Restart
   RptOrdenPagoAcuenta.PrintReport False
End If

FinOK:
    Exit Sub
UFAimprimir:
    ufa "Pago a cuenta grabado, falló la impresión " & donde, Me.Name & " - " & donde
    Resume FinOK
End Sub


Function ExistenTerceros() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If Not uCheques.chPropio(x) Then
            ExistenTerceros = True
            Exit Function
        End If
    Next x
End Function

Function ExistenPropios() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If uCheques.chPropio(x) Then
            ExistenPropios = True
            Exit Function
        End If
    Next x
End Function

Private Sub cmdProv_Click()
    On Error Resume Next
    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo, descripcion as [ Proveedor                                  ] from prov where activo = 1")
    If resu > "" Then
        TxtCodProv = frmBuscar.resultado
        txtprov = frmBuscar.resultado(2)
        txtNombre = frmBuscar.resultado(2)
        CUIT.Text = ObtenerCuit("Prov", s2n(TxtCodProv))
'        FormadePago (True)
        txtimporte.SetFocus
        ver_IIBBca
    End If
    cargar = "C"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub LimpioControles()
    TxtCodProv = ""
    txtprov = ""
    txtopago = ""
    txtNombre = ""
    CUIT.Text = ""
    dtFecha = Date
    txttipoiva = ""
    txtserie = ""
    cmbMoneda.ListIndex = -1
    txtcotiz = ""
    
    txtimporte = "0"
    txtimpcheques = "0"
    txtefectivo = "0"
    txttransf = "0"

    txtcodcuenta = ""
    txtcuenta = ""
    
    txtsaldo = ""
    cargar = ""
    
    Ope = ""
    
    txtcaja = ""
    txtcodcaja = "1"
    txtcodcaja_LostFocus
    
    uCheques.Borrar
    midDoc = 0
    uRetCompras.retgan = 0
    uRetCompras.retIB = 0
End Sub


Sub HabilitoControles(habilito As Boolean)
'    txtcodprov.Enabled = habilito
'    txtopago.Enabled = habilito
    txtserie.enabled = habilito
    txtNombre.enabled = habilito
    CUIT.enabled = habilito
    dtFecha.enabled = habilito
'    cmbformapago.Enabled = habilito
'    txtfvto.Enabled = habilito
'    optcontado.Enabled = habilito
'    optctacte.Enabled = habilito
    cmbMoneda.enabled = habilito
    txtcotiz.enabled = habilito
'    txtplan.Enabled = habilito
    
'    txtneto.Enabled = habilito
'    txtiva.Enabled = habilito
'    txtper3337.Enabled = habilito
    txtimporte.enabled = habilito
'    txtiva27.Enabled = habilito
'    txtexento.Enabled = habilito
'    txtreteniva.Enabled = habilito
'    txtimpint.Enabled = habilito
'    txtretengan.Enabled = habilito
'    txtper3431.Enabled = habilito
'    txtiva10.Enabled = habilito
'    txtpercep.Enabled = habilito
'    cmdprov.Enabled = habilito
'    txtnombre.Enabled = habilito
    
'    txtanio.Enabled = habilito
'    txtmes.Enabled = habilito
'    txtsuc.Enabled = habilito
'    txtserie.Enabled = habilito
'    txtcotiz.Enabled = habilito
'    cmbtipocompra.Enabled = habilito
'    cmbmoneda.Enabled = habilito
    uCheques.enabled = habilito
End Sub


Public Sub CargarDatos()
    
    Dim rs As New ADODB.Recordset
    Dim codigo As String

    If rsmov.State = 1 Then
        rsmov.Close
        Set rsmov = Nothing
    End If
    
    codigo = Trim(Me.Tag)
    
    If cargar = "CuentasBank" Then
        rs.Open "select * from Ctasbank where codigo = " & n2s(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcuenta = rs!codigo
            txtcuenta = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Cajas" Then
        rs.Open "select * from Cajas where codigo = " & n2s(txtcodcaja) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcaja = rs!codigo
            txtcaja = rs!responsable
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "C" Then
        rs.Open "select * from Prov where codigo = " & n2s(TxtCodProv) & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            TxtCodProv = rs!codigo
            txtprov = rs!DESCRIPCION
            txtNombre = rs!DESCRIPCION
            txttipoiva = ObtenerIvaProv("Prov", TxtCodProv)
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "BU" Then
        
        If TxtCodProv <> "" Then
            rsmov.Open "select * from " & ssDonde & " where fecha = " & ssFecha(dtFecha) & " and codpr = " & n2s(TxtCodProv) & " and tipodoc = '" & txttipocompra & "' and nrodoc = " & n2s(txtopago) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Else
            rsmov.Open "select * from " & ssDonde & " where fecha = " & ssFecha(dtFecha) & " and tipodoc = '" & txttipocompra & "' and nrodoc = " & n2s(txtopago) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        End If
        If Not rsmov.EOF Then
            CargoRegistro
        End If
        rsmov.Close
        Set rsmov = Nothing
    End If
    
End Sub

Sub CargoRegistro()
    If ON_ERROR_HABILITADO Then On Error GoTo ufacargo
    Dim z As Double
    
    TxtCodProv = rsmov!CODPR
    txtprov = ObtenerDescripcion("Prov", rsmov!CODPR)
    txtserie = rsmov!Serie
    txtopago = rsmov!NroDoc
    txtNombre = rsmov!razonsocialprov
    CUIT.Text = rsmov!cuitprov
'    cmbtipocompra = BuscoDato("TipoCompras", rsmov!tipocompra)
    dtFecha = rsmov!Fecha
    'txtRetGan = s2n(rsmov!retgan)
'    uRetCompras.retgan = rsmov!retgan
    
'    txtneto = rsmov!neto
'    txtiva = rsmov!iva_21
'    txtper3337 = rsmov!percepc
'    txtimporte = rsmov!total
'    txtiva27 = rsmov!iva_27
'    txtexento = rsmov!exento
'    txtreteniva = rsmov!iva_9
'    txtimpint = rsmov!imp_int
'    txtretengan = rsmov!ret_gan
'    txtper3431 = rsmov!der_est
'    txtiva10 = rsmov!iva_10
'    txtpercep = rsmov!perceib
    
'    txtanio = rsmov!anoimp
'    txtmes = rsmov!mesimp
'    txtsuc = rsmov!suc
'    txtserie = rsmov!serie
    z = s2n(nSinNull(rsmov!cotizacion), 4)
    If z = 0 Then z = 1
    txtcotiz = z
    
    If UCase(Trim(ssDonde)) = "COMPRAS" Then
    Else
        txtsaldo = rsmov!saldo / z
    End If
    cmbMoneda = ObtenerDescripcion("Monedas", rsmov!moneda)
    
    
    'txttransf = ObtenerTransferencia("Movibanc", rsmov!NroDoc, rsmov!CODPR) / z
    txttransf = s2n(obtenerDeSQL("select importe from movibanc where operacion='S' and iddoc=" & rsmov!iddoc) / z)
    'txtcodcuenta = ObtenerCuenta("Movicaja", rsmov!NroDoc, rsmov!CODPR)
    txtcodcuenta = s2n(obtenerDeSQL("select cuenta from movibanc where OPERACION='S' and iddoc=" & rsmov!iddoc))
    txtcuenta = ObtenerDescripcionCuentas("ctasbank", s2n(txtcodcuenta))
    txtcodcaja = ObtenerCaja("Movicaja", rsmov!NroDoc, rsmov!CODPR)
    txtcaja = ObtenerDescripcionCajas("Cajas", s2n(txtcodcaja))
    txtefectivo = ObtenerImporte("Movicaja", rsmov!NroDoc, rsmov!CODPR) / z
'    txtefectivo = obtenerdesql("select
    txtimpcheques = ObtenerTotalCheques("Movicaja", rsmov!NroDoc, rsmov!CODPR) / z
   uRetCompras.retgan = nSinNull(rsmov!retganpago) / z
   uRetCompras.retIB = nSinNull(rsmov!IBPAGO) / z

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    CargoCheques
'    LlenarGrilla GrillaMoviCaja, "Select Fecha, nro as 'Numero Cheque', Importe From CHEQUES " & _
'                        "Where ACTIVO = 1 And TDOC = '" & TIPODOC & "' AND NDOC = " & NroDoc, True
'    LlenarGrilla GrillaEfectivo, "Select Fecha, Importe From MOVICAJA " & _
'                        "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
'                            "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True

fin:
    Exit Sub
ufacargo:
    ufa "err al cargar datos", "carga op"
    Resume fin
End Sub

Private Sub CargoCheques()
    Dim rs As New ADODB.Recordset
    uCheques.Borrar

'    rs.Open "select Movicaja.*, Chq_comp.banco from Movicaja inner join Chq_comp on Movicaja.interno = Chq_comp.codigo where Movicaja.codprov = " & n2s(txtcodprov) & " and Movicaja.nrodoc = " & n2s(txtopago) & " and Movicaja.tipodoc = 'O/P' and Movicaja.tipo = 'P'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rs.Open "select Movicaja.*, Chq_comp.banco from Movicaja inner join Chq_comp on Movicaja.interno = Chq_comp.codigo where  Movicaja.iddoc=" & midDoc & " and Movicaja.nrodoc = " & n2s(txtopago) & " and Movicaja.tipodoc = 'O/P' and Movicaja.tipo = 'P'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        uCheques.metoCheque uCheques.rows + 1, rs!interno, "P"
        rs.MoveNext
    Wend
    rs.Close
    
    'rs.Open "select Movicaja.*, Cheques.banco_nro from Movicaja inner join Cheques on Movicaja.interno = Cheques.nroint where Movicaja.codprov = " & n2s(txtcodprov) & " and Movicaja.nrodoc = " & n2s(txtopago) & " and Movicaja.tipodoc = 'O/P' and Movicaja.tipo = 'C'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rs.Open "select Movicaja.*, Cheques.banco_nro from Movicaja inner join Cheques on Movicaja.interno = Cheques.nroint where  Movicaja.iddoc=" & midDoc & " and Movicaja.nrodoc = " & n2s(txtopago) & " and Movicaja.tipodoc = 'O/P' and Movicaja.tipo = 'C'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        uCheques.metoCheque uCheques.rows + 1, rs!interno, "T"
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
End Sub


Private Sub Form_Load()
'    CargaCombo2 cmbformapago, "FormasPago", "descripcion", "codigo", ""
'    CargaCombo2 cmbtipocompra, "TipoCompras", "descripcion", "codigo", ""
    CargaCombo3 cmbMoneda, "Monedas", "descripcion", "codigo", ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub


Sub Habilitobotones(buscar As Boolean, agregar As Boolean, eliminar As Boolean, _
                    Imprimir As Boolean, aceptar As Boolean, Cancelar As Boolean)
    cmdbuscar.enabled = buscar
    cmdcancelar.enabled = Cancelar
    cmdeliminar.enabled = eliminar
    cmdnuevo.enabled = agregar
    cmdok.enabled = aceptar
    cmdImprimir.enabled = Imprimir
End Sub

Function ObtenerCuenta(tabla As String, nDoc As Long, prov As Long) As Long
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset

Dim sqlstrCC As String, Cuenta As Long
    
    sqlstrCC = "Select movbanco from " + tabla + " where nrodoc = " & nDoc & " and codprov = " & prov & "  and tipo = 'T' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        Cuenta = rs!MovBanco
        sqlstrCC = "Select cuenta from Movibanc where movbanco = " & Cuenta & " and activo=1"
        rs1.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            ObtenerCuenta = rs1!Cuenta
        Else
            ObtenerCuenta = 0
        End If
        rs1.Close
        Set rs1 = Nothing
    Else
        ObtenerCuenta = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function ObtenerImporte(tabla As String, nDoc As Long, prov As Long) As Long
    Dim rs As New ADODB.Recordset

    Dim sqlstrCC As String
    
    sqlstrCC = "Select importe from " + tabla + _
                " where nrodoc = " & nDoc & " and codprov = " & prov & "  and TIPODOC = '" & RECIBOS_A_CUENTA_MOVICAJA & "' and tipo = 'E' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerImporte = rs!Importe
    Else
        ObtenerImporte = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function


Function ObtenerCaja(tabla As String, nDoc As Long, prov As Long) As Long
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String
    
    sqlstrCC = "Select caja from " + tabla + _
                " where nrodoc = " & nDoc & " and codprov = " & prov & " and TIPODOC = '" & RECIBOS_A_CUENTA_MOVICAJA & "' and tipo = 'E' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerCaja = rs!caja
    Else
        ObtenerCaja = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function ObtenerTotalCheques(tabla As String, nDoc As Long, prov As Long) As Double
Dim rs As New ADODB.Recordset

Dim Total As Double
Dim sqlstrCC As String
    
    sqlstrCC = "Select importe from " + tabla + _
                " where nrodoc = " & nDoc & " and codprov = " & prov & "  and TIPODOC = '" & RECIBOS_A_CUENTA_MOVICAJA & "' and (tipo = 'P' or tipo = 'C') and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Total = 0
    While Not rs.EOF
        Total = Total + rs!Importe
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    ObtenerTotalCheques = Total
    
End Function


Function ObtenerTransferencia(tabla As String, nDoc As Long, prov As Long) As Double
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String
    
    sqlstrCC = "Select importe from " + tabla + _
                    " where nrodoc = " & nDoc & " and codprov = " & prov & " and TIPODOC = '" & RECIBOS_A_CUENTA_MOVICAJA & "' and tipo = 'T' and activo=1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        ObtenerTransferencia = rs!Importe
    Else
        ObtenerTransferencia = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function BuscoDato(tabla As String, dato As Long) As String
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String
Dim i As Long
    
    sqlstrCC = "Select descripcion from " + tabla + _
                    " where Codigo = " & dato & " and TIPODOC = '" & RECIBOS_A_CUENTA_MOVICAJA & "' and activo = 1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        BuscoDato = rs!DESCRIPCION
    End If
    rs.Close
    Set rs = Nothing
    
End Function

'Private Sub Cuit_CuitInvalido(Nro As String)
'    MsgBox "Cuit Inválido"
'    cuit.SetFocus
'End Sub
'
''Private Sub optcontado_Click()
'    If cargar <> "BU" Then
'        cmbformapago.Enabled = False
'        txtfvto.Enabled = False
'        If fecha.Enabled = True Then
'            fecha.SetFocus
'        End If
'
'        frmCheques.grilla.Clear
'        frmCheques.txttotal = "0"
'        frmCheques.Limpiogrillas
'        frmCheques.InicioGrilla
'    End If
'End Sub

'Private Sub optctacte_Click()
'    cmbformapago.Enabled = True
'    txtfvto.Enabled = True
'    FormadePago (False)
'    txtcodcuenta.Enabled = False
'    cmbcuenta.Enabled = False
'    If fecha.Enabled = True Then
'        fecha.SetFocus
'    End If
'End Sub


Private Sub txtcodcaja_GotFocus()
    If Trim$(txtcodcaja) = "" Then txtcodcaja = "1"
    PintoFocoActivo
End Sub
Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub
Private Sub txtcodcuenta_GotFocus()
'    txtcodcuenta.SelStart = 0
'    txtcodcuenta.SelLength = Len(txtcodcuenta.Text)
    PintoFocoActivo
End Sub

Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcodprov_GotFocus()
'    If txtopago = "" And txtopago.Enabled = True Then
'        MsgBox "Debe ingresar un Nº de comprobante"
'    End If
    PintoFocoActivo
End Sub
Private Sub txtcodprov_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub
Private Sub txtcodprov_LostFocus()
    On Error Resume Next
    If Trim(TxtCodProv) <> "" Then
        txtprov = ObtenerDescripcion("Prov", s2n(TxtCodProv))
        If txtprov = "" Then
            MsgBox "Proveedor incorrecto"
            TxtCodProv.SetFocus
        Else
            If txtopago > "" Then 'And Not ExisteRecCompraMSG(s2n(TxtCodProv), s2n(txtopago)) Then   '(Not estarepetido("Compras", s2n(txtcodprov), s2n(txtopago)) And Not estarepetido("transcom", s2n(txtcodprov), s2n(txtopago))) Then
                txtNombre = ObtenerDescripcion("Prov", s2n(TxtCodProv))
                CUIT.Text = ObtenerCuit("Prov", s2n(TxtCodProv))
'                txtsuc = ObtenerSucursal("Prov", Val(txtcodprov))
                txttipoiva = ObtenerIvaProv("Prov", TxtCodProv)
                txtimporte.SetFocus
                cmdbuscar.enabled = True
            'Else
            '    MsgBox ""Se repite el Nº de comprobante para este proveedor o el nro no fue cargado"
            End If
        End If
'    Else
'        If txtopago <> "" Then
'            MsgBox "Debe ingresar un proveedor"
'            txtcodprov.SetFocus
'        Else
'            txtopago.SetFocus
'        End If
        ver_IIBBca
    End If
End Sub

Private Function ver_IIBBca()
    Dim Propio As Boolean
    Propio = obtenerDeSQL("select  conretiibbper from Prov where codigo = " & TxtCodProv)
    If Propio = True Then
        If MsgBox("El proveedor que selecciono tiene Retencion de IIBB personal." & Chr(13) & Chr(13) & "¿Desea utilizarlo?", vbCritical + vbYesNo, "Alvertencia") = vbYes Then
            uRetCompras.tieneIIBB = True
        Else
            uRetCompras.tieneIIBB = False
        End If
    Else
        uRetCompras.tieneIIBB = False
    End If
    Propio = obtenerDeSQL("select  conretganper from Prov where codigo = " & TxtCodProv)
    If Propio = True Then
        If MsgBox("El proveedor que selecciono tiene Retencion de Ganancias personal." & Chr(13) & Chr(13) & "¿Desea utilizarlo?", vbCritical + vbYesNo, "Alvertencia") = vbYes Then
            uRetCompras.tieneGAN = True
        Else
            uRetCompras.tieneGAN = False
        End If
    Else
        uRetCompras.tieneGAN = False
    End If
    
    
End Function

'Function estarepetido(Tabla As String, prov As Integer, codigo As Long) As Boolean
'    Dim rs As New ADODB.Recordset
'    rs.Open "select * from " & Tabla & " where codpr = " & prov & " and nrodoc = " & codigo & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    'If Not rs.EOF Then
'    '    estarepetido = True
'    'Else
'    '    estarepetido = False
'    'End If
'    estarepetido = Not rs.EOF
'    rs.Close
'    Set rs = Nothing
'End Function

'Function ObtenerSucursal(Tabla As String, codigo As Integer) As Integer
'Dim rs As New ADODB.Recordset
'
'    rs.Open "select suc from " & Tabla & " where codigo = " & Trim(codigo) & "", daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'
'    If Not rs.EOF Then
'        ObtenerSucursal = rs!suc
'    Else
'        ObtenerSucursal = 0
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'End Function

Function ObtenerCuit(tabla As String, codigo As Long) As String
    Dim rs As New ADODB.Recordset
    rs.Open "select cuit from " & tabla & " where codigo = " & Trim(codigo) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        ObtenerCuit = rs!CUIT
    Else
        ObtenerCuit = ""
    End If
    rs.Close
    Set rs = Nothing
End Function


Private Sub txtcotiz_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcotiz_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtcotiz_LostFocus()
    txtcotiz = s2n(txtcotiz)
End Sub

Private Sub txtEfectivo_GotFocus()
    'TxtEfectivo = s2n(txtimporte) - (s2n(txtimpcheques) + s2n(txttransf) + uRetCompras.TotalRet) 's2n(txtRetGan))
    txtefectivo = nuevoMonto(txtefectivo)
    PintoFocoActivo
End Sub

Private Sub txtefectivo_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Sub txtEfectivo_LostFocus()
    Dim Efectivo As Double
    
    Efectivo = s2n(txtefectivo)
    txtefectivo = Efectivo
    
    If Efectivo = s2n(txtimporte) Then
        'txtimpcheques.Enabled = False:
        uCheques.enabled = False
'        txttransf.enabled = False
'        txtcodcuenta.enabled = False
'        cmbcuenta.enabled = False
        cmbcaja.enabled = True
        txtcodcaja.enabled = True
    Else
        If Efectivo <> 0 Then '"" Then
            'If s2n(s2n(txtefectivo, 4) + s2n(txtimpcheques, 4) + s2n(txttransf, 4)) > s2n(txtimporte, 4) Then
            '    MsgBox "Con este valor esta superando al importe"
            '    Exit Sub
            'End If
            If CheMePase() Then Exit Sub
            
        '    If efectivo < s2n(txtimporte) Then
                'txtimpcheques.Enabled = True:
                uCheques.enabled = True
'                txttransf.enabled = True
'                txtcodcuenta.enabled = True
'                cmbcuenta.enabled = True
                cmbcaja.enabled = True
                txtcodcaja.enabled = True
        '    Else
        '        MsgBox "El importe en efectivo no puede superar al importe del comprobante"
        '        Exit Sub
        '    End If
        'Else
        '    txtefectivo = "0"
        End If
    End If
    
    nuevoMonto 0
End Sub

Private Sub txtopago_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtopago_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

'Private Sub txtopago_LostFocus()
'    If txtopago = "" Then
'        MsgBox "Debe ingresar el Nº de comprobante"
'        txtopago.SetFocus
'    End If
'End Sub

'Private Sub txtimpcheques_GotFocus()
'    'txtimpcheques = s2n(s2n(txtimporte, 4) - (s2n(txtefectivo, 4) + s2n(txttransf, 4)), 4)
'    txtimpcheques = nuevoMonto(txtimpcheques)
'    'txtimpcheques.SelStart = 0
'    'txtimpcheques.SelLength = Len(txtimpcheques.Text)
'    PintoFocoActivo
'End Sub
'Private Sub txtimpcheques_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub

'Private Sub txtimpcheques_LostFocus()
'
'    Dim habi As Boolean
'    txtimpcheques = s2n(txtimpcheques)
'
'    habi = (s2n(txtimpcheques) > 0) And (s2n(s2n(txtimpcheques) + s2n(txtefectivo)) <> s2n(txtimporte))
'    txttransf.Enabled = habi
'    txtcodcuenta.Enabled = habi
'    cmbcuenta.Enabled = habi
'
'End Sub

Private Sub txtimporte_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtimporte_LostFocus()
    Dim impo As Double
    
    txtimporte = s2n(txtimporte)
    impo = txtimporte
    uRetCompras.Calcular s2n(TxtCodProv), impo, impo, dtFecha
    'uRetCompras.Calcular s2n(TxtCodProv), impo / (1 + ProvCoefIVA(s2n(TxtCodProv))), dtFecha
    'txtRetGan = s2n(CalculaRetGan(s2n(txtcodprov), s2n(txtimporte), dtfecha))
    
    'refresco
    nuevoMonto 0
End Sub

Private Sub txtnombre_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtnombre_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttransf_GotFocus()
    txttransf = nuevoMonto(txttransf)
    PintoFocoActivo
End Sub
Private Sub txttransf_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub
Private Sub txttransf_LostFocus()
    Dim tran As Double, habil As Boolean
    tran = s2n(txttransf)
    txttransf = tran
    habil = tran <> 0
    If CheMePase Then
        habil = False
        Exit Sub
    Else
        txtcodcuenta.enabled = habil
        cmbcuenta.enabled = habil
    End If
    
    nuevoMonto 0
End Sub

Private Sub txtcodcuenta_LostFocus()
    If IsNumeric(txtcodcuenta) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & n2s(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
        
        If txtcuenta = "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcuenta = "0"
            'txtcodcuenta.SetFocus
        Else
            cargar = "CuentasBank"
            CargarDatos
        End If
    Else
        If txtcodcuenta <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcuenta = "0"
            txtcodcuenta.SetFocus
        End If
    End If
End Sub

Private Sub txtcodcaja_LostFocus()
    If IsNumeric(txtcodcaja) Then
        txtcaja = ObtenerDescripcionCajas("Cajas", s2n(txtcodcaja))
        If txtcaja = "" Then
            MsgBox "Código de caja incorrecto"
            txtcodcaja.SetFocus
        Else
            cargar = "Cajas"
            CargarDatos
        End If
    Else
        If txtcodcaja <> "" Then
            MsgBox "Código de caja incorrecto"
            txtcodcaja.SetFocus
        End If
    End If
End Sub

Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Function SumoPagos() As Double
    SumoPagos = s2n(s2n(txtefectivo) + uCheques.Total + s2n(txttransf) + uRetCompras.TotalRet)
End Function
Private Function CheMePase() As Boolean
    CheMePase = SumoPagos() > s2n(txtimporte)
    If CheMePase Then che "Se supero el importe a pagar"
End Function
Private Function nuevoMonto(cuanto)
    cuanto = s2n(cuanto)
    nuevoMonto = cuanto + (s2n(txtimporte) - SumoPagos())
    txtFaltaPagar = Round((s2n(txtimporte) - SumoPagos()), 4)
End Function

'4/3/5 Sebastian me paso con mod
'    cambio frmhelp x frmbuscar
'    fix consultas, ahora busca tipo doc !!! era grave
' agregue:
'   fix fechas string
'   habilitacion forma pago, no habilitaba cdo correspondia, ahora TRUCHO siempre habilitado
'23/3/5 si no pasaba por caja quebraba. fix
'   importe, s2n en lostfoco
'   parametro sistema contable
'30/3/5
'   quito el horrendo frmCheques
'31/3/5
'   fix codigo ch: ponia o/P en chPropio y FDC en ch3ros
'1/6/5
'   fix numero OP funcion NuevoCodigoOP
'20/4/6 simplificacion, fix preguntas a textbox numericos, caja predet.

Private Sub uCheques_cambio()
    txtimpcheques = uCheques.Total
    
    Dim habi As Boolean

   
    habi = (uCheques.Total > 0) And (s2n(uCheques.Total + s2n(txtefectivo)) <> s2n(txtimporte))
'    txttransf.enabled = habi
'    txtcodcuenta.enabled = habi
'    cmbcuenta.enabled = habi
    
    'txtFaltaPagar = nuevoMonto(0)
    nuevoMonto 0
End Sub

Private Sub uRetCompras_cambio(Total As Double)
    'reca
    nuevoMonto 0
End Sub

Function ChequeaChq() As Boolean
    Dim x As Long
    Dim inter As Long
    Dim Nro As Long
    
    ChequeaChq = True
    If uCheques.Total > 0 Then
        With uCheques
            For x = 1 To .rows
                If .chPropio(x) Then
                    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                        If .chNroInt(x) = 0 Then
                            inter = s2n(obtenerDeSQL("select codigo from chq_comp where nro = " & .chNumero(x) & " and banco=" & .chBancCod(x)))
                            If inter > 0 Then
                                MsgBox "El cheque Nro." & .chNumero(x) & " existe con interno " & inter & ", por lo que debe seleccionarlo.", , "ATENCION"
                                ChequeaChq = False
                                'Exit Function
                            End If
                        End If
                    End If
                Else
                    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                        If .chNroInt(x) = 0 Then
                            inter = s2n(obtenerDeSQL("select nroint from cheques where nro = " & .chNumero(x) & " and banco_nro=" & .chBancCod(x)))
                            If inter > 0 Then
                                MsgBox "El cheque Nro." & .chNumero(x) & " existe con interno " & inter & ", por lo que debe seleccionarlo.", , "ATENCION"
                                ChequeaChq = False
                                'Exit Function
                            End If
                        End If
                    End If
                End If
                Nro = s2n(obtenerDeSQL("select codigo from bancosgrales where codigo=" & .chBancCod(x)))
                If Nro = 0 Then
                    MsgBox "El banco seleccionado para el cheque Nro." & .chNumero(x) & " no existe en la base, verifiquelo antes de continuar.", , "ATENCION"
                    ChequeaChq = False
                End If
            Next
        End With
    End If
End Function

