VERSION 5.00
Begin VB.Form FrmCtasBancarias 
   Caption         =   "Ctas. Bancarias"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   Icon            =   "FrmCtasBancarias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      TabIndex        =   33
      Tag             =   "4"
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtpordebitar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   31
      Tag             =   "4"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtdepositado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   29
      Tag             =   "4"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtsaldo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ComboBox cmbmoneda 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "7"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ComboBox cmbsobrepesos 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmCtasBancarias.frx":08CA
      Left            =   9240
      List            =   "FrmCtasBancarias.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "5"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txttipo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmbtipcuenta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cuenta"
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
      TabIndex        =   25
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtdescuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   24
      Tag             =   "2"
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtdesbanco 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   23
      Tag             =   "2"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmbbanco 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Banco"
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
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtcodbanco 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtsobregiro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Tag             =   "4"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtnumero 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "3"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtcodigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
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
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "8"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   975
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "0"
      Top             =   3600
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin Gestion.ucCoDe uCuenta 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   6375
      _extentx        =   11245
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3450
      Left            =   120
      Top             =   120
      Width           =   9960
   End
   Begin VB.Label Label7 
      Caption         =   "Por debitar:"
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
      Left            =   7200
      TabIndex        =   32
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Depositado:"
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
      Left            =   3960
      TabIndex        =   30
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo:"
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
      Left            =   360
      TabIndex        =   28
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Cuenta:"
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
      Left            =   360
      TabIndex        =   26
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Cta. Contable"
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
      Left            =   360
      TabIndex        =   21
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Sobrepesos:"
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
      Left            =   7920
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Sobregiro:"
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
      Left            =   4560
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Código:"
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
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Número:"
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
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Banco:"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
End
Attribute VB_Name = "FrmCtasBancarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Dim rsctas As New ADODB.Recordset
Dim Ope As String
Dim numero As Long

Private Sub cmbbanco_Click()
    FrmHelp.Show
    Call CargarHelp("BancosGrales", "Codigo", "Descripcion", "Codigo", "Descripcion", "Codigo")
    FrmHelp.Tag = Me.Name
    cargar = "Bancos"
End Sub

Private Sub cmbcuenta_Click()
    FrmHelp.Show
    CargarHelp "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
    FrmHelp.Tag = Me.Name
    cargar = "Cuentas"
    cargar = "Cuentas"
End Sub



Private Sub cmbmoneda_GotFocus()
'    cmbmoneda.SelStart = 0
'    cmbmoneda.SelLength = Len(cmbmoneda.Text)
End Sub

Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmbsobrepesos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmbtipcuenta_Click()
    FrmHelp.Show
    Call CargarHelp("TipoCtas", "Codigo", "Descripcion", "Codigo", "Descripcion", "Codigo")
    FrmHelp.Tag = Me.Name
    cargar = "Tipocuentas"
End Sub

Private Sub cmdAceptar_Click()
    Dim Fecha

If txtTipo = "" Then
    MsgBox "Debe ingresar el tipo de cuenta"
    Exit Sub
End If

If txtcodbanco = "" Then
    MsgBox "Debe ingresar el código de banco"
    Exit Sub
End If

If txtNumero = "" Then
    MsgBox "Debe ingresar el número de la cuenta"
    Exit Sub
End If

'If txtCodCuenta = "" Then
'    MsgBox "Debe ingresar el código de la cuenta"
'    Exit Sub
'End If
    If uCuenta.DESCRIPCION = "" Then
        che "Debe ingresar el código de la cuenta"
        Exit Sub
    End If


If Ope <> "" Then
    Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
    If Ope = "A" Then
        DataEnvironment1.dbo_CUENTASBANK "A", val(txtCodigo), val(txtcodbanco), Trim(txtTipo), _
        Trim(txtNumero), s2n(txtsobregiro), IIf(cmbsobrepesos.Text = "SI", 1, 0), uCuenta.codigo, ObtenerCodigo("Monedas", cmbMoneda), Fecha, UsuarioSistema!codigo, 0, 0
    Else
        If Ope = "M" Then
            DataEnvironment1.dbo_CUENTASBANK "M", val(txtCodigo), val(txtcodbanco), Trim(txtTipo), _
            Trim(txtNumero), s2n(txtsobregiro), IIf(cmbsobrepesos = "SI", 1, 0), uCuenta.codigo, ObtenerCodigo("Monedas", cmbMoneda), 0, 0, 0, 0
        End If
    End If
    MsgBox "La operación fue realizada con éxito"
    LimpioControles
    Call Habilitobotones(True, True, True, True, True, True)
    Call HabilitoControles(False)
Else
    MsgBox "Operación no válida"
End If

End Sub

Private Sub cmdBuscar_Click()
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
    cargar = "CuentasBank"
    Call Habilitobotones(True, False, True, True, False, True)
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Call HabilitoControles(False)
    Call Habilitobotones(True, True, False, False, False, False)
End Sub
Public Sub CargarDatos()
Dim rs As New ADODB.Recordset, codigo

    codigo = val(Trim(Me.Tag))
    
    If cargar = "Tipocuentas" Then
        rs.Open "select * from TipoCtas where codigo = " & val(txtTipo) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtTipo = rs!codigo
            txtdescuenta = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    
    If cargar = "Bancos" Then
        rs.Open "select * from BancosGrales where codigo = " & val(txtcodbanco) & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            txtcodbanco = rs!codigo
            txtdesbanco = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "CuentasBank" Then
        rsctas.Open "select * from Ctasbank where activo = 1 and codigo = " & codigo & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rsctas.EOF Then
            CargoCtasBancarias
        End If
        rsctas.Close
        Set rsctas = Nothing
    End If
End Sub


Private Sub cmdeliminar_Click()
Dim a As String, Fecha

    a = MsgBox("Esta seguro de eliminar la Cta. Bancaria ?", vbYesNo)
    If a = vbYes Then
        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_CUENTASBANK "B", val(txtCodigo), 0, "", "", 0, 0, "", 0, 0, 0, UsuarioSistema!codigo, Fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Ivas", UsuarioSistema!codigo, Fecha, Time, "B"
        MsgBox "La Cta. Bancaria ha sido eliminada"
        LimpioControles
        Call HabilitoControles(False)
        Call Habilitobotones(True, True, False, False, False, False)
    End If
End Sub

Private Sub cmdmodificar_Click()
    Ope = "M"
    Call HabilitoControles(True)
    Call Habilitobotones(False, False, False, True, True, True)
End Sub

Private Sub cmdnuevo_Click()
Dim rs As New ADODB.Recordset

    rs.Open "Select max(codigo) as maxcod from Ctasbank", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rs!maxcod) Then
        If Not IsNull(rs!maxcod) Then
            txtCodigo = rs!maxcod + 1
            numero = rs!maxcod + 1
        Else
            txtCodigo = 1
            numero = 1
        End If
    Else
        txtCodigo = 1
        numero = 1
    End If
    rs.Close
    Set rs = Nothing
    
    cargocombos
    
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    
    Call HabilitoControles(True)
    Call Habilitobotones(False, False, False, False, True, True)
    Ope = "A"
End Sub
Sub cargocombos()
    If cmbsobrepesos.ListCount >= 1 Then
        cmbsobrepesos.ListIndex = 0
    End If
    
    If cmbMoneda.ListCount >= 1 Then
        cmbMoneda.ListIndex = 0
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'    MsgBox "FALTA HACER LA FCION PARA CALCULAR EL SALDO QUE SUPUESTAMENTE ES LA SUMA DE TODOS LOS CREDITOS - LA SUMA DE LOS DEBITOS, HAY QUE CHEQUEAR QUE LETRAS INCLUYEN C/U"
'    MsgBox "ADEMAS HABLAR CON LORENA PARA VER COMOSACAR EL CAMPO DEPOSITADO Y EL POR DEBITAR"
     CargaCombo cmbMoneda, "Monedas", "descripcion", "codigo", "activo = 1"
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub


Private Sub txtcodbanco_GotFocus()
    txtcodbanco.SelStart = 0
    txtcodbanco.SelLength = Len(txtcodbanco.Text)
End Sub

Private Sub txtcodbanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodbanco_LostFocus()
    If IsNumeric(txtcodbanco) Then
        txtdesbanco = ObtenerDescripcion("BancosGrales", val(txtcodbanco))
        If txtdesbanco = "" Then
            MsgBox "Código de banco incorrecto"
            txtcodbanco = "0"
            txtcodbanco.SetFocus
        End If
    Else
        If txtcodbanco <> "" Then
            MsgBox "Código de banco incorrecto"
            txtcodbanco = "0"
            txtcodbanco.SetFocus
        End If
    End If
End Sub

'Private Sub txtcodcuenta_GotFocus()
'    txtCodCuenta.SelStart = 0
'    txtCodCuenta.SelLength = Len(txtCodCuenta.Text)
'End Sub

'Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub

'Private Sub txtcodcuenta_LostFocus()
'    If IsNumeric(txtCodCuenta) Then
'        txtCuenta = ObtenerDescripcion("Cuentas", Val(txtCodCuenta))
'        If txtCuenta = "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtCodCuenta = "0"
'            txtCodCuenta.SetFocus
'        End If
'    Else
'        If txtCodCuenta <> "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtCodCuenta = "0"
'            txtCodCuenta.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtcodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtcodigo_LostFocus()
    If txtCodigo <> "" Then
        If val(txtCodigo) < numero Then
            MsgBox "El código no puede ser menor al último ingresado"
        End If
    Else
        MsgBox "Debe ingresar un código"
    End If
End Sub

Sub LimpioControles()
     
    txtCodigo = ""
    txtTipo = ""
    txtdescuenta = ""
    txtcodbanco = ""
    txtdesbanco = ""
    txtNumero = ""
    txtsobregiro = ""
    'txtCodCuenta = ""
    'txtCuenta = ""
    cmbsobrepesos.ListIndex = 0
    txtsaldo = ""
    txtdepositado = ""
    txtpordebitar = ""
    cmbMoneda.ListIndex = 0
    
    uCuenta.clear
    Ope = ""
End Sub

Sub CargoCtasBancarias()
     
    txtCodigo = rsctas!codigo
    txtTipo = rsctas!Tipo
    txtdescuenta = ObtenerDescripcion("TipoCtas", rsctas!Tipo)
    txtcodbanco = rsctas!Banco
    txtdesbanco = ObtenerDescripcion("BancosGrales", rsctas!Banco)
    txtNumero = rsctas!numero
    txtsobregiro = rsctas!sobregiro
    
    'txtCodCuenta = rsctas!cuenta_con
    uCuenta.codigo = rsctas!cuenta_con
    'txtCuenta = ObtenerDescripcion("Cuentas", rsctas!cuenta_con)
    
    'cmbsobrepesos.ListIndex = BuscarenCombo(cmbsobrepesos, IIf(rsctas!sobrepesos = True, "SI", "NO"))
    cmbsobrepesos.ListIndex = IIf(rsctas!sobrepesos = True, 1, 0)
    'cmbMoneda.ListIndex = BuscarEnCombo(cmbMoneda, ObtenerDescripcion("Monedas", rsctas!moneda))
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, ObtenerDescripcion("Monedas", rsctas!moneda))
'    txtsaldo = rscta!saldo
'    txtdepositado = rscta!depositado
'    txtpordebitar = rscta!xdebitar

End Sub

Sub HabilitoControles(habilito As Boolean)
    
    txtTipo.enabled = habilito
    txtdescuenta.enabled = habilito
    txtcodbanco.enabled = habilito
    txtdesbanco.enabled = habilito
    txtNumero.enabled = habilito
    txtsobregiro.enabled = habilito
'    txtCodCuenta.Enabled = habilito
    uCuenta.enabled = habilito
'    txtCuenta.Enabled = habilito
    cmbsobrepesos.enabled = habilito
    cmbbanco.enabled = habilito
'    cmbcuenta.Enabled = habilito
    cmbtipcuenta.enabled = habilito
    cmbMoneda.enabled = habilito
    
End Sub

Sub Habilitobotones(busco As Boolean, nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdbuscar.enabled = busco
    cmdnuevo.enabled = nuevo
    cmdmodificar.enabled = modifico
    cmdeliminar.enabled = elimino
    cmdAceptar.enabled = acepto
    cmdcancelar.enabled = Cancelo
End Sub


Private Sub txtdepositado_GotFocus()
    txtdepositado.SelStart = 0
    txtdepositado.SelLength = Len(txtdepositado.Text)

End Sub

Private Sub txtNumero_GotFocus()
    txtNumero.SelStart = 0
    txtNumero.SelLength = Len(txtNumero.Text)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtpordebitar_GotFocus()
    txtpordebitar.SelStart = 0
    txtpordebitar.SelLength = Len(txtpordebitar.Text)
End Sub

Private Sub txtsaldo_GotFocus()
    txtsaldo.SelStart = 0
    txtsaldo.SelLength = Len(txtsaldo.Text)
End Sub

Private Sub txtsobregiro_GotFocus()
    txtsobregiro.SelStart = 0
    txtsobregiro.SelLength = Len(txtsobregiro.Text)
End Sub

Private Sub txtsobregiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txttipo_GotFocus()
    txtTipo.SelStart = 0
    txtTipo.SelLength = Len(txtTipo.Text)
End Sub

Private Sub txttipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txttipo_LostFocus()
    If IsNumeric(txtTipo) Then
        txtdescuenta = ObtenerDescripcion("TipoCtas", val(txtTipo))
        If txtdescuenta = "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtTipo = "0"
            txtTipo.SetFocus
        End If
    Else
        If txtTipo <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtTipo = "0"
            txtTipo.SetFocus
        End If
    End If
End Sub


