VERSION 5.00
Begin VB.Form FrmCajas 
   Caption         =   "Cajas"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   Icon            =   "FrmCajas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbmoneda 
      Height          =   315
      ItemData        =   "FrmCajas.frx":08CA
      Left            =   6120
      List            =   "FrmCajas.frx":08D4
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "15"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   4170
      Picture         =   "FrmCajas.frx":08F8
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Anterior"
      Top             =   1665
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   4845
      Picture         =   "FrmCajas.frx":0C02
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Siguiente"
      Top             =   1665
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   5460
      Picture         =   "FrmCajas.frx":0F0C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ultimo"
      Top             =   1665
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   3555
      Picture         =   "FrmCajas.frx":1216
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Primero"
      Top             =   1665
      Width           =   615
   End
   Begin VB.ComboBox cmbctacontable 
      Height          =   315
      ItemData        =   "FrmCajas.frx":1520
      Left            =   1680
      List            =   "FrmCajas.frx":152A
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "15"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "4"
      Top             =   360
      Width           =   4455
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
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
      TabIndex        =   9
      Top             =   2520
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   5250
      TabIndex        =   18
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1410
      Left            =   120
      Top             =   120
      Width           =   8640
   End
   Begin VB.Label Label1 
      Caption         =   "Cta. Contable:"
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
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Descripcion:"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   360
      Width           =   1215
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
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FrmCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit ' mod 11/8/4



Dim Ope As String
Dim rsCaja As New ADODB.Recordset

Private Sub cmdaceptar_Click()

Dim fecha As Variant
Dim rs As New ADODB.Recordset
Dim codigo As Long


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtdescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtdescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from cajas", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not IsNull(rs!cod) Then
                codigo = rs!cod + 1
            Else
                codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DataEnvironment1.dbo_CAJA "A", codigo, Trim(txtdescripcion), ObtenerCodigoCue("cuentas", Trim(cmbctacontable.Text)), ObtenerCodigo("monedas", Trim(cmbmoneda.Text)), fecha, UsuarioSistema!codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.dbo_CAJA "M", Val(TxtCodigo), Trim(txtdescripcion), Cuenta_CuentaDeDescripcion(cmbctacontable.Text), ObtenerCodigo("monedas", Trim(cmbmoneda.Text)), 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(TxtCodigo), "Cajas", UsuarioSistema!codigo, fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Public Function Cuenta_CuentaDeDescripcion(queDescripcion As String) As String
    Cuenta_CuentaDeDescripcion = sSinNull(obtenerDeSQL("select cuenta from cuentas where descripcion  = '" & queDescripcion & "' "))
End Function
Public Function Cuenta_DescripcionDeCuenta(queCuenta As String) As String
    Cuenta_DescripcionDeCuenta = sSinNull(obtenerDeSQL("select Descripcion from cuentas where cuenta = '" & queCuenta & "' "))
End Function

Sub CargoRegistro()
    TxtCodigo = rsCaja!codigo
    If Not IsNull(rsCaja!responsable) Then
        txtdescripcion = rsCaja!responsable
    Else
        txtdescripcion = ""
    End If
    If Not IsNull(rsCaja!CUENTA) Then
        cmbctacontable.Text = Cuenta_DescripcionDeCuenta(rsCaja!CUENTA)
    Else
        cmbctacontable.Text = ""
    End If
    If Not IsNull(rsCaja!moneda) Then
        cmbmoneda.Text = ObtenerDescripcion("monedas", rsCaja!moneda)
    Else
        cmbmoneda.Text = ""
    End If
End Sub
Private Sub cmdanterior_Click()
    rsCaja.MovePrevious
    If Not rsCaja.BOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "Cajas", "Codigo", "Descripcion", "codigo", "responsable"
'    FrmHelp.Tag = Me.Name
    Dim resu As String

    'resu = frmBuscar.MostrarCodigoDescripcionActivo("Cajas")
    resu = frmBuscar.MostrarSql("select Codigo, Cuenta, Moneda from cajas where activo = 1")
    If resu > "" Then
        TxtCodigo = resu
        txtdescripcion = frmBuscar.resultado(2)
        CargarDatos
        Call HabilitoControles(True, False, True, False, True, False)
    End If

End Sub

Private Sub cmdcancelar_Click()
    Call HabilitoControles(False, False, False, True, False, True)
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoBotonesMoverse(False, False, False, False)
End Sub

Private Sub cmdeliminar_Click()
Dim fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_CAJA "B", Trim(TxtCodigo), "", 0, 0, 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(TxtCodigo), "Cajas", UsuarioSistema!codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsCaja.MoveFirst
    CargoRegistro
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsCaja.MoveNext
    If Not rsCaja.EOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsCaja.MoveLast
    CargoRegistro
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsCaja.State = 1 Then
        rsCaja.Close
        Set rsCaja = Nothing
    End If
End Sub

Private Sub txtDescripcion_GotFocus()

    txtdescripcion.SelStart = 0
    txtdescripcion.SelLength = Len(txtdescripcion.Text)

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbctacontable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtdescripcion.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
Dim rsCaja As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    
    cmbmoneda.ListIndex = BuscarenComboS(cmbmoneda, Const_PESOS)
    
    rsCaja.Open "select max(codigo) as cod from cajas", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsCaja!cod) Then
        TxtCodigo = rsCaja!cod + 1
    Else
        TxtCodigo = 1
    End If
    rsCaja.Close
    Set rsCaja = Nothing
    HabilitoTxt (False)
    txtdescripcion.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    TxtCodigo = ""
    txtdescripcion = ""
    cmbctacontable.ListIndex = -1
    cmbmoneda.ListIndex = -1
'    Me.Refresh
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtdescripcion.Locked = habilito
    cmbctacontable.Locked = habilito
    cmbmoneda.Locked = habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    
    cmdcancelar.enabled = hab1
    cmdaceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdmodificar.enabled = hab5
    CmdBuscar.enabled = hab6
    
End Sub
Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    
    cmdprimero.enabled = hab2
    cmdanterior.enabled = hab1
    cmdsiguiente.enabled = hab3
    cmdultimo.enabled = hab4
    
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    CargaCombo cmbctacontable, "cuentas", "descripcion", "cuenta", "imputable=1"
    CargaCombo cmbmoneda, "monedas", "descripcion", "codigo", ""
    LimpioTxt
    Call HabilitoControles(False, False, False, True, False, True)
    HabilitoBotonesMoverse False, False, False, False
End Sub
Sub CargarDatos()
    
    If rsCaja.State = 1 Then
        rsCaja.Close
        Set rsCaja = Nothing
    End If
    rsCaja.Open "select * from cajas where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsCaja.EOF Then
        rsCaja.MoveFirst
        rsCaja.Find "Codigo= " & Trim(TxtCodigo)
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub



' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'   fix param dbo_caja

