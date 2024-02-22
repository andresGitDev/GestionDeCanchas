VERSION 5.00
Begin VB.Form FrmAbmTiposUsuarios 
   Caption         =   "Abm de tipos de usuarios"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "FrmAbmTipoUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   2820
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
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
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   975
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
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
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
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
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
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   1965
      Picture         =   "FrmAbmTipoUsuarios.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   3870
      Picture         =   "FrmAbmTipoUsuarios.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ultimo"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3255
      Picture         =   "FrmAbmTipoUsuarios.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Siguiente"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2580
      Picture         =   "FrmAbmTipoUsuarios.frx":30C0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Anterior"
      Top             =   1560
      Width           =   675
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   855
      Width           =   4680
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion :"
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
      Left            =   375
      TabIndex        =   3
      Top             =   855
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo :"
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
      Height          =   375
      Left            =   405
      TabIndex        =   2
      Top             =   375
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1290
      Left            =   120
      Top             =   135
      Width           =   6720
   End
End
Attribute VB_Name = "FrmAbmTiposUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ope As String
Dim rsTipo As New ADODB.Recordset

Private Sub cmdAceptar_Click()

    Dim rs As New ADODB.Recordset
    Dim codigo As Long

    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from TipoUsuarios", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            codigo = nSinNull(rs!COD) + 1
            rs.Close
            Set rs = Nothing
            DataEnvironment1.dbo_TIPOUSUARIO "A", codigo, Trim(txtDescripcion), Date, UsuarioSistema!codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                DataEnvironment1.dbo_TIPOUSUARIO "M", Val(txtCodigo), Trim(txtDescripcion), 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FomasPago", UsuarioSistema!codigo, Date, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rsTipo.MovePrevious
    If Not rsTipo.BOF Then
        txtCodigo = rsTipo!codigo
        txtDescripcion = rsTipo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String

    'resu = frmBuscar.MostrarCodigoDescripcionActivo("TipoUsuarios")
    resu = frmBuscar.MostrarSql("select codigo as [ Codigo           ], descripcion as [ Descripcion                                ] from TipoUsuarios where activo=1")
    If resu > "" Then
        txtCodigo = resu
        CargarDatos
        Call HabilitoControles(True, False, True, False, True, False)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitoControles(False, False, False, True, False, True)
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoBotonesMoverse(False, False, False, False)
End Sub

Private Sub cmdeliminar_Click()
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        DataEnvironment1.dbo_TIPOUSUARIO "B", Trim(txtCodigo), "", 0, 0, UsuarioSistema!codigo, Date
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FomasPago", UsuarioSistema!codigo, Date, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsTipo.MoveFirst
    txtCodigo = rsTipo!codigo
    txtDescripcion = rsTipo!DESCRIPCION
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsTipo.MoveNext
    If Not rsTipo.EOF Then
        txtCodigo = rsTipo!codigo
        txtDescripcion = rsTipo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsTipo.MoveLast
    txtCodigo = rsTipo!codigo
    txtDescripcion = rsTipo!DESCRIPCION
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rsTipo.State = 1 Then
        rsTipo.Close
        Set rsTipo = Nothing
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = Len(txtDescripcion.Text)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtDescripcion.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
Dim rsTipoUsuario As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsTipoUsuario.Open "select max(codigo) as cod from TipoUsuarios", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsTipoUsuario!COD) Then
        txtCodigo = rsTipoUsuario!COD + 1
    Else
        txtCodigo = 1
    End If
    rsTipoUsuario.Close
    Set rsTipoUsuario = Nothing
    HabilitoTxt (False)
    txtDescripcion.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtCodigo = ""
    txtDescripcion = ""
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    cmdCancelar.enabled = hab1
    cmdAceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdmodificar.enabled = hab5
    cmdbuscar.enabled = hab6
End Sub
Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    cmdprimero.enabled = hab2
    cmdanterior.enabled = hab1
    cmdsiguiente.enabled = hab3
    cmdultimo.enabled = hab4
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioTxt
    Call HabilitoControles(False, False, False, True, False, True)
    HabilitoBotonesMoverse False, False, False, False
End Sub
Sub CargarDatos()
    
    If rsTipo.State = 1 Then
        rsTipo.Close
        Set rsTipo = Nothing
    End If
    rsTipo.Open "select * from TipoUsuarios where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsTipo.EOF Then
        rsTipo.MoveFirst
        rsTipo.Find "Codigo= " & str(Trim(txtCodigo))
        txtDescripcion = rsTipo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub
