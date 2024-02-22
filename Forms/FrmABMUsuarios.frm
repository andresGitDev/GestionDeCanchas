VERSION 5.00
Begin VB.Form FrmAbmUsuarios 
   Caption         =   "Usuarios"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   Icon            =   "FrmABMUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   6495.641
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmail 
      Height          =   285
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4920
      Width           =   2835
   End
   Begin VB.TextBox txtInicial 
      Height          =   285
      Left            =   2145
      TabIndex        =   34
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox txtClaveVieja 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2145
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2400
      Width           =   2910
   End
   Begin VB.TextBox txtClave2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2145
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2010
      Width           =   2910
   End
   Begin VB.CheckBox chkcambiar 
      Alignment       =   1  'Right Justify
      Caption         =   "Cambiar Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.TextBox txtporcentaje 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4440
      Width           =   2835
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3915
      Width           =   2850
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3435
      Width           =   2865
   End
   Begin VB.TextBox txtclave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2145
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1635
      Width           =   2910
   End
   Begin VB.TextBox txtusuario 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1230
      Width           =   2925
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2940
      Picture         =   "FrmABMUsuarios.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Anterior"
      Top             =   6360
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3615
      Picture         =   "FrmABMUsuarios.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Siguiente"
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   4230
      Picture         =   "FrmABMUsuarios.frx":2DB6
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ultimo"
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2325
      Picture         =   "FrmABMUsuarios.frx":30C0
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Primero"
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdsalir 
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
      Left            =   5969
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
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
      Left            =   4016
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
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
      Left            =   4993
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdmodificar 
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
      Left            =   2073
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
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
      Left            =   119
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
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
      Left            =   3039
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
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
      Left            =   1096
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox txtdescripcion 
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   825
      Width           =   2925
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2145
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   405
      Width           =   975
   End
   Begin VB.ComboBox CboTipoUsuario 
      Height          =   315
      Left            =   2145
      TabIndex        =   7
      Top             =   2940
      Width           =   2880
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   4935
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciales :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave Anterior :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   375
      TabIndex        =   33
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme Clave :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   2025
      Width           =   1380
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   390
      TabIndex        =   29
      Top             =   4455
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Telefono/s :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   375
      TabIndex        =   26
      Top             =   1635
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1275
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   6135
      Left            =   90
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   855
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   405
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Usuario: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   345
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "FrmAbmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ope As String
Dim rsusuario As New ADODB.Recordset

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAaceptar
    Dim rs As New ADODB.Recordset
    Dim codigo As Long
    Dim cambio As Long
    Dim PORCENTAJE As Double
    Dim resu

    If Trim(txtdescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtdescripcion.SetFocus
        Call HabilitoControles(True, True, False, False, False, False)
        Exit Sub
    Else
        If Trim(txtusuario) = "" Then
            MsgBox "Debe cargar el Usuario", 48, "Atencion"
            txtusuario.SetFocus
            Call HabilitoControles(True, True, False, False, False, False)
            Exit Sub
        Else
            If Trim(txtclave) = "" Then
                MsgBox "Debe cargar la clave", 48, "Atencion"
                txtclave.SetFocus
                Call HabilitoControles(True, True, False, False, False, False)
                Exit Sub
            Else
                If Ope = "A" Then
                    If txtclave <> txtClave2 Then
                        MsgBox "La clave que esta ingresando es diferente a la de confirmacion.", 48, "Atencion"
                        txtClave2.SetFocus
                        Call HabilitoControles(True, True, False, False, False, False)
                        Exit Sub
                    End If
                End If
    
                If Ope = "M" Then
                    Dim claveAnterior As String
                    claveAnterior = obtenerDeSQL("select clave from usuarios where codigo = " & txtCodigo)
                    If claveAnterior <> txtClaveVieja Then
                        MsgBox "La anterior clave no es correcta.", 48, "Atencion"
                        txtClaveVieja.SetFocus
                        Call HabilitoControles(True, True, False, False, False, False)
                        Exit Sub
                    Else
                        If txtclave <> txtClave2 Then
                            MsgBox "La clave que esta ingresando es diferente a la de confirmacion.", 48, "Atencion"
                            txtClave2.SetFocus
                            Call HabilitoControles(True, True, False, False, False, False)
                            Exit Sub
                        End If
                    End If
                End If

                If Ope = "A" Then
                    resu = obtenerDeSQL("select descripcion from usuarios where activo=1 and usuario='" & Trim(txtusuario.Text) & "' and clave='" & Trim(txtclave.Text) & "'")
                    If resu <> "" Then
                        MsgBox "ya existe este usuario y clave.", , "ATENCION"
                        Exit Sub
                    End If
                End If

                If Trim(txtporcentaje) = "" Then
                    PORCENTAJE = 0
                Else
                    PORCENTAJE = Trim(txtporcentaje)
                End If
                If Ope = "A" Then
                    Dim neww
                    neww = obtenerDeSQL("Select max(codigo) as cod from usuarios")
                    If IsNull(neww) Or IsEmpty(neww) Then
                        codigo = 1
                    Else
                        codigo = neww + 1
                    End If
                    
                    DataEnvironment1.dbo_USUARIO "A", codigo, Trim(txtdescripcion), Trim(txtusuario), Trim(txtclave), 0, ObtenerCodigo("TipoUsuarios", Trim(CboTipoUsuario.Text)), Trim(txtdireccion), Trim(txttel), Trim(txtlocalidad), CDbl(PORCENTAJE), Date, UsuarioActual, 0, Date
                    DataEnvironment1.Sistema.Execute "update usuarios set inicial=" & ssTexto(Trim(txtInicial)) & ",mail=" & ssTexto(Trim(txtmail)) & " where codigo=" & codigo
                    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                Else
                    
                    If Ope = "M" Then
                        If chkcambiar.Value = 1 Then
                            cambio = 1
                        Else
                            cambio = 0
                        End If
                        DataEnvironment1.dbo_USUARIO "M", Val(txtCodigo), Trim(txtdescripcion), Trim(txtusuario), Trim(txtclave), cambio, ObtenerCodigo("Tipousuarios", Trim(CboTipoUsuario.Text)), Trim(txtdireccion), Trim(txttel), Trim(txtlocalidad), CDbl(PORCENTAJE), 0, 0, 0, 0
                        DataEnvironment1.Sistema.Execute "update usuarios set inicial=" & ssTexto(Trim(txtInicial)) & ",mail=" & ssTexto(Trim(txtmail)) & " where codigo=" & txtCodigo
                        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Usuarios", UsuarioSistema!codigo, Date, Time, "M"
                        MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                    End If
                
                End If
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoControles(False, False, False, True, False, True)

fin:
    Set rs = Nothing
    Exit Sub
UFAaceptar:
    ufa "Err:", Me.Name ', Err
    Resume fin
End Sub
Sub CargoRegistro()
    txtCodigo = rsusuario!codigo
    txtdescripcion = rsusuario!DESCRIPCION
    txtusuario = rsusuario!usuario
    txtclave = rsusuario!Clave
    If Not IsNull(rsusuario!direccion) Then
        txtdireccion = rsusuario!direccion
    Else
        txtdireccion = ""
    End If
    If Not IsNull(rsusuario!PORCENTAJE) Then
        txtporcentaje = rsusuario!PORCENTAJE
    Else
        txtporcentaje = ""
    End If
    If Not IsNull(rsusuario!Telefono) Then
        txttel = rsusuario!Telefono
    Else
        txttel = ""
    End If
    If Not IsNull(rsusuario!Localidad) Then
        txtlocalidad = rsusuario!Localidad
    Else
        txtlocalidad = ""
    End If
    txtInicial.Text = sSinNull(rsusuario!inicial)
    
'''''  FALTA TABLA
    If Not IsNull(rsusuario!TIPOUSUARIO) Then
        CboTipoUsuario.Text = ObtenerDescripcion("TipoUsuarios", rsusuario!TIPOUSUARIO)
    Else
        CboTipoUsuario.ListIndex = -1
    End If
    
    txtmail.Text = sSinNull(rsusuario!mail)
End Sub
Private Sub cmdanterior_Click()
    rsusuario.MovePrevious
    If Not rsusuario.BOF Then
        CargoRegistro
        
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String

    resu = frmBuscar.MostrarCodigoDescripcionActivo("Usuarios")
    
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
    
    Label5.caption = "Clave :"
    Label11.Visible = False
    txtClaveVieja.Visible = False
End Sub

Private Sub cmdeliminar_Click()
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        DataEnvironment1.dbo_USUARIO "B", Trim(txtCodigo), "", "", "", 0, 0, "", "", "", 0, 0, 0, UsuarioSistema!codigo, Date
        DataEnvironment1.dbo_GRABARBITACORA Val(Trim(txtCodigo)), "Usuarios", UsuarioSistema!codigo, Date, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsusuario.MoveFirst
    CargoRegistro
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsusuario.MoveNext
    If Not rsusuario.EOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsusuario.MoveLast
    CargoRegistro
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rsusuario.State = 1 Then
        rsusuario.Close
        Set rsusuario = Nothing
    End If
End Sub

Private Sub txtClave2_GotFocus()
    txtClave2.SelStart = 0
    txtClave2.SelLength = Len(txtclave.Text)
End Sub

Private Sub txtClave2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    txtdescripcion.SelStart = 0
    txtdescripcion.SelLength = Len(txtdescripcion.Text)
End Sub
Private Sub txtusuario_GotFocus()
    txtusuario.SelStart = 0
    txtusuario.SelLength = Len(txtusuario.Text)
End Sub
Private Sub txtclave_GotFocus()
    txtclave.SelStart = 0
    txtclave.SelLength = Len(txtclave.Text)
End Sub
Private Sub txtDireccion_GotFocus()
    txtdireccion.SelStart = 0
    txtdireccion.SelLength = Len(txtdireccion.Text)
End Sub
Private Sub txttel_GotFocus()
    txttel.SelStart = 0
    txttel.SelLength = Len(txttel.Text)
End Sub
Private Sub txtlocalidad_GotFocus()
    txtlocalidad.SelStart = 0
    txtlocalidad.SelLength = Len(txtlocalidad.Text)
End Sub
Private Sub txtporcentaje_GotFocus()
    txtporcentaje.SelStart = 0
    txtporcentaje.SelLength = Len(txtporcentaje.Text)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txtclave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txttel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txtlocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub txtporcentaje_KeyPress(KeyAscii As Integer)
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
    
    txtclave.Text = ""
    Label5.caption = "Nueva Clave :"
    Label11.Visible = True
    txtClaveVieja.Visible = True
End Sub

Private Sub cmdnuevo_Click()
Dim rsusuario As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsusuario.Open "select max(codigo) as cod from Usuarios", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsusuario!COD) Then
        txtCodigo = rsusuario!COD + 1
    Else
        txtCodigo = 1
    End If
    rsusuario.Close
    Set rsusuario = Nothing
    HabilitoTxt (False)
    txtdescripcion.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
    Label11.Visible = False
    txtClaveVieja.Visible = False
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtCodigo = ""
    txtdescripcion = ""
    txtusuario = ""
    txtclave = ""
    txtClave2 = ""
    txtClaveVieja = ""
    txtdireccion = ""
    txttel = ""
    txtlocalidad = ""
    txtporcentaje = ""
    CboTipoUsuario.ListIndex = -1
    chkcambiar.Value = 0
    txtmail.Text = ""
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtdescripcion.Locked = habilito
    txtusuario.Locked = habilito
    txtclave.Locked = habilito
    txtClave2.Locked = habilito
    txtClaveVieja.Locked = habilito
    txtdireccion.Locked = habilito
    txttel.Locked = habilito
    txtlocalidad.Locked = habilito
    txtporcentaje.Locked = habilito
    CboTipoUsuario.Locked = habilito
    chkcambiar.enabled = Not habilito
    txtmail.Locked = habilito
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
    
'''' Lito '  25-10-4  Falta TABLA !!!???
    CargaCombo CboTipoUsuario, "TipoUsuarios", "descripcion", "codigo", ""

    Call HabilitoControles(False, False, False, True, False, True)
    HabilitoBotonesMoverse False, False, False, False
    Label11.Visible = False
    txtClaveVieja.Visible = False
End Sub
Sub CargarDatos()
    
    If rsusuario.State = 1 Then
        rsusuario.Close
        Set rsusuario = Nothing
    End If
    rsusuario.Open "select * from Usuarios where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsusuario.EOF Then
        rsusuario.MoveFirst
        rsusuario.Find "Codigo= " & str(Trim(txtCodigo))
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub

Private Sub txtusuario_Change()
Dim i As Long
    txtusuario.Text = UCase(txtusuario.Text)
    i = Len(txtusuario.Text)
    txtusuario.SelStart = i
End Sub
