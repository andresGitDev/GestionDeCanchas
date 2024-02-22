VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmConsignatarios 
   Caption         =   "Consignatarios"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   Icon            =   "FrmConsignatarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   ScaleHeight     =   3375
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtdesc 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   4935
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
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2955
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
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2955
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
      TabIndex        =   10
      Top             =   2955
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
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2955
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2955
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2955
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2955
      Width           =   975
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2280
      Picture         =   "FrmConsignatarios.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Primero"
      Top             =   2115
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   4185
      Picture         =   "FrmConsignatarios.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ultimo"
      Top             =   2115
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3570
      Picture         =   "FrmConsignatarios.frx":230E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Siguiente"
      Top             =   2115
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2895
      Picture         =   "FrmConsignatarios.frx":2618
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Anterior"
      Top             =   2115
      Width           =   675
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1155
      Width           =   4935
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1635
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Format          =   68812801
      CurrentDate     =   38052
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   240
      Width           =   855
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
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre :"
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
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label6 
      Caption         =   "Direccion :"
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
      Left            =   360
      TabIndex        =   16
      Top             =   1155
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Telefono/s :"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1635
      Width           =   975
   End
End
Attribute VB_Name = "FrmConsignatarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'11/8/4 falta sp
' FALTA TODO, INCOMPLETO


Option Explicit ' mod 11/8/4


Dim rscons As New ADODB.Recordset
Dim Ope As String
Dim numero As Long

'Private Sub cmdAceptar_Click()
'
'Dim fecha As Variant
'
'If Ope <> "" Then
'    fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'    If Ope = "A" Then
'        daTaenvironment1.dbo_CONSIGNACIONES "A", fecha, UsuarioSistema!Codigo, 0, 0
'    Else
'        If Ope = "M" Then
'            daTaenvironment1.dbo_CONSIGNACIONES "M", 0, 0, 0, 0
'
'            daTaenvironment1.dbo_GRABARBITACORA Val(Trim(txtCodigo)), "Usuarios", UsuarioSistema!Codigo, fecha, Time, "M"
'        End If
'    End If
'    MsgBox "La operación fue realizada con éxito"
'    LimpioControles
'    Call Habilitobotones(True, True, True, True, True, True)
'Else
'    MsgBox "Operación no válida"
'End If
'
'End Sub

Private Sub cmdBuscar_Click()
'    FrmHelp.Show
'    CargarHelp "Grupos", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    Call Habilitobotones(True, False, True, True, True, True)
    Dim resu As String

    resu = frmBuscar.MostrarCodigoDescripcionActivo("Grupos")
    If resu > "" Then
        txtCodigo = resu
        CargarDatos
        Call Habilitobotones(True, False, True, True, True, True)
    End If

End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Call HabilitoControles(False)
    Call Habilitobotones(True, True, False, False, False, True)
End Sub
Public Sub CargarDatos()
    Dim codigo
    
    If rscons.State = 1 Then
        rscons.Close
        Set rscons = Nothing
    End If

    codigo = Val(Trim(Me.Tag))
    rscons.Open "select * from Grupos where activo = true and codigo = " & codigo & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rscons.EOF Then
        CargoProducto
    End If
    rscons.Close
    Set rscons = Nothing

End Sub

Private Sub cmdeliminar_Click()
Dim fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_USUARIO "B", Trim(txtCodigo), "", "", "", 0, 0, "", "", "", 0, 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Val(Trim(txtCodigo)), "Grupos", UsuarioSistema!codigo, fecha, Time, "B"
        Call Habilitobotones(True, True, False, False, False, False)
        LimpioControles
        Call HabilitoControles(True)
    End If

End Sub

Private Sub cmdmodificar_Click()
    Ope = "M"
    Call HabilitoControles(True)
    Call Habilitobotones(True, False, False, True, True, True)
End Sub

Private Sub cmdnuevo_Click()
Dim rs As New ADODB.Recordset

    rs.Open "select max(codigo) + 1 as maxcodigo from Grupos", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        txtCodigo = rs!maxcodigo
        numero = rs!maxcodigo
    End If
    rs.Close
    Set rs = Nothing
    
    Call HabilitoControles(True)
    Call Habilitobotones(False, False, False, False, True, True)
    Ope = "A"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    HabilitoBotonesMoverse False, False, False, False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub txtcodigo_LostFocus()
    If txtCodigo <> "" Then
        If Val(txtCodigo) < numero Then
            MsgBox "El código no puede ser menor al último ingresado"
            txtCodigo.SetFocus
        End If
    Else
        MsgBox "Debe ingresar un código"
        txtCodigo.SetFocus
    End If
End Sub

Sub LimpioControles()
    txtCodigo = ""
    txtdesc = ""
    dtFecha = Date
    txtDireccion = ""
    txtTel = ""
    
    Ope = ""
End Sub

Sub CargoProducto()
'    txtCodigo = rscons!
'    txtdesc = rscons!
'    dtfecha = rscons!
'    txtdireccion = rscons!
'    txttel = rscons!

End Sub

Sub HabilitoControles(habilito As Boolean)
'    txtCodigo.Enabled = habilito
'    txtdesc.Enabled = habilito
'    dtfecha.Enabled = habilito
'    txtabrev.Enabled = habilito
'    txtdireccion = habilito
'    txttel = habilito
    
End Sub

Sub Habilitobotones(busco As Boolean, Nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdbuscar.enabled = busco
    cmdnuevo.enabled = Nuevo
    cmdmodificar.enabled = modifico
    cmdeliminar.enabled = elimino
    cmdaceptar.enabled = acepto
    cmdcancelar.enabled = Cancelo
End Sub

Private Sub txtdesc_LostFocus()
    If txtdesc = "" Then
        MsgBox "Debe ingresar una descripción"
    End If
End Sub

Private Sub cmdPrimero_Click()
    rscons.MoveFirst
    txtCodigo = rscons!codigo
    txtdesc = rscons!DESCRIPCION
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rscons.MoveNext
    If Not rscons.EOF Then
        txtCodigo = rscons!codigo
        txtdesc = rscons!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rscons.MoveLast
    txtCodigo = rscons!codigo
    txtdesc = rscons!DESCRIPCION
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub cmdanterior_Click()
    rscons.MovePrevious
    If Not rscons.BOF Then
        txtCodigo = rscons!codigo
        txtdesc = rscons!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rscons.State = 1 Then
        rscons.Close
        Set rscons = Nothing
    End If
End Sub

Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    
    cmdprimero.enabled = hab2
    cmdanterior.enabled = hab1
    cmdsiguiente.enabled = hab3
    cmdultimo.enabled = hab4
    
End Sub

' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR

' fix rs param mal escrito, no compilaba. bd x db
' fix habilitobotones en boton eliminar
Private Sub txttel_Change()

End Sub

Private Sub txttel_GotFocus()
    txtTel.SelStart = 0
    txtTel.SelLength = Len(txtTel.Text)
End Sub
