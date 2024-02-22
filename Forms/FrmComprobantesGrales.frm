VERSION 5.00
Begin VB.Form FrmComprobantesGrales 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CARGA DE COMPROBANTES"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   Icon            =   "FrmComprobantesGrales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
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
      Top             =   2385
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
      Top             =   2385
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
      Top             =   2385
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
      Top             =   2385
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
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2385
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
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2385
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
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2385
      Width           =   975
   End
   Begin VB.CommandButton cmdprimero 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2040
      Picture         =   "FrmComprobantesGrales.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Primero"
      Top             =   1545
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3945
      Picture         =   "FrmComprobantesGrales.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ultimo"
      Top             =   1545
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   3330
      Picture         =   "FrmComprobantesGrales.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Siguiente"
      Top             =   1545
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2655
      Picture         =   "FrmComprobantesGrales.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Anterior"
      Top             =   1545
      Width           =   675
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   840
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
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   345
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   450
      TabIndex        =   14
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1290
      Left            =   165
      Top             =   120
      Width           =   6720
   End
End
Attribute VB_Name = "FrmComprobantesGrales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4


Dim Ope As String
Dim rscomp As New ADODB.Recordset

Private Sub cmdAceptar_Click()

Dim fecha As Variant
Dim rs As New ADODB.Recordset
Dim Codigo As Long


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtdescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtdescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from TipoComprobantesGrales", DaTaEnvironment1.AMR, adOpenStatic, adLockReadOnly
            If Not IsNull(rs!cod) Then
                Codigo = rs!cod + 1
            Else
                Codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DaTaEnvironment1.dbo_COMPROBANTEGRAL "A", Codigo, Trim(txtdescripcion), fecha, UsuarioSistema!Codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DaTaEnvironment1.dbo_COMPROBANTEGRAL "M", Val(txtcodigo), Trim(txtdescripcion), 0, 0, 0, 0
                DaTaEnvironment1.dbo_GRABARBITACORA Trim(txtcodigo), "TipoComprobantesGrales", UsuarioSistema!Codigo, fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rscomp.MovePrevious
    If Not rscomp.BOF Then
        txtcodigo = rscomp!Codigo
        txtdescripcion = rscomp!descripcion
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "TipoComprobantesGrales", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
    
    Dim resu As String

    resu = frmBuscar.MostrarCodigoDescripcionActivo("TipoComprobantesGrales")
    If resu > "" Then
        txtcodigo = resu
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
Dim fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DaTaEnvironment1.dbo_COMPROBANTEGRAL "B", Trim(txtcodigo), "", 0, 0, UsuarioSistema!Codigo, fecha
        DaTaEnvironment1.dbo_GRABARBITACORA Trim(txtcodigo), "TipoComprobantesGrales", UsuarioSistema!Codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rscomp.MoveFirst
    txtcodigo = rscomp!Codigo
    txtdescripcion = rscomp!descripcion
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rscomp.MoveNext
    If Not rscomp.EOF Then
        txtcodigo = rscomp!Codigo
        txtdescripcion = rscomp!descripcion
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rscomp.MoveLast
    txtcodigo = rscomp!Codigo
    txtdescripcion = rscomp!descripcion
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rscomp.State = 1 Then
        rscomp.Close
        Set rscomp = Nothing
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

Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtdescripcion.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
Dim rscomp As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rscomp.Open "select max(codigo) as cod from TipoComprobantesGrales", DaTaEnvironment1.AMR, adOpenStatic, adLockReadOnly
    If Not IsNull(rscomp!cod) Then
        txtcodigo = rscomp!cod + 1
    Else
        txtcodigo = 1
    End If
    rscomp.Close
    Set rscomp = Nothing
    HabilitoTxt (False)
    txtdescripcion.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtcodigo = ""
    txtdescripcion = ""
    
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtdescripcion.Locked = habilito
    
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    
    cmdCancelar.Enabled = hab1
    cmdaceptar.Enabled = hab2
    cmdeliminar.Enabled = hab3
    cmdnuevo.Enabled = hab4
    cmdmodificar.Enabled = hab5
    cmdBuscar.Enabled = hab6
    
End Sub
Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    
    cmdprimero.Enabled = hab2
    cmdanterior.Enabled = hab1
    cmdsiguiente.Enabled = hab3
    cmdultimo.Enabled = hab4
    
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioTxt
    Call HabilitoControles(False, False, False, True, False, True)
    
    HabilitoBotonesMoverse False, False, False, False
End Sub
Sub CargarDatos()
    
    If rscomp.State = 1 Then
        rscomp.Close
        Set rscomp = Nothing
    End If
    rscomp.Open "select * from TipoComprobantesGrales where activo=1", DaTaEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    
    If Not rscomp.EOF Then
        rscomp.MoveFirst
        rscomp.Find "Codigo= " & str(Trim(txtcodigo))
        txtdescripcion = rscomp!descripcion
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub




' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'

