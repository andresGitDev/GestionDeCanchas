VERSION 5.00
Begin VB.Form frmTipoProd 
   Caption         =   "Tipos de productos"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   345
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2010
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   840
      Width           =   4680
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2655
      Picture         =   "frmTipoProd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   1545
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3330
      Picture         =   "frmTipoProd.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Siguiente"
      Top             =   1545
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   3945
      Picture         =   "frmTipoProd.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   1545
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2040
      Picture         =   "frmTipoProd.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   1545
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
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2385
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
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2385
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
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2385
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
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2385
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2385
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
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2385
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
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2385
      Width           =   975
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
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   1095
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
      Left            =   450
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmTipoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4

Dim Ope As String
Dim rsTipo As New ADODB.Recordset

Private Sub cmdAceptar_Click()

Dim Fecha As Variant
Dim rs As New ADODB.Recordset
Dim codigo As Long


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from tipoproductos", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not IsNull(rs!COD) Then
                codigo = rs!COD + 1
            Else
                codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DataEnvironment1.Sistema.Execute "insert into tipoproductos (codigo,descripcion,fecha_alta,usuario_alta,activo) values (" & txtCodigo.Text & ",'" & Trim(txtDescripcion) & "'," & ssFecha(Date) & "," & UsuarioActual & ",1)"
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.Sistema.Execute "update tipoproductos set descripcion='" & Trim(txtDescripcion.Text) & "' where codigo=" & Trim(txtCodigo)
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "tipoproductos", UsuarioSistema!codigo, Fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    '
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
    resu = frmBuscar.MostrarCodigoDescripcionActivo("tipoproductos")
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
Dim Fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.Sistema.Execute "update tipoproductos set usuario_baja=" & UsuarioActual & ",fecha_baja=" & ssFecha(Date) & ",activo=0 where codigo=" & Trim(txtCodigo)
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "tipoproductos", UsuarioSistema!codigo, Fecha, Time, "B"
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

    FrmKeyPress KeyAscii, True, True

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
Dim rsTipo As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsTipo.Open "select max(codigo) as cod from tipoproductos", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsTipo!COD) Then
        txtCodigo = rsTipo!COD + 1
    Else
        txtCodigo = 1
    End If
    rsTipo.Close
    Set rsTipo = Nothing
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
    cmdEliminar.enabled = hab3
    cmdNuevo.enabled = hab4
    cmdModificar.enabled = hab5
    cmdBuscar.enabled = hab6
    
End Sub
Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    
    cmdPrimero.enabled = hab2
    cmdAnterior.enabled = hab1
    cmdSiguiente.enabled = hab3
    cmdUltimo.enabled = hab4
    
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
    rsTipo.Open "select * from tipoproductos where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsTipo.EOF Then
        rsTipo.MoveFirst
        rsTipo.Find "Codigo= " & str(Trim(txtCodigo))
        txtDescripcion = rsTipo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub
