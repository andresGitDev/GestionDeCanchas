VERSION 5.00
Begin VB.Form FrmTransportes 
   Caption         =   "Transportes"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "FrmTransportes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   3240
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
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
      Left            =   1890
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   345
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4680
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2655
      Picture         =   "FrmTransportes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Anterior"
      Top             =   2505
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3330
      Picture         =   "FrmTransportes.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Siguiente"
      Top             =   2505
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   3945
      Picture         =   "FrmTransportes.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ultimo"
      Top             =   2505
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2040
      Picture         =   "FrmTransportes.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Primero"
      Top             =   2505
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
      TabIndex        =   9
      Top             =   3225
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
      TabIndex        =   8
      Top             =   3225
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
      TabIndex        =   7
      Top             =   3225
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
      TabIndex        =   6
      Top             =   3225
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
      TabIndex        =   5
      Top             =   3225
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
      TabIndex        =   4
      Top             =   3225
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
      TabIndex        =   3
      Top             =   3225
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Telfono/s :"
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
      Top             =   1815
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2250
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
      Left            =   360
      TabIndex        =   16
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
      Left            =   330
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "FrmTransportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit ' mod 12/8/4

Dim Ope As String
Dim rstrans As New ADODB.Recordset

Private Sub cmdAceptar_Click()

Dim fecha As Variant
Dim rs As New ADODB.Recordset
Dim codigo As Long


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from Transportes", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rs.Fields(0) > 0 Then
                codigo = rs!cod + 1
            Else
                codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DataEnvironment1.dbo_TRANSPORTE "A", codigo, Trim(txtDescripcion), Trim(txtDireccion), Trim(txtTel), fecha, UsuarioSistema!codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.dbo_TRANSPORTE "M", Val(txtCodigo), Trim(txtDescripcion), Trim(txtDireccion), Trim(txtTel), 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "transportes", UsuarioSistema!codigo, fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rstrans.MovePrevious
    If Not rstrans.BOF Then
        txtCodigo = rstrans!codigo
        txtDescripcion = rstrans!DESCRIPCION
        If Not IsNull(rstrans!direccion) Then
            txtDireccion = rstrans!direccion
        End If
        If Not IsNull(rstrans!Telefono) Then
            txtTel = rstrans!Telefono
        End If
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "Transportes", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("Transportes")
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
Dim fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_TRANSPORTE "B", Trim(txtCodigo), "", "", "", 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "transportes", UsuarioSistema!codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rstrans.MoveFirst
    txtCodigo = rstrans!codigo
    txtDescripcion = rstrans!DESCRIPCION
    If Not IsNull(rstrans!direccion) Then
        txtDireccion = rstrans!direccion
    End If
    If Not IsNull(rstrans!Telefono) Then
        txtTel = rstrans!Telefono
    End If
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rstrans.MoveNext
    If Not rstrans.EOF Then
    txtCodigo = rstrans!codigo
    txtDescripcion = rstrans!DESCRIPCION
    If Not IsNull(rstrans!direccion) Then
        txtDireccion = rstrans!direccion
    End If
    If Not IsNull(rstrans!Telefono) Then
        txtTel = rstrans!Telefono
    End If
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rstrans.MoveLast
    txtCodigo = rstrans!codigo
    txtDescripcion = rstrans!DESCRIPCION
    If Not IsNull(rstrans!direccion) Then
        txtDireccion = rstrans!direccion
    End If
    If Not IsNull(rstrans!Telefono) Then
        txtTel = rstrans!Telefono
    End If
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rstrans.State = 1 Then
        rstrans.Close
        Set rstrans = Nothing
    End If
End Sub

Private Sub txtDescripcion_GotFocus()

    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = Len(txtDescripcion.Text)

End Sub
Private Sub txtDireccion_GotFocus()

    txtDireccion.SelStart = 0
    txtDireccion.SelLength = Len(txtDireccion.Text)

End Sub
Private Sub txttel_GotFocus()

    txtTel.SelStart = 0
    txtTel.SelLength = Len(txtTel.Text)

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
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

Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtDescripcion.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
Dim rstransporte As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rstransporte.Open "select max(codigo) as cod from transportes", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rstransporte!cod) Then
        txtCodigo = rstransporte!cod + 1
    Else
        txtCodigo = 1
    End If
    rstransporte.Close
    Set rstransporte = Nothing
    HabilitoTxt (False)
    txtDescripcion.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtCodigo = ""
    txtDescripcion = ""
    txtDireccion = ""
    txtTel = ""
    
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    txtDireccion.Locked = habilito
    txtTel.Locked = habilito
    
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    
    cmdcancelar.enabled = hab1
    cmdaceptar.enabled = hab2
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
    
    If rstrans.State = 1 Then
        rstrans.Close
        Set rstrans = Nothing
    End If
    rstrans.Open "select * from transportes where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rstrans.EOF Then
        rstrans.MoveFirst
        rstrans.Find "Codigo= " & STR(Trim(txtCodigo))
        txtDescripcion = rstrans!DESCRIPCION
        If Not IsNull(rstrans!direccion) Then
            txtDireccion = rstrans!direccion
        End If
        If Not IsNull(rstrans!Telefono) Then
            txtTel = rstrans!Telefono
        End If
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub


' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
'   inhibo mov al cargar
'
