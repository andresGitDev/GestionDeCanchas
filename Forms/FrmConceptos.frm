VERSION 5.00
Begin VB.Form FrmConceptos 
   Caption         =   "Conceptos"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "FrmConceptos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkrechazo 
      Alignment       =   1  'Right Justify
      Caption         =   "Rechazo"
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
      Left            =   4440
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox cmbcomprobante 
      Height          =   315
      ItemData        =   "FrmConceptos.frx":08CA
      Left            =   2040
      List            =   "FrmConceptos.frx":08CC
      TabIndex        =   18
      Text            =   "cmbcomprobante"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ComboBox cmbmovimiento 
      Height          =   315
      ItemData        =   "FrmConceptos.frx":08CE
      Left            =   2040
      List            =   "FrmConceptos.frx":08DE
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1800
      Width           =   1455
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
      Left            =   2775
      Picture         =   "FrmConceptos.frx":08FF
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   2505
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3450
      Picture         =   "FrmConceptos.frx":0C09
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Siguiente"
      Top             =   2505
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   4065
      Picture         =   "FrmConceptos.frx":0F13
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   2505
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2160
      Picture         =   "FrmConceptos.frx":121D
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   2505
      Width           =   615
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
      Top             =   3345
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
      TabIndex        =   5
      Top             =   3345
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
      TabIndex        =   4
      Top             =   3345
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
      TabIndex        =   3
      Top             =   3345
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
      TabIndex        =   2
      Top             =   3345
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
      TabIndex        =   1
      Top             =   3345
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
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3345
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Movimiento: "
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
      Left            =   480
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Comprobante :"
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
      Left            =   480
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2205
      Left            =   80
      Top             =   120
      Width           =   6840
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
Attribute VB_Name = "FrmConceptos"
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
Dim codigo As Long
Dim rechazo As Long
Dim Movi As String


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    Else
        If Trim(cmbcomprobante.Text) = "" Then
            MsgBox "El comprobante no puede quedar en blanco", 48, "Atencion"
            Exit Sub
        Else
            If Trim(cmbmovimiento.Text) = "" Then
                MsgBox "El movimiento no puede quedar en blanco", 48, "Atencion"
                Exit Sub
            Else
                Select Case Trim(cmbmovimiento.Text)
                    Case "Suma"
                        Movi = "S"
                    Case "Resta"
                        Movi = "R"
                    Case "Ambos"
                        Movi = "A"
                    Case "Interno"
                        Movi = "I"
                End Select
                If Ope = "A" Then
                    rs.Open "Select max(codigo) as cod from conceptos", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
                    If rs.EOF = True And rs.BOF = True Then
                        codigo = 1
                    Else
                        If IsNull(rs!cod) Then
                            codigo = 1
                        Else
                            codigo = rs!cod + 1
                        End If
                    End If
                    rs.Close
                    Set rs = Nothing
                    fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                    If chkrechazo.Value = 1 Then
                        rechazo = 1
                    Else
                        rechazo = 0
                    End If
                    DataEnvironment1.dbo_CONCEPTO "A", codigo, ObtenerCodigo("TipoComprobantes", Trim(cmbcomprobante.Text)), Movi, rechazo, Trim(txtDescripcion), fecha, UsuarioSistema!codigo, 0, 0
                    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                Else
                    If Ope = "M" Then
                        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                        DataEnvironment1.dbo_CONCEPTO "M", Val(txtCodigo), ObtenerCodigo("TipoComprobantes", Trim(cmbcomprobante.Text)), Movi, rechazo, Trim(txtDescripcion), 0, 0, 0, 0
                        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Conceptos", UsuarioSistema!codigo, fecha, Time, "M"
                        MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                        
                    End If
                End If
                LimpioTxt
                HabilitoTxt (True)
            End If
        End If
    End If
    
End Sub

Private Sub cmdanterior_Click()
    rscomp.MovePrevious
    If Not rscomp.BOF Then
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
'    CargarHelp "Conceptos", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
    Dim resu As String

    resu = frmBuscar.MostrarCodigoDescripcionActivo("Conceptos")
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
        DataEnvironment1.dbo_CONCEPTO "B", Trim(txtCodigo), 0, "", 0, "", 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Conceptos", UsuarioSistema!codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rscomp.MoveFirst
    CargoRegistro
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rscomp.MoveNext
    If Not rscomp.EOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rscomp.MoveLast
    CargoRegistro
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
Sub CargoRegistro()
    txtCodigo = rscomp!codigo
    txtDescripcion = rscomp!DESCRIPCION
    cmbcomprobante.Text = ObtenerDescripcion("TipoComprobantes", rscomp!comprobante)
    Select Case rscomp!movimiento
        Case "S"
            cmbmovimiento.Text = "Suma"
        Case "R"
            cmbmovimiento.Text = "Resta"
        Case "A"
            cmbmovimiento.Text = "Ambos"
        Case "I"
            cmbmovimiento.Text = "Interno"
    End Select
    
    If rscomp!rechazo = 1 Then
        chkrechazo.Value = 1
    Else
        chkrechazo.Value = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rscomp.State = 1 Then
        rscomp.Close
        Set rscomp = Nothing
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
Private Sub cmbcomprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub
Private Sub cmbovmiento_KeyPress(KeyAscii As Integer)
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
Dim rscomp As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rscomp.Open "select max(codigo) as cod from conceptos", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rscomp!cod) Then
        txtCodigo = rscomp!cod + 1
    Else
        txtCodigo = 1
    End If
    rscomp.Close
    Set rscomp = Nothing
    cmbcomprobante.Text = "REMITO OFICIAL"
    cmbmovimiento.Text = "Suma"
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
    chkrechazo.Value = 0
    cmbcomprobante.ListIndex = -1
    cmbmovimiento.ListIndex = -1
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    chkrechazo.enabled = Not habilito
    cmbcomprobante.Locked = habilito
    cmbmovimiento.Locked = habilito
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
    CargaCombo cmbcomprobante, "tipocomprobantes", "descripcion", "codigo", ""
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoControles(False, False, False, True, False, True)
    HabilitoBotonesMoverse False, False, False, False
End Sub
Sub CargarDatos()

    
    If rscomp.State = 1 Then
        rscomp.Close
        Set rscomp = Nothing
    End If
    rscomp.Open "select * from Conceptos where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rscomp.EOF Then
        rscomp.MoveFirst
        rscomp.Find "Codigo= " & STR(Trim(txtCodigo))
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub


' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'


