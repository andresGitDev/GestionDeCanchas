VERSION 5.00
Begin VB.Form FrmFormasPagos 
   Caption         =   "Formas de pago"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "FrmFormasPagos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optctacte 
      Caption         =   "Cta. Cte."
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
      Height          =   240
      Left            =   660
      TabIndex        =   18
      Tag             =   "1"
      Top             =   1620
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton optcontado 
      Caption         =   "Contado"
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
      Height          =   240
      Left            =   2160
      TabIndex        =   17
      Top             =   1620
      Width           =   1335
   End
   Begin VB.TextBox txtdias 
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1005
      Width           =   735
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
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
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
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
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
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
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdprimero 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      Picture         =   "FrmFormasPagos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Primero"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4430
      Picture         =   "FrmFormasPagos.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ultimo"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3810
      Picture         =   "FrmFormasPagos.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Siguiente"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3130
      Picture         =   "FrmFormasPagos.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Anterior"
      Top             =   2280
      Width           =   675
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1005
      Width           =   3375
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
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   435
      Width           =   975
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   360
      TabIndex        =   19
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00400000&
      Height          =   465
      Left            =   3960
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton Opttarjeta 
         Caption         =   "Tarjeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1935
         TabIndex        =   22
         Top             =   165
         Width           =   1485
      End
      Begin VB.OptionButton Optefectivo 
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
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   165
         Width           =   1485
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Dias :"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   1020
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1935
      Left            =   120
      Top             =   165
      Width           =   7455
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
      Left            =   360
      TabIndex        =   10
      Top             =   1020
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
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   502
      Width           =   1095
   End
End
Attribute VB_Name = "FrmFormasPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4


Dim Ope As String
Dim rsFP As New ADODB.Recordset
Private Sub Controles()
 Call HabilitoControles(True, True, False, False, False, False)
End Sub
Private Sub cmdAceptar_Click()
Dim fecha As Variant
Dim rs As New ADODB.Recordset
Dim codigo As Long

    Call HabilitoControles(False, False, False, True, False, True)
    
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", vbInformation, "Atencion"
        txtDescripcion.SetFocus
        Controles
        Exit Sub
    Else
      If txtdias = "" Then
         MsgBox "Debe ingresar la cantidad de dias", vbInformation, "Atencion"
         txtdias.SetFocus
         Controles
         Exit Sub
      End If
      If optContado.Value = False And optCtaCte = False Then
         MsgBox "Debe condicion de Pago", vbInformation, "Atencion"
         optContado.SetFocus
         Controles
         Exit Sub
      End If
      If optContado = True And Optefectivo = False And Opttarjeta = False Then
        MsgBox "Debe seleccionar una Tipo de Pago", vbInformation, "Atencion"
        Frame3.Visible = True
        'Optefectivo.SetFocus
        
        Controles
        Exit Sub
      End If
      If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from formaspago", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'            If Not rs.EOF Then
'                Codigo = rs!cod + 1
'            Else
'                Codigo = 1
'            End If
            codigo = nSinNull(rs!cod) + 1
            rs.Close
            Set rs = Nothing
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DataEnvironment1.dbo_FORMAPAGO "A", codigo, Trim(txtDescripcion), Val(txtdias), Optefectivo, Opttarjeta, fecha, UsuarioSistema!codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
      Else
            If Ope = "M" Then
                fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.dbo_FORMAPAGO "M", Val(txtCodigo), Trim(txtDescripcion), Val(txtdias), Optefectivo, Opttarjeta, 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FomasPago", UsuarioSistema!codigo, fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                
            End If
      End If
    End If
    LimpioTxt
    HabilitoTxt (True)
    Frame3.Visible = False
    Optefectivo.Visible = False
    Opttarjeta.Visible = False
End Sub

Private Sub cmdanterior_Click()
    rsFP.MovePrevious
    If Not rsFP.BOF Then
        txtCodigo = rsFP!codigo
        txtDescripcion = rsFP!DESCRIPCION
        txtdias = rsFP!Dias
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "FormasPago", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
    
    Dim resu As String

    resu = frmBuscar.MostrarCodigoDescripcionActivo("FormasPago")
    If resu > "" Then
        txtCodigo = resu
        txtDescripcion = frmBuscar.resultado(2)
        CargarDatos
        Call HabilitoControles(True, False, True, False, True, False)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitoControles(False, False, False, True, False, True)
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoBotonesMoverse(False, False, False, False)
    Frame3.Visible = False
    Optefectivo.Visible = False
    Opttarjeta.Visible = False
End Sub

Private Sub cmdeliminar_Click()
Dim fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_FORMAPAGO "B", Trim(txtCodigo), "", 0, 0, 0, 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FomasPago", UsuarioSistema!codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsFP.MoveFirst
    txtCodigo = rsFP!codigo
    txtDescripcion = rsFP!DESCRIPCION
    txtdias = rsFP!Dias
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsFP.MoveNext
    If Not rsFP.EOF Then
        txtCodigo = rsFP!codigo
        txtDescripcion = rsFP!DESCRIPCION
        txtdias = rsFP!Dias
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsFP.MoveLast
    txtCodigo = rsFP!codigo
    txtDescripcion = rsFP!DESCRIPCION
    txtdias = rsFP!Dias
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub optcontado_Click()
    Frame3.Visible = True
End Sub

Private Sub optctacte_Click()
    Frame3.Visible = False
End Sub

Private Sub txtDescripcion_GotFocus()
    txtDescripcion.SelStart = 0
    txtDescripcion.SelLength = Len(txtDescripcion.Text)
End Sub
Private Sub txtdias_GotFocus()

    txtdias.SelStart = 0
    txtdias.SelLength = Len(txtdias.Text)

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Txtdias_KeyPress(KeyAscii As Integer)
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
Dim rsformapago As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsformapago.Open "select max(codigo) as cod from formaspago", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsformapago!cod) Then
        txtCodigo = rsformapago!cod + 1
    Else
        txtCodigo = 1
    End If
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
    txtdias = ""
    optContado = False
    optCtaCte = True
    Optefectivo = False
    Opttarjeta = False
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    txtdias.Locked = habilito
    optContado.enabled = Not habilito
    optCtaCte.enabled = Not habilito
    Optefectivo.enabled = Not habilito
    Opttarjeta.enabled = Not habilito

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
    
    If rsFP.State = 1 Then
        rsFP.Close
        Set rsFP = Nothing
    End If
    rsFP.Open "select * from formaspago where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsFP.EOF Then
        rsFP.MoveFirst
        rsFP.Find "Codigo= " & STR(Trim(txtCodigo))
        txtdias = rsFP!Dias
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub

' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'

