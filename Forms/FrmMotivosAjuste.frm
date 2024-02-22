VERSION 5.00
Begin VB.Form FrmMotivosAjuste 
   Caption         =   "Motivos de ajustes"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "FrmMotivosAjuste.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   1935
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   315
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1935
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   810
      Width           =   4815
   End
   Begin VB.CommandButton cmdanterior 
      Height          =   615
      Left            =   2580
      Picture         =   "FrmMotivosAjuste.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   1515
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3255
      Picture         =   "FrmMotivosAjuste.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Siguiente"
      Top             =   1515
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   3870
      Picture         =   "FrmMotivosAjuste.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   1515
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   1965
      Picture         =   "FrmMotivosAjuste.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   1515
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
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2355
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
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2355
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2355
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
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2355
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2355
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
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2355
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
      Left            =   1020
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2355
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1290
      Left            =   90
      Top             =   90
      Width           =   6810
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
      TabIndex        =   14
      Top             =   330
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
      Left            =   375
      TabIndex        =   13
      Top             =   810
      Width           =   1455
   End
End
Attribute VB_Name = "FrmMotivosAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4


Dim Ope As String
Dim rsmotivo As New ADODB.Recordset

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
            rs.Open "Select max(codigo) as cod from MotivosAjuste", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
                codigo = rs!cod + 1
            Else
                codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            'fecha = Date
            DataEnvironment1.dbo_MOTIVO "A", codigo, Trim(txtDescripcion), fecha, UsuarioSistema!codigo, 0, "01/01/1900"
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.dbo_MOTIVO "M", Val(txtCodigo), Trim(txtDescripcion), 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "MotivosAjuste", UsuarioSistema!codigo, fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rsmotivo.MovePrevious
    If Not rsmotivo.BOF Then
        txtCodigo = rsmotivo!codigo
        txtDescripcion = rsmotivo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "MotivosAjuste", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name

    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("MotivosAjuste")
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
        DataEnvironment1.dbo_MOTIVO "B", Trim(txtCodigo), "", 0, 0, UsuarioSistema!codigo, fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "MotivosAjuste", UsuarioSistema!codigo, fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsmotivo.MoveFirst
    txtCodigo = rsmotivo!codigo
    txtDescripcion = rsmotivo!DESCRIPCION
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsmotivo.MoveNext
    If Not rsmotivo.EOF Then
        txtCodigo = rsmotivo!codigo
        txtDescripcion = rsmotivo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsmotivo.MoveLast
    txtCodigo = rsmotivo!codigo
    txtDescripcion = rsmotivo!DESCRIPCION
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsmotivo.State = 1 Then
        rsmotivo.Close
        Set rsmotivo = Nothing
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
Dim rsmotivo As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsmotivo.Open "select max(codigo) as cod from MotivosAjuste", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsmotivo!cod) Then
        txtCodigo = rsmotivo!cod + 1
    Else
        txtCodigo = 1
    End If
    rsmotivo.Close
    Set rsmotivo = Nothing
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
    
    'HabilitoBotonesMoverse False, False, False, False
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioTxt
    Call HabilitoControles(False, False, False, True, False, True)
    
End Sub
Sub CargarDatos()
    
    If rsmotivo.State = 1 Then
        rsmotivo.Close
        Set rsmotivo = Nothing
    End If
    rsmotivo.Open "select * from MotivosAjuste where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsmotivo.EOF Then
        rsmotivo.MoveFirst
        rsmotivo.Find "Codigo= " & STR(Trim(txtCodigo))
        txtDescripcion = rsmotivo!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub


' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'

