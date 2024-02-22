VERSION 5.00
Begin VB.Form FrmBancosGenerales 
   Caption         =   "Bancos"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "FrmBancosGenerales.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2925
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1860
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   390
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   960
      Width           =   5295
   End
   Begin VB.CommandButton cmdanterior 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3075
      Picture         =   "FrmBancosGenerales.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   1635
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3750
      Picture         =   "FrmBancosGenerales.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Siguiente"
      Top             =   1635
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Enabled         =   0   'False
      Height          =   615
      Left            =   4365
      Picture         =   "FrmBancosGenerales.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   1635
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Enabled         =   0   'False
      Height          =   615
      Left            =   2460
      Picture         =   "FrmBancosGenerales.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   1635
      Width           =   615
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
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2460
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
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2460
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
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2460
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
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2460
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2460
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
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2460
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
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2460
      Width           =   975
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
      Left            =   300
      TabIndex        =   14
      Top             =   480
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
      Left            =   300
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   60
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "FrmBancosGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4


Dim Ope As String
Dim rsbanco As New ADODB.Recordset

Private Sub cmdaceptar_Click()

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
            rs.Open "Select max(codigo) as cod from BancosGrales", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
                codigo = nSinNull(rs!COD) + 1
            Else
                codigo = 1
            End If
            rs.Close
            Set rs = Nothing
            Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            DataEnvironment1.dbo_BANCOGRAL "A", codigo, Trim(txtDescripcion), Fecha, UsuarioSistema!codigo, 0, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                DataEnvironment1.dbo_BANCOGRAL "M", Val(txtCodigo), Trim(txtDescripcion), 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "BancosGrales", UsuarioSistema!codigo, Fecha, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rsbanco.MovePrevious
    If Not rsbanco.BOF Then
        txtCodigo = rsbanco!codigo
        txtDescripcion = rsbanco!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)

    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo, descripcion from BancosGrales where activo = 1")
    If resu > "" Then
        txtCodigo = resu
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
Dim Fecha As Variant
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        DataEnvironment1.dbo_BANCOGRAL "B", Trim(txtCodigo), "", 0, 0, UsuarioSistema!codigo, Fecha
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "BancosGrales", UsuarioSistema!codigo, Fecha, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsbanco.MoveFirst
    txtCodigo = rsbanco!codigo
    txtDescripcion = rsbanco!DESCRIPCION
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsbanco.MoveNext
    If Not rsbanco.EOF Then
        txtCodigo = rsbanco!codigo
        txtDescripcion = rsbanco!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsbanco.MoveLast
    txtCodigo = rsbanco!codigo
    txtDescripcion = rsbanco!DESCRIPCION
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsbanco.State = 1 Then
        rsbanco.Close
        Set rsbanco = Nothing
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
Dim rsbanco As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    rsbanco.Open "select max(codigo) as cod from BancosGrales", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsbanco!COD) Then
        txtCodigo = rsbanco!COD + 1
    Else
        txtCodigo = 1
    End If
    rsbanco.Close
    Set rsbanco = Nothing
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
    
    If rsbanco.State = 1 Then
        rsbanco.Close
        Set rsbanco = Nothing
    End If
    rsbanco.Open "select * from bancosGrales where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsbanco.EOF Then
        rsbanco.MoveFirst
        rsbanco.Find "Codigo= " & STR(Trim(txtCodigo))
        txtDescripcion = rsbanco!DESCRIPCION
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub


' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'

