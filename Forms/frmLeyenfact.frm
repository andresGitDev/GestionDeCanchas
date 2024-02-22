VERSION 5.00
Begin VB.Form frmLeyenfact 
   Caption         =   "Leyendas de factura"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   7755
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
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   390
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   1365
      Left            =   1920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   960
      Width           =   5295
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
      TabIndex        =   6
      Top             =   2835
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
      TabIndex        =   5
      Top             =   2835
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
      TabIndex        =   4
      Top             =   2835
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
      TabIndex        =   3
      Top             =   2835
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
      Top             =   2835
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
      TabIndex        =   1
      Top             =   2835
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2835
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
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   450
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
      Left            =   360
      TabIndex        =   9
      Top             =   975
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2415
      Left            =   120
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmLeyenfact"
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
Dim rs As New ADODB.Recordset
Dim codigo As Long

    Call HabilitoControles(False, False, False, True, False, True)
    
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", vbInformation, "Atencion"
        txtDescripcion.SetFocus
        Controles
        Exit Sub
    End If
      
    If Ope = "A" Then
          rs.Open "Select max(id) as cod from factleyenda", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
          codigo = nSinNull(rs!COD) + 1
          rs.Close
          Set rs = Nothing
          DataEnvironment1.Sistema.Execute "insert into factleyenda (leyenda) values ('" & Trim(txtDescripcion) & "')"
          MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    Else
          If Ope = "M" Then
              DataEnvironment1.Sistema.Execute "update factleyenda set leyenda='" & Trim(txtDescripcion) & "' where id=" & txtCodigo
              DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FactLeyenda", UsuarioSistema!codigo, Date, Time, "M"
              MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
              
          End If
    End If

    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String

    resu = frmBuscar.MostrarSql("select id as [ Codigo  ],leyenda as [ Leyenda                                                           ] from factleyenda")
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
End Sub

Private Sub cmdeliminar_Click()
Dim mensaje As String
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        DataEnvironment1.Sistema.Execute "delete from factleyenda where id=" & txtCodigo
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "FactLeyenda", UsuarioSistema!codigo, Date, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
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
 'KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
    rsformapago.Open "select max(id) as cod from factleyenda", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not IsNull(rsformapago!COD) Then
        txtCodigo = rsformapago!COD + 1
    Else
        txtCodigo = 1
    End If
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
    
    cmdcancelar.enabled = hab1
    cmdAceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdmodificar.enabled = hab5
    cmdbuscar.enabled = hab6
    
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioTxt
    Call HabilitoControles(False, False, False, True, False, True)
    
End Sub
Sub CargarDatos()
    
    If rsFP.State = 1 Then
        rsFP.Close
        Set rsFP = Nothing
    End If
    rsFP.Open "select * from formaspago where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rsFP.EOF Then
        rsFP.MoveFirst
        rsFP.Find "Codigo= " & str(Trim(txtCodigo))
    End If

End Sub

' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'

