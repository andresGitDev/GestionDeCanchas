VERSION 5.00
Begin VB.Form FrmABMUnidadTipos 
   Caption         =   "ABM de tipos de unidades"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   Icon            =   "FrmABMUnidadTipos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbreviatura 
      Height          =   285
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   2385
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
      Height          =   795
      Left            =   3285
      Picture         =   "FrmABMUnidadTipos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2595
      Width           =   870
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
      DisabledPicture =   "FrmABMUnidadTipos.frx":0E54
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
      Height          =   780
      Left            =   4155
      Picture         =   "FrmABMUnidadTipos.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1830
      Width           =   870
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
      DisabledPicture =   "FrmABMUnidadTipos.frx":1FE8
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4155
      Picture         =   "FrmABMUnidadTipos.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   270
      Width           =   870
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmABMUnidadTipos.frx":317C
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
      Height          =   780
      Left            =   4155
      Picture         =   "FrmABMUnidadTipos.frx":3A46
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmABMUnidadTipos.frx":4310
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
      Height          =   780
      Left            =   1545
      Picture         =   "FrmABMUnidadTipos.frx":4BDA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      DisabledPicture =   "FrmABMUnidadTipos.frx":54A4
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
      Height          =   780
      Left            =   2415
      Picture         =   "FrmABMUnidadTipos.frx":5D6E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmABMUnidadTipos.frx":6638
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   4155
      Picture         =   "FrmABMUnidadTipos.frx":6F02
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1050
      Width           =   870
   End
   Begin VB.CommandButton cmdprimero 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1920
      Picture         =   "FrmABMUnidadTipos.frx":77CC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Primero"
      Top             =   1965
      Width           =   285
   End
   Begin VB.CommandButton cmdultimo 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2820
      Picture         =   "FrmABMUnidadTipos.frx":7916
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ultimo"
      Top             =   1965
      Width           =   285
   End
   Begin VB.CommandButton cmdsiguiente 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2520
      Picture         =   "FrmABMUnidadTipos.frx":7A60
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Siguiente"
      Top             =   1965
      Width           =   285
   End
   Begin VB.CommandButton cmdanterior 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2220
      Picture         =   "FrmABMUnidadTipos.frx":7BAA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Anterior"
      Top             =   1965
      Width           =   285
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2385
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Abreviatura:"
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
      TabIndex        =   16
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   2415
      Left            =   90
      Top             =   105
      Width           =   3975
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
      TabIndex        =   15
      Top             =   960
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
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   195
      Width           =   1095
   End
End
Attribute VB_Name = "FrmABMUnidadTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ope As String
Dim rsUnidad As New ADODB.Recordset

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdaceptar_Click()
Dim codigo As Long
    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", vbInformation, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    If Trim(txtAbreviatura) = "" Then
        MsgBox "Debe cargar la abreviatura", vbInformation, "Atencion"
        txtAbreviatura.SetFocus
        Exit Sub
    End If
    
    ABMUnidadTipos Ope, txtCodigo, txtDescripcion, txtAbreviatura
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rsUnidad.MovePrevious
    If Not rsUnidad.BOF Then
        txtCodigo = rsUnidad!utcodigo
        txtDescripcion = rsUnidad!DESCRIPCION
        txtAbreviatura = rsUnidad!abreviatura
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select UtCodigo as [Nro],Descripcion,Abreviatura from umtipos order by utcodigo")
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
    If MsgBox("Esta seguro de querer eliminar este registro", vbYesNo + vbInformation, "Atencion") = vbYes Then
        ABMUnidadTipos "B", txtCodigo, "", ""
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsUnidad.MoveFirst
    txtCodigo = rsUnidad!utcodigo
    txtDescripcion = rsUnidad!DESCRIPCION
    txtAbreviatura = rsUnidad!abreviatura
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsUnidad.MoveNext
    If Not rsUnidad.EOF Then
        txtCodigo = rsUnidad!utcodigo
        txtDescripcion = rsUnidad!DESCRIPCION
        txtAbreviatura = rsUnidad!abreviatura
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsUnidad.MoveLast
    txtCodigo = rsUnidad!utcodigo
    txtDescripcion = rsUnidad!DESCRIPCION
    txtAbreviatura = rsUnidad!abreviatura
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsUnidad.State = 1 Then
        Set rsUnidad = Nothing
    End If
End Sub

Private Sub txtAbreviatura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
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
    Ope = "A"
    LimpioTxt
    txtCodigo = 0
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
    txtAbreviatura = ""
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    txtAbreviatura.Locked = habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    cmdcancelar.enabled = hab1
    cmdaceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdmodificar.enabled = hab5
    cmdbuscar.enabled = hab6
    HabilitoBotonesMoverse False, False, False, False
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
End Sub

Sub CargarDatos()
    If rsUnidad.State = 1 Then
        Set rsUnidad = Nothing
    End If
    rsUnidad.Open "select * from umtipos", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsUnidad.EOF Then
        rsUnidad.MoveFirst
        rsUnidad.Find "UTCodigo= " & txtCodigo
        txtDescripcion = sSinNull(rsUnidad!DESCRIPCION)
        txtAbreviatura = sSinNull(rsUnidad!abreviatura)
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub

Private Function ABMUnidadTipos(uOpe As String, uCodigo As Long, uDescrip As String, uAbreviatura As String) As Boolean
On Error GoTo MAL
Dim ABMUM As String, SP1 As Boolean, SP2 As Boolean, ET, MUESTRO As String
ABMUnidadTipos = True
Select Case uOpe
    Case "A":
        ABMUM = "INSERT INTO UMTIPOS (DESCRIPCION,ABREVIATURA) VALUES " _
            & "(" & ssTexto(uDescrip) & "," & ssTexto(uAbreviatura) & ")"
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "TUM guardada", vbInformation, "Tipo Unidad nueva"
    Case "M":
        ABMUM = "UPDATE UMTIPOS SET DESCRIPCION=" & ssTexto(uDescrip) _
            & ",ABREVIATURA=" & ssTexto(uAbreviatura) _
            & " where utcodigo=" & uCodigo
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "TUM actualizada", vbInformation, "Tipo Unidad modificada"
    Case "B":
        Set ET = Nothing
        ET = obtenerDeSQL("select * from unidadesmedida where TIPO=" & uCodigo)
        If IsNull(ET) Or IsEmpty(ET) Then
            SP1 = False
        Else
            SP1 = True
        End If
        Set ET = Nothing
        ET = obtenerDeSQL("select * from UMFACTOR where TIPO=" & uCodigo)
        If IsNull(ET) Or IsEmpty(ET) Then
            SP2 = False
        Else
            SP2 = True
        End If
        
        If SP1 And SP2 Then
            ABMUM = "DELETE FROM UMTIPOS WHERE UTCODIGO=" & uCodigo
            DataEnvironment1.Sistema.Execute ABMUM
            MsgBox "TUM eliminada", vbInformation, "Tipo Unidad borrada"
        Else
            MUESTRO = "El tipo de unidad no se puede eliminar por que:"
            If SP1 Then MUESTRO = MUESTRO & Chr(13) & " Esta asociado a una o mas unidades de medida"
            If SP2 Then MUESTRO = MUESTRO & Chr(13) & " Esta asociado a una o mas factores de medida"
            MsgBox MUESTRO, vbCritical, "No se eliminara el tipo"
            ABMUnidadTipos = False
        End If
End Select

Exit Function
MAL:
    MsgBox "Error en carga de TUM", vbCritical, "Tipo nueva no guardada"
    ABMUnidadTipos = False
End Function
