VERSION 5.00
Begin VB.Form FrmABMUnidadMedida 
   Caption         =   "ABM de unidades de medida"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   Icon            =   "FrmABMUnidadMedida.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2745
   End
   Begin VB.TextBox txtAbreviatura 
      Height          =   285
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1260
      Width           =   2745
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
      Left            =   5445
      Picture         =   "FrmABMUnidadMedida.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2595
      Width           =   870
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
      DisabledPicture =   "FrmABMUnidadMedida.frx":0E54
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
      Left            =   1965
      Picture         =   "FrmABMUnidadMedida.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
      DisabledPicture =   "FrmABMUnidadMedida.frx":1FE8
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
      Left            =   75
      Picture         =   "FrmABMUnidadMedida.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmABMUnidadMedida.frx":317C
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
      Left            =   2835
      Picture         =   "FrmABMUnidadMedida.frx":3A46
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmABMUnidadMedida.frx":4310
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
      Left            =   3705
      Picture         =   "FrmABMUnidadMedida.frx":4BDA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      DisabledPicture =   "FrmABMUnidadMedida.frx":54A4
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
      Left            =   4575
      Picture         =   "FrmABMUnidadMedida.frx":5D6E
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmABMUnidadMedida.frx":6638
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
      Left            =   955
      Picture         =   "FrmABMUnidadMedida.frx":6F02
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2610
      Width           =   870
   End
   Begin VB.CommandButton cmdprimero 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1920
      Picture         =   "FrmABMUnidadMedida.frx":77CC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Primero"
      Top             =   1965
      Width           =   645
   End
   Begin VB.CommandButton cmdultimo 
      Enabled         =   0   'False
      Height          =   390
      Left            =   3900
      Picture         =   "FrmABMUnidadMedida.frx":7916
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ultimo"
      Top             =   1965
      Width           =   645
   End
   Begin VB.CommandButton cmdsiguiente 
      Enabled         =   0   'False
      Height          =   390
      Left            =   3240
      Picture         =   "FrmABMUnidadMedida.frx":7A60
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Siguiente"
      Top             =   1965
      Width           =   645
   End
   Begin VB.CommandButton cmdanterior 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2580
      Picture         =   "FrmABMUnidadMedida.frx":7BAA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Anterior"
      Top             =   1965
      Width           =   645
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2745
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
      Height          =   360
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo:"
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
      Left            =   1545
      TabIndex        =   18
      Top             =   1575
      Width           =   1815
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
      Left            =   930
      TabIndex        =   17
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   2415
      Left            =   90
      Top             =   105
      Width           =   6255
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
      Left            =   900
      TabIndex        =   16
      Top             =   960
      Width           =   1815
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
      Left            =   2760
      TabIndex        =   15
      Top             =   195
      Width           =   1455
   End
End
Attribute VB_Name = "FrmABMUnidadMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ope As String
Dim rsUnidad As New ADODB.Recordset
Private uTipos() As datCD

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim codigo As Long
    
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
    
    ABMUnidadMedida Ope, txtCodigo, txtDescripcion, txtAbreviatura, uTipos(cboTipo.ListIndex).dCodigo
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoControles(False, False, False, True, False, True)
End Sub

Private Sub cmdanterior_Click()
    rsUnidad.MovePrevious
    If Not rsUnidad.BOF Then
        txtCodigo = rsUnidad!umcodigo
        txtDescripcion = rsUnidad!DESCRIPCION
        txtAbreviatura = rsUnidad!abreviatura
        cboTipo.ListIndex = OBT(CLng(nSinNull(rsUnidad!tipo)))
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select UMCodigo as [Nro],Descripcion,Abreviatura from unidadesmedida where activo=1 order by umcodigo")
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
    If MsgBox("Esta seguro de querer eliminar este registro", vbYesNo + vbInformation, "Atencion") = vbYes Then
        ABMUnidadMedida "B", txtCodigo, "", "", 0
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsUnidad.MoveFirst
    txtCodigo = rsUnidad!umcodigo
    txtDescripcion = rsUnidad!DESCRIPCION
    txtAbreviatura = rsUnidad!abreviatura
    cboTipo.ListIndex = OBT(CLng(nSinNull(rsUnidad!tipo)))
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsUnidad.MoveNext
    If Not rsUnidad.EOF Then
        txtCodigo = rsUnidad!umcodigo
        txtDescripcion = rsUnidad!DESCRIPCION
        txtAbreviatura = rsUnidad!abreviatura
        cboTipo.ListIndex = OBT(CLng(nSinNull(rsUnidad!tipo)))
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsUnidad.MoveLast
    txtCodigo = rsUnidad!umcodigo
    txtDescripcion = rsUnidad!DESCRIPCION
    txtAbreviatura = rsUnidad!abreviatura
    cboTipo.ListIndex = OBT(CLng(nSinNull(rsUnidad!tipo)))
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
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
    txtCodigo = s2n(obtenerDeSQL("select max(umcodigo) from UNIDADESMEDIDA")) + 1
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
    txtAbreviatura = ""
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    txtAbreviatura.Locked = habilito
    cboTipo.Locked = habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    cmdcancelar.enabled = hab1
    cmdAceptar.enabled = hab2
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
    CargoTipos
End Sub

Sub CargarDatos()
    If rsUnidad.State = 1 Then
        Set rsUnidad = Nothing
    End If
    rsUnidad.Open "select * from unidadesmedida where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsUnidad.EOF Then
        rsUnidad.MoveFirst
        rsUnidad.Find "UMCodigo= " & txtCodigo
        txtDescripcion = sSinNull(rsUnidad!DESCRIPCION)
        txtAbreviatura = sSinNull(rsUnidad!abreviatura)
        cboTipo.ListIndex = OBT(CLng(nSinNull(rsUnidad!tipo)))
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub

Private Function ABMUnidadMedida(uOpe As String, uCodigo As Long, uDescrip As String, uAbreviatura As String, uTipo As Long) As Boolean
On Error GoTo MAL
Dim ABMUM As String
ABMUnidadMedida = True
Select Case uOpe
    Case "A":
        ABMUM = "INSERT INTO UNIDADESMEDIDA (codigo,DESCRIPCION,ABREVIATURA,TIPO,ACTIVO) VALUES " _
            & "(" & s2n(uCodigo) & "," & ssTexto(uDescrip) & "," & ssTexto(uAbreviatura) & "," & uTipo & ",1)"
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UM guardada", vbInformation, "Unidad nueva"
    Case "M":
        ABMUM = "UPDATE UNIDADESMEDIDA SET DESCRIPCION=" & ssTexto(uDescrip) _
            & ",ABREVIATURA=" & ssTexto(uAbreviatura) _
            & ",TIPO=" & uTipo _
            & " where umcodigo=" & uCodigo
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UM actualizada", vbInformation, "Unidad modificada"
    Case "B":
        If MsgBox("¿Desea eliminar por completo la unidad?", vbYesNo + vbExclamation) = vbYes Then
            ABMUM = "DELETE FROM UNIDADESMEDIDA WHERE UMCODIGO=" & uCodigo
        Else
            ABMUM = "UPDATE UNIDADESMEDIDA  SET ACTIVO=0 WHERE UMCODIGO=" & uCodigo
        End If
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UM eliminada", vbInformation, "Unidad borrada"
End Select

Exit Function
MAL:
    MsgBox "Error en carga de UM", vbCritical, "Unidad nueva no guardada"
    ABMUnidadMedida = False
End Function

Private Function OBT(dCod As Long) As Long
Dim i As Long
OBT = 0
    For i = 0 To UBound(uTipos)
        If dCod = uTipos(i).dCodigo Then
            OBT = i
        End If
    Next
End Function

Private Sub CargoTipos()
Dim rsTipos As New ADODB.Recordset
Dim i As Long
rsTipos.Open "select * from umtipos order by utcodigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsTipos
    If .EOF Or .BOF Then
        ReDim uTipos(0)
        uTipos(0).dCodigo = 0
        uTipos(0).dDescripcion = "Sin Tipos"
        cboTipo.AddItem uTipos(0).dDescripcion
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            ReDim Preserve uTipos(i)
            uTipos(i).dCodigo = !utcodigo
            uTipos(i).dDescripcion = !abreviatura
            cboTipo.AddItem !abreviatura
            .MoveNext
        Next
    End If
    cboTipo.ListIndex = 0
End With
End Sub
