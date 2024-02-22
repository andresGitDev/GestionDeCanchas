VERSION 5.00
Begin VB.Form FrmABMUnidadFactor 
   Caption         =   "ABM de factor de conversion"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "FrmABMUnidadFactor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboUFinal 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1890
      Width           =   3000
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3000
   End
   Begin VB.TextBox txtFactor 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1260
      Width           =   3000
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
      Left            =   3840
      Picture         =   "FrmABMUnidadFactor.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   870
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
      DisabledPicture =   "FrmABMUnidadFactor.frx":0E54
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
      Left            =   4710
      Picture         =   "FrmABMUnidadFactor.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2235
      Width           =   870
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
      DisabledPicture =   "FrmABMUnidadFactor.frx":1FE8
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
      Left            =   4710
      Picture         =   "FrmABMUnidadFactor.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   675
      Width           =   870
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
      DisabledPicture =   "FrmABMUnidadFactor.frx":317C
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
      Left            =   4710
      Picture         =   "FrmABMUnidadFactor.frx":3A46
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3015
      Width           =   870
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      DisabledPicture =   "FrmABMUnidadFactor.frx":4310
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
      Left            =   2100
      Picture         =   "FrmABMUnidadFactor.frx":4BDA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3015
      Width           =   870
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      DisabledPicture =   "FrmABMUnidadFactor.frx":54A4
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
      Left            =   2970
      Picture         =   "FrmABMUnidadFactor.frx":5D6E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3015
      Width           =   870
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      DisabledPicture =   "FrmABMUnidadFactor.frx":6638
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
      Left            =   4710
      Picture         =   "FrmABMUnidadFactor.frx":6F02
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1455
      Width           =   870
   End
   Begin VB.CommandButton cmdprimero 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1920
      Picture         =   "FrmABMUnidadFactor.frx":77CC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Primero"
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton cmdultimo 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2820
      Picture         =   "FrmABMUnidadFactor.frx":7916
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton cmdsiguiente 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2520
      Picture         =   "FrmABMUnidadFactor.frx":7A60
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Siguiente"
      Top             =   2295
      Width           =   285
   End
   Begin VB.CommandButton cmdanterior 
      Enabled         =   0   'False
      Height          =   390
      Left            =   2220
      Picture         =   "FrmABMUnidadFactor.frx":7BAA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Anterior"
      Top             =   2295
      Width           =   285
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3000
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "(*) Valor por el cual se multiplicara el valor entrante"
      Height          =   330
      Left            =   105
      TabIndex        =   21
      Top             =   3825
      Width           =   4935
   End
   Begin VB.Label Label5 
      Caption         =   "Unidad Final"
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
      Left            =   255
      TabIndex        =   20
      Top             =   1905
      Width           =   1455
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
      Left            =   945
      TabIndex        =   19
      Top             =   1575
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Factor (*)"
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
      Left            =   525
      TabIndex        =   18
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   2760
      Left            =   90
      Top             =   120
      Width           =   4515
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   195
      Width           =   1095
   End
End
Attribute VB_Name = "FrmABMUnidadFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ope As String
Dim rsUnidad As New ADODB.Recordset
Private uTipos() As datCD
Private uMedidas() As datCD

Private Sub cboTipo_Change()
    CargoUnidades uTipos(cboTipo.ListIndex).dCodigo
End Sub

Private Sub cboTipo_Click()
    CargoUnidades uTipos(cboTipo.ListIndex).dCodigo
End Sub

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
    If Trim(txtFactor) = "" Then
        MsgBox "Debe cargar el factor", vbInformation, "Atencion"
        txtFactor.SetFocus
        Exit Sub
    End If
    
    ABMUnidadFactor Ope, s2n(txtCodigo), txtDescripcion, s2n(txtFactor), uMedidas(cboUFinal.ListIndex).dCodigo, uTipos(cboTipo.ListIndex).dCodigo
    LimpioTxt
    HabilitoTxt (True)
End Sub

Private Sub cmdanterior_Click()
    rsUnidad.MovePrevious
    If Not rsUnidad.BOF Then
        txtCodigo = rsUnidad!umcodigo
        txtDescripcion = rsUnidad!caracteristica
        txtFactor = rsUnidad!FACTOR
        cboTipo.ListIndex = OBT(nSinNull(rsUnidad!tipo))
        cboUFinal.ListIndex = OBT(nSinNull(rsUnidad!uMFinal))
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select UFCodigo as [Nro],Caracteristica from umfactor where activo=1 order by ufcodigo")
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
        ABMUnidadFactor "B", s2n(txtCodigo), "", 0, 0, 0
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rsUnidad.MoveFirst
    txtCodigo = rsUnidad!ufcodigo
    txtDescripcion = rsUnidad!caracteristica
    txtFactor = rsUnidad!FACTOR
    cboTipo.ListIndex = OBT(nSinNull(rsUnidad!tipo))
    cboUFinal.ListIndex = OBT(nSinNull(rsUnidad!uMFinal))
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rsUnidad.MoveNext
    If Not rsUnidad.EOF Then
        txtCodigo = rsUnidad!ufcodigo
        txtDescripcion = rsUnidad!caracteristica
        txtFactor = rsUnidad!FACTOR
        cboTipo.ListIndex = OBT(nSinNull(rsUnidad!tipo))
        cboUFinal.ListIndex = OBT(nSinNull(rsUnidad!uMFinal))
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsUnidad.MoveLast
    txtCodigo = rsUnidad!ufcodigo
    txtDescripcion = rsUnidad!caracteristica
    txtFactor = rsUnidad!FACTOR
    cboTipo.ListIndex = OBT(nSinNull(rsUnidad!tipo))
    cboUFinal.ListIndex = OBT(nSinNull(rsUnidad!uMFinal))
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
    txtFactor = ""
End Sub
Sub HabilitoTxt(habilito As Boolean)
    txtDescripcion.Locked = habilito
    txtFactor.Locked = habilito
    cboTipo.Locked = habilito
    cboUFinal.Locked = habilito
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
    CargoTipos
    CargoUnidades uTipos(cboTipo.ListIndex).dCodigo
End Sub

Sub CargarDatos()
    If rsUnidad.State = 1 Then
        Set rsUnidad = Nothing
    End If
    rsUnidad.Open "select * from umfactor where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsUnidad.EOF Then
        rsUnidad.MoveFirst
        rsUnidad.Find "UFCodigo= " & txtCodigo
        txtDescripcion = sSinNull(rsUnidad!caracteristica)
        txtFactor = sSinNull(rsUnidad!FACTOR)
        cboTipo.ListIndex = OBT(nSinNull(rsUnidad!tipo))
        cboUFinal.ListIndex = OBU(nSinNull(rsUnidad!uMFinal))
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub

Private Function ABMUnidadFactor(uOpe As String, uCodigo As Long, uDescrip As String, uFactor As Double, uMFinal As Long, uTipo As Long) As Boolean
On Error GoTo MAL
Dim ABMUM As String
ABMUnidadFactor = True
Select Case uOpe
    Case "A":
        ABMUM = "INSERT INTO UMFACTOR (CARACTERISTICA,FACTOR,UMFINAL,TIPO,ACTIVO) VALUES " _
            & "(" & ssTexto(uDescrip) & "," & x2s(uFactor) & "," & uMFinal & "," & uTipo & ",1)"
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UF guardada", vbInformation, "Factor nuevo"
    Case "M":
        ABMUM = "UPDATE UMFACTOR SET CARACTERISTICA=" & ssTexto(uDescrip) _
            & ",FACTOR=" & x2s(uFactor) _
            & ",UMFINAL=" & uMFinal _
            & ",TIPO=" & uTipo _
            & " where ufcodigo=" & uCodigo
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UF actualizada", vbInformation, "Factor modificado"
    Case "B":
        If MsgBox("¿Desea eliminar por completo el factor?", vbYesNo + vbExclamation) = vbYes Then
            ABMUM = "DELETE FROM UMFACTOR WHERE UFCODIGO=" & uCodigo
        Else
            ABMUM = "UPDATE UMFACTOR  SET ACTIVO=0 WHERE UFCODIGO=" & uCodigo
        End If
        DataEnvironment1.Sistema.Execute ABMUM
        MsgBox "UF eliminada", vbInformation, "Factor borrado"
End Select

Exit Function
MAL:
    MsgBox "Error en carga de UF", vbCritical, "Factor nuevo no guardada"
    ABMUnidadFactor = False
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

Private Function OBU(dCod As Long) As Long
Dim i As Long
OBU = 0
    For i = 0 To UBound(uMedidas)
        If dCod = uMedidas(i).dCodigo Then
            OBU = i
        End If
    Next
End Function

Private Sub CargoTipos()
Dim rsTipos As New ADODB.Recordset
Dim i As Long
rsTipos.Open "select * from umtipos order by utcodigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
cboUFinal.enabled = True
With rsTipos
    If .EOF Or .BOF Then
        ReDim uTipos(0)
        uTipos(0).dCodigo = 0
        uTipos(0).dDescripcion = "Sin Tipos"
        cboTipo.AddItem uTipos(0).dDescripcion
        cboUFinal.enabled = False
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

Private Function CargoUnidades(mTipo As Long)
Dim rsTipos As New ADODB.Recordset
Dim i As Long
rsTipos.Open "select * from unidadesmedida where tipo=" & mTipo & "  order by umcodigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsTipos
cboUFinal.clear
    If .EOF Or .BOF Then
        ReDim uMedidas(0)
        uMedidas(0).dCodigo = 0
        uMedidas(0).dDescripcion = "Sin Medida"
        cboUFinal.AddItem uMedidas(0).dDescripcion
    Else
        ReDim uMedidas(0)
        .MoveFirst
        For i = 0 To .RecordCount - 1
            ReDim Preserve uMedidas(i)
            uMedidas(i).dCodigo = !umcodigo
            uMedidas(i).dDescripcion = !abreviatura
            cboUFinal.AddItem !abreviatura
            .MoveNext
        Next
    End If
    cboUFinal.ListIndex = 0
End With
End Function
