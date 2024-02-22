VERSION 5.00
Begin VB.Form FrmCtasContables 
   Caption         =   "Cuentas contables"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "FrmCtasContables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucCoDe uSumariza 
      Height          =   315
      Left            =   1320
      TabIndex        =   27
      Top             =   2940
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.CheckBox chkmonNo 
      Alignment       =   1  'Right Justify
      Caption         =   "NO"
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
      Height          =   255
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox chkmonSi 
      Alignment       =   1  'Right Justify
      Caption         =   "SI"
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
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtrenglon 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   23
      Top             =   2520
      Width           =   2160
   End
   Begin VB.CheckBox chksaltoNo 
      Alignment       =   1  'Right Justify
      Caption         =   "NO"
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
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox chksaltoSi 
      Alignment       =   1  'Right Justify
      Caption         =   "SI"
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
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox chkimpNo 
      Alignment       =   1  'Right Justify
      Caption         =   "NO"
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
      Height          =   255
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox chkimpSi 
      Alignment       =   1  'Right Justify
      Caption         =   "SI"
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
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Width           =   855
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
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   345
      Width           =   1815
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
      Picture         =   "FrmCtasContables.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   3585
      Width           =   675
   End
   Begin VB.CommandButton cmdsiguiente 
      Height          =   615
      Left            =   3330
      Picture         =   "FrmCtasContables.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Siguiente"
      Top             =   3585
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      Height          =   615
      Left            =   3945
      Picture         =   "FrmCtasContables.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ultimo"
      Top             =   3585
      Width           =   615
   End
   Begin VB.CommandButton cmdprimero 
      Height          =   615
      Left            =   2040
      Picture         =   "FrmCtasContables.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Primero"
      Top             =   3585
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
      Top             =   4425
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
      Top             =   4425
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
      Top             =   4425
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
      Top             =   4425
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
      Top             =   4425
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
      Top             =   4425
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
      Top             =   4425
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Monetaria"
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
      Left            =   4920
      TabIndex        =   26
      Top             =   1320
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   735
      Left            =   4680
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Salto"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   735
      Left            =   2520
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Imputable"
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
      TabIndex        =   19
      Top             =   1320
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   735
      Left            =   360
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Renglones en blanco despues de Imprimir :"
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
      TabIndex        =   16
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Sumariza :"
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
      TabIndex        =   15
      Top             =   2940
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3330
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
Attribute VB_Name = "FrmCtasContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 11/8/4


Dim Ope As String
Dim idCta As Long
Dim rscta As New ADODB.Recordset

Private Sub cmdAceptar_Click()

    'Dim fecha As Variant
    Dim rs As New ADODB.Recordset
    Dim IMPUTABLE As Long
    Dim salto As Long
    Dim MONETARIA As Long
    Dim dat1 As Long
    Dim dat2 As String


    Call HabilitoControles(False, False, False, True, False, True)
    If Trim(txtDescripcion) = "" Then
        MsgBox "Debe cargar la descripcion", 48, "Atencion"
        txtDescripcion.SetFocus
        Exit Sub
    Else
        If Ope = "A" Then
            
            rs.Open "Select * from Cuentas where cuenta='" & uSumariza.codigo & Trim(txtCodigo) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rs.EOF = True And rs.BOF = True Then
            Else
                If Not rs.EOF Then
                    MsgBox "El codigo de cuenta ya existe", 48, "Atencion"
                    Exit Sub
                End If
            End If
            'rs.Close
            Set rs = Nothing
            Set rs = Nothing
            If chksaltoSi.Value = 1 Then
                salto = 1
            Else
                salto = 0
            End If
            If chkimpSi.Value = 1 Then
                IMPUTABLE = 1
            Else
                IMPUTABLE = 0
            End If
            If chkmonSi.Value = 1 Then
                MONETARIA = 1
            Else
                MONETARIA = 0
            End If
            If uSumariza.codigo = "" Then
               dat1 = 0
               dat2 = ""
            Else
                dat1 = uSumariza.codigo
                dat2 = uSumariza.codigo
            End If
'            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            'daTaenvironment1.dbo_CUENTA "A", Trim(txtCodigo), Trim(txtDescripcion), imputable, salto, monetaria, Val(Trim(txtrenglon)), ObtenerCodigoS("cuentas", Trim(txtsumariza)), fecha, UsuarioSistema!Codigo, 0, 0
            'DataEnvironment1.dbo_CUENTA "A", dat2 & Trim(txtCodigo), Trim(txtCodigo), Trim(txtDescripcion), IMPUTABLE, SALTO, MONETARIA, Val(Trim(txtrenglon)), dat1, Date, UsuarioSistema!CODIGO, 0, 0
            ABM_Cuentas "A", idCta, dat2 & Trim(txtCodigo), Trim(txtCodigo), Trim(txtDescripcion), IMPUTABLE, salto, MONETARIA, Val(Trim(txtRenglon)), dat1, Date, Date, UsuarioActual, 0
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
                If chksaltoSi.Value = 1 Then
                    salto = 1
                Else
                    salto = 0
                End If
                If chkimpSi.Value = 1 Then
                    IMPUTABLE = 1
                Else
                    IMPUTABLE = 0
                End If
                If chkmonSi.Value = 1 Then
                    MONETARIA = 1
                Else
                    MONETARIA = 0
                End If
                If uSumariza.codigo = "" Then
                   dat1 = 0
                   dat2 = txtCodigo '""
                Else
                    dat1 = uSumariza.codigo
                    dat2 = uSumariza.codigo & Trim(txtCodigo)
                End If
                'fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
                'daTaenvironment1.dbo_CUENTA "M", Val(txtCodigo), Trim(txtDescripcion), imputable, salto, monetaria, Trim(txtrenglon), ObtenerCodigoS("cuentas", Trim(txtsumariza)), 0, 0, 0, 0
                'DataEnvironment1.dbo_CUENTA "M", dat2, Val(txtCodigo), Trim(txtDescripcion), IMPUTABLE, SALTO, MONETARIA, Trim(txtrenglon), dat1, 0, 0, 0, 0
                DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Cuentas", UsuarioSistema!codigo, Date, Time, "M"
                If ABM_Cuentas("M", idCta, dat2, Val(txtCodigo), Trim(txtDescripcion), IMPUTABLE, salto, MONETARIA, Trim(txtRenglon), dat1, Date, Date, UsuarioActual, 0) Then
                    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
                Else
                    MsgBox "La Operación no se ha realizado con éxito.verifique el numero de cuenta.", vbCritical, "Atencion"
                End If
            End If
        End If
    End If
    LimpioTxt
    HabilitoTxt (True)
End Sub
Sub CargoRegistro()
    txtCodigo = rscta!cod 'rscta!Cuenta 'rscta.Fields(2)
    txtDescripcion = rscta!DESCRIPCION
    If rscta!IMPUTABLE = True Then
        chkimpSi.Value = 1
        chkimpNo.Value = 0
    Else
        chkimpNo.Value = 1
        chkimpSi.Value = 0
    End If
    If rscta!salto = True Then
        chksaltoSi.Value = 1
        chksaltoNo.Value = 0
    Else
        chksaltoNo.Value = 1
        chksaltoSi.Value = 0
    End If
    If rscta!MONETARIA = True Then
        chkmonSi.Value = 1
        chkmonNo.Value = 0
    Else
        chkmonNo.Value = 1
        chkmonSi.Value = 0
    End If
    If Not IsNull(rscta!RENGLON) Then
        txtRenglon = rscta!RENGLON
    Else
        txtRenglon = ""
    End If
    
    uSumariza.codigo = rscta!SUMARIZA
    'If Not IsNull(rscta!sumariza) Then
    '    txtsumariza = ObtenerDescripcionS("cuentas", rscta!sumariza)
    'Else
    '    txtsumariza = "0"
    'End If
End Sub
Private Sub cmdanterior_Click()
    rscta.MovePrevious
    If Not rscta.BOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

'Private Sub cmdayudactas_Click()
''    FrmHelp.Show
''    CargarHelp "cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
''    FrmHelp.Tag = "sumariza"
'    Dim resu As String
'    resu = frmBuscar.MostrarCodigoDescripcionActivo("cuentas")
'    If resu > "" Then
'        txtsumariza = frmBuscar.resultado(2)
'    End If
'End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
    Dim resu As String, an
    an = Array(0, 1000, 3000)
    'resu = frmBuscar.MostrarCodigoDescripcionActivo("cuentas")
    resu = frmBuscar.MostrarSql("select id as [i],cuenta as [ Cuenta                    ], descripcion [ Descripcion                                ] from cuentas where activo = 1 order by cuenta ", an)
    If resu > "" Then
        idCta = resu
        txtCodigo = frmBuscar.resultado(2)
        CargarDatos
        HabilitoControles True, False, True, False, True, False
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitoControles(False, False, False, True, False, True)
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoBotonesMoverse(False, False, False, False)
End Sub

Private Sub cmdeliminar_Click()
'Dim fecha As Variant
'Dim mensaje As String

'    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
'    If mensaje = 6 Then
    If confirma("Esta seguro de querer eliminar este registro") Then
'        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'        DataEnvironment1.dbo_CUENTA "B", uSumariza.CODIGO & Trim(txtCodigo), Trim(txtCodigo), "", 0, 0, 0, 0, 0, 0, 0, UsuarioSistema!CODIGO, Date
        ABM_Cuentas "B", idCta, uSumariza.codigo & Trim(txtCodigo), Trim(txtCodigo), "", 0, 0, 0, 0, 0, Date, Date, 0, UsuarioSistema!codigo
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtCodigo), "Cuentas", UsuarioSistema!codigo, Date, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
    End If
End Sub

Private Sub cmdPrimero_Click()
    rscta.MoveFirst
    CargoRegistro
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdsiguiente_Click()
    rscta.MoveNext
    If Not rscta.EOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rscta.MoveLast
    CargoRegistro
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    If rscta.State = 1 Then
        rscta.Close
        Set rscta = Nothing
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
'    txtdescripcion.SelStart = 0
'    txtdescripcion.SelLength = Len(txtdescripcion.Text)
    PintoFocoActivo
End Sub
'Private Sub txtsumariza_GotFocus()
''    txtsumariza.SelStart = 0
''    txtsumariza.SelLength = Len(txtsumariza.Text)
'    PintoFocoActivo
'End Sub

Private Sub txtrenglon_GotFocus()
    'txtrenglon.SelStart = 0
    'txtrenglon.SelLength = Len(txtrenglon.Text)
    PintoFocoActivo
End Sub
Private Sub txtcodigo_GotFocus()
'    txtcodigo.SelStart = 0
'    txtcodigo.SelLength = Len(txtcodigo.Text)
    PintoFocoActivo
End Sub
'Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtsumariza_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtrenglon_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'
Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtDescripcion.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
Dim rscta As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    HabilitoTxt (False)
    txtCodigo.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtCodigo = ""
    txtDescripcion = ""
    'txtsumariza = ""
    uSumariza.codigo = ""
    txtRenglon = "0"
    chkimpSi.Value = 0
    chkimpNo.Value = 0
    chkmonSi.Value = 0
    chkmonNo.Value = 0
    chksaltoSi.Value = 0
    chksaltoNo.Value = 0
End Sub
Sub HabilitoTxt(habilito As Boolean)
    ' OJO habilito bloquea, not habilito habilita ' Crazy, dont you think?
    txtCodigo.Locked = habilito
    txtDescripcion.Locked = habilito
    txtRenglon.Locked = habilito
    'txtsumariza.Locked = habilito
    uSumariza.enabled = Not habilito
    chkimpSi.enabled = Not habilito
    chkimpNo.enabled = Not habilito
    chkmonSi.enabled = Not habilito
    chkmonNo.enabled = Not habilito
    chksaltoSi.enabled = Not habilito
    chksaltoNo.enabled = Not habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    
    cmdcancelar.enabled = hab1
    cmdAceptar.enabled = hab2
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
 '   Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioTxt
    'uSumariza.ini "select descripcion from Cuentas where codigo = '###' and activo = 1", "select cuenta as [ Cuenta       ], descripcion as [ Descripcion               ] from Cuentas where activo = 1 ", True
    uSumariza.ini "select descripcion from Cuentas where cuenta = '###' and activo = 1", "select cuenta as [ Cuenta       ], descripcion as [ Descripcion               ] from Cuentas where activo = 1 ", True
    Call HabilitoControles(False, False, False, True, False, True)
    HabilitoBotonesMoverse False, False, False, False
End Sub
Sub CargarDatos()
    
    If rscta.State = 1 Then
        rscta.Close
        Set rscta = Nothing
    End If
    rscta.Open "select *,_codigo as cod from cuentas where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    
    If Not rscta.EOF Then
        rscta.MoveFirst
        rscta.Find "cuenta=" & str(Trim(txtCodigo))
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If

End Sub



' 11/8/4 Lito
'   inhibo mov al cargar
'   busq  frmHelp  pasa a frmBUSCAR
'18/2/5
'   fix fecha string >> con date
'   fix sumariza     >> con uc
'
' known: habilitacion controles, >> cuando tenga ganas le encajo mi menu
Private Sub uSumariza_cambio(codigo As Variant)

End Sub
