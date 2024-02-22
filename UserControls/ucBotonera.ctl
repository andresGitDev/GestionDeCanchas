VERSION 5.00
Begin VB.UserControl ucBotonera 
   Alignable       =   -1  'True
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   7440
   Begin VB.Frame fraBuscarYa 
      Height          =   555
      Left            =   4290
      TabIndex        =   17
      Top             =   0
      Width           =   1470
      Begin VB.CommandButton cmdIr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1005
         Picture         =   "ucBotonera.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Busqueda rapida"
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox txtBuscarYa 
         Height          =   315
         Left            =   75
         TabIndex        =   13
         Top             =   165
         Width           =   915
      End
   End
   Begin VB.Frame fraMov 
      Height          =   555
      Left            =   1620
      TabIndex        =   15
      Top             =   0
      Width           =   2940
      Begin VB.CommandButton cmdPrimero 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ucBotonera.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Primero"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdUltimo 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ucBotonera.ctx":06D4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ultimo"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdSiguiente 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1305
         Picture         =   "ucBotonera.ctx":081E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Siguiente"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdAnterior 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   690
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ucBotonera.ctx":0968
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Anterior"
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.Frame fraAbm 
      Height          =   975
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   7320
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         DisabledPicture =   "ucBotonera.ctx":0AB2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6360
         Picture         =   "ucBotonera.ctx":137C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         DisabledPicture =   "ucBotonera.ctx":1C46
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4560
         Picture         =   "ucBotonera.ctx":2510
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         DisabledPicture =   "ucBotonera.ctx":2DDA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5460
         Picture         =   "ucBotonera.ctx":36A4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         DisabledPicture =   "ucBotonera.ctx":3F6E
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2760
         Picture         =   "ucBotonera.ctx":4838
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         DisabledPicture =   "ucBotonera.ctx":5102
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   60
         Picture         =   "ucBotonera.ctx":59CC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         DisabledPicture =   "ucBotonera.ctx":6296
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3660
         Picture         =   "ucBotonera.ctx":6B60
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         DisabledPicture =   "ucBotonera.ctx":742A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   960
         Picture         =   "ucBotonera.ctx":7CF4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   900
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         DisabledPicture =   "ucBotonera.ctx":85BE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1860
         Picture         =   "ucBotonera.ctx":8E88
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.CheckBox chkAceptarBorraControles 
      Caption         =   "AceptarBorraControles"
      Height          =   285
      Left            =   5835
      TabIndex        =   16
      Top             =   210
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2340
   End
End
Attribute VB_Name = "ucBotonera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ucbBoton
    ucbNuevo
    ucbBuscar
    ucbImprimir
    ucbModificar
    ucbEliminar
    ucbAceptar
    ucbCancelar
    ucbSalir
    ucbPrimero
    ucbUltimo
    ucbAnterior
    ucbSiguiente
End Enum

Public Enum ucbEstado
    ucbocioso       '= 0
    ucbMostrando    '= 1
    ucbEditando     '= 2
   'ucbAgregando = 4            // La opcion de desglose AceptarAlta y AceptarModi fue mejor
End Enum

Private Enum ucbAM
    ucbNone
    ucbAlta
    ucbModi
End Enum

Public Event nuevo()
Public Event buscar()
Public Event Imprimir()
Public Event modificar()
Public Event eliminar()
Public Event aceptar()
Public Event Cancelar()
Public Event salir()
Public Event AceptarAlta()
Public Event AceptarModi()
Public Event BuscarYa(que)

Private mMsgConfirmaSalir    As String
Private mMsgConfirmaEliminar As String
Private mMsgConfirmaCancelar As String
'Private mAceptarBorraControles As Boolean

Private mRs                  As ADODB.Recordset
Private mMovimiento          As Boolean
Private mEstado              As ucbEstado
Private mAltaModi            As ucbAM


Public Event Clic(cual As ucbBoton, estado As ucbEstado)
Public Event HabilitarEdicion(sino As Boolean)
Public Event HabilitarEdicionAM(sino As Boolean, SiNoAlta As Boolean)
'Public Event HabilitarEdicionAlta(sino As Boolean)
'Public Event HabilitarEdicionModi(sino As Boolean)
Public Event BorrarControles()

Public Event SeMovio()

Private Const MaxHeightBOTONES = 375

Public Sub init(buscar As Boolean, nuevo As Boolean, modificar As Boolean, Imprimir As Boolean, eliminar As Boolean, Optional strSqlMov As String, Optional cnxn As ADODB.Connection, Optional BuscarYa As Boolean)
    cmdbuscar.Visible = buscar
    cmdnuevo.Visible = nuevo
    cmdImprimir.Visible = Imprimir
    cmdmodificar.Visible = modificar
    cmdeliminar.Visible = eliminar
    
    cmdAceptar.Visible = nuevo Or modificar
    cmdcancelar.Visible = nuevo Or modificar
    
    cmdSalir.Visible = True
    'txtBuscarYa.Visible = BuscarYa
    fraBuscarYa.Visible = BuscarYa
    
    mAltaModi = ucbNone
    mEstado = ucbocioso
    mMovimiento = (strSqlMov > "")
    
    ReverBotonesAbm True
    
    If mMovimiento Then
        Set mRs = New ADODB.Recordset
        mRs.Open strSqlMov, cnxn, adOpenStatic, adLockReadOnly
        'Enabled = True
        ReverBotonesMov True, True, True, True
    Else
        fraMov.Visible = False
    End If
    
    'mMsgConfirmaSalir = "¿ Salir ?"
    'mMsgConfirmaEliminar = "¿ Elimina ?"
    'mMsgConfirmaCancelar = "¿ Cancela ?"
'    mAceptarBorraControles = True
    'UserControl.BackColor = UserControl.ParentControls(0).BackColor
    'fraAbm.BackColor = UserControl.BackColor
    'fraMov.BackColor = UserControl.BackColor
    'tamanios
    
    RaiseEvent BorrarControles
End Sub

Public Sub AceptarOk(Optional strFind As String)
    On Error Resume Next
    mEstado = ucbMostrando
    mAltaModi = ucbNone
    ReverBotonesAbm True
    
    If Not mRs Is Nothing Then mRs.Requery
    If strFind > "" Then
        Find (strFind)
    ElseIf mABC() Then
        mEstado = ucbocioso
        ReverBotonesAbm True
        RaiseEvent BorrarControles ' No me gusta, pero es L&F BS
    End If
End Sub
Public Sub BuscarOK(Optional strFind As String)
    On Error Resume Next
    If strFind > "" Then Find (strFind)
    
    mEstado = ucbMostrando
    mAltaModi = ucbNone
    txtBuscarYa = ""
    
    ReverBotonesMov True, True, True, True
    ReverBotonesAbm
End Sub
Public Sub EliminarOK()
    On Error Resume Next
    mEstado = ucbocioso
    mAltaModi = ucbNone
    ReverBotonesAbm True
    RaiseEvent BorrarControles
    If Not mRs Is Nothing Then mRs.Requery
End Sub

Public Sub Requery()
    mRs.Requery
    'ReverBotonesMov
End Sub

Public Sub CancelarEdicion()
    If mEstado = ucbEditando Then cancelaEdicion
End Sub

Public Function Find(txtFind As String) 'As Boolean
    mRs.MoveFirst
    mRs.Find txtFind ', , adSearchForward, 0
End Function

Public Property Get estado() As ucbEstado
    estado = mEstado
End Property

Public Property Let BackColor(cual)
'    UserControl.BackColor = cual
'    fraMov.BackColor = cual
'    fraAbm.BackColor = cual
'    fraBuscarYa.BackColor = cual
End Property

Public Property Get BackColor()
    BackColor = UserControl.BackColor
End Property

Public Property Let MsgConfirmaSalir(str As String)
    mMsgConfirmaSalir = str
    PropertyChanged "MsgConfirmaSalir"
End Property
Public Property Get MsgConfirmaSalir() As String
    MsgConfirmaSalir = mMsgConfirmaSalir
End Property
Public Property Let MsgConfirmaEliminar(str As String)
    mMsgConfirmaEliminar = str
    PropertyChanged "MsgConfirmaEliminar"
End Property
Public Property Get MsgConfirmaEliminar() As String
    MsgConfirmaEliminar = mMsgConfirmaEliminar
End Property
Public Property Let MsgConfirmaCancelar(str As String)
    mMsgConfirmaCancelar = str
    PropertyChanged "MsgConfirmaCancelar"
End Property
Public Property Get MsgConfirmaCancelar() As String
    MsgConfirmaCancelar = mMsgConfirmaCancelar
End Property
Public Property Let CaptionEliminar(str As String)
    cmdeliminar.caption = str
    PropertyChanged "CaptionEliminar"
End Property
Public Property Get CaptionEliminar() As String
    CaptionEliminar = cmdeliminar.caption
End Property
Public Property Let CaptionImprimir(str As String)
    cmdImprimir.caption = str
    PropertyChanged "CaptionImprimir"
End Property
Public Property Get CaptionImprimir() As String
    CaptionImprimir = cmdImprimir.caption
End Property
Public Property Let AceptarBorraControles(sino As Boolean)
    mABC sino
    PropertyChanged "AceptarBorraControles"
End Property
Public Property Get AceptarBorraControles() As Boolean
    AceptarBorraControles = mABC 'mAceptarBorraControles
End Property


Private Sub cmdIr_Click()
    If Trim$(txtBuscarYa) > "" Then
        RaiseEvent BuscarYa(Trim$(txtBuscarYa))
    End If
End Sub

Private Sub txtBuscarYa_GotFocus()
    txtBuscarYa.SelStart = 0
    txtBuscarYa.SelLength = Len(txtBuscarYa)
End Sub
Private Sub txtBuscarYa_LostFocus()
    If Trim$(txtBuscarYa) > "" Then
        cmdIr.SetFocus
    End If
End Sub


Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 And UserControl.ActiveControl.Name = "txtBuscarYa" And Trim$(txtBuscarYa) > "" Then
        cmdIr.SetFocus
'        RaiseEvent BuscarYa(Trim$(txtBuscarYa))
    End If
End Sub

'Private Sub UserControl_Initialize()
'    On Error Resume Next
'    UserControl.BackColor = UserControl.ParentControls(0).BackColor
'End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim abc As Boolean
    On Error Resume Next
     UserControl.enabled = PropBag.ReadProperty("Enabled", True)
     mMsgConfirmaSalir = PropBag.ReadProperty("MsgConfirmaSalir", "¿ Cerrar formulario ?")
     mMsgConfirmaEliminar = PropBag.ReadProperty("MsgConfirmaEliminar", "")
     mMsgConfirmaCancelar = PropBag.ReadProperty("MsgConfirmaCancelar", "")
     cmdeliminar.caption = PropBag.ReadProperty("CaptionEliminar", "Eliminar")
     cmdImprimir.caption = PropBag.ReadProperty("CaptionImprimir", "&Imprimir")
     abc = PropBag.ReadProperty("AceptarBorraControles", True)
     mABC abc
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MsgConfirmaSalir", mMsgConfirmaSalir, "¿ Cerrar formulario ?"
    PropBag.WriteProperty "MsgConfirmaEliminar", mMsgConfirmaEliminar
    PropBag.WriteProperty "MsgConfirmaCancelar", mMsgConfirmaCancelar
    PropBag.WriteProperty "CaptionEliminar", cmdeliminar.caption, "Eliminar"
    PropBag.WriteProperty "CaptionImprimir", cmdImprimir.caption, "&Imprimir"
    PropBag.WriteProperty "AceptarBorraControles", mABC(), True
End Sub

Private Sub cmdanterior_Click()
    RaiseEvent Clic(ucbAnterior, mEstado)
    mRs.MovePrevious
    If mRs.BOF Then
        cmdPrimero_Click
    Else
        ReverBotonesMov True, True, True, True
        EventoCambio
    End If
End Sub
Private Sub cmdPrimero_Click()
    RaiseEvent Clic(ucbPrimero, mEstado)
    mRs.MoveFirst
    ReverBotonesMov False, False, True, True
    EventoCambio
End Sub
Private Sub cmdsiguiente_Click()
    RaiseEvent Clic(ucbSiguiente, mEstado)
    mRs.MoveNext
    If mRs.EOF Then
        cmdUltimo_Click
    Else
        ReverBotonesMov True, True, True, True
        EventoCambio
    End If
End Sub
Private Sub cmdUltimo_Click()
    RaiseEvent Clic(ucbUltimo, mEstado)
    mRs.MoveLast
    ReverBotonesMov True, True, False, False
    EventoCambio
End Sub


Private Sub ReverBotonesAbm(Optional conFoco As Boolean)
    cmdbuscar.enabled = (mEstado = ucbocioso Or mEstado = ucbMostrando)
    cmdnuevo.enabled = (mEstado = ucbocioso Or mEstado = ucbMostrando)
    cmdImprimir.enabled = (mEstado = ucbMostrando)
    cmdmodificar.enabled = (mEstado = ucbMostrando)
    cmdeliminar.enabled = (mEstado = ucbMostrando) '  = ucbeditando)
    cmdAceptar.enabled = (mEstado = ucbEditando)
    cmdcancelar.enabled = (mEstado = ucbEditando) 'Or mEstado = ucbAgregando)
    cmdSalir.enabled = True
    
    fraMov.enabled = (mEstado <> ucbEditando)
    fraBuscarYa.enabled = (mEstado <> ucbEditando)
    
    If conFoco Then enfocar
    
    'puedo usar un unico evento
    RaiseEvent HabilitarEdicion(mEstado = ucbEditando) ' or mEstado = ucbAgregando
    RaiseEvent HabilitarEdicionAM((mEstado = ucbEditando), (mEstado = ucbEditando) And (mAltaModi = ucbAlta))    ' or mEstado = ucbAgregando
End Sub

Private Sub enfocar()
    On Error Resume Next
    If cmdAceptar.enabled Then
        cmdAceptar.SetFocus
    ElseIf cmdnuevo.enabled Then
        cmdnuevo.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    RaiseEvent Clic(ucbAceptar, mEstado)
    
    If mAltaModi = ucbAlta Then RaiseEvent AceptarAlta
    If mAltaModi = ucbModi Then RaiseEvent AceptarModi
    RaiseEvent aceptar
End Sub

Private Sub cmdBuscar_Click()
    Dim Res As String
    
    RaiseEvent Clic(ucbBuscar, mEstado)
    RaiseEvent buscar
End Sub

Private Sub cmdCancelar_Click()
    If mMsgConfirmaCancelar > "" Then
        If Not confirma(mMsgConfirmaCancelar) Then Exit Sub
    End If
    cancelaEdicion
End Sub

Private Sub cancelaEdicion()
    mAltaModi = ucbNone

    RaiseEvent Cancelar
    If mABC() Then
        mEstado = ucbocioso
        ReverBotonesAbm True
        RaiseEvent Clic(ucbCancelar, mEstado)
        RaiseEvent BorrarControles ' No a todos les gusta esto
    Else
        mEstado = ucbMostrando
        ReverBotonesAbm True
        RaiseEvent Clic(ucbCancelar, mEstado)
        RaiseEvent SeMovio ' Mentira, pero debo recargar form por si modifico
    End If

End Sub

Private Sub cmdeliminar_Click()
    If mMsgConfirmaEliminar > "" Then
        If Not confirma(mMsgConfirmaEliminar) Then Exit Sub
    End If
    
    RaiseEvent Clic(ucbEliminar, mEstado)
    RaiseEvent eliminar
End Sub

Private Sub cmdImprimir_Click()
    mEstado = ucbMostrando
    ReverBotonesAbm
    RaiseEvent Clic(ucbImprimir, mEstado)
    RaiseEvent Imprimir
End Sub

Private Sub cmdmodificar_Click()
    mEstado = ucbEditando
    mAltaModi = ucbModi
    ReverBotonesAbm True
    RaiseEvent Clic(ucbModificar, mEstado)
    RaiseEvent modificar
End Sub

Private Sub cmdnuevo_Click()
    mAltaModi = ucbAlta
    RaiseEvent BorrarControles
    mEstado = ucbEditando
    ReverBotonesAbm True
    RaiseEvent Clic(ucbNuevo, mEstado)
    RaiseEvent nuevo
    
End Sub

Private Sub cmdSalir_Click()
    If mEstado = ucbEditando Then
        MsgBox "Para salir primero cancele el proceso.", vbInformation
        Exit Sub
    End If

    If mMsgConfirmaSalir > "" Then
        If Not confirma(mMsgConfirmaSalir) Then Exit Sub
    End If
    ' No estoy seguro de esto...
    mEstado = ucbocioso
    mAltaModi = ucbNone
    RaiseEvent BorrarControles
    '
    ReverBotonesAbm
    RaiseEvent Clic(ucbSalir, mEstado)
    RaiseEvent salir
End Sub

Private Sub tamanios()
    On Error Resume Next
    Dim uh As Long
    
    uh = UserControl.Height
    
    fraMov.Top = 0
    fraMov.Height = uh / 2
    fraMov.Visible = mMovimiento
'    fraBuscarYa.Height = uh / 2
'    cmdIr.Height = fraBuscarYa.Height
'    txtBuscarYa.Height = fraBuscarYa.Height
    
    If mMovimiento Or txtBuscarYa.Visible Then
        fraAbm.Top = uh / 2
        fraAbm.Height = uh / 2
''        hBotones uh / 2
    Else
        fraAbm.Top = maximo(uh - MaxHeightBOTONES, 0) '0
        fraAbm.Height = maximo(uh, MaxHeightBOTONES) 'uh
 ''       hBotones uh
    End If
        
'    If mMovimiento Then
'        fraMov.Visible = True
'        fraMov.Top = 0
'        fraAbm.Top = uh / 2
'        fraMov.Height = uh / 2
'        fraAbm.Height = uh / 2
'        hBotones uh / 2
'    Else
'        fraMov.Visible = False
'        fraMov.Top = 0
'        fraAbm.Top = maximo(uh - MaxHeightBOTONES, 0) '0
'        fraAbm.Height = maximo(uh, MaxHeightBOTONES) 'uh
'        hBotones uh
'    End If
End Sub
Private Function maximo(a, b)
    maximo = IIf(a > b, a, b)
End Function
'Private Sub hBotones(h As Long)
''    On Error Resume Next ' si no were un button
''    Dim boton As CommandButton
''    For Each boton In UserControl.Controls
''        boton.Height = h
''    Next
'End Sub

Private Sub ReverBotonesMov(pri As Boolean, ant As Boolean, sig As Boolean, ult As Boolean)
    On Error Resume Next
    cmdprimero.enabled = Not mRs.BOF And pri
    cmdanterior.enabled = Not mRs.BOF And ant
    cmdultimo.enabled = Not mRs.EOF And sig
    cmdsiguiente.enabled = Not mRs.EOF And ult
End Sub


Private Sub UserControl_Terminate()
    Set mRs = Nothing
End Sub

Public Property Get rs() As ADODB.Recordset
    Set rs = mRs
End Property

Private Sub EventoCambio()
    mEstado = ucbMostrando
    ReverBotonesAbm
    RaiseEvent SeMovio
End Sub

Private Function confirma(sMsg As String) As Boolean
    confirma = (MsgBox(sMsg, vbQuestion + vbYesNo, "Confirmacion") = vbYes)
End Function

Private Function mABC(Optional que) As Boolean
    If Not IsMissing(que) Then
        chkAceptarBorraControles.Value = IIf(que, vbChecked, vbUnchecked)
    End If
    mABC = (chkAceptarBorraControles.Value = vbChecked)
End Function


'Public Property Let estado(cual As ucbEstado)
'    mEstado = cual
'    ReverBotonesAbm
'End Property
'Public Sub EstadoMostrando()
'    mEstado = ucbMostrando
'    ReverBotonesAbm
'End Sub
'Public Sub EstadoVacio()
'    mEstado = ucbOcioso
'    ReverBotonesAbm
'End Sub


' 17/8/4 start
'   /8/4 ok
' 20/8/4 cambio forma de aviso, con metodos buscarOk y aceptarOK
' 23/8/4 find() y cambio buscar() para sincr RS
'        ampliac comentarios
' 24/8/4 baja de metodos
'        elimino encontro() y strbuscar, complica el esquema mental
'        ahora uso sist coherente para mantener estado
' 26/8/4 cosmetica
'        Nombre metodos, enum, parametros opcionales
' 3/9/4  Fix msgEliminar
' 13/9/4 .BackColor
'        .align
' 1/10/4 fallaba find, ahora hace movefirst antes
'        fix bt modif si MOV'
'14/10/4 Agregue Evento BorrarControles
'        Mod en eliminar requery autom
'22/10/4 cmdSalir.cancel true
'                   no se como poner cancel true al uc desde aca
'        optional find en aceptarOk()
'25-10-4 hab botones mov desp de buscarOk
'        Boton <Aceptar> = default
'        Aceptar Discriminado _AceptarAlta _AceptarModi
'19-11-4 AceptarOK dispara BorrarControles
'18-1-5  Fix: hard declare adodb.connection
'14-2-5  Fix: habilitacion botones desp de borrar control: relacionado al add:
'        Add: propiedad AceptarBorraControles.  true (BS L&F) default
'15-2-6  Bolus: toma color de fondo
'               altura max de botones, para meterle controles dentro sin q se borren
'22-3-5  txtbox busqueda rapida con evento buscarya(que)
'29-3-5 UserControl_AccessKeyPress on error
'4/4/5  Anule cambio tamaño botones automatico (hBotones())
'7/4/5  manejo focos
'18/8/6  HabilitarEdicionAM

