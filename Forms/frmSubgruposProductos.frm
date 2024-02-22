VERSION 5.00
Begin VB.Form frmSubgruposProductos 
   Caption         =   "Subgrupos de Productos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7455
   Icon            =   "frmSubgruposProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      Height          =   1560
      Left            =   0
      TabIndex        =   0
      Top             =   1530
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2752
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1965
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   4680
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
      Left            =   1965
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   345
      Width           =   975
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
      Left            =   405
      TabIndex        =   4
      Top             =   840
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
      Left            =   435
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1290
      Left            =   120
      Top             =   120
      Width           =   6720
   End
End
Attribute VB_Name = "frmSubgruposProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '   22/10/4
' Lito Explicit
' Reemplaza el viejo frm


'Form
Private Sub Form_Load()
'    CentrarMe Me
    ucMenu.init True, True, True, False, True, "select * from SubGruposProducto where activo=1 order by codigo", DataEnvironment1.Sistema
    ucMenu.MsgConfirmaEliminar = "Elimina este Grupo ? "
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

'Controls
Private Sub txtcodigo_GotFocus()
    GotFocusPinto txtCodigo
End Sub
Private Sub txtDescripcion_GotFocus()
    GotFocusPinto Txtdescripcion
End Sub


'--------------------------------- MENU -------------------------
Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    If FaltaDescr Then Exit Sub
    If YaTa() Then Exit Sub
    DataEnvironment1.dbo_SUBGRUPO "A", Trim(txtCodigo), Trim(Txtdescripcion), Date, UsuarioSistema!codigo, 0, 0
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk '"codigo = '" & txtCodigo & "'"
    GoTo fin
ufaErr:
    ufa "err al grabar alta", Me.Name ', Err
fin:
End Sub

Private Sub ucMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    If FaltaDescr Then Exit Sub

    DataEnvironment1.dbo_SUBGRUPO "M", Trim(txtCodigo), Trim(Txtdescripcion), 0, 0, 0, 0
    grabaBitacora "M", s2n(txtCodigo), "SubGruposProducto"
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk '"codigo = '" & txtCodigo & "'"
    GoTo fin
ufaErr:
    ufa "err al grabar modi", Me.Name ', Err
fin:
End Sub

Private Sub ucMenu_BorrarControles()
    FrmBorrarTxt Me
End Sub

Private Sub ucMenu_Buscar()
    Dim re As String
    re = frmBuscar.MostrarCodigoDescripcionActivo("SubGruposProducto")
    If re > "" Then
        CargaDatos re, frmBuscar.resultado(2)
        ucMenu.BuscarOK "codigo = '" & re & "'"
    End If
End Sub

Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    DataEnvironment1.dbo_SUBGRUPO "B", Trim(txtCodigo), "", 0, 0, UsuarioSistema!codigo, Date
    grabaBitacora "B", 0, "SubGrupoProducto"
    ucMenu.EliminarOK
    GoTo fin
ufaErr:
    ufa "Falla al eliminar", Me.Name ', Err
fin:
End Sub

Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    Txtdescripcion.Locked = Not sino
    txtCodigo.Locked = Not sino
End Sub
Private Sub ucMenu_Modificar()
    txtCodigo.Locked = True
    Txtdescripcion.SetFocus
End Sub
Private Sub ucMenu_Nuevo()
    txtCodigo.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
Private Sub ucMenu_SeMovio()
    CargaDatos ucMenu.rs!codigo, ucMenu.rs!descripcion
End Sub
'----------------------------- MENU ----------------------------------------


'-------- Privado Varios --------
Private Sub CargaDatos(co, de)
    txtCodigo = co
    Txtdescripcion = de
End Sub
Private Function FaltaDescr() As Boolean
    If Trim(Txtdescripcion) = "" Then
        FaltaDescr = True
        MsgBox "Falta cargar la descripcion", 48, "Atencion"
    End If
End Function
Private Function YaTa() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from SubGruposProducto where codigo='" & Trim(txtCodigo) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
       MsgBox "Codigo repetido", 48, "Atencion"
        YaTa = True
    End If
ufaErr:
fin:
    Set rs = Nothing
End Function
'-------- Privado Varios --------
