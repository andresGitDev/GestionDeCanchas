VERSION 5.00
Begin VB.Form frmAbmTipoProv 
   Caption         =   "Tipo Proveedores"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   Icon            =   "frmAbmTipoProv.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtId 
      Height          =   345
      Left            =   6270
      MaxLength       =   4
      TabIndex        =   1
      Top             =   675
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtBase 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1845
      TabIndex        =   4
      Top             =   1560
      Width           =   1725
   End
   Begin VB.TextBox txtCoef 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1845
      TabIndex        =   5
      Top             =   2040
      Width           =   1020
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   345
      Left            =   1860
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1095
      Width           =   5880
   End
   Begin VB.TextBox txtCodigo 
      Height          =   345
      Left            =   1860
      MaxLength       =   4
      TabIndex        =   2
      Top             =   660
      Width           =   1425
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1590
      Left            =   0
      TabIndex        =   0
      Top             =   2670
      Width           =   8625
      _extentx        =   15214
      _extenty        =   2805
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Proveedor para calculo de retencion de ganancias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   615
      TabIndex        =   11
      Top             =   165
      Width           =   6165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      Height          =   255
      Index           =   2
      Left            =   5055
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Base Imponible"
      Height          =   255
      Index           =   5
      Left            =   345
      TabIndex        =   9
      Top             =   1650
      Width           =   1350
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Coeficiente"
      Height          =   255
      Index           =   4
      Left            =   735
      TabIndex        =   8
      Top             =   2070
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   1
      Left            =   345
      TabIndex        =   7
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   705
      Width           =   1305
   End
End
Attribute VB_Name = "frmAbmTipoProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    uMenu.init True, True, True, False, True, "select * from provtiporetgan where activo = 1", DataEnvironment1.Sistema, True
End Sub

Private Sub cargar(cual)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    
    uMenu.Find "codigo = '" & cual & "' "
        
    With uMenu.rs
        txtCodigo = !codigo
        txtDescripcion = !DESCRIPCION
        txtBase = !BaseImponible
        txtCoef = !Coeficiente
    End With
    uMenu.BuscarOK "codigo = '" & Trim(cual) & "' "
    
    Exit Sub
ufaChe:
    ufa "error al cargar registro", "cargar provTiporetgan"
End Sub

Private Function existe(codi) As Boolean
    existe = Not IsEmpty(obtenerDeSQL("select id,codigo from provTiporetgan where codigo = '" & txtCodigo & "' and activo = 1 "))
End Function
Private Function TaTodo() As Boolean
    Dim t As Double
    If Trim(txtCodigo) = "" Or Trim(txtDescripcion) = "" Then
        che "falta codigo /descripcion"
        Exit Function
    End If
    t = s2n(txtCoef, 6)
    If Trim(txtBase) = "" Or Trim(txtCoef) = "" Then
        che "faltan base imponible/coeficiente"
        Exit Function
    End If
    If t >= 1 Then
        che "coeficiente debe ser menor a uno"
        Exit Function
    End If
    
    TaTodo = True
End Function

Private Sub txtBase_LostFocus()
    If txtBase <> "" Then
        txtBase.Text = s2n(txtBase.Text)
    End If
End Sub

Private Sub txtCoef_LostFocus()
    If txtCoef <> "" Then
        txtCoef.Text = s2n(txtCoef.Text)
    End If
End Sub

Private Sub uMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    Dim STR As String
     
    If Not TaTodo() Then Exit Sub
    
    If existe(txtCodigo) Then
        ufa "err: ya existe codigo", txtCodigo
        Exit Sub
    End If
    STR = "insert into provTiporetgan " & _
        " (codigo, Descripcion, BaseImponible, coeficiente ) values " & _
        " ('" & Trim(txtCodigo) & "', '" & Trim(txtDescripcion) & "', " & x2s(txtBase) & " , " & x2s(txtCoef) & " ) "
    DataEnvironment1.Sistema.Execute STR
    
    che "Tipo Proveedor agregado correctamente"
    uMenu.AceptarOk
fin:
    Exit Sub
UFAalta:
    ufa "Err en alta", "alta provTiporetgan"
    Resume fin
End Sub
Private Sub uMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufamodi
    If Trim(txtCodigo) = "" Then
        ufa "err: No se puede modificar ", "provTiporetgan txtcodigo = ''"
        Exit Sub
    End If
    
    If Not TaTodo Then Exit Sub
    If Not existe(txtCodigo) Then
        ufa "err: no se pudo modificar", " provTiporetgan " & txtCodigo
        Exit Sub
    End If
    DataEnvironment1.Sistema.Execute _
        "update provTiporetgan set Descripcion = '" & Trim(txtDescripcion) & "', BaseImponible = '" & x2s(txtBase) & "', coeficiente = '" & x2s(txtCoef) & "' "
    che "Modificado"
    uMenu.AceptarOk
fin:
    Exit Sub
ufamodi:
    ufa "err: no se pudo modificar ", "modi tipo prov gan " & txtCodigo
    Resume fin
End Sub

Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
End Sub
Private Sub uMenu_Buscar()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select Codigo , Descripcion as [Descripcion                                            ] from provTiporetgan where activo = 1")
    If resu > "" Then cargar resu
End Sub
Private Sub uMenu_BuscarYa(que As Variant)
    If Trim(que) > "" Then cargar que
End Sub
Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAbaja
    
    Dim prov As Long

    If Not existe(txtCodigo) Then
        ufa "err: No se puede eliminar ", "provTiporetgan txtcodigo = ''"
        Exit Sub
    End If
    
' *********** HAbilitar este control
'   prov = s2n(obtenerdesql("select codigo  from prov where               "))
    If prov > 0 Then
        che "No se puede eliminar, codigo usado en proveedor " & prov
        Exit Sub
    End If
  
    DataEnvironment1.Sistema.Execute _
        "update provTiporetgan set activo = 0 where codigo = '" & txtCodigo & "' "
    che "eliminado"
    uMenu.EliminarOK

fin:
    Exit Sub
UFAbaja:
    ufa "err: No se puede modificar ", "provTiporetgan txtcodigo = " & txtCodigo
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    txtCodigo.enabled = sino
    txtDescripcion.enabled = sino
    txtBase.enabled = sino
    txtCoef.enabled = sino
End Sub
Private Sub uMenu_Modificar()
    On Error Resume Next
    txtCodigo.Locked = True
    txtDescripcion.SetFocus
End Sub
Private Sub uMenu_Nuevo()
    On Error Resume Next
    txtCodigo.Locked = False
    txtCodigo.SetFocus
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
Private Sub uMenu_SeMovio()
    cargar uMenu.rs!codigo
End Sub
