VERSION 5.00
Begin VB.Form frmPRemito 
   Caption         =   "Punto de Remitos"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin Gestion.ucBotonera ucBotonera1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2566
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtDescri 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txtPunto 
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "ej. 0001"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Punto :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Limpio
    Frame1.enabled = False
    ucBotonera1.init True, True, True, False, True
    ucBotonera1.AceptarOk
End Sub

Private Sub txtPunto_LostFocus()
    txtPunto = Format(txtPunto, "0000")
End Sub

Private Sub ucBotonera1_AceptarAlta()
    Dim sql As String
    
    If txtPunto.Text = "" Then
        MsgBox "Debe cargar un numero de punto.", , "ATENCION"
        Exit Sub
    End If
    If obtenerDeSQL("select id from puntoremito where punto='" & Trim(txtPunto.Text) & "'") > 0 Then
        MsgBox "El punto ya existe, debe cambiarlo.", , "ATENCION"
        Exit Sub
    End If
    If obtenerDeSQL("select id from puntoremito where descripcion='" & Trim(txtDescri.Text) & "'") > 0 Then
        MsgBox "La descripcion ya existe verifique que no este intentando cargar el mismo punto.", , "ATENCION"
        Exit Sub
    End If
    If txtDescri.Text = "" Then txtDescri.Text = "-"
    
    sql = "insert into puntoremito(punto,descripcion,activo) values ('" & Trim(txtPunto.Text) & "','" & Trim(txtDescri.Text) & "',1)"
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se ha cargado con exito.", , "ATENCION"
    Limpio
    Frame1.enabled = False
    ucBotonera1.AceptarOk
End Sub

Private Sub ucBotonera1_AceptarModi()
    Dim sql As String
    
    If txtPunto.Text = "" Then
        MsgBox "Debe cargar un numero de punto.", , "ATENCION"
        Exit Sub
    End If
    If txtDescri.Text = "" Then txtDescri.Text = "-"
    
    sql = "update puntoremito set punto='" & Trim(txtPunto.Text) & "',descripcion='" & Trim(txtDescri.Text) & "' where id=" & lblID.caption
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se ha modificado con exito.", , "ATENCION"
    Limpio
    Frame1.enabled = False
    ucBotonera1.AceptarOk
End Sub

Private Sub ucBotonera1_buscar()
    Dim resu
    Limpio
    resu = frmBuscar.MostrarSql("select ID, Punto,Descripcion from puntoremito where activo=1")
    If resu <> "" Then
        lblID.caption = resu
        txtPunto.Text = frmBuscar.resultado(2)
        txtDescri.Text = frmBuscar.resultado(3)
    End If
    ucBotonera1.BuscarOK
End Sub

Private Sub ucBotonera1_Cancelar()
    Limpio
    ucBotonera1.AceptarOk
    Frame1.enabled = False
End Sub

Private Sub ucBotonera1_eliminar()
    Dim sql As String
    
    If MsgBox("Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, "ATENCION") = vbNo Then
        Exit Sub
    End If
    
    sql = "update puntoremito set activo=0 where id=" & lblID.caption
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se elimino con exito.", , "ATENCION"
    Limpio
    Frame1.enabled = False
    ucBotonera1.AceptarOk
End Sub

Private Sub ucBotonera1_modificar()
    Frame1.enabled = True
End Sub

Private Sub ucBotonera1_nuevo()
    Frame1.enabled = True
    Limpio
End Sub

Private Sub ucBotonera1_salir()
    Unload Me
End Sub

Private Sub Limpio()
    lblID.caption = s2n(obtenerDeSQL("select max(id) from puntoRemito")) + 1
    txtPunto.Text = ""
    txtDescri.Text = ""
End Sub
