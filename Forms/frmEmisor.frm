VERSION 5.00
Begin VB.Form frmEmisor 
   Caption         =   "Emisores"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Datos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3075
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7380
      Begin VB.TextBox txtObs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1830
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2525
         Width           =   3900
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   345
         Width           =   1170
      End
      Begin VB.TextBox txtNom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   2
         Top             =   720
         Width           =   4590
      End
      Begin VB.TextBox txtApell 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1185
         TabIndex        =   3
         Top             =   1065
         Width           =   4590
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   2140
         Width           =   4590
      End
      Begin VB.TextBox txtCargo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   4
         Top             =   1440
         Width           =   4590
      End
      Begin VB.TextBox txtSec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1185
         TabIndex        =   5
         Top             =   1800
         Width           =   4590
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Letra presupuesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Contacto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   405
         Width           =   1530
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   2185
         Width           =   1050
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1830
         Width           =   885
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Height          =   1440
      Left            =   0
      TabIndex        =   8
      Top             =   3225
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   2540
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "frmEmisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    ucMenu.init True, True, True, False, True
    
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from emisor")) + 1
    
End Sub

Private Sub ucMenu_AceptarAlta()
    Dim verif
    Dim prov As String
    Dim sql As String
    
    verif = obtenerDeSQL("select id from emisor where activo=1 and nombre='" & Trim(txtNom) & "' and apellido='" & Trim(txtApell) & "'")
    If verif = "" Or IsNull(verif) Or IsEmpty(verif) Then
    Else
        MsgBox "Este emisor ya existe.", , "ATENCION"
        Exit Sub
    End If
    
    'prov = obtenerDeSQL("select codigo from provincias where descripcion='" & Trim(CmbProvincia.Text) & "'")
    sql = "insert into emisor (nombre,apellido,cargo,sector,email,letra,activo) " & _
        " values( '" & Trim(txtNom.Text) & "','" & Trim(txtApell) & "','" & Trim(txtCargo) & "','" & Trim(txtSec) & "','" & Trim(txtEmail) & "','" & Trim(txtobs) & "',1 )"
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La alta se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    limpiar
    
End Sub

Private Sub ucMenu_AceptarModi()
    Dim verif
    Dim prov As String
    Dim sql As String
        
    sql = "update emisor set nombre='" & Trim(txtNom) & "',apellido='" & Trim(txtApell) & "',cargo='" & Trim(txtCargo) & "',sector='" & Trim(txtSec) & "',email='" & Trim(txtEmail) & "',letra='" & Trim(txtobs) & "' where id=" & Text3.Text
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La modificacion se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    limpiar
    
End Sub

Private Sub ucMenu_Buscar()
    Dim resu
    Dim a As String
    Dim rs As New ADODB.Recordset
    
    limpiar
    'resu = frmBuscar.MostrarSql("select id [ Codigo     ],nombre [    Nombre              ],apellido [ Apellido              ],cliente,sucursal from contacto where activo = 1", , , "", , , True)
    resu = frmBuscar.MostrarSql("select c.id [ Codigo     ],c.nombre [    Nombre              ],c.apellido [ Apellido              ] from emisor c where c.activo = 1", , , "", , , True)
    
    If resu <> "" Then
        rs.Open "Select * from emisor where id=" & frmBuscar.resultado(1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        Text3.Text = rs!ID
'        txtNom2.TextRTF = frmBuscar.resultado(2)
        txtNom = frmBuscar.resultado(2)
        txtApell.Text = frmBuscar.resultado(3)
        txtCargo.Text = sSinNull(rs!cargo)
        txtSec.Text = sSinNull(rs!sector)
        txtEmail.Text = sSinNull(rs!email)
        txtobs.Text = sSinNull(rs!letra)
        
        Set rs = Nothing
    End If
    ucMenu.BuscarOK
End Sub

Private Sub ucMenu_Cancelar()
    limpiar
    Frame1.enabled = False
End Sub

Private Sub ucMenu_eliminar()
    Dim verif
    Dim prov As String
    Dim sql As String
    
    If Text3.Text = "" Then
        MsgBox "No se puede eliminar el emisor, falta el id.", , "ATENCION"
        Exit Sub
    End If
    
    sql = "update emisor set activo=0 where id=" & Text3.Text
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La eliminacion se ha realizado con exito.", , "ATENCION"
    limpiar
    ucMenu.EliminarOK
    
End Sub

Private Sub ucMenu_Modificar()
    Frame1.enabled = True
End Sub

Private Sub ucMenu_Nuevo()
    Frame1.enabled = True
    limpiar
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from emisor")) + 1
End Sub

Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Private Sub limpiar()
    Text3.Text = ""
    txtNom.Text = ""
    txtApell.Text = ""
    txtCargo.Text = ""
    txtEmail.Text = ""
    txtobs.Text = ""
    txtSec.Text = ""
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from emisor")) + 1
End Sub





