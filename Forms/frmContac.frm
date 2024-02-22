VERSION 5.00
Begin VB.Form frmContac 
   Caption         =   "Contactos"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   7560
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
      Height          =   5595
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7380
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
         TabIndex        =   26
         Top             =   1800
         Width           =   4590
      End
      Begin VB.TextBox txtCel 
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
         TabIndex        =   24
         Top             =   2910
         Width           =   4590
      End
      Begin VB.TextBox txtInt 
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
         Left            =   1170
         TabIndex        =   22
         Top             =   2565
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
         TabIndex        =   20
         Top             =   1440
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
         Height          =   315
         Left            =   1185
         TabIndex        =   7
         Top             =   3630
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
         TabIndex        =   6
         Top             =   1065
         Width           =   4590
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
         TabIndex        =   5
         Top             =   720
         Width           =   4590
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
         TabIndex        =   4
         Top             =   345
         Width           =   1170
      End
      Begin VB.TextBox txtTel 
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
         Top             =   2175
         Width           =   4575
      End
      Begin VB.TextBox txtFax 
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
         TabIndex        =   2
         Top             =   3255
         Width           =   4590
      End
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
         Height          =   630
         Left            =   1470
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   4815
         Width           =   4260
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   285
         Left            =   1170
         TabIndex        =   8
         Top             =   4020
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uFlete 
         Height          =   285
         Left            =   1170
         TabIndex        =   9
         Top             =   4380
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         CodigoWidth     =   1000
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
         TabIndex        =   27
         Top             =   1830
         Width           =   885
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular:"
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
         TabIndex        =   25
         Top             =   2940
         Width           =   885
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Interno:"
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
         Left            =   255
         TabIndex        =   23
         Top             =   2595
         Width           =   885
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
         TabIndex        =   21
         Top             =   1470
         Width           =   885
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
         TabIndex        =   18
         Top             =   3675
         Width           =   1050
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
         TabIndex        =   17
         Top             =   1110
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         TabIndex        =   16
         Top             =   4020
         Width           =   810
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
         TabIndex        =   15
         Top             =   750
         Width           =   885
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
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
         Left            =   225
         TabIndex        =   13
         Top             =   2235
         Width           =   885
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
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
         Left            =   255
         TabIndex        =   12
         Top             =   3330
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Observacion:"
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
         Left            =   135
         TabIndex        =   11
         Top             =   4860
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
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
         Top             =   4380
         Width           =   915
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Height          =   1440
      Left            =   0
      TabIndex        =   19
      Top             =   5820
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   2540
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "frmContac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    ucMenu.init True, True, True, False, True
    
    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    uFlete.ini "select nombre from sucursal where cod = '###'", "select cod as [ Codigo ], nombre as [ Nombre                        ] from sucursal where activo = 1", False
    
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from contacto")) + 1
    
End Sub

Private Sub uCliente_cambio(codigo As Variant)
    CargoDatosCliente
End Sub
Private Sub CargoDatosCliente()
    On Error Resume Next
    Dim tmp
    
    tmp = obtenerDeSQL("select codigo from clientes where codigo = " & uCliente.codigo)
    
End Sub

Private Sub ucMenu_AceptarAlta()
    Dim verif
    Dim prov As String
    Dim sql As String
    
    verif = obtenerDeSQL("select cliente from contacto where activo=1 and nombre='" & Trim(txtNom) & "' and apellido='" & Trim(txtApell) & "' and sucursal=" & uFlete.codigo & " and cliente=" & uCliente.codigo)
    If verif = "" Or IsNull(verif) Or IsEmpty(verif) Then
    Else
        MsgBox "Este contacto ya existe para este cliente.", , "ATENCION"
        Exit Sub
    End If
    
    'prov = obtenerDeSQL("select codigo from provincias where descripcion='" & Trim(CmbProvincia.Text) & "'")
    sql = "insert into contacto (nombre,apellido,cargo,sector,telefono,interno,celular,fax,email,cliente,sucursal,nota,activo) " & _
        " values( '" & Trim(txtNom) & "','" & Trim(txtApell) & "','" & Trim(txtCargo) & "','" & Trim(txtSec) & "','" & Trim(txtTel) & "','" & Trim(txtInt) & "','" & Trim(txtCel) & "','" & Trim(txtFax) & "','" & Trim(txtEmail) & "'," & uCliente.codigo & "," & uFlete.codigo & ",'" & Trim(txtObs) & "',1 )"
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La alta se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    
End Sub

Private Sub ucMenu_AceptarModi()
    Dim verif
    Dim prov As String
    Dim sql As String
        
    sql = "update contacto set cliente=" & uCliente.codigo & ",sucursal=" & uFlete.codigo & ",nombre='" & Trim(txtNom) & "',apellido='" & Trim(txtApell) & "',cargo='" & Trim(txtCargo) & "',sector='" & Trim(txtSec) & "',telefono='" & Trim(txtTel) & "',interno='" & Trim(txtInt) & "',celular='" & Trim(txtCel) & "',fax='" & Trim(txtFax) & "',email='" & Trim(txtEmail) & "',nota='" & Trim(txtObs) & "' where id=" & Text3.Text
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La modificacion se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    
End Sub

Private Sub ucMenu_Buscar()
    Dim resu
    Dim a As String
    Dim rs As New ADODB.Recordset
    
    'resu = frmBuscar.MostrarSql("select id [ Codigo     ],nombre [    Nombre              ],apellido [ Apellido              ],cliente,sucursal from contacto where activo = 1", , , "", , , True)
    resu = frmBuscar.MostrarSql("select c.id [ Codigo     ],c.nombre [    Nombre              ],c.apellido [ Apellido              ],isnull(cl.descripcion,'') as descripcion,isnull(s.nombre,'') as nombre " & _
            "from contacto c left outer join clientes cl on cl.codigo=c.cliente left outer join sucursal s on s.cod=c.sucursal where c.activo = 1", , , "", , , True)
    
    If resu <> "" Then
        rs.Open "Select * from contacto where id=" & frmBuscar.resultado(1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        Text3.Text = rs!ID
        txtNom.Text = frmBuscar.resultado(2)
        txtApell.Text = frmBuscar.resultado(3)
        txtCargo.Text = sSinNull(rs!cargo)
        txtSec.Text = sSinNull(rs!sector)
        txtTel.Text = sSinNull(rs!Telefono)
        txtInt.Text = sSinNull(rs!interno)
        txtCel.Text = sSinNull(rs!celular)
        txtFax.Text = sSinNull(rs!fax)
        txtEmail.Text = sSinNull(rs!email)
        uCliente.codigo = sSinNull(rs!cliente)
        uFlete.codigo = sSinNull(rs!sucursal)
        txtObs.Text = sSinNull(rs!nota)
        
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
        MsgBox "No se puede eliminar el contacto, falta el id.", , "ATENCION"
        Exit Sub
    End If
    
    sql = "update contacto set activo=0 where id=" & Text3.Text
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
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from contacto")) + 1
End Sub

Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Private Sub uFlete_cambio(codigo As Variant)
    On Error Resume Next
    Dim tmp
    
    tmp = obtenerDeSQL("select cod from sucursal where cod = " & uFlete.codigo)
    
End Sub

Private Sub limpiar()
    Text3.Text = ""
    txtNom.Text = ""
    txtApell.Text = ""
    uCliente.clear
    txtCargo.Text = ""
    txtTel.Text = ""
    txtInt.Text = ""
    txtCel.Text = ""
    txtFax.Text = ""
    txtEmail.Text = ""
    txtObs.Text = ""
    txtSec.Text = ""
    uFlete.clear
    Text3.Text = nSinNull(obtenerDeSQL("select max(id) from contacto")) + 1
End Sub


