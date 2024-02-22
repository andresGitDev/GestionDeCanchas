VERSION 5.00
Begin VB.Form frmsucursal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ABM Sucursal"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   3945
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   11055
      Begin VB.TextBox Text6 
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
         Left            =   8205
         TabIndex        =   21
         Top             =   2880
         Width           =   2550
      End
      Begin VB.TextBox Text5 
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
         Left            =   4605
         TabIndex        =   19
         Top             =   2880
         Width           =   2070
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   17
         Top             =   2880
         Width           =   2070
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
         Height          =   285
         Left            =   9840
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   11
         Top             =   975
         Width           =   4590
      End
      Begin VB.TextBox Calle 
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
         TabIndex        =   5
         Top             =   1845
         Width           =   4695
      End
      Begin VB.TextBox txtLocalidad 
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
         Height          =   285
         Left            =   6930
         TabIndex        =   4
         Top             =   1860
         Width           =   2775
      End
      Begin VB.ComboBox CmbProvincia 
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
         Left            =   1155
         TabIndex        =   3
         Top             =   2355
         Width           =   4590
      End
      Begin VB.CommandButton cmdLocBusco 
         Caption         =   "Localidad"
         Height          =   345
         Left            =   9765
         TabIndex        =   2
         Top             =   1875
         Width           =   1200
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   285
         Left            =   1170
         TabIndex        =   9
         Top             =   1440
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uFlete 
         Height          =   285
         Left            =   1170
         TabIndex        =   23
         Top             =   3480
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Flete:"
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
         TabIndex        =   24
         Top             =   3480
         Width           =   810
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
         Left            =   6870
         TabIndex        =   22
         Top             =   2925
         Width           =   1215
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
         Left            =   3960
         TabIndex        =   20
         Top             =   2925
         Width           =   645
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
         Left            =   240
         TabIndex        =   18
         Top             =   2925
         Width           =   885
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
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
         Left            =   9000
         TabIndex        =   16
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Sucursal:"
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
         Left            =   360
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
         TabIndex        =   12
         Top             =   1005
         Width           =   885
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
         TabIndex        =   10
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion:"
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
         Left            =   165
         TabIndex        =   8
         Top             =   1890
         Width           =   930
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Height          =   300
         Left            =   165
         TabIndex        =   7
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
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
         Left            =   5940
         TabIndex        =   6
         Top             =   1905
         Width           =   1050
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Height          =   1440
      Left            =   0
      TabIndex        =   0
      Top             =   4095
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   2540
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "frmsucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLocBusco_Click()
Dim resu As Variant, CodigoProv As String
    If CmbProvincia.Text <> "" Then
        CodigoProv = "'" & obtenerDeSQL("select codigo from provincias where descripcion = '" & CmbProvincia.Text & "'") & "'"
        resu = frmBuscar.MostrarSql("select numero as NUM,localidad as LOCALIDAD , partido as PARTIDO from localidades where provincia = " & CodigoProv, , "Localidades", "-")
    Else
        resu = frmBuscar.MostrarSql("select numero as NUM,localidad as LOCALIDAD , partido as PARTIDO from localidades", , "Localidades", "-")
    End If
    If resu > "" Then
        txtlocalidad = obtenerDeSQL("select localidad from localidades where numero = " & resu)
    End If
End Sub

Private Sub Form_Load()
    ucMenu.init True, True, True, False, True
    
    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    uFlete.ini "select descripcion from transportes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from transportes where activo = 1", False
    CargaProvincia
    
    Text3.Text = obtenerDeSQL("select max(cod) from sucursal") + 1
    
End Sub
Sub CargaProvincia()
    CmbProvincia.AddItem "BUENOS AIRES"
    CmbProvincia.AddItem "CAPITAL FEDERAL"
    CmbProvincia.AddItem "CATAMARCA"
    CmbProvincia.AddItem "CHACO"
    CmbProvincia.AddItem "CHUBUT"
    CmbProvincia.AddItem "CORDOBA"
    CmbProvincia.AddItem "CORRIENTES"
    CmbProvincia.AddItem "ENTRE RIOS"
    CmbProvincia.AddItem "FORMOSA"
    CmbProvincia.AddItem "JUJUY"
    CmbProvincia.AddItem "LA PAMPA"
    CmbProvincia.AddItem "LA RIOJA"
    CmbProvincia.AddItem "MENDOZA"
    CmbProvincia.AddItem "MISIONES"
    CmbProvincia.AddItem "NEUQUEN"
    CmbProvincia.AddItem "RIO NEGRO"
    CmbProvincia.AddItem "SALTA"
    CmbProvincia.AddItem "SAN JUAN"
    CmbProvincia.AddItem "SAN LUIS"
    CmbProvincia.AddItem "SANTA CRUZ"
    CmbProvincia.AddItem "SANTA FE"
    CmbProvincia.AddItem "SANTIAGO DEL ESTERO"
    CmbProvincia.AddItem "TIERRA DEL FUEGO"
    CmbProvincia.AddItem "TUCUMAN"
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
    
    verif = obtenerDeSQL("select cliente from sucursal where sucursal=" & Text2.Text & " and cliente=" & uCliente.codigo)
    If verif = "" Or IsNull(verif) Or IsEmpty(verif) Then
    Else
        MsgBox "Esta sucursal ya existe para este cliente, seleccione otro nro de sucursal.", , "ATENCION"
        Exit Sub
    End If
    
    prov = obtenerDeSQL("select codigo from provincias where descripcion='" & Trim(CmbProvincia.Text) & "'")
    sql = "insert into sucursal (cod,cliente,sucursal,nombre,direccion,local_su,prov_su,telefono,fax,contacto,flete,unidades,saldo,fecha_alta,fecha_baja,usuario_alta,usuario_baja,activo) " & _
        " values( " & Text3 & ", " & uCliente.codigo & "," & Text2 & ",'" & Text1 & "','" & Calle & "','" & txtlocalidad & "','" & prov & "','" & Text4 & "','" & Text5 & "','" & Text6 & "'," & uFlete.codigo & ",'U',0," & ssFecha(Date) & "," & ssFecha("01/01/1900") & "," & UsuarioActual & ",0,1 )"
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La alta se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    
End Sub

Private Sub ucMenu_AceptarModi()
    Dim verif
    Dim prov As String
    Dim sql As String
    
    verif = obtenerDeSQL("select cliente from sucursal where sucursal=" & Text2.Text & " and cliente=" & uCliente.codigo)
    If verif = "" Or IsNull(verif) Then
        MsgBox "Esta sucursal ya existe para este cliente, seleccione otro nro de sucursal.", , "ATENCION"
        Exit Sub
    End If
    
    prov = obtenerDeSQL("select codigo from provincias where descripcion='" & Trim(CmbProvincia.Text) & "'")
    sql = "update sucursal set cliente=" & uCliente.codigo & ",sucursal=" & Text2 & ",nombre='" & Text1 & "',direccion='" & Calle & "',local_su='" & txtlocalidad & "',prov_su='" & prov & "',telefono='" & Text4 & "',fax='" & Text5 & "',contacto='" & Text6 & "',flete=" & uFlete.codigo & " where cod=" & Text3.Text
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La modificacion se ha realizado con exito.", , "ATENCION"
    ucMenu.AceptarOk
    
End Sub

Private Sub ucMenu_Buscar()
    Dim resu
    Dim a As String
    Dim rs As New ADODB.Recordset
    
    resu = frmBuscar.MostrarSql("select sucursal,nombre [    Nombre              ],direccion [     Direccion           ],cliente from sucursal where activo = 1", , , "", , , True)
    
    If resu <> "" Then
        rs.Open "Select * from sucursal where sucursal=" & frmBuscar.resultado(1) & " and cliente=" & frmBuscar.resultado(4), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        Text3.Text = rs!COD
        Text2.Text = frmBuscar.resultado(1)
        Text1.Text = frmBuscar.resultado(2)
        Calle.Text = frmBuscar.resultado(3)
        txtlocalidad.Text = sSinNull(rs!local_su)
        Text4.Text = sSinNull(rs!Telefono)
        Text5.Text = sSinNull(rs!fax)
        Text6.Text = sSinNull(rs!contacto)
                
        uFlete.codigo = rs!flete
        uCliente.codigo = frmBuscar.resultado(4)
        If Not IsNull(rs!prov_su) Then
            a = rs!prov_su
            Set rs = Nothing
            rs.Open "Select * from provincias where codigo='" & a & "' and activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If IsNull(rs!DESCRIPCION) Or IsEmpty(rs!DESCRIPCION) Or (rs.EOF = True And rs.BOF = True) Then
            Else
                CmbProvincia.Text = rs!DESCRIPCION
            End If
        End If
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
    
    If Text2.Text = "" Then
        MsgBox "No se puede eliminar una sucursal, falta ingresar el numero.", , "ATENCION"
        Exit Sub
    End If
    
    sql = "update sucursal set activo=0 where cod=" & Text3.Text
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "La eliminacion se ha realizado con exito.", , "ATENCION"
    ucMenu.EliminarOK
    
End Sub

Private Sub ucMenu_Modificar()
    Frame1.enabled = True
End Sub

Private Sub ucMenu_Nuevo()
    Frame1.enabled = True
    Text3.Text = obtenerDeSQL("select max(cod) from sucursal") + 1
End Sub

Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Private Sub uFlete_cambio(codigo As Variant)
    On Error Resume Next
    Dim tmp
    
    tmp = obtenerDeSQL("select codigo from transportes where codigo = " & uFlete.codigo)
    
End Sub

Private Sub limpiar()
    Text3.Text = ""
    Text2.Text = ""
    Text1.Text = ""
    uCliente.clear
    Calle.Text = ""
    txtlocalidad.Text = ""
    CmbProvincia.ListIndex = 0
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    uFlete.clear
    Text3.Text = obtenerDeSQL("select max(cod) from sucursal") + 1
End Sub
