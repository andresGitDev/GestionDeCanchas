VERSION 5.00
Begin VB.Form FrmLisProveedores 
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   Icon            =   "FrmLisProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optcateg 
      Caption         =   "Categoría"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton optelegir 
      Caption         =   "Elegir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   360
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton opttodos 
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtcateg 
      Height          =   285
      Left            =   3360
      TabIndex        =   18
      Tag             =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbcateg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Categoría"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtdescateg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   16
      Tag             =   "2"
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1500
      Left            =   240
      TabIndex        =   15
      Top             =   225
      Width           =   1935
   End
   Begin VB.Frame frameorden 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   780
      Left            =   225
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   8775
      Begin VB.OptionButton optcodigo 
         Caption         =   "Código"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Descripción"
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
         Height          =   240
         Left            =   5760
         TabIndex        =   13
         Top             =   260
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.TextBox cargar 
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      Tag             =   "1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtdesdesde 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   8
      Tag             =   "2"
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmbdesde 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prov."
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtdesde 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Tag             =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txthasta 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbhasta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Prov."
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtdeshasta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   975
      Left            =   120
      Top             =   1920
      Width           =   9015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label lblcateg 
      Caption         =   "Categoría:"
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
      Left            =   2295
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbldesde 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desde:"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblhasta 
      Caption         =   "Hasta:"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmLisProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 16/9/4


Private Sub cmbcateg_Click()
    FrmHelp.Show
    CargarHelp "provCategoria", "Codigo", "Descripción", "codigo", "descripcion", "codigo"
    FrmHelp.Tag = Me.Name
    cargar = "ProvCateg"
End Sub

Private Sub cmbdesde_Click()
    FrmHelp.Show
    CargarHelp "Prov", "Codigo", "Descripción", "codigo", "descripcion", "codigo"
    FrmHelp.Tag = Me.Name
    cargar = "ProvDesde"
End Sub

Private Sub cmbhasta_Click()
    FrmHelp.Show
    CargarHelp "Prov", "Codigo", "Descripción", "codigo", "descripcion", "codigo"
    FrmHelp.Tag = Me.Name
    cargar = "ProvHasta"
End Sub

Private Sub cmdAceptar_Click()
    Dim desde
    Dim hasta
        
    If optcodigo = False And optDesc = False Then
        MsgBox "debe ingresar un tipo de orden para realizar el listado"
        Exit Sub
    End If
    
    'el listado de proveedores los paso con:
    'OC=ordenado por codigo, OD=ordenado por descripcion
    'TC=todos por codigo, TD=todos por descripcion
    'CC=categoria por codigo, CD=categoria por descripcion
        
    If optelegir = True Then
        If optcodigo = True Then
            desde = val(txtdesde)
            hasta = val(txthasta)
            'DataEnvironment1.LisProveedoresPorCodigo Val(txtdesde), Val(txthasta)
            'DataEnvironment1.dbo_LisProveedor "OC", desde, hasta
            'rptProveedoresElijoCodigo.Show vbModal
            TraerProveedores "LisProveedor", "OC", desde, hasta
            Set rptProveedoresListados.DataControl1.Recordset = RStraer
            rptProveedoresListados.Field6.Text = RStraer.RecordCount
            rptProveedoresListados.Field7.Text = Date
            rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR CODIGO"
            rptProveedoresListados.Show vbModal
            Set RStraer = Nothing
            'DataEnvironment1.rsdbo_LisProveedor.Close
            'DataEnvironment1.rsLisProductosTodosPorCodigo.Close
        Else
            desde = val(txtdesde)
            hasta = val(txthasta)
            'DataEnvironment1.LisProveedoresPorDescripcion Val(txtdesde), Val(txthasta)
            
            'DataEnvironment1.dbo_LisProveedor "OD", desde, hasta
            'rptProveedoresElijoDescripcion.Show vbModal
            TraerProveedores "LisProveedor", "OD", desde, hasta
            Set rptProveedoresListados.DataControl1.Recordset = RStraer
            rptProveedoresListados.Field6.Text = RStraer.RecordCount
            rptProveedoresListados.Field7.Text = Date
            rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR DESCRIPCION"
            rptProveedoresListados.Show vbModal
            Set RStraer = Nothing
            'DataEnvironment1.rsdbo_LisProveedor.Close
            'DataEnvironment1.rsLisProveedoresPorDescripcion.Close
        End If
    Else
        If opttodos = True Then
            If optcodigo = True Then
                desde = 0
                hasta = 0
                'DataEnvironment1.LisProveedoresTodosPorCodigo
                
                'DataEnvironment1.dbo_LisProveedor "TC", DESDE, HASTA
                'rptProveedoresTodosCodigo.Show vbModal
                TraerProveedores "LisProveedor", "TC", desde, hasta
                Set rptProveedoresListados.DataControl1.Recordset = RStraer
                rptProveedoresListados.Field6.Text = RStraer.RecordCount
                rptProveedoresListados.Field7.Text = Date
                rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR CODIGO"
                rptProveedoresListados.Show vbModal
                Set RStraer = Nothing
                'DataEnvironment1.rsdbo_LisProveedor.Close
                'DataEnvironment1.rsLisProveedoresTodosPorCodigo.Close
            Else
                desde = 0
                hasta = 0
                'DataEnvironment1.LisProveedoresTodosPorDescripcion
                
                'DataEnvironment1.dbo_LisProveedor "TD", desde, hasta
                'rptProveedoresTodosDescripcion.Show vbModal
                TraerProveedores "LisProveedor", "TD", desde, hasta
                Set rptProveedoresListados.DataControl1.Recordset = RStraer
                rptProveedoresListados.Field6.Text = RStraer.RecordCount
                rptProveedoresListados.Field7.Text = Date
                rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR DESCRIPCION"
                rptProveedoresListados.Show vbModal
                Set RStraer = Nothing
                'DataEnvironment1.rsdbo_LisProveedor.Close
                'DataEnvironment1.rsLisProveedoresTodosPorDescripcion.Close
            End If
        Else
            If txtcateg <> "" Then
                rptProveedoresListados.lblcateg.Visible = True
                rptProveedoresListados.Label11.Visible = True
                If optcodigo = True Then
                    desde = val(txtcateg)
                    hasta = 0
                    'DataEnvironment1.LisProveedoresCategPorCodigo Val(txtcateg)
                    
                    'DataEnvironment1.dbo_LisProveedor "CC", desde, hasta
                    'rptProveedoresCategCodigo.Show vbModal
                    TraerProveedores "LisProveedor", "CC", desde, hasta
                    rptProveedoresListados.lblcateg.caption = txtdescateg
                    Set rptProveedoresListados.DataControl1.Recordset = RStraer
                    rptProveedoresListados.Field6.Text = RStraer.RecordCount
                    rptProveedoresListados.Field7.Text = Date
                    rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR CATEGORIA"
                    rptProveedoresListados.Show vbModal
                    Set RStraer = Nothing
                    'DataEnvironment1.rsdbo_LisProveedor.Close
                    'DataEnvironment1.rsLisProveedoresCategPorCodigo.Close
                Else
                    desde = val(txtcateg)
                    hasta = 0
                    'DataEnvironment1.LisProveedoresCategPorDescripcion Val(txtcateg)
                    
                    'DataEnvironment1.dbo_LisProveedor "CD", desde, hasta
                    'rptProveedoresCategDescripcion.Sections("Medio").Controls("lblcateg").caption = txtdescateg
                    'rptProveedoresCategDescripcion.Show vbModal
                    TraerProveedores "LisProveedor", "CD", desde, hasta
                    rptProveedoresListados.lblcateg.caption = txtdescateg
                    Set rptProveedoresListados.DataControl1.Recordset = RStraer
                    rptProveedoresListados.Field6.Text = RStraer.RecordCount
                    rptProveedoresListados.Field7.Text = Date
                    rptProveedoresListados.Label2.caption = "LISTADO DE PROVEEDORES POR CATEGORIA"
                    rptProveedoresListados.Show vbModal
                    Set RStraer = Nothing
                    'DataEnvironment1.rsdbo_LisProveedor.Close
                    'DataEnvironment1.rsLisProveedoresCategPorDescripcion.Close
                End If
            Else
                MsgBox "Debe ingresar una categoría"
            End If
        End If
    End If
        
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    VerOrden (False)
    VerTextos (False)
    VerCateg (False)
End Sub

Private Sub LimpioControles()
    txtdesde = ""
    txtdesdesde = ""
    txthasta = ""
    txtdeshasta = ""
    cargar = ""
    optcodigo = False
    optDesc = False
    optelegir = False
    opttodos = False
End Sub

Private Sub VerTextos(habilito As Boolean)
    txtdesde.Visible = habilito
    txtdesdesde.Visible = habilito
    txthasta.Visible = habilito
    txtdeshasta.Visible = habilito
'    optcodigo.Visible = habilito
'    optDesc.Visible = habilito
    cmbdesde.Visible = habilito
    cmbhasta.Visible = habilito
    lbldesde.Visible = habilito
    lblhasta.Visible = habilito
'    optelegir.Visible = habilito
'    opttodos.Visible = habilito
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub optcateg_Click()
    VerTextos (False)
    VerOrden (True)
    VerCateg (True)
End Sub

Private Sub optcateg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub



Private Sub optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub



Private Sub optDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optelegir_Click()
    VerTextos (True)
    VerOrden (True)
    VerCateg (False)
End Sub

Private Sub optelegir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub opttodos_Click()
    VerTextos (False)
    VerOrden (True)
    VerCateg (False)
End Sub

Private Sub opttodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtcateg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcateg_Lostfocus()
    If IsNumeric(txtcateg) Then
        txtdescateg = ObtenerDescripcion("provCategoria", val(txtcateg))
        If txtdescateg = "" Then
            MsgBox "Codigo de Categoría incorrecta"
            txtcateg.SetFocus
        End If
    Else
        If txtcateg <> "" Then
            MsgBox "Codigo de Categoría incorrecta"
            txtcateg = "0"
            txtcateg.SetFocus
        End If
    End If
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtdesde_LostFocus()
    If IsNumeric(txtdesde) Then
        txtdesdesde = ObtenerDescripcion("Prov", val(txtdesde))
        If txtdesdesde = "" Then
            MsgBox "Codigo de Proveedor incorrecto"
            txtdesde.SetFocus
        End If
    Else
        If txtdesde <> "" Then
            MsgBox "Codigo de Proveedor incorrecto"
            txtdesde = "0"
            txtdesde.SetFocus
        End If
    End If
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txthasta_LostFocus()
    If IsNumeric(txthasta) Then
        txthasta = ObtenerDescripcion("Prov", val(txthasta))
        If txthasta = "" Then
            MsgBox "Codigo de Proveedor incorrecto"
            txthasta.SetFocus
        End If
    Else
        If txthasta <> "" Then
            MsgBox "Codigo de Proveedor incorrecta"
            txthasta = "0"
            txthasta.SetFocus
        End If
    End If
End Sub

Public Sub CargarDatos()
Dim rs As New ADODB.Recordset, codigo
    
    codigo = val(Trim(Me.Tag))
    
    If cargar = "ProvDesde" Then
        rs.Open "select * from Prov where codigo = " & codigo & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            txtdesde = rs!codigo
            txtdesdesde = rs!DESCRIPCION
        End If
        
        rs.Close
        Set rs = Nothing
    End If

    If cargar = "ProvHasta" Then
        rs.Open "select * from Prov where codigo = " & codigo & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            txthasta = rs!codigo
            txtdeshasta = rs!DESCRIPCION
        End If
        
        rs.Close
        Set rs = Nothing
    End If

    If cargar = "ProvCateg" Then
        rs.Open "select * from provCategoria where codigo = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            txtcateg = rs!codigo
            txtdescateg = rs!DESCRIPCION
        End If
        
        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub VerOrden(habilito As Boolean)
    frameorden.Visible = habilito
    optcodigo.Visible = habilito
    optDesc.Visible = habilito
End Sub

Private Sub VerCateg(habilito As Boolean)
    lblcateg.Visible = habilito
    txtcateg.Visible = habilito
    cmbcateg.Visible = habilito
    txtdescateg.Visible = habilito
End Sub

