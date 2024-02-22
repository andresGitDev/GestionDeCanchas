VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListadoCheques 
   Caption         =   "Listado de Cheques Propios"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "FrmListadoCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcodcuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4320
      TabIndex        =   14
      Tag             =   "13"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmbcuenta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuenta"
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
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtcuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Tag             =   "2"
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
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
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
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
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   975
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Frame FrameCli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2280
      Left            =   225
      TabIndex        =   3
      Top             =   180
      Width           =   3075
      Begin VB.OptionButton optdebitarse 
         Caption         =   "Cheques Por Debitarse"
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
         Height          =   195
         Left            =   255
         TabIndex        =   18
         Top             =   345
         Width           =   2535
      End
      Begin VB.OptionButton optdebitados 
         Caption         =   "Cheques Debitados"
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
         Height          =   195
         Left            =   255
         TabIndex        =   17
         Top             =   825
         Width           =   2295
      End
      Begin VB.OptionButton optcartera 
         Caption         =   "Cheques En Cartera"
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
         Height          =   195
         Left            =   255
         TabIndex        =   16
         Top             =   1305
         Width           =   2295
      End
      Begin VB.OptionButton optanulados 
         Caption         =   "Cheques Anulados"
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
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
   End
   Begin VB.OptionButton opttodos 
      BackColor       =   &H00808080&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optelegir 
      BackColor       =   &H00808080&
      Caption         =   "Elegir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton optcateg 
      BackColor       =   &H00808080&
      Caption         =   "Categoría"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker fechadesde 
      Height          =   255
      Left            =   4350
      TabIndex        =   5
      Top             =   360
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   189399041
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   255
      Left            =   4350
      TabIndex        =   6
      Top             =   840
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   189399041
      CurrentDate     =   38052
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   8175
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta:"
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
      Left            =   3480
      TabIndex        =   15
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblhasta 
      Caption         =   "Hasta"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbldesde 
      Caption         =   "Desde"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "FrmListadoCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4


Private Sub cmbcuenta_Click()
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
End Sub

Private Sub cmdAceptar_Click()
    Dim rs As New ADODB.Recordset
    Dim codcuenta As Long, Banco As String, tipocta As String, numero As String

    CierroTodo

    rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        codcuenta = rs!codigo
        Banco = ObtenerDescripcion("BancosGrales", rs!Banco)
        tipocta = ObtenerDescripcion("TipoCtas", rs!Tipo)
        numero = rs!numero
    End If
    rs.Close
    Set rs = Nothing
    
    If optanulados = True Then
        
        DataEnvironment1.LisChequesAnulados fechadesde, fechahasta, val(txtcodcuenta)
        
        rptChequesAnulados.Sections("Medio").Controls("lblcodcta").caption = codcuenta
        rptChequesAnulados.Sections("Medio").Controls("lblbanco").caption = Banco
        rptChequesAnulados.Sections("Medio").Controls("lbltipocta").caption = tipocta
        rptChequesAnulados.Sections("Medio").Controls("lblnumero").caption = numero
        
        'paso el nombre de la empresa al encabezado del reporte
        rptChequesAnulados.Sections("Encabezado").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        
        rptChequesAnulados.Show 'vbModal
        'DataEnvironment1.rsLisChequesAnulados.Close
    
    End If
    
    
    If optCartera = True Then
        DataEnvironment1.LisChequesCartera val(txtcodcuenta)
        
        rptChequesCartera.Sections("Medio").Controls("lblcodcta").caption = codcuenta
        rptChequesCartera.Sections("Medio").Controls("lblbanco").caption = Banco
        rptChequesCartera.Sections("Medio").Controls("lbltipocta").caption = tipocta
        rptChequesCartera.Sections("Medio").Controls("lblnumero").caption = numero
        
        'paso el nombre de la empresa al encabezado del reporte
        rptChequesCartera.Sections("Encabezado").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        
        rptChequesCartera.Show 'vbModal
        'DataEnvironment1.rsLisChequesCartera.Close
    
    End If
    
    
    If optdebitados = True Then
        DataEnvironment1.LisChequesDebitados fechadesde, fechahasta, val(txtcodcuenta)
        
        rptChequesDebitados.Sections("Medio").Controls("lblcodcta").caption = codcuenta
        rptChequesDebitados.Sections("Medio").Controls("lblbanco").caption = Banco
        rptChequesDebitados.Sections("Medio").Controls("lbltipocta").caption = tipocta
        rptChequesDebitados.Sections("Medio").Controls("lblnumero").caption = numero
        
        'paso el nombre de la empresa al encabezado del reporte
        rptChequesDebitados.Sections("Encabezado").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        
        rptChequesDebitados.Show 'vbModal
        'DataEnvironment1.rsLisChequesDebitados.Close
    
    End If
    
    
    If optdebitarse = True Then
        DataEnvironment1.LisChequesDebitarse fechadesde, fechahasta, val(txtcodcuenta)
        
        rptChequesDebitarse.Sections("Medio").Controls("lblcodcta").caption = codcuenta
        rptChequesDebitarse.Sections("Medio").Controls("lblbanco").caption = Banco
        rptChequesDebitarse.Sections("Medio").Controls("lbltipocta").caption = tipocta
        rptChequesDebitarse.Sections("Medio").Controls("lblnumero").caption = numero
        
        'paso el nombre de la empresa al encabezado del reporte
        rptChequesDebitarse.Sections("Encabezado").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        
        rptChequesDebitarse.Show 'vbModal
        'DataEnvironment1.rsLisChequesDebitarse.Close
    
    End If
    
'    Cancelo
End Sub

Public Sub CargarDatos()
    Dim rs As New ADODB.Recordset
  
'    codigo = Val(Trim(Me.Tag))
    
    rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        txtcuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        cmdAceptar.enabled = True
        cmdcancelar.enabled = True
    End If
    rs.Close
    Set rs = Nothing

End Sub
  
Private Sub cmdCancelar_Click()
    Cancelo
End Sub

Private Sub Cancelo()
    fechadesde = Date
    fechahasta = Date
    txtcodcuenta = ""
    txtcuenta = ""
    cmdAceptar.enabled = False
    cmdcancelar.enabled = False
    optanulados.Value = False
    optCartera.Value = False
    optdebitados.Value = False
    optdebitarse.Value = False
    HabilitoControles (False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    fechadesde = Date
    fechahasta = Date
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub


Private Sub Form_Unload(cancel As Integer)
    CierroTodo
End Sub
Private Sub CierroTodo()
    On Error Resume Next
    DataEnvironment1.rsLisChequesAnulados.Close
    DataEnvironment1.rsLisChequesCartera.Close
    DataEnvironment1.rsLisChequesDebitados.Close
    DataEnvironment1.rsLisChequesDebitarse.Close
End Sub
Private Sub optanulados_Click()
    HabilitoControles (True)
End Sub

'Private Sub optanulados_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optcartera_Click()
    HabilitoControles (True)
    fechadesde.enabled = False
    fechahasta.enabled = False
End Sub

'Private Sub optcartera_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optdebitados_Click()
    HabilitoControles (True)
End Sub

'Private Sub optdebitados_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optdebitarse_Click()
    HabilitoControles (True)
End Sub

Private Sub HabilitoControles(habilito As Boolean)
    fechadesde.enabled = habilito
    fechahasta.enabled = habilito
    txtcodcuenta.enabled = habilito
    cmbcuenta.enabled = habilito
End Sub

'Private Sub optdebitarse_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub

Private Sub txtcodcuenta_LostFocus()
    If IsNumeric(txtcodcuenta) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
        
        If txtcuenta = "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcuenta.SetFocus
        Else
            CargarDatos
        End If
    Else
        If txtcodcuenta <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcuenta = "0"
            txtcodcuenta.SetFocus
        End If
    End If
End Sub
