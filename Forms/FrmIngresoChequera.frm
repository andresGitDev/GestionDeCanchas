VERSION 5.00
Begin VB.Form FrmIngresoChequera 
   Caption         =   "Ingreso de Chequera"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "FrmIngresoChequera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   4050
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optingreso 
      Alignment       =   1  'Right Justify
      Caption         =   "Ingreso"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   240
      Width           =   1575
   End
   Begin VB.OptionButton optanulo 
      Alignment       =   1  'Right Justify
      Caption         =   "Anulación"
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
      Left            =   3120
      TabIndex        =   16
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txttipocta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Tag             =   "2"
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtdesbanco 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Tag             =   "2"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox txtinicial 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "1"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtcantidad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "2"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtnumcta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Tag             =   "2"
      Top             =   2400
      Width           =   4455
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox txtcodcta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "0"
      Top             =   600
      Width           =   1335
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "5"
      Top             =   3600
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      Top             =   3600
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "4"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3330
      Left            =   120
      Top             =   120
      Width           =   6000
   End
   Begin VB.Label Label14 
      Caption         =   "Banco:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Cta.:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Nº Ch. Inicial:"
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
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Nº Cuenta:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Cant. Cheques:"
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
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblcuenta 
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "FrmIngresoChequera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Private Sub cmbcuenta_Click()
'    FrmHelp.Show
'    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
'    FrmHelp.Tag = Me.Name
    
    Dim resu As String

    resu = frmBuscar.MostrarSql("select ctasBank.codigo, descripcion as [Banco         ], Numero from ctasBank inner join BancosGrales on CtasBank.Banco = BancosGrales.Codigo where ctasBank.activo = 1 order by CtasBank.codigo ")
    If resu > "" Then
        txtcodcta = resu
        CargarDatos
    End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub optanulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optingreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtcantidad_GotFocus()
    txtCantidad.SelStart = 0
    txtCantidad.SelLength = Len(txtCantidad.Text)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtCantidad_LostFocus()
    If Not IsNumeric(txtCantidad) Then
        MsgBox "Cantidad incorrecta"
'        txtcantidad = "0"
'        txtcantidad.SetFocus
    End If
End Sub

Private Sub txtcodcta_GotFocus()
    txtcodcta.SelStart = 0
    txtcodcta.SelLength = Len(txtcodcta.Text)
    
    If optingreso = False And optanulo = False Then
        MsgBox "Debe ingresar un tipo de movimiento"
    End If

End Sub

Private Sub txtcodcta_LostFocus()
    If IsNumeric(txtcodcta) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
'            txtdescuenta = ObtenerDescripcion("BancosGrales", rs!banco) & " - " & rs!numero
            txtdesbanco = ObtenerDescripcionBancos("BancosGrales", rs!Banco)
            txttipocta = rs!Tipo & " - " & ObtenerDescripcion("TipoCtas", val(rs!Tipo))
            txtnumcta = rs!numero
            txtInicial.SetFocus
        Else
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta = "0"
            txtcodcta.SetFocus
        End If
        rs.Close
        Set rs = Nothing
        
    Else
        If txtcodcta <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta = "0"
            txtcodcta.SetFocus
        End If
    End If
End Sub

Public Sub CargarDatos()
Dim rs As New ADODB.Recordset

'    If rsefec.State = 1 Then
'        rsefec.Close
'        Set rsefec = Nothing
'    End If
    
    
    'codigo = Val(Trim(Me.Tag))
       
    rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        txtcodcta = rs!codigo
'        txtdescuenta = ObtenerDescripcionBancos("BancosGrales", rs!banco) & " - " & rs!numero
        txtdesbanco = ObtenerDescripcionBancos("BancosGrales", rs!Banco)
        txttipocta = rs!Tipo & " - " & ObtenerDescripcion("TipoCtas", val(rs!Tipo))
        txtnumcta = rs!numero
        txtInicial.SetFocus
    End If
    rs.Close
    Set rs = Nothing
        
End Sub

Private Sub cmdAceptar_Click()
Dim rs As New ADODB.Recordset
Dim valormaximo As Long
Dim x As Long, maximo As Long, numero As Long, Fecha

    If optingreso = False And optanulo = False Then
        MsgBox "Debe ingresar una da las dos operaciones (Ingreso / Anulación)"
        Exit Sub
    End If
    
    If txtcodcta = "" Then
        MsgBox "Debe ingresar el código de cuenta"
        Exit Sub
    End If
    
    If txtInicial = "" Then
        MsgBox "Debe ingresar el Nº inicial de los cheques"
        Exit Sub
    End If
    
    If txtCantidad = "" Then
        MsgBox "Debe ingresar la cantidad de cheques"
        Exit Sub
    End If
    
    If optingreso = True Then
        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
        
        rs.Open "select max(codigo) as maximo from chq_comp", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not IsNull(rs!maximo) Then
            valormaximo = rs!maximo
        Else
            valormaximo = 1
        End If
        rs.Close
        Set rs = Nothing
        
        numero = val(txtInicial) - 1
        
        For x = 1 To val(txtCantidad)
            valormaximo = valormaximo + 1
            numero = numero + 1
            DataEnvironment1.dbo_INGRESOCHEQUERA valormaximo, 0, numero, ObtenerCodigo("BancosGrales", txtdesbanco), val(txtcodcta), _
            0, 0, "", 0, "C", 0, 0, Fecha, UsuarioSistema!codigo, 0, 0, 1
        Next
    Else
        Dim cad As String
        cad = "update chq_comp set estado = 'N' where (nro >= " & x2s(txtInicial) & " and nro <= " & x2s(s2n(txtInicial) + s2n(txtCantidad)) & ") and estado = 'C' and importe=0 and cuentabancaria=" & txtcodcta
        DataEnvironment1.Sistema.Execute cad
        'DataEnvironment1.Sistema.Execute "update chq_comp set estado = 'N' where nro >= " & Val(txtinicial) & " and nro <= " & Val(txtcantidad) & " and estado = 'C'"
    End If
    
    MsgBox "La operación fue realizada con éxito"
    LimpioControles
    Call Habilitobotones(False)

End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Habilitobotones (False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub optanulo_Click()
    Habilitobotones (True)
End Sub

Private Sub optingreso_Click()
    Habilitobotones (True)
End Sub

Private Sub LimpioControles()
    txtcodcta = ""
'    txtdescuenta = ""
    txtdesbanco = ""
    txttipocta = ""
    txtnumcta = ""
    txtInicial = ""
    txtCantidad = ""
    optanulo = False
    optingreso = False
End Sub

Private Sub txtdesbanco_GotFocus()
    txtdesbanco.SelStart = 0
    txtdesbanco.SelLength = Len(txtdesbanco.Text)
End Sub

Private Sub txtinicial_GotFocus()
    txtInicial.SelStart = 0
    txtInicial.SelLength = Len(txtInicial.Text)
End Sub

Private Sub txtinicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtinicial_LostFocus()

    If optingreso = True Then
        If IsNumeric(txtInicial) Then
            Dim rs As New ADODB.Recordset

            rs.Open "select nro from chq_comp where nro = " & val(txtInicial) & " and cuentabancaria = " & val(txtcodcta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF And txtInicial <> "" Then
                MsgBox "Este Nro. de cheque ya existe para esta cuenta bancaria, intente nuevamente"
                txtInicial.SetFocus
            End If
            rs.Close
            Set rs = Nothing
        Else
            If txtInicial <> "" Then
                MsgBox "Nro. inicial incorrecto"
'               txtinicial = "0"
'               txtinicial.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub Habilitobotones(habilito As Boolean)
    txtcodcta.enabled = habilito
    cmbcuenta.enabled = habilito
    txtInicial.enabled = habilito
    txtCantidad.enabled = habilito
End Sub

Private Sub txtnumcta_GotFocus()
    txttipocta.SelStart = 0
    txttipocta.SelLength = Len(txttipocta.Text)
End Sub

Private Sub txttipocta_GotFocus()
    txttipocta.SelStart = 0
    txttipocta.SelLength = Len(txttipocta.Text)
End Sub


