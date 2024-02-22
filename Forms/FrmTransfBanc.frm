VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransfBanc 
   Caption         =   "Transferencias Bancarias"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   Icon            =   "FrmTransfBanc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6960
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboEjercicio 
      Height          =   315
      Left            =   3600
      TabIndex        =   27
      Text            =   "Ejercicio"
      Top             =   2160
      Width           =   990
   End
   Begin Gestion.ucTipoCompra uCuentas 
      Height          =   3690
      Left            =   30
      TabIndex        =   25
      Top             =   2760
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   6509
   End
   Begin VB.Frame fraMenu 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   75
      TabIndex        =   16
      Top             =   6405
      Width           =   7485
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
         Left            =   6390
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   105
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
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "10"
         Top             =   120
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
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton cmdnuevo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Nuevo"
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
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "0"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Eliminar"
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
         Left            =   2085
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdbuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "0"
         Top             =   105
         Width           =   975
      End
      Begin VB.CommandButton CmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
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
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox movbanc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox cargar 
      Height          =   285
      Left            =   5745
      TabIndex        =   14
      Tag             =   "4"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtimporte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1785
      TabIndex        =   3
      Tag             =   "4"
      Top             =   1785
      Width           =   1575
   End
   Begin VB.TextBox txtconcepto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1785
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "3"
      Top             =   1395
      Width           =   6930
   End
   Begin VB.TextBox txtcuentao 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4425
      TabIndex        =   9
      Tag             =   "2"
      Top             =   480
      Width           =   4320
   End
   Begin VB.CommandButton cmbcuentao 
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
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtcodctao 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtcuentad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4425
      TabIndex        =   6
      Tag             =   "2"
      Top             =   930
      Width           =   4290
   End
   Begin VB.CommandButton cmbcuentad 
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
      Left            =   3345
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox txtcodctad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1785
      TabIndex        =   1
      Tag             =   "2"
      Top             =   930
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Tag             =   "5"
      Top             =   2160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   186712065
      CurrentDate     =   38052
   End
   Begin VB.Label Label34 
      Caption         =   "Ejercicio"
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Mov. bancario :"
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
      TabIndex        =   26
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblIdDoc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   5385
      TabIndex        =   24
      Top             =   180
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2475
      Left            =   90
      Top             =   75
      Width           =   8925
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha:"
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
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Importe:"
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
      Index           =   2
      Left            =   225
      TabIndex        =   12
      Top             =   1770
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto:"
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
      Index           =   1
      Left            =   225
      TabIndex        =   11
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta Origen:"
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
      Left            =   225
      TabIndex        =   10
      Top             =   495
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Cuenta Destino:"
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
      Index           =   0
      Left            =   225
      TabIndex        =   7
      Top             =   930
      Width           =   1455
   End
End
Attribute VB_Name = "FrmTransfBanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4


Dim midDoc As Long
Dim modifico As Boolean
Dim Ope As String


'Private Sub LimpioImputacion()
'    txtcuentacod = ""
'    txtcuenta = ""
'    txtconc = ""
'    txtvalor = ""
'End Sub

'Private Sub cmbcuenta_Click()
'    FrmHelp.Show
'    CargarHelpCuentas "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "Cuentas"
'End Sub

Private Sub cmbcuentad_Click()
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
    cargar = "Cuentad"
End Sub

Private Sub cmbcuentao_Click()
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
    cargar = "Cuentao"
End Sub


Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta

    Dim maximobanc1 As Long, maximocaja As Long, maximobanc2 As Long  '  1 movicaja, 2 movibanc
', x As Long
    Dim valorcuentacon1 As String, valorcuentacon2 As String
    Dim asie As New Asiento, i As Long
'Dim rs As New ADODB.Recordset
   
    If s2n(txtimporte) = 0 Then 's2n(txtcodctao) = 0 Or
        che "Faltan ingresar datos"
        Exit Sub
    End If
        
    If Ope <> "A" Then Exit Sub
    
    
    
    If Trim(txtcodctad) > "" And Trim(txtcodctao) > "" Then ' entre bancos   txtcodctao
        'If Trim(txtcodctad.Text) = "" Then
        '    MsgBox "Debe ingresar la cuenta de destino."
        '    Exit Sub
        'End If
        maximobanc1 = nuevoCodigo("movibanc", "movbanco")
        movbanc = maximobanc1
        maximobanc2 = maximobanc1 + 1
        maximocaja = nuevoCodigo("movicaja", "movimiento")
        valorcuentacon1 = verCuentaContableBanco(val(txtcodctao))
        valorcuentacon2 = verCuentaContableBanco(val(txtcodctad)) 'IIf(txtcodctad = "", txtcodctao, verCuentaContableBanco(Val(txtcodctad)))
        
        DE_BeginTrans
            
            midDoc = NuevoDocumento("TRB", maximobanc1, 0, 0)
            lblIDDOC = midDoc
            
            asie.nuevo "Transferencia bancaria ", dtFecha, "TRB"
            asie.AgregarItem valorcuentacon1, 0, s2n(txtimporte)
            asie.AgregarItem valorcuentacon2, s2n(txtimporte), 0
 
            
            abmTransfMoviBanc abmAlta, val(txtcodctao), "S", Trim(txtconcepto), dtFecha, "T", s2n(txtimporte), maximobanc1, midDoc
            abmTransfMoviBanc abmAlta, val(txtcodctad), "E", Trim(txtconcepto), dtFecha, "T", s2n(txtimporte), maximobanc2, midDoc
            

            abmTransfMoviCaja abmAlta, 0, maximocaja, "T", "E", s2n(txtimporte), "TRANSF. " & Trim(txtconcepto), _
            dtFecha, valorcuentacon1, maximobanc1, midDoc
            
            asie.Grabar midDoc, , leerEjercicioId(cboEjercicio)
        
            
        DE_CommitTrans
        che "Transferencia grabada "
        
    ElseIf Trim(txtcodctao) > "" Then            'ACA banco a cuentascontables
        
        If s2n(uCuentas.Total) <> s2n(txtimporte) Then
            che "no coinciden los montos"
            Exit Sub
        End If
        
'        If Trim(txtcodctad.Text) = "" Then
'            MsgBox "Debe ingresar al menos la cuenta de destino."
'            Exit Sub
'        End If
        
        maximobanc1 = nuevoCodigo("movibanc", "movbanco")
        movbanc = maximobanc1
'        maximobanc2 = maximobanc1 + 1
        maximocaja = nuevoCodigo("movicaja", "movimiento")
        valorcuentacon1 = verCuentaContableBanco(val(txtcodctao))
'        valorcuentacon2 = verCuentaContableBanco(Val(txtcodctad))
        
        DE_BeginTrans
            
            midDoc = NuevoDocumento("TRB", maximobanc1, 0, 0)
            lblIDDOC = midDoc
            
            asie.nuevo "Transferencia bancaria ", dtFecha, "TRB"
            asie.AgregarItem valorcuentacon1, 0, s2n(txtimporte) 'valorcuentacon2
 '           asie.AgregarItem valorcuentacon2, s2n(txtimporte), 0

            For i = 1 To uCuentas.rows
                asie.AgregarItem uCuentas.imCuenta(i), uCuentas.imMonto(i), 0
            Next i
            
            abmTransfMoviBanc abmAlta, val(txtcodctao), "S", Trim(txtconcepto), dtFecha, "T", s2n(txtimporte), maximobanc1, midDoc
'            abmTransfMoviBanc abmAlta, Val(txtcodctad), "E", Trim(txtconcepto), dtfecha, "E", s2n(txtimporte), maximobanc2, midDoc
            
            abmTransfMoviCaja abmAlta, 0, maximocaja, "T", "E", s2n(txtimporte), "TRANSF. " & Trim(txtconcepto), _
                         dtFecha, valorcuentacon1, maximobanc1, midDoc
            'valorcuentacon1
            
            asie.Grabar midDoc, , leerEjercicioId(cboEjercicio)
                   
        DE_CommitTrans
        che "Transferencia grabada "
        
    ElseIf Trim(txtcodctad) > "" Then
        If s2n(uCuentas.Total) <> s2n(txtimporte) Then
            che "no coinciden los montos"
            Exit Sub
        End If
        
'        If Trim(txtcodctad.Text) = "" Then
'            MsgBox "Debe ingresar al menos la cuenta de destino."
'            Exit Sub
'        End If
        
        maximobanc1 = nuevoCodigo("movibanc", "movbanco")
        movbanc = maximobanc1
        maximobanc2 = maximobanc1 'maximobanc1 + 1
        maximocaja = nuevoCodigo("movicaja", "movimiento")
'        valorcuentacon1 = verCuentaContableBanco(Val(txtcodctao))
        valorcuentacon2 = verCuentaContableBanco(val(txtcodctad))
        
        DE_BeginTrans
            
            midDoc = NuevoDocumento("TRB", maximobanc1, 0, 0)
            lblIDDOC = midDoc
            
            asie.nuevo "Transferencia bancaria ", dtFecha, "TRB"
'            asie.AgregarItem valorcuentacon1, 0, s2n(txtimporte) 'valorcuentacon2
            asie.AgregarItem valorcuentacon2, s2n(txtimporte), 0

            For i = 1 To uCuentas.rows
                asie.AgregarItem uCuentas.imCuenta(i), 0, uCuentas.imMonto(i)
            Next i
            
'            abmTransfMoviBanc abmAlta, Val(txtcodctao), "S", Trim(txtconcepto), dtfecha, "E", s2n(txtimporte), maximobanc1, midDoc
            abmTransfMoviBanc abmAlta, val(txtcodctad), "E", Trim(txtconcepto), dtFecha, "T", s2n(txtimporte), maximobanc2, midDoc
            
            abmTransfMoviCaja abmAlta, 0, maximocaja, "T", "I", s2n(txtimporte), "TRANSF. " & Trim(txtconcepto), _
                         dtFecha, valorcuentacon1, maximobanc1, midDoc
            'valorcuentacon1
            
            asie.Grabar midDoc, , leerEjercicioId(cboEjercicio)
                   
        DE_CommitTrans
        che "Transferencia grabada "
        
    End If
    
    ImprimirTransferenciaBanc midDoc
    
    LimpioControles
'    InicioGrilla
    HabilitoControles (False)
'    habilitogrillaenable (False)
    Call Habilitobotones(True, True, True, True, True, True)
    
fin:
    Exit Sub
UFAalta:
    ufa "Prg: Fallo el alta", "alta transferencia"
    DE_RollbackTrans
    Resume fin
End Sub


''Private Sub cmdAceptar_Click()
''Dim rs As New ADODB.Recordset
''Dim maximobanc1 As Long, maximobanc2 As Long, maximocaja As Long, x As Long
'''Dim fecha As String, fechamovi As String,
''Dim valorcuentacon1 As String, valorcuentacon2 As String
''
'''If txtTotal <> txtImporte Then
'''    MsgBox "No coincide el importe ingresado con el importe total"
'''    Exit Sub
'''End If
''
''If txtcodctao <> "" And txtImporte <> "" Then
''
'''        fechamovi = Month(dtFecha) & "/" & Day(dtFecha) & "/" & Year(dtFecha)
'''        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
''
''        If Ope = "A" Then
''            rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''            If Not IsNull(rs!maxcodigo) Then
''                maximobanc1 = rs!maxcodigo + 1
''            Else
''                maximobanc1 = 1
''            End If
''            rs.Close
''            Set rs = Nothing
''
''            DataEnvironment1.dbo_TRANSFMOVIBANC "A", Val(txtcodctao), "S", Trim(TxtConcepto), dtfecha, "E", s2n(txtImporte), _
''                        maximobanc1, Date, UsuarioSistema!codigo, 0, 0
''
''            rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''            If Not IsNull(rs!maxcodigo) Then
''                maximobanc2 = rs!maxcodigo + 1
''            Else
''                maximobanc2 = 1
''            End If
''            rs.Close
''            Set rs = Nothing
''
''            DataEnvironment1.dbo_TRANSFMOVIBANC "A", Val(txtcodctad), "E", Trim(TxtConcepto), dtfecha, "E", s2n(txtImporte), _
''            maximobanc2, Date, UsuarioSistema!codigo, 0, 0
''
''            rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''            If Not IsNull(rs!maxcodigo) Then
''                maximocaja = rs!maxcodigo + 1
''            Else
''                maximocaja = 1
''            End If
''            rs.Close
''            Set rs = Nothing
''
''            rs.Open "select cuenta_con from Ctasbank where codigo = " & Val(txtcodctao) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''            If Not rs.EOF Then
''                valorcuentacon1 = rs!cuenta_con
''            End If
''            rs.Close
''            Set rs = Nothing
''
'''            DataEnvironment1.dbo_TRANSFMOVICAJA "A", 0, maximocaja, "T", "E", s2n(txtImporte), "TRANSF. " & Trim(TxtConcepto), _
'''            dtfecha, valorcuentacon1, maximobanc1, Date, UsuarioSistema!codigo, 0, 0
''
''            abmTransfMoviCaja abmAlta, 0, maximocaja,  "T", "E", s2n(txtImporte), "TRANSF. " & Trim(TxtConcepto), _
''                         dtfecha, valorcuentacon1, maximobanc1,
''
''
''            rs.Open "select cuenta_con from Ctasbank where codigo = " & Val(txtcodctad) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''            If Not rs.EOF Then
''                valorcuentacon2 = rs!cuenta_con
''            End If
''            rs.Close
''            Set rs = Nothing
''
'''            If txtcodctad <> "" Then
'''                DataEnvironment1.dbo_TRANSFDETALLE "A", maximocaja, s2n(txtimporte), valorcuentacon2, "TRANSF. " & Trim(txtConcepto), "TR", _
'''                dtfecha
'''            Else
'''                For x = 1 To grilla.rows - 1
'''                    DataEnvironment1.dbo_TRANSFDETALLE "A", maximocaja, s2n(grilla.TextMatrix(x, 3)), Val(grilla.TextMatrix(x, 0)), IIf(Trim(grilla.TextMatrix(x, 2)) <> "", Trim(grilla.TextMatrix(x, 2)), Trim(txtConcepto)), "TR", _
'''                    dtfecha
'''                Next
'''            End If
''        Else
''            'ACA IRIA TODO LO DE MODIFICACIONES
''        End If
''
''    LimpioControles
''    InicioGrilla
''    HabilitoControles (False)
'''    habilitogrilla (False)
''    habilitogrillaenable (False)
''    Call Habilitobotones(True, True, True, True, True, True)
''    ImprimirTransferenciaBanc
''Else
''    MsgBox "Faltan ingresar datos"
''End If
''End Sub
'
'Sub InicioGrilla()
'    grilla.Clear
'    'grilla.ColWidth(1) = 1700
'    grilla.TextMatrix(0, 0) = "Cuenta"
'    grilla.TextMatrix(0, 1) = "Descripción"
'    grilla.TextMatrix(0, 2) = "Concepto"
'    grilla.TextMatrix(0, 3) = "Importe"
'    grilla.rows = 2
'End Sub

Private Sub cmdBuscar_Click()
'    cargar = "Movibanc"
'    FrmHelp.Show
'    CargarHelpMovibanc "MOVIBANC", "Mov. Bancario", "Fecha - Importe", "Movbanco", "Fecha", "Importe", "", "Movbanco"
'    FrmHelp.Tag = Me.Name
'    Call Habilitobotones(True, False, True, True, True, True)
    Dim resu As String


    With frmBuscar
    'resu = .MostrarSql("select movbanco, importe, fecha, descripcion, rd.iddoc as _H_iddoc from movibanc as mb inner join registrodocumentos as rd on rd.iddoc = mb.iddoc where rd.iddoc >0 and operacion = 'E'  and rd.activo = 1 order by movbanco")
'    resu = "select movbanco, importe, fecha, descripcion, rd.iddoc as _H_iddoc from movibanc as mb inner join registrodocumentos as rd on rd.iddoc = mb.iddoc where rd.activo = 1 and (rd.iddoc >0 and operacion = 'E') or (operacion='S' and (select count(iddoc) from movibanc as m where m.iddoc=mb.iddoc)=1) order by movbanco"
    resu = .MostrarSql("select movbanco, importe, fecha, descripcion, rd.iddoc as _H_iddoc from movibanc as mb inner join registrodocumentos as rd on rd.iddoc = mb.iddoc where rd.activo = 1 and (rd.iddoc >0 and operacion = 'E') or (operacion='S' and (select count(iddoc) from movibanc as m where m.iddoc=mb.iddoc and m.activo=1)=1) order by movbanco")
    If resu > "" Then
        midDoc = .resultado(5)
        lblIDDOC = .resultado(5)
        movbanc = .resultado(1)
        txtimporte = .resultado(2)
        dtFecha = .resultado(3)
        txtconcepto = .resultado(4)
        txtcodctad = obtenerDeSQL("select cuenta from movibanc where operacion = 'E' and iddoc = " & midDoc & " and movbanco=" & movbanc)
        'txtcuentad = obtenerDeSQL("select b.Descripcion+' - '+c.numero from ctasbank c inner join bancosgrales b on b.codigo=c.banco where c.codigo=" & txtcodctad)
        If txtcodctad = "" Or txtcodctad = "0" Then
            txtcuentad.Text = "Sin Asignar"
            txtcodctad = 0
        Else
            txtcuentad = obtenerDeSQL("select b.Descripcion+' - '+c.numero from ctasbank c inner join bancosgrales b on b.codigo=c.banco where c.codigo=" & txtcodctad)
        End If
        
        If txtcodctad = 0 Then
            txtcodctao = obtenerDeSQL("select cuenta from movibanc where operacion='S' and movbanco = '" & movbanc & "' and  iddoc = " & midDoc)
        Else
            txtcodctao = obtenerDeSQL("select cuenta from movibanc where operacion='S' and movbanco = '" & movbanc - 1 & "' and  iddoc = " & midDoc)
        End If
        If txtcodctao = "" Or txtcodctao = "0" Then
            txtcuentao.Text = "Sin Asignar"
            txtcodctao = 0
        Else
            txtcuentao = obtenerDeSQL("select b.Descripcion+' - '+c.numero from ctasbank c inner join bancosgrales b on b.codigo=c.banco where c.codigo=" & txtcodctao)
        End If
                
        cmdbuscar.enabled = True
        cmdnuevo.enabled = True
        cmdeliminar.enabled = True
        cmdImprimir.enabled = True
        cmdAceptar.enabled = False
        cmdcancelar.enabled = True
    
    End If
    End With
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
'    Limpiotextosgrilla
'    InicioGrilla
    HabilitoControles (False)
'    habilitogrilla (False)
'    habilitogrillaenable (False)
    Call Habilitobotones(True, True, False, False, False, True)
    cmdImprimir.enabled = False
End Sub

'Private Sub cmdcargar_Click()
'Dim Valor As Double, totalgrilla
'
'    If txtvalor <> "" Then
'        If modifico = False Then
'            Valor = s2n(txtvalor)
'            If (Valor <= s2n(txtImporte)) And (Valor + s2n(txtTotal) <= s2n(txtImporte)) Then
'                If txtTotal <> "" Then
'                    If s2n(txtTotal) + Valor <= s2n(txtImporte) Then
'                        CargogrillaTotal
'                    Else
'                        MsgBox "Con este valor el importe total serìa superado", vbInformation
'                    End If
'                Else
'                    CargogrillaTotal
'                End If
'                Limpiotextosgrilla
'                If txtcuentacod.Enabled = True And txtImporte <> txtTotal Then
'                    txtcuentacod.SetFocus
'                End If
'            Else
'                MsgBox "El valor a ingresar no puede superar al importe original"
'                txtvalor.SetFocus
'            End If
'        Else
'            totalgrilla = sumogrilla()
'            If totalgrilla - s2n(grilla.TextMatrix(grilla.row, 3)) + s2n(txtvalor) <= s2n(txtImporte) Then
'                grilla.TextMatrix(grilla.row, 0) = txtcuentacod
'                grilla.TextMatrix(grilla.row, 1) = txtCuenta
'                grilla.TextMatrix(grilla.row, 2) = txtconc
'                grilla.TextMatrix(grilla.row, 3) = txtvalor
'                txtTotal = sumogrilla()
''                LimpioImputacion
'                modifico = False
'                grilla.SetFocus
'            Else
''                MsgBox "El valor a ingresar no puede superar al total"
''                txtvalor.SetFocus
''            End If
'        End If
'    Else
'        MsgBox "Debe ingresar un valor"
'        txtvalor.SetFocus
'    End If
'
'End Sub

'Function sumogrilla() As Double
'Dim x As Long
'Dim Total As Double
'
'    For x = 1 To grilla.rows - 1
'        Total = Total + s2n(grilla.TextMatrix(x, 3))
'    Next
'    sumogrilla = Total
'
'End Function
'
'Private Sub MuestroGrilla()
'    txtcuentacod = grilla.TextMatrix(grilla.row, 0)
'    txtcuenta = grilla.TextMatrix(grilla.row, 1)
'    txtconc = grilla.TextMatrix(grilla.row, 2)
'    txtvalor = grilla.TextMatrix(grilla.row, 3)
'End Sub

Private Sub cmdeliminar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAbaja
    
    If Not confirma("Esta seguro de querer eliminar este registro") Then Exit Sub
    If midDoc = 0 Then
        che " prg: sin id para borrar "
        Exit Sub
    End If
    
    DE_BeginTrans
        
            BorroDocumento midDoc
            abmTransfMoviBanc abmbaja, 0, "", "", Date, "", 0, val(movbanc), midDoc  ' baja los 2
            abmTransfMoviCaja abmbaja, 0, 0, "", "", 0, "", Date, "", 0, midDoc
            grabaBitacora "B", midDoc, "movicaja movibanc"
        
    DE_CommitTrans
    che "eliminado"
        
    Call Habilitobotones(True, True, False, False, False, True)
    LimpioControles
    HabilitoControles (False)
    
fin:
    Exit Sub
UFAbaja:
    DE_RollbackTrans
    ufa "error en baja", "eliminar transf " & midDoc
    Resume fin
End Sub

'Private Sub cmdGrilla_Click()
'
'End Sub

'Private Sub cmdeliminar_Click()
'Dim rs As New ADODB.Recordset
'Dim Fecha As Variant
'Dim mensaje As String
'
'    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
'    If mensaje = 6 Then
'        Fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'
'        DataEnvironment1.dbo_TRANSFMOVIBANC "B", 0, "", "", 0, "", 0, Val(movbanc), 0, 0, Fecha, UsuarioSistema!codigo
'
'        DataEnvironment1.dbo_TRANSFMOVIBANC "B", 0, "", "", 0, "", 0, (Val(movbanc) + 1), 0, 0, Fecha, UsuarioSistema!codigo
'
'        rs.Open "select movimiento from MOVICAJA where movbanco = " & Val(movbanc) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'        If Not rs.EOF Then
'            DataEnvironment1.dbo_TRANSFMOVICAJA "B", 0, 0, "", "", 0, "", 0, "", Val(movbanc), 0, 0, Fecha, UsuarioSistema!codigo
'            DataEnvironment1.dbo_TRANSFDETALLE "B", rs!movimiento, 0, "", "", "", 0
'        End If
'        rs.Close
'        Set rs = Nothing
'
'        DataEnvironment1.dbo_GRABARBITACORA Val(movbanc), "Movibanc", UsuarioSistema!codigo, Fecha, Time, "B"
'
'        Call Habilitobotones(True, True, False, False, False, True)
'        LimpioControles
'        HabilitoControles (False)
'        InicioGrilla
''        habilitogrilla (False)
'    End If
'
'End Sub

'Private Sub cmdmodificar_Click()
'    Ope = "M"
'    HabilitoControles (True)
'    Call Habilitobotones(True, False, False, True, True, True)
'    If grilla.Visible = True Then
'        habilitogrillaenable (True)
'    End If
'End Sub

Private Sub cmdImprimir_Click()
    ImprimirTransferenciaBanc midDoc
End Sub

Private Sub cmdnuevo_Click()
    LimpioControles
    Call Habilitobotones(False, False, False, False, True, True)
    cmdImprimir.enabled = False
    HabilitoControles (True)
    Ope = "A"
    modifico = False
    txtcodctao.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub dtfecha_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uCuentas.enabled = False
    
    cmdbuscar.enabled = True
    cmdnuevo.enabled = True
    cmdeliminar.enabled = False
    cmdImprimir.enabled = False
    cmdAceptar.enabled = False
    cmdcancelar.enabled = False
    
    Dim EjerA As New ADODB.Recordset
    EjerA.Open "select * from ejercicio", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    EjerA.MoveFirst
    While Not EjerA.EOF
        cboEjercicio.AddItem EjerA!denominacion 'EjerA!idejercicio
        EjerA.MoveNext
    Wend
    cboEjercicio = leerEjercicioDenominacion() ' mIdEjercicioActivo
    If UsuarioActual() <> 19 Then
        cboEjercicio.Visible = False
        Label34.Visible = False
    End If
    
End Sub

'Private Sub grilla_Click()
'    modifico = True
'    MuestroGrilla
'End Sub

Private Sub txtcodctad_GotFocus()
    txtcodctad.SelStart = 0
    txtcodctad.SelLength = Len(txtcodctad.Text)
End Sub

Private Sub txtcodctad_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcodctao_GotFocus()
    txtcodctao.SelStart = 0
    txtcodctao.SelLength = Len(txtcodctao.Text)
End Sub

Private Sub txtcodctao_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcodctao_LostFocus()
    If IsNumeric(txtcodctao) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodctao) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuentao = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
        
        If txtcuentao = "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodctao.SetFocus
        Else
            cargar = "Cuentao"
            CargarDatos
        End If
    Else
        If txtcodctao <> "" Then
            MsgBox "Código de cuenta incorrecto"
            txtcodctao = "0"
            txtcodctao.SetFocus
        End If
    End If
End Sub

Private Sub txtcodctad_LostFocus()
    If txtcodctad <> "" Then
        If IsNumeric(txtcodctad) Then
            Dim rs As New ADODB.Recordset
            
            rs.Open "select * from Ctasbank where codigo = " & val(txtcodctad) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtcuentad = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
            End If
            rs.Close
            Set rs = Nothing
            
            If txtcuentad = "" Then
                MsgBox "Codigo de cuenta incorrecto"
                txtcodctad.SetFocus
            Else
                cargar = "Cuentad"
'                habilitogrilla (False)
'                habilitogrillaenable (False)
                CargarDatos
            End If
        Else
            If txtcodctad <> "" Then
                MsgBox "Codigo de cuenta incorrecto"
                txtcodctad = "0"
                txtcodctad.SetFocus
            End If
        End If
    Else
        txtcuentad = ""
'        InicioGrilla
'        habilitogrilla (True)
    End If
End Sub

Public Sub CargarDatos()
Dim rs As New ADODB.Recordset
Dim codigo As Long

    codigo = val(Trim(Me.Tag))
       
'    If cargar = "Cuentas" Then
'        If txtcuentacod = "" Then
'            txtcuentacod = Trim(str(codigo))
'        End If
'        If Not noestaenlagrilla(txtcuentacod, grilla) And esimputable(txtcuentacod) Then
'            rs.Open "select * from Cuentas where codigo = " & Val(txtcuentacod) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If Not rs.EOF Then
'                txtcuentacod = rs!codigo
'                txtcuenta = rs!descripcion
'                txtconc.SetFocus
'            End If
'            rs.Close
'            Set rs = Nothing
'        Else
'            MsgBox "El concepto ya se encuentra cargado"
'            txtcuentacod = ""
'            txtcuentacod.SetFocus
'        End If
'    End If
       
       
    If cargar = "Cuentao" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodctao) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodctao = rs!codigo
            txtcuentao = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Cuentad" Then
        If txtcodctad <> txtcodctao Then
            rs.Open "select * from Ctasbank where codigo = " & val(txtcodctad) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtcodctad = rs!codigo
                txtcuentad = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
            End If
            rs.Close
            Set rs = Nothing
        Else
            MsgBox "La cuenta de destino no puede ser la misma que la de origen"
            txtcodctad = ""
            txtcodctad.SetFocus
        End If
    End If
        
    If cargar = "Movibanc" Then
        rs.Open "select Movibanc.*, Ctasbank.banco, Ctasbank.numero from Movibanc inner join Ctasbank on Movibanc.cuenta = Ctasbank.codigo where Movibanc.movbanco = " & codigo & " and Movibanc.activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodctao = rs!Cuenta
            txtcuentao = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
            txtconcepto = rs!DESCRIPCION
            txtimporte = rs!Importe
            dtFecha = rs!Fecha
            
            midDoc = rs!iddoc: lblIDDOC = midDoc
            
        End If
        rs.Close
        Set rs = Nothing
            
        codigo = codigo + 1
            
        rs.Open "select Movibanc.*, Ctasbank.banco, Ctasbank.numero from Movibanc inner join Ctasbank on Movibanc.cuenta = Ctasbank.codigo where Movibanc.movbanco = " & codigo & " and Movibanc.activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodctad = rs!Cuenta
            txtcuentad = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
'        Else
'            codigo = codigo - 1
'            rs.Close
'            Set rs = Nothing
'            rs.Open "select movimiento from Movicaja where movbanco = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If Not rs.EOF Then
'                Cargogrilla (rs!movimiento)
'            End If
        End If
        rs.Close
        Set rs = Nothing
    End If
        
End Sub

'Private Sub habilitogrillaenable(habilito As Boolean)
''    Label11.Enabled = habilito
''    txtcuentacod.Enabled = habilito
''    cmbcuenta.Enabled = habilito
''    Label7.Enabled = habilito
''    txtconc.Enabled = habilito
''    Label9.Enabled = habilito
''    txtvalor.Enabled = habilito
''    cmdcargar.Enabled = habilito
''    grilla.Enabled = habilito
''    cmbeliminofila.Enabled = habilito
'End Sub

'Sub habilitogrilla(habilito As Boolean)
''    Label11.Visible = habilito
''    txtcuentacod.Visible = habilito
''    cmbcuenta.Visible = habilito
''    txtCuenta.Visible = habilito
''    Label7.Visible = habilito
''    txtconc.Visible = habilito
''    Label9.Visible = habilito
''    txtvalor.Visible = habilito
''    cmdcargar.Visible = habilito
''    grilla.Visible = habilito
''    cmbeliminofila.Visible = habilito
''    Label8.Visible = habilito
''    TxtTotal.Visible = habilito
'End Sub

Sub LimpioControles()
    '    txtcodctao = ""
    '    txtcuentao = ""
    '    txtcodctad = ""
    '    txtcuentad = ""
    '    txtconcepto = ""
    '    txtimporte = "0"
    '    txtcuentacod = ""
    '    txtcuenta = ""
    '    txtconc = ""
    '    txtvalor = "0"
    '    txttotal = "0"
    '    movbanc = ""
    '    txttotal = "0"
    FrmBorrarTxt Me
    
    Ope = ""
    cargar = ""
    midDoc = 0
    lblIDDOC = ""
    dtFecha = Date
    uCuentas.Borrar
End Sub

Private Sub HabilitoControles(habilito As Boolean)
    txtcodctao.enabled = habilito
    txtcodctad.enabled = habilito
    txtconcepto.enabled = habilito
    txtimporte.enabled = habilito
    dtFecha.enabled = habilito
    cmbcuentao.enabled = habilito
    cmbcuentad.enabled = habilito
End Sub

'Private Sub txtconc_Change()
'Dim i As Long
'    txtconc.Text = UCase(txtconc.Text)
'    i = Len(txtconc.Text)
'    txtconc.SelStart = i
'End Sub
'
'Private Sub txtconc_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub txtconcepto_Change()
Dim i As Long
    txtconcepto.Text = UCase(txtconcepto.Text)
    i = Len(txtconcepto.Text)
    txtconcepto.SelStart = i
End Sub

Private Sub txtConcepto_GotFocus()
    txtconcepto.SelStart = 0
    txtconcepto.SelLength = Len(txtconcepto.Text)
    
'    If txtcodctad = "" Then
'        InicioGrilla
''        habilitogrilla (True)
'    End If
End Sub

'Private Sub txtconcepto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtcuentacod_GotFocus()
'Dim rs As New ADODB.Recordset
'
'    rs.Open "select dato_fijo from datos", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs.EOF Then
'        If rs!DATO_FIJO = 7 Then
'            txtcuentacod = "1"
'            txtcuentacod.Enabled = False
'            txtcuenta = "COMPRAS"
'            txtconc = "COMPRAS"
'            txtconc.Enabled = False
'            txtvalor = txtimporte
'            txtvalor.Enabled = False
'            cmbcuenta.Enabled = False
'            cmdcargar.Enabled = False
'            cmbeliminofila.Enabled = False
'            CargogrillaTotal
'        End If
'    End If
'    rs.Close
'
'End Sub

Private Sub txtcuentacod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

'Private Sub txtcuentacod_LostFocus()
'    If IsNumeric(txtcuentacod) Then
'        If Not noestaenlagrilla(txtcuentacod, grilla) And esimputable(Val(txtcuentacod)) Then
'            txtcuenta = ObtenerDescripcion("Cuentas", Val(txtcuentacod))
'            If txtcuenta = "" Then
'                MsgBox "Codigo de cuenta incorrecto"
'                txtcuentacod.SetFocus
'            Else
'                cargar = "Cuentas"
'                CargarDatos
'            End If
'        Else
'            MsgBox "El concepto ya se encuentra cargado o la cuenta no es imputable"
'            txtcuentacod = ""
'            txtcuentacod.SetFocus
'        End If
'    Else
'        If txtcuentacod <> "" Then
'            MsgBox "La cuenta es incorrecta"
'            txtcuentacod = "0"
'            txtcuentacod.SetFocus
'        End If
'    End If
'End Sub

Private Sub txtimporte_GotFocus()
    txtimporte.SelStart = 0
    txtimporte.SelLength = Len(txtimporte.Text)
    
'    If txtcodctad = "" And grilla.Visible = False Then
'        InicioGrilla
''        habilitogrilla (True)
'    End If
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtimporte_LostFocus()
    If IsNumeric(txtimporte) Then
        txtimporte = s2n(txtimporte)
        If Trim(txtcodctad) = "" Or Trim(txtcodctao) = "" Then
            'habilitogrillaenable (True)
            uCuentas.enabled = True
            uCuentas.Total_a_Imputar = s2n(txtimporte)
        End If
    Else
        If txtimporte <> "" Then
            MsgBox "Debe ingresar un importe"
            txtimporte = "0"
            txtimporte.SetFocus
        End If
    End If
End Sub

'Private Sub cmbeliminofila_Click()
'    If grilla.TextMatrix(grilla.row, grilla.col) <> "" Then
'        If grilla.rows > 1 Then
'            txtTotal = s2n(txtTotal) - s2n(grilla.TextMatrix(grilla.row, 3))
'            If grilla.rows = 2 Then
'                grilla.TextMatrix(1, 0) = ""
'                grilla.TextMatrix(1, 1) = ""
'                grilla.TextMatrix(1, 2) = ""
'                grilla.TextMatrix(1, 3) = ""
'            Else
'                grilla.RemoveItem (grilla.row)
'            End If
'        Else
'            MsgBox "No hay productos para eliminar o no ha seleccionado ninguno de ellos"
'        End If
'    End If
'End Sub

'Private Sub Cargogrilla(movimiento As Long)
'Dim rs1 As New ADODB.Recordset
'
'    rs1.Open "select * from DetalleMovcajas where movimiento = " & movimiento & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs1.EOF Then
''        InicioGrilla
''        habilitogrilla (True)
'        grilla.rows = 2
'        grilla.row = 0
'        While Not rs1.EOF
'            grilla.row = grilla.row + 1
'            grilla.TextMatrix(grilla.row, 0) = rs1!cuenta
'            grilla.TextMatrix(grilla.row, 1) = ObtenerDescripcion("Cuentas", Val(rs1!cuenta))
'            grilla.TextMatrix(grilla.row, 2) = rs1!concepto
'            grilla.TextMatrix(grilla.row, 3) = rs1!Importe
'            If txtTotal <> "" Then
'                txtTotal = s2n(txtTotal) + s2n(rs1!Importe)
'            Else
'                txtTotal = s2n(rs1!Importe)
'            End If
'            rs1.MoveNext
'            If Not rs1.EOF Then
'                grilla.rows = grilla.rows + 1
'            End If
'        Wend
'    End If
'    rs1.Close
'    Set rs1 = Nothing
'End Sub

'Private Sub CargogrillaTotal()
'Dim Valor As Double
'
'    If grilla.rows = 2 Then
'        grilla.Row = 1
'        grilla.Col = 0
'        If Trim(grilla.Text) = "" Then
'            grilla.Row = 1
'            grilla.Col = 0
'            grilla.Text = txtcuentacod
'            grilla.Col = 1
'            grilla.Text = txtcuenta
'            grilla.Col = 2
'            grilla.Text = txtconc
'            grilla.Col = 3
'            grilla.Text = txtvalor
'        Else
'            grilla.AddItem txtcuentacod & Chr(9) & txtcuenta & Chr(9) & txtconc & Chr(9) & txtvalor
'        End If
'    Else
'        grilla.AddItem txtcuentacod & Chr(9) & txtcuenta & Chr(9) & txtconc & Chr(9) & txtvalor
'    End If
'    If TxtTotal <> "" Then
'        Valor = s2n(txtvalor)
'        TxtTotal = s2n(TxtTotal) + Valor
'    Else
'        TxtTotal = s2n(txtvalor)
'    End If
'    If TxtTotal = txtimporte Then
'        MsgBox "El detalle ha sido completado"
''        habilitogrillaenable (False)
'    End If
'End Sub
'
'Private Sub Limpiotextosgrilla()
'    txtcuentacod = ""
'    txtcuenta = ""
'    txtconc = ""
'    txtvalor = ""
'End Sub

Sub Habilitobotones(busco As Boolean, nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
    cmdbuscar.enabled = busco
    cmdnuevo.enabled = nuevo
'    cmdmodificar.Enabled = modifico
    cmdeliminar.enabled = elimino
    cmdAceptar.enabled = acepto
    cmdcancelar.enabled = Cancelo
End Sub

'Private Sub txtvalor_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub

'Private Sub txtvalor_LostFocus()
'    If IsNumeric(txtvalor) Then
''        InicioGrilla
'        If grilla.Visible = False Then
''            habilitogrilla (True)
'        End If
'        habilitogrillaenable (True)
'        txtvalor = s2n(txtvalor)
'    Else
'        If txtvalor <> "" Then
'            MsgBox "Debe ingresar un importe"
'            txtvalor = "0"
'            txtvalor.SetFocus
'        End If
'    End If
'End Sub

'3/12/4 fecha string x date
'
