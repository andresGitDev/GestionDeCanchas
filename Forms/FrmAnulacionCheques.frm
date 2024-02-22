VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAnulacionCheques 
   Caption         =   "Anulación de Cheques Propios"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   Icon            =   "FrmAnulacionCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "5"
      Top             =   5280
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "4"
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmbver 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Cheques"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      Top             =   240
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Tag             =   "0"
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   450
      _Version        =   393216
      Format          =   82247681
      CurrentDate     =   38052
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmAnulacionCheques.frx":08CA
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      MergeCells      =   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).BandIndent=   10
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   5055
      Left            =   120
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label lblcartel 
      Caption         =   """Click"" para marcar el cheque a anular"
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
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de Débito"
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
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "FrmAnulacionCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4 '****************fecha

Private Sub cmbver_Click()
    InicioGrilla
    Cargogrilla
    lblcartel.Visible = True
End Sub

Private Sub cmdaceptar_Click()
Dim x As Long, maximo As Long, d_q_cuenta
Dim rs As New ADODB.Recordset
Dim mensaje As String, cad_update As String
    
    mensaje = MsgBox("Esta seguro de realizar la operación", vbYesNo, "Atencion")
    If mensaje = 6 Then
        rs.Open "select max(movbanco) as maxmov from movibanc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Or rs!maxmov <> Null Then
            maximo = rs!maxmov
        Else
            maximo = 0
        End If
        rs.Close
        Set rs = Nothing
        
        For x = 1 To Grilla.rows - 1
            If Grilla.TextMatrix(x, 5) = "        X" Then
                maximo = maximo + 1
                
                
                d_q_cuenta = obtenerDeSQL("select cuentabancaria from Chq_comp where codigo =" & Val(Grilla.TextMatrix(x, 0)))
                If d_q_cuenta = 0 Then
                    If MsgBox("El cheque " & Val(Grilla.TextMatrix(x, 0)) & " no pertenece a ninguna cuenta bancaria." & Chr(13) & "Por favor indique una a continuacion, gracias.", vbInformation + vbYesNo) = vbYes Then
                        d_q_cuenta = frmBuscar.MostrarSql("select c.codigo as [CODIGO], c.banco as [BANCO - Nº],b.descripcion as  [NOMBRE  ],c.numero as [CUENTA - Nº] from ctasbank c inner join bancosgrales b on c.banco=b.codigo where c.activo=1", , "Cuentas bancarias", " - ")
                        DataEnvironment1.Sistema.Execute "update Chq_comp set cuentabancaria = " & d_q_cuenta & " where codigo = " & Val(Grilla.TextMatrix(x, 0))
                    Else
                        d_q_cuenta = 0
                        MsgBox "Debe ingresar una cuenta para este cheque.", vbInformation
                        Exit Sub
                    End If
                End If
                    
                DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", d_q_cuenta, "P", "Rechazo de Cheque", dtFecha, "C" _
                  , Val(Grilla.TextMatrix(x, 0)), s2n(Grilla.TextMatrix(x, 2)), nuevoCodigo("movibanc", "movbanco"), 0, Date, UsuarioSistema!codigo
                
                cad_update = "UPDATE CHQ_COMP SET ESTADO='N', FECHA_OPERACION = " & ssFecha(dtFecha) & " WHERE CODIGO = " & Val(Grilla.TextMatrix(x, 0))
                DataEnvironment1.Sistema.Execute cad_update
                'DataEnvironment1.dbo_ANULOCHEQUES Val(grilla.TextMatrix(x, 0)), dtFecha
                DataEnvironment1.dbo_GRABARBITACORA Val(Grilla.TextMatrix(x, 0)), "chq_comp", Val(UsuarioSistema!codigo), Date, Time, "B"
            End If
        Next
        InicioGrilla
        cmdaceptar.enabled = False
        cmdcancelar.enabled = False
        lblcartel.Visible = False
    End If
    
End Sub

Sub InicioGrilla()
    Grilla.clear
    Grilla.ColWidth(3) = 5000
    Grilla.ColWidth(6) = 0
    Grilla.TextMatrix(0, 0) = "Nº Interno"
    Grilla.TextMatrix(0, 1) = "Fecha Cheque"
    Grilla.TextMatrix(0, 2) = "Importe"
    Grilla.TextMatrix(0, 3) = "Descripción de la Cuenta"
    Grilla.TextMatrix(0, 4) = "Nº Cheque"
    Grilla.TextMatrix(0, 5) = "   "
    Grilla.TextMatrix(0, 6) = ""
    Grilla.rows = 2
End Sub

Private Sub Cargogrilla()
Dim rs As New ADODB.Recordset

    rs.Open "select * from Chq_comp where estado = 'B' or estado = 'C' and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rs.EOF Then
        cmdaceptar.enabled = True
        cmdcancelar.enabled = True
        While Not rs.EOF
            If Grilla.rows = 2 Then
                Grilla.Row = 1
                Grilla.Col = 0
                If Trim(Grilla.Text) = "" Then
                    Grilla.Row = 1
                    Grilla.Col = 0
                    Grilla.Text = rs!codigo
                    Grilla.Col = 1
                    Grilla.Text = rs!fecha_cheque
                    Grilla.Col = 2
                    Grilla.Text = rs!Importe
                    Grilla.Col = 3
                    Grilla.Text = ObtenerDescCtaBancaria(rs!cuentabancaria)
                    Grilla.Col = 4
                    Grilla.Text = rs!Nro
                    Grilla.Col = 6
                    Grilla.Text = rs!cuentabancaria
                Else
                    Grilla.AddItem rs!codigo & Chr(9) & rs!fecha_cheque & Chr(9) & rs!Importe & Chr(9) & ObtenerDescCtaBancaria(rs!cuentabancaria) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!cuentabancaria
                End If
            Else
                Grilla.AddItem rs!codigo & Chr(9) & rs!fecha_cheque & Chr(9) & rs!Importe & Chr(9) & ObtenerDescCtaBancaria(rs!cuentabancaria) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!cuentabancaria
            End If
            rs.MoveNext
        Wend
    Else
        MsgBox "No hay cheques para anular"
    End If
End Sub

Function ObtenerDescCtaBancaria(dato As Long) As String

Dim rs1 As New ADODB.Recordset

    rs1.Open "select * from Ctasbank where codigo = " & dato & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs1.EOF Then
        ObtenerDescCtaBancaria = ObtenerDescripcionBancos("BancosGrales", rs1!Banco) & " - " & rs1!numero & " - " & rs1!tipo & " - " & ObtenerDescripcion("TipoCtas", Val(rs1!tipo))
    End If
    rs1.Close
    Set rs1 = Nothing

End Function

Private Sub cmdcancelar_Click()
    InicioGrilla
    cmdaceptar.enabled = False
    cmdcancelar.enabled = False
    lblcartel.Visible = False
    dtFecha = Date
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub dtfecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    lblcartel.Visible = False
    dtFecha = Date
End Sub

Private Sub grilla_Click()
    If Grilla.TextMatrix(Grilla.Row, 5) <> "        X" Then
        Grilla.TextMatrix(Grilla.Row, 5) = "        X"
    Else
        Grilla.TextMatrix(Grilla.Row, 5) = ""
    End If
End Sub


