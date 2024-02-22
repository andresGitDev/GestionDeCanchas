VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmExtractoBanc 
   Caption         =   "Extracto Bancario"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   Icon            =   "FrmExtractoBanc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ver Cheques"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtcodcta1 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmbcuenta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuenta"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtcuenta1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Tag             =   "2"
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtcodcta2 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmbcuenta2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuenta"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtcuenta2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox cargo 
      Height          =   285
      Left            =   7560
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker fechadesde 
      Height          =   255
      Left            =   1230
      TabIndex        =   10
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   255
      Left            =   4230
      TabIndex        =   11
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   38052
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmExtractoBanc.frx":08CA
      Height          =   3855
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).BandIndent=   10
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   5535
      Left            =   120
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "FrmExtractoBanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Private Sub cargo_GotFocus()
    cargo.SelStart = 0
    cargo.SelLength = Len(cargo.Text)
End Sub

Private Sub cmbcuenta_Click()
    cargo = "Ctasbank1"
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
End Sub

Private Sub cmbcuenta2_Click()
    cargo = "Ctasbank2"
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
End Sub

Private Sub cmdAceptar_Click()
           
Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
Dim Fecha As String, Banco As String, nrocheque As String, Cuenta As Long
   
    InicioGrilla
   
    rs.Open "select movibanc.*, ctasbank.numero, tipoctas.descripcion, bancosgrales.descripcion as desbanco from (((movibanc inner join ctasbank on movibanc.cuenta = ctasbank.codigo) inner join bancosgrales on ctasbank.banco = bancosgrales.codigo) inner join tipoctas on ctasbank.tipo = tipoctas.codigo) where movibanc.fecha >= " & ssFecha(fechadesde) & " and movibanc.fecha <= " & ssFecha(fechahasta) & " and movibanc.cuenta >= " & val(txtcodcta1) & " and movibanc.cuenta <= " & val(txtcodcta2) & " order by movibanc.cuenta, movibanc.fecha", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
            While Not rs.EOF
                
                If Cuenta <> rs!Cuenta Then
                    If GRILLA.TextMatrix(GRILLA.Row, GRILLA.Col) <> "" Then
                        GRILLA.AddItem rs!Cuenta & " " & rs!desbanco & " " & rs!DESCRIPCION & Chr(9) & rs!numero & Chr(9) & "Saldo: "
                    Else
                        GRILLA.TextMatrix(GRILLA.Row, 0) = rs!Cuenta & " " & rs!desbanco & " " & rs!DESCRIPCION
                        GRILLA.TextMatrix(GRILLA.Row, 1) = rs!numero
                        GRILLA.TextMatrix(GRILLA.Row, 2) = "Saldo: "
                    End If
                End If
                
                'fecha = Month(rs!fecha) & "/" & Day(rs!fecha) & "/" & Year(rs!fecha)
                
                nrocheque = ""
                Banco = ""
                If rs!documento = "P" Then
                    rs1.Open "select chq_comp.nro, bancosgrales.descripcion from chq_comp inner join bancosgrales on chq_comp.banco = bancosgrales.codigo where chq_comp.codigo = " & rs!interno & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not rs1.EOF Then
                        If rs1!Nro <> Null Or rs1!Nro <> "" Then
                            nrocheque = rs1!Nro
                        End If
                        If rs1!DESCRIPCION <> Null Or rs1!DESCRIPCION <> "" Then
                            Banco = rs1!DESCRIPCION
                        End If
                    End If
                    rs1.Close
                    Set rs1 = Nothing
                Else
                    If rs!documento = "C" Then
                        rs1.Open "select cheques.nro, bancosgrales.descripcion from cheques inner join bancosgrales on cheques.banco_nro = bancosgrales.codigo where cheques.nroint = " & rs!interno & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If Not rs1.EOF Then
                            If rs1!Nro <> Null Or rs1!Nro <> "" Then
                                nrocheque = rs1!Nro
                            End If
                            If rs1!DESCRIPCION <> Null Or rs1!DESCRIPCION <> "" Then
                                Banco = rs1!DESCRIPCION
                            End If
                        End If
                        rs1.Close
                        Set rs1 = Nothing
                    End If
                End If
                        
                       
                If rs!OPERACION = "A" Or rs!OPERACION = "D" Or rs!OPERACION = "E" Or rs!OPERACION = "T" Then
                           
        '            If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
        '                grilla.AddItem Chr(9) & rs!fecha & Chr(9) & IIf(rs!descripcion <> "" Or rs!descripcion <> Null, rs!descripcion, "") & Chr(9) & 0 & Chr(9) & rs!importe & Chr(9) & Chr(9) & rs!documento & Chr(9) & _
        '                IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0) & Chr(9) & Banco & Chr(9) & nrocheque & Chr(9) & rs!movbanco
        '            Else
        '                grilla.TextMatrix(grilla.Row, 1) = rs!fecha
        '                grilla.TextMatrix(grilla.Row, 2) = IIf(rs!descripcion <> "" Or rs!descripcion <> Null, rs!descripcion, "")
        '                grilla.TextMatrix(grilla.Row, 3) = 0
        '                grilla.TextMatrix(grilla.Row, 4) = rs!importe
        '                grilla.TextMatrix(grilla.Row, 6) = rs!documento
        '                grilla.TextMatrix(grilla.Row, 7) = IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0)
        '                grilla.TextMatrix(grilla.Row, 8) = Banco
        '                grilla.TextMatrix(grilla.Row, 9) = nrocheque
        '                grilla.TextMatrix(grilla.Row, 10) = rs!movbanco
        '            End If
        '
        '        Else
                    
                    If GRILLA.TextMatrix(GRILLA.Row, GRILLA.Col) <> "" Then
                        GRILLA.AddItem Chr(9) & Fecha & Chr(9) & IIf(rs!DESCRIPCION <> "" Or rs!DESCRIPCION <> Null, rs!DESCRIPCION, "") & Chr(9) & rs!Importe & Chr(9) & Chr(9) & 0 & Chr(9) & rs!documento & Chr(9) & _
                        IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0) & Chr(9) & Banco & Chr(9) & nrocheque & Chr(9) & rs!MovBanco
                    Else
                        GRILLA.TextMatrix(GRILLA.Row, 1) = Fecha
                        GRILLA.TextMatrix(GRILLA.Row, 2) = IIf(rs!DESCRIPCION <> "" Or rs!DESCRIPCION <> Null, rs!DESCRIPCION, "")
                        GRILLA.TextMatrix(GRILLA.Row, 3) = rs!Importe
                        GRILLA.TextMatrix(GRILLA.Row, 4) = 0
                        GRILLA.TextMatrix(GRILLA.Row, 6) = rs!documento
                        GRILLA.TextMatrix(GRILLA.Row, 7) = IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0)
                        GRILLA.TextMatrix(GRILLA.Row, 8) = Banco
                        GRILLA.TextMatrix(GRILLA.Row, 9) = nrocheque
                        GRILLA.TextMatrix(GRILLA.Row, 10) = rs!MovBanco
                    End If
                                
                                
                'este parrafo es el que di vuelta
                Else
                    If GRILLA.TextMatrix(GRILLA.Row, GRILLA.Col) <> "" Then
                        GRILLA.AddItem Chr(9) & rs!Fecha & Chr(9) & IIf(rs!DESCRIPCION <> "" Or rs!DESCRIPCION <> Null, rs!DESCRIPCION, "") & Chr(9) & 0 & Chr(9) & rs!Importe & Chr(9) & Chr(9) & rs!documento & Chr(9) & _
                        IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0) & Chr(9) & Banco & Chr(9) & nrocheque & Chr(9) & rs!MovBanco
                    Else
                        GRILLA.TextMatrix(GRILLA.Row, 1) = rs!Fecha
                        GRILLA.TextMatrix(GRILLA.Row, 2) = IIf(rs!DESCRIPCION <> "" Or rs!DESCRIPCION <> Null, rs!DESCRIPCION, "")
                        GRILLA.TextMatrix(GRILLA.Row, 3) = 0
                        GRILLA.TextMatrix(GRILLA.Row, 4) = rs!Importe
                        GRILLA.TextMatrix(GRILLA.Row, 6) = rs!documento
                        GRILLA.TextMatrix(GRILLA.Row, 7) = IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0)
                        GRILLA.TextMatrix(GRILLA.Row, 8) = Banco
                        GRILLA.TextMatrix(GRILLA.Row, 9) = nrocheque
                        GRILLA.TextMatrix(GRILLA.Row, 10) = rs!MovBanco
                    End If
                'fin del parrafo
                                
                                
                                
                End If
                
                Cuenta = rs!Cuenta
                
                rs.MoveNext
                
            Wend
            rs.Close
    Else
        MsgBox "No se encontro ningun registro.", vbInformation
    End If

    Set rs = Nothing
    
'    daTaenvironment1.LisPorMovimiento
'    rptPorMovimiento.Show vbModal
'    daTaenvironment1.rsLisPorMovimiento.Close
'
'    daTaenvironment1.dbo_LISTADOXMOVIMIENTO "B", 0, "", 0, 0, "", 0, "", "", 0
    
'    LimpioControles
    cargo = ""
    
End Sub


Private Sub cmdCancelar_Click()
    LimpioControles
    InicioGrilla
    cargo = ""
    fechadesde = Date
    fechahasta = Date
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub fechadesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Private Sub fechahasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Public Sub CargarDatos()
Dim rs As New ADODB.Recordset
   
'    codigo = Val(Trim(Me.Tag))
       
    If cargo = "Ctasbank1" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta1) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcta1 = rs!codigo
            txtcuenta1 = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargo = "Ctasbank2" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta2) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcta2 = rs!codigo
            txtcuenta2 = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
End Sub

Private Sub Form_Load()
    InicioGrilla
    cargo = ""
    fechadesde = Date
    fechahasta = Date
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub txtcodcta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtcodcta2_GotFocus()
    txtcodcta2.SelStart = 0
    txtcodcta2.SelLength = Len(txtcodcta2.Text)
End Sub

Private Sub txtcodcta1_GotFocus()
    txtcodcta1.SelStart = 0
    txtcodcta1.SelLength = Len(txtcodcta1.Text)
End Sub

Private Sub txtcodcta1_LostFocus()
    If IsNumeric(txtcodcta1) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta1) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta1 = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        Else
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta1 = "0"
            txtcodcta1.SetFocus
        End If
        rs.Close
        Set rs = Nothing
    Else
        If txtcodcta1 <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta1 = "0"
            txtcodcta1.SetFocus
        End If
    End If
End Sub

Private Sub txtcodcta2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtcodcta2_LostFocus()
    If IsNumeric(txtcodcta2) Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcta2) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta2 = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        Else
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta2 = "0"
            txtcodcta2.SetFocus
        End If
        rs.Close
        Set rs = Nothing
    Else
        If txtcodcta2 <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcta2 = "0"
            txtcodcta2.SetFocus
        End If
    End If
End Sub


Private Sub LimpioControles()
    txtcodcta1 = ""
    txtcodcta2 = ""
    txtcuenta1 = ""
    txtcuenta2 = ""
    fechadesde = Date
    fechahasta = Date
End Sub

'Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
'Dim fecha As String, banco As String, nrocheque As String, cuenta As long
'
'    daTaenvironment1.dbo_LISTADOXMOVIMIENTO "B", 0, "", 0, 0, "", 0, "", "", 0
'
'    rs.Open "select movibanc.*, ctasbank.numero, tipoctas.descripcion, bancosgrales.descripcion as desbanco from (((movibanc inner join ctasbank on movibanc.cuenta = ctasbank.codigo) inner join bancosgrales on ctasbank.banco = bancosgrales.codigo) inner join tipoctas on ctasbank.tipo = tipoctas.codigo) where movibanc.fecha >= '" & fechadesde & "' and movibanc.fecha <= '" & fechahasta & "' and movibanc.cuenta >= " & Val(txtcodcta1) & " and movibanc.cuenta <= " & Val(txtcodcta2) & " order by movibanc.fecha, movibanc.cuenta", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    While Not rs.EOF
'
'        If cuenta <> rs!cuenta Then
'            'mando saldo anterior
'        End If
'
'        fecha = Month(rs!fecha) & "/" & Day(rs!fecha) & "/" & Year(rs!fecha)
'
'        nrocheque = ""
'        banco = ""
'        If rs!documento = "P" Then
'            rs1.Open "select chq_comp.nro, bancosgrales.descripcion from chq_comp inner join bancosgrales on chq_comp.banco = bancosgrales.codigo where chq_comp.codigo = " & rs!interno & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            If Not rs1.EOF Then
'                If rs1!Nro <> Null Or rs1!Nro <> "" Then
'                    nrocheque = rs1!Nro
'                End If
'                If rs1!descripcion <> Null Or rs1!descripcion <> "" Then
'                    banco = rs1!descripcion
'                End If
'            End If
'            rs1.Close
'            Set rs1 = Nothing
'        Else
'            If rs!documento = "C" Then
'                rs1.Open "select cheques.nro, bancosgrales.descripcion from cheques inner join bancosgrales on cheques.banco_nro = bancosgrales.codigo where cheques.nroint = " & rs!interno & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                If Not rs1.EOF Then
'                    If rs1!Nro <> Null Or rs1!Nro <> "" Then
'                        nrocheque = rs1!Nro
'                    End If
'                    If rs1!descripcion <> Null Or rs1!descripcion <> "" Then
'                        banco = rs1!descripcion
'                    End If
'                End If
'                rs1.Close
'                Set rs1 = Nothing
'            End If
'        End If
'
'        If rs!descripcion <> Null Then
'
'        End If
'
'        If rs!operacion = "A" Or rs!operacion = "D" Or rs!operacion = "E" Or rs!operacion = "T" Then
'            daTaenvironment1.dbo_LISTADOXMOVIMIENTO "A", fecha, IIf(rs!descripcion <> "" Or rs!descripcion <> Null, rs!descripcion, ""), 0, rs!importe, rs!documento, IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0), banco, nrocheque, rs!movbanco
'        Else
'            daTaenvironment1.dbo_LISTADOXMOVIMIENTO "A", fecha, IIf(rs!descripcion <> "" Or rs!descripcion <> Null, rs!descripcion, ""), rs!importe, 0, rs!documento, IIf(rs!interno <> Null Or rs!interno <> 0, rs!interno, 0), banco, nrocheque, rs!movbanco
'        End If
'
'        cuenta = rs!cuenta
'
'        rs.MoveNext
'
'    Wend
'    rs.Close
'    Set rs = Nothing
'
'    daTaenvironment1.LisPorMovimiento
'    rptPorMovimiento.Show vbModal
'    daTaenvironment1.rsLisPorMovimiento.Close
'
'    daTaenvironment1.dbo_LISTADOXMOVIMIENTO "B", 0, "", 0, 0, "", 0, "", "", 0
'
'    LimpioControles
'    cargo = ""
    

Private Sub InicioGrilla()
    GRILLA.clear
    GRILLA.ColWidth(0) = 4000
    GRILLA.ColWidth(1) = 1000
    GRILLA.ColWidth(2) = 2000
    GRILLA.ColWidth(8) = 2500
    GRILLA.TextMatrix(0, 1) = "Fecha"
    GRILLA.TextMatrix(0, 2) = "Descripción"
    GRILLA.TextMatrix(0, 3) = "Débitos"
    GRILLA.TextMatrix(0, 4) = "Créditos"
    GRILLA.TextMatrix(0, 5) = "Saldo"
    GRILLA.TextMatrix(0, 6) = "Doc."
    GRILLA.TextMatrix(0, 7) = "Nº Int."
    GRILLA.TextMatrix(0, 8) = "Banco"
    GRILLA.TextMatrix(0, 9) = "Nº Cheque"
    GRILLA.TextMatrix(0, 10) = "Nº Mov."
    GRILLA.rows = 2
End Sub


