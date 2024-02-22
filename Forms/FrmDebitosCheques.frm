VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDebitosCheques 
   Caption         =   "Débito de Cheques"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "FrmDebitosCheques.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      Top             =   550
      Width           =   1455
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "4"
      Top             =   5310
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      Top             =   5310
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "5"
      Top             =   5310
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Tag             =   "0"
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   450
      _Version        =   393216
      Format          =   62521345
      CurrentDate     =   38052
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmDebitosCheques.frx":000C
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   1020
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7011
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
   Begin Gestion.uCtaBanco uCtaBanco 
      Height          =   375
      Left            =   285
      TabIndex        =   8
      Top             =   570
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   661
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   750
      Left            =   9105
      TabIndex        =   9
      Top             =   225
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   6900
      TabIndex        =   10
      Tag             =   "0"
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   450
      _Version        =   393216
      Format          =   62521345
      CurrentDate     =   38052
   End
   Begin VB.Label Label1 
      Caption         =   "Importe total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha hasta"
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
      Left            =   5640
      TabIndex        =   11
      Top             =   240
      Width           =   1170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   4995
      Left            =   120
      Top             =   105
      Width           =   9855
   End
   Begin VB.Label lblcartel 
      Caption         =   """Click"" para marcar el cheque a debitar"
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
      Top             =   5190
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
Attribute VB_Name = "FrmDebitosCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Private Sub cmbver_Click()
    InicioGrilla
    Cargogrilla
    lblcartel.Visible = True
End Sub

Private Sub cmdAceptar_Click()
Dim x As Long, maximo As Long
Dim rs As New ADODB.Recordset
Dim mensaje As String, iddoc_puto As Long
    
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
        
        For x = 1 To grilla.rows - 1
            If grilla.TextMatrix(x, 5) = "        X" Then
                maximo = maximo + 1
                'DataEnvironment1.dbo_DEBITOSCHEQUES Val(grilla.TextMatrix(x, 0)), Val(grilla.TextMatrix(x, 6)), dtFecha, s2n(grilla.TextMatrix(x, 2)), maximo
                iddoc_puto = nSinNull(obtenerDeSQL("select iddoc from chq_comp where codigo = " & Val(grilla.TextMatrix(x, 0))))
                DataEnvironment1.Sistema.Execute "UPDATE CHQ_COMP  SET ESTADO='B', FECHA_OPERACION=" & ssFecha(dtfecha) & " WHERE CODIGO=" & Val(grilla.TextMatrix(x, 0))
                DataEnvironment1.Sistema.Execute "INSERT INTO MOVIBANC (CUENTA, OPERACION, DESCRIPCION, FECHA, DOCUMENTO, INTERNO, IMPORTE,TIPDOC, MOVBANCO, ACTIVO, FECHA_ALTA, USUARIO_ALTA, IDDOC)  VALUES(" & Val(grilla.TextMatrix(x, 6)) & ", 'B', 'Debito de Cheque', " & ssFecha(dtfecha) & ", 'P'," & Val(grilla.TextMatrix(x, 0)) & ", '" & x2s(grilla.TextMatrix(x, 2)) & "','DEB'," & maximo & ", 1 , " & ssFecha(Date) & "," & UsuarioActual & ", " & iddoc_puto & " )"
                
            End If
        Next
        InicioGrilla
        cmdaceptar.enabled = False
        cmdcancelar.enabled = False
        lblcartel.Visible = False
    End If
    
End Sub

Sub InicioGrilla()
    grilla.clear
    grilla.ColWidth(3) = 5000
    grilla.ColWidth(6) = 0
    grilla.TextMatrix(0, 0) = "Nº Interno"
    grilla.TextMatrix(0, 1) = "Fecha Vencimiento"
    grilla.TextMatrix(0, 2) = "Importe"
    grilla.TextMatrix(0, 3) = "Proveedor"
    grilla.TextMatrix(0, 4) = "Nº Cheque"
    grilla.TextMatrix(0, 5) = "   "
    grilla.TextMatrix(0, 6) = ""
    grilla.rows = 2
End Sub

Private Sub Cargogrilla()
Dim i As Long
Dim Aux As Double
Dim rs As New ADODB.Recordset

    rs.Open "select * from Chq_comp where estado = 'T' and activo = 1 and cuentabancaria = " & uCtaBanco.codigo & " and fecha_cheque<=" & ssFecha(DTPicker1.Value) & " order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rs.EOF Then
        cmdaceptar.enabled = True
        cmdcancelar.enabled = True
        While Not rs.EOF
            If grilla.rows = 2 Then
                grilla.Row = 1
                grilla.Col = 0
                If Trim(grilla.Text) = "" Then
                    grilla.Row = 1
                    grilla.Col = 0
                    grilla.Text = rs!codigo
                    grilla.Col = 1
                    grilla.Text = rs!fechadeposito
                    grilla.Col = 2
                    grilla.Text = rs!Importe
                    grilla.Col = 3
                    grilla.Text = ObtenerDescripcion("Prov", rs!Proveedor)
                    grilla.Col = 4
                    grilla.Text = rs!Nro
                    grilla.Col = 6
                    grilla.Text = rs!cuentabancaria
                Else
                    grilla.AddItem rs!codigo & Chr(9) & rs!fechadeposito & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Prov", rs!Proveedor) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!cuentabancaria
                End If
            Else
                grilla.AddItem rs!codigo & Chr(9) & rs!fechadeposito & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Prov", rs!Proveedor) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!cuentabancaria
            End If
            rs.MoveNext
        Wend
        i = 1
        While i < grilla.rows
            Aux = Aux + grilla.TextMatrix(i, 2)
            i = i + 1
        Wend
        Text1.Text = Aux
        
    Else
        MsgBox "No hay cheques para debitar"
    End If
End Sub

Private Sub cmdCancelar_Click()
    InicioGrilla
    cmdaceptar.enabled = False
    cmdcancelar.enabled = False
    lblcartel.Visible = False
    dtfecha = Date
    DTPicker1 = Date
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True
End Sub

Private Sub Form_Load()
    lblcartel.Visible = False
    dtfecha = Date
    DTPicker1 = Date
    ucXls1.ini grilla, "C:\ChequesPorDebitar.xls", "Cheques de Terceros"
        
End Sub

Private Sub grilla_Click()
    If grilla.TextMatrix(grilla.Row, 5) <> "        X" Then
        grilla.TextMatrix(grilla.Row, 5) = "        X"
    Else
        grilla.TextMatrix(grilla.Row, 5) = ""
    End If
End Sub
