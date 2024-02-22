VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLisSaldoCliSinDetalle 
   Caption         =   "Listado de Saldos por Cliente"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "FrmLisSaldoCliSinDetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid gSaldos 
      Height          =   4245
      Left            =   120
      TabIndex        =   21
      Top             =   3525
      Visible         =   0   'False
      Width           =   7485
      _cx             =   13203
      _cy             =   7488
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   990
      Left            =   7665
      TabIndex        =   20
      Top             =   75
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1746
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   195
      TabIndex        =   16
      Top             =   2040
      Width           =   7230
      Begin VB.OptionButton OptALL 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5700
         TabIndex        =   19
         Top             =   60
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton OptMin 
         Alignment       =   1  'Right Justify
         Caption         =   "Clientes Minoristas"
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
         Left            =   2685
         TabIndex        =   18
         Top             =   60
         Width           =   2310
      End
      Begin VB.OptionButton OptMay 
         Alignment       =   1  'Right Justify
         Caption         =   "Clientes Mayoristas"
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
         Left            =   -45
         TabIndex        =   17
         Top             =   75
         Width           =   2145
      End
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2970
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2970
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2970
      Width           =   975
   End
   Begin VB.OptionButton opttodos 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5880
      TabIndex        =   5
      Top             =   1680
      Width           =   1335
   End
   Begin VB.OptionButton optneg 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo Negativo"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
   End
   Begin VB.OptionButton optpos 
      Alignment       =   1  'Right Justify
      Caption         =   "Saldo Positivo"
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
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtclienteh 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtcodclih 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtcliented 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton cmdayudacli 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtcodclid 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Format          =   203620353
      CurrentDate     =   38052
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta Cliente :"
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
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Desde Cliente :"
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
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2505
      Left            =   120
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label5 
      Caption         =   "Saldo a "
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
      Left            =   5040
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmLisSaldoCliSinDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private smSaldoCliTmp As String


Private Sub cmdAceptar_Click()
    Dim rsClie As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim TotalHaber As Variant
    Dim TotalDebe As Variant
    Dim Total As Variant
    Dim str As String
    Dim clie As String, clieDes As String

'    If smSaldoCliTmp = "" Then
        smSaldoCliTmp = TablaTempCrear(tt_SaldoCliTemp)
        '
'    End If

    If OptMay.Value = True Then
        rsClie.Open "select * from clientes where codigo>=" & txtcodclid & " and codigo<=" & txtcodclih & " and mayorista = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    ElseIf OptMin.Value = True Then
        rsClie.Open "select * from clientes where codigo>=" & txtcodclid & " and codigo<=" & txtcodclih & " and mayorista <> 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    ElseIf OptALL.Value = True Then
        rsClie.Open "select * from clientes where codigo>=" & txtcodclid & " and codigo<=" & txtcodclih, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    End If
    
    

    If Not rsClie.EOF Then
        Do While Not rsClie.EOF
            clie = rsClie!codigo
            clieDes = ssStr(rsClie!DESCRIPCION)
            TotalDebe = 0: TotalHaber = 0
            
            
            rs1.Open "select sum(saldo) as deuda,cliente from FacturaVenta where fecha<=" & ssFecha(dtFecha) & " and activo=1 and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='FAE'  or tipodoc='FEA' or tipodoc='FEB' or tipodoc='FEC'  or tipodoc='DEA' or tipodoc='DEB' or tipodoc='DEC'  or tipodoc='NDA' or tipodoc='NDB' or tipodoc='NDR' or tipodoc='NDE' or tipodoc='ACD') and (cliente = " & clie & "   ) group by cliente", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If (rs1.EOF = True And rs1.BOF = True) Or IsNull(rs1!deuda) Or IsEmpty(rs1!deuda) Then
                TotalDebe = 0
            Else
                TotalDebe = s2n(rs1!deuda)
            End If
            rs1.Close
            
            rs2.Open "select sum(saldo) as afavor,cliente from FacturaVenta where fecha<=" & ssFecha(dtFecha) & " and  activo=1 and (tipodoc='RAA' or tipodoc='NCA' or tipodoc='NCB' or tipodoc='NCE' or tipodoc='CEA' or tipodoc='CEB' or tipodoc='CEC' or tipodoc='ACC')  and cliente=" & clie & " group by cliente", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF Then
                TotalHaber = s2n(rs2!afavor)
            End If
            rs2.Close
            Set rs2 = Nothing
            
            
            Total = s2n(TotalDebe) - s2n(TotalHaber)
            If opttodos.Value = True Then
                If Total <> 0 Then
                    DataEnvironment1.Sistema.Execute "insert into  " & smSaldoCliTmp & " (codigo,descripcion,saldo)values(" & clie & ",'" & clieDes & "'," & Replace(s2n(Total), ",", ".") & ")"
                End If
            Else
                If optpos.Value = True Then
                    If Total > 0 Then
                        DataEnvironment1.Sistema.Execute "insert into  " & smSaldoCliTmp & " (codigo,descripcion,saldo)values(" & clie & ",'" & clieDes & "'," & n2s(Total) & ")"
                    End If
                Else
                    If optneg.Value = True Then
                        If Total < 0 Then
                            DataEnvironment1.Sistema.Execute "insert into  " & smSaldoCliTmp & " (codigo,descripcion,saldo)values(" & clie & ",'" & clieDes & "'," & s2n(Total) & ")"
                        End If
                    End If
                End If
            End If
            rsClie.MoveNext

        Loop
    End If
    rsClie.Close
    Set rs1 = Nothing
    str = "select * from " & smSaldoCliTmp & " order by codigo"
    RptLisSaldoCliSinDetalle.Data.Connection = DataEnvironment1.Sistema
    RptLisSaldoCliSinDetalle.Data.Source = str
    RptLisSaldoCliSinDetalle.lblfecha = Date
    RptLisSaldoCliSinDetalle.Show
    
    LlenarGrilla gSaldos, str, True
    
    If gSaldos.rows > 1 Then
        ucXls1.enabled = True
    Else
        ucXls1.enabled = False
    End If
End Sub

Private Sub cmdayudacli_Click()
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
    If resu > "" Then
        txtcodclid = frmBuscar.resultado
        txtcliented = frmBuscar.resultado(2)
    End If
End Sub
Private Sub cmdCancelar_Click()
    LimpioCampos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
    If resu > "" Then
        txtcodclih = frmBuscar.resultado
        txtclienteh = frmBuscar.resultado(2)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    LimpioCampos
    ucXls1.ini gSaldos, "C:\SaldosSinDetalleCleinets.xls", "Saldos de clientes"
End Sub
Sub LimpioCampos()
    txtcodclih = "9999"
    txtclienteh = ""
    txtcodclid = "0"
    txtcliented = ""
    optpos.Value = False
    optneg.Value = False
    opttodos.Value = True
    dtFecha.Value = Date
End Sub

Private Sub txtCodCliD_GotFocus()
    txtcodclid.SelStart = 0
    txtcodclid.SelLength = Len(txtcodclid.Text)
End Sub

Private Sub txtCodCliH_GotFocus()
    txtcodclih.SelStart = 0
    txtcodclih.SelLength = Len(txtcodclih.Text)
End Sub

