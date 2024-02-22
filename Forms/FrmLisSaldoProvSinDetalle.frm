VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisSaldoProvSinDetalle 
   Caption         =   "Listado de Saldos de Proveedores sin Detalle"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   Icon            =   "FrmLisSaldoProvSinDetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   855
      Left            =   7680
      TabIndex        =   16
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin VB.TextBox txtcodclid 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   1215
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
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtcliented 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox txtcodclih 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
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
      TabIndex        =   7
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtclienteh 
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
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
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
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
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
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
      TabIndex        =   2
      Top             =   2160
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
      TabIndex        =   1
      Top             =   2160
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   240
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Format          =   72089601
      CurrentDate     =   38052
   End
   Begin VSFlex7LCtl.VSFlexGrid gSaldo 
      Height          =   3015
      Left            =   0
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   7575
      _cx             =   13361
      _cy             =   5318
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
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label6 
      Caption         =   "Desde Prov :"
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
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta Prov :"
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
End
Attribute VB_Name = "FrmLisSaldoProvSinDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 10/3/5


Private smSaldoProTmp As String
'

Private Sub cmdaceptar_Click()
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    
    Dim Total As Variant
    Dim STR As String

    'If smSaldoProTmp = "" Then
    smSaldoProTmp = TablaTempCrear(tt_SaldoCliTemp)
    
    'daTaenvironment1.Sistema.Execute "delete from  " & smSaldoProTmp & " "
    
    rs1.Open "select sum(saldo) as afavor,codpr from transcom where fecha<= " & ssFecha(dtFecha) & "  and activo=1 and (tipodoc='RAC' or tipodoc='APC' or tipodoc='N/C') and (codpr>=" & Val(txtcodclid) & " and codpr<=" & Val(txtcodclih) & ") and saldo > 0 group by codpr", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        Do While Not rs1.EOF
            DataEnvironment1.Sistema.Execute "insert into " & smSaldoProTmp & " (codigo,descripcion,saldo)values(" & rs1!CODPR & ",'" & ObtenerDescripcion("prov", rs1!CODPR) & "'," & Replace(s2n(rs1!afavor), ",", ".") & ")"
            rs1.MoveNext
        Loop
    End If
    rs1.Close
    Set rs1 = Nothing
    
'*******************************************************************************

    rs2.Open "select sum(saldo) as deuda,codpr from transcom where fecha<= " & ssFecha(dtFecha) & "  and activo=1 and (tipodoc='FAC' or tipodoc='APD' or tipodoc='N/D') and (codpr>=" & Val(txtcodclid) & " and codpr<=" & Val(txtcodclih) & ") and saldo > 0 group by codpr", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs2.EOF Then
        Do While Not rs2.EOF
            DataEnvironment1.Sistema.Execute "insert into " & smSaldoProTmp & "  (codigo,descripcion,saldo)values(" & rs2!CODPR & ",'" & ObtenerDescripcion("prov", rs2!CODPR) & "',-" & Replace(s2n(rs2!deuda), ",", ".") & ")"
            rs2.MoveNext
        Loop
    End If
    rs2.Close
    Set rs2 = Nothing

'********************************************************************************
    Dim str2 As String
    If opttodos.Value = True Then
        STR = "select sum(saldo) as sal,codigo,descripcion from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo)<0 or sum(saldo)>5 "
        str2 = "select codigo as Prov,descripcion as razon,sum(saldo) as Saldo from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo)<0 or sum(saldo)>5 "
    Else
        If optpos.Value = True Then
            STR = "select sum(saldo) as sal,codigo,descripcion from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo) > 5"
            str2 = "select codigo as Prov,descripcion as razon,sum(saldo) as Saldo from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo) > 0"
        Else
            If optneg.Value = True Then
                STR = "select sum(saldo) as sal,codigo,descripcion from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo) < 0"
                str2 = "select codigo as Prov,descripcion as razon,sum(saldo) as Saldo from " & smSaldoProTmp & "  group by codigo,descripcion having sum(saldo) < 0"
            End If
        End If
    End If
    
'*********************************************************************************

    
    RptLisSaldoProvSinDetalle.Data.Connection = DataEnvironment1.Sistema
    RptLisSaldoProvSinDetalle.Data.Source = STR
    RptLisSaldoProvSinDetalle.lblfecha = dtFecha.Value 'Date
    RptLisSaldoProvSinDetalle.Show
    
    LlenarGrilla gSaldo, str2, True

End Sub

Private Sub cmdayudacli_Click()
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("prov")
    If resu > "" Then
        txtcodclid = frmBuscar.resultado
        txtcliented = frmBuscar.resultado(2)
    End If
End Sub
Private Sub cmdcancelar_Click()
    LimpioCampos
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("prov")
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
    ucXls1.ini gSaldo, "C:\SaldoProveedorSinDetalle.xls", "Saldo Proveedor sin Detalle"
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

'Private Sub Form_Unload(cancel As Integer)
'    TablaTempBorrar smSaldoProTmp
'End Sub

Private Sub txtCodCliD_GotFocus()
    txtcodclid.SelStart = 0
    txtcodclid.SelLength = Len(txtcodclid.Text)
End Sub

Private Sub txtCodCliH_GotFocus()
    txtcodclih.SelStart = 0
    txtcodclih.SelLength = Len(txtcodclih.Text)
End Sub


'10/3/5 tablatemp
'

