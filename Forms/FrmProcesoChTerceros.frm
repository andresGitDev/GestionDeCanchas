VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmProcesoChTerceros 
   Caption         =   "Administración de Cheques"
   ClientHeight    =   8040
   ClientLeft      =   180
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "FrmProcesoChTerceros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   8040
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid gImprime 
      Height          =   3660
      Left            =   8970
      TabIndex        =   33
      Top             =   2550
      Visible         =   0   'False
      Width           =   2865
      _cx             =   5054
      _cy             =   6456
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
   Begin VB.Frame fraBoton 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   240
      TabIndex        =   9
      Top             =   6510
      Width           =   8610
      Begin Gestion.ucCoDe uGastos 
         Height          =   315
         Left            =   1140
         TabIndex        =   31
         Top             =   930
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.TextBox txtGastos 
         Height          =   300
         Left            =   135
         TabIndex        =   28
         Text            =   "0"
         Top             =   930
         Width           =   960
      End
      Begin VB.CommandButton cmdIChequesT 
         Caption         =   "Ingresar cheques de T"
         Height          =   330
         Left            =   6660
         TabIndex        =   27
         Top             =   510
         Width           =   1845
      End
      Begin VB.CommandButton cmdAsientoDcanje 
         Caption         =   "&Generar Asiento de canje"
         Height          =   390
         Left            =   4950
         TabIndex        =   26
         Top             =   1515
         Visible         =   0   'False
         Width           =   2175
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   930
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   930
         Width           =   975
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
         Left            =   5985
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   930
         Width           =   975
      End
      Begin VB.TextBox txtcuenta 
         Height          =   285
         Left            =   4020
         TabIndex        =   16
         Tag             =   "2"
         Top             =   540
         Visible         =   0   'False
         Width           =   2295
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
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   495
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtcodcuenta 
         Height          =   285
         Left            =   1125
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cmbplazo 
         Height          =   315
         ItemData        =   "FrmProcesoChTerceros.frx":08CA
         Left            =   4980
         List            =   "FrmProcesoChTerceros.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtcaja 
         Height          =   285
         Left            =   3945
         TabIndex        =   12
         Tag             =   "2"
         Top             =   495
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmbcaja 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Caja"
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
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   495
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtcodcaja 
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Tag             =   "1"
         Top             =   495
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker fechaoperacion 
         Height          =   255
         Left            =   1620
         TabIndex        =   20
         Top             =   90
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   450
         _Version        =   393216
         Format          =   132775937
         CurrentDate     =   38052
      End
      Begin Gestion.ucCoDe uCajaEfectivo 
         Height          =   315
         Left            =   945
         TabIndex        =   32
         Top             =   510
         Width           =   5505
         _ExtentX        =   9234
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label lblplazo 
         Caption         =   "Plazo en horas:"
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
         Left            =   3465
         TabIndex        =   24
         Top             =   90
         Visible         =   0   'False
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
         Left            =   150
         TabIndex        =   23
         Top             =   525
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbloperacion 
         Caption         =   "F. de Operación"
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
         Left            =   60
         TabIndex        =   22
         Top             =   90
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblcaja 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Caja:"
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
         Left            =   150
         TabIndex        =   21
         Top             =   495
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.TextBox txtbanco 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8910
      TabIndex        =   8
      Tag             =   "6"
      Top             =   2085
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox cantidad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8910
      TabIndex        =   7
      Tag             =   "6"
      Top             =   1485
      Visible         =   0   'False
      Width           =   495
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
      Height          =   1290
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   8580
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         Left            =   6600
         TabIndex        =   34
         Text            =   "Ejercicio"
         Top             =   480
         Width           =   990
      End
      Begin VB.OptionButton optcobro 
         Caption         =   "Cobrar Chq en Canje"
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
         Left            =   5160
         TabIndex        =   29
         Top             =   975
         Width           =   3195
      End
      Begin VB.CheckBox chkFiltroFecha 
         Caption         =   "Mostrar cheques desde:"
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   150
         TabIndex        =   25
         Top             =   150
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton optVerTodo 
         Caption         =   "Ver todos"
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
         Left            =   7215
         TabIndex        =   4
         Top             =   225
         Width           =   1215
      End
      Begin VB.OptionButton optacreditar 
         Caption         =   "Acreditar"
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
         Left            =   1470
         TabIndex        =   1
         Top             =   945
         Width           =   1170
      End
      Begin VB.OptionButton optdepositar 
         Caption         =   "Depositar"
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
         Left            =   150
         TabIndex        =   0
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optcanje 
         Caption         =   "Canjear"
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
         Left            =   3990
         TabIndex        =   3
         Top             =   975
         Width           =   1215
      End
      Begin VB.OptionButton optrechazar 
         Caption         =   "Rechazar"
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
         Left            =   2730
         TabIndex        =   2
         Top             =   960
         Width           =   1320
      End
      Begin Gestion.ucFecha uFecha 
         Height          =   285
         Left            =   2385
         TabIndex        =   30
         Top             =   270
         Width           =   1245
         _ExtentX        =   2805
         _ExtentY        =   503
         FechaInit       =   0
      End
      Begin VB.Label Label34 
         Caption         =   "Ejercicio"
         Height          =   255
         Left            =   7680
         TabIndex        =   35
         Top             =   540
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillacheques 
      Bindings        =   "FrmProcesoChTerceros.frx":08EE
      Height          =   4950
      Left            =   105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1485
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8731
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "FrmProcesoChTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ver que datos pongo en concepto asiento

Option Explicit '16/9/4


Private midDoc As Long

Dim sologrilla As Long

Private Sub cmbcuenta_Click()
    sologrilla = 1
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
End Sub


'Private Sub cmbplazo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAaceptar
    Dim vdias
    Dim rs As New ADODB.Recordset
    Dim maximobanc1 As Long, maxcheque As Long, maximocaja As Long, x As Long, valcartera As String, depcuenta As Long, cuentaconcaja As Long
    'Dim fecha As String, fechaoper As String,
    Dim cuentacon As String, Importe As Double, asse As String
    Dim asiCh As New Asiento, aConcepto As String, cueBan As String
        
    If optVerTodo Then Exit Sub
    '
        
If val(cantidad) > 0 Then
    
''        fechach = Month(fechacheque) & "/" & Day(fechacheque) & "/" & Year(fechacheque)
''        fechaing = Month(fechacingreso) & "/" & Day(fechaingreso) & "/" & Year(fechaingreso)
        
'        fechaoper = Month(fechaoperacion) & "/" & Day(fechaoperacion) & "/" & Year(fechaoperacion)
 '        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
 
 
 
 
            If optdepositar = True Then
                If txtcodcuenta = "0" Or txtcodcuenta = "" Then
                    MsgBox "Falta indicar la cuenta de destino.", vbInformation
                    txtcodcuenta.SetFocus
                    Exit Sub
                Else
                    aConcepto = "Deposito de cheques " ' & banco & numeroch & ???
                End If
            End If
            If optacreditar = True Then
                aConcepto = "Acreditacion de cheques "
            End If
            If optrechazar = True Then
                aConcepto = "Rechazo de cheques "
            End If
            If optcanje = True Then
                aConcepto = "Canje de cheques "
                If uCajaEfectivo.codigo = "0" Or uCajaEfectivo.codigo = "" Then
                    MsgBox "Debe ingresar la caja.", , "ATENCION"
                    uCajaEfectivo.SetFocus
                    'Exit Sub
                Else
                    aConcepto = "Canje de cheques "
                End If
                If s2n(txtGastos) > 0 Then
                    If Trim(uGastos.codigo) = "" Then
                        MsgBox "Debe ingresar la cuenta de gastos.", , "ATENCION"
                        Exit Sub
                    End If
                End If
                
            End If
            
            If optcobro = True Then
                aConcepto = "Cobro de cheques en canje"
                If uCajaEfectivo.codigo = "0" Or uCajaEfectivo.codigo = "" Then
                    MsgBox "Debe ingresar la caja.", , "ATENCION"
                    uCajaEfectivo.SetFocus
                    Exit Sub
                End If
                If s2n(txtGastos) > 0 Then
                    If Trim(uGastos.codigo) = "" Then
                        MsgBox "Debe ingresar la cuenta de gastos.", , "ATENCION"
                        Exit Sub
                    End If
                End If
                
            End If
            asiCh.nuevo aConcepto, fechaoperacion, "CH3"
 
 
        '***************************************
        DE_BeginTrans
        
        midDoc = NuevoDocumento("ch3", nuevoCodigo("registrodocumentos", "nrodoc", "tipodoc = 'ch3'"), 0, 0)
        'para canje
        Dim aCaja As Double, aCartera As Double, aCanje As Double, AGastos As Double
        aCaja = 0
        aCartera = pTotalEnCheques
        aCanje = 0
        For x = 1 To grillacheques.rows - 1
            If grillacheques.TextMatrix(x, 5) <> "" Then
                If optdepositar Then
                    vdias = DateDiff("d", CDate(grillacheques.TextMatrix(x, 1)), fechaoperacion)
                    If vdias > 30 Then
                        MsgBox "El cheque " & grillacheques.TextMatrix(x, 4) & " excede de 30 dias al deposito. " & (vdias - 30) & " pasados del vencimiento", vbInformation
                        grillacheques.TextMatrix(x, 5) = ""
                        GoTo vencido
                    End If
                End If
''                rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''                If Not IsNull(rs!maxcodigo) Then
''                    maximobanc1 = rs!maxcodigo + 1
''                Else
''                    maximobanc1 = 1
''                End If
''                rs.Close
''                Set rs = Nothing
                maximobanc1 = nuevoCodigo("MoviBanc", "MovBanco")
                               
                
            
                If optdepositar = False Then
                    rs.Open "select dep_cuenta from Cheques where nroint = " & val(grillacheques.TextMatrix(x, 0)) & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If Not rs.EOF Then
                        depcuenta = rs!dep_cuenta
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
            
            
            
                'movibanc
                If optdepositar = True Then
                    asse = "optDepo chmovibanc"
                    Importe = s2n(grillacheques.TextMatrix(x, 2))
                    cueBan = sSinNull(obtenerDeSQL("select cuenta_con from ctasbank where codigo = '" & x2s(txtcodcuenta) & "' "))
                    
                    DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", val(txtcodcuenta), "D", "deposito de Cheque", fechaoperacion, "C" _
                        , val(grillacheques.TextMatrix(x, 0)), Importe, maximobanc1, midDoc, Date, UsuarioSistema!codigo
                        
                    'ASIENTO
                    asiCh.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, Importe
                    asiCh.AcumularItem cueBan, Importe, 0
                End If
                
                If optacreditar = True Then

                    DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", depcuenta, "A", "Acreditación de Cheque", fechaoperacion, "C" _
                      , val(grillacheques.TextMatrix(x, 0)), s2n(grillacheques.TextMatrix(x, 2)), maximobanc1, midDoc, fechaoperacion, UsuarioSistema!codigo
                End If
                
                If optrechazar = True Then
                    
                    Importe = s2n(grillacheques.TextMatrix(x, 2))
                    cueBan = sSinNull(obtenerDeSQL("select cuenta_con from ctasbank where codigo = '" & depcuenta & "' "))
                    
                    asse = "optRech chMoviBanc"
                    DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", depcuenta, "R", "Rechazo de Cheque", fechaoperacion, "C" _
                      , val(grillacheques.TextMatrix(x, 0)), s2n(grillacheques.TextMatrix(x, 2)), maximobanc1, midDoc, fechaoperacion, UsuarioSistema!codigo
                    'ASIENTO
                    asiCh.AcumularItem CuentaParam(ID_Cuenta_M_CH_RECHAZADOS), Importe, 0
                    asiCh.AcumularItem cueBan, 0, Importe
                    
                End If
                
                If optcanje = True Then
                    asse = "optCanj chmovibanc"
                    DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", 0, "J", "Canje de Cheque", fechaoperacion, "C" _
                        , val(grillacheques.TextMatrix(x, 0)), s2n(grillacheques.TextMatrix(x, 2)), maximobanc1, midDoc, fechaoperacion, UsuarioSistema!codigo
                    'ASIENTO
                    Importe = s2n(grillacheques.TextMatrix(x, 2))
                    
                    'asiCh.AcumularItem CuentaParam(ID_Cuenta_M_CH_CANJE), importe, 0
                    'asiCh.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, importe
                    
                End If

                maximocaja = nuevoCodigo("MoviCaja", "Movimiento")
                valcartera = CuentaParam(ID_Cuenta_M_CH_CARTERA)
                
                
                                               
                rs.Open "select banco, cuenta_con from Ctasbank where codigo = " & val(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                If Not rs.EOF Then
                    txtbanco = rs!Banco
                    cuentacon = rs!cuenta_con
                End If
                rs.Close
                Set rs = Nothing
                
                
                rs.Open "select cuenta from Cajas where codigo = " & uCajaEfectivo.codigo & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                ' Val(txtcodcaja)'antes
                If Not rs.EOF Then
                    cuentaconcaja = rs!Cuenta
                End If
                rs.Close
                Set rs = Nothing
                                                               
                
                If optdepositar = True Then
                    asse = "optdepo mc ch3"
                    
                    DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", 0, maximocaja, "C", "E", s2n(grillacheques.TextMatrix(x, 2)), "deposito " & ObtenerDescripcion("BancosGrales", val(txtbanco)), _
                            fechaoperacion, val(grillacheques.TextMatrix(x, 0)), valcartera, maximobanc1, midDoc, Date, UsuarioSistema!codigo

'''''                    DataEnvironment1.dbo_INGCHEQUEDETALLE "A", maximocaja, s2n(grillacheques.TextMatrix(x, 2)), 0, cuentacon, "deposito " & ObtenerDescripcion("BancosGrales", Val(txtbanco)), "DC", _
                    fechaoperacion

                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", val(grillacheques.TextMatrix(x, 0)), 0, "", val(cmbplazo), val(txtcodcuenta), "", 0, fechaoperacion, "D", 0, "", 0, 0, 0, 0, midDoc
                End If
                
                If optacreditar = True Then
                    asse = "optacre ch3 "
                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", val(grillacheques.TextMatrix(x, 0)), 0, "", 0, depcuenta, "", 0, fechaoperacion, "A", 0, "", 0, fechaoperacion, 0, 0, 0
                End If
                
                If optrechazar = True Then
                    asse = "optrech ch3 "
                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", val(grillacheques.TextMatrix(x, 0)), 0, "", 0, depcuenta, "", 0, fechaoperacion, "R", 0, "", 0, fechaoperacion, 0, 0, 0
                End If
                Dim valorcaja As Double
                If optcanje = True Then
                    asse = "optcanj ch3"
                    'DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", , maximocaja, "C", "E", s2n(grillacheques.TextMatrix(x, 2)), "Canje de Cheque " & grillacheques.TextMatrix(x, 0), _
                        fechaoperacion, Val(grillacheques.TextMatrix(x, 0)), valcartera, 0, midDoc, Date, UsuarioSistema!codigo
                
                    'DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", Val(txtcodcaja), maximocaja, "E", "I", s2n(grillacheques.TextMatrix(x, 2)), "deposito " & ObtenerDescripcion("BancosGrales", Val(txtbanco)), _
                        fechaoperacion, 0, cuentaconcaja, 0, midDoc, Date, UsuarioSistema!codigo
                    'DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", cuentaconcaja, maximocaja, "E", "I", s2n(grillacheques.TextMatrix(x, 2)), "deposito " & ObtenerDescripcion("BancosGrales", Val(txtbanco)), _
                        fechaoperacion, 0, cuentaconcaja, 0, midDoc, Date, UsuarioSistema!codigo
                    
                    
                                        
                    If pTotalEnCheques >= s2n(grillacheques.TextMatrix(x, 2)) Then
                        pTotalEnCheques = pTotalEnCheques - s2n(grillacheques.TextMatrix(x, 2))
                        valorcaja = 0
                    Else
                        valorcaja = s2n(grillacheques.TextMatrix(x, 2)) - pTotalEnCheques
                        pTotalEnCheques = 0
                    End If
                    aCaja = aCaja + valorcaja
                                        
                    
                    If valorcaja > 0 Then
                        DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", uCajaEfectivo.codigo, maximocaja, "E", "I", valorcaja, "deposito " & ObtenerDescripcion("BancosGrales", val(txtbanco)), _
                            fechaoperacion, 0, cuentaconcaja, 0, midDoc, Date, UsuarioSistema!codigo
                    End If
                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", val(grillacheques.TextMatrix(x, 0)), 0, "", 0, depcuenta, "", 0, fechaoperacion, "J", 0, "", 0, 0, 0, 0, midDoc
                    

                    
                    
                End If
                
                
                If optcobro = True Then
                    If pTotalEnCheques >= s2n(grillacheques.TextMatrix(x, 2)) Then
                        pTotalEnCheques = pTotalEnCheques - s2n(grillacheques.TextMatrix(x, 2))
                        valorcaja = 0
                    Else
                        valorcaja = s2n(grillacheques.TextMatrix(x, 2)) - pTotalEnCheques
                        pTotalEnCheques = 0
                    End If
                    aCaja = aCaja + valorcaja
                                        
                    
                    If valorcaja > 0 Then
                        DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", uCajaEfectivo.codigo, maximocaja, "E", "I", valorcaja, "deposito " & ObtenerDescripcion("BancosGrales", val(txtbanco)), _
                            fechaoperacion, 0, cuentaconcaja, 0, midDoc, Date, UsuarioSistema!codigo
                    End If
                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", val(grillacheques.TextMatrix(x, 0)), 0, "", 0, depcuenta, "", 0, fechaoperacion, "D", 0, "", 0, 0, 0, 0, midDoc
                    
                End If
                
                DataEnvironment1.dbo_GRABARBITACORA Trim(grillacheques.TextMatrix(x, 0)), "Cheques", UsuarioSistema!codigo, Date, Time, "M"
                
            End If
vencido:
        Next
        If optcanje = True Then
            aCanje = s2n(aCaja + aCartera)
            AGastos = s2n(txtGastos)
            If cuentaconcaja = 0 Then
                If aCartera > 0 Then
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), s2n(aCartera - AGastos), 0
                    asiCh.AgregarItem uGastos.codigo, AGastos, 0
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, s2n(aCartera)
                Else
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CANJE), s2n(aCanje), 0
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, aCanje
                End If
            Else
                asiCh.AgregarItem cuentaconcaja, s2n(aCaja - AGastos), 0
                asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), s2n(aCartera), 0
                asiCh.AgregarItem uGastos.codigo, AGastos, 0
                asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, aCanje
            End If
        End If
        
        If optcobro = True Then
            aCanje = s2n(aCaja + aCartera)
            AGastos = s2n(txtGastos)
            If cuentaconcaja = 0 Then
                If aCartera > 0 Then
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), s2n(aCartera - AGastos), 0
                    asiCh.AgregarItem uGastos.codigo, AGastos, 0
                    asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CANJE), 0, s2n(aCartera)
                Else
                    'asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CANJE), s2n(aCanje), 0
                    'asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, aCanje
                End If
            Else
                asiCh.AgregarItem cuentaconcaja, s2n(aCaja - AGastos), 0
                asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), s2n(aCartera), 0
                asiCh.AgregarItem uGastos.codigo, AGastos, 0
                asiCh.AgregarItem CuentaParam(ID_Cuenta_M_CH_CANJE), 0, aCanje
            End If
        End If
        
        If asiCh.CantItems > 0 Then
            If asiCh.Grabar(midDoc, , leerEjercicioId(cboEjercicio)) = 0 Then
                DE_RollbackTrans
                ufa "Err al grabar asiento ", Me.Name & " - " '& sAssert
                Exit Sub
            End If
        End If
        
        DE_CommitTrans
        '***************************************
        
        
'       daTaenvironment1.dbo_DETALLEGTOSTEMP "A", Val(grillacheques.TextMatrix(x, 0)), Trim(grillacheques.TextMatrix(x, 1)), Trim(grillacheques.TextMatrix(x, 2)), s2n(grillacheques.TextMatrix(x, 3))
'        enletras = NroEnLetras(s2n(Replace(txtimporte, ".", ",")))
'
'        daTaenvironment1.LisChequesTerceros
'        rptChequesTerceros.Sections("Encabezado").Controls("lblnumero").Caption = interno
'        rptChequesTerceros.Sections("Encabezado").Controls("lblnumcheque").Caption = txtnumero
'        rptChequesTerceros.Sections("Encabezado").Controls("lblconcepto").Caption = txtconcepto
'        rptChequesTerceros.Sections("Encabezado").Controls("lblcodbanco").Caption = txtcodbanco
'        rptChequesTerceros.Sections("Encabezado").Controls("lblcliente").Caption = txtcodcli
'        rptChequesTerceros.Sections("Encabezado").Controls("lblprocedencia").Caption = cmbprocedencia
'        rptChequesTerceros.Sections("Encabezado").Controls("lblfechacheque").Caption = fechacheque
'        rptChequesTerceros.Sections("Encabezado").Controls("lblfechaingreso").Caption = fechaingreso
'        rptChequesTerceros.Sections("Encabezado").Controls("lblmovcaja").Caption = maximocaja
'        rptChequesTerceros.Sections("Encabezado").Controls("lblmovbanco").Caption = maximobanc1
'
'        rptChequesTerceros.Sections("Medio").Controls("lblenpesos").Caption = enletras
'        rptChequesTerceros.Sections("Medio").Controls("lbltotoperacion").Caption = txtimporte
'
'        rptChequesTerceros.Show vbModal
'        daTaenvironment1.rsLisChequesTerceros.Close
'
'        daTaenvironment1.dbo_DETALLEGTOSTEMP "B", 0, "", "", 0
    
    MsgBox "Operación Realizada con éxito", vbOKOnly
    
    If MsgBox("Imprimir detalle", vbYesNo) = vbYes Then
        ImprimeProceso
    End If
    
    LimpioControles
    Iniciogrillacheques
    habilitogrillachequesenable (False)
    cmdAceptar.enabled = False
    cmdcancelar.enabled = False
    cmbplazo.enabled = False
    txtcodcuenta.enabled = False
    cmbcuenta.enabled = False
    optdepositar = False
    optacreditar = False
    optrechazar = False
    optcanje = False
Else
    MsgBox "Faltan ingresar datos"
End If

fin:
    Set rs = Nothing
    Exit Sub
UFAaceptar:
    DE_RollbackTrans
    ufa "err al grabar ", "aceptar " & asse
    Resume fin
End Sub

Private Function ImprimeProceso()
Dim sTitulo As String, sQue As String, i As Long, iRows As Long
    If optdepositar Then
        sQue = " DEPOSITO "
    ElseIf optacreditar Then
        sQue = " ACREDITACION "
    ElseIf optrechazar Then
        sQue = " RECHAZO "
    ElseIf optcanje Then
        sQue = " CANJE "
    ElseIf optcobro Then
        sQue = " COBRO "
    End If
    sTitulo = "Detalle de " & sQue & " de Cheques"
With gImprime
    .rows = 1
    .cols = 0
    .cols = 5
    .TextMatrix(0, 0) = "NRO INTERNO"
    .ColWidth(0) = 2000
    .TextMatrix(0, 1) = "FECHA"
    .ColWidth(1) = 1000
    .TextMatrix(0, 2) = "IMPORTE"
    .ColWidth(2) = 1000
    .TextMatrix(0, 3) = "CLIENTE"
    .ColWidth(3) = 4000
    .TextMatrix(0, 4) = "NUMERO"
    .ColWidth(4) = 2000
    For i = 1 To grillacheques.rows - 1
        If Trim(grillacheques.TextMatrix(i, 5)) = "X" Then
            .AddItem ""
            iRows = .rows - 1
            .TextMatrix(iRows, 0) = grillacheques.TextMatrix(i, 0)
            .TextMatrix(iRows, 1) = grillacheques.TextMatrix(i, 1)
            .TextMatrix(iRows, 2) = grillacheques.TextMatrix(i, 2)
            .TextMatrix(iRows, 3) = grillacheques.TextMatrix(i, 3)
            .TextMatrix(iRows, 4) = grillacheques.TextMatrix(i, 4)
        End If
    Next
    .AddItem ""
    .TextMatrix(.rows - 1, 3) = "TOTAL"
    For i = 1 To gImprime.rows - 1
        If .TextMatrix(i, 0) <> "" Then
            .TextMatrix(.rows - 1, 2) = s2n(.TextMatrix(.rows - 1, 2)) + s2n(.TextMatrix(i, 2))
        End If
    Next
    
    PrintG gImprime, pVertical, sTitulo, fechaoperacion, sTitulo & "(CTA : " & txtcuenta & ")"
End With
End Function

Private Sub cmdAsientoDcanje_Click()
    frmAsientoManual.Show
End Sub

Private Sub cmdCancelar_Click()
    optacreditar = False
    optdepositar = False
    optrechazar = False
    cmdAceptar.enabled = False
    cmdcancelar.enabled = False
    Iniciogrillacheques
    habilitogrillachequesenable (False)
    habilitoplazoycuenta (False)
    habilitocaja (False)
End Sub

Sub Iniciogrillacheques()
    grillacheques.clear
    'grillacheques.ColWidth(1) = 1700
    grillacheques.TextMatrix(0, 0) = "Nº Interno"
    grillacheques.TextMatrix(0, 1) = "Fecha Ch."
    grillacheques.TextMatrix(0, 2) = "Importe"
    grillacheques.TextMatrix(0, 3) = "Cliente"
    grillacheques.TextMatrix(0, 4) = "Nº Cheque"
    grillacheques.TextMatrix(0, 5) = "Marcar"
    grillacheques.rows = 2
End Sub

Private Sub cmdIChequesT_Click()
FrmIngChequesTerceros.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



'Private Sub fechaoperacion_GotFocus()
''    fechaoperacion.SelStart = 0
''    fechaoperacion.SelLength = Len(fechaoperacion.Text)
'End Sub

'Private Sub fechaoperacion_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub Form_Load()
'    MsgBox "HAY QUE CAMBIAR TODO EL CODIGO DEL BOTON 'ACEPTAR' YA QUE ES UNA COPIA DEL FRMINGCHEQUESTERCEROS"
    uCajaEfectivo.ini "select Responsable from cajas where activo = 1 and codigo = ###", "select Codigo as [ Codigo   ], responsable as [  Responsable      ]  from cajas where activo = 1"
    uGastos.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    cmdAsientoDcanje.Visible = False
    txtGastos.Visible = False
    uGastos.Visible = False
    pTotalEnCheques = 0
    uFecha.ini ucPrimerDiaAnio
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub



Private Sub Form_Resize()
    Anclar fraBoton, Me, anclarAbajo + anclarIzquierda
    Anclar grillacheques, Me, anclarLadosTodos
End Sub

Private Sub grillacheques_Click()
    If sologrilla <> 1 Then
        If grillacheques.TextMatrix(grillacheques.Row, 5) <> "        X" Then
            grillacheques.TextMatrix(grillacheques.Row, 5) = "        X"
            cantidad = str(val(cantidad) + 1)
        Else
            grillacheques.TextMatrix(grillacheques.Row, 5) = ""
            cantidad = str(val(cantidad) - 1)
        End If
    End If
    sologrilla = 0
End Sub

Private Function condi()
    ' puedo poner if's
    condi = " and fecha > " & uFecha.ssFecha & " order by fecha "
    
End Function

Private Sub Cargogrillacheques()
    ' ESTO SE PUEDE LIMPIAR UN TOCO
    Dim rs As New ADODB.Recordset

    relojito

    If optdepositar = True Then
        rs.Open "select * from Cheques where estado = 'C' and activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            cmbplazo.enabled = True
            txtcodcuenta.enabled = True
            cmbcuenta.enabled = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = (sSinNull(rs!Fecha))
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = nSinNull(rs!Nro)
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para depositar"
        End If
    End If
    
    If optacreditar = True Then
        rs.Open "select * from Cheques where estado = 'D' and activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = rs!Fecha
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = rs!Nro
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para depositar"
        End If
    End If
    
    If optrechazar = True Then
        rs.Open "select * from Cheques where estado = 'A' and activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = rs!Fecha
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = rs!Nro
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para depositar"
        End If
    End If
    
    If optcanje = True Then
        rs.Open "select * from Cheques where estado = 'C' and activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            cmdAsientoDcanje.enabled = False
            txtGastos.Visible = True
            uGastos.Visible = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = sSinNull(rs!Fecha)
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = sSinNull(rs!Nro)
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para canjear"
            cmdAsientoDcanje.enabled = False
            txtGastos.Visible = False
            uGastos.Visible = False
        End If
    End If
    
    If optcobro = True Then
        rs.Open "select * from Cheques where estado = 'J' and activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            cmdAsientoDcanje.enabled = False
            txtGastos.Visible = True
            uGastos.Visible = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = sSinNull(rs!Fecha)
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = sSinNull(rs!Nro)
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para cobrar"
            cmdAsientoDcanje.enabled = False
            txtGastos.Visible = False
            uGastos.Visible = False
        End If
    End If
    If optVerTodo = True Then
        rs.Open "select * from Cheques where activo = 1 " & condi() _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            grillacheques.Visible = True
            cmdAceptar.enabled = True
            cmdcancelar.enabled = True
            While Not rs.EOF
                If grillacheques.rows = 2 Then
                    grillacheques.Row = 1
                    grillacheques.Col = 0
                    If Trim(grillacheques.Text) = "" Then
                        grillacheques.Row = 1
                        grillacheques.Col = 0
                        grillacheques.Text = rs!NroInt
                        grillacheques.Col = 1
                        grillacheques.Text = sSinNull(rs!Fecha)
                        grillacheques.Col = 2
                        grillacheques.Text = rs!Importe
                        grillacheques.Col = 3
                        grillacheques.Text = ObtenerDescripcion("Clientes", rs!cliente)
                        grillacheques.Col = 4
                        grillacheques.Text = sSinNull(rs!Nro)
                    Else
                        grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                    End If
                Else
                    grillacheques.AddItem rs!NroInt & Chr(9) & rs!Fecha & Chr(9) & rs!Importe & Chr(9) & ObtenerDescripcion("Clientes", rs!cliente) & Chr(9) & rs!Nro & Chr(9) & Chr(9) & rs!Nro
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox "No hay cheques para nada"
        End If
    End If
    
    relojito False
    
End Sub

Private Sub optacreditar_Click()
    Iniciogrillacheques
    LimpioControles
    Cargogrillacheques
    habilitogrillachequesenable (True)
    habilitocaja (False)
    habilitoplazoycuenta (False)
    lbloperacion.Visible = True
    fechaoperacion.Visible = True
    cmdIChequesT.Visible = False
End Sub

'Private Sub optacreditar_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub optcanje_Click()
    Iniciogrillacheques
    LimpioControles
    Cargogrillacheques
    habilitoplazoycuenta (False)
    habilitocaja (True)
    cmdIChequesT.Visible = True
End Sub

Private Sub optcobro_Click()
    Iniciogrillacheques
    LimpioControles
    Cargogrillacheques
    habilitoplazoycuenta (False)
    habilitocaja (True)
    cmdIChequesT.Visible = True
End Sub

Private Sub optdepositar_Click()
    Iniciogrillacheques
    LimpioControles
    grillacheques.Redraw = False
    Cargogrillacheques
    grillacheques.Redraw = True
    habilitogrillachequesenable (True)
    habilitocaja (False)
    habilitoplazoycuenta (True)
    cmdIChequesT.Visible = False
End Sub

Sub CargarDatos()
Dim rs As New ADODB.Recordset
    
    rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        txtcodcuenta = rs!codigo
        txtcuenta = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub LimpioControles()
    fechaoperacion = Date
    cantidad = ""
    cmbplazo.ListIndex = 0
'''    txtCodCuenta = ""
'''    txtCuenta = ""
    txtbanco = ""
End Sub

Private Sub habilitogrillachequesenable(habilito As Boolean)
    grillacheques.enabled = habilito
End Sub

'
'Private Sub optdepositar_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub


Private Sub optrechazar_Click()
    Iniciogrillacheques
    LimpioControles
    Cargogrillacheques
    habilitoplazoycuenta (False)
    habilitocaja (False)
    lbloperacion.Visible = True
    fechaoperacion.Visible = True
    cmdIChequesT.Visible = False
End Sub

Private Sub optVerTodo_Click()
    Iniciogrillacheques
    LimpioControles
    Cargogrillacheques
    habilitoplazoycuenta (False)
    habilitocaja (False)
    cmdIChequesT.Visible = False
End Sub

'Private Sub optrechazar_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub



'''Private Sub txtcodcaja_GotFocus()
'''    txtcodcaja.SelStart = 0
'''    txtcodcaja.SelLength = Len(txtcodcaja.Text)
'''End Sub

Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

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
        End If
    Else
        If txtcodcuenta <> "" Then
            MsgBox "Codigo de cuenta incorrecto"
            txtcodcuenta = "0"
            txtcodcuenta.SetFocus
        End If
    End If
End Sub

Private Sub habilitoplazoycuenta(habilito As Boolean)
    lbloperacion.Visible = habilito
    fechaoperacion.Visible = habilito
    lblplazo.Visible = habilito
    lblcuenta.Visible = habilito
    cmbplazo.Visible = habilito
    txtcodcuenta.Visible = habilito
    cmbcuenta.Visible = habilito
    txtcuenta.Visible = habilito
End Sub

Private Sub habilitocaja(habilito As Boolean)
    lblcaja.Visible = habilito
'    txtcodcaja.Visible = habilito
'    cmbcaja.Visible = habilito
'    txtcaja.Visible = habilito
    uCajaEfectivo.Visible = habilito
    cmdAsientoDcanje.Visible = habilito
    txtGastos.Visible = habilito
    uGastos.Visible = habilito
    lbloperacion.Visible = habilito
    fechaoperacion.Visible = habilito
    
End Sub

Private Sub txtGastos_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub
