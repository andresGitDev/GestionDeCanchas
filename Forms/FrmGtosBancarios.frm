VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmGastosBancarios 
   Caption         =   "Debitos/Creditos Bancarios"
   ClientHeight    =   9525
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9765
   Icon            =   "FrmGtosBancarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTodo 
      BorderStyle     =   0  'None
      Height          =   7650
      Left            =   -45
      TabIndex        =   13
      Top             =   15
      Width           =   9900
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         Left            =   3600
         TabIndex        =   33
         Text            =   "Ejercicio"
         Top             =   240
         Width           =   990
      End
      Begin Gestion.uCtaBanco uCtaBanco 
         Height          =   345
         Left            =   1185
         TabIndex        =   2
         Top             =   1065
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   609
      End
      Begin VB.TextBox txtIdDoc 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7305
         TabIndex        =   32
         Top             =   105
         Width           =   1335
      End
      Begin VB.TextBox txtMovBanc 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7320
         TabIndex        =   31
         Top             =   510
         Width           =   1335
      End
      Begin VB.TextBox txtnumcta 
         Height          =   285
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1845
         Width           =   6120
      End
      Begin VB.TextBox txtimporte 
         Height          =   285
         Left            =   2565
         TabIndex        =   6
         Top             =   2940
         Width           =   1335
      End
      Begin VB.TextBox txtconcepto 
         Height          =   285
         Left            =   2565
         TabIndex        =   5
         Top             =   2550
         Width           =   6120
      End
      Begin VB.TextBox txttipocta 
         Height          =   285
         Left            =   2565
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1485
         Width           =   6135
      End
      Begin VB.TextBox txtvalor 
         Height          =   285
         Left            =   1605
         TabIndex        =   9
         Top             =   4590
         Width           =   1335
      End
      Begin VB.TextBox txtconc 
         Height          =   330
         Left            =   1605
         TabIndex        =   8
         Top             =   4110
         Width           =   5805
      End
      Begin VB.CommandButton cmdcargar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5925
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4950
         Width           =   975
      End
      Begin VB.CommandButton cmbeliminofila 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Eliminar Fila"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5925
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5550
         Width           =   975
      End
      Begin VB.TextBox txttotal 
         Height          =   285
         Left            =   5925
         TabIndex        =   15
         Tag             =   "8"
         Top             =   7110
         Width           =   1095
      End
      Begin VB.OptionButton optgtos 
         Alignment       =   1  'Right Justify
         Caption         =   "Debitos Bancario"
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
         Left            =   810
         TabIndex        =   0
         Tag             =   "0"
         Top             =   240
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.OptionButton optcredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Crédito Bancario"
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
         Left            =   960
         TabIndex        =   1
         Tag             =   "1"
         Top             =   570
         Width           =   1830
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2490
         Left            =   150
         TabIndex        =   14
         Top             =   4950
         Width           =   5655
         _cx             =   9975
         _cy             =   4392
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
         FixedCols       =   0
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
      Begin Gestion.ucCoDe uCuenta 
         Height          =   375
         Left            =   1635
         TabIndex        =   7
         Top             =   3615
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
      Begin MSComCtl2.DTPicker dtfecha 
         Height          =   255
         Left            =   7245
         TabIndex        =   16
         Top             =   2925
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   450
         _Version        =   393216
         Format          =   188547073
         CurrentDate     =   38052
      End
      Begin VB.Label Label13 
         Caption         =   "Ejercicio"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "id:"
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
         Left            =   6165
         TabIndex        =   30
         Top             =   135
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Movimiento:"
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
         Left            =   6045
         TabIndex        =   29
         Top             =   525
         Width           =   1125
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
         Left            =   270
         TabIndex        =   26
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label12 
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
         Left            =   330
         TabIndex        =   25
         Top             =   2925
         Width           =   975
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
         Left            =   285
         TabIndex        =   24
         Top             =   1845
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto/Responsable:"
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
         Left            =   285
         TabIndex        =   23
         Top             =   2550
         Width           =   2295
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
         Left            =   285
         TabIndex        =   22
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha"
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
         Left            =   6525
         TabIndex        =   21
         Top             =   2925
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   285
         TabIndex        =   20
         Top             =   4110
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   285
         TabIndex        =   19
         Top             =   4590
         Width           =   1215
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
         Left            =   285
         TabIndex        =   18
         Top             =   3630
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL"
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
         Left            =   6165
         TabIndex        =   17
         Top             =   6750
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   165
         X2              =   9765
         Y1              =   3390
         Y2              =   3390
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   7830
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   2990
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucFecha uFechaDesde 
         Height          =   285
         Left            =   1305
         TabIndex        =   28
         Top             =   15
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         FechaInit       =   5
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscare Desde:"
         Height          =   240
         Left            =   30
         TabIndex        =   27
         Top             =   45
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FrmGastosBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '16/9/4  ' ***************** txtcodcli
'14/2/5 li: reempl fecha string

Dim midDoc As Long
Dim mMovCaja As Long   ' para compatibilidad con lo viejo
Dim mMovBanc As Long  '   "  "  "


'grilla
Private Enum gricosos
    griCUEN
    griDESC
    griCONC
    griIMPO
End Enum

'Dim rsefec As New ADODB.Recordset
'Dim Ope As String
'Dim modifico As Boolean
'Dim numero As Long


'Private Sub cmbcotizacion_Click()
'    FrmCotizaciones.cmbmoneda = txtmoneda
'    FrmCotizaciones.cmbmoneda.Enabled = False
'    FrmCotizaciones.Show vbModal
'    txtcotiz = FrmCotizaciones.txtcotizacion '**********************
'    txtConcepto.SetFocus
'End Sub

'Private Sub cmbcta_Click()
'    FrmHelp.Show
'    CargarHelpCuentas "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "Cuentas"
'End Sub

'Private Sub cmbcuenta_Click()
'    FrmHelp.Show
'    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
'    FrmHelp.Tag = Me.Name
'    cargar = "CuentasBank"
'End Sub



Private Sub cmbeliminofila_Click()
'    If grilla.TextMatrix(grilla.row, grilla.col) <> "" Then
'        If grilla.rows > 1 Then
'            txttotal = s2n(txttotal) - s2n(grilla.TextMatrix(grilla.row, 3))
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
    
    If GRILLA.Row > 0 Then GRILLA.RemoveItem GRILLA.Row
    recalGrilla
End Sub
Private Function recalGrilla()
    On Error Resume Next
    Dim i As Long, tot As Double
    For i = 1 To GRILLA.rows - 1
        tot = tot + CDbl(GRILLA.TextMatrix(i, griIMPO))
    Next i
    txttotal = Format(tot, "standard")
End Function
'Private Sub cmdAceptar_Click()
'    Dim rs As New ADODB.Recordset
'    'Dim fecha As Variant
'    Dim i As Long
'    Dim enletras As String
'    Dim iddoc As Long
'
''If txtcodcta = "" Then
''    MsgBox "Debe ingresar el código de cuenta"
''    Exit Sub
''End If
'
'    If uCtaBanco.codigo = 0 Then
'        che "falta cuenta banco"
'        Exit Sub
'    End If
'
''
''If optcredito = False And optgtos = False Then
''    MsgBox "Debe ingresar si es un Gasto o un Crédito Bancario"
''    Exit Sub
''End If
'
'If s2n(txttotal) <> s2n(txtimporte) Then
'    MsgBox "No coincide el importe ingresado con el importe total"
'    Exit Sub
'End If
'
'If Ope <> "" Then
'
''        If txtmoneda <> "Pesos" Then
''            rs.Open "select * from Cotizaciones where moneda = " & ObtenerCodigo("Monedas", txtmoneda) & " and Fecha = cdate('" & Date & "') and activo = 1", daTaenvironment1.Sistema, adOpenStatic, adLockOptimistic
''            If rs.EOF Then
''               MsgBox "La moneda asociada a la caja ingresada no se encuentra actualizada"
''               Exit Sub
''            End If
''        End If
'
''        If txtcotiz <> "" Then
'            If Ope = "A" Then
'                If optgtos = True Then
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "G", Trim(txtconcepto), _
'                       dtfecha, "G", s2n(txtimporte), Val(txtmovbank), Date, UsuarioSistema!codigo, 0, 0, 1
'
'
'                    DataEnvironment1.dbo_MOVICAJAS "A", 0, Val(txtmovcaja), _
'                        0, "G", "E", s2n(txtimporte), Trim(txtconcepto), dtfecha, Trim(txtcuentacon), Val(txtmovbank), s2n(txtcotiz), iddoc, Date, UsuarioSistema!codigo
'
''''                    For i = 1 To grilla.rows - 1
''''                        DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
''''                        s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "GB"
''''                        DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
''''                    Next
'                Else
'                    DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "T", Trim(txtconcepto), _
'                    dtfecha, "T", s2n(txtimporte), Val(txtmovbank), Date, UsuarioSistema!codigo, 0, 0, 1
'
'
'                    DataEnvironment1.dbo_MOVICAJAS "A", 0, Val(txtmovcaja), _
'                        0, "D", "I", s2n(txtimporte), Trim(txtconcepto), dtfecha, Trim(txtcuentacon), Val(txtmovbank), s2n(txtcotiz), iddoc, Date, UsuarioSistema!codigo
'
''''                    For i = 1 To grilla.rows - 1
''''                        DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
''''                        s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "CB"
''''                        DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
''''                    Next
'                End If
''            Else
''                If Ope = "M" Then
''                    If optgtos = True Then
''                        DataEnvironment1.dbo_MOVIBANCOS "M", Val(txtcodcta), "G", Trim(txtConcepto), _
''                        dtfecha, "G", s2n(txtimporte), Val(txtmovbank), 0, 0, 0, 0, 0
''
''                        DataEnvironment1.dbo_MOVICAJAS "M", 0, Val(txtmovcaja), _
''                            0, "G", "E", s2n(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacon), Val(txtmovbank), s2n(txtcotiz), 0, 0, 0
''
''                        DataEnvironment1.Sistema.Execute "delete from DetalleMovCajas where movimiento = " & Val(txtmovcaja) & ""
''
''                        For i = 1 To grilla.rows - 1
''                            DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
''                            s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "GB"
''                            DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
''                        Next
''                    Else
''                        DataEnvironment1.dbo_MOVIBANCOS "M", Val(txtcodcta), "T", Trim(txtConcepto), _
''                        dtfecha, "T", s2n(txtimporte), Val(txtmovbank), 0, 0, 0, 0, 0
''
''                        DataEnvironment1.dbo_MOVICAJAS "M", 0, Val(txtmovcaja), _
''                        0, "D", "I", s2n(txtimporte), Trim(txtConcepto), dtfecha, Trim(txtcuentacon), Val(txtmovbank), s2n(txtcotiz), 0, 0, 0, 0, 0
''
''                        DataEnvironment1.Sistema.Execute "delete from DetalleMovCajas where movimiento = " & Val(txtmovcaja) & ""
''
''                        For i = 1 To grilla.rows - 1
''                            DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
''                            s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "CB"
''                            DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
''                        Next
''                    End If
''                    DataEnvironment1.dbo_GRABARBITACORA Val(Trim(txtmovcaja)), "Usuarios", UsuarioSistema!codigo, Date, Time, "M"
''                End If
'            End If
'
'            enletras = NroEnLetras(Val(txtimporte))
'
'            DataEnvironment1.BuscoMovBancario
'            rptGastoBancario.Sections("Encabezado").Controls("lblnumero").caption = txtmovbank
'            rptGastoBancario.Sections("Encabezado").Controls("lblconcepto").caption = txtconcepto
'            rptGastoBancario.Sections("Encabezado").Controls("lblcuenta").caption = ObtenerDescripcion("TipoCtas", Val(txttipocta))
'            rptGastoBancario.Sections("Encabezado").Controls("lbldesbanco").caption = txtdesbanco
'
'            rptGastoBancario.Sections("Medio").Controls("lblenpesos").caption = enletras
'            rptGastoBancario.Sections("Medio").Controls("lbltotoperacion").caption = txtimporte
'
'            rptGastoBancario.Show vbModal
'            DataEnvironment1.rsBuscoMovBancario.Close
'
'            DataEnvironment1.dbo_DETALLEGTOSTEMP "B", 0, "", "", 0
'
'            MsgBox "La operación fue realizada con éxito"
'            LimpioControles
'            Call Habilitobotones(True, True, True, True, True, True)
'            Call HabilitoControles(False)
'            Call MonedaVisible(False)
'            grilla.Clear
'            InicioGrilla
'            cargar = ""
''        Else
''            MsgBox "Debe ingresar la cotización del día"
''        End If
'Else
'    MsgBox "Operación no válida"
'End If
'
'End Sub


'Private Sub cmdBuscar_Click()
'    cargar = "Movibanc"
'    FrmHelp.Show
'    CargarHelpGtosBanc "MOVIBANC", "Cuenta", "Fecha", "Cuenta", "Fecha", "Movbanco", "Cuenta, fecha"
'    FrmHelp.Tag = Me.Name
'    Call Habilitobotones(True, False, True, True, True, True)
'End Sub
'
'Private Sub cmdCancelar_Click()
'    grilla.Clear
'    InicioGrilla
'    LimpioControles
'    Limpiotextosgrilla
' '   Call HabilitoControles(False)
''    habilitogrillaenable (False)
' '   Call Habilitobotones(True, True, False, False, False, True)
'    Call MonedaVisible(False)
' '   cargar = ""
'End Sub

'Public Sub CargarDatos(nromov As Long, iddoc As Long)
'    Dim rs As New ADODB.Recordset
'    Dim mon As Long, codigo As Long
'    'Dim fecha As Variant
'
''    If rsefec.State = 1 Then
''        rsefec.Close
''        Set rsefec = Nothing
''    End If
'
''    codigo = Val(Trim(Me.Tag))
'
''
''    If cargar = "CuentasBank" Then
''        rs.Open "select * from Ctasbank where codigo = " & uCtaBanco.codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''        If Not rs.EOF Then
''            'txtcodcta = rs!codigo
''            uCtaBanco.codigo = rs!codigo
''            'txtdescuenta = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
''            txtdesbanco = ObtenerDescripcionBancos("BancosGrales", rs!Banco)
''            txttipocta = rs!tipo & " - " & ObtenerDescripcion("TipoCtas", Val(rs!tipo))
''            txtnumcta = rs!numero
''            txtcuentacon = rs!cuenta_con
''            If Not IsNull(rs!moneda) Then
''                txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
''            End If
''            txtconcepto.SetFocus
''        End If
''        rs.Close
''        Set rs = Nothing
'
'''        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
''        If ObtenerCodigo("Monedas", txtmoneda) <> 1 Then
''            rs.Open "select * from Cotizaciones where Fecha = " & ssFecha(dtfecha) & " and moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''            If Not rs.EOF Then
''                txtmoneda = ObtenerDescripcion("Monedas", ObtenerCodigo("Monedas", txtmoneda))
''                txtcotiz = ObtenerCotizacion("Cotizaciones", ObtenerCodigo("Monedas", txtmoneda))
''            Else
''                MsgBox "Debe ingresar la cotización del día"
''                'ACA DEBO LLAMAR AL FORMULARIO DE ABM DE COTIZACIONES
''            End If
''            MonedaVisible (True)
''            rs.Close
''            Set rs = Nothing
''        End If
''
''    End If
'
'    'SACO LOS DATOS DE MOVIBANC CUYA CUENTA (CODIGO) SEA IGUAL A LA ELEGIDA EN EL HELP O A MANO
'
'    If cargar = "Movibanc" Then
'        rsefec.Open "select * from MOVIBANC where activo = 1 and cuenta = " & uCtaBanco.codigo & " and movbanco = " & Val(txtmovbank) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'        If Not rsefec.EOF Then
'            cargodatos
'        End If
'        rsefec.Close
'        Set rsefec = Nothing
'    End If
'
'End Sub



Private Sub cmdcargar_Click()
    Dim Valor As Double ', totalgrilla As Double

    Valor = s2n(txtvalor)
    If Valor = 0 Then Exit Sub
                    
    If s2n(txttotal) + Valor <= s2n(txtimporte) Then
        Cargogrilla
        Limpiotextosgrilla
        If uCuenta.enabled = True Then uCuenta.SetFocus
    Else
        MsgBox "Con este valor el importe total serìa superado", vbInformation
    End If
'                Else
'                    Cargogrilla
''                End If
'            Else
'                MsgBox "El valor a ingresar no puede superar al importe original"
'                txtvalor.SetFocus
'            End If
'        Else
'            totalgrilla = sumogrilla()
'            If totalgrilla - s2n(grilla.TextMatrix(grilla.row, 3)) + s2n(txtvalor) <= s2n(txtimporte) Then
'                grilla.TextMatrix(grilla.row, 0) = uCuenta.codigo 'txtcodcuenta
'                grilla.TextMatrix(grilla.row, 1) = uCuenta.descripcion  'txtcuenta
'                grilla.TextMatrix(grilla.row, 2) = txtconc
'                grilla.TextMatrix(grilla.row, 3) = txtvalor
'                txttotal = sumogrilla()
'                LimpioImputacion
'                modifico = False
'                grilla.SetFocus
'            Else
'                MsgBox "El valor a ingresar no puede superar al total"
'                txtvalor.SetFocus
'            End If
'        End If
'    Else
'        MsgBox "Debe ingresar un valor"
'        txtvalor.SetFocus
'    End If
End Sub

Private Sub LimpioImputacion()
    uCuenta.clear
    txtconc = ""
    txtvalor = ""
End Sub
'
'Function sumogrilla() As Double
'    Dim x As Long
'    Dim Total As Double
'
'    For x = 1 To grilla.rows - 1
'        Total = Total + s2n(grilla.TextMatrix(x, 3))
'    Next
'    sumogrilla = Total
'
'End Function

'Private Sub MuestroGrilla()
''    txtcodcuenta = grilla.TextMatrix(grilla.row, 0)
''    txtcuenta = grilla.TextMatrix(grilla.row, 1)
'    uCuenta.codigo = grilla.TextMatrix(grilla.row, 0)
'    txtconc = grilla.TextMatrix(grilla.row, 2)
'    txtvalor = grilla.TextMatrix(grilla.row, 3)
'End Sub
'Private Sub MonedaVisible(habilito As Boolean)
'    lblmoneda.Visible = habilito
'    lblcotiz.Visible = habilito
'    txtmoneda.Visible = habilito
'    txtcotiz.Visible = habilito
'    cmbcotizacion.Visible = habilito
'End Sub
'Private Sub cmdeliminar_Click()
'
'    'Dim fecha As Variant
''    Dim mensaje As String
''
''    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
''    If mensaje = 6 Then
''        fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
'
'' *************compil
'        'daTaenvironment1.dbo_MOVICAJAS "B", 0, Trim(txtmovcaja), 0, "", "", 0, "", 0, 0, UsuarioSistema!codigo, fecha, 0
'        Dim iddoc As Long
'        iddoc = s2n(lblId)
'        DataEnvironment1.dbo_MOVICAJAS "B", 0, Trim(txtmovcaja), 0, "", "", 0, "", 0, "", 0, 0, iddoc, Date, UsuarioSistema!codigo
'        DataEnvironment1.dbo_MOVIBANCOS "B", 0, "", "", 0, "", 0, Val(txtmovbank), 0, 0, UsuarioSistema!codigo, Date, 0
'
'        DataEnvironment1.dbo_GRABARBITACORA Val(Trim(txtmovcaja)), "", UsuarioSistema!codigo, Date, Time, "B"
''        Call Habilitobotones(True, True, False, False, True, False)
''        Call HabilitoControles(True)
''        LimpioControles
''        InicioGrilla
''    End If
'
'End Sub


'''Private Sub cmdmodificar_Click()
'''    Ope = "M"
'''    Call HabilitoControles(True)
'''    Call Habilitobotones(True, False, False, True, True, True)
'''    habilitogrillaenable (True)
'''    Call MonedaVisible(True)
'''End Sub

'Private Sub cmdnuevo_Click()
''Dim rs As New ADODB.Recordset
'
''    Call HabilitoControles(True)
''    Call Habilitobotones(False, False, False, False, True, True)
''    Call habilitogrilla(True)
''    LimpioControles
'
'    rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not IsNull(rs!maxcodigo) Then
'        txtmovbank = rs!maxcodigo + 1
'    End If
''    rs.Close
''    Set rs = Nothing
'
'    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not IsNull(rs!maxcodigo) Then
'        txtmovcaja = rs!maxcodigo + 1
'    End If
'    rs.Close
'    Set rs = Nothing
'
'    Ope = "A"
'    modifico = False
'End Sub


Private Sub Form_Load()
    InicioGrilla
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' ", "select cuenta as [ Cuenta           ], Descripcion as [  Descripcion                ] from cuentas where activo = 1 and imputable = 1", True
    uMenu.init True, True, False, True, True
    
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
        Label13.Visible = False
    End If
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

'Private Sub grilla_Click()
'    modifico = True
'    MuestroGrilla
'End Sub

Private Sub optcredito_Click()
    On Error Resume Next
    uCtaBanco.SetFocus
End Sub
Private Sub optgtos_Click()
    On Error Resume Next
    uCtaBanco.SetFocus
End Sub

'Private Sub txtcodcta_GotFocus()
'    txtcodcta.SelStart = 0
'    txtcodcta.SelLength = Len(txtcodcta.Text)
'End Sub
'
'Private Sub txtcodcta_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub
'
'Private Sub txtcodcta_LostFocus()
''    If optcredito = False And optgtos = False Then
''        MsgBox "Debe ingresar si es un Gasto o un Crédito Bancario"
''        Exit Sub
''    End If
'
' '   If IsNumeric(txtcodcta) Then
'        Dim rs As New ADODB.Recordset
'
'        rs.Open "select * from Ctasbank where codigo = " & Val(txtcodcta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'        If Not rs.EOF Then
'            txtdescuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
'        End If
'        rs.Close
'        Set rs = Nothing
'
'        If txtdescuenta = "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcta = ""
'            txtcodcta.SetFocus
'        Else
'            cargar = "CuentasBank"
'            CargarDatos
'        End If
''    Else
''        If txtcodcta = "" Then
''            MsgBox "Código de cuenta incorrecto"
''            txtcodcta = ""
''            txtcodcta.SetFocus
''        End If
''    End If
'End Sub



''Private Sub txtcodcuenta_GotFocus()
''Dim rs As New ADODB.Recordset
''
''    rs.Open "select dato_fijo from datos", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''    If Not rs.EOF Then
''        If rs!DATO_FIJO = 7 Then
''            txtcodcuenta = "1"
''            txtcodcuenta.Enabled = False
''            txtcuenta = "COMPRAS"
''            txtconc = "COMPRAS"
''            txtconc.Enabled = False
''            txtvalor = txtimporte
''            txtvalor.Enabled = False
''            cmbcuenta.Enabled = False
''            cmdcargar.Enabled = False
''            cmbeliminofila.Enabled = False
''            Cargogrilla
''        End If
''    End If
''    rs.Close
''
''End Sub

'Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub

'Private Sub txtconc_Change()
'Dim i As Long
'    txtconc.Text = UCase(txtconc.Text)
'    i = Len(txtconc.Text)
'    txtconc.SelStart = i
'End Sub

Private Sub txtconc_GotFocus()
    If uCuenta.codigo = "" Then
        MsgBox "Debe cargar la cuenta"
        uCuenta.SetFocus
    End If
    txtconc = txtconcepto
End Sub

'Private Sub txtcodcuenta_LostFocus()
'    If IsNumeric(txtcodcuenta) Then
'        If Not noestaenlagrilla(txtcodcuenta, grilla) And esimputable(Val(txtcodcuenta)) Then
'            txtcuenta = ObtenerDescripcion("Cuentas", Val(txtcodcuenta))
'            If txtcuenta = "" Then
'                MsgBox "Codigo de cuenta incorrecto"
'                txtcodcuenta = "0"
'                txtcodcuenta.SetFocus
'            Else
'                cargar = "Cuentas"
'                CargarDatos
'            End If
'        Else
'            MsgBox "El concepto ya se encuentra cargado o la cuenta no es imputable"
'            txtcodcuenta = ""
'            txtcodcuenta.SetFocus
'        End If
'    Else
'        If txtcodcuenta <> "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = "0"
'            txtcodcuenta.SetFocus
'        End If
'    End If
'End Sub

Sub LimpioControles()
    dtFecha = Date
'    txtcodcta = ""
'    txtdescuenta = ""
'    txtdesbanco = ""
  '  txttipocta = ""
   ' txtnumcta = ""
   ' txtconcepto = ""
    'txtimporte = ""
'    txtmovbank = ""
'    txtcuentacon = ""
    'txttotal = "0"
'    txtmovcaja = ""
'    txtcotiz = "0"

    FrmBorrarTxt Me
    optCredito.Value = False
    optgtos.Value = True

    uCtaBanco.codigo = 0
    mMovCaja = 0
    mMovBanc = 0
    midDoc = 0

'    Ope = ""
    InicioGrilla
End Sub

Sub cargodatos(NroMov As Long, iddoc As Long)
    Dim sWhe As String
    Dim movcaja As Long
    Dim rsMB As New ADODB.Recordset
        
        
    If iddoc > 0 Then
        sWhe = " where iddoc = '" & iddoc & "' "
    ElseIf NroMov > 0 Then
        sWhe = " where movbanco  = '" & NroMov & "' "
    Else
        ufa "mov no identificado", "gastos bancarios"
        Exit Sub
    End If
    
    With rsMB
        .Open "select * from movibanc " & sWhe, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        uCtaBanco.codigo = !Cuenta
        txttipocta = uCtaBanco.Tipo
        
        dtFecha = !Fecha
        txtimporte = !Importe
        txtIdDoc = nSinNull(!iddoc)
        optgtos = (!OPERACION = "S") 'debito bancario
        optCredito = (!OPERACION = "E") 'credito bancario
        txtconcepto = !DESCRIPCION
        txtMovBanc = !MovBanco
        
        ' con idDoc ya no haria falta
        movcaja = s2n(obtenerDeSQL("select movimiento from movicaja where movbanco = " & !MovBanco))
        
    End With
        
    'si  llego hasta aca sin err, pongo var de modulo
    midDoc = iddoc
    mMovBanc = NroMov
    mMovCaja = movcaja
    
    Set rsMB = Nothing
End Sub

    'Dim rs As New ADODB.Recordset
    'Dim rs1 As New ADODB.Recordset
    'Dim mon As Long, movimiento As Long
    
    'PONGO EL NUM DE CUENTA EN EL PRIMER TEXTO
    
    'txtcodcta = rsefec!cuenta
        
    'TRAIGO LOS CAMPOS RESTANTES DE ESA CUENTA BANCARIA DE LA TABLA CTASBANK
        
'    rs.Open "select * from Ctasbank where codigo = " & uCtaBanco.codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs.EOF Then
'        txtcodcta = rs!codigo
'        uCtaBanco.codigo = rs!codigo
'        rs1.Open "select * from Ctasbank where codigo = " & uCtaBanco.codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'        If Not rs1.EOF Then
'            txtdescuenta = ObtenerDescripcion("BancosGrales", rs1!Banco) & " - " & rs1!numero
'        End If
'        rs1.Close
'        Set rs1 = Nothing
        
'        If rsefec!operacion = "G" Then
'            optgtos = True
'        Else
'            optcredito = True
'        End If
'        txtdesbanco = ObtenerDescripcion("BancosGrales", rs!Banco)
'        txttipocta = rs!Tipo
'        txtnumcta = rs!numero
        'txtconcepto = rsefec!descripcion
'        txtimporte = rsefec!Importe
'        dtfecha = rsefec!Fecha
'        txtcuentacon = rs!cuenta_con
'        txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
    'End If
    'rs.Close
    'Set rs = Nothing
    
''    'TRAIGO LA COTIZACION REGISTRADA
''
'''    Call MonedaVisible(True)
''
''    rs.Open "select * from Cotizaciones where Fecha = " & ssFecha(dtfecha) & " and moneda = " & ObtenerCodigo("Monedas", txtmoneda) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''    If Not rs.EOF Then
''        MonedaVisible (True)
''        txtmoneda = ObtenerDescripcion("Monedas", rs!moneda)
''        txtcotiz = rs!cotizacion
''        lblID = rs!iddoc
''    End If
''    rs.Close
''    Set rs = Nothing
        
        
    'SACO EL NUM DE MOVIMIENTO DE LA TABLA MOVICAJA CUYO MOV. BANCARIO SEA IGUAL AL TRAIDO DE MOVIBANC
    
'    rs.Open "select movimiento from MOVICAJA where movbanco = " & Val(txtmovbank) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs.EOF Then
        'movimiento = rs!movimiento
        'txtmovcaja = rs!movimiento
'    End If
'    rs.Close
'    Set rs = Nothing
        
        
''    rs.Open "select * from DetalleMovcajas where movimiento = " & movimiento & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
''    If Not rs.EOF Then
''        habilitogrilla (True)
''        grilla.rows = 2
''        grilla.row = 0
''        While Not rs.EOF
''            grilla.row = grilla.row + 1
''            grilla.TextMatrix(grilla.row, 0) = rs!cuenta
''            grilla.TextMatrix(grilla.row, 1) = ObtenerDescripcion("Cuentas", Val(rs!cuenta))
''            grilla.TextMatrix(grilla.row, 2) = rs!concepto
''            grilla.TextMatrix(grilla.row, 3) = rs!Importe
''            If txttotal <> "" Then
''                txttotal = s2n(txttotal) + s2n(rs!Importe)
''            Else
''                txttotal = s2n(rs!Importe)
''            End If
''            rs.MoveNext
''            If Not rs.EOF Then
''                grilla.rows = grilla.rows + 1
''            End If
''        Wend
''    End If
''    rs.Close
''    Set rs = Nothing
    
    
    
'End Sub

'Sub HabilitoControles(habilito As Boolean)
''    txtcodcta.Enabled = habilito
''    txtdesbanco.Enabled = habilito
''    txttipocta.Enabled = habilito
''    txtnumcta.Enabled = habilito
'    optCredito.Enabled = habilito
'    optgtos.Enabled = habilito
'    txtConcepto.Enabled = habilito
'    txtimporte.Enabled = habilito
'    dtfecha.Enabled = habilito
'    txtimporte.Enabled = habilito
''    cmbcuenta.Enabled = habilito
''    txtcodcta.Enabled = habilito
'    uCtaBanco.Enabled = habilito
'End Sub

'Sub Habilitobotones(busco As Boolean, Nuevo As Boolean, modifico As Boolean, elimino As Boolean, acepto As Boolean, Cancelo As Boolean)
'    cmdBuscar.Enabled = busco
'    cmdnuevo.Enabled = Nuevo
'    cmdModificar.Enabled = modifico
'    cmdEliminar.Enabled = elimino
'    cmdaceptar.Enabled = acepto
'    cmdcancelar.Enabled = Cancelo
'End Sub


'******************* compil
'Private Sub Form_Unload(Cancel As Integer)
'    If rsefec.State = 1 Then
'        rsefec.Close
'        Set rsefec = Nothing
'    End If
'End Sub

Sub InicioGrilla()
    GRILLA.clear
    'grilla.ColWidth(1) = 1700
    GRILLA.TextMatrix(0, griCUEN) = "Cuenta"
    GRILLA.TextMatrix(0, griDESC) = "Descripción"
    GRILLA.TextMatrix(0, griCONC) = "Concepto"
    GRILLA.TextMatrix(0, griIMPO) = "Importe"
    GRILLA.rows = 1
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label2.Visible = habilito

'    txtcodcuenta.Visible = habilito
'    cmbcta.Visible = habilito
'    txtcuenta.Visible = habilito
    uCuenta.Visible = habilito
    
    Label6.Visible = habilito
    txtconc.Visible = habilito
    Label3.Visible = habilito
    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
    GRILLA.Visible = habilito
    cmbeliminofila.Visible = habilito
    txttotal.Visible = habilito
End Sub

'Private Sub txtconc_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtconcepto_Change()
'Dim i As Long
'    txtconcepto.Text = UCase(txtconcepto.Text)
'    i = Len(txtconcepto.Text)
'    txtconcepto.SelStart = i
'End Sub

Private Sub txtConcepto_GotFocus()
    txtconcepto.SelStart = 0
    txtconcepto.SelLength = Len(txtconcepto.Text)
End Sub


'Private Sub txtcotiz_GotFocus()
'    txtcotiz.SelStart = 0
'    txtcotiz.SelLength = Len(txtcotiz.Text)
'End Sub
'
'Private Sub txtcotiz_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub
'
'
'Private Sub txtcotiz_LostFocus()
'    If Not IsNumeric(txtcotiz) Then
'        MsgBox "Cotización incorrecta"
'        txtcotiz = "0"
'        txtcotiz.SetFocus
'    End If
'End Sub

'Private Sub txtdesbanco_GotFocus()
'    txtdesbanco.SelStart = 0
'    txtdesbanco.SelLength = Len(txtdesbanco.Text)
'End Sub


Private Sub txtimporte_GotFocus()
    txtimporte.SelStart = 0
    txtimporte.SelLength = Len(txtimporte.Text)
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Sub txtimporte_LostFocus()
    If Not IsNumeric(txtimporte) Then
        MsgBox "Debe ingresar un importe correcto"
        txtimporte = "0"
        txtimporte.SetFocus
    Else
        Call habilitogrillaenable(True)
        txtimporte = s2n(txtimporte)
    End If
End Sub

Private Sub Limpiotextosgrilla()
'    txtcodcuenta = ""
'    txtcuenta = ""
    uCuenta.clear
    txtconc = ""
    txtvalor = ""
End Sub


Private Sub Cargogrilla()
    GRILLA.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
    'txttotal = s2n(txttotal) + s2n(txtvalor) 'Valor
    recalGrilla
    If s2n(txttotal) = s2n(txtimporte) Then MsgBox "El detalle ha sido completado"
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)

'    txtcodcuenta.Enabled = habilito
    'cmbcta.Enabled = habilito
    'txtcuenta.Enabled = habilito
    uCuenta.enabled = habilito
    
    txtconc.enabled = habilito
    txtvalor.enabled = habilito
    cmdcargar.enabled = habilito
    GRILLA.enabled = habilito
    cmbeliminofila.enabled = habilito
End Sub


'Private Sub txtmoneda_GotFocus()
'    txtmoneda.SelStart = 0
'    txtmoneda.SelLength = Len(txtmoneda.Text)
'End Sub


'Private Sub txtnumcta_GotFocus()
'    txtnumcta.SelStart = 0
'    txtnumcta.SelLength = Len(txtnumcta.Text)
'End Sub


'Private Sub txttipocta_GotFocus()
'    txttipocta.SelStart = 0
'    txttipocta.SelLength = Len(txttipocta.Text)
'End Sub


Private Sub txtvalor_GotFocus()
    txtvalor.SelStart = 0
    txtvalor.SelLength = Len(txtvalor.Text)
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Sub txtvalor_LostFocus()
    If IsNumeric(txtvalor) Then
'        InicioGrilla
        If GRILLA.Visible = False Then
            habilitogrilla (True)
        End If
        habilitogrillaenable (True)
        txtvalor = s2n(txtvalor)
    Else
        If txtvalor <> "" Then
            MsgBox "Debe ingresar un importe correcto"
            txtvalor = "0"
            txtvalor.SetFocus
        End If
    End If
End Sub

Private Function TaTodo() As Boolean
    If uCtaBanco.codigo = 0 Then
        che "falta cuenta banco"
        Exit Function
    End If
    If s2n(CDbl(txttotal)) <> s2n(txtimporte) Then
        che "No coincide el importe ingresado con el importe total"
        Exit Function
    End If
    TaTodo = True
End Function

Private Function alta() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    'Dim iddoc As Long, nMovBanco As Long, nMovCaja As Long
    Dim asie As New Asiento
    Dim i As Long
    Dim Origen As String
    
    mMovBanc = nuevoCodigo("movibanc", "movBanco")
    mMovCaja = nuevoCodigo("movicaja", "movimiento")

    'txtMovBanc = nMovBanco
    
    DE_BeginTrans
        midDoc = NuevoDocumento("GBK", mMovBanc, 0, 0)
        If optgtos Then
            Origen = "GB"
            DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "S", Trim(txtconcepto), _
                    dtFecha, "G", s2n(txtimporte), mMovBanc, midDoc, Date, UsuarioSistema!codigo, 1
                    
            DataEnvironment1.dbo_MOVICAJAS "A", 0, mMovCaja, _
                    0, "G", "E", s2n(txtimporte), Trim(txtconcepto), dtFecha, uCtaBanco.CuentaContable, _
                    mMovBanc, 0, midDoc, Date, UsuarioSistema!codigo
    '''                    For i = 1 To grilla.rows - 1
    '''                        DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
    '''                        s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "GB"
    '''                        DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
    '''                    Next
            
            asie.nuevo "Gasto Bancario ", dtFecha, "GB"
            asie.AgregarItem uCtaBanco.CuentaContable, 0, s2n(txtimporte) ', "Gasto bancario " &
            
            For i = 1 To GRILLA.rows - 1
                'DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
                        s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "GB"
                'DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
                asie.AgregarItem GRILLA.TextMatrix(i, griCUEN), _
                            CDbl(GRILLA.TextMatrix(i, griIMPO)), 0, _
                            GRILLA.TextMatrix(i, griCONC)
            Next i
        Else
            Origen = "CB"
            DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanco.codigo, "E", Trim(txtconcepto), _
                    dtFecha, "T", s2n(txtimporte), mMovBanc, midDoc, Date, UsuarioSistema!codigo, 1
            
            DataEnvironment1.dbo_MOVICAJAS "A", 0, mMovCaja, _
                    0, "D", "I", s2n(txtimporte), Trim(txtconcepto), dtFecha, _
                    uCtaBanco.CuentaContable, mMovBanc, 0, midDoc, _
                    Date, UsuarioSistema!codigo
                    
    '''                    For i = 1 To grilla.rows - 1
    '''                        DataEnvironment1.dbo_DETMOVCAJAS "A", Val(txtmovcaja), _
    '''                        s2n(grilla.TextMatrix(i, 3)), 0, Trim(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 2)), "CB"
    '''                        DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(i, 0)), Trim(grilla.TextMatrix(i, 1)), Trim(grilla.TextMatrix(i, 2)), s2n(grilla.TextMatrix(i, 3))
    '''                    Next
            asie.nuevo "Credito Bancario ", dtFecha, "CB"
            asie.AgregarItem uCtaBanco.CuentaContable, s2n(txtimporte), 0
            
            For i = 1 To GRILLA.rows - 1
                asie.AgregarItem GRILLA.TextMatrix(i, griCUEN), _
                             0, CDbl(GRILLA.TextMatrix(i, griIMPO)), _
                            GRILLA.TextMatrix(i, griCONC)
            Next i
            
        End If
        
        asie.Grabar midDoc, , leerEjercicioId(cboEjercicio)
        
    
    DE_CommitTrans
    alta = True ' Antes de imprimir, ya se grabo
    che "grabado movimiento " & mMovBanc
    
    Imprimir mMovBanc, Origen
fin:
    Exit Function
UFAalta:
    DE_RollbackTrans
    ufa "err:no se pudo grabar", "alta gasto banc"
    Resume fin
End Function

'Private Sub Imprimir()
'    Dim enletras As String
'    enletras = NroEnLetras(Val(txtImporte))
'
'    DataEnvironment1.BuscoMovBancario
'    rptGastoBancario.Sections("Encabezado").Controls("lblnumero").caption = mMovBanc
'    rptGastoBancario.Sections("Encabezado").Controls("lblconcepto").caption = txtConcepto
'    rptGastoBancario.Sections("Encabezado").Controls("lblcuenta").caption = ObtenerDescripcion("TipoCtas", Val(txttipocta))
'    rptGastoBancario.Sections("Encabezado").Controls("lbldesbanco").caption = uCtaBanco.descripcion
'    rptGastoBancario.Sections("Medio").Controls("lblenpesos").caption = enletras
'    rptGastoBancario.Sections("Medio").Controls("lbltotoperacion").caption = txtImporte
'    rptGastoBancario.Show vbModal
'    DataEnvironment1.rsBuscoMovBancario.Close
'End Sub

Private Sub uCtaBanco_cambio(codigo As Variant)
    txttipocta = uCtaBanco.Tipo
    txtnumcta = uCtaBanco.NroCuenta
End Sub

Private Sub uMenu_AceptarAlta()
    If Not TaTodo Then Exit Sub
    If alta Then uMenu.AceptarOk
End Sub
Private Sub uMenu_BorrarControles()
    LimpioControles
End Sub

Private Sub uMenu_Buscar() ' busca gasto G , falta agregar credito , cual letra?  C ?
    Dim resu
    
    With frmBuscar
        resu = .MostrarSql( _
                    "select CUENTA ,  fecha as [ FECHA   ], movbanco as MOV , importe as [IMPORTE], iddoc  " & _
                    " from movibanc " & _
                    " where activo = 1 and fecha > " & uFechaDesde.ssFecha & _
                    " and  (operacion = 'S'  or (operacion = 'E'  and documento = 'T') ) " & _
                    " order by iddoc desc, movbanco desc ")
        If resu > "" Then
            cargodatos s2n(.resultado(3)), s2n(.resultado(5))
            uMenu.BuscarOK
        End If
    End With
End Sub

Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAbaja
        
        DE_BeginTrans
            If midDoc > 0 Then
                BorroDocumento midDoc
            End If
                
            DataEnvironment1.dbo_MOVICAJAS "B", 0, mMovCaja, 0, "", "", 0, "", 0, "", 0, 0, midDoc, Date, UsuarioSistema!codigo
            DataEnvironment1.dbo_MOVIBANCOS "B", 0, "", "", 0, "", 0, mMovBanc, midDoc, Date, UsuarioSistema!codigo, 1
                    
            DataEnvironment1.dbo_GRABARBITACORA mMovCaja, "", UsuarioSistema!codigo, Date, Time, "B"
        DE_CommitTrans
        
        che "eliminado"
        uMenu.EliminarOK
fin:
    Exit Sub
UFAbaja:
    DE_RollbackTrans
    ufa "prg: Fallo la baja", ""
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fraTodo.enabled = sino
End Sub
Private Sub uMenu_Imprimir()
Dim Origen As String

If optgtos = True Then Origen = "GB"
If optCredito = True Then Origen = "CB"

    Imprimir txtMovBanc, Origen
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
