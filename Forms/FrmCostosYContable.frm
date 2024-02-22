VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmCostosYContable 
   Caption         =   "Centro de costos"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   Icon            =   "FrmCostosYContable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   5565
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIva 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,000%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   5
      EndProperty
      Height          =   285
      Left            =   5895
      TabIndex        =   34
      ToolTipText     =   "ejemplo de iva 21"
      Top             =   1815
      Width           =   1080
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   3960
      TabIndex        =   33
      Top             =   1815
      Width           =   1110
   End
   Begin VB.TextBox txtimptotal 
      Height          =   285
      Left            =   2040
      TabIndex        =   31
      Tag             =   "9"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txttotalcentro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   29
      Tag             =   "8"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.CommandButton cmbagregarcostos 
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
      Left            =   6000
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2565
      Width           =   975
   End
   Begin VB.CommandButton cmbvolver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3735
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4965
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4965
      Width           =   975
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   28
      Tag             =   "9"
      Top             =   345
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtimporte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9795
      TabIndex        =   26
      Tag             =   "9"
      Top             =   585
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtimp 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Text            =   "0"
      Top             =   1815
      Width           =   1335
   End
   Begin VB.TextBox txtdescodigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3975
      TabIndex        =   23
      Tag             =   "2"
      Text            =   "Presione el boton para indicar codigo"
      Top             =   1455
      Width           =   3015
   End
   Begin VB.CommandButton cmbcodigo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Text            =   "0"
      Top             =   1455
      Width           =   1335
   End
   Begin VB.CommandButton cmbeliminocostos 
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
      Left            =   6000
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3165
      Width           =   975
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10245
      TabIndex        =   7
      Tag             =   "8"
      Top             =   2490
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtvalor 
      Height          =   285
      Left            =   9825
      TabIndex        =   6
      Top             =   1230
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox txtconc 
      Height          =   285
      Left            =   9780
      TabIndex        =   5
      Top             =   885
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtcuenta 
      Height          =   285
      Left            =   3975
      TabIndex        =   4
      Tag             =   "2"
      Text            =   "Presione el boton para indicar cuenta"
      Top             =   1080
      Width           =   3015
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
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1035
      Width           =   855
   End
   Begin VB.TextBox txtcuentacod 
      Height          =   285
      Left            =   1335
      TabIndex        =   2
      Text            =   "0"
      Top             =   1080
      Width           =   1335
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
      Height          =   255
      Left            =   10380
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
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
      Height          =   300
      Left            =   10365
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1740
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmCostosYContable.frx":08CA
      Height          =   675
      Left            =   10260
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   555
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   1191
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillacostos 
      Bindings        =   "FrmCostosYContable.frx":08DC
      Height          =   1815
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2580
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label13 
      Caption         =   "Ejemplo: 21"
      Height          =   255
      Left            =   5880
      TabIndex        =   37
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "Neto :"
      Height          =   270
      Left            =   3465
      TabIndex        =   36
      Top             =   1845
      Width           =   510
   End
   Begin VB.Label Label10 
      Caption         =   "IVA :"
      Height          =   240
      Left            =   5460
      TabIndex        =   35
      Top             =   1845
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   4635
      Left            =   105
      Top             =   135
      Width           =   7095
   End
   Begin VB.Label Label6 
      Caption         =   "Importe a imputar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   6240
      TabIndex        =   30
      Top             =   3765
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Importe a imputar:"
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
      Left            =   7860
      TabIndex        =   27
      Top             =   585
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   7080
      Y1              =   240
      Y2              =   240
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
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "CENTRO DE COSTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   240
      Width           =   2175
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
      Left            =   10365
      TabIndex        =   13
      Top             =   2130
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8670
      TabIndex        =   12
      Top             =   915
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8715
      TabIndex        =   11
      Top             =   1245
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   315
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FrmCostosYContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4
'4/1/5

Private Sub cmbagregarcostos_Click()
Dim Valor As Double
    'fix 27/10/4
    Valor = s2nt(txtimp)
    'controles
    If Valor = 0 Then
        MsgBox "Debe ingresar un valor"
        'txtimp.SetFocus
        Exit Sub
    End If
    If Trim(txtCodigo) = "" Then
        MsgBox "Debe Ingresar Codigo"
        txtCodigo.SetFocus
        Exit Sub
    End If
    If Valor + s2n(txttotalcentro) > s2n(txtimptotal) Then
        MsgBox "El valor a ingresar supera el importe original"
        'txtimp.SetFocus
        Exit Sub
    End If
    If Trim(txtcuentacod) = "" Then
        MsgBox "Debe ingresar una cuenta"
        txtcuentacod.SetFocus
        Exit Sub
    End If
    'recalcular
    CargogrillaTotalCostos
    
    Limpiotextosgrilla
    If txtCodigo.enabled = True And txtimptotal <> txttotalcentro Then
        txtCodigo.SetFocus
    End If
End Sub

Private Sub cmbcodigo_Click()
    FrmHelp.Show
    CargarHelp "CentrodeCostos", "Codigo", "Descripcion", "codigo", "descripcion"
    FrmHelp.Tag = Me.Name
    cargar = "Centro"
End Sub

Private Sub cmbcuenta_Click()
    FrmHelp.Show
    'CargarHelpCuentas "Cuentas", "_Codigo", "Descripcion", "_codigo", "descripcion"
    CargarHelpCuentas "Cuentas", "cuenta", "Descripcion", "cuenta", "descripcion"
    FrmHelp.Tag = Me.Name
    cargar = "Cuentas"
End Sub


Private Sub cmbeliminocostos_Click()
    If grillacostos.TextMatrix(grillacostos.Row, grillacostos.Col) <> "" Then
        If grillacostos.rows > 1 Then
            txttotalcentro = s2nt(txttotalcentro) - (CLng(grillacostos.TextMatrix(grillacostos.Row, 2)) + CLng(grillacostos.TextMatrix(grillacostos.Row, 3)))
            If grillacostos.rows = 2 Then
                grillacostos.TextMatrix(1, 0) = ""
                grillacostos.TextMatrix(1, 1) = ""
                grillacostos.TextMatrix(1, 2) = ""
                grillacostos.TextMatrix(1, 3) = ""
                grillacostos.TextMatrix(1, 4) = ""
                grillacostos.TextMatrix(1, 5) = ""
            Else
                grillacostos.RemoveItem (grillacostos.Row)
            End If
        Else
            MsgBox "No hay productos para eliminar o no ha seleccionado ninguno de ellos"
        End If
    End If
End Sub

Private Sub cmbvolver_Click()
    Me.Hide
End Sub

Private Sub cmdAceptar_Click()
    MsgBox "Operación Realizada con éxito", vbOKOnly
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Limpiotextosgrilla
    InicioGrilla
    InicioGrillaCostos
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub


Public Function CargarImputacion(Importe As Double, Total As Double, cta As Long)
On Error GoTo ufaErr
    
    Dim rs As New ADODB.Recordset
    
    LimpioControles
    InicioGrilla
    InicioGrillaCostos
    
    txtimptotal = Total
'    txtimporte = Importe
    
    rs.Open "select dato_fijo from datos", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        If rs!DATO_FIJO = 7 Then
            
'            txtTotal = ""
            txttotalcentro = ""
            
            txtcuentacod = val(cta)
            txtcuenta = ObtenerDescripcion("cuentas", cta)
            txtconc = txtcuenta '"COMPRAS"
'            txtvalor = Importe
            
'            grilla.rows = 2
'            grilla.TextMatrix(1, 0) = ""
'            CargogrillaTotal
            
            'If grillacostos.Visible = True Then
                txtCodigo = "1"
                txtdescodigo = "GENERAL"
                txtimp = Total
                grillacostos.rows = 2
                grillacostos.TextMatrix(1, 0) = ""
                
                'CargogrillaTotalCostos  'con esto cargo la grilla de abajo
                
            'End If
        End If
    End If
    
fin:
    Set rs = Nothing
    Exit Function
ufaErr:
    ufa "", "activate" & Me.Name
    Resume fin
End Function


Private Sub Form_Load()
    txtimptotal.Text = FrmOrdenPago.txttot
End Sub

'Private Sub grilla_DblClick()
'Dim C As long
'    If grilla.Row <> 0 Then
'        txtcuentacod = grilla.TextMatrix(grilla.Row, 0)
'        txtcuenta = grilla.TextMatrix(grilla.Row, 1)
'        txtconc = grilla.TextMatrix(grilla.Row, 2)
'        txtvalor = grilla.TextMatrix(grilla.Row, 3)
'        If grilla.Row = 1 Then
'            For C = 0 To grilla.cols - 1
'                grilla.Col = C
'                grilla.Text = ""
'            Next C
'        Else
'            grilla.RemoveItem (grilla.Row)
'        End If
'        txttotal = txttotal - txtvalor
'    End If
'End Sub

Private Sub grillacostos_DblClick()
    Dim C As Long
    If grillacostos.Row <> 0 And grillacostos.TextMatrix(grillacostos.Row, 2) > "" Then
        txtCodigo = grillacostos.TextMatrix(grillacostos.Row, 0)
        txtdescodigo = grillacostos.TextMatrix(grillacostos.Row, 1)
        'txtimp = grillacostos.TextMatrix(grillacostos.Row, 2)
        txtimp = CLng(grillacostos.TextMatrix(grillacostos.Row, 2)) + CLng(grillacostos.TextMatrix(grillacostos.Row, 3))
        txtneto = grillacostos.TextMatrix(grillacostos.Row, 2)
        txtIva = grillacostos.TextMatrix(grillacostos.Row, 3)
        If grillacostos.Row = 1 Then
            For C = 0 To grillacostos.cols - 1
                grillacostos.Col = C
                grillacostos.Text = ""
            Next C
        Else
            grillacostos.RemoveItem (grillacostos.Row)
        End If
        txttotalcentro = txttotalcentro - txtimp
    End If
End Sub

Private Sub txtcodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtcodigo_LostFocus()

    If txtCodigo <> "" Then
        If Not noestaenlagrilla(txtCodigo, grillacostos) Then
            txtdescodigo = ObtenerDescripcion("centrodecostos", val(txtCodigo))
            If txtdescodigo = "" Then
                MsgBox "Codigo de cuenta incorrecto"
                txtCodigo.SetFocus
            Else
                cargar = "Centro"
                CargarDatos
            End If
        Else
            MsgBox "El concepto ya se encuentra cargado o el código no es imputable"
            txtCodigo = ""
            txtCodigo.SetFocus
        End If
    End If

End Sub

'Private Sub txtconc_Change()
'Dim i As long
'    txtconc.Text = UCase(txtconc.Text)
'    i = Len(txtconc.Text)
'    txtconc.SelStart = i
'End Sub

'Private Sub txtconc_GotFocus()
'    If txtcuentacod = "" Then
'        MsgBox "Debe cargar la cuenta"
'        txtcuentacod.SetFocus
'    End If
'End Sub

Private Sub txtcuenta_GotFocus()
    txtcuenta.SelStart = 0
    txtcuenta.SelLength = Len(txtcuenta.Text)
End Sub

Private Sub txtcuentacod_GotFocus()
    txtcuentacod.SelStart = 0
    txtcuentacod.SelLength = Len(txtcuentacod.Text)
End Sub

Private Sub txtimp_GotFocus()
    txtimp.SelStart = 0
    txtimp.SelLength = Len(txtimp.Text)
    txtimp = txtimptotal
End Sub

'Private Sub txtimporte_GotFocus()
'    txtimporte.SelStart = 0
'    txtimporte.SelLength = Len(txtimporte.Text)
'End Sub

Private Sub txtimptotal_GotFocus()
    txtimptotal.SelStart = 0
    txtimptotal.SelLength = Len(txtimptotal.Text)
End Sub

Private Sub txtiva_LostFocus()
    If txtIva = "0" Or txtIva = "" Then
        Exit Sub
    Else
        txtIva = s2nt(txtIva)
        If txtneto = "" Then
            Exit Sub
        Else
            If txtneto = 0 Then
                Exit Sub
            End If
            txtimp = txtneto * ((txtIva / 100) + 1)
        End If
    End If
End Sub

Private Sub txtneto_LostFocus()
    If txtneto = "" Then
        Exit Sub
    Else
        If txtneto = 0 Then
            Exit Sub
        End If
        txtneto = s2nt(txtneto)
        If txtIva = "" Then
            txtimp = 0
            txtimp = txtneto
            txtIva = 0
        Else
            txtimp = 0
            txtimp = txtneto * ((s2nt(txtIva) / 100) + 1)
        End If
    End If
End Sub

'Private Sub txttotal_GotFocus()
'    txttotal.SelStart = 0
'    txttotal.SelLength = Len(txttotal.Text)
'End Sub

Private Sub txttotalcentro_GotFocus()
    txttotalcentro.SelStart = 0
    txttotalcentro.SelLength = Len(txttotalcentro.Text)
End Sub

'Private Sub txtvalor_GotFocus()
'    txtvalor.SelStart = 0
'    txtvalor.SelLength = Len(txtvalor.Text)
'End Sub

'Private Sub txtvalor_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtvalor_LostFocus()
'    Dim Valor As Double
'
'    Valor = s2n(txtvalor)
'    txtvalor = Valor
'
'   If Valor = 0 Then Exit Sub'

'    If grilla.Visible = False Then
'        habilitogrilla (True)
'    End If
'    habilitogrillaenable (True)
'
'End Sub

'Private Sub cmdcargar_Click()
'Dim Valor As Double
'
'    'fix 27/10/4
'    Valor = s2n(txtvalor)
'    'controles
'    If Valor = 0 Then
'        MsgBox "Debe ingresar un valor"
'        txtvalor.SetFocus
'        Exit Sub
'    End If
'
'    If Trim(txtcuentacod) = "" Then
'        MsgBox "Debe Ingresar Codigo"
'        txtcuentacod.SetFocus
'        Exit Sub
'    End If
    
'    If Valor + s2n(txttotal) > s2n(txtimporte) Then
'        MsgBox "El valor a ingresar supera el importe original"
'        If txtimporte.Enabled Then txtimporte.SetFocus
'        Exit Sub
'    End If
    
'    CargogrillaTotal
    
'    Limpiotextosgrilla
'    If txtcuentacod.Enabled = True And txtimporte <> txttotal Then
'        txtcuentacod.SetFocus
'    End If
'End Sub

'Private Sub cmbeliminofila_Click()
'    If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
'        If grilla.rows > 1 Then
'            txttotal = s2nt(txttotal) - s2nt(grilla.TextMatrix(grilla.Row, 3))
'            If grilla.rows = 2 Then
'                grilla.TextMatrix(1, 0) = ""
'                grilla.TextMatrix(1, 1) = ""
'                grilla.TextMatrix(1, 2) = ""
'                grilla.TextMatrix(1, 3) = ""
'            Else
'                grilla.RemoveItem (grilla.Row)
'            End If
'        Else
'            MsgBox "No hay productos para eliminar o no ha seleccionado ninguno de ellos"
'        End If
'    End If
'End Sub

Sub LimpioControles()
    txtcuentacod = ""
'    txtimporte = ""
    txtimptotal = ""
    txtcuenta = ""
    txtconc = ""
'    txtvalor = ""
'    txtTotal = "0"
    txttotalcentro = "0"
    txtCodigo = ""
    txtdescodigo = ""
    txtimp = ""
End Sub

Sub InicioGrillaCostos()
    grillacostos.clear
    'grillacostos.ColWidth(1) = 1700
    grillacostos.TextMatrix(0, 0) = "Código"
    grillacostos.TextMatrix(0, 1) = "Descripción"
    'grillacostos.TextMatrix(0, 2) = "Importe"
    grillacostos.TextMatrix(0, 2) = "Neto"
    grillacostos.TextMatrix(0, 3) = "IVA"
    grillacostos.TextMatrix(0, 4) = "Nro Cuenta"
    grillacostos.TextMatrix(0, 5) = "Descripcion Cuenta"
    grillacostos.rows = 2
End Sub
Public Sub InicioGrilla()
    grillacostos.clear
    grillacostos.rows = 2
    grillacostos.ColWidth(1) = 1700
    
    grillacostos.TextMatrix(0, 0) = "Cuenta"
    grillacostos.TextMatrix(0, 1) = "Descripción"
    grillacostos.TextMatrix(0, 2) = "Concepto"
    grillacostos.TextMatrix(0, 3) = "Importe"
    
    
    
    ''grilla.Clear
    ''grilla.rows = 2
    ''grilla.ColWidth(1) = 1700
    ''grilla.TextMatrix(0, 0) = "Cuenta"
    ''grilla.TextMatrix(0, 1) = "Descripción"
    ''grilla.TextMatrix(0, 2) = "Concepto"
    ''grilla.TextMatrix(0, 3) = "Importe"
    ''grilla.rows = 2
End Sub

Public Sub CargarDatos()
Dim rs As New ADODB.Recordset, codigo As Long
    
    codigo = val(Trim(Me.Tag))
       
    If cargar = "Cuentas" Then
        'If txtcuentacod = "" Then
'            txtcuentacod = Trim(str(Codigo))
        'End If
        If Not noestaenlagrilla(txtcuentacod, GRILLA) And esimputable(txtcuentacod) Then
            rs.Open "select *,_codigo as cod from Cuentas where cuenta = " & val(txtcuentacod) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtcuentacod = rs!Cuenta
                txtcuenta = rs!DESCRIPCION
'                txtconc.SetFocus
            End If
            rs.Close
            Set rs = Nothing
        Else
            MsgBox "El concepto ya se encuentra cargado"
            txtcuentacod = ""
            txtcuentacod.SetFocus
        End If
    End If
               
    If cargar = "Centro" Then
        'If txtCodigo = "" Then
            txtCodigo = Trim(str(codigo))
        'End If
        If Not noestaenlagrilla(txtCodigo, GRILLA) Then 'estaba en grillacostos
            rs.Open "select * from centrodecostos where codigo = " & val(txtCodigo) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
                txtCodigo = rs!codigo
                txtdescodigo = rs!DESCRIPCION
 '               txtimp.SetFocus
            End If
            rs.Close
            Set rs = Nothing
        Else
            MsgBox "El concepto ya se encuentra cargado"
            txtCodigo.SetFocus
        End If
    End If
               
    
End Sub

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

Private Sub txtcuentacod_LostFocus()

    If txtcuentacod <> "" Then
        If Not noestaenlagrilla(txtcuentacod, GRILLA) And esimputable(val(txtcuentacod)) Then
            txtcuenta = ObtenerDescripcion("Cuentas", val(txtcuentacod))
            If txtcuenta = "" Then
                MsgBox "Codigo de cuenta incorrecto"
                txtcuentacod.SetFocus
            Else
                cargar = "Cuentas"
                CargarDatos
            End If
        Else
            MsgBox "El concepto ya se encuentra cargado o la cuenta no es imputable"
            txtcuentacod = ""
            txtcuentacod.SetFocus
        End If
    End If
        
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)
    Label11.enabled = habilito
'    txtcuentacod.Enabled = habilito
    cmbcuenta.enabled = habilito
    Label7.enabled = habilito
    txtconc.enabled = habilito
    Label9.enabled = habilito
'    txtvalor.Enabled = habilito
    cmdcargar.enabled = habilito
'    grilla.Enabled = habilito
    cmbeliminofila.enabled = habilito
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label11.Visible = habilito
'    txtcuentacod.Visible = habilito
    cmbcuenta.Visible = habilito
    txtcuenta.Visible = habilito
    Label4.Visible = habilito
    txtconc.Visible = habilito
    Label9.Visible = habilito
'    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
'    grilla.Visible = habilito
    cmbeliminofila.Visible = habilito
    Label8.Visible = habilito
'    txtTotal.Visible = habilito
End Sub

'Public Sub CargogrillaTotal()
'    Dim Valor As Double
    
    
'    If grilla.TextMatrix(1, 0) = "" Then
'        grilla.Row = 1
'        grilla.Col = 0
'        grilla.Text = txtcuentacod
'        grilla.Col = 1
'        grilla.Text = txtcuenta
'        grilla.Col = 2
'        grilla.Text = txtconc
'        grilla.Col = 3
'        grilla.Text = txtvalor
'    Else
'        grilla.AddItem txtcuentacod & Chr(9) & txtcuenta & Chr(9) & txtconc & Chr(9) & txtvalor
'    End If
'    If txttotal <> "" Then
'        Valor = s2nt(txtvalor)
'        txtTotal = s2nt(txtTotal) + Valor
'    Else
'        txtTotal = s2nt(txtvalor)
'    End If
'    txtcuentacod = ""
'    txtcuenta = ""
'    txtconc = ""
'    txtvalor = "0"
'End Sub

Public Sub CargogrillaTotalCostos()
Dim Valor As Double

    If grillacostos.TextMatrix(1, 0) = "" Then
        grillacostos.Row = 1
        grillacostos.Col = 0
        grillacostos.Text = txtCodigo
        grillacostos.Col = 1
        grillacostos.Text = txtdescodigo
        grillacostos.Col = 2
        'grillacostos.Text = txtimp
        grillacostos.Text = txtneto
        grillacostos.Col = 3
        grillacostos.Text = txtneto * (txtIva / 100)
        grillacostos.Col = 4
        grillacostos.Text = txtcuentacod
        grillacostos.Col = 5
        grillacostos.Text = txtcuenta
    Else
        'grillacostos.AddItem TxtCodigo & Chr(9) & txtdescodigo & Chr(9) & txtimp
        grillacostos.AddItem txtCodigo & Chr(9) & txtdescodigo & Chr(9) & txtneto & Chr(9) & txtneto * (txtIva / 100) & Chr(9) & txtcuentacod & Chr(9) & txtcuenta
    End If
    If txttotalcentro <> "" Then
        Valor = s2nt(txtimp)
        txttotalcentro = s2nt(txttotalcentro) + Valor
    Else
        txttotalcentro = s2nt(txtimp)
    End If
    txtcuentacod = ""
    txtcuenta = ""
    txtCodigo = ""
    txtdescodigo = ""
    txtimp = "0"
End Sub

Private Sub Limpiotextosgrilla()
    txtcuentacod = ""
    txtcuenta = ""
    txtconc = ""
'    txtvalor = ""
    txtneto = "0"
    txtIva = "0"
End Sub


Private Sub txtimp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub



' ********************************esto no es necesario xq quiere ing neto y iva a mano
'Private Sub txtimp_LostFocus()
'    If IsNumeric(txtimp) Then
'        If grillacostos.Visible = False Then
'            habilitogrillaCostos (True)
'        End If
'        habilitogrillaenableCostos (True)
'        txtimp = s2nt(txtimp)
'        recalcular
'    Else
'        If txtimp <> "" Then
'            MsgBox "Debe ingresar un importe correcto"
'            txtimp = "0"
'            txtimp.SetFocus
'            recalcular
'        End If
'    End If
'End Sub

Private Sub habilitogrillaenableCostos(habilito As Boolean)
    txtCodigo.enabled = habilito
    txtimp.enabled = habilito
    cmbcodigo.enabled = habilito
    cmbagregarcostos.enabled = habilito
    cmbeliminocostos.enabled = habilito
    grillacostos.enabled = habilito
End Sub

Sub habilitogrillaCostos(habilito As Boolean)
    Line1.Visible = habilito
    Label6.Visible = habilito
    txtimptotal.Visible = habilito
    Label1.Visible = habilito
    Label2.Visible = habilito
    Label3.Visible = habilito
    txtCodigo.Visible = habilito
    txtdescodigo.Visible = habilito
    txtimp.Visible = habilito
    cmbcodigo.Visible = habilito
    cmbagregarcostos.Visible = habilito
    cmbeliminocostos.Visible = habilito
    grillacostos.Visible = habilito
    Label4.Visible = habilito
    txttotalcentro.Visible = habilito
End Sub

'Private Sub txtimp_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtimp_LostFocus()
'    txtimp = s2n(txtimp)
'    recalcular
'End Sub

Private Sub recalcular()
    Dim tmp
    Dim tipoiva
    Dim COD As Long
    'cmbingresar.Enabled = s2n(txtimporte) <> 0

    'If uProv.codigo = 0 Then
    '    txtneto = ""
    '    txtiva = ""
    '    txtexento = ""
    'Else
    
        'If FrmFactProv.txtimporte.Enabled = True Then
        If vieneDE = "FrmFactProv" Then
            'tmp = obtenerDeSQL("select letra from ivas where codigo = " & s2n(FrmFactProv.txttipoiva.Text))
            tipoiva = s2n(FrmFactProv.cboIva.Text) 's2n(FrmFactProv.txttipoiva.Text)
        'ElseIf FrmNotaDeCredito.txtimporte.Enabled = True Then
        ElseIf vieneDE = "frmNotaCredDebCompra" Then
            'tmp = obtenerDeSQL("select letra from ivas where codigo = " & s2n(FrmNotaDeCredito.txttipoiva.Text))
            tipoiva = s2n(frmNotaCredDebCompra.cboIva.Text)
        'ElseIf FrmNotaDeDebito.txtimporte.Enabled = True Then
        ElseIf vieneDE = "frmNotaCredDebCompra" Then
            'tmp = obtenerDeSQL("select letra from ivas where codigo = " & s2n(FrmNotaDeDebito.txttipoiva.Text))
            tipoiva = s2n(frmNotaCredDebCompra.cboIva.Text)
        ElseIf vieneDE = "frmFacturaVenta" Then
            COD = ComboCodigo(frmFacturaVenta.cmbTipoIva)
            tipoiva = s2n(COD)
        ElseIf vieneDE = "frmNotaCreDebVenta" Then
            COD = ComboCodigo(frmNotaCreDebVenta.cmbTipoIva)
            tipoiva = s2n(COD)
        End If
        
        
        tmp = obtenerDeSQL("select letra from ivas where codigo = " & tipoiva)
        If IsEmpty(tmp) Then
            ufa "prg: no encuentro condicion iva tabla ivas", " recalcular "
        ElseIf tmp = "B" Or tmp = "E" Then
            txtneto = txtimp
            txtIva = "0"
        Else
            'tmp = obtenerDeSQL(" select Porcentaje from Porcentajesiva where  activo = 1 and iva = " & s2n(FrmFactProv.txttipoiva.Text))
            tmp = obtenerDeSQL(" select Porcentaje from Porcentajesiva where  activo = 1 and iva = " & tipoiva)
            
            If IsEmpty(tmp) Then
                txtneto = 0 'txtimporte
                txtIva = 0
            Else
                txtneto = s2n(s2n(txtimp) / (1 + s2n(tmp)))
                txtIva = s2n(s2n(txtimp) - s2n(txtneto))
            End If
        End If
    'End If
End Sub

'Private Sub cmbagregarcostos_Click()
'Dim valor As Double
'
'    If txtimp <> "" Then
'        valor = s2nt(txtimp)
'
'        If (valor <= s2nt(txtimptotal)) And (valor + s2nt(txttotalcentro) <= s2nt(txtimptotal)) Then
'
'            If txttotalcentro <> "" Then
'                If s2nt(txttotalcentro) + valor <= s2nt(txtimptotal) Then
'                    CargogrillaTotalCostos
'                Else
'                    MsgBox "Con este valor el importe total serìa superado", vbInformation
'                End If
'            Else
'                CargogrillaTotalCostos
'            End If
'            Limpiotextosgrilla
'            If txtcodigo.Enabled = True And txtimptotal <> txttotalcentro Then
'                txtcodigo.SetFocus
'            End If
'        Else
'            MsgBox "El valor a ingresar no puede superar al importe original"
'            txtimp.SetFocus
'        End If
'    Else
'        MsgBox "Debe ingresar un valor"
'        txtimp.SetFocus
'    End If
'End Sub
'Private Sub cmbeliminocostos_Click()
'
'
'    If grillacostos.TextMatrix(grillacostos.row, grillacostos.col) <> "" Then
'        If grillacostos.rows > 1 Then
'            txttotalcentro = s2nt(txttotalcentro) - s2nt(grillacostos.TextMatrix(grillacostos.row, 3))
'            If grillacostos.rows = 2 Then
'                grillacostos.TextMatrix(1, 0) = ""
'                grillacostos.TextMatrix(1, 1) = ""
'                grillacostos.TextMatrix(1, 2) = ""
'                grillacostos.TextMatrix(1, 3) = ""
'            Else
'                grillacostos.RemoveItem (grillacostos.row)
'            End If
'        Else
'            MsgBox "No hay productos para eliminar o no ha seleccionado ninguno de ellos"
'        End If
'    End If
'
'
'End Sub
'
'Private Sub cmdcargar_Click()
'Dim valor As Double
'
'    If txtvalor <> "" Then
'        valor = s2nt(txtvalor)
'        If (valor <= s2nt(txtimporte)) And (valor + s2nt(txtTotal) <= s2nt(txtimporte)) Then
'            If txtTotal <> "" Then
'                If s2nt(txtTotal) + valor <= s2nt(txtimporte) Then
'                    CargogrillaTotal
'                Else
'                    MsgBox "Con este valor el importe total serìa superado", vbInformation
'                End If
'            Else
'                CargogrillaTotal
'            End If
'            Limpiotextosgrilla
'            If txtcuentacod.Enabled = True And txtimporte <> txtTotal Then
'                txtcuentacod.SetFocus
'            End If
'        Else
'            MsgBox "El valor a ingresar no puede superar al importe original"
'            txtvalor.SetFocus
'        End If
'    Else
'        MsgBox "Debe ingresar un valor"
'        txtvalor.SetFocus
'    End If
'End Sub

