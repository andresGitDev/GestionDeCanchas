VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAjusteDeCosto 
   Caption         =   "Ajuste de costos"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmAjusteDeCosto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar"
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
      Left            =   120
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5565
      Width           =   975
   End
   Begin VB.CommandButton cmdElimi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eliminar"
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
      Left            =   1320
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5565
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Ajuste de debito"
      Height          =   315
      Left            =   5430
      TabIndex        =   40
      Top             =   675
      Width           =   1440
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ajuste de credito"
      Height          =   300
      Left            =   3840
      TabIndex        =   39
      Top             =   675
      Width           =   1545
   End
   Begin VB.TextBox txtMotivo 
      Height          =   285
      Left            =   1005
      TabIndex        =   38
      Top             =   4860
      Width           =   5970
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
      Left            =   9930
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1665
      Visible         =   0   'False
      Width           =   510
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
      Height          =   270
      Left            =   9945
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1350
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtcuentacod 
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Top             =   1245
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
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   840
   End
   Begin VB.TextBox txtcuenta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   18
      Tag             =   "2"
      Top             =   1245
      Width           =   3015
   End
   Begin VB.TextBox txtconc 
      Height          =   285
      Left            =   9345
      TabIndex        =   17
      Top             =   420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txtvalor 
      Height          =   285
      Left            =   9315
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10515
      TabIndex        =   15
      Tag             =   "8"
      Top             =   1665
      Visible         =   0   'False
      Width           =   195
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
      Left            =   5985
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3315
      Width           =   975
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Left            =   1305
      TabIndex        =   13
      Top             =   1620
      Width           =   1335
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
      Height          =   315
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1605
      Width           =   855
   End
   Begin VB.TextBox txtdescodigo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   11
      Tag             =   "2"
      Top             =   1620
      Width           =   3015
   End
   Begin VB.TextBox txtimp 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1305
      TabIndex        =   10
      Top             =   1980
      Width           =   1335
   End
   Begin VB.TextBox txtimporte 
      Height          =   285
      Left            =   9585
      TabIndex        =   9
      Tag             =   "9"
      Top             =   210
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5595
      TabIndex        =   8
      Tag             =   "9"
      Top             =   315
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   4905
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5565
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
      Left            =   3705
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5565
      Width           =   975
   End
   Begin VB.CommandButton cmbvolver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
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
      Left            =   6105
      MaskColor       =   &H00400000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5565
      Width           =   975
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
      Left            =   5985
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2715
      Width           =   975
   End
   Begin VB.TextBox txttotalcentro 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5985
      TabIndex        =   3
      Tag             =   "8"
      Top             =   4275
      Width           =   1095
   End
   Begin VB.TextBox txtimptotal 
      Height          =   285
      Left            =   2025
      TabIndex        =   2
      Tag             =   "9"
      Top             =   675
      Width           =   1335
   End
   Begin VB.TextBox txtNeto 
      Height          =   285
      Left            =   3945
      TabIndex        =   1
      Top             =   1980
      Width           =   1110
   End
   Begin VB.TextBox txtIva 
      Height          =   285
      Left            =   5880
      TabIndex        =   0
      Top             =   1980
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "frmAjusteDeCosto.frx":08CA
      Height          =   555
      Left            =   10050
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   979
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
      Bindings        =   "frmAjusteDeCosto.frx":08DC
      Height          =   1815
      Left            =   225
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2715
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
      Caption         =   "Motivo :"
      Height          =   225
      Left            =   240
      TabIndex        =   37
      Top             =   4890
      Width           =   630
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
      Left            =   240
      TabIndex        =   36
      Top             =   1230
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
      Left            =   8235
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   975
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
      Left            =   8280
      TabIndex        =   34
      Top             =   435
      Visible         =   0   'False
      Width           =   975
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
      Left            =   10800
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Left            =   2985
      TabIndex        =   32
      Top             =   240
      Width           =   2175
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
      Left            =   225
      TabIndex        =   31
      Top             =   1620
      Width           =   975
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
      Left            =   225
      TabIndex        =   30
      Top             =   1980
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   225
      X2              =   7065
      Y1              =   240
      Y2              =   240
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
      Left            =   7650
      TabIndex        =   29
      Top             =   225
      Visible         =   0   'False
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
      Left            =   6225
      TabIndex        =   28
      Top             =   3915
      Width           =   735
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
      Left            =   225
      TabIndex        =   27
      Top             =   675
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   5235
      Left            =   90
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label10 
      Caption         =   "IVA :"
      Height          =   240
      Left            =   5445
      TabIndex        =   26
      Top             =   2010
      Width           =   360
   End
   Begin VB.Label Label12 
      Caption         =   "Neto :"
      Height          =   270
      Left            =   3450
      TabIndex        =   25
      Top             =   2010
      Width           =   510
   End
End
Attribute VB_Name = "frmAjusteDeCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tdocaux As String
Public ndocaux As Long

Private Sub cmbagregarcostos_Click()
Dim Valor As Double
    Valor = s2nt(txtimp)

    If Valor = 0 Then
        MsgBox "Debe ingresar un valor"
        txtimp.SetFocus
        Exit Sub
    End If
    If Trim(txtCodigo) = "" Then
        MsgBox "Debe Ingresar Codigo"
        txtCodigo.SetFocus
        Exit Sub
    End If
    If Valor + s2n(txttotalcentro) > s2n(txtimptotal) Then
        MsgBox "El valor a ingresar supera el importe original"
        Exit Sub
    End If
    If Trim(txtcuentacod.Text) = "" Then
        MsgBox "Debe ingresar la cuenta."
        txtcuentacod.SetFocus
        Exit Sub
    End If
    
    recalcular
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
    CargarHelpCuentas "Cuentas", "cuenta", "Descripcion", "cuenta", "descripcion", "cuenta"
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
    Dim sAssert As String
    Dim rsmax As New ADODB.Recordset
    Dim Nro As Long
    Dim x
    Dim dtFecha As Date
    Dim maximocaja As Long
    
    dtFecha = Date
    
    If MODO_ON_ERROR_ABM_ON Then On Error GoTo UfaOK
    
    If Option1.Value = False And Option2.Value = False Then
        MsgBox "Debe especificar si es ajuste por CREDITO o DEBITO.", vbInformation, "ADVERTENCIA"
        Exit Sub
    End If
    
    If gEMPR_ConSistContable Then
        
        If s2n(frmAjusteDeCosto.txttotalcentro) = 0 Then
            MsgBox "No se realizaron las imputaciones en centro de costos"
            Exit Sub
        End If
    End If
    
    If gEMPR_ConSistContable Then
            sAssert = " dbo_INGCOMPRASDETALLE "
            
            If Option1.Value = True Then
                rsmax.Open "select max(ndoc) as mas from detallecentrocostos where activo=1 and tdoc='AJC'", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                If IsNull(rsmax!mas) Or IsEmpty(rsmax!mas) Then
                    Nro = 1
                Else
                    Nro = rsmax!mas + 1
                End If
                
                Set rsmax = Nothing
            Else
                rsmax.Open "select max(ndoc) as mas from detallecentrocostos where activo=1 and tdoc='AJD'", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                If IsNull(rsmax!mas) Or IsEmpty(rsmax!mas) Then
                    Nro = 1
                Else
                    Nro = rsmax!mas + 1
                End If
                
                Set rsmax = Nothing
            End If
            
            sAssert = " dbo_INGCENTROCOSTOS "
            
            If Option1.Value = "1" Then 'ajuste de credito
            
                'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                'el tipo de documento "AJC" es ajuste de credito y "AJD" ajuste de debito
                For x = 1 To frmAjusteDeCosto.grillacostos.rows - 1
                    
                    DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(frmAjusteDeCosto.grillacostos.TextMatrix(x, 0)), _
                    dtFecha, "AJC", Nro, s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 2)), s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 3)), s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 2)) + s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, Trim(txtMotivo), frmAjusteDeCosto.grillacostos.TextMatrix(x, 4), 0
                    
                Next
            Else           'ajuste de debito
                
                'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                'el tipo de documento "AJC" es ajuste de credito y "AJD" ajuste de debito
                For x = 1 To frmAjusteDeCosto.grillacostos.rows - 1
                
                    DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(frmAjusteDeCosto.grillacostos.TextMatrix(x, 0)), _
                    dtFecha, "AJD", Nro, s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 2)), s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 3)), s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 2)) + s2n(frmAjusteDeCosto.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, Trim(txtMotivo), frmAjusteDeCosto.grillacostos.TextMatrix(x, 4), 0
                                    
                Next
            End If
    End If
        
        frmAjusteDeCosto.LimpioControles
        frmAjusteDeCosto.InicioGrillaCostos
        Option1.Value = False
        Option2.Value = False
    
    MsgBox "Operación Realizada con éxito", vbOKOnly
    
    
FinOK:
    Exit Sub
UfaOK:
    DE_RollbackTrans
    ufa "Err al grabar el ajuste ", Me.Name & " - " & sAssert
    Resume FinOK
End Sub

Private Sub cmdBuscar_Click()
    Dim s As String
    Dim sql As String
    Dim impo As Double
    Dim rs As New ADODB.Recordset
    Dim Aux As New ADODB.Recordset
    
    cmdcancelar.Value = True
    
    s = "select distinct(ndoc) as [ Numero ], tdoc as [  Tipo  ],sum(importe) as [  Importe   ] from detallecentrocostos where (tdoc='AJD' or tdoc='AJC') and activo=1 group by ndoc,tdoc"
    With frmBuscar
        If frmBuscar.MostrarSql(s) > "" Then
            sql = "select * from detallecentrocostos where tdoc='" & .resultado(2) & "' and ndoc=" & .resultado(1)
            rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If .resultado(2) = "AJD" Then
                Option2.Value = True
                Option1.Value = False
            Else
                Option1.Value = True
                Option2.Value = False
            End If
            While Not rs.EOF
                impo = impo + rs!Importe
                rs.MoveNext
            Wend
            txtimptotal = impo
            
            rs.MoveFirst
            'txtMotivo = rs!motivo
            While Not rs.EOF
                                
                txtimp = rs!Importe
                txtneto = rs!Neto
                txtIva = ((rs!ivapesos * 100) / rs!Neto) 'saco iva en %
                txtCodigo.Text = rs!codigo
                s = "select * from CentrodeCostos where codigo=" & rs!codigo
                Aux.Open s, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                txtdescodigo.Text = Aux!DESCRIPCION
                Set Aux = Nothing
                txtcuentacod.Text = rs!Cuenta
                s = "select * from cuentas where cuenta=" & rs!Cuenta
                Aux.Open s, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                txtcuenta.Text = Aux!DESCRIPCION
                cmbagregarcostos.Value = True
                
                rs.MoveNext
                Set Aux = Nothing
            Wend
            
            rs.MoveFirst
            txtMotivo = rs!motivo
            ndocaux = .resultado(1)
            tdocaux = .resultado(2)
                        
        End If
    End With
    
    Set rs = Nothing
    cmdElimi.enabled = True
    
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Limpiotextosgrilla
    InicioGrillaCostos
    Option1.Value = False
    Option2.Value = False
End Sub

Private Sub cmdElimi_Click()
    Dim sql As String
    Dim i As Long
    
    If grillacostos.rows <= 1 Then
        cmdElimi.enabled = False
        Exit Sub
    End If
    i = 1
    While i < grillacostos.rows
        sql = "update detallecentrocostos set activo=0 where tdoc='" & tdocaux & "' and ndoc=" & ndocaux & " and codigo=" & grillacostos.TextMatrix(i, 0)
        DataEnvironment1.Sistema.Execute sql
        i = i + 1
    Wend
    MsgBox "La operacion se ha realizado con exito.", , "Importante"
    cmdcancelar.Value = True
    cmdElimi.enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub
Public Function CargarImputacion(Importe As Double, Total As Double, cta As Long)
On Error GoTo ufaErr
    
    Dim rs As New ADODB.Recordset
    
    LimpioControles
    InicioGrillaCostos
    
    txtimptotal = Total
    
    rs.Open "select dato_fijo from datos", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        If rs!DATO_FIJO = 7 Then
            
            txttotalcentro = ""
            
            txtcuentacod = val(cta)
            txtcuenta = ObtenerDescripcion("cuentas", cta)
            txtconc = "COMPRAS"
            txtvalor = Importe
            
            txtCodigo = "1"
            txtdescodigo = "GENERAL"
            txtimp = Total
            grillacostos.rows = 2
            grillacostos.TextMatrix(1, 0) = ""
                
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
    LimpioControles
    Limpiotextosgrilla
'    InicioGrilla
    InicioGrillaCostos
    cmdElimi.enabled = False
End Sub

Private Sub grillacostos_DblClick()
    Dim C As Long
    If grillacostos.Row <> 0 And grillacostos.TextMatrix(grillacostos.Row, 2) > "" Then
        txtCodigo = grillacostos.TextMatrix(grillacostos.Row, 0)
        txtdescodigo = grillacostos.TextMatrix(grillacostos.Row, 1)
        txtimp = CLng(grillacostos.TextMatrix(grillacostos.Row, 2)) + CLng(grillacostos.TextMatrix(grillacostos.Row, 3))
        txtneto = grillacostos.TextMatrix(grillacostos.Row, 2)
        txtIva = grillacostos.TextMatrix(grillacostos.Row, 3)
        txtcuentacod = grillacostos.TextMatrix(grillacostos.Row, 4)
        txtcuenta = grillacostos.TextMatrix(grillacostos.Row, 5)
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
End Sub


Private Sub txtimptotal_GotFocus()
    txtimptotal.SelStart = 0
    txtimptotal.SelLength = Len(txtimptotal.Text)
End Sub

Private Sub txtimptotal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtimptotal_LostFocus()
    cmbcodigo.SetFocus
End Sub

Private Sub txtiva_LostFocus()
    If txtIva = 0 Or txtIva = "" Then
        Exit Sub
    Else
        txtIva = s2nt(txtIva)
        If txtneto = "" Or txtneto = 0 Then
            Exit Sub
        Else
            txtimp = txtneto * ((txtIva / 100) + 1)
        End If
    End If
End Sub

Private Sub txtneto_LostFocus()
    If txtneto = "" Or txtneto = 0 Then
        Exit Sub
    Else
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

Private Sub txttotalcentro_GotFocus()
    txttotalcentro.SelStart = 0
    txttotalcentro.SelLength = Len(txttotalcentro.Text)
End Sub

Sub LimpioControles()
    txtcuentacod = ""
    txtimptotal = ""
    txtcuenta = ""
    txtvalor = ""
    txttotalcentro = "0"
    txtCodigo = ""
    txtdescodigo = ""
    txtimp = ""
    txtMotivo = ""
End Sub

Sub InicioGrillaCostos()
    grillacostos.clear
    grillacostos.TextMatrix(0, 0) = "Código"
    grillacostos.TextMatrix(0, 1) = "Descripción"
    grillacostos.TextMatrix(0, 2) = "Neto"
    grillacostos.TextMatrix(0, 3) = "IVA"
    grillacostos.TextMatrix(0, 4) = "Nro Cuenta"
    grillacostos.TextMatrix(0, 5) = "Descripcion Cuenta"
    grillacostos.rows = 2
End Sub

Public Sub CargarDatos()
Dim rs As New ADODB.Recordset, codigo As Long
    
    codigo = val(Trim(Me.Tag))
       
    If cargar = "Cuentas" Then
        If Not noestaenlagrilla(txtcuentacod, GRILLA) And esimputable(txtcuentacod) Then
            rs.Open "select *, _codigo as cod from Cuentas where cuenta= " & val(txtcuentacod) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                txtcuentacod = rs!Cuenta
                txtcuenta = rs!DESCRIPCION
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
        txtCodigo = Trim(str(codigo))
        
        If Not noestaenlagrilla(txtCodigo, GRILLA) Then
            rs.Open "select * from centrodecostos where codigo = " & val(txtCodigo) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
                txtCodigo = rs!codigo
                txtdescodigo = rs!DESCRIPCION
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
    txtcuentacod.enabled = habilito
    cmbcuenta.enabled = habilito
    Label7.enabled = habilito
    txtconc.enabled = habilito
    Label9.enabled = habilito
    txtvalor.enabled = habilito
    cmdcargar.enabled = habilito
    GRILLA.enabled = habilito
    cmbeliminofila.enabled = habilito
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label11.Visible = habilito
    txtcuentacod.Visible = habilito
    cmbcuenta.Visible = habilito
    txtcuenta.Visible = habilito
    Label4.Visible = habilito
    txtconc.Visible = habilito
    Label9.Visible = habilito
    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
    GRILLA.Visible = habilito
    cmbeliminofila.Visible = habilito
    Label8.Visible = habilito
    txttotal.Visible = habilito
End Sub

Public Sub CargogrillaTotalCostos()
Dim Valor As Double

    If grillacostos.TextMatrix(1, 0) = "" Then
        grillacostos.Row = 1
        grillacostos.Col = 0
        grillacostos.Text = txtCodigo
        grillacostos.Col = 1
        grillacostos.Text = txtdescodigo
        grillacostos.Col = 2
        grillacostos.Text = txtneto
        grillacostos.Col = 3
        grillacostos.Text = txtneto * (txtIva / 100)
        grillacostos.Col = 4
        grillacostos.Text = txtcuentacod
        grillacostos.Col = 5
        grillacostos.Text = txtcuenta
    Else
        grillacostos.AddItem txtCodigo & Chr(9) & txtdescodigo & Chr(9) & txtneto & Chr(9) & txtneto * (txtIva / 100) & Chr(9) & txtcuentacod & Chr(9) & txtcuenta
    End If
    If txttotalcentro <> "" Then
        Valor = s2nt(txtimp)
        txttotalcentro = s2nt(txttotalcentro) + Valor
    Else
        txttotalcentro = s2nt(txtimp)
    End If
    txtCodigo = ""
    txtdescodigo = ""
    txtimp = "0"
End Sub

Private Sub Limpiotextosgrilla()
    txtcuentacod = ""
    txtcuenta = ""
    txtvalor = ""
    txtneto = "0"
    txtIva = "0"
    txtMotivo = ""
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

Private Sub txtimp_LostFocus()
    If IsNumeric(txtimp) Then
        If grillacostos.Visible = False Then
            habilitogrillaCostos (True)
        End If
        habilitogrillaenableCostos (True)
        txtimp = s2nt(txtimp)
        recalcular
        cmbagregarcostos.SetFocus
    Else
        If txtimp <> "" Then
            MsgBox "Debe ingresar un importe correcto"
            txtimp = "0"
            recalcular
            cmbagregarcostos.SetFocus
        End If
    End If
End Sub

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

Private Sub recalcular()
    Dim tmp
    Dim tipoiva
    
    If FrmFactProv.txtimporte.enabled = True Then
        tipoiva = s2n(FrmFactProv.cboIva.Text)
    ElseIf frmNotaCredDebCompra.txtimporte.enabled = True Then
        tipoiva = s2n(frmNotaCredDebCompra.cboIva.Text)
    ElseIf frmNotaCredDebCompra.txtimporte.enabled = True Then
        tipoiva = s2n(frmNotaCredDebCompra.cboIva.Text)
    End If
    
    tmp = "B"
    If IsEmpty(tmp) Then
        ufa "prg: no encuentro condicion iva tabla ivas", " recalcular "
    ElseIf tmp = "B" Or tmp = "E" Then
        'txtneto = txtimp
        'txtIva = "0"
    Else
        tmp = obtenerDeSQL(" select Porcentaje from Porcentajesiva where  activo = 1 and iva = " & tipoiva)
        
        If IsEmpty(tmp) Then
            txtneto = 0 'txtimporte
        Else
            txtneto = s2n(s2n(txtimp) / (1 + s2n(tmp)))
            txtIva = s2n(s2n(txtimp) - s2n(txtneto))
        End If
    End If
End Sub

