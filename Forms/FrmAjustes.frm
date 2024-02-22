VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmAjustes 
   Caption         =   "Ajustes a Comprobantes"
   ClientHeight    =   7875
   ClientLeft      =   90
   ClientTop       =   450
   ClientWidth     =   7380
   Icon            =   "FrmAjustes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Tag             =   "habilitogrilla"
   Begin Gestion.ucCoDe uCuenta 
      Height          =   315
      Left            =   1440
      TabIndex        =   36
      Top             =   3330
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.TextBox txtserie 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtfinal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   27
      Tag             =   "8"
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox txtvalor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtconc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   3840
      Width           =   4455
   End
   Begin VB.CommandButton cmdcargar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
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
      TabIndex        =   9
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmbeliminofila 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar Fila"
      Enabled         =   0   'False
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
      TabIndex        =   26
      Top             =   5280
      Width           =   975
   End
   Begin VB.OptionButton optdebito 
      Caption         =   "Débito"
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
      Left            =   2280
      TabIndex        =   25
      Tag             =   "1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.OptionButton optcredito 
      Caption         =   "Crédito"
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
      Left            =   600
      TabIndex        =   24
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Framedevol 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   585
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox cargar 
      Height          =   285
      Left            =   5760
      TabIndex        =   23
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
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
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtcodmotivo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmbmotivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Motivo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "2"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtmotivo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txttotal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtnrodoc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtcodprov 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdprov 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedor"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "2"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtprov 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   17104897
      CurrentDate     =   37934
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Bindings        =   "FrmAjustes.frx":08CA
      Height          =   2535
      Left            =   240
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorSel    =   16777215
      BackColorBkg    =   16777215
      Enabled         =   0   'False
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   7170
      Left            =   120
      Top             =   120
      Width           =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   7200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label6 
      Caption         =   "Nº Doc."
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
      TabIndex        =   35
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6240
      TabIndex        =   33
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label7 
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
      Left            =   240
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   30
      Top             =   4320
      Width           =   1215
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Motivo"
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
      TabIndex        =   20
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Total"
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
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Serie"
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
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "FrmAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'16/9/4 '*********************** txtcodcli
'4/1/5 fecha string


Private midDoc As Long

Private Sub cmbeliminofila_Click()
    If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
        If grilla.rows > 1 Then
            txttotal = s2n(txttotal) - s2n(grilla.TextMatrix(grilla.Row, 3))
            If grilla.rows = 2 Then
                grilla.TextMatrix(1, 0) = ""
                grilla.TextMatrix(1, 1) = ""
                grilla.TextMatrix(1, 2) = ""
                grilla.TextMatrix(1, 3) = ""
            Else
                grilla.RemoveItem (grilla.Row)
            End If
            recalculo
        Else
            MsgBox "No hay productos para eliminar o no ha seleccionado ninguno de ellos"
        End If
    End If
End Sub

Private Sub cmbmotivo_Click()
    FrmHelp.Show
    CargarHelp "MotivosAjuste", "Codigo", "Descripcion", "codigo", "descripcion"
    FrmHelp.Tag = "FrmAjustes"
    cargar = "Motivo"
End Sub

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim x As Long
    
    If optCredito = False And optDebito = False Then
        MsgBox "Debe ingresar si el ajuste es por débito o crédito"
        Exit Sub
    End If
    
    If TxtCodProv = "" Then
        MsgBox "Debe ingresar un proveedor"
        Exit Sub
    End If
    
    If txttotal = "" Then
        MsgBox "Debe ingresar un importe"
        Exit Sub
    End If
    
    If TrabaIva(Fecha.Value) Then
        MsgBox "La fecha del comprobante esta dentro de las fechas trabadas para emision," & Chr(13) & "verifiquelo con su contadora.", , "ATENCION"
        Exit Sub
    End If
    
    If grilla.TextMatrix(1, 0) = "" Then
        MsgBox "Debe ingresar la imputación contable"
        Exit Sub
    End If
        
        Dim iddoc As Long, TIPODOC As String, NroDoc As Long
        Dim AsientoCompra As New Asiento
        
        '************************************************
        DE_BeginTrans
        

            
        If optCredito = True Then
            TIPODOC = "APC"
            NroDoc = Val(txtnrodoc)
            iddoc = NuevoDocumento(TIPODOC, NroDoc, Val(TxtCodProv), 0)
            
            'CABECERA asiento
            AsientoCompra.nuevo "Aj " & txtprov, Fecha, "APC"
            'DEBE
            AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(txttotal), 0
            
            DataEnvironment1.dbo_INGCOMPRASCTACTE "A", Fecha, Year(Fecha), Month(Fecha), Val(TxtCodProv), "", "", 0, "APC", Val(txtnrodoc) _
                , 0, 0, s2n(txttotal), 0, s2n(txttotal), Fecha, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
                , Val(txtserie), 0, 0, 0, 0, 0, 0, Val(txtcodmotivo), Date, UsuarioSistema!codigo, iddoc, 0, 0, "", "", 0, 0
        
            For x = 1 To grilla.rows - 1
                'HABER
                AsientoCompra.AcumularItem grilla.TextMatrix(x, 0), 0, grilla.TextMatrix(x, 3)
            Next
            
             DataEnvironment1.dbo_MODIFICONUMAJUSTESCREDITO Val(txtnrodoc)
            If siAsiento("AsientosCompras") Then
                If AsientoCompra.Grabar(iddoc) = 0 Then
                    DE_RollbackTrans
                    Exit Sub
                End If
             End If
            
        Else
            TIPODOC = "APD"
            NroDoc = Val(txtnrodoc)
            iddoc = NuevoDocumento(TIPODOC, NroDoc, Val(TxtCodProv), 0)
            
            'CABECERA asiento
            AsientoCompra.nuevo "Aj " & txtprov, Fecha, "APD"
            'HABER
            AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(txttotal)
        
            DataEnvironment1.dbo_INGCOMPRASCTACTE "A", Fecha, Year(Fecha), Month(Fecha), Val(TxtCodProv), "", "", 0, "APD", Val(txtnrodoc), _
               0, 0, s2n(txttotal), 0, s2n(txttotal), Fecha, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
               Val(txtserie), 0, 0, 0, 0, 0, 0, Val(txtcodmotivo), Date, UsuarioSistema!codigo, iddoc, 0, 0, _
               "", "", 0, 0
        
            For x = 1 To grilla.rows - 1
                'DEBE
                AsientoCompra.AcumularItem grilla.TextMatrix(x, 0), grilla.TextMatrix(x, 3), 0
            Next
            
            DataEnvironment1.dbo_MODIFICONUMAJUSTESDEBITO Val(txtnrodoc)
            If siAsiento("AsientosCompras") Then AsientoCompra.Grabar (iddoc)
            
        End If
        DE_CommitTrans
        '************************************************
        
        Call HabilitoControlesAbajo(False, False, False, True, False, True)
        MsgBox "El ajuste se ha realizado con éxito", vbInformation
        
        ImprimirAjusteCompra
        
        LimpioControles
        HabilitoControles (False)
        habilitogrilla (False)
        InicioGrilla
        
fin:
    Exit Sub
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar ", "alta APC o APD " & txtnrodoc
    Resume fin
End Sub

Private Sub ImprimirAjusteCompra()
If optCredito Then
    RptAjustesGrales.TxtAjuste = "AJUSTE CREDITO A PROVEEDOR "
   Else
    RptAjustesGrales.TxtAjuste = "AJUSTE DEBITO  A PROVEEDOR"
End If
    RptAjustesGrales.TxtCliProv = txtprov
    RptAjustesGrales.TxtFecha = Fecha
    RptAjustesGrales.lblfecha = Date
    RptAjustesGrales.txttotal = txttotal
    RptAjustesGrales.txtNro = Format(txtnrodoc, "00000000")
    RptAjustesGrales.TxtTotalenLetras = enletras(txttotal)
    RptAjustesGrales.Restart
    If PREVIEW_IMPRESIONES Then
        RptAjustesGrales.Show
    Else
        RptAjustesGrales.PrintReport False
    End If
End Sub

Private Sub cmdBuscar_Click()
    If optCredito = False And optDebito = False Then
        MsgBox "Debe marcar si es un débito o crédito"
        Exit Sub
    End If
    
    FrmHelp.Show
    CargarHelpComp "Transcom", "Fecha", "Tipo Doc.", "Nº Doc.", "Importe", "fecha", "tipodoc", "nrodoc", "total", TxtCodProv, "fecha,nrodoc", IIf(optCredito = True, "APC", "APD")
    FrmHelp.Tag = "FrmAjustes"
    Call HabilitoControlesAbajo(True, False, True, False, True, True)
    cargar = "BU"
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    Limpiotextosgrilla
    HabilitoControles (False)
    habilitogrilla (False)
    Call HabilitoControlesAbajo(False, False, False, True, False, True)
    InicioGrilla
End Sub

Private Sub cmdcargar_Click()
Dim Valor As Double
    
     Valor = s2n(txtvalor)
     
     If Valor = 0 Then
        che "debe ingresar un valor"
        Exit Sub
    End If
    
    If uCuenta.codigo = "" Then
        che "debe ingresar una cuenta"
        Exit Sub
    End If
    
    If Valor + s2n(txtfinal) > s2n(txttotal) Then
        che "Con este valor el importe total serìa superado"
        Exit Sub
    End If

    CargogrillaTotal
    Limpiotextosgrilla
        
    If uCuenta.enabled = True And s2n(txttotal) <> s2n(txtfinal) Then
        uCuenta.SetFocus
    End If
End Sub

Private Sub cmdeliminar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim

    Dim mensaje As String, x As Long
    
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        
        If optCredito = True Then
            DataEnvironment1.dbo_INGCOMPRASCTACTE "B", 0, 0, 0, Val(TxtCodProv), "", "", 0, "", Val(txtnrodoc) _
                , 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Val(txtserie), 0, 0, 0, 0, 0, 0, 0, Date, UsuarioSistema!codigo, 0, 0, 0 _
                , "", "", 0, 0
        Else
            DataEnvironment1.dbo_INGCOMPRASCTACTE "B", 0, 0, 0, Val(TxtCodProv), "", "", 0, "", Val(txtnrodoc) _
                , 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 _
                , Val(txtserie), 0, 0, 0, 0, 0, 0, 0, Fecha, UsuarioSistema!codigo, 0, 0, 0 _
                , "", "", 0, 0
        End If
            
            'borr doc y asiento
            BorroDocumento midDoc
            If siAsiento("AsientosCompras") Then
                If Not AsientoBaja_idDoc(midDoc) Then
                    MsgBox "No se pudo borrar asiento." & Chr(13) & "(idDoc " & midDoc & ")", vbCritical
                    'ufa "err no se pudo borrar documento", "middoc " & midDoc
                    DE_RollbackTrans
                    Exit Sub
                End If
            End If
        
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtnrodoc), "Transcom", UsuarioSistema!codigo, Fecha, Time, "B"
        
        Call HabilitoControlesAbajo(True, True, False, False, False, False)
        LimpioControles
        HabilitoControles (False)
        InicioGrilla
    End If
    
fin:
    Exit Sub
UFAelim:
    DE_RollbackTrans
    ufa "err al eliminar", "elim APD o APC " & txtnrodoc
    Resume fin
End Sub

Private Sub cmdnuevo_Click()
    LimpioControles
    HabilitoControles (True)
    Call HabilitoControlesAbajo(True, True, False, False, False, False)
End Sub

Private Sub cmdProv_Click()
    FrmHelp.Show
    CargarHelp "Prov", "Codigo", "Descripcion", "codigo", "descripcion"
    FrmHelp.Tag = "FrmAjustes"
    cargar = "Prov"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    InicioGrilla
    habilitogrilla (False)
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' ", "select cuenta as [ Cuenta           ], Descripcion as [  Descripcion                ] from cuentas where activo = 1 and imputable = 1", True
End Sub

Private Sub optcredito_Click()
Dim rs As New ADODB.Recordset

    rs.Open "select max(num_apc) as maximo from bs", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not IsNull(rs!maximo) Then
        txtnrodoc = rs!maximo + 1
    Else
        txtnrodoc = 1
    End If
    rs.Close
    Set rs = Nothing
    
End Sub

Private Sub optdebito_Click()
Dim rs As New ADODB.Recordset

    rs.Open "select max(num_apd) as maximo from bs", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not IsNull(rs!maximo) Then
        txtnrodoc = rs!maximo + 1
    Else
        txtnrodoc = 1
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub txtcodmotivo_GotFocus()
    txtcodmotivo.SelStart = 0
    txtcodmotivo.SelLength = Len(txtcodmotivo.Text)
End Sub

Private Sub txtcodmotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If

End Sub

Private Sub txtcodmotivo_LostFocus()
    If Trim(txtcodmotivo) <> "" Then
        txtMotivo = ObtenerDescripcion("MotivosAjuste", Val(txtcodmotivo))
        If txtMotivo = "" Then
            MsgBox "Motivo incorrecto"
            txtcodmotivo.SetFocus
        End If
    End If
End Sub

Private Sub txtcodprov_GotFocus()
    TxtCodProv.SelStart = 0
    TxtCodProv.SelLength = Len(TxtCodProv.Text)
End Sub

Private Sub txtcodprov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcodprov_LostFocus()
    If Trim(TxtCodProv) <> "" Then
        txtprov = ObtenerDescripcion("Prov", Val(TxtCodProv))
        If txtprov = "" Then
            MsgBox "Proveedor incorrecto"
            TxtCodProv.SetFocus
        End If
    End If
End Sub

Public Sub CargarDatos()

Dim rs As New ADODB.Recordset
Dim codigo As String
Dim opc As String
   
    codigo = Trim(Me.Tag)
        
    If cargar = "Prov" Then
        rs.Open "select * from Prov where codigo = " & Val(TxtCodProv) & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            TxtCodProv = rs!codigo
            txtprov = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "Motivo" Then
        rs.Open "select * from MotivosAjuste where codigo = " & Val(txtcodmotivo) & " and activo = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            txtcodmotivo = rs!codigo
            txtMotivo = rs!DESCRIPCION
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    If cargar = "BU" Then
        If optCredito = True Then
            opc = "APC"
        Else
            opc = "APD"
        End If
        
        rs.Open "select * from Transcom where nrodoc = " & Val(txtnrodoc) & " and tipodoc='" & Trim(opc) & "' and activo = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic

        If Not rs.EOF Then
            midDoc = nSinNull(rs!iddoc)
            If rs!TIPODOC = "APC" Then
                optCredito = True
            Else
                optDebito = True
            End If
            TxtCodProv = rs!CODPR
            txtprov = ObtenerDescripcion("Prov", rs!CODPR)
            Fecha = rs!Fecha
            txtserie = nSinNull(rs!Serie)
            txtnrodoc = rs!NroDoc
            txttotal = rs!Total
            txtcodmotivo = rs!motivoajuste
            txtMotivo = ObtenerDescripcion("Motivosajuste", rs!motivoajuste)
        End If
        rs.Close
        Set rs = Nothing
    End If
        
End Sub

Private Sub LimpioControles()
    optCredito = False
    optDebito = False
    TxtCodProv = ""
    txtprov = ""
    Fecha = Date
    txttotal = ""
    txtserie = ""
    txtnrodoc = ""
    txtcodmotivo = ""
    txtMotivo = ""
    
    uCuenta.clear
    
    txtconc = ""
    txtvalor = ""
    txtfinal = "0"
    cargar = ""
    midDoc = 0
End Sub

Private Sub HabilitoControles(habilito As Boolean)
    TxtCodProv.enabled = habilito
    txtprov.enabled = habilito
    Fecha.enabled = habilito
    txttotal.enabled = habilito
    txtcodmotivo.enabled = habilito
    txtMotivo.enabled = habilito
    cmdProv.enabled = habilito
    cmbmotivo.enabled = habilito
    txtserie.enabled = habilito
End Sub

Sub Habilitobotones(buscar As Boolean, agregar As Boolean, eliminar As Boolean, aceptar As Boolean, Cancelar As Boolean)
    cmdbuscar.enabled = buscar
    cmdcancelar.enabled = Cancelar
    cmdeliminar.enabled = eliminar
    cmdnuevo.enabled = agregar
    cmdAceptar.enabled = aceptar
End Sub

Private Sub txtconc_GotFocus()
    txtconc.SelStart = 0
    txtconc.SelLength = Len(txtconc.Text)
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)
    Label11.enabled = habilito
    uCuenta.enabled = habilito
    
    Label7.enabled = habilito
    txtconc.enabled = habilito
    Label9.enabled = habilito
    txtvalor.enabled = habilito
    cmdcargar.enabled = habilito
    grilla.enabled = habilito
    cmbeliminofila.enabled = habilito
    txtfinal.enabled = habilito
End Sub

Sub habilitogrilla(habilito As Boolean)
    Label11.Visible = habilito
    
    uCuenta.Visible = habilito
    
    Label7.Visible = habilito
    txtconc.Visible = habilito
    Label9.Visible = habilito
    txtvalor.Visible = habilito
    cmdcargar.Visible = habilito
    grilla.Visible = habilito
    cmbeliminofila.Visible = habilito
    Label8.Visible = habilito
    txtfinal.Visible = habilito
End Sub

Private Sub Cargogrilla(interno As Long)
Dim rs1 As New ADODB.Recordset

    rs1.Open "select DetalleMovcajas.* from DetalleMovcajas inner join Movicaja on Movicaja.movimiento = DetalleMovcajas.movimiento where Movicaja.interno = " & interno & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    If Not rs1.EOF Then
        InicioGrilla
        grilla.rows = 2
        grilla.Row = 0
        While Not rs1.EOF
            grilla.Row = grilla.Row + 1
            grilla.TextMatrix(grilla.Row, 0) = rs1!Cuenta
            grilla.TextMatrix(grilla.Row, 1) = ObtenerDescripcion("Cuentas", Val(rs1!Cuenta))
            grilla.TextMatrix(grilla.Row, 2) = rs1!concepto
            grilla.TextMatrix(grilla.Row, 3) = rs1!Importe
            If txttotal <> "" Then
                txttotal = s2n(txttotal) + s2n(rs1!Importe)
            Else
                txttotal = s2n(rs1!Importe)
            End If
            rs1.MoveNext
            If Not rs1.EOF Then
                grilla.rows = grilla.rows + 1
            End If
        Wend
    End If
    rs1.Close
    Set rs1 = Nothing
End Sub

Private Sub CargogrillaTotal()
Dim Valor As Double

    If grilla.rows = 2 Then
        grilla.Row = 1
        grilla.Col = 0
        If Trim(grilla.Text) = "" Then
            grilla.Row = 1
            grilla.Col = 0
            grilla.Text = uCuenta.codigo 'txtcuentacod
            grilla.Col = 1
            grilla.Text = uCuenta.DESCRIPCION 'txtcuenta
            grilla.Col = 2
            grilla.Text = txtconc
            grilla.Col = 3
            grilla.Text = txtvalor
        Else
            grilla.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
        End If
    Else
        grilla.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
    End If
    
    recalculo
End Sub

Private Sub recalculo()
    Dim i As Long, tot As Double
    
    For i = 1 To grilla.rows - 1
        tot = tot + s2n(grilla.TextMatrix(i, 3))
    Next i
    txtfinal = tot
End Sub

Private Sub txtnrodoc_GotFocus()
    txtnrodoc.SelStart = 0
    txtnrodoc.SelLength = Len(txtnrodoc.Text)
End Sub

Private Sub txtSerie_GotFocus()
    txtserie.SelStart = 0
    txtserie.SelLength = Len(txtserie.Text)
End Sub

Private Sub txtserie_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttotal_GotFocus()
    txttotal.SelStart = 0
    txttotal.SelLength = Len(txttotal.Text)
End Sub

Private Sub txttotal_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTotal_LostFocus()
    
    If Trim(txttotal) = "" Then Exit Sub
    
    If Not IsNumeric(txttotal) Then
        MsgBox "Debe ingresar un importe"
        txttotal.SetFocus
    Else
        habilitogrilla (True)
        habilitogrillaenable (True)
        Dim rs As New ADODB.Recordset

        txttotal = s2n(txttotal)
    End If
    
End Sub

Private Sub Limpiotextosgrilla()
    uCuenta.clear
    txtconc = ""
    txtvalor = ""
End Sub

Sub HabilitoControlesAbajo(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    
    cmdcancelar.enabled = hab1
    cmdAceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdbuscar.enabled = hab6
    
End Sub

Sub InicioGrilla()
    grilla.clear
    grilla.TextMatrix(0, 0) = "Cuenta"
    grilla.TextMatrix(0, 1) = "Descripción"
    grilla.TextMatrix(0, 2) = "Concepto"
    grilla.TextMatrix(0, 3) = "Importe"
    grilla.rows = 2
End Sub

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
    If Trim(txtvalor) = "" Then Exit Sub
    
    habilitogrillaenable (True)
    txtvalor = s2n(txtvalor)
End Sub
