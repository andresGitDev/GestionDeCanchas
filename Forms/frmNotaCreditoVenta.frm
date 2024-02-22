VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaCreditoVenta 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Emision Nota Credito"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin ProyectoAMR.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      Height          =   435
      Left            =   0
      TabIndex        =   22
      Top             =   5550
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   767
      MsgConfirmaSalir=   "¿ Cerrar formulario ? "
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   "¿ Cancela edicion ?"
      CaptionEliminar =   "&Eliminar"
      CaptionImprimir =   "&Imprimir"
   End
   Begin VB.Frame fra 
      BackColor       =   &H00C0C0C0&
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9555
      Begin ProyectoAMR.ucCoDe uCliente 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   1380
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2295
         Left            =   1380
         TabIndex        =   5
         Top             =   2880
         Width           =   5655
         _cx             =   9975
         _cy             =   4048
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
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4560
         Width           =   945
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   4200
         Width           =   945
      End
      Begin VB.TextBox txtPIVA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8820
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox txtMonto 
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   2340
         Width           =   1635
      End
      Begin VB.ComboBox cmbTipoIva 
         Height          =   315
         Left            =   6840
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2415
      End
      Begin VB.TextBox TxtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   900
         Width           =   1245
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   320
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   975
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   1800
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   5340
         TabIndex        =   0
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   51576833
         CurrentDate     =   38126
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iva:"
         Height          =   255
         Left            =   7380
         TabIndex        =   21
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Neto:"
         Height          =   255
         Left            =   7260
         TabIndex        =   20
         Top             =   4260
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Monto Total:"
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
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2340
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo IVA:"
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
         Height          =   255
         Left            =   6840
         TabIndex        =   15
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comprobante:"
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
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha :"
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
         Height          =   255
         Left            =   4500
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo:"
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
         Height          =   255
         Left            =   540
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FormaPago:"
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
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente:"
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
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1380
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmNotaCreditoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '19/11/4

'Private WithEvents cliente As LiCodigo

Public Enum FTipoNota
    Tipo_NotaCredito
    Tipo_NotaDebito
End Enum
Private mTipoNota As FTipoNota

Private gDESC As Long
Private g As LiGrilla

'
'Private Sub cliente_cambio(Codigo) ' As Integer)
'    On Error Resume Next
'    Dim ac As Variant
'    ac = obtenerDeSQL("select iva, FormaPago from clientes where codigo = " & Codigo)
'
'    cmbFormaPago.ListIndex = BuscarEnCombo(cmbFormaPago, ac(1))
'    cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, ac(0))
'    txtPIVA = s2n(obtenerDeSQL("select porcentaje from porcentajesiva where activo = 1 and iva =  " & ComboCodigo(cmbTipoIva)))
'    rever
'End Sub
''
'Private Sub cmdAceptar_Click()
'    Dim tmpfec As Date, tipoForm As String
'
'    If s2n(txtCodigo) > 0 Then
'        MsgBox "Ya fue Grabada"
'        Exit Sub
'    End If
'
'    If s2n(txtMonto) = 0 Then
'        MsgBox "Monto no Valido"
'        Exit Sub
'    End If
'
''    tipoformu = obtenerDeSQL("select letra from ")
''    'alta
'    txtCodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
'    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then    'If ComboCodigo(cmbTipoIva) = IVA_ConsumidorFinal Then
'        txtTipoDoc = TipoDoc_NCREDITO_B
'        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_B) + 1
'        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_B)
'    Else
'        txtTipoDoc = TipoDoc_NCREDITO_A
'        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_A) + 1
'        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_A)
'    End If
'
'
'    If tmpfec > dtFecha Then
'        MsgBox "Fecha menor que la ultima factura cargada"
'        Exit Sub
'    End If
'
'
'    If confirma("Nro Factura: " & TxtNroFactura) Then
'        If GrabaFactura() Then
''            ucBoton.AceptarOk
'            MsgBox "Nro Factura: " & TxtNroFactura & vbCrLf & "Grabado"
'            ImprimirComprobante s2n(txtCodigo)
'        End If
'    Else
'        txtCodigo = ""
'        TxtNroFactura = ""
'    End If
'
'End Sub

'Private Sub cmdReImprimir_Click()
''    If s2n(txtCodigo) = 0 Then Exit Sub
''    ImprimirComprobante (s2n(txtCodigo))
'End Sub
'
''Private Sub cmdSalir_Click()
''    Unload Me
''End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    'Set cliente = New LiCodigo
    Set g = New LiGrilla

    CentrarMe Me
    fra.Top = 0
    fra.Left = 0
    fra.Height = Me.ScaleHeight
    fra.Width = Me.ScaleWidth
    
    dtFecha = Date

    comboSql cmbFormaPago, "select descripcion, codigo from formaspago where activo = 1"
    comboSql cmbTipoIva, "select descripcion, codigo from ivas"
    
    'cliente.init cmbCliente, txtCodCliente, "clientes", False, False, cmdCliente, "activo = 1"
    uCliente.ini "select descripcion from clientes where codigo = ###", "select codigo as [ Codigo    ], descripcion as [ Cliemte                                              ] from clientes order by descripcion"
    g.init grilla
    gDESC = g.AddCol("Descripcion" & Space(90), "S")
    g.rows = 20
    uMenu.init False, True, False, True, False
End Sub


Private Function GrabaFactura() As Boolean
    On Error GoTo UFAgraba
    
'    Dim i As Long
    Dim asse As String ' assert
    Dim ac As Variant, i As Long
'    Dim cant, prod, formu, puni, plis, pedi, ptot, remi, item, depot
'    Dim codtipodoc, serie, intBajaStock As Integer
    
    GrabaFactura = False
 
'*** una transaccion aqui ..... *********************

    asse = " Actualizo tabla parametros BS "
    
    AumentarParametroN CAMPO_BS_CodFactura_VENTA, s2n(txtCodigo)
    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "A" Then
        AumentarParametroD CAMPO_BS_FecFACTURA_A, dtFecha
        AumentarParametroN CAMPO_BS_NroFACTURA_A, s2n(TxtNroFactura)
    ElseIf TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then
        AumentarParametroD CAMPO_BS_FecFACTURA_B, dtFecha
        AumentarParametroN CAMPO_BS_NroFACTURA_B, s2n(TxtNroFactura)
    Else
        ufa "PrgErr: TipoDoc No reconocido", Me.Name, Err
    End If
    
    asse = "Graba Cabecera "
    ac = obtenerDeSQL("select provincia, cuit from clientes where codigo = " & cliente.Codigo)
    DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, CLng(dtFecha), CLng(dtFecha), ComboCodigo(cmbFormaPago), cliente.Codigo, cliente.descripcion, ac(0), ac(1), ComboCodigo(cmbTipoIva), 0, s2n(txtNeto), s2n(txtPIVA), s2n(txtIva), s2n(txtMonto), s2n(txtMonto), 0, 0, 0, UsuarioActual(), CLng(Date), 0, 0, 0, 0, 0          ' s2n(txtcotizacion),
    
    asse = "Graba detalle"
    For i = 1 To g.rows - 1
        If Trim(g.tx(i, gDESC)) > "" Then
            DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, 0, True, "", Trim(g.tx(i, gDESC)), "", 0, 0, 0, 0, 0, 0
        End If
    Next i
    
    'contabilidad?
    '
    
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura = True
    GoTo FIN
    
UFAgraba:
    ufa "Err al grabar ", Me.Name & " - grabaFactura() - " & asse, Err
FIN:

End Function

'Private Sub txtMonto_Change()
'    rever
'End Sub

Private Sub txtMonto_LostFocus()
    rever
End Sub

Private Sub rever()
    txtMonto = s2n(txtMonto)
    txtIva = s2n(txtMonto * s2n(txtPIVA))
    txtNeto = s2n(s2n(txtMonto) - s2n(txtIva))
End Sub
Private Sub uCliente_cambio(Codigo As Variant)
    On Error Resume Next
    Dim ac As Variant
    ac = obtenerDeSQL("select iva, FormaPago from clientes where codigo = " & Codigo)
   
    cmbFormaPago.ListIndex = BuscarEnCombo(cmbFormaPago, ac(1))
    cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, ac(0))
    txtPIVA = s2n(obtenerDeSQL("select porcentaje from porcentajesiva where activo = 1 and iva =  " & ComboCodigo(cmbTipoIva)))
    rever
End Sub

Public Sub mostrar(que As FTipoNota)
    mTipoNota = que
End Sub

'----------------------MENU -----------------------
Private Sub uMenu_AceptarAlta()
    Dim tmpfec As Date, tipoForm As String

    If s2n(txtCodigo) > 0 Then
        MsgBox "Ya fue Grabada"
        Exit Sub
    End If
    If s2n(txtMonto) = 0 Then
        MsgBox "Monto no Valido"
        Exit Sub
    End If

'    tipoformu = obtenerDeSQL("select letra from ")
'    'alta
    txtCodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then    'If ComboCodigo(cmbTipoIva) = IVA_ConsumidorFinal Then
        If mTipoNota = Tipo_NotaCredito Then
            txtTipoDoc = TipoDoc_NCREDITO_B
        Else
            txtTipoDoc = TipoDoc_NDEBITO_B
        End If
        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_B) + 1
        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_B)
    Else
        If mTipoNota = Tipo_NotaCredito Then
            txtTipoDoc = TipoDoc_NCREDITO_A
        Else
            txtTipoDoc = TipoDoc_NDEBITO_A
        End If
        TxtNroFactura = obtenerParametro(CAMPO_BS_NroFACTURA_A) + 1
        tmpfec = obtenerParametro(CAMPO_BS_FecFACTURA_A)
    End If
    
    If tmpfec > dtFecha Then
        MsgBox "Fecha menor que la ultima factura cargada"
        Exit Sub
    End If
    
    If confirma("Nro Factura: " & TxtNroFactura) Then
        If GrabaFactura() Then
'            ucBoton.AceptarOk
            MsgBox "Nro Factura: " & TxtNroFactura & vbCrLf & "Grabado"
            uMenu.AceptarOk
            ImprimirComprobante s2n(txtCodigo)
        End If
    Else
        txtCodigo = ""
        TxtNroFactura = ""
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    FrmBorrarCbo Me
    'cliente.Codigo = 0
    uCliente.Codigo = 0
End Sub
Private Sub uMenu_HabilitarEdicion(SiNo As Boolean)
    fra.Enabled = SiNo
End Sub
Private Sub uMenu_Imprimir()
    If s2n(txtCodigo) = 0 Then Exit Sub
    ImprimirComprobante (s2n(txtCodigo))
End Sub

Private Sub uMenu_nuevo()
    dtFecha.SetFocus
    'uCliente.SetFocus
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub
'----------------------MENU -----------------------


'19/11/4
'    adapt licodigo, +cmd, +where
'16/12/4
'   ucLiCode cliente,
'   add grilla descripcion
'   unifico frm debito/ credito,
'


