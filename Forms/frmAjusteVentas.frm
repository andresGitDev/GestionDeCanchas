VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAjusteVentas 
   Caption         =   "Ajuste a Comprobantes Ventas"
   ClientHeight    =   5415
   ClientLeft      =   180
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmAjusteVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   1155
   End
   Begin VB.Frame Framedevol 
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
      Height          =   1005
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton optCredito 
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
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optDebito 
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
         Left            =   180
         TabIndex        =   18
         Tag             =   "1"
         Top             =   660
         Width           =   1335
      End
   End
   Begin VB.Frame fraAjuste 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   60
      TabIndex        =   10
      Top             =   1020
      Width           =   7635
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
      End
      Begin Gestion.ucCoDe uMotivo 
         Height          =   330
         Left            =   1080
         TabIndex        =   4
         Top             =   2340
         Visible         =   0   'False
         Width           =   6315
         _ExtentX        =   9128
         _ExtentY        =   582
         CodigoWidth     =   1000
      End
      Begin Gestion.ucFecha uFecha 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   540
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
         FechaInit       =   0
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   60
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   582
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uCuenta 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1860
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta"
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
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   1920
         Visible         =   0   'False
         Width           =   975
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
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   2340
         Visible         =   0   'False
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
         Left            =   60
         TabIndex        =   13
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente"
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
         TabIndex        =   12
         Top             =   60
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
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1515
      Left            =   0
      TabIndex        =   8
      Top             =   3900
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2672
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucFecha uHasta 
         Height          =   255
         Left            =   6360
         TabIndex        =   16
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         FechaInit       =   0
      End
      Begin Gestion.ucFecha uDesde 
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   450
         FechaInit       =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buscar Entre:"
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
         Left            =   4140
         TabIndex        =   24
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraContable 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   8355
      Begin VB.CommandButton cmdBorrarItem 
         Caption         =   "Borrar Item"
         Height          =   255
         Left            =   7320
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2235
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   7215
         _cx             =   12726
         _cy             =   3942
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
      Begin VB.Label Label11 
         Caption         =   "Cuentas:"
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
         TabIndex        =   9
         Top             =   0
         Width           =   1275
      End
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
      Left            =   4380
      TabIndex        =   23
      Top             =   540
      Width           =   735
   End
   Begin VB.Label txtCodigo 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5220
      TabIndex        =   22
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label9 
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
      Left            =   4380
      TabIndex        =   21
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmAjusteVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' New 1/4/5 no puedo creerlo ... que sea NEW a esta altura

'CURIOSIDAD 1
' uso  Transacciones, a pesar de ser 1 solo SP, para ayudar al multiusuario.

'CURIOSIDAD 2
' no borre la grilla y la parte contable, QUE NO SE USA en ventas,
' porque intuyo q la usare para reemplazar el ajuste COMPRAS, que jamas probe.
' --- NO BORRAR ---

Private CondicionTipoDoc As String
Private midDoc As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    Dim s As String
    CondicionTipoDoc = " ( tipodoc = '" & TipoDoc_AJ_CREDITO & "' or tipodoc = '" & TipoDoc_AJ_DEBITO & "' ) "
    
    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    uMotivo.ini "select descripcion from MotivosAjuste where codigo = ### ", "select Codigo, descripcion as [ Descripcion                               ] from MotivosAjuste where activo = 1", False

    s = "select * from FacturaVenta where activo = 1 and " & CondicionTipoDoc & " order by codigo"
    DE_abrir
    uMenu.init True, True, False, True, True, s, DataEnvironment1.Sistema, True
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    fraContable.Visible = False 'gEMPR_ConSistContable
    uDesde.ini (ucPrimerDiaAnio)
    uHasta.ini (ucHoy)
    
End Sub

Private Sub txttotal_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtTotal_LostFocus()
    txttotal = s2n(txttotal)
End Sub

Private Function TaTodo() As Boolean
    Dim i As Long
    
    If uCliente.codigo = 0 Then
        che "Falta Cliente"
        Exit Function
    End If
    If s2n(txttotal) = 0 Then
        che "Falta importe"
        Exit Function
    End If
    TaTodo = True
End Function

Private Function GrabaFactura() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    Dim Total As Double, codigo As Long, numero As Long, tipoiva, CUIT, iddoc As Long, asVta As New Asiento
    
    GrabaFactura = False
    
    tipoiva = s2n(obtenerDeSQL("select iva from clientes where codigo = " & uCliente.codigo))
    CUIT = sSinNull(obtenerDeSQL("select cuit from clientes where codigo = " & uCliente.codigo))
    Total = s2n(txttotal)
    
   
    '*** una transaccion aqui ..... *********************
    DE_BeginTrans
    
    
        codigo = nuevoCodigo("FacturaVenta", "codigo")
        numero = nuevoCodigo("FacturaVenta", "NroFactura", "TipoDoc = '" & StrTipoDoc() & "'")  'CondicionTipoDoc)
        iddoc = NuevoDocumento(StrTipoDoc, numero, 0, 0)
        
        asVta.nuevo "Aj " & uCliente.DESCRIPCION, uFecha.dtFecha, StrTipoDoc()
        
        Dim tiene_c, CUENTA_C As String
        tiene_c = obtenerDeSQL("select tiene_cuenta from clientes where codigo = " & uCliente.codigo)
        If tiene_c = 1 Then
            CUENTA_C = obtenerDeSQL("select cuenta from clientes where codigo = " & uCliente.codigo)
        Else
            CUENTA_C = CuentaParam(ID_Cuenta_V_DEUDxVENTAS)
        End If
        
        If optCredito Then
            asVta.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), s2n(Total), 0
            asVta.AcumularItem CUENTA_C, 0, s2n(txttotal)
        Else
            asVta.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), 0, s2n(txttotal)
            asVta.AcumularItem CUENTA_C, s2n(txttotal), 0
        End If
                
        DataEnvironment1.dbo_abmFacturaVenta "A", codigo, StrTipoDoc, numero, uFecha.dtFecha, uFecha.dtFecha, 0, 0, uCliente.codigo, uCliente.DESCRIPCION, "", CUIT, tipoiva, 0, 0, 0, 0, Total, Total, 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, uMotivo.codigo, 0, iddoc
                
        If asVta.Grabar(iddoc) = 0 Then
            DE_RollbackTrans
            ufa "Err al grabar asiento ", x2s(iddoc)
            GoTo fin
        End If

    
    DE_CommitTrans
    '*** una transaccion hasta aqui ..... *********************
    txtCodigo = codigo
    txtNumero = numero
    GrabaFactura = True
fin:
    Exit Function
UfaGraba:
    DE_RollbackTrans
    ufa "Err al grabar ", Me.Name & " - grabaAjusteventas - "
    Resume fin
End Function

Private Sub CargaDatos(codigo As Long)
    On Error Resume Next
    Dim tempo
    tempo = obtenerDeSQL("select codigo, tipodoc, nroFactura, fecha, cliente, total, MotivoAjuste, iddoc from facturaventa where codigo = " & codigo)
    txtCodigo = tempo(0)
    optCredito.Value = (Trim$(tempo(1)) = TipoDoc_AJ_CREDITO)
    optDebito.Value = (Trim$(tempo(1)) = TipoDoc_AJ_DEBITO)
    txtNumero = tempo(2)
    uFecha.dtFecha tempo(3)
    uCliente.codigo = tempo(4)
    txttotal = s2n(tempo(5))
    uMotivo.codigo = tempo(6)
    midDoc = s2n(tempo(7))
'''    'contable carga grilla
    uCuenta.codigo = ""
End Sub

Private Function StrTipoDoc()
    StrTipoDoc = IIf(optCredito, TipoDoc_AJ_CREDITO, TipoDoc_AJ_DEBITO)
End Function

Private Sub ImprimirAjuste()
 Dim sql As String
 If optCredito Then
    RptAjustesGrales.TxtAjuste = "AJUSTE CREDITO  A CLIENTE "
   Else
    RptAjustesGrales.TxtAjuste = "AJUSTE DEBITO   A CLIENTE"
 End If
    RptAjustesGrales.TxtCliProv = uCliente.DESCRIPCION
    RptAjustesGrales.TxtFecha = uFecha.dtFecha
    RptAjustesGrales.lblfecha = Date
    RptAjustesGrales.txttotal = txttotal
    RptAjustesGrales.txtNro = Format(txtNumero, "00000000")
    RptAjustesGrales.TxtTotalenLetras = enletras(txttotal)
    RptAjustesGrales.Restart
    If PREVIEW_IMPRESIONES Then
        RptAjustesGrales.Show
    Else
        RptAjustesGrales.PrintReport False
    End If
End Sub


'----------------------- MENU --------------------------------------
Private Sub uMenu_AceptarAlta()
    If TaTodo() Then
        If GrabaFactura() Then
            che "Ajuste Nro " & txtNumero & vbCrLf & "Grabado"
            ImprimirAjuste
            uMenu.AceptarOk
        End If
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    txtCodigo = ""
    uFecha.dtFecha Date
    uMotivo.clear
    uCliente.clear
    uCuenta.clear
    midDoc = 0
'    uCuenta.codigo = CuentaParam(ID_Cuenta_V_VENTAS)
'''    g.Borrar
'''    g.rows = TOT_ROWS
End Sub
Private Sub uMenu_Buscar()
    Dim resu
    resu = frmBuscar.MostrarSql("select NroFactura as [ Numero       ], TipoDoc as [ Tipo ], cliente as [ Cod Cliente ],codigo as [Cod] from FacturaVenta where activo = 1 and tipodoc = '" & StrTipoDoc() & "' and fecha " & ssBetween(uDesde.dtFecha, uHasta.dtFecha) & " order by NroFactura desc ")
    If resu > "" Then
        CargaDatos s2n(frmBuscar.resultado(4))
        uMenu.BuscarOK "codigo = " & frmBuscar.resultado(4)
    End If
End Sub
Private Sub uMenu_BuscarYa(que As Variant)
    Dim codi
    'codi = obtenerDeSQL("select codigo from FacturaVenta where activo = 1 and " & CondicionTipoDoc & " and NroFactura = " & que)
    codi = obtenerDeSQL("select codigo from FacturaVenta where activo = 1 and tipodoc = '" & StrTipoDoc() & "' and NroFactura = " & que)
    If Not IsEmpty(codi) Then
        CargaDatos (codi)
        uMenu.BuscarOK "codigo = " & codi
    End If
End Sub
Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaBORRA
    If s2n(txtCodigo) = 0 Then Exit Sub ' por las dudas
    
    DE_BeginTrans
    
        DataEnvironment1.dbo_abmFacturaVenta "B", s2n(txtCodigo), "", 0, 0, 0, 0, 0, 0, "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        grabaBitacora "B", s2n(txtCodigo), "FacturaVenta"
            
            'Baja Doc y asiento
            If Not BorroDocumento(midDoc) Or Not AsientoBaja_idDoc(midDoc) Then
                MsgBox "No se pudo borrar documento ni asiento." & Chr(13) & "(idDoc " & midDoc & ")", vbCritical
                'ufa "err al borrar documento", " middoc = " & midDoc
                DE_RollbackTrans
                GoTo fin:
            End If
    
    DE_CommitTrans
    
    che "Eliminado"
    uMenu.EliminarOK
fin:
    Exit Sub
UfaBORRA:
    DE_RollbackTrans
    ufa "err al eliminar", "Eliminar ajuste"
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
'''    fraContable.Enabled = sino
    fraAjuste.enabled = sino
End Sub

Private Sub uMenu_Imprimir()
    ImprimirAjuste
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub
Private Sub uMenu_SeMovio()
    CargaDatos uMenu.rs!codigo
End Sub
'----------------------- MENU --------------------------------------

'4/4/5
'
