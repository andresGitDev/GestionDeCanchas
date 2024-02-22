VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNotaDcheq 
   Caption         =   "Nota de debito por cheque rechazado"
   ClientHeight    =   7830
   ClientLeft      =   765
   ClientTop       =   450
   ClientWidth     =   10230
   Icon            =   "frmNotaDcheq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   10080
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   38
         Top             =   3540
         Width           =   945
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   7740
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3555
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   8445
         TabIndex        =   35
         Top             =   3150
         Width           =   960
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   1530
         TabIndex        =   18
         Top             =   1395
         Width           =   2355
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   320
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   105
         Width           =   1035
      End
      Begin VB.TextBox TxtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   2580
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   600
         Width           =   1245
      End
      Begin VB.ComboBox cmbTipoIva 
         Height          =   315
         Left            =   6900
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1020
         Width           =   2415
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1545
         TabIndex        =   13
         Top             =   1785
         Width           =   1455
      End
      Begin VB.TextBox txtPIVA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7815
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4845
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4110
         Width           =   945
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   10
         Top             =   4830
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   5160
         Width           =   945
      End
      Begin VB.TextBox txtIIBB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8460
         TabIndex        =   4
         Top             =   4515
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Gastos Administrativos"
         Height          =   585
         Left            =   8085
         TabIndex        =   2
         ToolTipText     =   "Genera Resibo a Cuenta."
         Top             =   5610
         Visible         =   0   'False
         Width           =   1050
      End
      Begin Gestion.ucFecha uFechaBuscaCheque 
         Height          =   330
         Left            =   8295
         TabIndex        =   3
         Top             =   2220
         Width           =   975
         _ExtentX        =   2990
         _ExtentY        =   582
         FechaInit       =   0
      End
      Begin Gestion.ucCoDe uCuenta 
         Height          =   315
         Left            =   1530
         TabIndex        =   5
         Top             =   2670
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uCheques 
         Height          =   315
         Left            =   1545
         TabIndex        =   7
         Top             =   2235
         Width           =   4605
         _ExtentX        =   6800
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1020
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2025
         Left            =   405
         TabIndex        =   9
         Top             =   3690
         Width           =   5775
         _cx             =   10186
         _cy             =   3572
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
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   5460
         TabIndex        =   19
         Top             =   525
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62390273
         CurrentDate     =   38126
      End
      Begin Gestion.ucCoDe ucCoDe1 
         Height          =   315
         Left            =   1530
         TabIndex        =   36
         Top             =   3135
         Width           =   5145
         _ExtentX        =   9234
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta gastos:"
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
         Index           =   3
         Left            =   135
         TabIndex        =   40
         Top             =   3150
         Width           =   1350
      End
      Begin VB.Label Label17 
         Caption         =   "Iva Gasto:"
         Height          =   255
         Index           =   3
         Left            =   6900
         TabIndex        =   39
         Top             =   3615
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos admin:"
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
         Index           =   2
         Left            =   6960
         TabIndex        =   34
         Top             =   3180
         Width           =   1515
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   780
         TabIndex        =   33
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label14 
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
         Left            =   300
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
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
         Left            =   720
         TabIndex        =   31
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label4 
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
         Left            =   4620
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   6900
         TabIndex        =   28
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Neto:"
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
         Left            =   960
         TabIndex        =   27
         Top             =   1845
         Width           =   570
      End
      Begin VB.Label Label18 
         Caption         =   "Neto:"
         Height          =   255
         Left            =   7230
         TabIndex        =   26
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Iva :"
         Height          =   255
         Index           =   0
         Left            =   7215
         TabIndex        =   25
         Top             =   4875
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCheque 
         Caption         =   "Nro Int Cheque"
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
         Height          =   315
         Left            =   165
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Total:"
         Height          =   255
         Index           =   1
         Left            =   7170
         TabIndex        =   23
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   22
         Top             =   2685
         Width           =   795
      End
      Begin VB.Label Label17 
         Caption         =   "IIBB:"
         Height          =   255
         Index           =   2
         Left            =   7245
         TabIndex        =   21
         Top             =   4575
         Width           =   375
      End
      Begin VB.Label lblChequeBusca 
         Caption         =   "Busca cheque desde:"
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
         Height          =   315
         Left            =   6135
         TabIndex        =   20
         Top             =   2250
         Width           =   2070
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Cancel          =   -1  'True
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   2672
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "frmNotaDcheq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '19/11/4

'Private WithEvents cliente As LiCodigo

'Public Enum FTipoNota
'    Tipo_NotaCredito
'    Tipo_NotaDebito
'    Tipo_NotaDebitoChRechazado
'End Enum
Private mTipoNota As FTipoNota

Private gDESC As Long
Private g As LiGrilla
Private Const CANT_RENGLONES = 5

Private Sub Command1_Click()
'    frmAsientoManual.Show
ImprimirComprobante s2n("61")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
'    rever
    Text1.Text = 0
End Sub

Private Sub Form_Load()
    Set g = New LiGrilla

'    CentrarMe Me
'    fra.Top = 0
'    fra.Left = 0
'    fra.Height = Me.ScaleHeight
'    fra.Width = Me.ScaleWidth
    
    dtFecha = Date

    comboSql cmbformapago, "select descripcion, codigo from formaspago where activo = 1"
    comboSql cmbTipoIva, "select descripcion, codigo from ivas"
    
    uCliente.ini "select descripcion from clientes where codigo = ### ", "select codigo as [ Codigo    ], descripcion as [ Cliemte                                              ] from clientes order by descripcion"
    g.init grilla
    gDESC = g.AddCol("Descripcion" & Space(90), "S")
    g.rows = CANT_RENGLONES
    
    'le saque el boton imprimir Hasta q haya busqueda
    If mTipoNota = Tipo_NotaDebitoChRechazado Then
        uMenu.init True, True, False, True, True
    Else
        uMenu.init False, True, False, False, False
    End If
    'uCheques.ini "select importe from cheques where NroInt = ### and estado = 'R' and activo = 1", "select NroInt, fecha, cliente, Importe from cheques where activo = 1 and estado = 'R' order by NroInt desc", False
    uCheques.ini "select importe from cheques where NroInt = ### and (estado = 'T' or estado = 'C' or estado = 'A' or estado = 'R' or estado = 'J') and activo = 1", "select NroInt, fecha, cliente, Importe, estado from cheques where activo = 1 and (estado = 'T' or estado = 'C' or estado='A' ) and fecha > " & uFechaBuscaCheque.ssFecha & " order by NroInt desc", False
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    ucCoDe1.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    txtPIVA = "1,21"
    Text2.Text = "1,21"
    txtIIBB = 0
End Sub


Private Function GrabaFactura() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    

    Dim asse As String ' assert
    Dim ac As Variant, i As Long, tota As Double, iddoc As Long
    Dim AsientoVenta As New Asiento, tipo As String
    Dim ContraAsientoVenta As New Asiento
    Dim TextoAsientoComprobante As String
    Dim Consulta As String
    Dim totND As Double
    Dim Neto As Double
    Dim cant As Double
    Dim Valor As Integer
    Dim cadena As String
    
    If Text1.Text = "" Then
        MsgBox "Debe ingresar el gasto administrativo."
        Exit Function
    End If
    
    GrabaFactura = False
    tota = s2n(txttotal)
 
 
'*** una transaccion aqui ..... *********************
    DE_BeginTrans
    
    asse = " Actualizo tabla parametros BS "
    
    
'''    AumentarParametroN CAMPO_BS_CodFactura_VENTA, s2n(txtcodigo)
''    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "A" Then
''        AumentarParametroD CAMPO_BS_FecFACTURA_A, dtfecha
''        AumentarParametroN CAMPO_BS_NroFACTURA_A, s2n(TxtNroFactura)
''    ElseIf TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then
''        AumentarParametroD CAMPO_BS_FecFACTURA_B, dtfecha
''        AumentarParametroN CAMPO_BS_NroFACTURA_B, s2n(TxtNroFactura)
''    Else
''        ufa "PrgErr: TipoDoc No reconocido", Me.Name ', Err
''    End If
    tipo = IIf(Left(txtTipoDoc, 2) = "NC", "N.Credito venta", "N.Debito venta")
    iddoc = NuevoDocumento(txtTipoDoc, TxtNroFactura, 0, 0)
    
    'DEBERIA ESTAR EN CABECERA ASIENTO
    TextoAsientoComprobante = tipo & TxtNroFactura
    
    
    AsientoVenta.nuevo tipo & " " & uCliente.DESCRIPCION, dtFecha, txtTipoDoc
     'ContraAsientoVenta.Nuevo tipo & " " & uCliente.DESCRIPCION, dtFecha, txtTipoDoc
    
    
    asse = "Graba Cabecera "
    ac = obtenerDeSQL("select provincia, cuit from clientes where codigo = " & uCliente.codigo)
    

    
    asse = "Graba detalle ND ch rech"

    If mTipoNota <> Tipo_NotaDebitoChRechazado Then
        If s2n(txtPIVA) = 0 Then
            Valor = 0
        Else
            Valor = 1
        End If
        DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtFecha, dtFecha, ComboCodigo(cmbformapago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n(txtMonto), s2n(txtPIVA) - Valor, s2n(txtIva), tota, tota, 0, 0, 0, UsuarioActual(), Date, 0, 0, s2n(txtIIBB), 0, 0, 0, 0, 0, 0, iddoc
        's2n(txtPIVA), s2n(TxtIVA)
    Else
        If s2n(txtPIVA) = 0 Then
            Valor = 0
        Else
            Valor = 1
        End If
        DataEnvironment1.dbo_abmFacturaVenta "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, dtFecha, dtFecha, ComboCodigo(cmbformapago), 0, uCliente.codigo, uCliente.DESCRIPCION, sSinNull(ac(0)), sSinNull(ac(1)), ComboCodigo(cmbTipoIva), 0, s2n(txtMonto), s2n(txtPIVA) - Valor, s2n(txtIva), tota, tota, 0, 0, 0, UsuarioActual(), Date, 0, 0, s2n(txtIIBB), 0, 0, 0, 1, 0, s2n(txttotal), iddoc
        's2n(txtPIVA), s2n(TxtIVA)
        
     
  
        
        DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, 1, True, uCheques.codigo, DatosCheque(), "", txtMonto, txtMonto, 0, 0, 0, 0, iddoc
        cadena = "update facturaventadetalle set _iva=0 where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura & " and producto='" & uCheques.codigo & "'"
        DataEnvironment1.Sistema.Execute cadena
        
'        If s2n(txttotal) > 0 Then
'            DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, 0, True, "0", "Gastos gravados  " & x2s(txttotal) & "", "", 0, 0, 0, 0, 0, 0, iddoc
'        End If
        
        
'        DataEnvironment1.Sistema.Execute _
'            "update cheques set estado = 'R' where nroint = " & uCheques.codigo
        
        If s2n(Text1.Text) > 0 Then
            Consulta = "Insert Into gastoadmin (nronota, tipodoc,fecha, importe,cuenta) " & _
                                "Values (" & TxtNroFactura & ", '" & txtTipoDoc & "', " & ssFecha(dtFecha) & ", '" & x2s(Text1.Text) & "', '" & ucCoDe1.codigo & "')"
                    
            DataEnvironment1.Sistema.Execute Consulta
        End If


        Dim d_q_cuenta
        d_q_cuenta = obtenerDeSQL("select dep_cuenta from cheques where nroint =" & uCheques.codigo)
        'If d_q_cuenta = 0 Then
        '    If MsgBox("El cheque " & uCheques.codigo & " no tiene cuenta de deposito." & Chr(13) & "Por favor indique una a continuacion, gracias.", vbInformation + vbYesNo) = vbYes Then
        '        d_q_cuenta = frmBuscar.MostrarSql("select c.codigo as [CODIGO], c.banco as [BANCO - Nº],b.descripcion as  [NOMBRE  ],c.numero as [CUENTA - Nº] from ctasbank c inner join bancosgrales b on c.banco=b.codigo where c.activo=1", , "Cuentas bancarias", " - ")
        '        DataEnvironment1.Sistema.Execute "update cheques set dep_cuenta = " & d_q_cuenta & " where nroint = " & uCheques.codigo
        '    Else
        '        d_q_cuenta = 0
        '        MsgBox "Debe ingresar una cuenta para este cheque.", vbInformation
        '        Exit Function
        '    End If
        'End If
            

            
        DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", d_q_cuenta, "R", "Rechazo de Cheque", dtFecha, "C" _
          , uCheques.codigo, s2n(txttotal), nuevoCodigo("movibanc", "movbanco"), iddoc, Date, UsuarioSistema!codigo
        
        ' con iddoc tengo un problema: son 3 operaciones: entra, quizas sale, rechazo.
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", uCheques.codigo, 0, "", 0, 0, "", 0, dtFecha, "R", 0, "", 0, Date, 0, 0, iddoc
    End If
        
    asse = "Graba detalle text"
    For i = 1 To g.rows - 1
        If Trim(g.tx(i, gDESC)) > "" Then
            If Trim(g.tx(i + 1, gDESC)) > "" Then
                Neto = 0
                cant = 0
            Else
                Neto = s2n(Text1.Text) 's2n(txtMonto)
                cant = 0 '1
            End If
            DataEnvironment1.dbo_abmFacturaVentaDetalle "A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, cant, True, i, Trim(g.tx(i, gDESC)), IIf(i = 1, uCuenta.codigo, ucCoDe1.codigo), Neto, Neto, 0, 0, 0, 0, iddoc
            If i = 2 Then
                cadena = "update facturaventadetalle set _iva=" & x2s((Text2.Text - 1) * 100) & " where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura & " and producto='2'"
                DataEnvironment1.Sistema.Execute cadena
            End If
        End If
    Next i
    
    
'    Consulta = "Insert Into gastoadmin (nronota, tipodoc,fecha, importe,cuenta) " & _
'                            "Values (" & TxtNroFactura & ", '" & txtTipoDoc & "', " & ssFecha(dtFecha) & ", '" & x2s(Text1.Text) & "', '" & ucCoDe1.codigo & "')"
'
'    DataEnvironment1.Sistema.Execute Consulta
    
    If mTipoNota = Tipo_NotaCredito Then 'credito
        AsientoVenta.AgregarItem uCuenta.codigo, s2n(txtneto), 0, TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(txtIva), 0, TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(txtIIBB), 0
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), 0, s2n(txttotal), TextoAsientoComprobante
    ElseIf mTipoNota = Tipo_NotaDebito Then
        totND = s2n(totND + txtMonto)
    
        AsientoVenta.AgregarItem uCuenta.codigo, 0, totND, TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva), TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(txtIIBB)
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), s2n(txttotal), 0, TextoAsientoComprobante
        
        'AsientoVenta.AgregarItem ucCoDe1.codigo, 0, Text1.Text, TextoAsientoComprobante
        'AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(Text3.Text), TextoAsientoComprobante
    ElseIf mTipoNota = Tipo_NotaDebitoChRechazado Then
        ' = s2n(TxtTotal)

        'If mTipoNota = Tipo_NotaDebitoChRechazado Then
            totND = s2n(totND + txtMonto)
        'Else
        '    totND = s2n(totND + txtMonto)
        'End If
        
        'aca esta el asiento manual para gabi
        Dim ec, cNota As String
        ec = obtenerDeSQL("select tiene_cuenta,cuenta from clientes where codigo=" & uCliente.codigo)
        If ec(0) = 1 Then
            cNota = Trim(ec(1))
        Else
            If uCuenta.codigo = "" Then
                MsgBox "El cliente no tiene cuenta personal, se necesita que indique una cuenta contable para el asiento.", vbCritical
                GoTo UfaGraba
            End If
            cNota = Trim(uCuenta.codigo)
        End If
        
        '************************************
        
        
                
        AsientoVenta.AgregarItem cNota, 0, totND, TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva), TextoAsientoComprobante
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(txtIIBB)
        AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), s2n(txttotal), 0, TextoAsientoComprobante
        If ucCoDe1.codigo > "" Then
            AsientoVenta.AgregarItem ucCoDe1.codigo, 0, Text1.Text, TextoAsientoComprobante
            AsientoVenta.AgregarItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(Text3.Text), TextoAsientoComprobante
        End If
    End If
    'If
    
    'ContraAsientoVenta.AgregarItem uCuenta.codigo, s2n(txtMonto), 0, TextoAsientoComprobante
    'ContraAsientoVenta.AgregarItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(txtMonto)
    'ContraAsientoVenta.Grabar iddoc
    
    AsientoVenta.Grabar iddoc
    '= 0 Then
'        DE_RollbackTrans
'        ufa "Err al grabar asiento", " " & iddoc
'        GoTo fin
'    End If
   
    DE_CommitTrans
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura = True
    GoTo fin
    
UfaGraba:
    DE_RollbackTrans
    ufa "Err al grabar ", Me.Name & " - grabaFactura() - " & asse ', Err
fin:

End Function




Private Sub uMenu_eliminar()
Dim cadena_del As String, movi_b, estado_c As String
    If MsgBox("Esta seguro de eliminar la nota de debito.", vbYesNo + vbCritical) = vbYes Then
        If txtCodigo <> 0 Then
            
            'inactivo la factura
            cadena_del = "update facturaventa set activo=0 where codigo=" & txtCodigo
            DataEnvironment1.Sistema.Execute cadena_del
            'delete del detalle
            cadena_del = "delete facturaventadetalle where codigofactura=" & txtCodigo
            DataEnvironment1.Sistema.Execute cadena_del
            'inactivo movimiento bancario
            movi_b = obtenerDeSQL("select movbanco from movibanc where operacion='R' and interno=" & uCheques.codigo)
            cadena_del = "update movibanc set activo=0 where movbanco=" & movi_b
            DataEnvironment1.Sistema.Execute cadena_del
            'revierto el estado del cheque
            estado_c = ""
            While estado_c <> "C" And estado_c <> "T" And estado_c <> "A"
                estado_c = InputBox("Elija el estado del cheque :" & Chr(13) & "C- Cartera" & Chr(13) & "T- Transferido" & Chr(13) & "A-Acreditado", "Estado del cheque", "C")
            Wend
            cadena_del = "update cheques set estado='" & estado_c & "' where nroint=" & uCheques.codigo
            DataEnvironment1.Sistema.Execute cadena_del
            'delete de gastos
            cadena_del = "delete gastoadmin where nronota=" & TxtNroFactura
            DataEnvironment1.Sistema.Execute cadena_del
            
        Else
            MsgBox "No se puede eliminar esta nota.", vbCritical
        End If
    End If
    MsgBox "Nota borrada.", vbInformation
    uMenu.EliminarOK
End Sub

Private Function DatosCheque() As String
    On Error Resume Next
    Dim tmp
    tmp = obtenerDeSQL("select nro, fecha, importe, descripcion from cheques inner join bancosGrales on cheques.banco_nro = BancosGrales.codigo where cheques.activo = 1 and cheques.NroInt = " & uCheques.codigo)
    'LO SIG LO HAGO POR EN TMP(0) SE GUARDA EL NUMERO DE CHEQUE, PERO NECESITO QUE TENGA EL TOTAL
'    tmp(0) = TxtTotal
    DatosCheque = "CHEQUE Nº " & tmp(0) & " " & Format(tmp(1), "dd/mm/yy") & " " & tmp(3) '& "  " & x2s(tmp(2))
End Function



Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Text1.Text <> 0 And Text1.Text <> "" Then
            grilla.TextMatrix(2, 0) = "GASTOS ADMINISTRATIVOS " & s2n(Text1) + s2n(Text3) & " $ (pesos)"
        End If
        Text3.SetFocus
    End If
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text <> "" Then
        rever
    Else
        Text1.Text = 0
    End If
    If Text1.Text <> 0 And Text1.Text <> "" Then
        grilla.TextMatrix(2, 0) = "GASTOS ADMINISTRATIVOS " & s2n(Text1) + s2n(Text3) & " $ (pesos)"
    End If
End Sub

Private Sub txtIIBB_LostFocus()
    rever
    RefrescoDeValores
End Sub

Private Sub txtiva_LostFocus()
    txtIva = s2n(txtIva)
    rever
End Sub

Private Sub RefrescoDeValores() '*********modificado el 16/5/7***raul
    Dim nrocheque As Variant
    Dim montoCheque As Variant
    Dim montoCHsinIva As Variant
    
    'EN UCHEQUES.DESCRIPCION SE COLOCABA EL MONTO PERO EL MONTO YA ESTABA EN OTRO FOCO
    'POR ESO UTILISE UCHEQUES.DESCRIPCION PARA QUE TENGA EL NUMERO DE CHEQUE
    'COMO EL CONTROL DEVUELTE EL MONTO Y SI LO MODIFICO CAMBIO A TODOS LO QUE LO UTILISEN
    'CON LO SIG HAGO COMO UN REFRESCO DE LO QUE QUIERO QUE TENGA
    If mTipoNota = Tipo_NotaDebitoChRechazado Then
        If (uCheques.codigo > 0) Then
            nrocheque = obtenerDeSQL("select nro from cheques where nroint = " & uCheques.codigo)
            uCheques.EditaDescripcion = True
            uCheques.DESCRIPCION = (nrocheque)
            montoCheque = s2n(obtenerDeSQL("select importe from cheques where nroint = " & uCheques.codigo), 2)
            montoCHsinIva = montoCheque '/ txtPIVA
            txtMonto = s2n(montoCHsinIva)
            txtneto = s2n(montoCHsinIva)
        End If
    End If
End Sub




Private Sub txtMonto_GotFocus()
    frmPintoFoco Me
    rever
    RefrescoDeValores
End Sub

Private Sub txtMonto_LostFocus()
    rever
    RefrescoDeValores
End Sub

Private Sub rever() '******modificado el 16/5/7***raul
    Dim valorCiva As Double
    Dim valorSiva As Double
    Dim mCh As Double
    
    If mTipoNota = Tipo_NotaDebitoChRechazado Then
        'COPIO EL TOTAL DEL CHEQUE, ESTE VALOR ES SACADO DE LA BD
        valorCiva = s2n(obtenerDeSQL("select importe from cheques where nroint = " & uCheques.codigo), 2) 'IIf(uCheques.codigo = 0, 0, s2n(txtMonto, 4)) ESTO ESTABA ANTES
        
        'VALIDO EL VALOR DEL CHEQUE PARA PONERLE LA COMA
        valorCiva = s2n(valorCiva, 2)
        
        'CALCULO EL VALOR DEL CHEQUE SIN IVA
        If txtPIVA = "" Then txtPIVA = 1
        If txtPIVA = 0 Then txtPIVA = 1
        valorSiva = valorCiva ' / s2n(txtPIVA, 2) ' permito mods manuales'If s2n(TxtIVA) = 0 Then txtIVA = valorCiva / s2n(txtPIVA, 4) ' permito mods manuales
        txtneto.Text = s2n(valorSiva)
        txtPIVA = "1,21"
        Text2.Text = "1,21"
        
        'CON ESTO SIEMPRE CALCULO EL VALOR DEL IVA
        txtIva = s2n(valorCiva - valorSiva)  'If s2n(TxtIVA) = 0 Then TxtIVA = valorCiva - valorSiva ESTO ESTABA ANTES
        Text3.Text = s2n(Text1.Text * (s2n(Text2.Text, 2) - 1))
        
        'ACA CALCULO EL TOTAL SIN EL VALOR DEL IVA
        'txtNeto = s2n(valorCiva - txtIva, 2) ' s2n(txtMonto, 4) + mCh
        
        'Y POR ULTIMO RECALCULO EL TOTAL SUMANDO EL VALOR SIN IVA + EL IVA + LOS INGRESOS BRUTOS
        txttotal = s2n(txtneto.Text) + s2n(txtIva.Text) + s2n(txtIIBB) + s2n(Text1.Text) + s2n(Text3.Text) 's2n(valorSiva) + s2n(txtIva, 2) + s2n(txtIIBB) 's2n(s2n(txtMonto, 2) + s2n(TxtIVA, 2) + mCh + s2n(txtIIBB)) ESTO ESTABA ANTES
    End If
    
    If mTipoNota = Tipo_NotaDebito Then
        valorCiva = s2n(txtMonto)
        txtMonto = s2n(txtMonto)
        txtPIVA = "1,21"
        valorSiva = s2n(valorCiva * txtPIVA) ' permito mods manuales
        txtIva = s2n(Text1.Text) * (txtPIVA - 1) 's2n(valorSiva - valorCiva)
        txttotal = s2n(txtneto) + s2n(txtIva) + s2n(txtIIBB) 's2n(valorSiva)   's2n(s2n(txtMonto) - s2n(txtIva))
        'txtNeto = s2n(valorSiva) - s2n(txtIva) + s2n(txtIIBB)
    End If
            
    If mTipoNota = Tipo_NotaCredito Then '******modificado el 16/5/7***raul
        valorSiva = s2n(txtMonto)
        txtMonto = s2n(txtMonto)
        txtPIVA = "1,21"
        valorCiva = s2n(valorSiva * txtPIVA)
        txtIva = s2n(valorCiva - valorSiva)
        txttotal = s2n(valorCiva) 's2n(s2n(txtMonto) - s2n(txtIva))
        txtneto = s2n(valorCiva) - s2n(txtIva) + s2n(txtIIBB)
    End If

    
    
    
End Sub


Private Sub txtPIVA_LostFocus()
    rever
End Sub

Private Sub uCheques_Buscar()
    Dim ss As String
'    ss = "select NroInt, Fecha, Cliente, Importe from cheques where estado = 'R' and activo = 1 "
'    ESTADO T=TRANSFERIDO  C=CARTERA
    ss = "select NroInt, fecha, cliente, Importe, estado from cheques where (activo = 1) and (estado = 'T' or estado = 'C' or estado='A' or estado='J')  and  fecha > " & uFechaBuscaCheque.ssFecha
    If uCliente.codigo > 0 Then ss = ss & " and cliente = " & uCliente.codigo
    ss = ss & " order by NroInt desc"
    uCheques.strSqlBuscar = ss
End Sub

Private Sub uCheques_cambio(codigo As Variant)
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    Dim tclie As Long
    Dim nrocheque As Long
    Dim montoCheque As Variant, BANCO_NRO As Integer, Banco As String


        'busca el numero de cliente que tenga ese numero de cheque
        tclie = obtenerDeSQL("select cliente from cheques where nroint = " & uCheques.codigo)
        'busca el numero de cheque(este es el numero con el que viene el cheque) correspondiente a ese cheque(ucheques.codigo es igual al numero interno que se le da)
        nrocheque = obtenerDeSQL("select nro from cheques where nroint = " & uCheques.codigo)
        'busca el monto de ese cheque
        
        montoCheque = obtenerDeSQL("select importe from cheques where nro = " & nrocheque)
        
        If uCheques.codigo > 0 Then
            BANCO_NRO = obtenerDeSQL("select banco_nro from cheques where nroint = " & uCheques.codigo)
            Banco = obtenerDeSQL("select descripcion from bancosgrales where codigo = " & BANCO_NRO)
            grilla.TextMatrix(1, 0) = "CHEQUE Nº " & nrocheque & " BANCO " & Banco
        End If

        If tclie > 0 Then
            If tclie <> uCliente.codigo Then
                If uCliente.codigo > 0 Then che "el cheque corresponde a otro cliente"
                uCliente.codigo = tclie
            End If
        End If
    'End If

    RefrescoDeValores
fin:
End Sub

Private Sub uCheques_LostFocus()
    rever
    'RefrescoDeValores
End Sub

Private Sub uCliente_cambio(codigo As Variant)
    On Error Resume Next
    Dim ac As Variant
    ac = obtenerDeSQL("select iva, FormaPago from clientes where codigo = " & codigo)
   
    cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, ac(1))
    cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, ac(0))
    'txtPIVA = s2n(obtenerDeSQL("select porcentaje from porcentajesiva where activo = 1 and iva =  " & ComboCodigo(cmbTipoIva)))
    'EL VALOR ANTERIOR LO SACABA DE LA BASE DE DATOS PERO NO ERA CORRECTO EL VALOR, ENTONCES SE LO PASO POR CODIGO
    txtPIVA = "1,21"
    txtIIBB = 0
    BuscoNroYTipo
    
    rever
End Sub

Private Function BuscoNroYTipo() As Boolean
    Dim tmpfec, letra As String
    BuscoNroYTipo = True
    Dim ss As String, andtipo As String, tmp
    letra = TipoFormVenta(ComboCodigo(cmbTipoIva))

    If letra = "B" Then
        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_B & "' or TipoDoc = '" & TipoDoc_FACTURA_B & "'  or TipoDoc = '" & TipoDoc_NDEBITO_B & "' ) "
    ElseIf letra = "A" Then
        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_A & "' or TipoDoc = '" & TipoDoc_FACTURA_A & "'  or TipoDoc = '" & TipoDoc_NDEBITO_A & "' ) "
    Else
        ufa "prg: No se encontro letra doc para Tipo Iva :" & cmbTipoIva, Me.Name ', 0
        BuscoNroYTipo = False
        Exit Function
    End If
    
    ' si vacio, lo lleno
    If Trim(TxtNroFactura) = "" Then TxtNroFactura = obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo) + 1

    BuscoNroYTipo = RevisaNroYFechaOk("FacturaVenta", "NroFactura", "Fecha", s2n(TxtNroFactura, 0), dtFecha, andtipo) ' Then Exit Sub
'    'existe factura
'    ss = "select codigo from facturaVenta where  NroFactura = " & TxtNroFactura & andtipo
'    tmp = obtenerDeSQL(ss)
'    If Not IsEmpty(tmp) Then
'        che "Factura Existente con el codigo interno  " & tmp
'        uMenu.SetFocus
'        BuscoNroYTipo = False
'        Exit Function
'    End If
'
'    Dim maxfac, maxfe, minfe, minfac
'    'fecha factura menor mas alta
'    ss = "select max(NroFactura) from FacturaVenta where activo = 1 and NroFactura < " & TxtNroFactura & andtipo
'    maxfac = obtenerDeSQL(ss)
'    ss = "select Fecha from FacturaVenta where NroFactura = " & maxfac & andtipo
'    maxfe = CDate(obtenerDeSQL(ss))
'    If dtFecha < maxfe Then
'        che " Fecha Factura " & dtFecha & " menor que de factura " & maxfac & " " & maxfe
'        BuscoNroYTipo = False
'        Exit Function
'    End If
'
'    'fecha factura mayor mas baja
'    ss = "select min(NroFactura) from FacturaVenta where activo = 1 and NroFactura > " & TxtNroFactura & andtipo
'    minfac = obtenerDeSQL(ss)
'    If IsNull(minfac) Then Exit Function
'
'    ss = "select Fecha from FacturaVenta where NroFactura = " & minfac & andtipo
'    minfe = (obtenerDeSQL(ss))
'
'    minfe = CDate(minfe)
'    If dtFecha > minfe Then
'        che " Fecha Factura " & dtFecha & " mayor que de factura " & minfac & " " & minfe
'        BuscoNroYTipo = False
'        Exit Function
'    End If
End Function



Public Sub mostrar(que As FTipoNota)
    mTipoNota = que
    Select Case que
    Case Tipo_NotaCredito
        Me.caption = "Emision Nota Credito"
        txtTipoDoc = "NC"
        'fraCheque.Visible = False
        'uCheques.Visible = False
        'lblCheque.Visible = False
        vercheqes False
        veoGastos False
    Case Tipo_NotaDebito
        Me.caption = "Emision Nota Debito"
        txtTipoDoc = "ND"
        'fraCheque.Visible = False
        'uCheques.Visible = False
        'lblCheque.Visible = False
        vercheqes False
        veoGastos False
    Case Tipo_NotaDebitoChRechazado
        Me.caption = "Emision Nota Debito Cheque Rechazado"
        txtTipoDoc = "ND"
        'fraCheque.Visible = True
        'uCheques.Visible = True
        'lblCheque.Visible = True
        vercheqes True
        veoGastos True
    End Select
    Me.Show
End Sub

Private Sub vercheqes(sino As Boolean)
        uCheques.Visible = sino
        lblCheque.Visible = sino
        lblChequeBusca.Visible = sino
        uFechaBuscaCheque.Visible = sino
End Sub

Private Function veoGastos(vg As Boolean)
    ucCoDe1.Visible = vg
    Text1.Visible = vg
    Text2.Visible = vg
    Text3.Visible = vg
    Label1(3).Visible = vg
    Label1(2).Visible = vg
    Label17(3).Visible = vg
    
    txtIva.Visible = Not vg
    txtPIVA.Visible = Not vg
    Label17(0).Visible = Not vg
End Function




'----------------------MENU -----------------------
Private Sub uMenu_AceptarAlta()
    
    Dim tmpfec As Date, tipoForm As String, andtipo  As String

    If TrabaIva(dtFecha.Value) Then
        MsgBox "La fecha del comprobante esta dentro de las fechas trabadas para emision," & Chr(13) & "verifiquelo con su contadora.", , "ATENCION"
        Exit Sub
    End If
    
    If Not PuedoVentas(dtFecha) Then
        'msg en funcion
        Exit Sub
    End If

    If s2n(txttotal) = 0 Then
        che "Monto no Valido"
        txtMonto.SetFocus
        Exit Sub
    End If
    If mTipoNota = Tipo_NotaDebitoChRechazado And uCheques.codigo = 0 Then
        che "Falta cheque"
        uCheques.SetFocus
        Exit Sub
    End If
    'If uCuenta.codigo = "" Then
    '    che "falta cuenta contable"
    '    Exit Sub
    'End If
    
'    'alta
    'txtcodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
    txtCodigo = nuevoCodigo("FacturaVenta", "codigo")
    If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "B" Then
        If mTipoNota = Tipo_NotaCredito Then
            txtTipoDoc = TipoDoc_NCREDITO_B
        Else
            txtTipoDoc = TipoDoc_NDEBITO_B
        End If
'        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_B & "' or TipoDoc = '" & TipoDoc_FACTURA_B & "'  or TipoDoc = '" & TipoDoc_NDEBITO_B & "' ) "
    Else
        If mTipoNota = Tipo_NotaCredito Then
            txtTipoDoc = TipoDoc_NCREDITO_A
        Else
            txtTipoDoc = TipoDoc_NDEBITO_A
        End If
'        andtipo = " ( TipoDoc = '" & TipoDoc_NCREDITO_A & "' or TipoDoc = '" & TipoDoc_FACTURA_A & "'  or TipoDoc = '" & TipoDoc_NDEBITO_A & "' ) "
    End If
    
    If Not BuscoNroYTipo() Then Exit Sub
'    If Not RevisaNroYFechaOk("FacturaVenta", "NroFactura", "Fecha", s2n(TxtNroFactura, 0), dtFecha, andtipo) Then Exit Sub
    
    If confirma("Nro de Nota: " & TxtNroFactura) Then
        If GrabaFactura() Then
            MsgBox "Nro Factura: " & TxtNroFactura & vbCrLf & "Grabado"
            If mTipoNota = Tipo_NotaCredito Then
                'ImprimirComprobante s2n(txtCodigo)
                ImprimirAMRAT (s2n(txtCodigo)), True, True
            Else
                'ImprimirComprobante s2n(txtCodigo)
                ImprimirAMRAT (s2n(txtCodigo)), True, True
            End If
            uMenu.AceptarOk
        End If
'    Else
'        txtCodigo = ""
'        TxtNroFactura = ""
    End If
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    FrmBorrarCbo Me
    uCliente.codigo = 0
    uCheques.clear
    uCuenta.clear
    ucCoDe1.clear
    g.Borrar
    g.rows = CANT_RENGLONES
    Text1.Text = 0
    txtneto.Text = ""
    txtIva.Text = ""
    Text3.Text = ""
    txtIIBB.Text = ""
    txttotal.Text = ""
End Sub
Private Sub uMenu_Buscar() 'similiar a la busqueda de factura venta
If ON_ERROR_HABILITADO Then On Error GoTo fin3
Dim re As Variant, WhereTipo As String, WhereFecha As String, a(10) As Variant, auxx As Double
    If mTipoNota = Tipo_NotaDebito Or mTipoNota = Tipo_NotaDebitoChRechazado Then
        WhereTipo = " (TipoDoc = 'NDA' or TipoDoc = 'NDB' or TipoDoc = 'NDE') " 'no esta verificado si funciona
    Else
        'WhereTipo = " (TipoDoc = 'NCA' or TipoDoc = 'NCB' or TipoDoc = 'NCE') "'no esta verificado si funciona
    End If
    
    WhereFecha = "fecha " & ssBetween(CDate("01/01/" & Year(Date)), Date)
    
    
    With frmBuscar

        re = .MostrarSql("select f.Codigo as Codigo, f.TipoDoc, f.NroFactura AS NroNota, f.Cliente, c.descripcion as [ Nombre                        ], f.Fecha as [Fecha ], f.Activo from facturaventa as f left join clientes as c on c.codigo = f.cliente where " & WhereTipo & " and " & WhereFecha & " order by f.NroFactura desc ", , "Nota Debito", " <-> ", "Activo", "Anulada", False)
        
        
        If re > "" Then
            If .resultado(7) = "Activo" Then
                'esto esta hecho asi nomas si anda bien joya sino arreglenlo
                txtCodigo = .resultado(1)
                txtTipoDoc = Trim(.resultado(2))
                TxtNroFactura = .resultado(3)
                uCliente.codigo = .resultado(4)
                ' = .resultado(5) 'descp cliente

                Text2 = obtenerDeSQL("select porcentajeiva from facturaventa where codigo=" & txtCodigo) 'ivaPorc
                Text3 = obtenerDeSQL("select iva from facturaventa where codigo=" & txtCodigo) 'ivaValor
                txttotal = obtenerDeSQL("select total from facturaventa where codigo=" & txtCodigo)
                txtneto = obtenerDeSQL("select neto from facturaventa where codigo=" & txtCodigo)
                auxx = txttotal - txtneto
                Text1 = s2n(auxx / 1.21)
                Text3 = s2n(auxx - Text1)
                dtFecha = .resultado(6)
                LlenarGrilla grilla, "select cast(producto as int) as producto, Descripcion from facturaventadetalle where codigofactura =" & txtCodigo & " and  producto<>'1' ORDER BY Producto DESC", True
                If grilla.rows > 1 Then grilla.ColWidth(0) = 0
                If grilla.rows > 1 Then grilla.ColWidth(1) = 6000
                If grilla.rows > 1 Then uCheques.codigo = Trim(grilla.TextMatrix(1, 0))
                uCuenta.codigo = obtenerDeSQL("select formula from facturaventadetalle where producto=1 and codigofactura=" & txtCodigo)
                ucCoDe1.codigo = obtenerDeSQL("select formula from facturaventadetalle where producto=2 and codigofactura=" & txtCodigo)
                uMenu.BuscarOK
            Else
                MsgBox "No se puede visualizar un documento anulado.", vbInformation
            End If
        End If
    End With
Exit Sub
fin3:
MsgBox "Se producio un error al buscar.", vbCritical
End Sub



Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fra.enabled = sino
End Sub
Private Sub uMenu_Imprimir()
    If s2n(txtCodigo) = 0 Then Exit Sub
    'ImprimirComprobante (s2n(txtCodigo))
    ImprimirAMRAT (s2n(txtCodigo)), True, True
'    If MsgBox("Desea imprimir el triplicado?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
'        ImprimirAMRAT (s2n(txtCodigo)), True
'    End If
End Sub
Private Sub uMenu_Nuevo()
    'dtFecha.SetFocus
    'uCliente.SetFocus
    Text1.Text = 0
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
'17/12/4
'   foco en cliente
'28/2/5
'   null
'18/4/5
'   cheque rechazado
'20/5/5
'   nro factura manual, verifica fechas y correlatividad
'26/5/5
'   codigo FV desde tabla FacturaVenta no BS
'31/5/5
'   fix montos neto, tot,  (y cheque)
'   iva editable
'   pone nro al principio
'

