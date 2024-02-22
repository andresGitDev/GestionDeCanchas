VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLiberarCHtercero 
   Caption         =   "Salida de Cheques de terceros"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   Icon            =   "frmLiberarCHtercero.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTodo 
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   45
      TabIndex        =   12
      Top             =   60
      Width           =   8295
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         Left            =   3840
         TabIndex        =   31
         Text            =   "Ejercicio"
         Top             =   1920
         Width           =   990
      End
      Begin VB.Frame fraX 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   75
         TabIndex        =   19
         Top             =   2580
         Width           =   7815
      End
      Begin VB.TextBox txtint 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   495
         Width           =   1935
      End
      Begin VB.TextBox txtnumcheque 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   17
         Tag             =   "2"
         Top             =   1215
         Width           =   5775
      End
      Begin VB.TextBox txtimporte 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1935
         Width           =   1335
      End
      Begin VB.TextBox txtconcepto 
         Height          =   285
         Left            =   2055
         TabIndex        =   15
         Tag             =   "2"
         Top             =   1575
         Width           =   5775
      End
      Begin VB.TextBox txtbanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "2"
         Top             =   855
         Width           =   5775
      End
      Begin Gestion.ucCoDe uChe 
         Height          =   285
         Left            =   2025
         TabIndex        =   13
         Top             =   150
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   503
         CodigoWidth     =   1000
      End
      Begin MSComCtl2.DTPicker fechaoper 
         Height          =   255
         Left            =   6525
         TabIndex        =   20
         Top             =   2295
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   450
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   38052
      End
      Begin MSComCtl2.DTPicker fechacheque 
         Height          =   255
         Left            =   2025
         TabIndex        =   21
         Top             =   2295
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   61538305
         CurrentDate     =   38052
      End
      Begin VB.Label Label34 
         Caption         =   "Ejercicio"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label lbliddoc 
         Caption         =   "Label1"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   120
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   3060
         Left            =   30
         Top             =   15
         Width           =   8175
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Cheque:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   2295
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Interno Cheque:"
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
         Left            =   135
         TabIndex        =   27
         Top             =   495
         Width           =   1695
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
         Left            =   120
         TabIndex        =   26
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Nº Cheque:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1215
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto/Resp.:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1575
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Banco:"
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
         Left            =   105
         TabIndex        =   23
         Top             =   855
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Operación:"
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
         Left            =   4875
         TabIndex        =   22
         Top             =   2295
         Width           =   1605
      End
   End
   Begin VB.Frame fraGrilla 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   15
      TabIndex        =   0
      Top             =   3240
      Width           =   8370
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6075
         TabIndex        =   5
         Tag             =   "8"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txtvalor 
         Height          =   285
         Left            =   2055
         TabIndex        =   4
         Top             =   825
         Width           =   1335
      End
      Begin VB.TextBox txtconc 
         Height          =   285
         Left            =   2055
         TabIndex        =   3
         Top             =   465
         Width           =   4455
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
         Left            =   6015
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1215
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
         Left            =   6015
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1695
         Width           =   975
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   1860
         Left            =   375
         TabIndex        =   6
         Top             =   1170
         Width           =   5580
         _cx             =   9842
         _cy             =   3281
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
         Height          =   300
         Left            =   2055
         TabIndex        =   7
         Top             =   135
         Width           =   5865
         _ExtentX        =   10213
         _ExtentY        =   529
         CodigoWidth     =   1000
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
         Left            =   6225
         TabIndex        =   11
         Top             =   2340
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
         Left            =   435
         TabIndex        =   10
         Top             =   510
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
         Left            =   450
         TabIndex        =   9
         Top             =   855
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
         Left            =   420
         TabIndex        =   8
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00400000&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   3150
         Left            =   75
         Top             =   15
         Width           =   8160
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1650
      Left            =   0
      TabIndex        =   29
      Top             =   6540
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   2910
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "frmLiberarCHtercero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private midDoc As Long


Private Sub cmbeliminofila_Click()
    If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
        If grilla.Row > 0 Then grilla.RemoveItem (grilla.Row)
    End If
    revisoTotalGrilla
End Sub

Private Sub cmdcargar_Click()
    
    Dim totalgrilla As Double

    If s2n(txtvalor) = 0 Then
        che "falta ingresar valor"
        txtvalor.SetFocus
        Exit Sub
    End If
    If (s2n(txtvalor) <= s2n(txtimporte)) And (s2n(txtvalor) + s2n(txttotal) > s2n(txtimporte)) Then
        MsgBox "El valor a ingresar no puede superar al importe original"
        txtvalor.SetFocus
        Exit Sub
    End If
    If (s2n(txttotal) + s2n(txtvalor)) > s2n(txtimporte) Then
        che "Con este valor el importe total serìa superado"
        Exit Sub
    End If
    
    'todo ok
    Cargogrilla
    Limpiotextosgrilla
    'If txtcodcuenta.enabled Then txtcodcuenta.SetFocus

End Sub

Private Sub Form_Load()
    uCuenta.ini uCuentaIni1Imput, uCuentaIni2Imput, True
 
    uChe.ini " select b.descripcion  from Cheques c inner join bancosgrales b on c.banco_nro=b.codigo  where c.activo = 1 and c.estado = 'C' and c.nroint = ### ", _
        " select c.nroint [Interno],c.nro [Nro de cheque],b.descripcion [Descripcion                        ] ,c.importe [Importe],c.Fecha from cheques c inner join bancosgrales b on c.banco_nro=b.codigo where c.activo=1 and c.estado='C' order by nroint"
    
    uMenu.init True, True, False, False, True
        
    fechacheque.Value = Date
    fechaoper.Value = Date
    InicioGrilla
    
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

Private Sub txtconc_Click()
    If txtconcepto <> "" Then
        txtconc.Text = txtconcepto.Text
    End If
End Sub

Private Sub txtvalor_Click()
    If txtimporte <> "" Then
        txtvalor.Text = txtimporte.Text
    End If
End Sub

Private Sub uChe_cambio(codigo As Variant)
    If uChe.codigo = "0" Then
    Else
        txtint.Text = uChe.codigo
        txtbanco.Text = uChe.DESCRIPCION
        txtnumcheque.Text = frmBuscar.resultado(2)
        txtimporte.Text = obtenerDeSQL("select importe from cheques where nroint=" & Trim(txtint))
        fechacheque.Value = obtenerDeSQL("select fecha from cheques where nroint=" & Trim(txtint))
    End If
End Sub

Private Sub uMenu_Aceptar()
If ON_ERROR_HABILITADO Then On Error GoTo UFAaceptar
    Dim vdias
    Dim rs As New ADODB.Recordset
    Dim maximobanc1 As Long, maxcheque As Long, maximocaja As Long, x As Long, valcartera As String, depcuenta As Long, cuentaconcaja As Long
    Dim cuentacon As String, Importe As Double, asse As String
    Dim asiCh As New Asiento, aConcepto As String, cueBan As String, nroInterno As Long
    
     If grilla.rows < 1 Then
        MsgBox "No ha cargado los datos en la grilla.", , "ATENCION"
        Exit Sub
     End If
'    If optrechazar = True Then
        aConcepto = Trim(txtconcepto) '"Rechazo de cheques "
'    End If
    
    asiCh.nuevo aConcepto, fechaoper, "CH3"

    '***************************************
    DE_BeginTrans
    
    midDoc = NuevoDocumento("ch3", nuevoCodigo("registrodocumentos", "nrodoc", "tipodoc = 'ch3'"), 0, 0)
    Importe = s2n(txtimporte.Text)
    'ASIENTO
    asiCh.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, Importe
    
    For x = 1 To grilla.rows - 1
        If grilla.TextMatrix(x, 0) <> "" Then
            
            maximobanc1 = nuevoCodigo("MoviBanc", "MovBanco")
            
            rs.Open "select dep_cuenta from Cheques where nroint = " & Val(txtint.Text) & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                depcuenta = rs!dep_cuenta
            End If
            rs.Close
            Set rs = Nothing
                
            
'            If optrechazar = True Then
                
                Importe = s2n(grilla.TextMatrix(x, 3))
                cueBan = Trim(grilla.TextMatrix(x, 0)) 'sSinNull(obtenerDeSQL("select cuenta_con from ctasbank where codigo = '" & depcuenta & "' "))
                nroInterno = uChe.codigo
                asse = "optRech chMoviBanc"
                
                'ASIENTO
                asiCh.AcumularItem cueBan, Importe, 0
                
'            End If
            
            maximocaja = nuevoCodigo("MoviCaja", "Movimiento")
            
            valcartera = CuentaParam(ID_Cuenta_M_CH_CARTERA)
            
                                                           
'            If optrechazar = True Then
                asse = "optrech ch3 "
                'S=salida de cheque de terceros
                DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", nroInterno, 0, "", 0, depcuenta, "", 0, fechaoper, "S", 0, "", 0, fechaoper, 0, 0, midDoc
'            End If
            
        End If
vencido:
    Next
    
    DataEnvironment1.dbo_GRABARBITACORA Trim(txtint.Text), "ChequeSalida", UsuarioSistema!codigo, Date, Time, "M"
    
    If asiCh.CantItems > 0 Then
        If asiCh.Grabar(midDoc, , leerEjercicioId(cboEjercicio)) = 0 Then
            DE_RollbackTrans
            ufa "Err al grabar asiento ", Me.Name & " - " '& sAssert
            Exit Sub
        End If
    End If
    
    DE_CommitTrans
    '***************************************
        
    
    MsgBox "Operación Realizada con éxito", vbOKOnly
    uMenu.AceptarOk
    limpio


fin:
    Set rs = Nothing
    Exit Sub
UFAaceptar:
    DE_RollbackTrans
    ufa "err al grabar ", "aceptar " & asse
    Resume fin
    
End Sub

Private Sub uMenu_Buscar()
    Dim resu
    
    resu = frmBuscar.MostrarSql("select distinct(Nroint),Nro,Fecha_oper [Fecha   ],Importe,iddoc_egreso [iddoc] from cheques c inner join bitacora b on b.codigo=c.nroint where estado='S' and tabla='ChequeSalida'")
    If resu <> "" Then
        uChe.codigo = frmBuscar.resultado(1)
        txtint = frmBuscar.resultado(1)
        txtbanco.Text = obtenerDeSQL("select b.descripcion  from Cheques c inner join bancosgrales b on c.banco_nro=b.codigo  where c.activo = 1 and c.estado = 'S' and c.nroint =" & frmBuscar.resultado(1))
        txtnumcheque.Text = frmBuscar.resultado(2)
        txtimporte.Text = obtenerDeSQL("select importe from cheques where nroint=" & Trim(txtint))
        fechacheque.Value = obtenerDeSQL("select fecha from cheques where nroint=" & Trim(txtint))
        txtconcepto.Text = obtenerDeSQL("select concepto from asientos where iddoc=" & frmBuscar.resultado(5))
        lblIdDoc = frmBuscar.resultado(5)
        uMenu.BuscarOK
    End If
    
End Sub

Private Sub uMenu_Cancelar()
    limpio
End Sub

Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    Dim tempCH
    Dim asse As String, nInterno As Long
    Dim depcuenta As Long
    Dim rs As New ADODB.Recordset
    
    nInterno = s2n(txtint.Text)
    
    tempCH = obtenerDeSQL("select estado     from cheques  where NroInt = " & nInterno)
        
    If IsEmpty(tempCH) Then
        ufa "prg: err al eliminar. no se encuentra cheque", "eliminar " & nInterno
        Exit Sub
    ElseIf tempCH <> "S" Then
        che "Cheque no esta figura como entregado"
        Exit Sub
    End If
    
    rs.Open "select dep_cuenta from Cheques where nroint = " & Val(txtint.Text) & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        depcuenta = rs!dep_cuenta
    End If
    rs.Close
    Set rs = Nothing
            
    DE_BeginTrans
        
        BorroDocumento lblIdDoc
        
        asse = "cheques " 'cheques
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "M", nInterno, 0, "", 0, depcuenta, "", 0, fechaoper, "C", 0, "", 0, fechaoper, 0, 0, lblIdDoc
               
        
        'asse = "movcjdet" 'movcjdet detalleMovCajas
'''''        DataEnvironment1.dbo_INGCHEQUEDETALLE "B", s2n(tempMC), 0, 0, 0, "", "", 0
        
        asse = "bitacora"
        DataEnvironment1.dbo_GRABARBITACORA Trim(txtint.Text), "ChequeSalida", UsuarioSistema!codigo, Date, Time, "B"
        
        AsientoBaja_idDoc lblIdDoc
        
    DE_CommitTrans
    
    limpio
    MsgBox "Se ha eliminado con exito.", , "ATENCION"
    uMenu.EliminarOK
    
fin:
    Exit Sub
UFAelim:
    ufa "prg: err en la eliminacion", "cheques 3ros salida " & asse
    DE_RollbackTrans
    Resume fin

End Sub

Private Sub uMenu_Nuevo()
    limpio
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub

Private Sub limpio()
    uChe.clear
    txtint.Text = ""
    txtbanco.Text = ""
    txtnumcheque.Text = ""
    txtconcepto.Text = ""
    txtimporte.Text = ""
    fechacheque.Value = Date
    fechaoper.Value = Date
    uCuenta.clear
    txtconc.Text = ""
    txtvalor.Text = ""
    txttotal.Text = ""
    InicioGrilla
End Sub

Sub InicioGrilla()
    grilla.clear
    grilla.rows = 1
    grilla.cols = 4
    'grilla.ColWidth(1) = 1700
    grilla.TextMatrix(0, 0) = "Cuenta"
    grilla.TextMatrix(0, 1) = "Descripción"
    grilla.TextMatrix(0, 2) = "Concepto"
    grilla.TextMatrix(0, 3) = "Importe"
End Sub

Private Sub Cargogrilla()
    grilla.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor

    If txttotal <> "" Then
        txttotal = s2n(txttotal) + s2n(txtvalor)
    Else
        txttotal = s2n(txtvalor)
    End If
    If txttotal = txtimporte Then
        MsgBox "El detalle ha sido completado"
    End If
End Sub

Private Sub Limpiotextosgrilla()
    txtconc = ""
    txtvalor = ""
    uCuenta.clear
End Sub

Private Sub revisoTotalGrilla()
    Dim i As Long, tot As Double
    With grilla
        For i = 1 To .rows - 1
            tot = tot + s2n(.TextMatrix(i, 3))
        Next i
    End With
    txttotal = tot
End Sub
