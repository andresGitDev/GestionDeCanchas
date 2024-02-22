VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLibracionCheques 
   Caption         =   "Libración de Cheques"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   Icon            =   "FrmLibracionCheques.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGrilla 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   15
      TabIndex        =   36
      Top             =   4260
      Width           =   8370
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
         TabIndex        =   21
         Top             =   1695
         Width           =   975
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
         TabIndex        =   20
         Top             =   1215
         Width           =   975
      End
      Begin VB.TextBox txtconc 
         Height          =   285
         Left            =   2055
         TabIndex        =   18
         Top             =   465
         Width           =   4455
      End
      Begin VB.TextBox txtvalor 
         Height          =   285
         Left            =   2055
         TabIndex        =   19
         Top             =   825
         Width           =   1335
      End
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6075
         TabIndex        =   37
         Tag             =   "8"
         Top             =   2700
         Width           =   1095
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   1860
         Left            =   375
         TabIndex        =   38
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
         TabIndex        =   17
         Top             =   135
         Width           =   5865
         _extentx        =   10213
         _extenty        =   529
         codigowidth     =   1000
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00400000&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   3150
         Left            =   75
         Top             =   30
         Width           =   8160
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
         TabIndex        =   42
         Top             =   135
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
         TabIndex        =   41
         Top             =   855
         Width           =   1215
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
         TabIndex        =   40
         Top             =   510
         Width           =   1215
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
         TabIndex        =   39
         Top             =   2340
         Width           =   735
      End
   End
   Begin VB.Frame fraTodo 
      BorderStyle     =   0  'None
      Height          =   4155
      Left            =   45
      TabIndex        =   22
      Top             =   75
      Width           =   8295
      Begin VB.ComboBox cboEjercicio 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Text            =   "Ejercicio"
         Top             =   120
         Width           =   990
      End
      Begin Gestion.ucCoDe uChe 
         Height          =   285
         Left            =   2025
         TabIndex        =   5
         Top             =   540
         Width           =   5760
         _extentx        =   10160
         _extenty        =   503
         codigowidth     =   1000
      End
      Begin VB.Frame fraOpcion 
         Height          =   420
         Left            =   2055
         TabIndex        =   1
         Top             =   45
         Width           =   4725
         Begin VB.OptionButton optX 
            Caption         =   "Otros"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3555
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   150
            Width           =   960
         End
         Begin VB.OptionButton optX 
            Caption         =   "Proveedor"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   1905
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   150
            Width           =   1575
         End
         Begin VB.OptionButton optX 
            Caption         =   "Caja"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   2
            Tag             =   "0"
            Top             =   135
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.TextBox txttipocta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Tag             =   "2"
         Top             =   1965
         Width           =   5775
      End
      Begin VB.TextBox txtbanco 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Tag             =   "2"
         Top             =   1605
         Width           =   5775
      End
      Begin VB.TextBox txtconcepto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2055
         TabIndex        =   12
         Tag             =   "2"
         Top             =   2685
         Width           =   5775
      End
      Begin VB.TextBox txtimporte 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   3045
         Width           =   1350
      End
      Begin VB.TextBox txtnumcheque 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Tag             =   "2"
         Top             =   2325
         Width           =   5775
      End
      Begin VB.TextBox txtint 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   885
         Width           =   1935
      End
      Begin VB.TextBox txtcodcuenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Tag             =   "2"
         Top             =   1245
         Width           =   1935
      End
      Begin VB.TextBox txtnumcuenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Tag             =   "2"
         Top             =   1245
         Width           =   2535
      End
      Begin VB.Frame fraX 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   75
         TabIndex        =   23
         Top             =   3690
         Width           =   7815
         Begin Gestion.ucCoDe uX 
            Height          =   300
            Left            =   1950
            TabIndex        =   16
            Top             =   30
            Width           =   5880
            _extentx        =   10372
            _extenty        =   529
            codigowidth     =   1000
         End
         Begin VB.Label lblX 
            Caption         =   "Caja:"
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
            Top             =   30
            Width           =   975
         End
      End
      Begin MSComCtl2.DTPicker fechaoper 
         Height          =   255
         Left            =   5715
         TabIndex        =   15
         Top             =   3405
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   77070337
         CurrentDate     =   38052
      End
      Begin MSComCtl2.DTPicker fechacheque 
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   3405
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   77070337
         CurrentDate     =   38052
      End
      Begin VB.Label Label34 
         Caption         =   "Ejercicio"
         Height          =   255
         Left            =   1200
         TabIndex        =   44
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Libracion :"
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
         Left            =   3915
         TabIndex        =   35
         Top             =   3405
         Width           =   1695
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
         Left            =   120
         TabIndex        =   34
         Top             =   1605
         Width           =   975
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
         Left            =   120
         TabIndex        =   33
         Top             =   1965
         Width           =   975
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
         TabIndex        =   32
         Top             =   2685
         Width           =   1575
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
         TabIndex        =   31
         Top             =   2325
         Width           =   1095
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
         TabIndex        =   30
         Top             =   3045
         Width           =   975
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
         TabIndex        =   29
         Top             =   885
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Cód. Cuenta:"
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
         Top             =   1245
         Width           =   1215
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
         TabIndex        =   27
         Top             =   3405
         Width           =   1455
      End
      Begin VB.Label lblcartel 
         BackStyle       =   0  'Transparent
         Caption         =   "Libración de Cheque contra Otros"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1995
         TabIndex        =   26
         Top             =   3690
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Label Label13 
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
         Left            =   4200
         TabIndex        =   25
         Top             =   1245
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   4110
         Left            =   30
         Top             =   15
         Width           =   8175
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1530
      Left            =   0
      TabIndex        =   0
      Top             =   7515
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   2699
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
End
Attribute VB_Name = "FrmLibracionCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit '16/9/4  ************************* optlibracheques

'/* --------------------------------------
' una LIB ahora tiene registro iddoc, PERO no agregue iddoc a viejas tablas
' Se referencia iddoc a traves de movimiento de movicaja
'
' afecta
'
'   CHQ_COMP    mod
'   MOVICAJA    alta
'   MOVIBANC    alta
'   asientos    alta
'
'*/


' ************   PASAR A PARAMETRIZACION   ******************
Private Const BS_CHQCOMP_USAR_NROINT = False
' ************   PASAR A PARAMETRIZACION   ******************


Private gCUEN As Long
Private gDESC As Long
Private gCONC As Long
Private gIMPO As Long


Private mViejo As enumChequeContraQue
Private Enum enumChequeContraQue
    chequeCaja = 0
    chequeProv = 1
    chequeOtro = 2
End Enum

Private Sub LimpioImputacion()
    uCuenta.clear
    uX.clear
    txtconc = ""
    txtvalor = ""
End Sub

Private Sub LimpioControles()
'    optcaja.Value = True
'    optprov.Value = False
'    optotros.Value = False
    optX(chequeCaja).Value = True
    
    uX.clear
    uCuenta.clear
    uChe.clear
    txtint = ""
    txtcodcuenta = ""
    txtbanco = ""
    txttipocta = ""
    txtnumcheque = ""
    txtconcepto = ""
    txtimporte = ""
    fechacheque = Date
    fechaoper = Date
'    txtcod = ""
'    cargar = ""
'    txtcuentabanc = ""
'    txtdesc = ""
    txttotal = "0"
    txtnumcuenta = ""
'    Ope = ""
End Sub

Private Sub HabilitoControles(habilito As Boolean)
    
    optX(0).enabled = habilito
    optX(1).enabled = habilito
    optX(2).enabled = habilito
    
    txtint.enabled = habilito
'    txtcodcuenta.Enabled = habilito
'    txtbanco.Enabled = habilito
'    txttipocta.Enabled = habilito
'    txtnumcheque.Enabled = habilito
    txtconcepto.enabled = habilito
    txtimporte.enabled = habilito
    fechacheque.enabled = habilito
    fechaoper.enabled = habilito
'    txtcod.Enabled = habilito
'    txtdesc.Enabled = habilito
'    cmbcheque.Enabled = habilito
'    cmbcambio.Enabled = habilito
    uX.enabled = habilito
End Sub


'Private Sub cmbcheque_Click()
'    Dim ss As String, re As String
'    ss = "SELECT ch.CODIGO as [CodigoInterno], ba.descripcion as [ Banco            ], ch.NRO as [ Numero        ] FROM CHQ_COMP ch LEFT OUTER JOIN BancosGrales ba ON ba.codigo = ch.Banco WHERE (ch.ESTADO = 'C')"
'    re = frmBuscar.MostrarSql(ss)
'''''''''        rs.Open "select * from " + Tabla & " where activo = 1 and estado = 'C' order by " & order
'    If re > "" Then
'        CargaCheque s2n(re)
'    End If
'End Sub

Private Sub revisoTotalGrilla()
    Dim i As Long, tot As Double
    With grilla
        For i = 1 To .rows - 1
            tot = tot + s2n(.TextMatrix(i, 3))
        Next i
    End With
    txttotal = tot
End Sub

Private Sub cmbeliminofila_Click()
    If grilla.TextMatrix(grilla.Row, grilla.Col) <> "" Then
        If grilla.Row > 0 Then grilla.RemoveItem (grilla.Row)
    End If
    revisoTotalGrilla
End Sub

Private Sub imprimoPantalla(interno As Long)
    Dim rsNroPago As New ADODB.Recordset
    Dim enletras, tipo, ContraQue
    Dim sql As String

        If optX(chequeCaja) Then
            tipo = "C"
            ContraQue = "CAJA "
            GoTo ImprimirCaja
        ElseIf optX(chequeOtro) Then
            tipo = "O"
            ContraQue = "OTROS"
            RptLibracionChequeProv.ContraQue = ContraQue
            RptLibracionChequeProv.txtProveedor = ""
            RptLibracionChequeProv.TxtCodProv = ""
            RptLibracionChequeProv.lblproveedor = ""
            GoTo Imprimir
        Else  'optX(chequeProv)
            tipo = "P"
            ContraQue = "PROVEEDOR"
            RptLibracionChequeProv.ContraQue = ContraQue
            RptLibracionChequeProv.txtProveedor = uX.DESCRIPCION
            RptLibracionChequeProv.TxtCodProv = uX.codigo
            GoTo Imprimir
        End If
    
ImprimirCaja:
        enletras = NroEnLetras(s2n(txtimporte))
        'DataEnvironment1.LibracionCheques
        rptLibracionCheques.Sections("Encabezado").Controls("label7").caption = fechaoper
        rptLibracionCheques.Sections("Encabezado").Controls("lblnumbanco").caption = txtbanco
        rptLibracionCheques.Sections("Encabezado").Controls("lblcuenta").caption = txtcodcuenta
        rptLibracionCheques.Sections("Encabezado").Controls("lblnumerocta").caption = txtnumcuenta
        rptLibracionCheques.Sections("Encabezado").Controls("lbldesc").caption = ContraQue

        rptLibracionCheques.Sections("Medio").Controls("lblinterno").caption = txtint
        rptLibracionCheques.Sections("Medio").Controls("lblnumero").caption = txtnumcheque
        rptLibracionCheques.Sections("Medio").Controls("lblctabanc").caption = txtcodcuenta
        rptLibracionCheques.Sections("Medio").Controls("lblbanco").caption = txtbanco
        rptLibracionCheques.Sections("Medio").Controls("lbltipocta").caption = txttipocta
        rptLibracionCheques.Sections("Medio").Controls("lblfecha").caption = fechacheque
        rptLibracionCheques.Sections("Medio").Controls("lblconcepto").caption = txtconcepto
        rptLibracionCheques.Sections("Medio").Controls("lblenletras").caption = enletras
        rptLibracionCheques.Sections("Medio").Controls("lblimporte").caption = txtimporte
        
        rptLibracionCheques.Sections("Detalle").Height = 0
        
        sql = "select * from Movicaja  MC inner join RegistroDocumentos on RegistroDocumentos.idDoc = MC.idDoc " & _
                " where MC.interno = " & Val(interno) & " and MC.tipodoc = 'LIB' order by MC.movimiento"
        
        rsNroPago.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If rsNroPago.EOF And rsNroPago.BOF Then Exit Sub
        rptLibracionCheques.Sections("Encabezado").Controls("LblOrdenPago").caption = rsNroPago!NumeroDePago
        Set rsNroPago = Nothing
        
        
        
        
        Set rptLibracionCheques.Sections("Encabezado").Controls("image1").Picture = FrmPrincipal.imgLogoSimple  ' LOGO DE LA EMPRESA PREDETERMINADO
        rptLibracionCheques.Show vbModal
        GoTo fin
Imprimir:
        RptLibracionChequeProv.Fecha = fechaoper 'Date
        RptLibracionChequeProv.TxtChequeInterno = txtint
        RptLibracionChequeProv.TxtNroCheque = txtnumcheque
        RptLibracionChequeProv.CodCuentaBanc = txtcodcuenta
        RptLibracionChequeProv.TxtNumeroCuenta = txtnumcuenta
        RptLibracionChequeProv.txtbanco = txtbanco
        RptLibracionChequeProv.TxtTipoCuenta = txttipocta
        RptLibracionChequeProv.TxtFechaCheque = fechacheque
        RptLibracionChequeProv.txtconcepto = txtconcepto
        RptLibracionChequeProv.txtimporte.Text = NroEnLetras(CDbl(txtimporte))
        RptLibracionChequeProv.Importe = txtimporte
        
        sql = "SELECT RegistroDocumentos.NumeroDePago,MoviCaja.MOVIMIENTO, Asientos.NroAsiento, MAYOR.Debe, MAYOR.Haber, MoviCaja.INTERNO, MoviCaja.TipoDoc, CUENTAS.DESCRIPCION, CUENTAS.Cuenta " & _
        " FROM (((MoviCaja INNER JOIN Asientos ON MoviCaja.idDoc = Asientos.idDoc) INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento) INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) " & _
        " inner join RegistroDocumentos on RegistroDocumentos.idDoc = MoviCaja.idDoc " & _
        " WHERE (((MAYOR.Debe)>0) AND ((MoviCaja.INTERNO)=" & txtint & ") AND ((MoviCaja.TipoDoc)='LIB'));"
        
        RptLibracionChequeProv.DataLibracion.Connection = DataEnvironment1.Sistema
        RptLibracionChequeProv.DataLibracion.Source = sql
        RptLibracionChequeProv.Show vbModal
fin:
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
    If txtcodcuenta.enabled Then txtcodcuenta.SetFocus

End Sub

Function sumogrilla() As Double
    Dim x As Long
    Dim Total As Double
    
    For x = 1 To grilla.rows - 1
        Total = Total + s2n(grilla.TextMatrix(x, 3))
    Next
    sumogrilla = Total
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uCuenta.ini uCuentaIni1Imput, uCuentaIni2Imput, True
 ' inicializo caja/prov
    If BS_CHQCOMP_USAR_NROINT Then
        uChe.ini _
            " select b.descripcion from chq_comp c left outer join bancosgrales b on c.banco = b.codigo where c.activo = 1 and c.estado = 'C' and c.codigo = ### ", _
            " SELECT  ch.CODIGO as [CodigoInterno], ch.Nro as [Numero              ], ba.descripcion as [ Banco            ], ch.NRO as [ Numero        ] FROM CHQ_COMP ch LEFT OUTER JOIN BancosGrales ba ON ba.codigo = ch.Banco WHERE ch.activo=1 and (ch.ESTADO = 'C')"
    Else
        uChe.ini _
            " select b.descripcion from chq_comp c left outer join bancosgrales b on c.banco = b.codigo where c.activo = 1 and c.estado = 'C' and c.nro = ###", _
            " SELECT ch.Nro as [Numero              ], ch.CODIGO as [CodigoInterno], ba.descripcion as [ Banco            ], ch.NRO as [ Numero        ] FROM CHQ_COMP ch LEFT OUTER JOIN BancosGrales ba ON ba.codigo = ch.Banco WHERE ch.activo=1 and (ch.ESTADO = 'C')"
    End If
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
        Label34.Visible = False
    End If
    
    
End Sub

Private Function ChequeContraQue(que As enumChequeContraQue)

    If mViejo = que Then Exit Function  '  na'pa'cer
    
    Select Case que
    Case chequeCaja
        lblcartel.Visible = False
        fraX.Visible = True
        lblX.caption = "Caja: "
        uX.ini "select responsable from cajas where codigo = ### and activo = 1", "select codigo, responsable as [ Descripcion            ] from cajas where activo = 1 "
        uX.codigo = 1
        fraGrilla.enabled = False
    Case chequeProv
        lblcartel.Visible = False
        fraX.Visible = True
        lblX.caption = "Prov: "
        uX.ini "select descripcion from prov where codigo = ### and activo  =1", "select codigo, descripcion from prov where activo = 1 order by codigo"
        fraGrilla.enabled = fraTodo.enabled
    Case chequeOtro
        lblcartel.Visible = True
        fraX.Visible = False
        uX.clear
        fraGrilla.enabled = fraTodo.enabled
    End Select
    mViejo = que
End Function


Private Sub CargaCheque(cual)
    Dim rs As New ADODB.Recordset
    Dim s As String
    
    
    If BS_CHQCOMP_USAR_NROINT Then
        
        s = "select Chq_comp.*, Chq_comp.codigo as interno, Ctasbank.tipo, Ctasbank.numero from Chq_comp inner join Ctasbank on Chq_comp.cuentabancaria = Ctasbank.codigo where Chq_comp.codigo = " & cual & " and Chq_comp.activo = 1"
    Else
        s = "select Chq_comp.*, Chq_comp.codigo as interno, Ctasbank.tipo, Ctasbank.numero from Chq_comp inner join Ctasbank on Chq_comp.cuentabancaria = Ctasbank.codigo where Chq_comp.nro = " & cual & " and Chq_comp.activo = 1"
    End If
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        txtcodcuenta = rs!cuentabancaria
        txtnumcuenta = rs!numero
        txtbanco = ObtenerDescripcionBancos("BancosGrales", rs!Banco)
        txttipocta = ObtenerDescripcion("TipoCtas", Val(rs!tipo))
        txtint = rs!interno
        txtnumcheque = rs!Nro
        txtconcepto.SetFocus
    Else
        'MsgBox "El cheque no existe", vbInformation
        txtint = ""
        'txtint.SetFocus
        txtnumcheque = ""
        txtcodcuenta = ""
        txtnumcuenta = ""
        txtbanco = ""
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Sub CargarMovimientos()
    Dim Proveedor As Long, rs As New ADODB.Recordset
    
        Proveedor = 0
        InicioGrilla
        rs.Open "select Chq_comp.*, Ctasbank.tipo, Ctasbank.numero from Chq_comp inner join Ctasbank on Chq_comp.cuentabancaria = Ctasbank.codigo where Chq_comp.codigo = " & Val(txtint) & " and Chq_comp.activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcuenta = rs!cuentabancaria
            txtnumcuenta = rs!numero
            txtbanco = ObtenerDescripcionBancos("BancosGrales", rs!Banco)
            txttipocta = ObtenerDescripcion("TipoCtas", Val(rs!tipo))
            txtnumcheque = rs!Nro
            fechacheque = rs!fechadeposito
            Proveedor = rs!Proveedor
        End If
        rs.Close
        Set rs = Nothing
        
        rs.Open "select * from Movicaja where interno = " & Val(txtint) & " and tipodoc = 'LIB' order by movimiento", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                                             
        If Not rs.EOF Then
            txtconcepto = rs!concepto
            txtimporte = rs!Importe
            fechaoper = rs!Fecha
            If rs!tipo_libracion = "C" Then
                optX(chequeCaja).Value = True
                ChequeContraQue (chequeCaja)
                uX.codigo = rs!caja
            Else
                If rs!tipo_libracion = "O" Then
                    optX(chequeOtro).Value = True
                    ChequeContraQue chequeOtro
                Else
                    optX(chequeProv).Value = True
                    ChequeContraQue chequeProv
                    uX.codigo = Proveedor
                End If
                IngresoGrilla (rs!iddoc)
            End If
        End If

    Set rs = Nothing
End Sub

Private Sub IngresoGrilla(iddoc As Long)
Dim rs1 As New ADODB.Recordset

    If iddoc = 0 Then
        'che "datos no migrados"
        Exit Sub
    End If
    
    'rs1.Open "select * from Detallemovcajas where movimiento = " & Val(movimiento) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    rs1.Open "select * from asientos inner join mayor on mayor.idAsiento = asientos.idAsiento where iddoc = '" & iddoc & "' and debe > 0 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs1.EOF Then
        habilitogrillaenable (False)
        While Not rs1.EOF
            Call CargogrillaTabla(rs1!Cuenta, rs1!COMPROBANTE, rs1!Debe)
            rs1.MoveNext
        Wend
    End If
    rs1.Close
    Set rs1 = Nothing

End Sub



Private Sub CargogrillaTabla(Cuenta As Long, concepto As String, Importe As Double)
'    If grilla.rows = 2 Then
'        grilla.row = 1
'        grilla.col = 0
'        If Trim(grilla.Text) = "" Then
'            grilla.row = 1
'            grilla.col = 0
'            grilla.Text = cuenta
'            grilla.col = 1
'            grilla.Text = ObtenerDescripcion("Cuentas", Val(grilla.TextMatrix(grilla.row, grilla.col)))
'            grilla.col = 2
'            grilla.Text = concepto
'            grilla.col = 3
'            grilla.Text = Importe
'        Else
'            grilla.AddItem cuenta & Chr(9) & ObtenerDescripcion("Cuentas", Val(grilla.TextMatrix(grilla.row, grilla.col))) & Chr(9) & concepto & Chr(9) & Importe
'        End If
'    Else
        grilla.AddItem Cuenta & Chr(9) & obtenerDeSQL("select descripcion from cuentas where cuenta = '" & Cuenta & "' ") & Chr(9) & concepto & Chr(9) & Importe
'    End If
'    If txttotal <> "" Then
        txttotal = s2n(txttotal) + s2n(Importe)
'    Else
'        txttotal = s2n(Importe)
'    End If
End Sub

Private Sub optX_Click(Index As Integer)
    ChequeContraQue CLng(Index)
End Sub
Private Sub optX_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    ChequeContraQue CLng(Index)
End Sub
Private Sub optX_LostFocus(Index As Integer)
    ChequeContraQue CLng(Index)
End Sub

''debi pasarlo a array controles
'Private Sub optcaja_Click()
'ChequeContraQue
'End Sub
'Private Sub optcaja_KeyUp(KeyCode As Integer, Shift As Integer)
'ChequeContraQue
'End Sub
'Private Sub optcaja_LostFocus()
'ChequeContraQue
'End Sub
'Private Sub optotros_Click()
'ChequeContraQue
'End Sub
'Private Sub optotros_KeyUp(KeyCode As Integer, Shift As Integer)
'ChequeContraQue
'End Sub
'Private Sub optotros_LostFocus()
'ChequeContraQue
'End Sub
'Private Sub optprov_Click()
'ChequeContraQue
'End Sub
'Private Sub optprov_KeyUp(KeyCode As Integer, Shift As Integer)
'ChequeContraQue
'End Sub
'Private Sub optprov_LostFocus()
'ChequeContraQue
'End Sub


Private Sub txtbanco_GotFocus()
    txtbanco.SelStart = 0
    txtbanco.SelLength = Len(txtbanco.Text)
End Sub
Private Sub txtcodcuenta_GotFocus()
    txtcodcuenta.SelStart = 0
    txtcodcuenta.SelLength = Len(txtcodcuenta.Text)
End Sub
Private Sub txtconc_GotFocus()
    If Trim(txtconc) = "" Then txtconc = txtconcepto
    
    txtconc.SelStart = 0
    txtconc.SelLength = Len(txtconc.Text)
End Sub
Private Sub txtConcepto_GotFocus()
    txtconcepto.SelStart = 0
    txtconcepto.SelLength = Len(txtconcepto.Text)
End Sub
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
'        MsgBox "Debe ingresar un importe"
'        txtimporte = "0"
'        txtimporte.SetFocus
    Else
        If optX(chequeProv) = True Then
            habilitogrilla (True)
        End If
        txtimporte = s2n(txtimporte)
    End If
End Sub

'Private Sub txtint_GotFocus()
'    txtint.SelStart = 0
'    txtint.SelLength = Len(txtint.Text)
'End Sub

'Private Sub txtint_KeyPress(KeyAscii As Integer)
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'End Sub

'Private Sub txtint_LostFocus()
'
'    If Trim(txtint) <> "" Then
'        If IsNumeric(txtint) Then
''            cargar = "Cheques"
'            CargaCheque s2n(txtint)
'        End If
'    'Else
'    '    MsgBox "Nº de cheque incorrecto"
''        txtcod = ""
'        'txtcod.SetFocus
'    End If
'
'End Sub


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
    txttotal.Visible = habilito
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

Sub InicioGrilla()
    grilla.clear
    grilla.rows = 1
    grilla.cols = 4
    'grilla.ColWidth(1) = 1700
    grilla.TextMatrix(0, 0) = "Cuenta"
    grilla.TextMatrix(0, 1) = "Descripción"
    grilla.TextMatrix(0, 2) = "Concepto"
    grilla.TextMatrix(0, 3) = "Importe"
    gCUEN = 0
    gDESC = 1
    gCONC = 2
    gIMPO = 3
End Sub

Private Sub habilitogrillaenable(habilito As Boolean)
    Label2.enabled = habilito

    uCuenta.enabled = habilito
    
    Label6.enabled = habilito
    txtconc.enabled = habilito
    Label3.enabled = habilito
    txtvalor.enabled = habilito
    cmdcargar.enabled = habilito
    grilla.enabled = habilito
    cmbeliminofila.enabled = habilito
End Sub

Private Sub Limpiotextosgrilla()
    txtconc = ""
    txtvalor = ""
    uCuenta.clear
End Sub
Private Sub txtnumcheque_GotFocus()
    txtnumcheque.SelStart = 0
    txtnumcheque.SelLength = Len(txtnumcheque.Text)
End Sub

Private Sub txtnumcuenta_GotFocus()
    txtnumcuenta.SelStart = 0
    txtnumcuenta.SelLength = Len(txtnumcuenta.Text)
End Sub

Private Sub txttipocta_GotFocus()
    txttipocta.SelStart = 0
    txttipocta.SelLength = Len(txttipocta.Text)
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
    txtvalor = s2n(txtvalor)
    
'    If optX(chequeCaja) = False And optX(chequeProv) = False And optotros = False And txtvalor <> "" Then
'        MsgBox "Debe ingresar un tipo de movimiento"
'        Exit Sub
'    End If
'    If IsNumeric(txtvalor) Then
''        InicioGrilla
'        If grilla.Visible = False Then
'            habilitogrilla (True)
'        End If
'        habilitogrillaenable (True)
'        txtvalor = s2n(txtvalor)
'    Else
'        If txtvalor <> "" Then
'            MsgBox "Debe ingresar un importe"
'            txtvalor = "0"
'            txtvalor.SetFocus
'        End If
'    End If
End Sub


Private Function alta() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta

    Dim valmaximo As Long, maximobanco As Long, valmaximo2, i
    Dim valorcuentacon As String, cuentaCaja As String
    Dim asiLib As New Asiento, iddoc As Long

    If Not TodoOk Then Exit Function
    
    maximobanco = nuevoCodigo("movibanc", "movBanco")
    valmaximo = nuevoCodigo("movicaja", "movimiento")
    valorcuentacon = sSinNull(obtenerDeSQL("select cuenta_con from ctasBank  inner join cuentas on cuentas.cuenta = ctasBank.cuenta_con where ctasBank.codigo = '" & Val(txtcodcuenta) & "' "))
    If valorcuentacon = "" Then
        che "Falta definicion de cuenta contable para cuenta banco"
        Exit Function
    End If
    If optX(chequeCaja) Then
        cuentaCaja = sSinNull(obtenerDeSQL("select cajas.cuenta from cajas inner join cuentas on cuentas.cuenta = cajas.cuenta where cajas.codigo = '" & uX.codigo & "' "))
        If cuentaCaja = "" Then
            che "Falta definicion de cuenta contable para cuenta caja"
            Exit Function
        End If
    End If
        
    DE_BeginTrans
        
        iddoc = NuevoDocumento("LIB", valmaximo, 0, NuevoNroPago())
        If optX(chequeCaja) Then
               
            DataEnvironment1.dbo_MOVLIBMOVICAJA 0, valmaximo, "P", "E", s2n(txtimporte), _
                Trim(txtconcepto), fechaoper, Val(txtint), "LIB", valmaximo, Trim(valorcuentacon), _
                maximobanco, "C", iddoc, Date, UsuarioSistema!codigo
           
            valmaximo2 = nuevoCodigo("movicaja", "movimiento")

                       
            DataEnvironment1.dbo_MOVLIBMOVICAJA uX.codigo, valmaximo2, "E", "I", s2n(txtimporte), _
                Trim(txtconcepto), fechaoper, Val(txtint), "LIB", valmaximo2, Trim(valorcuentacon), _
               maximobanco, "C", iddoc, Date, UsuarioSistema!codigo

            DataEnvironment1.dbo_MOVLIBMOVIBANC Val(txtcodcuenta), "L", Trim(txtconcepto), fechaoper, _
                "P", Val(txtint), s2n(txtimporte), "LIB", valmaximo, valmaximo, valmaximo, _
                iddoc, Date, UsuarioSistema!codigo
                                                             
                               
            DataEnvironment1.dbo_MOVLIBCHEQUES Val(txtint), fechacheque, s2n(txtimporte), valmaximo, "LIB", "T", fechaoper, _
                fechaoper, 0, _
                iddoc, Date, UsuarioSistema!codigo
                
            asiLib.nuevo "Libracion cheque contra caja", fechaoper, "LA"
            asiLib.AgregarItem cuentaCaja, s2n(txtimporte), 0
            asiLib.AgregarItem valorcuentacon, 0, s2n(txtimporte)
            'asiLib.AgregarItem cuentaCaja, 0, s2n(txtimporte)
        End If
                   
        If optX(chequeProv) = True Then
            DataEnvironment1.dbo_MOVLIBMOVICAJA 0, valmaximo, "P", "E", s2n(txtimporte), _
                Trim(txtconcepto), fechaoper, Val(txtint), "LIB", valmaximo, Trim(valorcuentacon), _
                maximobanco, "P", iddoc, Date, UsuarioSistema!codigo
                
            DataEnvironment1.dbo_MOVLIBMOVIBANC Val(txtcodcuenta), "L", Trim(txtconcepto), fechaoper, _
                 "P", Val(txtint), s2n(txtimporte), "LIB", valmaximo, valmaximo, valmaximo, _
                 iddoc, Date, UsuarioSistema!codigo
            
            DataEnvironment1.dbo_MOVLIBCHEQUES Val(txtint), fechacheque, s2n(txtimporte), valmaximo, "LIB", "T", fechaoper, _
                 fechaoper, uX.codigo, _
                 iddoc, Date, UsuarioSistema!codigo
           
            asiLib.nuevo "Libracion cheque a proveedor", fechaoper, "LC"
            asiLib.AgregarItem valorcuentacon, 0, s2n(txtimporte)
            For i = 1 To grilla.rows - 1
                asiLib.AgregarItem grilla.TextMatrix(i, gCUEN), s2n(grilla.TextMatrix(i, gIMPO)), 0, grilla.TextMatrix(i, gCONC)
            Next
           
        End If
            
        If optX(chequeOtro) = True Then
            DataEnvironment1.dbo_MOVLIBMOVICAJA 0, valmaximo, "P", "E", s2n(txtimporte), _
                Trim(txtconcepto), fechaoper, Val(txtint), "LIB", valmaximo, Trim(valorcuentacon), _
                maximobanco, "O", iddoc, Date, UsuarioSistema!codigo
                
            DataEnvironment1.dbo_MOVLIBMOVIBANC Val(txtcodcuenta), "L", Trim(txtconcepto), fechaoper, _
                "P", Val(txtint), s2n(txtimporte), "LIB", valmaximo, valmaximo, valmaximo, _
                iddoc, Date, UsuarioSistema!codigo
            
'            For i = 1 To grilla.rows - 1
'                DataEnvironment1.dbo_MOVLIBDETALLE valmaximo, s2n(grilla.TextMatrix(i, 3)), grilla.TextMatrix(i, 0), IIf(txtconc <> "", Trim(txtconc), Trim(txtConcepto)), _
'                "LV", "LIB", valmaximo, fechaoper
'            Next
            
            DataEnvironment1.dbo_MOVLIBCHEQUES Val(txtint), fechacheque, s2n(txtimporte), valmaximo, "LIB", "T", fechaoper, _
                fechaoper, 0, _
                iddoc, Date, UsuarioSistema!codigo
                
            asiLib.nuevo "Libracion cheque ", fechaoper, "LV"
            asiLib.AgregarItem valorcuentacon, 0, s2n(txtimporte)
            For i = 1 To grilla.rows - 1
                asiLib.AgregarItem grilla.TextMatrix(i, gCUEN), s2n(grilla.TextMatrix(i, gIMPO)), 0, grilla.TextMatrix(i, gCONC)
            Next
               
        End If
        
         asiLib.Grabar iddoc, , leerEjercicioId(cboEjercicio)                '  tiene err.raise
'        If asiLib.Grabar(iddoc) = 0 Then
'            DE_RollbackTrans
''            che "no se pudo hacer asiento " & asiLib.Diferencia
'            Exit Function
'        End If
        
    DE_CommitTrans
    alta = True
        
    imprimoPantalla txtint
        
    MsgBox "La operación fue realizada con éxito"
    InicioGrilla

fin:
    Set asiLib = Nothing
    Exit Function

UFAalta:
    DE_RollbackTrans
    
    ufa "prg:fallo el alta", "aceptar libracion cheques"
    Resume fin
End Function


Private Function TodoOk() As Boolean
    If Trim(txtint) = "" Then
        MsgBox "Debe ingresar el número interno del cheque"
        Exit Function
    End If
    If (optX(chequeCaja) Or optX(chequeProv)) And uX.codigo = 0 Then
        MsgBox "Debe ingresar el código de caja o proveedor"
        Exit Function
    End If
    If optX(chequeProv) = True And (txtimporte <> txttotal) Then
        MsgBox "Los datos no estan cargados correctamente"
        Exit Function
    End If
    revisoTotalGrilla
    If Not optX(chequeCaja) Then ' si es caja
        If s2n(txttotal) <> s2n(txtimporte) Then
            che "no coinciden totales "
            Exit Function
        End If
    End If
    TodoOk = True
End Function


Private Sub uChe_cambio(codigo As Variant)
    CargaCheque codigo
End Sub

Private Sub uMenu_AceptarAlta()
    If alta Then uMenu.AceptarOk
End Sub

Private Sub uMenu_BorrarControles()
    LimpioControles
    InicioGrilla
    mViejo = -1
    ChequeContraQue (chequeCaja)
End Sub
Private Sub uMenu_Buscar()
    
    Dim re As String, ss As String
    ss = "select interno, fecha, importe from Movicaja  where tipodoc = 'LIB' and caja = 0 order by movimiento desc"
    re = frmBuscar.MostrarSql(ss)
    If re = "" Then Exit Sub
    
    txtint = re
    CargarMovimientos
    uMenu.BuscarOK
End Sub
Private Sub uMenu_eliminar()
    On Error GoTo ufaErr
    
    Dim i As Variant, iddoc As Long
    i = obtenerDeSQL("select movimiento, iddoc from movicaja where interno =  '" & Val(txtint) & "' and tipodoc = 'LIB' ")
    
    If IsEmpty(i) Then
        ufa "prg: no encontre cheque en tabla", "movicaja interno = " & txtint
    Else
        iddoc = i(1)
        If iddoc > 0 Then
            DE_BeginTrans
                'If BorroDocumento(iddoc) Then
                BorroDocumento (iddoc)
                DataEnvironment1.dbo_ELIMINOLIBRACION i(0), i(1)
                grabaBitacora "B", s2n(txtint), "chq_comp"
                AsientoBaja_idDoc iddoc
            DE_CommitTrans
            uMenu.EliminarOK
            'End If
        Else
            DataEnvironment1.dbo_ELIMINOLIBRACION i(0), i(1)
            grabaBitacora "B", s2n(txtint), "chq_comp"
            uMenu.EliminarOK
        End If
    End If
fin:
    
    Exit Sub
ufaErr:
    DE_RollbackTrans
    ufa "fallo eliminacion", "elim libracion"
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoControles sino ' al pedo, pero queda
    fraTodo.enabled = sino
End Sub
Private Sub uMenu_Imprimir()
    imprimoPantalla txtint
End Sub


Private Sub uMenu_Nuevo()
    On Error Resume Next
    optX(0).SetFocus
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub

