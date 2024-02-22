VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmExtractoTonka 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Extracto Bancario"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin Gestion.ucFecha uFeDe 
      Height          =   300
      Left            =   810
      TabIndex        =   0
      Top             =   60
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      fechainit       =   0
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7455
      TabIndex        =   8
      Top             =   1530
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   810
      TabIndex        =   4
      Top             =   1530
      Width           =   885
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Operacion"
      Height          =   330
      Index           =   1
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   75
      Width           =   840
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Libracion"
      Height          =   345
      Index           =   0
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   75
      Value           =   -1  'True
      Width           =   855
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4680
      Left            =   135
      TabIndex        =   12
      Top             =   2040
      Width           =   11085
      _cx             =   19553
      _cy             =   8255
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
      AllowUserResizing=   1
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
   Begin Gestion.uCtaBanco uCtaBanDe 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   750
      Width           =   7890
      _extentx        =   13917
      _extenty        =   556
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   5145
      TabIndex        =   7
      Top             =   1545
      Width           =   855
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      Top             =   1560
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   635
   End
   Begin Gestion.uCtaBanco uCtaBanHa 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   1125
      Width           =   7890
      _extentx        =   12700
      _extenty        =   556
   End
   Begin Gestion.ucFecha uFeHa 
      Height          =   300
      Left            =   825
      TabIndex        =   1
      Top             =   405
      Width           =   945
      _extentx        =   1667
      _extenty        =   529
      fechainit       =   4
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "considerar fecha: "
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
      Index           =   1
      Left            =   4890
      TabIndex        =   15
      Top             =   105
      Width           =   1725
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   75
      TabIndex        =   11
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hasta"
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
      Left            =   120
      TabIndex        =   9
      Top             =   405
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desde"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   105
      Width           =   615
   End
End
Attribute VB_Name = "FrmExtractoTonka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' un asco. lo hice mierda el:  26/6/6, todavia falta saldos.
'
Private Const s_OPE_CRED_EMIS = "ETDV"  ' la V es propuesta, ch propios rechazados, NO implementado en ningun lado aun
Private Const s_OPE_DEBI_EMIS = "SGLR"

Private Const s_OPE_CRED_ACRE = "ETA"  ' no muestro rechazos porque no cierra con acreditaciones
Private Const s_OPE_DEBI_ACRE = "SGB"
'

Private Sub cmdAceptar_Click()
    Dim rs As New ADODB.Recordset
    
    Dim Banco As String, nrocheque As String, cuenta As Long
    Dim ss As String, ssi As String, ssic As String, ssid As String
    Dim tempo As Variant
    Dim deb As Double, cre As Double, fech As String
    Dim saldoC As Double, SaldoD As Double, saldoRestoD As Double, saldoRestoC As Double
    
    InicioGrilla
    
    
    If optFecha(0) Then ' fecha libracion/deposito
        ssi = "select sum(importe) as sumacred from movibanc " & _
            " where activo = 1 " & _
            " and fecha < " & uFeDe.ssFecha
        ssic = " and (operacion = 'E' or operacion = 'T' or operacion = 'D' or operacion = 'V' ) "
        ssid = " and (operacion = 'S' or operacion = 'G' or operacion = 'L' or operacion = 'R' ) "


        ss = " select m.*, ctasbank.numero, tipoctas.descripcion as tides, bancosgrales.descripcion as desbanco " & _
            " from (((movibanc m inner join ctasbank on m.cuenta = ctasbank.codigo) inner join bancosgrales on ctasbank.banco = bancosgrales.codigo) left join tipoctas on ctasbank.tipo = tipoctas.codigo) " & _
            " where m.activo = 1 " & _
            " and ( m.fecha " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & " ) " & _
            " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
            " order by m.cuenta, m.fecha, m.interno"

    Else                ' fecha operacion debito/acreditacion
    
        ssi = "select sum(importe) as sumacred from movibanc " & _
            " where activo = 1 " & _
            " and fecha < " & uFeDe.ssFecha
        ssic = " and (operacion = 'E' or operacion = 'T' or operacion = 'A' ) "
        ssid = " and (operacion = 'S' or operacion = 'G' or operacion = 'B' ) "
    
        ss = " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, q.FECHA_OPERACION AS Fe_oper " & _
            " FROM MoviBanc m INNER JOIN " & _
            " CTASBANK cu ON m.CUENTA = cu.CODIGO INNER JOIN " & _
            " BancosGrales b ON cu.BANCO = b.codigo LEFT OUTER JOIN " & _
            " CHQ_COMP q ON m.INTERNO = q.CODIGO LEFT OUTER JOIN " & _
            " TipoCtas t ON cu.TIPO = t.Codigo " & _
            " where m.activo = 1 " & _
            " and m.fecha " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & _
            " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
            " order by m.cuenta, m.fecha, m.interno"
            
    End If
    
    
    With rs
        .Open ss, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            
            If cuenta <> !cuenta Then  ' corte
'                If grilla.rows > 1 Then grilla.AddItem cuenta
                saldoC = s2n(obtenerDeSQL(ssi & ssic & " and cuenta = " & !cuenta))
                SaldoD = s2n(obtenerDeSQL(ssi & ssid & " and cuenta = " & !cuenta))
                saldoRestoC = IIf(saldoC - SaldoD > 0, Round(saldoC - SaldoD, 2), 0)
                saldoRestoD = IIf(SaldoD - saldoC > 0, Round(SaldoD - saldoC, 2), 0)
                grilla.AddItem !cuenta & vbTab & !cuenta & " " & !desbanco & vbTab & !numero & vbTab & saldoRestoD & vbTab & saldoRestoC
            End If
            
            fech = !fecha
            chequeBanco nrocheque, Banco, !documento, s2n(!interno)
            credidebi cre, deb, !importe, !operacion
                       
            grilla.AddItem !cuenta & vbTab _
                & fech & vbTab _
                & sSinNull(!descripcion) & vbTab _
                & deb & vbTab _
                & cre & vbTab _
                & " " & vbTab _
                & !documento & vbTab _
                & nSinNull(!interno) & vbTab _
                & Banco & vbTab _
                & nrocheque & vbTab _
                & !MovBanco
            
            cuenta = !cuenta
            .MoveNext
            
        Wend
    End With
    Set rs = Nothing
    'grillaSumarizo grilla, Array(3, 4)
    grilla.SubtotalPosition = flexSTBelow
    
    grilla.subtotal flexSTSum, 0, 3, , , , True
    grilla.subtotal flexSTSum, 0, 4, , , , True
    
    Dim i:
    For i = 1 To grilla.rows - 1
    If grilla.IsSubtotal(i) Then grilla.TextMatrix(i, 5) = grilla.TextMatrix(i, 4) - grilla.TextMatrix(i, 3)
    Next i
    
End Sub

Private Function credidebi(refCred As Double, refDebi As Double, importe As Double, operacion As String)
    ' devuelve columnas credito y debito  en parametros byref
    
    refCred = 0
    refDebi = 0
    
    If optFecha(0) Then  ' por fecha propio/libracion 3ro/deposito
        If InStr(s_OPE_CRED_EMIS, operacion) > 0 Then
            refCred = importe
        ElseIf InStr(s_OPE_DEBI_EMIS, operacion) > 0 Then
            refDebi = importe
        End If
    Else                 ' por fecha propio/debito    3ro/acreditacion
        If InStr(s_OPE_CRED_ACRE, operacion) > 0 Then
            refCred = importe
        ElseIf InStr(s_OPE_DEBI_ACRE, operacion) > 0 Then
            refDebi = importe
        End If
    End If
End Function

Private Function chequeBanco(refQueNro As String, refQueBanco As String, docu As String, interno As Long)
    ' devuelve banco y cheque en parametros byref
    Dim tempo
    refQueBanco = ""
    refQueNro = ""
    If docu = "P" Then    ' ch propios
        tempo = obtenerDeSQL("select chq_comp.nro, bancosgrales.descripcion from chq_comp inner join bancosgrales on chq_comp.banco = bancosgrales.codigo where chq_comp.codigo = " & interno)
        If Not IsEmpty(tempo) Then
            refQueNro = tempo(0)
            refQueBanco = tempo(1)
        End If
    ElseIf docu = "C" Then   'ch 3ros
        tempo = obtenerDeSQL("select cheques.nro, bancosgrales.descripcion from cheques inner join bancosgrales on cheques.banco_nro = bancosgrales.codigo where cheques.nroint = " & interno)
        If Not IsEmpty(tempo) Then
            refQueNro = tempo(0)
            refQueBanco = tempo(1)
        End If
    End If
End Function

Private Sub cmdImprimir_Click()
    If grilla.rows < 2 Then Exit Sub
    
    grilla.GridLines = flexGridNone
    grilla.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 8
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Extracto bancario "
    'FrmImpresiones.VSPrinter.Paragraph = "Para cuentas entre " &
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & uFeDe.dtfecha & " - " & uFeHa.dtfecha
    FrmImpresiones.VSPrinter.Paragraph = " "
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    
    FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d "
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla.GridLines = flexGridFlat

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InicioGrilla
    ucXls1.ini grilla, "C:\ExtractoBanc"
    Form_Resize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub InicioGrilla()
    grilla.cols = 11
    grilla.rows = 1
    grillaWidth grilla, Array(0, 2100, 2500, 2000, 1200, 1000, 1000, 800, 800, 1250, 1100, 1100)
    grillaTitulos grilla, Array("", "Fecha", "Descripción", "Débitos", "Créditos", "Saldo", "Doc.", "Nº Int.", "Banco", "Nº Cheque", "Nº Mov.")
End Sub

Private Sub uFeDe_LostFocus()
    uFeHa.setUltDiaMes uFeDe.Mes, uFeDe.Anio
End Sub
