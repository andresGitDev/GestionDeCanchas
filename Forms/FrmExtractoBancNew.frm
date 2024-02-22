VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmExtractoBancNew 
   Caption         =   "Extracto Bancario"
   ClientHeight    =   9030
   ClientLeft      =   750
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "FrmExtractoBancNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FrmExtractoBancNew.frx":08CA
      Left            =   8520
      List            =   "FrmExtractoBancNew.frx":08FE
      TabIndex        =   19
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   7560
      TabIndex        =   14
      Top             =   8040
      Width           =   3615
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   870
         Left            =   2625
         Picture         =   "FrmExtractoBancNew.frx":09FE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   870
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Mostrar"
         Height          =   870
         Left            =   0
         Picture         =   "FrmExtractoBancNew.frx":12C8
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   870
         Left            =   855
         Picture         =   "FrmExtractoBancNew.frx":1B92
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   840
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   885
         Left            =   1695
         TabIndex        =   18
         Top             =   0
         Width           =   930
         _extentx        =   1640
         _extenty        =   1561
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FrmExtractoBancNew.frx":245C
      Left            =   5160
      List            =   "FrmExtractoBancNew.frx":246C
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   120
      Width           =   2415
   End
   Begin Gestion.uCtaBanco uCtaBanHa 
      Height          =   300
      Left            =   1065
      TabIndex        =   11
      Top             =   870
      Width           =   7965
      _extentx        =   14049
      _extenty        =   529
   End
   Begin Gestion.uCtaBanco uCtaBanDe 
      Height          =   360
      Left            =   1065
      TabIndex        =   10
      Top             =   480
      Width           =   7905
      _extentx        =   13944
      _extenty        =   635
   End
   Begin Gestion.ucFecha uFeHa 
      Height          =   300
      Left            =   9915
      TabIndex        =   9
      Top             =   885
      Width           =   1005
      _extentx        =   1773
      _extenty        =   529
      fechainit       =   0
   End
   Begin Gestion.ucFecha uFeDe 
      Height          =   330
      Left            =   9915
      TabIndex        =   8
      Top             =   495
      Width           =   1005
      _extentx        =   1773
      _extenty        =   582
      fechainit       =   0
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Vencimiento"
      Height          =   330
      Index           =   1
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   1110
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Operacion"
      Height          =   345
      Index           =   0
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Value           =   -1  'True
      Width           =   1080
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   6765
      Left            =   30
      TabIndex        =   4
      Top             =   1245
      Width           =   11190
      _cx             =   19738
      _cy             =   11933
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
   Begin VB.Label Label4 
      Caption         =   "Por tipo :"
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Filtrar por :"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Considerar Fecha: "
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
      Left            =   210
      TabIndex        =   7
      Top             =   105
      Width           =   1725
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
      Left            =   195
      TabIndex        =   3
      Top             =   930
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
      Left            =   180
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9165
      TabIndex        =   1
      Top             =   915
      Width           =   615
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9150
      TabIndex        =   0
      Top             =   555
      Width           =   615
   End
End
Attribute VB_Name = "frmExtractoBancNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' un asco. lo hice mierda el:  26/6/6, todavia falta saldos.
'
Private Const s_OPE_CRED_EMIS = "DEP" '"ETDV" ' "ETDV"  ' la V es propuesta, ch propios rechazados, NO implementado en ningun lado aun
Private Const s_OPE_DEBI_EMIS = "LGSR" '"SGLR"

Private Const s_OPE_CRED_ACRE = "AEP" '"ETA"  ' no muestro rechazos porque no cierra con acreditaciones
Private Const s_OPE_DEBI_ACRE = "BGRS" '"SGBL" '"SGB"
'

Private Sub cmdAceptar_Click()
    Dim rs As New ADODB.Recordset
    
    Dim Banco As String, nrocheque As String, Cuenta As Long
    Dim ss As String, ssi As String, ssic As String, ssid As String, ss_Ope As String
    Dim filtro As String
    Dim tempo As Variant
    Dim deb As Double, cre As Double, fech As String
    Dim saldoC As Double, SaldoD As Double, saldoRestoD As Double, saldoRestoC As Double
    Dim AAUUXX1
    
    InicioGrilla
    
    If Combo1.ListIndex = 0 Then 'todos
        filtro = ""
    ElseIf Combo1.ListIndex = 1 Then 'credito
        filtro = " and (operacion='E' and documento='T') "
    ElseIf Combo1.ListIndex = 2 Then 'debito
        filtro = " and (operacion='S' and documento='G' and (m.descripcion not like 'Iva 21%' and m.descripcion not like 'Iva 10.5%' and m.descripcion not like 'Gastos%' and m.descripcion not like 'Iva 27%' and m.descripcion not like 'Mant. Cuenta%' and m.descripcion not like 'Mant. Cuenta Sueldos%' and m.descripcion not like 'Gastos por Chequera%' and m.descripcion not like 'Gastos varios%' and m.descripcion not like 'Comisiones por Gestion de Cheq%' and m.descripcion not like 'Imp por Sobre Giro%' and m.descripcion not like 'Comisiones por transferencias%' and m.descripcion not like 'Percepcion de IIBB%' and m.descripcion not like 'Sircreb%' and m.descripcion not like 'Imp por Deb%' and m.descripcion not like 'Imp por Cred%') ) "
    ElseIf Combo1.ListIndex = 3 Then 'gastos
        filtro = " and (operacion='S' and documento='G' "
        If Combo2.ListIndex = 0 Then
            filtro = filtro & " and (m.descripcion like 'Iva 21%' or m.descripcion like 'Iva 10.5%' or m.descripcion like 'Gastos%' or m.descripcion like 'Iva 27%' or m.descripcion like 'Mant. Cuenta%' or m.descripcion like 'Mant. Cuenta Sueldos%' or m.descripcion like 'Gastos por Chequera%' or m.descripcion like 'Gastos varios%' or m.descripcion like 'Comisiones por Gestion de Cheq%' or m.descripcion like 'Imp por Sobre Giro%' or m.descripcion like 'Comisiones por transferencias%' or m.descripcion like 'Percepcion de IIBB%' or m.descripcion like 'Sircreb%' or m.descripcion like 'Imp por Deb%' or m.descripcion like 'Imp por Cred%') ) "
        ElseIf Combo2.ListIndex = 1 Then
            filtro = filtro & " and (m.descripcion like 'Iva 21%' )) "
        ElseIf Combo2.ListIndex = 2 Then
            filtro = filtro & " and (m.descripcion like 'Iva 10.5%' ) ) "
        ElseIf Combo2.ListIndex = 3 Then
            filtro = filtro & " and (m.descripcion like 'Iva 27%' ) ) "
        ElseIf Combo2.ListIndex = 4 Then
            filtro = filtro & " and (m.descripcion like 'Gastos%' and m.descripcion not like 'Gastos por Chequera%' and m.descripcion not like 'Gastos varios%' ) ) "
        ElseIf Combo2.ListIndex = 5 Then
            filtro = filtro & " and (m.descripcion like 'Mant. Cuenta%' and m.descripcion not like 'Mant. Cuenta Sueldos%' ) ) "
        ElseIf Combo2.ListIndex = 6 Then
            filtro = filtro & " and (m.descripcion like 'Mant. Cuenta Sueldos%' ) ) "
        ElseIf Combo2.ListIndex = 7 Then
            filtro = filtro & " and (m.descripcion like 'Gastos por Chequera%' ) ) "
        ElseIf Combo2.ListIndex = 8 Then
            filtro = filtro & " and (m.descripcion like 'Gastos varios%' ) ) "
        ElseIf Combo2.ListIndex = 9 Then
            filtro = filtro & " and (m.descripcion like 'Comisiones por Gestion de Cheq%' ) ) "
        ElseIf Combo2.ListIndex = 10 Then
            filtro = filtro & " and (m.descripcion like 'Imp por Sobre Giro%' ) ) "
        ElseIf Combo2.ListIndex = 11 Then
            filtro = filtro & " and (m.descripcion like 'Comisiones por transferencias%' ) ) "
        ElseIf Combo2.ListIndex = 12 Then
            filtro = filtro & " and (m.descripcion like 'Percepcion de IIBB%' ) ) "
        ElseIf Combo2.ListIndex = 13 Then
            filtro = filtro & " and (m.descripcion like 'Sircreb%' ) ) "
        ElseIf Combo2.ListIndex = 14 Then
            filtro = filtro & " and (m.descripcion like 'Imp por Deb%' ) ) "
        ElseIf Combo2.ListIndex = 15 Then
            filtro = filtro & " and (m.descripcion like 'Imp por Cred%') ) "
        End If
    End If
    
    If optfecha(0) Then ' fecha libracion/deposito tiene que ver con la feha del cheque
        ssi = "select sum(importe) as sumacred from movibanc " & _
            " where activo = 1 " & _
            " and fecha < " & uFeDe.ssFecha
        ssic = " and (operacion = 'D' or operacion = 'E' ) " '" and (operacion = 'E' or operacion = 'T' or operacion = 'D' or operacion = 'V' ) "
        ssid = " and (operacion = 'L'  or operacion = 'G' or operacion = 'R'  or operacion = 'S' ) " '" and (operacion = 'S' or operacion = 'G' or operacion = 'L' or operacion = 'R' ) "
        ss_Ope = " and (operacion = 'D' or operacion = 'E' or operacion = 'L'  or operacion = 'G' or operacion = 'R' or operacion = 'S' ) " ' todos los estados
        

        'ss = " select m.*, ctasbank.numero, tipoctas.descripcion as tides, bancosgrales.descripcion as desbanco " & _
            " from (((movibanc m inner join ctasbank on m.cuenta = ctasbank.codigo) inner join bancosgrales on ctasbank.banco = bancosgrales.codigo) left join tipoctas on ctasbank.tipo = tipoctas.codigo) " & _
            " where m.activo = 1 " & _
            " and ( m.fecha " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & " ) " & _
            " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
            " order by m.cuenta, m.fecha, m.interno"
             
        'ss = " SELECT m.*, ctasbank.numero, tipoctas.descripcion as tides, bancosgrales.descripcion as desbanco,CHQ_COMP.fecha_cheque as fe_cheq  " & _
            " FROM MoviBanc m INNER JOIN " & _
            " CTASBANK ON m.CUENTA = ctasbank.CODIGO INNER JOIN " & _
            " BancosGrales  ON ctasbank.BANCO = BancosGrales.codigo LEFT OUTER JOIN " & _
            " CHQ_COMP ON m.INTERNO = CHQ_COMP.CODIGO LEFT OUTER JOIN " & _
            " TipoCtas  ON ctasbank.TIPO = TipoCtas.Codigo " & _
            " where m.activo = 1 " & _
            " and CHQ_COMP.fecha_cheque " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & _
            " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
            " order by m.cuenta, CHQ_COMP.fecha_cheque, m.interno"
'        ss = "SELECT m.*, ctasbank.numero, tipoctas.descripcion as tides, bancosgrales.descripcion as desbanco,CHQ_COMP.fecha_cheque as Fecha_cheque_opera " & _
'            " FROM MoviBanc m " & _
'            " INNER JOIN CTASBANK ON m.CUENTA = ctasbank.CODIGO " & _
'            " INNER JOIN  BancosGrales  ON ctasbank.BANCO = BancosGrales.codigo " & _
'            " LEFT OUTER JOIN  CHQ_COMP ON m.INTERNO = CHQ_COMP.CODIGO " & _
'            " LEFT OUTER JOIN  TipoCtas  ON ctasbank.TIPO = TipoCtas.Codigo " & _
'            " where m.activo = 1  and CHQ_COMP.fecha_cheque " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
'            " Union " & _
'            " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, q.FECHA_OPERACION AS Fecha_cheque_opera " & _
'            " FROM MoviBanc m INNER JOIN  CTASBANK cu ON m.CUENTA = cu.CODIGO " & _
'            " INNER JOIN  BancosGrales b ON cu.BANCO = b.codigo " & _
'            " LEFT OUTER JOIN  CHQ_COMP q ON m.INTERNO = q.CODIGO " & _
'            " LEFT OUTER JOIN  TipoCtas t ON cu.TIPO = t.Codigo " & _
'            " where m.activo = 1 and q.FECHA_OPERACION  " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
'            " Union " & _
'            " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, m.FECHA AS Fecha_cheque_opera " & _
'            " FROM MoviBanc m INNER JOIN  CTASBANK cu ON m.CUENTA = cu.CODIGO " & _
'            " INNER JOIN  BancosGrales b ON cu.BANCO = b.codigo " & _
'            " LEFT OUTER JOIN  TipoCtas t ON cu.TIPO = t.Codigo " & _
'            " where m.activo = 1 and m.FECHA " & ssBetween(uFeDe.dtfecha, uFeHa.dtfecha) & " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & _
'            " order by m.cuenta, m.fecha, m.interno"
        ss = " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, q.FECHA_OPERACION AS Fecha_cheque_opera " & _
            " FROM MoviBanc m INNER JOIN  CTASBANK cu ON m.CUENTA = cu.CODIGO " & _
            " INNER JOIN  BancosGrales b ON cu.BANCO = b.codigo " & _
            " LEFT OUTER JOIN  CHQ_COMP q ON m.INTERNO = q.CODIGO " & _
            " LEFT OUTER JOIN  TipoCtas t ON cu.TIPO = t.Codigo " & _
            " where m.activo = 1 and (q.FECHA_OPERACION  " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & " or q.FECHA_cheque  " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & ") and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & ssic & ssid & filtro & _
            " Union " & _
            " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, m.FECHA AS Fecha_cheque_opera " & _
            " FROM MoviBanc m INNER JOIN  CTASBANK cu ON m.CUENTA = cu.CODIGO " & _
            " INNER JOIN  BancosGrales b ON cu.BANCO = b.codigo " & _
            " LEFT OUTER JOIN  TipoCtas t ON cu.TIPO = t.Codigo " & _
            " where m.activo = 1 and m.FECHA " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & ss_Ope & filtro & _
            " order by m.cuenta, m.fecha, m.interno"

    Else                ' fecha operacion debito/acreditacion
    
        ssi = "select sum(importe) as sumacred from movibanc " & _
            " where activo = 1 " & _
            " and fecha < " & uFeDe.ssFecha
        ssic = " and (operacion = 'A' or operacion = 'E'  ) " '" and (operacion = 'E' or operacion = 'T' or operacion = 'A' or operacion = 'J' ) " 'ingresos a banco
        ssid = " and (operacion = 'B' or operacion = 'G' or operacion = 'R' or operacion = 'S' )" '" and (operacion = 'S' or operacion = 'G' or operacion = 'B' or operacion = 'R') " 'egreso
        ss_Ope = " and (operacion = 'A' or operacion = 'E' or operacion = 'B'  or operacion = 'G' or operacion = 'R' or operacion = 'S' ) " '" and (operacion = 'E' or operacion = 'R' or operacion = 'T' or operacion = 'A' or operacion = 'S' or operacion = 'G' or operacion = 'B' or operacion = 'J') "  ' todo
        
        ss = " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, q.FECHA_OPERACION AS Fecha_cheque_opera " & _
            " FROM MoviBanc m INNER JOIN " & _
            " CTASBANK cu ON m.CUENTA = cu.CODIGO INNER JOIN " & _
            " BancosGrales b ON cu.BANCO = b.codigo LEFT OUTER JOIN " & _
            " CHQ_COMP q ON m.INTERNO = q.CODIGO LEFT OUTER JOIN " & _
            " TipoCtas t ON cu.TIPO = t.Codigo " & _
            " where m.activo = 1 " & _
            " and q.FECHA_OPERACION " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & _
            " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & ssic & ssid & filtro & _
            " Union " & _
            " SELECT m.*, cu.NUMERO , t.Descripcion AS TiDes, b.descripcion AS desbanco, m.FECHA AS Fecha_cheque_opera " & _
            " FROM MoviBanc m INNER JOIN  CTASBANK cu ON m.CUENTA = cu.CODIGO " & _
            " INNER JOIN  BancosGrales b ON cu.BANCO = b.codigo " & _
            " LEFT OUTER JOIN  TipoCtas t ON cu.TIPO = t.Codigo " & _
            " where m.activo = 1 and m.FECHA " & ssBetween(uFeDe.dtFecha, uFeHa.dtFecha) & " and ( m.cuenta between " & uCtaBanDe.codigo & " and " & uCtaBanHa.codigo & " ) " & ss_Ope & filtro & _
            " order by m.cuenta, m.fecha, m.interno"
            
    End If
    
    Dim inte As Long
    Dim mov As Long
    
    With rs

            .Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        'MsgBox "" & rs.RecordCount
        
        inte = 0
        mov = 0
        If .EOF And .BOF Then
            MsgBox "No existen datos para ese periodo.", vbInformation, "Atencion"
            Exit Sub
        Else
        .MoveFirst
            While Not .EOF
                
                 If IIf(IsNull(!interno), 0, !interno) <> inte Or !MovBanco <> mov Then
                    Dim Aux
                    If !OPERACION = "R" Then
                        Aux = !OPERACION
                    End If
                    If Cuenta <> !Cuenta Then  ' corte
                '       If grilla.rows > 1 Then grilla.AddItem cuenta
                        saldoC = s2n(obtenerDeSQL(ssi & ssic & " and cuenta = " & !Cuenta))
                        SaldoD = s2n(obtenerDeSQL(ssi & ssid & " and cuenta = " & !Cuenta))
                        saldoRestoC = IIf(saldoC - SaldoD > 0, Round(saldoC - SaldoD, 2), 0)
                        saldoRestoD = IIf(SaldoD - saldoC > 0, Round(SaldoD - saldoC, 2), 0)
                        GRILLA.AddItem !Cuenta & vbTab & !Cuenta & " " & !desbanco & vbTab & !numero & vbTab & saldoRestoD & vbTab & saldoRestoC
                    End If
    
    
                    If optfecha(0) Then
                        If IsNull(!Fecha_cheque_opera) Or IsEmpty(!Fecha_cheque_opera) Then
                            If IsNull(!FECHA_OPERACION) Or IsEmpty(!FECHA_OPERACION) Then
                                fech = !Fecha
                            Else
                                fech = !FECHA_OPERACION
                            End If
                        Else
                            fech = !Fecha_cheque_opera
                        End If
                    ElseIf optfecha(1) Then
                        If IsNull(!Fecha_cheque_opera) Or IsEmpty(!Fecha_cheque_opera) Then
                            fech = !Fecha
                        Else
                            fech = !Fecha_cheque_opera
                        End If
                    End If
                    'fech = !fecha
    
                    chequeBanco nrocheque, Banco, !documento, s2n(!interno)
                    credidebi cre, deb, !Importe, !OPERACION
        
                    GRILLA.AddItem !Cuenta & vbTab _
                        & fech & vbTab _
                        & sSinNull(!DESCRIPCION) & vbTab _
                        & deb & vbTab _
                        & cre & vbTab _
                        & " " & vbTab _
                        & !documento & vbTab _
                        & nSinNull(!interno) & vbTab _
                        & Banco & vbTab _
                        & nrocheque & vbTab _
                        & !MovBanco
                End If
               
                Cuenta = !Cuenta
                inte = IIf(IsNull(!interno), 1, !interno)
                mov = !MovBanco
                
                .MoveNext
            Wend
        End If
    End With
    
    Set rs = Nothing
    

    
    'grillaSumarizo grilla, Array(3, 4)
    GRILLA.SubtotalPosition = flexSTBelow
    
    GRILLA.subtotal flexSTSum, 0, 3, , , , True
    GRILLA.subtotal flexSTSum, 0, 4, , , , True
    
    Dim i
    For i = 1 To GRILLA.rows - 1
        'If Grilla.IsSubtotal(i) Then Grilla.TextMatrix(i, 5) = Grilla.TextMatrix(i, 4) - Grilla.TextMatrix(i, 3)
        If GRILLA.rows - 1 = i Then
            GRILLA.TextMatrix(i, 5) = GRILLA.TextMatrix(i, 4) - GRILLA.TextMatrix(i, 3)
        Else
            If i = 1 Then
                GRILLA.TextMatrix(i, 5) = GRILLA.TextMatrix(i, 4) - GRILLA.TextMatrix(i, 3)
            Else
                GRILLA.TextMatrix(i, 5) = s2n(GRILLA.TextMatrix(i - 1, 5), 4) + s2n(GRILLA.TextMatrix(i, 4), 4) - s2n(GRILLA.TextMatrix(i, 3), 4)
            End If
        End If
    Next i
    
    If GRILLA.rows > 1 Then GRILLA.TextMatrix(GRILLA.rows - 1, 5) = s2n(GRILLA.TextMatrix(GRILLA.rows - 1, 5))
    
    
End Sub

Private Function credidebi(refCred As Double, refDebi As Double, Importe As Double, OPERACION As String)
    ' devuelve columnas credito y debito  en parametros byref
    
    refCred = 0
    refDebi = 0
    
    If optfecha(0) Then  ' por fecha propio/libracion 3ro/deposito
        If InStr(s_OPE_CRED_EMIS, OPERACION) > 0 Then
            refCred = Importe
        ElseIf InStr(s_OPE_DEBI_EMIS, OPERACION) > 0 Then
            refDebi = Importe
        End If
    Else                 ' por fecha propio/debito    3ro/acreditacion
        If InStr(s_OPE_CRED_ACRE, OPERACION) > 0 Then
            refCred = Importe
        ElseIf InStr(s_OPE_DEBI_ACRE, OPERACION) > 0 Then
            refDebi = Importe
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
    If GRILLA.rows < 2 Then Exit Sub
   
    GRILLA.GridLines = flexGridNone
    GRILLA.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = GRILLA.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 8
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Extracto bancario "
    'FrmImpresiones.VSPrinter.Paragraph = "Para cuentas entre " &
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & uFeDe.dtFecha & " - " & uFeHa.dtFecha
    FrmImpresiones.VSPrinter.Paragraph = " "
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    
    FrmImpresiones.VSPrinter.RenderControl = GRILLA.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d "
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    GRILLA.GridLines = flexGridFlat

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 3 Then
        Combo2.enabled = True
    Else
        Combo2.enabled = False
    End If
End Sub

Private Sub Form_Load()
    InicioGrilla
    ucXls1.ini GRILLA, "C:\ExtractoBanc"
    ucXls1.caption = "Guardar en Excel"
    Form_Resize
    Combo2.ListIndex = 0
    Combo1.ListIndex = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Resize()
    Anclar GRILLA, Me, anclarLadosTodos
'    Anclar cmdAceptar, Me, anclarAbajo + anclarDerecha
'    Anclar cmdImprimir, Me, anclarAbajo + anclarDerecha
'    Anclar ucXls1, Me, anclarAbajo + anclarDerecha
'    Anclar cmdSalir, Me, anclarAbajo + anclarDerecha
    Anclar Frame1, Me, anclarAbajo + anclarDerecha
End Sub

Private Sub InicioGrilla()
    GRILLA.cols = 11
    GRILLA.rows = 1
    grillaWidth GRILLA, Array(0, 2100, 2500, 2000, 1200, 1000, 1000, 800, 800, 1250, 1100, 1100)
    grillaTitulos GRILLA, Array("", "Fecha", "Descripción", "Debitos", "Creditos", "Saldo", "Doc.", "Nº Int.", "Banco", "Nº Cheque", "Nº Mov.")
End Sub


Private Sub uFeDe_LostFocus()
    'uFeHa.setUltDiaMes uFeDe.Mes, uFeDe.Anio
End Sub
