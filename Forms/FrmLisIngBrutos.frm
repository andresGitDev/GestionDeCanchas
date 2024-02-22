VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmLisIngBrutos 
   Caption         =   "Listado de Ingresos Brutos por Provincia"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "FrmLisIngBrutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucEntreFechas ucEntreFechas1 
      Height          =   330
      Left            =   45
      TabIndex        =   8
      Top             =   330
      Width           =   5445
      _ExtentX        =   5715
      _ExtentY        =   820
   End
   Begin VB.TextBox TotalIVA 
      Height          =   330
      Left            =   4065
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox TotalNeto 
      Height          =   330
      Left            =   2565
      TabIndex        =   5
      Top             =   4440
      Width           =   1485
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   960
      Left            =   3555
      Picture         =   "FrmLisIngBrutos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   975
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla 
      Height          =   3630
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   5640
      _cx             =   9948
      _cy             =   6403
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
      Rows            =   1
      Cols            =   3
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
      AutoSearch      =   2
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
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   960
      Left            =   2565
      Picture         =   "FrmLisIngBrutos.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4845
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   960
      Left            =   4545
      Picture         =   "FrmLisIngBrutos.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1635
      TabIndex        =   7
      Top             =   4470
      Width           =   2370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione el Periodo a Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   45
      TabIndex        =   3
      Top             =   60
      Width           =   4485
   End
End
Attribute VB_Name = "FrmLisIngBrutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sql As String
Dim str As String


Private Sub CmdEjecutar_Click()
Dim TNeto As Double, TIVA As Double
'
'    sql = "SELECT Provincias.descripcion, SUM(FacturaVenta.Neto) AS SumaNeto, SUM(FacturaVenta.Iva) " & _
'    "AS SumaIVA FROM FacturaVenta INNER JOIN Provincias ON " & _
'    "FacturaVenta.Provincia = Provincias.codigo WHERE FacturaVenta.fecha BETWEEN " & _
'    "'" & ucEntreFechas1.desde & "' AND '" & ucEntreFechas1.hasta & "' and (FacturaVenta.tipodoc='FAA' OR " & _
'    "FacturaVenta.TIPODOC='FAB' OR FacturaVenta.TIPODOC='NCA' OR FacturaVenta.TIPODOC='NDA' OR " & _
'    "FacturaVenta.TIPODOC='NCB' OR FacturaVenta.TIPODOC='NDB') and FacturaVenta.ACTIVO = 1 GROUP BY " & _
'    "Provincias.descripcion ORDER BY provincias.descripcion "
'
'    rs.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'Do While Not rs.EOF
'    Grilla.AddItem rs!descripcion & Chr(9) & Format$(rs!sumaneto, "standard") & Chr(9) & Format$(rs!sumaiva, "standard")
'    TNeto = TNeto + Format$(rs!sumaneto, "standard")
'    TIVA = TIVA + Format$(rs!sumaiva, "standard")
'    rs.MoveNext
'Loop
'rs.Close
'TotalNeto = Format(TNeto, "standard")
'TotalIva = Format(TIVA, "standard")

    Dim rs As New ADODB.Recordset
    Dim Consulta As String
    Dim signo As Variant
    Dim sTmpIvaVentas As String
    Dim Neto As Double
    Dim Iva As Double
    
    sTmpIvaVentas = TablaTempCrear(tt_IIBB)
    ControlPrevioProv
    Consulta = "SELECT Provincias.descripcion,FacturaVenta.codigo, FacturaVenta.TipoDoc,FacturaVenta.Neto,FacturaVenta.Iva " & _
                      "FROM FacturaVenta INNER JOIN Provincias ON FacturaVenta.Provincia = Provincias.codigo " & _
                      " WHERE FacturaVenta.fecha " & ssBetween(ucEntreFechas1.desde, ucEntreFechas1.hasta) & " and " & _
    "(FacturaVenta.tipodoc='FAA' OR FacturaVenta.TIPODOC='FAB' OR FacturaVenta.TIPODOC='NCA' OR " & _
    "FacturaVenta.TIPODOC='NDA' OR FacturaVenta.TIPODOC='NCB' OR FacturaVenta.TIPODOC='NDB') " & _
    "and FacturaVenta.ACTIVO = 1 and provincias.activo=1 ORDER BY provincias.descripcion"
    
    rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            signo = ""
            If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then
                signo = "-"
            End If
            Neto = rs!Neto
            Iva = rs!Iva
            If Trim(rs!TIPODOC) = "FAB" Then
                Iva = Format$((s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=21"), 2) / 121) * 21, "standard")
                If Iva = 0 Then
                    Iva = Format$((s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!codigo & ") and _iva=10.5"), 2) / s2n(110.5)) * s2n(10.5), "standard")
                End If
                Neto = Neto - Iva
            End If
            Consulta = "Insert Into " & sTmpIvaVentas & " (provincia,neto,iva) " & _
            "Values ( '" & rs!DESCRIPCION & "' , " & signo & x2s(Neto) & ", " & signo & x2s(nSinNull(Iva)) & ")"
            DataEnvironment1.Sistema.Execute Consulta
            rs.MoveNext
        Wend
    End If

    Set rs = Nothing
                                                                                
    str = "SELECT Provincia, SUM(Neto) AS SumaNeto, SUM(Iva) " & _
    "AS SumaIVA FROM " & sTmpIvaVentas & " GROUP BY Provincia ORDER BY provincia"
    rs.Open str, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    Dim g As Integer

    Cargogrilla
    Do While Not rs.EOF
        grilla.AddItem rs!Provincia & Chr(9) & Format$(rs!sumaneto, "standard") & Chr(9) & Format$(rs!sumaiva, "standard")
        TNeto = TNeto + Format$(rs!sumaneto, "standard")
        TIVA = TIVA + Format$(rs!sumaiva, "standard")
        rs.MoveNext
    Loop
    rs.Close
    TotalNeto = Format(TNeto, "standard")
    Totaliva = Format(TIVA, "standard")


End Sub
Sub ControlPrevioProv()
    Dim rs As New ADODB.Recordset
    Dim rsC As New ADODB.Recordset
    Dim Sql2 As String
    Dim valor1 As Single, valor2 As Single
    valor1 = 0
    valor2 = 0
    'sql = "SELECT COUNT(Provincia) AS Cuantos From FacturaVenta WHERE (tipodoc = 'FAA' or tipodoc = 'FAB' or tipodoc = 'NCA' or tipodoc = 'NCB' or tipodoc = 'NDA'  or tipodoc = 'NDB' ) AND  (Provincia = '') AND (Fecha BETWEEN  '" & ucEntreFechas1.desde & "' AND '" & ucEntreFechas1.hasta & "') and activo = 1"
    sql = "SELECT COUNT(Provincia) AS Cuantos From FacturaVenta WHERE (tipodoc = 'FAA' or tipodoc = 'FAB' or tipodoc = 'NCA' or tipodoc = 'NCB' or tipodoc = 'NDA'  or tipodoc = 'NDB' ) AND  (Provincia = '') AND (Fecha " & ssBetween(ucEntreFechas1.desde, ucEntreFechas1.hasta) & ") and activo = 1"
    '"WHERE FECHA " & ssBetween(dtfechad, dtfechah) &
    Sql2 = "select count(provincia) as Cuantos FROM Clientes WHERE (Provincia='')"
  
    rs.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    rsC.Open Sql2, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    
    If Not rs.EOF Then valor1 = rs!cuantos
    If Not rsC.EOF Then valor2 = rsC!cuantos
    If valor1 = 0 And valor2 = 0 Then
    Else
        MsgBox "Las Facturas sin Provincias asignadas son : " & valor1 & Chr(13) & Chr(13) & "Los Clientes sin Provincias asignadas son : " & valor2 & Chr(13) & Chr(13) & "Esto hara que el Informe no sea exacto", vbInformation, "Aviso"
    End If


    Set rs = Nothing
    Set rsC = Nothing
End Sub


Private Sub cmdImprimir_Click()
    RptListIIBB.DataC.Connection = DataEnvironment1.Sistema
    RptListIIBB.DataC.Source = str
    RptListIIBB.TotalNeto.Text = TotalNeto
    RptListIIBB.Totaliva.Text = Totaliva
    RptListIIBB.fechadesde.Text = ucEntreFechas1.desde
    RptListIIBB.fechahasta.Text = ucEntreFechas1.hasta
    RptListIIBB.Show vbModal
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
    Cargogrilla
End Sub

Sub Cargogrilla()
grilla.clear
grilla.rows = 1
grilla.TextMatrix(0, 0) = " Provincia "
grilla.TextMatrix(0, 1) = " Neto "
grilla.TextMatrix(0, 2) = "IVA"
grilla.ColWidth(0) = 2500
grilla.ColWidth(1) = 1500
grilla.ColWidth(2) = 1500

End Sub


