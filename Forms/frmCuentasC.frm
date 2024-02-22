VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmCuentasC 
   Caption         =   "Cuentas de clientes"
   ClientHeight    =   11775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "frmCuentasC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11775
   ScaleWidth      =   11085
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ver"
      Height          =   855
      Left            =   1125
      TabIndex        =   2
      Top             =   15
      Width           =   900
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   915
      Left            =   75
      TabIndex        =   1
      Top             =   15
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   1614
   End
   Begin VSFlex7LCtl.VSFlexGrid gCuentas 
      Height          =   10635
      Left            =   75
      TabIndex        =   0
      Top             =   990
      Width           =   10860
      _cx             =   19156
      _cy             =   18759
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
End
Attribute VB_Name = "frmCuentasC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVer_Click()
cVer
End Sub

Private Sub Form_Load()
ucXls1.ini gCuentas, "C:\CuentasClientes.xls"
cVer
End Sub

'Private Function cVer()
'Dim sVeo As String, i As Long
'Dim cCuentas, j As Long, cCargar As String
'sVeo = "Select CODIGO AS CLIENTE, DESCRIPCION,CUENTASVENTAS AS [CUENTAS CONTABLES],'' AS DETALLE FROM CLIENTES WHERE ACTIVO=1"
'LlenarGrilla gCuentas, sVeo, False
'With gCuentas
'    If .rows > 1 Then
'        .ColWidth(0) = 1000
'        .ColWidth(1) = 3500
'        .ColWidth(2) = 8500
'        .ColWidth(3) = 8500
'        .ColAlignment(2) = flexAlignLeftCenter
'        For i = 1 To .rows - 1
'            .TextMatrix(i, 2) = Replace(.TextMatrix(i, 2), "#", "")
'            cCuentas = Split(.TextMatrix(i, 2), ",")
'            If IsArray(cCuentas) Then
'                cCargar = ""
'                For j = 0 To UBound(cCuentas)
'                    If Trim(cCuentas(j)) = "" Then
'                    Else
'                        cCargar = cCargar & " # " & obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(cCuentas(j)))
'                    End If
'                Next
'                .TextMatrix(i, 3) = Trim(cCargar)
'            Else
'                If Trim(cCuentas) = "" Then
'                    .TextMatrix(i, 3) = "SIN CUENTAS CONTABLES"
'                Else
'                    .TextMatrix(i, 3) = obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(cCuentas))
'                End If
'            End If
'        Next
'    End If
'End With
'End Function

Private Function cVer()
Dim sVeo As String, i As Long
Dim cCuentas, j As Long, cCargar As String
Dim str As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
sVeo = "Select CODIGO AS CLIENTE, DESCRIPCION,'' AS [CUENTAS CONTABLES],'' AS DETALLE FROM CLIENTES WHERE activo=1"
LlenarGrilla gCuentas, sVeo, False
gCuentas.rows = 1
sVeo = "Select CODIGO AS CLIENTE, DESCRIPCION,'' AS [CUENTAS CONTABLES],'' AS DETALLE FROM CLIENTES WHERE ACTIVO=1"
rs2.Open sVeo, DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
With gCuentas
    While Not rs2.EOF
        .AddItem rs2!cliente & Chr(9) & rs2!DESCRIPCION
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 1500
        .ColWidth(3) = 3500
        .ColAlignment(2) = flexAlignLeftCenter
        
        str = "select cuentasventas from clientes where codigo=" & rs2!cliente
        rs.Open str, DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not IsNull(rs!cuentasventas) Then
            cCuentas = Split(Replace(Trim(rs!cuentasventas), "#", ""), ",")
            For j = 0 To UBound(cCuentas)
                If cCuentas(j) <> "" Then
                    .AddItem "" & Chr(9) & "" & Chr(9) & cCuentas(j) & Chr(9) & obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(cCuentas(j)))
                End If
            Next
        End If
        Set rs = Nothing
        
        rs2.MoveNext
    Wend
    Set rs2 = Nothing
End With
End Function

