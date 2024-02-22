VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisMovProv2 
   Caption         =   "Listado de composicion de saldo de proveedores sin detalles"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   3840
      TabIndex        =   17
      Top             =   1380
      Width           =   3615
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Con Saldo"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   210
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Saldo"
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame fraBoton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   120
      TabIndex        =   11
      Top             =   6465
      Width           =   11175
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
         Height          =   495
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   15
         Width           =   975
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   -15
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
         Height          =   495
         Left            =   3075
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   -15
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6375
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   495
         Left            =   7410
         TabIndex        =   16
         Top             =   15
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
      End
   End
   Begin VB.Frame fraGrilla 
      BackColor       =   &H00E0E0E0&
      Height          =   4380
      Left            =   105
      TabIndex        =   10
      Top             =   1995
      Width           =   11055
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   4035
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   7117
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLisMovProv2.frx":0000
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   8775
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   945
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61865985
         CurrentDate     =   39173
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   930
         TabIndex        =   7
         Top             =   585
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61865985
         CurrentDate     =   39347
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo al:"
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
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   585
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   240
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin Gestion.ucCoDe uProvD 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   315
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uProvH 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   270
         TabIndex        =   4
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   285
         TabIndex        =   3
         Top             =   705
         Width           =   1095
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla 
      Height          =   1830
      Left            =   315
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4470
      Visible         =   0   'False
      Width           =   10935
      _cx             =   19288
      _cy             =   3228
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
      Rows            =   2
      Cols            =   2
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
End
Attribute VB_Name = "frmLisMovProv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'NOTA  8/6/6
'   La tabla tiene los campos numericos como STRING,
'   asi que con el sql NO hay que transformar el numero
'   al formato yanki (con punto)  " SET SALDO =   12.34  "
'   sino que se trabaja con coma, " SET SALDO =  '12,34' "


Private Const CONST_AJUSTE_PROV_CREDITO = "APC"
Private Const CONST_AJUSTE_PROV_DEBITO = "APD"
Private Const CONST_FACTURA = "FAC"
Private Const CONST_NOTAS_DEBITOS = "N/D"
Private Const CONST_NOTAS_CREDITOS = "N/C"
Private Const CONST_RECIBOS_CUENTA = "RAC" 'ANTES RAC
                                            ' nota 6/4/6:   ?????
Private Const CONST_RECIBOS = "REC"
Private Const CONST_IMPUTACION = "IMP"
Public TablaTemp As String
Private Const CONST_CONTADO = 1


Private Function CalcularSaldoAnterior(CodigoProveedor As Long, fechahasta As Date) As Double
    'CONSULTAS DE LAS DISTINTAS TABLAS PARA OBTENER EL SALDO A LA FECHA
    'UNA VEZ REALIZADAS LAS CONSULTAS HAY QUE RECORRER LAS TUPLAS Y SUMAR Y/O RESTAR DE ACUERDO AL TIPO DE DOCUMENTO

'SELECT TIPODOC, FORMADEPAGO, sum(total)
'From TRANSCOM
'WHERE ACTIVO = 1 AND CODPR = 1001 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY TIPODOC, FORMADEPAGO

'SELECT CLIENTE, SUM(TOTAL)
'From RECIBOS
'WHERE ACTIVO = 1 AND CLIENTE = 1000 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY CLIENTE

'SELECT TIPODOC, sum(total)
'From COMPRAS
'WHERE ACTIVO = 1 AND CODPR = 1001 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY TIPODOC

    Dim Debe As Double
    Dim haber As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim rsaux As New ADODB.Recordset
    Dim rsaux2 As New ADODB.Recordset
    Dim Consulta As String, tot As Double

    Debe = 0
    haber = 0
    With rsCuenta
        'TABLA TRANSCOM
        Consulta = "Select TIPODOC, Sum(TOTAL) as Total " & _
                " From TRANSCOM " & _
                " Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                " And FECHA < " & ssFecha(fechahasta) & _
                " Group By TIPODOC"

        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Select Case x2s(!TIPODOC)
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber + s2n(!Total, 4)
            Case Else
                Debe = Debe + s2n(!Total, 4)
            End Select
            .MoveNext
        Wend
        .Close

        
        'TABLA REC_COMP  ' grabadas con el nuevo
        Consulta = "Select SUM(TOTAL) AS TOTAL " & _
                    " From REC_COMP " & _
                    " Where ACTIVO = 1 and iddoc >0 " & _
                    " And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta)
        Debe = Debe + s2n(obtenerDeSQL(Consulta), 4)
        'TABLA REC_COMP ' grabadas con el viejo
        Consulta = "Select SUM(TOTAL + RetGanPago + IBPago ) AS TOTAL " & _
                    " From REC_COMP " & _
                    " Where ACTIVO = 1 and iddoc < 1 " & _
                    " And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta)
        Debe = Debe + s2n(obtenerDeSQL(Consulta), 4)
        
    
        'TABLA COMPRAS
        Consulta = "Select TIPODOC,  SUM(TOTAL) AS TOTAL " & _
                    " From COMPRAS " & _
                    " Where formadepago<>-1 and ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta) & " and contado = 0 " & _
                    " Group by TIPODOC"
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            tot = s2n(!Total, 4)
            Select Case x2s(!TIPODOC)
            
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber + tot
            Case Else
                Debe = Debe + tot
            End Select
            .MoveNext
        Wend
        .Close
    
    End With
    
    
    '******************* esto es para restar las facturas q tienen recibos
    'busco recibos
    Consulta = "Select * From REC_COMP r inner join relfnr_c c on " & _
                    " c.ndoc=r.nro and c.iddoc=r.iddoc Where r.ACTIVO = 1 " & _
                    " And r.CODPR = " & CodigoProveedor & _
                    " And r.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) 'ssFecha(fechahasta)
                    
    rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rsaux.EOF
        Consulta = "select * from compras where codpr=" & CodigoProveedor & " and fecha < " & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!Fact
        rsaux2.Open "select * from compras where codpr=" & CodigoProveedor & " and fecha < " & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!Fact, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rsaux2.EOF = True And rsaux2.BOF = True Then
            Set rsaux2 = Nothing
            rsaux2.Open "select * from transcom where codpr=" & CodigoProveedor & " and fecha<" & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!Fact, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If rsaux2.EOF = True And rsaux2.BOF = True Then
            Else
                Select Case x2s(rsaux!tfac)
                Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                    haber = haber - s2n(rsaux!Impor, 4)
                Case Else
                    Debe = Debe - s2n(rsaux!Impor, 4)
                End Select
            End If
        Else
            'Consulta = "select * from transcom where codpr=" & CodigoProveedor & " and fecha<" & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!fact
            Select Case x2s(rsaux!tfac)
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber - s2n(rsaux!Impor, 4)
            Case Else
                Debe = Debe - s2n(rsaux!Impor, 4)
            End Select
        End If
        Set rsaux2 = Nothing
        rsaux.MoveNext
    Wend
    
    Set rsaux = Nothing
    'busco imputaciones
    Consulta = "Select * From imppro r inner join relfnr_c c on " & _
                    " c.ndoc=r.nro and c.iddoc=r.iddoc Where r.ACTIVO = 1 " & _
                    " And c.tdoc='IMP' and r.CODPR = " & CodigoProveedor & _
                    " And r.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) 'ssFecha(fechahasta)
                    
    rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rsaux.EOF
        rsaux2.Open "select * from compras where codpr=" & CodigoProveedor & " and fecha<" & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!Fact, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rsaux2.EOF = True And rsaux2.BOF = True Then
            Set rsaux2 = Nothing
            rsaux2.Open "select * from transcom where codpr=" & CodigoProveedor & " and fecha<" & ssFecha(fechahasta) & " and tipodoc='" & rsaux!tfac & "' and nrodoc=" & rsaux!Fact, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If rsaux2.EOF = True And rsaux2.BOF = True Then
            Else
                Select Case x2s(rsaux!tfac) '!TIPODOC
                Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                    haber = haber - s2n(rsaux!Impor, 4)
                Case Else
                    Debe = Debe - s2n(rsaux!Impor, 4)
                End Select
            End If
        Else
            tot = s2n(rsaux!Impor, 4)
            
            If s2n(rsaux!FormadePago) = -1 Then
                tot = 0
            End If
            
            Select Case x2s(rsaux!tfac)
            
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber - tot
            Case Else
                Debe = Debe - tot
            End Select
        End If
        Set rsaux2 = Nothing
        rsaux.MoveNext
    Wend
    '************************************************************************************************
    
    
    CalcularSaldoAnterior = Debe - haber
    
    Set rsCuenta = Nothing
End Function

Private Sub CalcularSaldo()
    Dim rsaux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Variant
    Dim CodigoProv As Long
    Dim CodigoProvActual As Long

    With rsaux
        Consulta = "Select * From " & TablaTemp & " Order By CODIGO_PROV, FECHA, ID"
        
        .Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic

        While Not .EOF
            CodigoProv = !CODIGO_PROV
            CodigoProvActual = CodigoProv
            saldo = 0

            While CodigoProv = CodigoProvActual
                If Not IsNull(!Debe) And Not IsNull(!haber) Then
                        saldo = s2n(saldo) + s2n(rsaux!Debe) - s2n(rsaux!haber)
                End If
                
                Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & !ID
                DataEnvironment1.Sistema.Execute Consulta
                
                .MoveNext
                If .EOF Then
                    CodigoProvActual = 0
                Else
                    CodigoProvActual = !CODIGO_PROV
                End If
            Wend
            
        Wend
    End With
    Set rsaux = Nothing
End Sub

Private Sub CrearConsulta()
    Dim rsProv As New ADODB.Recordset
    
    With rsProv
        .Open " Select CODIGO, DESCRIPCION " & _
              " From PROV " & _
              " Where CODIGO >= " & uProvD.codigo & _
              "   and CODIGO <= " & uProvH.codigo & _
              " Order By CODIGO", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
        While Not .EOF
            CrearConsulta_Prov !codigo, ssStr(!DESCRIPCION)
            .MoveNext
        Wend
    
    End With
    CalcularSaldo
    
    Set rsProv = Nothing
End Sub

Private Sub CrearConsulta_Prov(CodigoProv As Long, DescripcionProv As String)
    
    Dim Saldo_Cuenta As Double
    Dim Consulta As String
    Dim rs As New ADODB.Recordset
    Dim sDebe As String, sHaber As String
    Dim tot_ConSinRet As Double
    Dim tot As Double
    Dim Total As Double
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
                                
    'saldo_Cuenta = Round(CalcularSaldoAnterior(CodigoProv, dtfechad.Value), 2)
    Saldo_Cuenta = Round(frmLisMovProv.CalcularSaldoAnterior2(CodigoProv, dtfechad.Value), 2)
    If Saldo_Cuenta < 0 Then
        sDebe = "'0'"
        sHaber = " '" & Abs(Saldo_Cuenta) & "' "
    Else
        sDebe = " '" & Saldo_Cuenta & "' "
        sHaber = "'0'"
    End If
    
    'If Option2.Value = True Then
    'End If
    
    If Option3.Value = True Then 'son sin saldo = 0
        If Saldo_Cuenta = 0 Then
        Else
            Consulta = " Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                       " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(dtfechad.Value) & _
                        ", 'SI', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
            DataEnvironment1.Sistema.Execute Consulta
        End If
    End If
    
    If Option4.Value = True Then
        Consulta = " Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                       " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(dtfechad.Value) & _
                        ", 'SI', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
        DataEnvironment1.Sistema.Execute Consulta
    End If
                            
                            
    Saldo_Cuenta = 0
    
    
    With rs
        'TABLA TRANSCOM
        Consulta = "Select FECHA, TIPODOC, NRODOC,TOTAL, SALDO, RAZONSOCIALPROV, FORMADEPAGO " & _
                " From TRANSCOM Where ACTIVO = 1 " & _
                " AND CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            Total = !Total
            If !saldo = !Total Then 'si son iguales es xq esta para pagar toda la factura
                tot = s2n(!saldo) ' tot numerico, lo redondeo
                
                Select Case x2s(!TIPODOC)
                 Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                    sDebe = "'0'"
                    sHaber = " '" & tot & "' "
                Case Else
                    sDebe = " '" & tot & "' "
                    sHaber = "'0'"
                End Select
                
                Consulta = " Insert Into  " & TablaTemp & _
                    " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, saldo, DEBE, HABER) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & ", " & _
                    " '" & ssStr(rs!TIPODOC) & "', '" & (rs!NroDoc) & "', '" & Saldo_Cuenta & "', " & _
                    sDebe & ", " & sHaber & ") "
           
                DataEnvironment1.Sistema.Execute Consulta
    '                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
    '                If rs!TIPODOC = CONST_FACTURA And rs!FormadePago = CONST_CONTADO Then
    ''                    Saldo_Cuenta = Saldo_Cuenta - rs!Total
    '                    Consulta = "Insert Into LIST_MOV_CUENTA_PROV (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
    '                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
    '                                        "VALUES (" & CodigoProv & ", '" & x2s(rs!razonsocialprov) & "', " & ssFecha(rs!fecha) & _
    '                                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
    '                    daTaenvironment1.Sistema.Execute Consulta
    '                End If
                .MoveNext
            Else   'entonces aca va si la factura tiene un pago parcial
                Set rs2 = Nothing
                rs2.Open "select * from  relfnr_c r where impor<>0 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                'rs2.Open "select * from  relfnr_c r inner join rec_comp c on c.nro=r.ndoc and c.iddoc=r.iddoc  where impor<>0 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " and fecha " & ssBetween(dtfechad.Value, dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                'rs2.MoveFirst
                If rs2.RecordCount > 0 Then rs2.MoveFirst
                'Total = s2n(!Total)
                While Not rs2.EOF
                    
                    If rs2!tdoc = "REC" Then
                        Set rs4 = Nothing
                        rs4.Open "select * from relfnr_c r inner join rec_comp c on c.nro=r.ndoc and c.iddoc=r.iddoc where c.activo=1 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " and r.tdoc='REC' AND NDOC=" & rs2!nDoc & " and fecha " & ssBetween(dtfechad.Value, dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                       
                        If rs4.EOF And rs4.BOF Then
                        Else
                            rs4.MoveFirst
                            While Not rs4.EOF
                                Total = Total - rs4!Impor
                                rs4.MoveNext
                            Wend
                        End If
                    ElseIf rs2!tdoc = "IMP" Then
                        Set rs4 = Nothing
                        rs4.Open "select * from relfnr_c r inner join imppro c on c.nro=r.ndoc and c.iddoc=r.iddoc where c.activo=1 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " and r.tdoc='IMP'  AND NDOC=" & rs2!nDoc & "  and fecha " & ssBetween(dtfechad.Value, dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        
                        If rs4.EOF And rs4.BOF Then
                        Else
                            rs4.MoveFirst
                            While Not rs4.EOF
                                Total = Total - rs4!Impor
                                rs4.MoveNext
                            Wend
                        End If
                    End If
                    rs2.MoveNext
                Wend
                    tot = s2n(Total)
                    Select Case x2s(!TIPODOC)
                    Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                        sDebe = "'0'"
                        sHaber = " '" & tot & "' "
                    Case Else
                        sDebe = " '" & tot & "' "
                        sHaber = "'0'"
                    End Select
                    
                    
                    Consulta = " Insert Into  " & TablaTemp & _
                    " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, saldo, DEBE, HABER) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & ", " & _
                    " '" & ssStr(rs!TIPODOC) & "', '" & (rs!NroDoc) & "', '" & Saldo_Cuenta & "', " & _
                    sDebe & ", " & sHaber & ") "
                    
                    DataEnvironment1.Sistema.Execute Consulta
                
                .MoveNext
            End If
        Wend
        .Close
        Set rs2 = Nothing
        
       
        Total = 0
        'TABLA COMPRAS
        Consulta = "Select CODPR, RAZONSOCIALPROV, FECHA, TIPODOC, NRODOC, TOTAL, CONTADO,FormadePago From COMPRAS Where ACTIVO = 1 And " & _
                    "CODPR = " & CodigoProv & " And FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
              
        
        While Not .EOF
            Total = !Total
            
            If s2n(!FormadePago) = -1 Then
                GoTo porcanje
            End If
            
            
            rs2.Open " Select tfac, facT , impor,tdoc,ndoc,totaldocu,iddoc " & _
                    " From RELFNR_C  Where FACT = " & !NroDoc & " AND (TFAC='" & !TIPODOC & "') " & _
                    " and prov = " & CodigoProv, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            '& " and iddoc>0"
            If rs2.RecordCount > 0 Then rs2.MoveFirst
            
            If rs2.EOF = True And rs2.BOF = True Then
                tot = s2n(!Total)
                Select Case x2s(!TIPODOC)
                 Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                    sDebe = "'0'"
                    sHaber = " '" & tot & "' "
                Case Else
                    sDebe = " '" & tot & "' "
                    sHaber = "'0'"
                End Select
    
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                    " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                    ", '" & !TIPODOC & "', '" & !NroDoc & "', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
                
                DataEnvironment1.Sistema.Execute Consulta
                
                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
                If !TIPODOC = CONST_FACTURA And !contado Then
                    Consulta = "Insert Into " & TablaTemp & " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                        " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                        " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                        ", 'CON', '" & !NroDoc & "', '" & tot & "', '0', '" & Saldo_Cuenta & "')"
                    DataEnvironment1.Sistema.Execute Consulta
                End If
            Else
                rs2.MoveFirst
                While Not rs2.EOF
                    If Trim(rs2!tdoc) = "REC" Then
                        Set rs3 = Nothing
                        rs3.Open "select * from relfnr_c r inner join rec_comp c on c.nro=r.ndoc and c.iddoc=r.iddoc where c.activo=1 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " and r.tdoc='REC' AND NDOC=" & rs2!nDoc & " and fecha " & ssBetween(dtfechad.Value, dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                        rs3.Open " Select * From REC_COMP  Where activo=1 and nro= " & rs2!nDoc & " and codpr = " & CodigoProv & " and iddoc=" & rs2!iddoc & " and fecha<=" & ssFecha(dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If rs3.EOF = True And rs3.BOF = True Then 'rs3.RecordCount > 0 Then
'                            tot = s2n(rs2!impor) 's2n(!Total)
                        Else
                            rs3.MoveFirst
                            While Not rs3.EOF
                                Total = Total - rs2!Impor
                                rs3.MoveNext
                            Wend
                        End If
                    ElseIf Trim(rs2!tdoc) = "IMP" Then
                        Set rs3 = Nothing
                        rs3.Open "select * from relfnr_c r inner join imppro c on c.nro=r.ndoc and c.iddoc=r.iddoc where c.activo=1 and prov=" & CodigoProv & " and tfac='" & !TIPODOC & "' and fact=" & !NroDoc & " and r.tdoc='IMP'  AND NDOC=" & rs2!nDoc & "  and fecha " & ssBetween(dtfechad.Value, dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'                        rs3.Open " Select * From imppro  Where nro= " & rs2!nDoc & " and codpr = " & CodigoProv & " and iddoc=" & rs2!iddoc & " and fecha<=" & ssFecha(dtfechah.Value), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                        If rs3.EOF = True And rs3.BOF = True Then
'                            tot = s2n(rs2!impor)
                        Else
                            rs3.MoveFirst
                            While Not rs3.EOF
                                Total = Total - rs2!Impor
                                rs3.MoveNext
                            Wend
                        End If
                    End If
                    
                    Set rs3 = Nothing
                    rs2.MoveNext
                Wend
                tot = s2n(Total)
                Select Case x2s(!TIPODOC)
                    Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                        sDebe = "'0'"
                        sHaber = " '" & tot & "' "
                    Case Else
                        sDebe = " '" & tot & "' "
                        sHaber = "'0'"
                End Select
    
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                    " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                    ", '" & !TIPODOC & "', '" & !NroDoc & "', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
                
                DataEnvironment1.Sistema.Execute Consulta
                
                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
                If !TIPODOC = CONST_FACTURA And !contado Then
                    Consulta = "Insert Into " & TablaTemp & " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                        " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                        " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                        ", 'CON', '" & !NroDoc & "', '" & tot & "', '0', '" & Saldo_Cuenta & "')"
                    DataEnvironment1.Sistema.Execute Consulta
                End If
                
            End If
            
porcanje:
            
            Set rs2 = Nothing
          
            .MoveNext
        Wend
        .Close
    End With
        
    Set rs = Nothing
    
End Sub

Private Sub cmdAceptar_Click()
    'If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    If uProvD.codigo = 0 Then Exit Sub
    If uProvH.codigo = 0 Then uProvH.codigo = uProvD.codigo
    'If Trim(txtCodProvd.Text) <> "" And Trim(txtcodprovh.Text) <> "" Then
        
    relojito
    
    TablaTemp = TablaTempCrear("(" _
        & "[ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_PROV] [numeric](18, 0) NULL ," _
        & "[DESCRIPCION_PROV] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FECHA] [datetime] NULL ," _
        & "[TIPO_DOCUMENTO] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[NRO_DOCUMENTO] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DEBE] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[HABER] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[SALDO] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[Obs] [char] (10) NULL " _
        & ") ON [PRIMARY]")
    DataEnvironment1.Sistema.Execute "ALTER TABLE  " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [PK_" & TablaTemp & "] PRIMARY KEY  CLUSTERED" _
        & "([id])  ON [PRIMARY]"
    DataEnvironment1.Sistema.Execute "ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [DF_ " & TablaTemp & "1] DEFAULT (0) FOR [FECHA]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "2] DEFAULT (0) FOR [DEBE]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "3] DEFAULT (0) FOR [HABER]," _
        & "CONSTRAINT [DF_ " & TablaTemp & "4] DEFAULT (0) FOR [SALDO]"
    DataEnvironment1.Sistema.Execute "CREATE  INDEX [IX_ " & TablaTemp & "] ON " & TablaTemp & " ([ID]) ON [PRIMARY]"


    CrearConsulta
    
    
        'LlenarGrilla Grilla, _
                " Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', " & _
                " FECHA, TIPO_DOCUMENTO AS DOC, NRO_DOCUMENTO AS NUMERO, " & _
                " DEBE, HABER, SALDO, '' as [SaldoFinal] " & _
                " From " & TablaTemp & _
                " Order By CODIGO_PROV, FECHA, ID", False
        LlenarGrilla grilla, _
                " Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', " & _
                " FECHA, TIPO_DOCUMENTO AS DOC, NRO_DOCUMENTO AS NUMERO, " & _
                " DEBE, HABER, SALDO, '' as [SaldoFinal] " & _
                " From " & TablaTemp & _
                " Order By CODIGO_PROV, FECHA, ID", False
        grillaMarcoSaldosFinales grilla, 0, 8, 7
        limpioGrilla 8
        'grilla.Visible = True
    
    grillaWidth grilla, Array(740, 2800, 1000, 600, 900, 900, 900, 900, 900)
    relojito False 'ver que hace y dejarlo si sirve...
    
    grilla.AddItem Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & sumoTOTc
'fin:
'    relojito False
'    Exit Sub
'ufaChe:
'    che "err en consulta "
'    GoTo fin

    filtragrilla
End Sub
Private Function limpioGrilla(Col As Long) 'limpio la grilla y tabla temporal de los importes con cero, incluyendo si el total es cero borro historial de cliente
    Dim i As Long
    Dim j As Long
    Dim cli As Long
    Dim Borrar As String
    
    i = 1
    While i < grilla.rows
        If Trim(grilla.TextMatrix(i, Col)) = "0" Then
            cli = grilla.TextMatrix(i, 0)
            j = 1
            While j < grilla.rows
                If grilla.TextMatrix(j, 0) = CStr(cli) Then
                    grilla.TextMatrix(j, 0) = ""
                End If
                j = j + 1
            Wend
        End If
        i = i + 1
    Wend
    
    j = 1
    While j < grilla.rows
        If grilla.TextMatrix(j, 0) = "" Then
            Borrar = "delete from " & TablaTemp & " where descripcion_prov='" & grilla.TextMatrix(j, 1) & "'"
            DataEnvironment1.Sistema.Execute Borrar
            grilla.RemoveItem (j)
        Else
            j = j + 1
        End If
        'j = j + 1
    Wend
End Function
Private Function filtragrilla()
    Dim i As Long
    Dim cli As Long
    Dim Valor As Double
    Dim CUIT As String
    Dim nom As String
    
    grilla2.rows = 1
    i = 1
    Valor = 0
    If grilla.rows > 1 Then
        If s2n(grilla.TextMatrix(i, 5)) > 0 Then
            Valor = s2n(grilla.TextMatrix(i, 5))
        Else
            Valor = -s2n(grilla.TextMatrix(i, 6))
        End If
        cli = s2n(grilla.TextMatrix(i, 0))
        i = i + 1
        While i < grilla.rows
            If grilla.TextMatrix(i, 0) <> "" Then
                If cli = grilla.TextMatrix(i, 0) Then
                    If s2n(grilla.TextMatrix(i, 5)) <> 0 Then
                        Valor = Valor + s2n(grilla.TextMatrix(i, 5))
                    Else
                        Valor = Valor - s2n(grilla.TextMatrix(i, 6))
                    End If
                Else
                    CUIT = obtenerDeSQL("select cuit from prov where codigo=" & cli)
                    nom = obtenerDeSQL("select descripcion from prov where codigo=" & cli)
                    grilla2.AddItem CUIT & Chr(9) & nom & Chr(9) & s2n(Valor)
                    Valor = 0
                    cli = grilla.TextMatrix(i, 0)
                    If s2n(grilla.TextMatrix(i, 5)) > 0 Then
                        Valor = Valor + s2n(grilla.TextMatrix(i, 5))
                    Else
                        Valor = Valor - s2n(grilla.TextMatrix(i, 6))
                    End If
                End If
            End If
            i = i + 1
        Wend
        CUIT = obtenerDeSQL("select cuit from prov where codigo=" & cli)
        nom = obtenerDeSQL("select descripcion from prov where codigo=" & cli)
        grilla2.AddItem CUIT & Chr(9) & nom & Chr(9) & s2n(Valor)
    Else
        MsgBox "No se han encontrado datos para la busqueda.", , "ATENCION"
    End If
End Function

Private Function sumoTOTc() As Double
Dim d As Long, t As Double
t = 0
    For d = 1 To grilla.rows - 1
        If grilla.TextMatrix(d, 8) <> "" Then
            t = t + grilla.TextMatrix(d, 8)
        End If
    Next
sumoTOTc = t
End Function

Private Sub cmdImprimir_Click()
    Dim Consulta As String
    Dim rsempresa As New ADODB.Recordset
    Dim i As Long

    
    
    rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    'RptLisMovCtaProv2.ImageLOGO.Picture = FrmPrincipal.imgLogoSimple 'LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
    
    If grilla2.rows < 2 Then Exit Sub
    
    grilla2.GridLines = flexGridNone
    grilla2.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.MarginLeft = 2000
    FrmImpresiones.VSPrinter.Orientation = orPortrait ' orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla2.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 20
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 16
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Listado de composicion de proveedores sin detalle"
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & dtfechad.Value & " - " & dtfechah.Value  '& "     Rango de Cuentas : " & CmbCtaD & "  -  " & CmbCtaH
    FrmImpresiones.VSPrinter.Paragraph = " "
    
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    FrmImpresiones.VSPrinter.RenderControl = grilla2.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla2.GridLines = flexGridFlat
    
    Set rsempresa = Nothing
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next
    uProvH.codigo = obtenerDeSQL("select max(codigo) from prov where activo = 1")
    uProvD.codigo = 1
    dtfechad.Value = "01/04/2007" 'Date
    dtfechah.Value = Date
    ucXls1.ini grilla, "C:\LisCompCtaProv", "Listado movimiento cuenta proveedores"
    uProvD.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    Dim p1 As String, p2 As String
    p1 = "Select descripcion from prov where codigo = '###' and activo = 1"
    p2 = "select codigo, descripcion as [Nombre                         ] from prov where activo = 1"
    uProvD.ini p1, p2, False
    uProvH.ini p1, p2, False
    cmdCancelar_Click
    Form_Resize
    ini
    ucXls1.ini grilla2, "C:\CompProvSinDet.xls", "Composicion de proveedor sin detalle"
End Sub

Private Sub ini()
    grilla2.rows = 1
    grilla2.TextMatrix(0, 0) = "CUIT"
    grilla2.TextMatrix(0, 1) = "NOMBRE"
    grilla2.TextMatrix(0, 2) = "SALDO"
End Sub

Private Sub Form_Resize()
    Anclar fraBoton, Me, anclarAbajo + anclarIzquierda
    Anclar fraGrilla, Me, anclarLadosTodos
    Anclar grilla, fraGrilla, anclarLadosTodos
End Sub

'Private Sub ucXls1_Clic(Cancel As Boolean)
'    Dim p As String, fe As String
'    fe = " entre " & dtfechad & " y " & dtfechah & ""
'    If uProvD.codigo = uProvH.codigo Then
'        p = " para " & uProvD.descripcion
'    Else
'        p = "prov " & uProvD.descripcion & " a " & uProvH.descripcion
'    End If
'    ucXls1.aTitulo = "Listado mov cuenta proveedores " & p & fe
'End Sub

Private Sub ucXls1_Clic(cancel As Boolean)
    ucXls1.ini grilla2, "C:\CompProvSin.xls", "Composicion de proveedor sin detalle"
End Sub
