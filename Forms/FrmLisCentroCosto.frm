VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmLisCentroCosto 
   Caption         =   "Listado de centro de costos"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12420
   Icon            =   "FrmLisCentroCosto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12420
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   855
      Left            =   720
      TabIndex        =   16
      Top             =   6360
      Width           =   855
      _extentx        =   1508
      _extenty        =   1508
   End
   Begin Gestion.ucCoDe ucCoDeCosto 
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Top             =   720
      Width           =   4455
      _extentx        =   7858
      _extenty        =   503
      codigowidth     =   1000
   End
   Begin Gestion.ucEntreFechas ucEntreFechas1 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   3255
      _extentx        =   5741
      _extenty        =   503
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4890
      TabIndex        =   13
      Top             =   6870
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6675
      TabIndex        =   12
      Top             =   6870
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   330
      Left            =   8460
      TabIndex        =   8
      Top             =   6345
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8295
      TabIndex        =   4
      Top             =   6885
      Width           =   1305
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3090
      TabIndex        =   3
      Top             =   6870
      Width           =   1305
   End
   Begin VB.TextBox TotalNeto 
      Enabled         =   0   'False
      Height          =   330
      Left            =   2880
      TabIndex        =   1
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox TotalIVA 
      Enabled         =   0   'False
      Height          =   330
      Left            =   5760
      TabIndex        =   0
      Top             =   6360
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla 
      Height          =   4800
      Left            =   105
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1260
      Width           =   12195
      _cx             =   21511
      _cy             =   8467
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLisCentroCosto.frx":08CA
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de costo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6075
      TabIndex        =   11
      Top             =   735
      Width           =   1425
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde - Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   705
      TabIndex        =   10
      Top             =   705
      Width           =   1410
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   7860
      TabIndex        =   9
      Top             =   6375
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione los Datos a Ver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   165
      TabIndex        =   7
      Top             =   150
      Width           =   3210
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Neto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   6
      Top             =   6360
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4830
      TabIndex        =   5
      Top             =   6375
      Width           =   915
   End
End
Attribute VB_Name = "FrmLisCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sql As String
Dim Consulta As String
'Private rsCost As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Dim i As Long
    
    limpio
    'grilla.Clear
    ucCoDeCosto.codigo = ""
    ucCoDeCosto.DESCRIPCION = ""
End Sub
Public Sub limpio()
    Dim i As Long
    
    For i = 1 To grilla.rows - 1
        grilla.RemoveItem (1)
    Next i
    
    TotalNeto.Text = ""
    TotalIva.Text = ""
    txttotal.Text = ""
    
End Sub

Private Sub CmdEjecutar_Click()
    Dim TNeto As Double, TIVA As Double, TIMPO As Double
    Dim PNeto As Double, PIVA As Double, PIMPO As Double
    Dim netoD, netoH, ivaD, ivaH
    Dim tdoc
    Dim whe
    Dim orden
    Dim desc As String, cod As String, descrip As String
    Dim v As Double
    
    Dim rsCost As New ADODB.Recordset
    Dim rsCost2 As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    'Set rsCost = Nothing

    
    If ucCoDeCosto.codigo = "" Or ucCoDeCosto.DESCRIPCION = "" Then
        MsgBox "Debe ingresar un centro de costo."
        Exit Sub
    End If
    
    limpio
    
    'Dim signo As Variant
    'Dim sTmpIvaVentas As String
    
    'sTmpIvaVentas = TablaTempCrear(tt_IIBB)
    'ControlPrevioProv
    If ucCoDeCosto.codigo = "" Or ucCoDeCosto.codigo = 0 Then
        whe = ""
    Else
        whe = " and c.codigo=" & ucCoDeCosto.codigo
    End If
    
    orden = " order by d.tdoc,c.codigo"
        
    Consulta = "select d.codigo,c.descripcion,d.fecha,d.tdoc,d.ndoc,d.neto,d.ivapesos,d.importe,d.motivo,d.prov from " & _
                " detallecentrocostos D inner join centrodecostos C on D.codigo=c.codigo  where d.activo=1 and FECHA " & ssBetween(ucEntreFechas1.desde, ucEntreFechas1.hasta) & whe & orden
    
    rsCost.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    'If Not rs.EOF Then
    '    rs.MoveFirst
    '    While Not rs.EOF
    '        signo = ""
    '        If Trim(rs!TIPODOC) = "NCA" Or Trim(rs!TIPODOC) = "NCB" Then
    '            signo = "-"
    '        End If
    '        Consulta = "Insert Into " & sTmpIvaVentas & " (provincia,neto,iva) " & _
    '        "Values ( '" & rs!descripcion & "' , " & signo & x2s(rs!Neto) & ", " & signo & x2s(nSinNull(rs!Iva)) & ")"
    '        DataEnvironment1.Sistema.Execute Consulta
    '        rs.MoveNext
    '    Wend
    'End If
    
    'rs.Close
    'Set rs = Nothing
                                                                                
    'STR = "SELECT Provincia, SUM(Neto) AS SumaNeto, SUM(Iva) " & _
    '"AS SumaIVA FROM " & sTmpIvaVentas & " GROUP BY Provincia ORDER BY provincia"
    'rs.Open STR, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    
    Do While Not rsCost.EOF
        'If rsCost!nDoc = 2769 Then
        '    MsgBox ""
        'End If
        desc = "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc
        rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsCost2.EOF = True And rsCost2.BOF = True Then
            Set rsCost2 = Nothing
            desc = "select * from compras where tipodoc='" & rsCost!tdoc & "' and nrodoc=" & rsCost!nDoc & " and codpr=" & rsCost!prov
            rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If rsCost2.EOF = True And rsCost2.BOF = True Then
                Set rsCost2 = Nothing
                desc = "select * from transcom where tipodoc='" & rsCost!tdoc & "' and nrodoc=" & rsCost!nDoc & " and codpr=" & rsCost!prov
                rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If rsCost2.EOF = True And rsCost2.BOF = True Then
                    Set rsCost2 = Nothing
                    cod = " - "
                    descrip = rsCost!motivo '" - "
                Else
                    cod = rsCost2!CODPR
                    descrip = rsCost2!razonsocialprov
                End If
            Else
                cod = rsCost2!CODPR
                descrip = rsCost2!razonsocialprov
            End If
        Else
            cod = rsCost2!cliente
            descrip = rsCost2!razonsocial
        End If
        Set rsCost2 = Nothing
        
        'PNeto = PNeto + Format$(rsCost!Neto, "standard")
        'PIVA = PIVA + Format$(rsCost!ivapesos, "standard")
        'PIMPO = PIMPO + Format$(rsCost!Importe, "standard")
        tdoc = rsCost!tdoc
        'rs.MoveNext
        
        '********************************************************
            
        If rsCost.EOF = True Then
        Else
            Select Case rsCost!tdoc
                Case "AJC": 'ajuste de credito
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    v = 1
                Case "AJD": 'ajuste de debito
                    TNeto = TNeto - Format$(rsCost!Neto, "standard")
                    TIVA = TIVA - Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO - Format$(rsCost!importe, "standard")
                    netoD = 0 '" "
                    netoH = Format$(rsCost!Neto, "standard")
                    ivaD = 0 '" "
                    ivaH = Format$(rsCost!ivapesos, "standard")
                    v = 1
                Case "FAC": 'factura de compra
                    TNeto = TNeto - Format$(rsCost!Neto, "standard")
                    TIVA = TIVA - Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO - Format$(rsCost!importe, "standard")
                    netoD = 0 '" "
                    netoH = Format$(rsCost!Neto, "standard")
                    ivaD = 0 '" "
                    ivaH = Format$(rsCost!ivapesos, "standard")
                    v = 1
                Case "NCC": 'nota de credito de compra
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    v = 1
                Case "NDC": 'nota de debito de compra
                    TNeto = TNeto - Format$(rsCost!Neto, "standard")
                    TIVA = TIVA - Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO - Format$(rsCost!importe, "standard")
                    netoD = 0 '" "
                    netoH = Format$(rsCost!Neto, "standard")
                    ivaD = 0 '" "
                    ivaH = Format$(rsCost!ivapesos, "standard")
                    v = 1
                Case "FAV": 'factura de venta
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    Set rs = Nothing
                Case "FAA": 'factura de venta
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    Set rs = Nothing
                Case "FAB": 'factura de venta
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    Set rs = Nothing
                Case "FAE": 'factura de exterior
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto + Format$(rsCost!Neto * v, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos * v, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe * v, "standard")
                    netoD = Format$(rsCost!Neto * v, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos * v, "standard")
                    ivaH = 0 '" "
                    Set rs = Nothing
                Case "NCA": 'nota de credito por devolucion en venta
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto - Format$(rsCost!Neto, "standard")
                    TIVA = TIVA - Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO - Format$(rsCost!importe, "standard")
                    netoD = 0 '" "
                    netoH = Format$(rsCost!Neto, "standard")
                    ivaD = 0 '" "
                    ivaH = Format$(rsCost!ivapesos, "standard")
                    Set rs = Nothing
                Case "NC", "NCE":   'nota de credito en venta por cliente
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto - Format$(rsCost!Neto, "standard")
                    TIVA = TIVA - Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO - Format$(rsCost!importe, "standard")
                    netoD = 0 '" "
                    netoH = Format$(rsCost!Neto, "standard")
                    ivaD = 0 '" "
                    ivaH = Format$(rsCost!ivapesos, "standard")
                    Set rs = Nothing
                Case "ND":  'nota de debito en venta por cliente
                    rs.Open "select * from facturaventa where tipodoc='" & rsCost!tdoc & "' and nrofactura=" & rsCost!nDoc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rs!moneda = 0 Or rs!moneda = 1 Then
                        v = 1
                    Else
                        v = rs!cotizacion
                    End If
                    
                    TNeto = TNeto + Format$(rsCost!Neto, "standard")
                    TIVA = TIVA + Format$(rsCost!ivapesos, "standard")
                    TIMPO = TIMPO + Format$(rsCost!importe, "standard")
                    netoD = Format$(rsCost!Neto, "standard")
                    netoH = 0 '" "
                    ivaD = Format$(rsCost!ivapesos, "standard")
                    ivaH = 0 '" "
                    Set rs = Nothing
            End Select
        End If
        
        PNeto = PNeto + Format$(rsCost!Neto * v, "standard")
        PIVA = PIVA + Format$(rsCost!ivapesos * v, "standard")
        PIMPO = PIMPO + Format$(rsCost!importe * v, "standard")
        
        grilla.AddItem cod & Chr(9) & descrip & Chr(9) & Format(rsCost!fecha, "DD/MM/YYYY") & Chr(9) & rsCost!tdoc & Chr(9) & rsCost!nDoc & Chr(9) & Format$((netoD), "standard") & Chr(9) & Format$((Format$((netoH), "standard") - (2 * Format$((netoH), "standard"))), "standard") & Chr(9) & Format$(ivaD, "standard") & Chr(9) & Format$((Format$(ivaH, "standard") - (2 * Format$(ivaH, "standard"))), "standard") & Chr(9) & IIf((Format$(ivaH, "standard") - (2 * Format$(ivaH, "standard"))) < 0, Format$((Format$(rsCost!importe * v, "standard") - (2 * Format$(rsCost!importe * v, "standard"))), "standard"), Format$(rsCost!importe * v, "standard"))
        
        rsCost.MoveNext
        
        If rsCost.EOF = True Then
            'Grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & Format$(PIVA, "standard") & Chr(9) & Format$(PIMPO, "standard")
            '**************************************************
            Select Case tdoc
                Case "AJC": 'ajuste de credito >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "AJD": 'ajuste de debito >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "FAC": 'factura de compra >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "NCC": 'nota de credito de compra >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "NDC": 'nota de debito de compra >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "FAV": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAA": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAB": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "NCA": 'nota de credito por devolucion en venta >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "NC":  'nota de credito en venta por cliente >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "ND":  'nota de debito en venta por cliente >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAE": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
            End Select
            '************************************************************************
'            g.tx i, gTOTA, dato
            PNeto = 0
            PIVA = 0
            PIMPO = 0
        ElseIf Not rsCost!tdoc = tdoc Then
            'Grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & Format$(PIVA, "standard") & Chr(9) & Format$(PIMPO, "standard")
            '*******************************************
            Select Case tdoc
                Case "AJC": 'ajuste de credito >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "AJD": 'ajuste de debito >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "FAC": 'factura de compra >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "NCC": 'nota de credito de compra >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "NDC": 'nota de debito de compra >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "FAV": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAA": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAB": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "NCA": 'nota de credito por devolucion en venta >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "NC":  'nota de credito en venta por cliente >>>>>haber
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$((Format$(PNeto, "standard") - (2 * Format$(PNeto, "standard"))), "standard") & Chr(9) & "" & Chr(9) & Format$((Format$(PIVA, "standard") - (2 * Format$(PIVA, "standard"))), "standard") & Chr(9) & Format$((Format$(PIMPO, "standard") - (2 * Format$(PIMPO, "standard"))), "standard")
                Case "ND":  'nota de debito en venta por cliente >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
                Case "FAE": 'factura de venta >>>>>debe
                    grilla.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(PNeto, "standard") & Chr(9) & "" & Chr(9) & Format$(PIVA, "standard") & Chr(9) & "" & Chr(9) & Format$(PIMPO, "standard")
            End Select
            '**************************************************************************
            
            'g.tx i, gTOTA, dato
            PNeto = 0
            PIVA = 0
            PIMPO = 0
        End If
            '    Wend
            '    .Close
            'End With
            
            'Label5.caption = importeTot
            
            
        '**********************************************************
    Loop
    Set rsCost = Nothing
    'rs.Close
    'TotalNeto = Format(TNeto, "standard")
    'TotalIVA = Format(TIVA, "standard")
    
    TotalNeto.Text = TNeto
    TotalIva.Text = TIVA
    txttotal.Text = TIMPO
    
End Sub

Private Sub cmdImprimir_Click()
    Dim sql As String
    Dim desc As String
    Dim descrip As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rsCost2 As New ADODB.Recordset
    Dim v As Double
    
    If grilla.rows > 1 Then
        'rptListCosto.DataC.Connection = DataEnvironment1.Sistema
        'rptListCosto.DataC.Source = Consulta
        
        '**************************************************************************
        Set rs2 = New ADODB.Recordset

        With rs2
        '>>>>>>>>> agregar dos campos mas por el debe y haber y en reporte tambien!! haber-restan
            .Fields.Append "tdoc", adChar, 4, adFldUpdatable
            .Fields.Append "Descripcion", adVarChar, 100, adFldUpdatable
            .Fields.Append "fecha", adVarChar, 10, adFldUpdatable
            .Fields.Append "netoD", adDouble, 10, adFldUpdatable
            .Fields.Append "netoH", adDouble, 10, adFldUpdatable '1
            .Fields.Append "ivapesosD", adDouble, 10, adFldUpdatable
            .Fields.Append "ivapesosH", adDouble, 10, adFldUpdatable '2
            .Fields.Append "importe", adDouble, 10, adFldUpdatable
            ' Utilice el tipo de cursor Keyset para permitir la actualización
            ' de los registros.
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open
        End With
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
        With rs2
            Do While Not rs.EOF
                
                desc = "select * from facturaventa where tipodoc='" & rs!tdoc & "' and nrofactura=" & rs!nDoc
                rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If rsCost2.EOF = True And rsCost2.BOF = True Then
                    Set rsCost2 = Nothing
                    desc = "select * from compras where tipodoc='" & rs!tdoc & "' and nrodoc=" & rs!nDoc & " and codpr=" & rs!prov
                    rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If rsCost2.EOF = True And rsCost2.BOF = True Then
                        Set rsCost2 = Nothing
                        desc = "select * from transcom where tipodoc='" & rs!tdoc & "' and nrodoc=" & rs!nDoc & " and codpr=" & rs!prov
                        rsCost2.Open desc, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                        If rsCost2.EOF = True And rsCost2.BOF = True Then
                            Set rsCost2 = Nothing
                            descrip = " - "
                        Else
                            descrip = rsCost2!razonsocialprov
                        End If
                    Else
                        descrip = rsCost2!razonsocialprov
                    End If
                    v = 1
                Else
                    descrip = rsCost2!razonsocial
                    If rsCost2!moneda = 0 Or rsCost2!moneda = 1 Then
                        v = 1
                    Else
                        v = rsCost2!cotizacion
                    End If
                End If
                Set rsCost2 = Nothing
                
                .AddNew
                !tdoc = rs!tdoc
                !DESCRIPCION = descrip
                !fecha = rs!fecha

                If rs.EOF = True Then
                Else
                    Select Case rs!tdoc
                        Case "AJC": 'ajuste de credito
                            '>>> debe +
                            !netoD = rs!Neto
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos
                            !ivapesosH = 0 '" "
                            !importe = rs!importe
                        Case "AJD": 'ajuste de debito
                            '>>> haber -
                            !netoD = 0 '" "
                            !netoH = rs!Neto - (2 * rs!Neto)
                            !ivapesosD = 0 '" "
                            !ivapesosH = rs!ivapesos - (2 * rs!ivapesos)
                            !importe = rs!importe - (2 * rs!importe)
                        Case "FAC": 'factura de compra
                            '>>> haber -
                            !netoD = 0 '" "
                            !netoH = rs!Neto - (2 * rs!Neto)       'asi paso a negativo
                            !ivapesosD = 0 '" "
                            !ivapesosH = rs!ivapesos - (2 * rs!ivapesos)
                            !importe = rs!importe - (2 * rs!importe)
                        Case "NCC": 'nota de credito de compra
                            '>>> debe +
                            !netoD = rs!Neto
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos
                            !ivapesosH = 0 '" "
                            !importe = rs!importe
                        Case "NDC": 'nota de debito de compra
                            '>>> haber -
                            !netoD = 0 '" "
                            !netoH = rs!Neto - (2 * rs!Neto)
                            !ivapesosD = 0 '" "
                            !ivapesosH = rs!ivapesos - (2 * rs!ivapesos)
                            !importe = rs!importe - (2 * rs!importe)
                        Case "FAV": 'factura de venta
                            '>>> debe +
                            !netoD = rs!Neto * v
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos * v
                            !ivapesosH = 0 '" "
                            !importe = rs!importe * v
                        Case "FAA": 'factura de venta
                            '>>> debe +
                            !netoD = rs!Neto * v
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos * v
                            !ivapesosH = 0 '" "
                            !importe = rs!importe * v
                        Case "FAB": 'factura de venta
                            '>>> debe +
                            !netoD = rs!Neto * v
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos * v
                            !ivapesosH = 0 '" "
                            !importe = rs!importe * v
                        Case "FAE": 'factura de exterior
                            '>>> debe +
                            !netoD = rs!Neto * v
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos * v
                            !ivapesosH = 0 '" "
                            !importe = rs!importe * v
                        Case "NCA": 'nota de credito por devolucion en venta
                            '>>> haber -
                            !netoD = 0 '" "
                            !netoH = (rs!Neto * v) - (2 * (rs!Neto * v))
                            !ivapesosD = 0 '" "
                            !ivapesosH = (rs!ivapesos * v) - (2 * (rs!ivapesos * v))
                            !importe = (rs!importe * v) - (2 * (rs!importe * v))
                        Case "NC", "NCE": 'nota de credito en venta por cliente
                            '>>> haber -
                            !netoD = 0 '" "
                            !netoH = (rs!Neto * v) - (2 * (rs!Neto * v))
                            !ivapesosD = 0 '" "
                            !ivapesosH = (rs!ivapesos * v) - (2 * (rs!ivapesos * v))
                            !importe = (rs!importe * v) - (2 * (rs!importe * v))
                        Case "ND":  'nota de debito en venta por cliente
                            '>>> debe +
                            !netoD = rs!Neto * v
                            !netoH = 0 '" "
                            !ivapesosD = rs!ivapesos * v
                            !ivapesosH = 0 '" "
                            !importe = rs!importe * v
                    End Select
                End If
                '!NetoD = rs!Neto
                '!netoH=
                '!ivapesosD=
                '!ivapesosH = rs!ivapesos
                '!Importe = rs!Importe 's2n()
                .Update
                '.Bookmark = .LastModified
                
                rs.MoveNext
            Loop
        End With
        
        rs2.MoveFirst
        Set rptListCosto.DataC.Recordset = rs2
        '*************************************************************************
        
        rptListCosto.Label10 = " Nº " & Format(ucCoDeCosto.codigo, "0000")
        rptListCosto.Label11 = ucCoDeCosto.DESCRIPCION
        
        Set rs = Nothing
        sql = "select descripcion,cliente,orden_compra, presupuestopesos,num_presupuesto from " & _
                " centrodecostos where activo=1 and codigo=" & ucCoDeCosto.codigo
                
        rs.Open sql, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
        If IsNull(rs!cliente) Then
            rptListCosto.Label13 = " - "
        Else
            rptListCosto.Label13 = rs!cliente
        End If
        If IsNull(rs!orden_compra) Then
            rptListCosto.Label15 = " - "
        Else
            rptListCosto.Label15 = rs!orden_compra
        End If
        If IsNull(rs!num_presupuesto) Then
            rptListCosto.Label17 = " - "
        Else
            rptListCosto.Label17 = rs!num_presupuesto
        End If
        If IsNull(rs!PresupuestoPESOS) Then
            rptListCosto.Label19 = " - "
        Else
            rptListCosto.Label19 = rs!PresupuestoPESOS
        End If
        
        'rptListCosto.TotalNeto.Text = Format$(TotalNeto, "standard")
        'rptListCosto.TotalIva.Text = Format$(TotalIva, "standard")
        rptListCosto.TotalImpo.Text = Format$(txttotal, "standard")
        rptListCosto.fechadesde.Text = ucEntreFechas1.desde
        rptListCosto.fechahasta.Text = ucEntreFechas1.hasta
        rptListCosto.Show vbModal
        
        Set rs = Nothing
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Cargogrilla
    ucCoDeCosto.ini "Select DESCRIPCION from centrodecostos Where CODIGO = '###'", _
                        "Select CODIGO, DESCRIPCION From centrodecostos Where ACTIVO = 1", False
    ucXls1.ini grilla, "CentroCosto.xls"
End Sub

Sub Cargogrilla()
    grilla.clear
    grilla.rows = 1
    grilla.TextMatrix(0, 0) = " Cod cliente/proveedor "
    grilla.TextMatrix(0, 1) = " Descripcion cliente/proveedor"
    grilla.TextMatrix(0, 2) = " Fecha "
    grilla.TextMatrix(0, 3) = " Tipo de documento "
    grilla.TextMatrix(0, 4) = " Numero de documento "
    grilla.TextMatrix(0, 5) = " Neto DEBE"
    grilla.TextMatrix(0, 6) = " Neto HABER"
    grilla.TextMatrix(0, 7) = " Iva en pesos DEBE"
    grilla.TextMatrix(0, 8) = " Iva en pesos HABER"
    grilla.TextMatrix(0, 9) = " Importe "
    grilla.ColWidth(0) = 1800
    grilla.ColWidth(1) = 2400
    grilla.ColWidth(2) = 1100
    grilla.ColWidth(3) = 1600
    grilla.ColWidth(4) = 1900
    grilla.ColWidth(5) = 1500
    grilla.ColWidth(6) = 1500
    grilla.ColWidth(7) = 1500
    grilla.ColWidth(8) = 1500
    grilla.ColWidth(9) = 1500

End Sub


