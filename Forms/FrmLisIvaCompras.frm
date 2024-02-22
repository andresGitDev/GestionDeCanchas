VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLisIvaCompras 
   Caption         =   "Subdiario Compras"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   Icon            =   "FrmLisIvaCompras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4350
      TabIndex        =   11
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton optmes 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Mes de Imputacion"
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
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   90
         Width           =   2415
      End
      Begin VB.ComboBox cmbmes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisIvaCompras.frx":08CA
         Left            =   915
         List            =   "FrmLisIvaCompras.frx":08F2
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   495
         Width           =   1935
      End
      Begin VB.ComboBox cmbaño 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisIvaCompras.frx":095B
         Left            =   3405
         List            =   "FrmLisIvaCompras.frx":095D
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   495
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00400000&
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   765
         Left            =   105
         Top             =   180
         Width           =   4665
      End
      Begin VB.Label Label1 
         Caption         =   "Mes"
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
         Left            =   315
         TabIndex        =   16
         Top             =   510
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Año"
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
         Left            =   2925
         TabIndex        =   15
         Top             =   510
         Width           =   495
      End
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
      Height          =   375
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1095
      Width           =   975
   End
   Begin Gestion.ucXls uXls 
      Height          =   945
      Left            =   9210
      TabIndex        =   5
      Top             =   60
      Width           =   795
      _extentx        =   1402
      _extenty        =   1667
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   228065281
      CurrentDate     =   38252
   End
   Begin VB.OptionButton optfecha 
      Alignment       =   1  'Right Justify
      Caption         =   "Entre Fechas"
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
      Left            =   285
      TabIndex        =   0
      Top             =   45
      Value           =   -1  'True
      Width           =   1695
   End
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
      Height          =   375
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1065
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
      Height          =   375
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1095
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
      Height          =   375
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   300
      Left            =   2865
      TabIndex        =   2
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   228065281
      CurrentDate     =   38252
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3375
      Left            =   90
      TabIndex        =   10
      Top             =   1605
      Width           =   9855
      _cx             =   17383
      _cy             =   5953
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
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLisIvaCompras.frx":095F
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
   Begin VB.Label Label3 
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
      Left            =   2250
      TabIndex        =   9
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   165
      TabIndex        =   8
      Top             =   435
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   810
      Left            =   45
      Top             =   165
      Width           =   4260
   End
End
Attribute VB_Name = "FrmLisIvaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 5/10/4  '18/10/4  10/3/5  3/11/5

'le puse ON ERROR UFA

Private sTablaTemp As String
'Private Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoynro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto] [float] NULL , [ivaresp] [float] NULL , [rg3337] [float] NULL , [iva27] [float] NULL , [iva10] [float] NULL , [imptotal] [float] NULL , [retenciva] [float] NULL , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL, [letra] [varchar] (2) NULL)"
'Private Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoynro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto] [float] NULL , [ivaresp] [float] NULL , [rg3337] [float] NULL , [iva27] [float] NULL , [iva10] [float] NULL , [retenciva] [float] NULL , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL    , [imptotal] [float] NULL , [NoGrabado] [float] NULL     )"
Private Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoynro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto21] [float] NULL ,  [neto10] [float] NULL ,  [neto27] [float] NULL  , [iva21] [float] NULL ,  [iva10] [float] NULL , [iva27] [float] NULL  , [rg3337] [float] NUll , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL    , [imptotal] [float] NULL , [NoGrabado] [float] NULL     )"

Private Sub cmdAceptar_Click()
Dim Opp As Long
    If optFecha.Value = True Then
        Opp = 1
    Else
        Opp = 2
    End If
    CargoGC GRILLA, dtfechad, dtfechah, cmbmes.ListIndex + 1, val(Trim(cmbaño.Text)), Opp
End Sub

Public Function CargoGC(ByRef gg As Object, ByVal FechaD As Date, ByVal FechaH As Date, ByVal Mes As Long, ByVal Anio As Long, ByVal Opcion As Long)
    If ON_ERROR_HABILITADO Then On Error GoTo UFAlistado
    
    Dim str As String, rs As New ADODB.Recordset, Consulta As String, signo As Variant, ssql As String, scampos As String, sWhereDate As String, sWhere As String
    Dim cNeto21 As Double, cNeto10 As Double, cNeto27 As Double, cIva21 As Double, cIva10 As Double, cIva27 As Double, cRetGan As Double, i As Long
    relojito True

    sTablaTemp = TablaTempCrear(tt_iva_compras_temp)
    With rs
        If Opcion = 1 Then
            sWhereDate = " c.fecha " & ssBetween(FechaD, FechaH) & " "
        ElseIf Opcion = 2 Then
            sWhereDate = " c.mesimp=" & Mes & " and c.anoimp=" & Anio & " "
        End If
        sWhere = " where  " & sWhereDate & " and c.tipodoc in ('FAC','N/D','N/C') and c.activo=1"
        scampos = " c.Fecha,c.razonsocialprov,c.cuitprov,c.TIPODOC, i.letraprov, c.suc, c.NroDoc,c.percepc,c.Total,c.imp_int,c.der_est,c.EXENTO,c.ibcapital,c.ibprovincia,c.nogravado,c.iva_21,c.iva_10,c.iva_27,c.neto "
        ssql = " SELECT " & scampos & ", i.letraprov FROM TRANSCOM as c INNER JOIN Ivas as i ON c.TIPOIVA = i.codigo " & sWhere & _
                " union " & _
                " SELECT " & scampos & ", i.letraprov FROM COMPRAS  as c INNER JOIN Ivas as i ON c.TIPOIVA = i.codigo " & sWhere

       rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
       If Not rs.EOF And Not rs.BOF Then
           Do While Not rs.EOF
               signo = ""
               If Trim(rs!TIPODOC) = "N/C" Then
                   signo = "-"
               End If

               cIva21 = 0
               cIva10 = 0
               cIva27 = 0
               cNeto21 = 0
               cNeto10 = 0
               cNeto27 = 0
               cIva21 = s2n(nSinNull(!IVA_21))
               cIva10 = s2n(nSinNull(!iva_10))
               cIva27 = s2n(nSinNull(!IVA_27))
               If cIva21 + cIva10 + cIva27 = cIva21 Or cIva21 + cIva10 + cIva27 = cIva10 Or cIva21 + cIva10 + cIva27 = cIva27 Then
                   If cIva21 + cIva10 + cIva27 = cIva21 Then
                       cNeto21 = s2n(nSinNull(!Neto))
                       cNeto10 = 0
                       cNeto27 = 0
                   ElseIf cIva21 + cIva10 + cIva27 = cIva10 Then
                       cNeto21 = 0
                       cNeto10 = s2n(nSinNull(!Neto))
                       cNeto27 = 0
                   ElseIf cIva21 + cIva10 + cIva27 = cIva27 Then
                       cNeto21 = 0
                       cNeto10 = 0
                       cNeto27 = s2n(nSinNull(!Neto))
                   End If
               Else
                   cNeto21 = s2n((100 * nSinNull(!IVA_21)) / 21)
                   cNeto10 = s2n((100 * nSinNull(!iva_10)) / 10.5)
                   cNeto27 = s2n((100 * nSinNull(!IVA_27)) / 27)
               End If
               cRetGan = 0 's2n(!retgan)
               Consulta = "insert into " & sTablaTemp _
                   & " ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337, " _
                   & " imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values ( " _
                   & ssFecha(rs!Fecha) & ",'" & !razonsocialprov & "','" & !cuitprov & "','" _
                   & TipoyNro(!TIPODOC, !letraprov, !suc, !NroDoc) & "'," & signo & x2s(cNeto21) & ", " & signo & x2s(cNeto10) & " , " & signo & x2s(cNeto27) & " ," & signo & x2s(cIva21) & "," _
                   & signo & x2s(cIva10) & "," & signo & x2s(cIva27) & "," & signo & x2s(!percepc) & "," & signo & x2s(rs!Total) & "," & signo & x2s(!imp_int) _
                   & "," & signo & x2s(cRetGan) & "," & signo & x2s(!der_est) & "," & signo & x2s(!EXENTO) & "," & signo & x2s(rs!ibcapital) & "," & signo & x2s(rs!ibprovincia) & "," & x2s(IIf(signo = "-", -1, 1) * s2n(!nogravado)) & ")"
               
               '!retganpago
               
               DataEnvironment1.Sistema.Execute Consulta
               rs.MoveNext
           Loop
       End If
       Set rs = Nothing
        
    End With
    
    str = "select * from " & sTablaTemp & " order by fecha"
    RptIvaCompras.Data.Connection = DataEnvironment1.Sistema
    RptIvaCompras.Data.Source = str
    RptIvaCompras.lblfecha = Date
    RptIvaCompras.Printer.PaperSize = vbPRPSLegal 'hoja oficio
    scampos = " fecha as [Fecha], [razonsocial] as [Razon social] , [nrocuit] as [Nro CUIT] , [tipoynro] as [Documento],  [neto21] as [ Neto 21],  [neto10] as [ Neto 10.5],  [neto27] as [ Neto 27]   , [iva21] as  [IVA 21] , [iva10] as [IVA 10.5]  , [iva27] as [IVA 27], [rg3337] as [RG 3337], [IB_CAPITAL]+ [IB_PROVINCIA] as [       IIBB], [exento] as [ Exento],[Nograbado] as [No Grabado], [imptotal] as [          TOTAL     ]"
    
    LlenarGrilla gg, "select " & scampos & " from " & sTablaTemp & " order by fecha", False

    
    If gg.rows > 1 Then
        i = 1
        While i < gg.rows
            gg.TextMatrix(i, 4) = s2n(gg.TextMatrix(i, 4), 2, True)
            gg.TextMatrix(i, 5) = s2n(gg.TextMatrix(i, 5), 2, True)
            gg.TextMatrix(i, 6) = s2n(gg.TextMatrix(i, 6), 2, True)
            gg.TextMatrix(i, 7) = s2n(gg.TextMatrix(i, 7), 2, True)
            gg.TextMatrix(i, 8) = s2n(gg.TextMatrix(i, 8), 2, True)
            gg.TextMatrix(i, 9) = s2n(gg.TextMatrix(i, 9), 2, True)
            gg.TextMatrix(i, 10) = s2n(gg.TextMatrix(i, 10), 2, True)
            gg.TextMatrix(i, 11) = s2n(gg.TextMatrix(i, 11), 2, True)
            gg.TextMatrix(i, 12) = s2n(gg.TextMatrix(i, 12), 2, True)
            gg.TextMatrix(i, 13) = s2n(gg.TextMatrix(i, 13), 2, True)
            gg.TextMatrix(i, 14) = s2n(gg.TextMatrix(i, 14), 2, True)
            i = i + 1
        Wend
        gg.ColWidth(1) = 2900
        gg.ColWidth(2) = 1200
        gg.ColWidth(3) = 1800
        gg.ColWidth(4) = 1100
        gg.ColWidth(5) = 1100
        gg.ColWidth(8) = 1100
        gg.ColWidth(14) = 1100
        
        sumarizo gg, Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
    End If
    

    relojito False
Exit Function
UFAlistado:
    ufa "err prg", " --consulta=  " & Consulta & " -- " ', Err
End Function

Private Sub cmdCancelar_Click()
    dtfechad = Date
    dtfechah = Date
    cmbmes.Text = ""
    cmbaño.Text = ""
End Sub

Private Sub cmdImprimir_Click()
    Dim bb As Boolean
    Dim i As Long
    Dim j As Long
    Dim cant As Long
    Dim Arrai As Variant
    
    bb = confirma("imprime fecha de emision")
    i = 1
    cant = 36
    ReDim Arrai(10)
    While i < GRILLA.rows
        
        If i = cant Then
            'sumarizo grilla, Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
            
            GRILLA.AddItem "" & Chr(9) & "SUBTOTAL DE TRANSPORTE" & Chr(9) & "" & Chr(9) & "" & Chr(9) & s2n(Arrai(0), 2, True) & Chr(9) & s2n(Arrai(1), 2, True) & Chr(9) & s2n(Arrai(2), 2, True) & Chr(9) & _
                        s2n(Arrai(3), 2, True) & Chr(9) & s2n(Arrai(4), 2, True) & Chr(9) & s2n(Arrai(5), 2, True) & Chr(9) & s2n(Arrai(6), 2, True) & Chr(9) & s2n(Arrai(7), 2, True) & Chr(9) & s2n(Arrai(8), 2, True) & Chr(9) & s2n(Arrai(9), 2, True) & Chr(9) & s2n(Arrai(10), 2, True), i
            GRILLA.AddItem "" & Chr(9) & "SUBTOTAL DE TRANSPORTE" & Chr(9) & "" & Chr(9) & "" & Chr(9) & s2n(Arrai(0), 2, True) & Chr(9) & s2n(Arrai(1), 2, True) & Chr(9) & s2n(Arrai(2), 2, True) & Chr(9) & _
                        s2n(Arrai(3), 2, True) & Chr(9) & s2n(Arrai(4), 2, True) & Chr(9) & s2n(Arrai(5), 2, True) & Chr(9) & s2n(Arrai(6), 2, True) & Chr(9) & s2n(Arrai(7), 2, True) & Chr(9) & s2n(Arrai(8), 2, True) & Chr(9) & s2n(Arrai(9), 2, True) & Chr(9) & s2n(Arrai(10), 2, True), i + 1
            cant = cant + 38
            i = i + 2
        End If
        j = 0
        While j < 11
            Arrai(j) = s2n(Arrai(j), 2) + s2n(GRILLA.TextMatrix(i, 4 + j), 2)
            j = j + 1
        Wend
        i = i + 1
    Wend
    
    If optFecha.Value = True Then
        PrintG GRILLA, pHorizontal, "IVA COMPRAS", IIf(bb, Date, "01/01/1900"), "SUBDIARIO COMPRAS DESDE : " & dtfechad.Value & " AL " & dtfechah.Value, vbPRPSLegal
    Else
        PrintG GRILLA, pHorizontal, "IVA COMPRAS", IIf(bb, Date, "01/01/1900"), "SUBDIARIO COMPRAS DEl MES : " & cmbmes.Text & " DEL " & cmbaño.Text, vbPRPSLegal
    End If
    
    i = 1
    While i < GRILLA.rows
        If Trim(GRILLA.TextMatrix(i, 1)) = "SUBTOTAL DE TRANSPORTE" Then
            GRILLA.RemoveItem i
            i = i - 1
        End If
        i = i + 1
    Wend
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub dtfechad_LostFocus()
    dtfechah = CDate(DateSerial(Year(dtfechad), Month(dtfechad) + 1, 0))
End Sub

Private Sub Form_Load()
    Dim d As Long, h As Long, i As Long
    h = Year(Date)
    d = Year(Date) - 6
    For i = d To h
        cmbaño.AddItem i
    Next
    
    dtfechad = Date - Day(Date) + 1
    dtfechah = Date
    cmbmes.Text = ""
    cmbaño.Text = ""
    If gEMPR_idEmpresa <> 2 Then
'       Frame1.Visible = True
    Else
       Frame1.Visible = False
    End If
    '   optmes.Value = True
    'End If
    uXls.ini GRILLA, "c:\SubdCompras", "Subdiario de compras " & dtfechad & "  -  " & dtfechah
    Form_Resize
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Resize()
    Anclar GRILLA, Me, anclarLadosTodos
End Sub

''Private Sub Form_Unload(cancel As Integer)
''    TablaTempBorrar sTablaTemp
''End Sub

Private Sub optfecha_Click()
    If optFecha.Value = True Then
        cmbmes.enabled = False
        cmbaño.enabled = False
        dtfechad.enabled = True
        dtfechah.enabled = True
        optmes.Value = False
    Else
        If optmes.Value = True Then
            cmbmes.enabled = True
            cmbaño.enabled = True
            dtfechad.enabled = False
            dtfechah.enabled = False
        End If
    End If
End Sub

Private Sub optmes_Click()
    If optmes.Value = True Then
        cmbmes.enabled = True
        cmbaño.enabled = True
        dtfechad.enabled = False
        dtfechah.enabled = False
        optFecha.Value = False
    Else
        If optFecha.Value = True Then
            cmbmes.enabled = False
            cmbaño.enabled = False
            dtfechad.enabled = True
            dtfechah.enabled = True
            optmes.Value = False
        End If
    End If
End Sub

Private Sub uXls_Clic(cancel As Boolean)
    uXls.aTitulo = "Subdiario Compras " & dtfechad & "  -  " & dtfechah
End Sub
Public Sub sumarizo(GRILLA, a)
    Dim i As Long
    With GRILLA
        .SubtotalPosition = flexSTBelow
        For i = 0 To UBound(a):        .subtotal flexSTSum, -1, a(i), , , , True, , , True: Next
'        .TextMatrix(.rows - 1, 0) = " Totales"
    End With
End Sub

Public Function TipoyNro(tdoc As String, letra As String, suc As Long, Nro As Long) As String
    TipoyNro = Left(tdoc & " " & letra & "   ", 6) & Format(suc, CuantosDigitosPV()) & "-" & Format(Nro, "00000000")
End Function

Public Function TYN(suc As Long, Nro As Long) As String
    TYN = Format(suc, CuantosDigitosPV()) & "-" & Format(Nro, "00000000")
End Function

'10/3/5 tabla temp
'  /9/5 agergado  "and ivas.letra = 'A'" en los SQL
'?/11/5  SACADO iva.letra de listado OJO arreglar brehm
'

