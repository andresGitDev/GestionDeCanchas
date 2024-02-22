VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListadoChequesTerceros 
   Caption         =   "Listado de Cheques de Terceros"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "FrmListadoChequesTerceros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
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
      Left            =   8610
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1785
      Width           =   1185
   End
   Begin VB.Frame fraFechas 
      ForeColor       =   &H00000000&
      Height          =   915
      Left            =   3930
      TabIndex        =   2
      Top             =   180
      Width           =   4575
      Begin VB.OptionButton optFeIng 
         Caption         =   "Fecha Ingreso"
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
         Left            =   255
         TabIndex        =   13
         Top             =   570
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optFeChe 
         Caption         =   "Fecha Cheque"
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
         Left            =   255
         TabIndex        =   12
         Top             =   225
         Width           =   1770
      End
      Begin MSComCtl2.DTPicker dtFDesde 
         Height          =   300
         Left            =   2865
         TabIndex        =   14
         Tag             =   "6"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   38052
      End
      Begin MSComCtl2.DTPicker dtFHasta 
         Height          =   300
         Left            =   2865
         TabIndex        =   15
         Tag             =   "7"
         Top             =   510
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62914561
         CurrentDate     =   38052
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
         Left            =   2220
         TabIndex        =   17
         Top             =   225
         Width           =   615
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
         Left            =   2220
         TabIndex        =   16
         Top             =   540
         Width           =   615
      End
   End
   Begin VB.Frame fraEstados 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   240
      TabIndex        =   1
      Tag             =   "0"
      Top             =   180
      Width           =   3585
      Begin VB.OptionButton OptSalidos 
         Caption         =   "Cheques de Terceros Entregados"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   3375
      End
      Begin VB.OptionButton optRechazados 
         Caption         =   "Cheques rechazados"
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
         Left            =   120
         TabIndex        =   11
         Tag             =   "4"
         Top             =   1125
         Width           =   2655
      End
      Begin VB.OptionButton optCartera 
         Caption         =   "Cheques en Cartera"
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
         Left            =   120
         TabIndex        =   10
         Tag             =   "3"
         Top             =   825
         Width           =   2655
      End
      Begin VB.OptionButton optTransferidos 
         Caption         =   "Cheques Entregados por OP"
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
         Left            =   120
         TabIndex        =   9
         Tag             =   "2"
         Top             =   510
         Width           =   3015
      End
      Begin VB.OptionButton optAcreditarse 
         Caption         =   "Cheques por Acreditarse"
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
         Left            =   120
         TabIndex        =   0
         Tag             =   "1"
         Top             =   210
         Width           =   2655
      End
      Begin VB.OptionButton optIngresados 
         Caption         =   "Cheques ingresados"
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
         TabIndex        =   8
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3930
      Left            =   105
      TabIndex        =   7
      Top             =   2850
      Width           =   9990
      _cx             =   17621
      _cy             =   6932
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
      FormatString    =   $"FrmListadoChequesTerceros.frx":08CA
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
      ExplorerBar     =   7
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
   Begin Gestion.ucCoDe uCliente 
      Height          =   345
      Left            =   4695
      TabIndex        =   3
      Top             =   1200
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   609
      CodigoWidth     =   1000
   End
   Begin Gestion.ucXls uXls 
      Height          =   375
      Left            =   5370
      TabIndex        =   6
      Top             =   1785
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ver "
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
      Left            =   4065
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "8"
      Top             =   1785
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Luego de ver los datos, presione sobre el nombre de la columna para ordenar o escriba sobre un reglon para buscar."
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   2445
      Width           =   9930
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   2265
      Left            =   90
      Top             =   135
      Width           =   9975
   End
   Begin VB.Label lblcliente 
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3930
      TabIndex        =   4
      Top             =   1245
      Width           =   735
   End
End
Attribute VB_Name = "FrmListadoChequesTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 16/9/4
' 14/7/6 lo Rehice.
    
    
Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
   
    Dim rs As New ADODB.Recordset, s As String
    
    'pal corte
    Dim Mes As Long, Anio As Long, rsMes  As Long, rsAnio As Long
    Dim Total As Double, TotalTotal As Double
    
    'pal where
    Dim campoFecha As String, queRangoFecha As String, queCliente As String, queEstado As String
    
    relojito
    InicioGrilla
    
    campoFecha = IIf(optFeIng, "FECHA_INGR", "FECHA")
    queRangoFecha = " and ( " & campoFecha & ssBetween(dtFDesde, dtFHasta) & " ) "
    queCliente = IIf(uCliente.codigo > 0, "and cheques.cliente = " & uCliente.codigo, "")
    
    queEstado = ""
    If optCartera = True Then queEstado = " and estado = 'C' "
    If optAcreditarse = True Then queEstado = " and estado = 'D' "
    If optRechazados = True Then queEstado = " and estado = 'R' "
    If optTransferidos Then queEstado = " and estado = 'T' "
    If OptSalidos = True Then queEstado = " and estado='S' "
    
    s = " select cheques.*," & _
        " BancosGrales.descripcion as NombreBanco, " & _
        " clientes.descripcion as NombreCliente, " & _
        " prov.descripcion as NombreProv " & _
        " FROM CHEQUES  " & _
        " LEFT OUTER JOIN BancosGrales ON CHEQUES.BANCO_NRO = BancosGrales.codigo " & _
        " LEFT OUTER JOIN  Clientes ON CHEQUES.CLIENTE = Clientes.codigo " & _
        " left outer join prov on cheques.prov = prov.codigo " & _
        " where cheques.activo = 1 " & queRangoFecha & queEstado & queCliente & _
        " order by  " & campoFecha
    
'    Debug.Print s
    With rs
    .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                  
    If Not .EOF Then
        Mes = Month(.Fields(campoFecha))
        Anio = Year(.Fields(campoFecha))
        
        While Not .EOF
            rsMes = Month(.Fields(campoFecha))
            rsAnio = Year(.Fields(campoFecha))
            
            If Not (rsMes = Mes) Or (rsMes = Mes And rsAnio <> Anio) Then
                Grilla.AddItem "Total " & ObtenerMes(Mes) & " " & Anio & ": " & Total
                Total = 0
            End If
                        
            Grilla.AddItem vbTab & _
                !Fecha & vbTab & _
                !NroInt & vbTab & _
                !nombreBanco & vbTab & _
                !Nro & vbTab & _
                !Importe & vbTab & _
                !nombreCliente & vbTab & _
                IIf(!procedencia = "T", "TERCEROS", "PROPIO") & vbTab & _
                !tdoc & vbTab & _
                !nDoc & vbTab & _
                !fecha_ingr & vbTab & _
                verEstado(!estado) & vbTab & _
                verDestino(!estado, sSinNull(!NombreProv), !NroInt) & vbTab & _
                Left(!tdocProv & "    ", 4) & IIf(s2n(!ndocProv) = 0, "", !ndocProv)

            Total = Total + s2n(!Importe)
            TotalTotal = TotalTotal + s2n(!Importe)
            
            Mes = rsMes
            Anio = rsAnio
            
            .MoveNext
        Wend
        
        Grilla.AddItem "Total " & ObtenerMes(Mes) & " " & Anio & ": " & Total
        Grilla.AddItem "Total Gral:         " & Format(TotalTotal, "standard")
    End If
    End With
    
fin:
    relojito False
    Set rs = Nothing
    Exit Sub
ufa:
    ufa "err listado", ""
    Resume fin
End Sub
Private Function verDestino(estado As String, Nombre As String, inter As Long)
    If estado = "C" Or estado = "T" Then
        verDestino = sSinNull(Nombre)
    ElseIf estado = "J" Then
        verDestino = obtenerDeSQL("select c.sector from movibanc b inner join movicaja m on m.iddoc=b.iddoc inner join cajas c on c.codigo=m.caja where b.activo=1 and b.operacion='" & estado & "' and b.interno=" & inter)
    Else
        verDestino = obtenerDeSQL("select b.descripcion from movibanc m inner join ctasbank c on c.codigo=m.cuenta inner join bancosgrales b on b.codigo=c.banco where m.activo=1 and m.operacion='" & estado & "' and m.interno=" & inter)
    End If
End Function
  
Private Function verEstado(cual As String) As String
    Dim s
    s = "C-Cartera       T-Transferido   R-Rechazado    A-Acreditado   D-Depositado    J-Canjeado     "
    verEstado = Mid(s, InStr(s, cual & "-") + 2, 12)
End Function


Function ObtenerMes(codigo As Long) As String
    Dim a
    a = Split("X ENE FEB MAR ABR MAY JUN JUL AGO SEP OCT NOV DIC")
    ObtenerMes = a(codigo)
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtFDesde = CDate("1/1/" & Year(Date))
    dtFHasta = Date
    InicioGrilla
    uCliente.ini "select descripcion from clientes where codigo = ### ", "select codigo as [ Codigo    ], descripcion as [ Cliemte                                              ] from clientes order by descripcion"
    uXls.ini Grilla, "C:\ListadoCheques", "Listado de Cheques"
    Form_Resize
    'Grilla.Editable = flexEDKbdMouse
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Resize()
    Anclar Grilla, Me, anclarLadosTodos
End Sub

Private Sub InicioGrilla()
    With Grilla
        .clear
        .cols = 14
        .rows = 1
        .FormatString = " | Fecha | NroInt | Banco | Nro Cheque | Importe| Cliente | Procedencia | Tipo | NroDoc| FechaIngreso | Estado | Destino | Documento "
        .ColFormat(5) = "#.00"
        .ColAlignment(5) = flexAlignRightCenter
    End With
    grillaWidth Grilla, Array(900, 1050, 750, 2370, 1080, 945, 1560, 1080, 570, 450, 1290)
End Sub
