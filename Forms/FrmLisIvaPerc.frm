VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisIvaPerc 
   Caption         =   "Listado de percepcion de iva"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
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
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1095
      Width           =   975
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
      Left            =   8205
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1065
      Width           =   975
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
      Left            =   360
      TabIndex        =   9
      Top             =   45
      Value           =   -1  'True
      Width           =   1695
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
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1095
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cmbaño 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisIvaPerc.frx":0000
         Left            =   3405
         List            =   "FrmLisIvaPerc.frx":0002
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   495
         Width           =   1335
      End
      Begin VB.ComboBox cmbmes 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisIvaPerc.frx":0004
         Left            =   915
         List            =   "FrmLisIvaPerc.frx":002C
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   495
         Width           =   1935
      End
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
         TabIndex        =   1
         Top             =   90
         Width           =   2415
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
         TabIndex        =   5
         Top             =   510
         Width           =   495
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
         TabIndex        =   4
         Top             =   510
         Width           =   495
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
   End
   Begin Gestion.ucXls uXls 
      Height          =   945
      Left            =   9285
      TabIndex        =   7
      Top             =   60
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1667
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   300
      Left            =   915
      TabIndex        =   8
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   161939457
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   300
      Left            =   2940
      TabIndex        =   13
      Top             =   465
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   161939457
      CurrentDate     =   38252
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3375
      Left            =   165
      TabIndex        =   14
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
      FormatString    =   $"FrmLisIvaPerc.frx":0095
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
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   810
      Left            =   120
      Top             =   165
      Width           =   4260
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
      Left            =   240
      TabIndex        =   16
      Top             =   435
      Width           =   1335
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
      Left            =   2325
      TabIndex        =   15
      Top             =   465
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLisIvaPerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 5/10/4  '18/10/4  10/3/5  3/11/5

Private sTablaTemp As String
Private Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoynro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto21] [float] NULL ,  [neto10] [float] NULL ,  [neto27] [float] NULL  , [iva21] [float] NULL ,  [iva10] [float] NULL , [iva27] [float] NULL  , [rg3337] [float] NUll , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL    , [imptotal] [float] NULL , [NoGrabado] [float] NULL     )"

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAlistado
    
    Dim str As String
    Dim rs As New ADODB.Recordset
    Dim Consulta As String
    Dim signo As Variant
    Dim ssql As String
    Dim asse As String
    Dim scampos As String
    Dim cNeto21 As Double, cNeto10 As Double, cNeto27 As Double, cIva21 As Double, cIva10 As Double, cIva27 As Double, no_grav As Double
    Dim cRetGan As Double
    Dim i As Long
    relojito True

    sTablaTemp = TablaTempCrear(tt_iva_compras_temp)
    With rs
    If optfecha.Value = True Then
        
        asse = "transcom1"
        ssql = "SELECT TRANSCOM.*, Ivas.letraprov FROM TRANSCOM INNER JOIN Ivas ON TRANSCOM.TIPOIVA = Ivas.codigo where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and transcom.activo=1 and (TRANSCOM.percepc<>0 or transcom.der_est<>0)"
        rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
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
                cRetGan = 0
                Consulta = "insert into " & sTablaTemp _
                    & " ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337, " _
                    & " imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values ( " _
                    & ssFecha(rs!Fecha) & ",'" & !razonsocialprov & "','" & !cuitprov & "','" _
                    & TipoyNro(!TIPODOC, !letraprov, !suc, !NroDoc) & "'," & signo & x2s(cNeto21) & ", " & signo & x2s(cNeto10) & " , " & signo & x2s(cNeto27) & " ," & signo & x2s(cIva21) & "," _
                    & signo & x2s(cIva10) & "," & signo & x2s(cIva27) & "," & signo & x2s(!percepc) & "," & signo & x2s(rs!Total) & "," & signo & x2s(!imp_int) _
                    & "," & signo & x2s(cRetGan) & "," & signo & x2s(!der_est) & "," & signo & x2s(!EXENTO) & "," & signo & x2s(rs!ibcapital) & "," & signo & x2s(rs!ibprovincia) & "," & signo & x2s(no_grav) & ")"
                
                DataEnvironment1.Sistema.Execute Consulta
                'DataEnvironment1.Sistema.Execute "insert into #TmpTbl1209124835877 ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337,  imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values (  '20151118' ,'BGH S.A.','30-50361289-1','N/C A 0127-00042725',-0, -6683.52 , -0 ,-0,-701.81,-0,-200.51,-7819.8,-0,-0,-0,-0,-233.97,-0,-0.01)"
                rs.MoveNext
            Loop
        End If
        rs.Close
        
        asse = "compras1"
        ssql = " SELECT COMPRAS.*, Ivas.letraprov FROM COMPRAS INNER JOIN Ivas ON COMPRAS.TIPOIVA = Ivas.codigo where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and compras.activo=1 and (compras.percepc<>0 or compras.der_est<>0)"
        rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
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
                cRetGan = 0
                no_grav = s2n(!nogravado)
                If no_grav < 0 Then
                    no_grav = -no_grav
                    'Stop
                End If
                
                Consulta = "insert into " & sTablaTemp _
                    & " ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337, " _
                    & " imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values ( " _
                    & ssFecha(rs!Fecha) & ",'" & !razonsocialprov & "','" & !cuitprov & "','" _
                    & TipoyNro(!TIPODOC, !letraprov, !suc, !NroDoc) & "'," & signo & x2s(cNeto21) & ", " & signo & x2s(cNeto10) & " , " & signo & x2s(cNeto27) & " ," & signo & x2s(cIva21) & "," _
                    & signo & x2s(cIva10) & "," & signo & x2s(cIva27) & "," & signo & x2s(!percepc) & "," & signo & x2s(rs!Total) & "," & signo & x2s(!imp_int) _
                    & "," & signo & x2s(cRetGan) & "," & signo & x2s(!der_est) & "," & signo & x2s(!EXENTO) & "," & signo & x2s(rs!ibcapital) & "," & signo & x2s(rs!ibprovincia) & "," & signo & x2s(no_grav) & ")"
                DataEnvironment1.Sistema.Execute Consulta
                rs.MoveNext
            Loop
        End If
        rs.Close
        RptIvaCompras.lblTitulo = "Listado de percepcion de iva del " & CStr(dtfechad) & " al " & CStr(dtfechah)
    Else
        asse = "transcom2"
        If optmes.Value = True Then
          ssql = "SELECT TRANSCOM.*, Ivas.letraprov FROM TRANSCOM INNER JOIN Ivas ON TRANSCOM.TIPOIVA = Ivas.codigo where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and transcom.activo=1 and (TRANSCOM.percepc<>0 or transcom.der_est<>0)"
         rs.Open ssql, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
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
                        cNeto21 = s2n((100 * cIva21) / 21)
                        cNeto10 = s2n((100 * cIva10) / 10.5)
                        cNeto27 = s2n((100 * cIva27) / 27)
                    End If
                    cRetGan = 0 's2n(!retgan)
                    Consulta = "insert into " & sTablaTemp _
                        & " ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337, " _
                        & " imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values ( " _
                        & ssFecha(rs!Fecha) & ",'" & !razonsocialprov & "','" & !cuitprov & "','" _
                        & TipoyNro(!TIPODOC, !letraprov, !suc, !NroDoc) & "'," & signo & x2s(cNeto21) & ", " & signo & x2s(cNeto10) & " , " & signo & x2s(cNeto27) & " ," & signo & x2s(cIva21) & "," _
                        & signo & x2s(cIva10) & "," & signo & x2s(cIva27) & "," & signo & x2s(!percepc) & "," & signo & x2s(rs!Total) & "," & signo & x2s(!imp_int) _
                        & "," & signo & x2s(cRetGan) & "," & signo & x2s(!der_est) & "," & signo & x2s(!EXENTO) & "," & signo & x2s(rs!ibcapital) & "," & signo & x2s(rs!ibprovincia) & "," & signo & x2s(!nogravado) & ")"
                    DataEnvironment1.Sistema.Execute Consulta
                    rs.MoveNext
                Loop
            End If
            rs.Close
            
            asse = "compras2"
            ssql = " SELECT COMPRAS.*, Ivas.letraprov FROM COMPRAS INNER JOIN Ivas ON COMPRAS.TIPOIVA = Ivas.codigo where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and compras.activo=1 and (compras.percepc<>0 or compras.der_est<>0)"  'and ivas.letra = 'A'"
            rs.Open ssql, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                Do While Not rs.EOF
                    signo = ""
                    If Trim(rs!TIPODOC) = "N/C" Then      'Or Trim(rs!TIPODOC) = "NCB"
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
                        cNeto21 = s2n((100 * cIva21) / 21)
                        cNeto10 = s2n((100 * cIva10) / 10.5)
                        cNeto27 = s2n((100 * cIva27) / 27)
                    End If
                    cRetGan = 0 's2n(!retgan)
                    Consulta = "insert into " & sTablaTemp _
                        & " ( fecha, razonsocial, nrocuit, tipoynro, neto21,neto10,neto27, iva21, iva10, iva27, rg3337, " _
                        & " imptotal,  impint, retencgan, rg3431, exento, IB_CAPITAL, IB_PROVINCIA,nograbado ) values ( " _
                        & ssFecha(rs!Fecha) & ",'" & !razonsocialprov & "','" & !cuitprov & "','" _
                        & TipoyNro(!TIPODOC, !letraprov, !suc, !NroDoc) & "'," & signo & x2s(cNeto21) & ", " & signo & x2s(cNeto10) & " , " & signo & x2s(cNeto27) & " ," & signo & x2s(cIva21) & "," _
                        & signo & x2s(cIva10) & "," & signo & x2s(cIva27) & "," & signo & x2s(!percepc) & " ," & signo & x2s(rs!Total) & "," & signo & x2s(!imp_int) _
                        & "," & signo & x2s(cRetGan) & "," & signo & x2s(!der_est) & "," & signo & x2s(!EXENTO) & "," & signo & x2s(rs!ibcapital) & "," & signo & x2s(rs!ibprovincia) & "," & signo & x2s(!nogravado) & ")"
                    DataEnvironment1.Sistema.Execute Consulta
                    rs.MoveNext
                Loop
            End If
            rs.Close
            RptIvaCompras.lblTitulo = "Listado de percepcion de Iva del mes " & Trim(cmbmes.Text) & " - Año " & Trim(cmbaño.Text)
        Else
            MsgBox "Debe seleccionar alguna de las opciones", 48, "Atencion"
        End If
    End If
    End With
    
    asse = "reporte"
    str = "select * from " & sTablaTemp & " order by fecha"
    RptIvaCompras.Data.Connection = DataEnvironment1.Sistema
    RptIvaCompras.Data.Source = str
    RptIvaCompras.lblfecha = Date
    RptIvaCompras.Printer.PaperSize = vbPRPSLegal 'hoja oficio
    
    scampos = " fecha as [Fecha], [razonsocial] as [Razon social] , [nrocuit] as [Nro CUIT] , [tipoynro] as [Documento],  [neto21] as [ Neto 21],  [neto10] as [ Neto 10.5],  [neto27] as [ Neto 27]   , [iva21] as  [IVA 21] , [iva10] as [IVA 10.5]  , [iva27] as [IVA 27], [rg3337] as [RG 3337], [IB_CAPITAL]+ [IB_PROVINCIA] as [       IIBB], [rg3431] as [ RG 3431],[Nograbado] as [No Grabado], [imptotal] as [          TOTAL     ]"
    
    LlenarGrilla GRILLA, "select " & scampos & " from " & sTablaTemp & " order by fecha", False
    
    i = 1
    While i < GRILLA.rows
        GRILLA.TextMatrix(i, 4) = s2n(GRILLA.TextMatrix(i, 4), 2, True)
        GRILLA.TextMatrix(i, 5) = s2n(GRILLA.TextMatrix(i, 5), 2, True)
        GRILLA.TextMatrix(i, 6) = s2n(GRILLA.TextMatrix(i, 6), 2, True)
        GRILLA.TextMatrix(i, 7) = s2n(GRILLA.TextMatrix(i, 7), 2, True)
        GRILLA.TextMatrix(i, 8) = s2n(GRILLA.TextMatrix(i, 8), 2, True)
        GRILLA.TextMatrix(i, 9) = s2n(GRILLA.TextMatrix(i, 9), 2, True)
        GRILLA.TextMatrix(i, 10) = s2n(GRILLA.TextMatrix(i, 10), 2, True)
        GRILLA.TextMatrix(i, 11) = s2n(GRILLA.TextMatrix(i, 11), 2, True)
        GRILLA.TextMatrix(i, 12) = s2n(GRILLA.TextMatrix(i, 12), 2, True)
        GRILLA.TextMatrix(i, 13) = s2n(GRILLA.TextMatrix(i, 13), 2, True)
        GRILLA.TextMatrix(i, 14) = s2n(GRILLA.TextMatrix(i, 14), 2, True)
        i = i + 1
    Wend
    
    GRILLA.ColWidth(1) = 2900
    GRILLA.ColWidth(2) = 1200
    GRILLA.ColWidth(3) = 1800
    GRILLA.ColWidth(4) = 1100
    GRILLA.ColWidth(5) = 1100
    GRILLA.ColWidth(8) = 1100
    GRILLA.ColWidth(14) = 1100
    
    GRILLA.ColHidden(4) = True
    GRILLA.ColHidden(5) = True
    GRILLA.ColHidden(6) = True
    GRILLA.ColHidden(7) = True
    GRILLA.ColHidden(8) = True
    GRILLA.ColHidden(9) = True
    GRILLA.ColHidden(11) = True
    GRILLA.ColHidden(13) = True
    GRILLA.ColHidden(14) = True
    
    sumarizo GRILLA, Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
    
fin:
    Set rs = Nothing
    relojito False
    Exit Sub
UFAlistado:
    ufa "err prg", asse & " --consulta=  " & Consulta & " -- " ', Err
    Resume fin
End Sub

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
    
    If optfecha.Value = True Then
        PrintG GRILLA, pHorizontal, "PERCEPCIONES DE IVA", IIf(bb, Date, "01/01/1900"), "PERCEPCIONES DE IVA DESDE : " & dtfechad.Value & " AL " & dtfechah.Value, vbPRPSLegal
    Else
        PrintG GRILLA, pHorizontal, "PERCEPCIONES DE IVA", IIf(bb, Date, "01/01/1900"), "PERCEPCIONES DE IVA DEl MES : " & cmbmes.Text & " DEL " & cmbaño.Text, vbPRPSLegal
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
    uXls.ini GRILLA, "c:\ListPerc", "Listado de precepciones " & dtfechad & "  -  " & dtfechah
    Form_Resize
    optfecha.Value = True
    optmes.Value = False
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Resize()
    Anclar GRILLA, Me, anclarLadosTodos
End Sub

Private Sub optfecha_Click()
    If optfecha.Value = True Then
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
        optfecha.Value = False
    Else
        If optfecha.Value = True Then
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
    TipoyNro = Left(tdoc & " " & letra & "   ", 6) & Format(suc, "0000") & "-" & Format(Nro, "00000000")
End Function

Public Function TYN(suc As Long, Nro As Long) As String
    TYN = Format(suc, "0000") & "-" & Format(Nro, "00000000")
End Function


