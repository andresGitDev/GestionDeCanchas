VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmLisTipoCompras 
   Caption         =   "Listado de Totales x Tipo de Compras"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   Icon            =   "FrmLisTipoCompras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excel"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.ComboBox cmbaño 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmLisTipoCompras.frx":08CA
      Left            =   4560
      List            =   "FrmLisTipoCompras.frx":08CC
      TabIndex        =   6
      Text            =   "Año"
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cmbmes 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmLisTipoCompras.frx":08CE
      Left            =   1320
      List            =   "FrmLisTipoCompras.frx":08F6
      TabIndex        =   5
      Text            =   "Mes"
      Top             =   600
      Width           =   1935
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58785793
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58785793
      CurrentDate     =   38252
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Hasta"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Desde"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   765
      Left            =   360
      Top             =   1560
      Width           =   5760
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   765
      Left            =   360
      Top             =   360
      Width           =   5760
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2445
      Left            =   120
      Top             =   120
      Width           =   6240
   End
End
Attribute VB_Name = "FrmLisTipoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 5/10/4

Private smTipoComprasTmp As String
Private Const tt_TipoComprasTemp = _
"(TipoCompra  int, " & _
"[descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
"[neto] [float] NULL, [exento] [float] NULL , " & _
"[NoGravado] [float] NULL ," & _
"[rg3431] [float] NULL ," & _
"[retencgan] [float] NULL ,[iva] [float] NULL, [impint] [float] NULL ," & _
"[IB_CAPITAL] [float] NULL ,[IB_PROVINCIA] [float] NULL,[imptotal] [float] NULL,[rg3337] [float] NULL)"

Private Sub cmdaceptar_Click()
    Dim STR As String
    Dim rs As New ADODB.Recordset
    Dim rstrans As New ADODB.Recordset
    Dim Neto As Variant
    Dim Iva As Variant
    Dim wheFecha As String
    
    smTipoComprasTmp = TablaTempCrear(tt_TipoComprasTemp)
              
            If optFecha Then
                wheFecha = " fecha>=" & ssFecha(dtfechad) & " and fecha<=" & ssFecha(dtfechah)
            ElseIf optmes Then
                wheFecha = " mesimp = " & (cmbmes.ListIndex + 1) & " And anoimp = " & Val(Trim(cmbaño.Text)) & ""
            End If
            
            Dim s As String
            
            s = "select tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados, " & _
            " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10 ) as ivas,sum(Imp_Int)as impint, " & _
            " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
            " from TRANSCOM inner join tipocompras on transcom.tipocompra=tipocompras.codigo " & _
            " where " & wheFecha & " and (tipodoc='FAC' or tipodoc='N/D' ) and transcom.activo=1 " & _
            " group by tipocompra, tipocompras.descripcion"
            Insertar s, 1
            
            s = "select  tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados, " & _
            " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias, sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas, sum(Imp_Int)as impint," & _
            " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
            " from TRANSCOM inner join tipocompras on transcom.tipocompra=tipocompras.codigo " & _
            " where " & wheFecha & " and (tipodoc='N/C') and transcom.activo=1 " & _
            " group by  tipocompra, tipocompras.descripcion"
            Insertar s, -1
            
            s = "select tipocompra,  tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados," & _
            " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas,sum(Imp_Int)as impint, " & _
            " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
            " from COMPRAS inner join tipocompras on compras.tipocompra=tipocompras.codigo " & _
            " where " & wheFecha & " and (tipodoc='FAC' or tipodoc='N/D' ) and compras.activo=1 " & _
            " group by  tipocompra, tipocompras.descripcion"
            Insertar s, 1

            s = "select  tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados," & _
            " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas,sum(Imp_Int)as impint, " & _
            " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
            " from COMPRAS inner join tipocompras on compras.tipocompra=tipocompras.codigo " & _
            " where " & wheFecha & " and (tipodoc='N/C') and compras.activo=1  group by  tipocompra, tipocompras.descripcion"
            Insertar s, -1
            

            STR = "select  descripcion,sum(neto) as netos,sum(exento)as exento,sum(NoGravado)as NoGravado, " & _
            " sum(rg3431) as rg3431,sum(rg3337) as rg3337,sum(retencgan)as retencgan,sum(iva) as iva,sum(impint)as impint, " & _
            " sum(IB_capital) as IBcapital,sum(IB_provincia) as IBprovincia,sum(Imptotal) as ImpTotal  " & _
            " from " & smTipoComprasTmp & " group by tipocompra, descripcion  order by tipocompra "
            
            rptTotalesTipoCompras.Data.Connection = DataEnvironment1.Sistema
            rptTotalesTipoCompras.Data.Source = STR
            rptTotalesTipoCompras.lblfecha = Date
            rptTotalesTipoCompras.lblTitulo = "Listado de Tipo de Compras del mes " & Trim(cmbmes.Text) & " -Año " & Trim(cmbaño.Text)
            rptTotalesTipoCompras.Show

    
End Sub

Private Sub Insertar(ss As String, signo)  ' descr, Neto, Iva)
    Dim ars As New ADODB.Recordset

    With ars
        .Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
           DataEnvironment1.Sistema.Execute _
                "insert into " & smTipoComprasTmp & "  (tipocompra, descripcion,neto,exento,NoGravado," & _
                "rg3431,retencgan,iva,impint,IB_capital,IB_provincia,ImpTotal,rg3337) " & _
                " values ('" & !Tipocompra & "', '" & !des & "','" & x2s(!netos * signo) & "'," & _
                "'" & x2s(!exentos * signo) & "','" & x2s(!NoGravados * signo) & "', " & _
                "'" & x2s(!rg_3431 * signo) & "','" & x2s(!retgacias * signo) & "','" & x2s(!IVAS * signo) & "', " & _
                "'" & x2s(!impint * signo) & "','" & x2s(!IBcapitales * signo) & "','" & x2s(!IBprovincias * signo) & "'," & _
                "'" & x2s(!Totales * signo) & "','" & x2s(!rg_3337 * signo) & "' )"
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub cmdcancelar_Click()
    'dtfechad = Date - 30 '"01/" & Month(Date) - 1 & "/" & Year(Date) 'Date
    'dtfechah = Date
    'cmbmes.Text = ""
    'cmbaño.Text = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim arch As String
    
    Dim STR As String
    Dim rs As New ADODB.Recordset
    Dim rstrans As New ADODB.Recordset
    Dim Neto As Variant
    Dim Iva As Variant
    
    smTipoComprasTmp = TablaTempCrear(tt_TipoComprasTemp)
    
    Dim s As String
    
    s = "select tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados, " & _
    " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10 ) as ivas,sum(Imp_Int)as impint, " & _
    " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
    " from TRANSCOM inner join tipocompras on transcom.tipocompra=tipocompras.codigo " & _
    " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' ) and transcom.activo=1 " & _
    " group by tipocompra, tipocompras.descripcion"
    Insertar s, 1
    
    s = "select  tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados, " & _
    " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias, sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas, sum(Imp_Int)as impint," & _
    " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
    " from TRANSCOM inner join tipocompras on transcom.tipocompra=tipocompras.codigo " & _
    " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and transcom.activo=1 " & _
    " group by  tipocompra, tipocompras.descripcion"
    Insertar s, -1
    
    s = "select tipocompra,  tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados," & _
    " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas,sum(Imp_Int)as impint, " & _
    " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
    " from COMPRAS inner join tipocompras on compras.tipocompra=tipocompras.codigo " & _
    " where mesimp=" & cmbmes.ListIndex + 1 & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' ) and compras.activo=1 " & _
    " group by  tipocompra, tipocompras.descripcion"
    Insertar s, 1

    s = "select  tipocompra, tipocompras.descripcion as des,sum(neto)as netos, sum(exento)as exentos,sum(NoGravado)as NoGravados," & _
    " sum(der_est)as rg_3431,sum(percepc) as rg_3337,sum(ret_gan) as retgacias ,sum(iva_21+ iva_27 + iva_9 + iva_10) as ivas,sum(Imp_Int)as impint, " & _
    " sum(IBcapital) as IBcapitales,sum(IBprovincia) as IBprovincias,sum(total) as Totales " & _
    " from COMPRAS inner join tipocompras on compras.tipocompra=tipocompras.codigo " & _
    " where mesimp=" & cmbmes.ListIndex + 1 & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and compras.activo=1  group by  tipocompra, tipocompras.descripcion"
    Insertar s, -1
    

    STR = "select  descripcion,sum(neto) as netos,sum(exento)as exento,sum(NoGravado)as NoGravado, " & _
    " sum(rg3431) as rg3431,sum(rg3337) as rg3337,sum(retencgan)as retencgan,sum(iva) as iva,sum(impint)as impint, " & _
    " sum(IB_capital) as IBcapital,sum(IB_provincia) as IBprovincia,sum(Imptotal) as ImpTotal  " & _
    " from " & smTipoComprasTmp & " group by tipocompra, descripcion  order by tipocompra "
            
    Set rs = Nothing
    rs.Open STR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    CommonDialog1.Filter = "Archivos de planilla de calculo(*.xls)|*.xls|Todos los archivos (*.*)|*.*|"
    CommonDialog1.ShowSave
    arch = CommonDialog1.FileName
    
    VinculoXl arch, "Listado de Tipo de Compras del mes " & Trim(cmbmes.Text) & " -Año " & Trim(cmbaño.Text), , , rs
End Sub

Private Sub Form_Load()
Dim i As Long, anio As Long
Dim ejj
ejj = obtenerDeSQL("select fechainicio,fechafin from ejercicio where activo=1")
    dtfechad = CDate(ejj(0))
    dtfechah = CDate(ejj(1))

'    dtfechad = Date - 30 '"01/" & Month(Date) - 1 & "/" & Year(Date) 'Date
'    dtfechah = Date
'    cmbmes.Text = ""
'    cmbaño.Text = ""
anio = Year(Date) - 10
For i = 1 To 10
    cmbaño.AddItem anio + i
Next
End Sub



Private Sub optfecha_Click()
    If optmes.Value = True Then
        cmbmes.enabled = True
        cmbaño.enabled = True
        dtfechad.enabled = False
        dtfechah.enabled = False
    Else
        If optFecha.Value = True Then
            cmbmes.enabled = False
            cmbaño.enabled = False
            dtfechad.enabled = True
            dtfechah.enabled = True
        End If
    End If
End Sub

Private Sub optmes_Click()
    If optmes.Value = True Then
        cmbmes.enabled = True
        cmbaño.enabled = True
        dtfechad.enabled = False
        dtfechah.enabled = False
    Else
        If optFecha.Value = True Then
            cmbmes.enabled = False
            cmbaño.enabled = False
            dtfechad.enabled = True
            dtfechah.enabled = True
        End If
    End If
End Sub


            


'[imptotal] [float] NULL ,---------------------------------------------
'[impint] [float] NULL ,--------IMP_INT-----------------------------------------
'[retencgan] [float] NULL ,---RET_GAN------------------------------------
'[rg3431] [float] NULL ,-----DER_EST-----------------------------------

