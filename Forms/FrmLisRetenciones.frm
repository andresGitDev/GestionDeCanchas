VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisRetenciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Listado de Retenciones"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chktodas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Todas"
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
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   555
      Width           =   1005
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
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1425
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
      Left            =   1545
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1425
      Width           =   975
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
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1425
      Width           =   975
   End
   Begin VB.ComboBox cmbret 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmLisRetenciones.frx":0000
      Left            =   1230
      List            =   "FrmLisRetenciones.frx":0028
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   930
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   285
      Left            =   1470
      TabIndex        =   4
      Top             =   195
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   94306305
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   195
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   94306305
      CurrentDate     =   38252
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      Height          =   1275
      Left            =   60
      Top             =   60
      Width           =   5970
   End
   Begin VB.Label Label1 
      BackColor       =   &H00745134&
      BackStyle       =   0  'Transparent
      Caption         =   "Retencion"
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
      Left            =   165
      TabIndex        =   8
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackColor       =   &H00745134&
      BackStyle       =   0  'Transparent
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
      Left            =   150
      TabIndex        =   7
      Top             =   210
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00745134&
      BackStyle       =   0  'Transparent
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
      Left            =   3315
      TabIndex        =   6
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLisRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tt_LisRet_temp As String
Private Sub chktodas_Click()
    If chktodas.Value = 1 Then
        cmbret.Enabled = False
    Else
        cmbret.Enabled = True
    End If
End Sub

Private Sub cmdaceptar_Click()
'On Error GoTo UFAlistado
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim STR As String
Dim COMP As String

    tt_LisRet_temp = " ([fecha] [datetime] NULL ,[tiporet] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[nro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[importe] [float] NULL ,[comprobante] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[cliente] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"
    sTablaTemp = TablaTempCrear(tt_LisRet_temp)
    If chktodas.Value = 1 Then
        rs.Open "select * from facturaventa inner join tiporetenciones on  facturaventa.tipodoc=tiporetenciones.codigo where fecha>=" & ssFecha(dtfechad) & " and fecha<=" & ssFecha(dtfechah) & " and activo=1", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
                If rs!nrofactura = 10528 Then
                   MsgBox "a"
                End If
                rs2.Open "select * from facturaretencion where tdoc_ret='" & rs!TIPODOC & "' and ndoc_ret=" & rs!nrofactura, DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
                If Not rs2.EOF Then
                    'COMP = rs2!codfactura
                    COMP = obtenerDeSQL("SELECT nrofactura FROM facturaventa WHERE codigo=" & rs2!codfactura & "")
                Else
                    COMP = ""
                End If
                rs2.Close
                Set rs2 = Nothing
                DataEnvironment1.AMR.Execute "insert into " & sTablaTemp & "(fecha,tiporet,nro,importe,comprobante,cliente)values(" & ssFecha(rs!fecha) & ",'" & Mid(ObtenerDescripcionS("Tiporetenciones", Trim(rs!TIPODOC)), 11) & "','" & rs!nrofactura & "'," & Replace(rs!Total, ",", ".") & ",'" & COMP & "','" & ObtenerDescripcion("clientes", rs!cliente) & "')"
                rs.MoveNext
            Loop
        Else
            MsgBox "No hay Datos para mostrar", 48, "Atencion"
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
    Else
        rs.Open "select * from facturaventa where fecha>=" & ssFecha(dtfechad) & " and fecha<=" & ssFecha(dtfechah) & " and activo=1 and tipodoc='" & ObtenerCodigoS("tiporetenciones", Trim(cmbret.Text)) & "'", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
                rs2.Open "select * from facturaretencion where tdoc_ret='" & rs!TIPODOC & "' and ndoc_ret=" & rs!nrofactura, DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
                If Not rs2.EOF Then
                    'COMP = rs2!codfactura
                    COMP = obtenerDeSQL("SELECT nrofactura FROM facturaventa WHERE codigo=" & rs2!codfactura & "")
                Else
                    COMP = ""
                End If
                rs2.Close
                Set rs2 = Nothing
                DataEnvironment1.AMR.Execute "insert into " & sTablaTemp & "(fecha,tiporet,nro,importe,comprobante,cliente)values(" & ssFecha(rs!fecha) & ",'" & Mid(ObtenerDescripcionS("Tiporetenciones", Trim(rs!TIPODOC)), 11) & "','" & rs!nrofactura & "'," & Replace(rs!Total, ",", ".") & ",'" & COMP & "','" & ObtenerDescripcion("clientes", rs!cliente) & "')"
                rs.MoveNext
            Loop
        Else
            MsgBox "No hay Datos para mostrar", 48, "Atencion"
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    RptRetenciones.lbltitulo = "Listado de Retenciones desde el " & dtfechad.Value & " al " & dtfechah.Value
    STR = "select * from " & sTablaTemp & " order by fecha "
    RptRetenciones.Data.Connection = DataEnvironment1.AMR
    RptRetenciones.Data.Source = STR
    RptRetenciones.Fiecliente.DataField = "CLIENTE"
    RptRetenciones.Fiecomp.DataField = "COMPROBANTE"
    RptRetenciones.Fiefecha.DataField = "FECHA"
    RptRetenciones.Fieimporte.DataField = "IMPORTE"
    RptRetenciones.Fienro.DataField = "NRO"
    RptRetenciones.Fietipo.DataField = "TIPORET"
    RptRetenciones.lblfecha = Date
    RptRetenciones.Show
    
fin:
    Exit Sub
UFAlistado:
    MsgBox "err en listado"
    Resume fin
End Sub

Private Sub cmdcancelar_Click()
    LimpioControles
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Sub CargaCombo(Combo As Object, Tabla As String, campo As String, Bound As String, wer As String)

Dim rsCargacombo As New ADODB.Recordset
Dim sqlstrCC As String
Dim i As Long
    If Bound <> "" Then
        sqlstrCC = "Select " + campo + " as NN," + Bound + " from " + Tabla
    Else
        sqlstrCC = "Select " + campo + " as NN" + Bound + " from " + Tabla
    End If
    If wer <> "" Then
        sqlstrCC = sqlstrCC + " and " + wer
    End If
    sqlstrCC = sqlstrCC + " order by " + campo
    rsCargacombo.Open sqlstrCC, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    Combo.Clear
    If Not rsCargacombo.EOF And Not rsCargacombo.BOF Then
        rsCargacombo.MoveFirst
        i = 0
        While Not rsCargacombo.EOF
            Combo.AddItem Trim(rsCargacombo.Fields("NN"))
            Combo.ItemData(i) = i
            i = i + 1
            rsCargacombo.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
    Call CargaCombo(cmbret, "tiporetenciones", "descripcion", "codigo", "")
    LimpioControles
End Sub
Sub LimpioControles()
    dtfechad.Value = Date - 30 '"01/" & Month(Date) - 1 & "/" & Year(Date) 'Date
    dtfechah.Value = Date
    chktodas.Value = 1
    cmbret.ListIndex = -1
End Sub

