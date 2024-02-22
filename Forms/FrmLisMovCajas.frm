VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisCajas 
   Caption         =   "Listado de Mov de Cajas"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   Icon            =   "FrmLisMovCajas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opttodosef3 
      Caption         =   "Todos"
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
      Left            =   4560
      TabIndex        =   21
      Tag             =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optsoloch3 
      Caption         =   "Solo Ch. 3º"
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
      Left            =   2640
      TabIndex        =   20
      Tag             =   "1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton optsolofvo 
      Caption         =   "Solo Efvo."
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
      Left            =   600
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtcodcaja 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Tag             =   "1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbcaja 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caja"
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
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtcaja 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Tag             =   "2"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38252
   End
   Begin VB.Frame Framedos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   585
      Left            =   480
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   3375
      Begin VB.OptionButton optbuscocaja 
         Caption         =   "Buscar Caja"
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
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   1455
      End
      Begin VB.OptionButton opttodas 
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
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Tag             =   "1"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frameuno 
      Height          =   615
      Left            =   480
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   5775
      Begin VB.OptionButton optef3 
         Caption         =   "Efvo. / Ch. 3º"
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
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton opttodos 
         Caption         =   "Todos"
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
         Left            =   2160
         TabIndex        =   15
         Tag             =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton optsolobancos 
         Caption         =   "Solo Bancos"
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
         Left            =   4080
         TabIndex        =   14
         Tag             =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblcaja 
      Caption         =   "Nº Caja:"
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
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Left            =   3600
      TabIndex        =   6
      Top             =   480
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
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLisCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msCajaTemp As String

Private Sub LimpioPrimero(habilito As Boolean)
    Frameuno.Visible = habilito
    optef3.Visible = habilito
    opttodos.Visible = habilito
    optsolobancos.Visible = habilito
End Sub

Private Sub LimpioSegundo(habilito As Boolean)
    Framedos.Visible = habilito
    optsolofvo.Visible = habilito
    optsoloch3.Visible = habilito
    opttodosef3.Visible = habilito
End Sub

Private Sub cmbcaja_Click()
    Dim sql As String

    sql = "select codigo as [Código   ], sector  as [ Sector              ],  responsable as [Responsable            ] from cajas where activo = 1 order by codigo"
    
    frmBuscar.MostrarSql (sql)
     
    If frmBuscar.resultado(1) <> "" Then
        txtcodcaja = frmBuscar.resultado(1)
        txtCaja = frmBuscar.resultado(2) '  ObtenerDescripcionCajas("Cajas", frmBuscar.resultado(1))
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpioPrimero (False)
    LimpioSegundo (False)
    optbuscocaja = False
    opttodas = False
    dtfechad = Date
    dtfechah = Date
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
    Dim STR As String, tipo As String
    Dim Total As Double, saldoanterior As Double, diant As Variant
    Dim punto1 As Double, punto2 As Double, punto3 As Double, punto4 As Double, punto5 As Double, punto6 As Double, punto7 As Double, punto8 As Double, punto9 As Double
    Dim rs As New ADODB.Recordset

'generacion tablatemp
    'If msCajaTemp = "" Then
    msCajaTemp = TablaTempCrear(tt_CajasTemp)
    DataEnvironment1.Sistema.Execute "CREATE  INDEX ixi ON " & msCajaTemp & " ([fecha]) ON [PRIMARY] "
    'End If

    'daTaenvironment1.Sistema.Execute "delete from " & msCajaTemp 'CajasTemp"

    
    If optbuscocaja = True Then
        
        'PARA UNA CAJA SOLA
        rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and caja = " & Val(txtcodcaja) & " and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
            While Not rs.EOF
                diant = rs!fecha
                Total = 0
                Do While rs!fecha = diant
                    Select Case rs!tipo
                        Case "E":   tipo = "Efectivo"
                        Case "C":   tipo = "Ch. 3º"
                        Case "P":   tipo = "Ch. Propio"
                        Case "T":   tipo = "Transferencia"
                        Case "G":   tipo = "Gasto B."
                        Case "D":   tipo = "Créd. B."
                    End Select
                    If rs!Ing_egr = "I" Then
                        DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                        & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                    Else
                        DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                        & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                    End If
                    
                    If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                        punto1 = punto1 + rs!importe
                        If Left(rs!concepto, 6) = "FacCdo" Then
                            punto2 = punto2 + rs!importe
                        Else
                            punto3 = punto3 + rs!importe
                        End If
                    Else
                        If rs!TIPODOC = "FAC" Then
                            punto4 = punto4 + rs!importe
                            If buscotipopago(rs!codfp) = 1 Then
                                punto5 = punto5 + rs!importe 'TARJETA
                            Else
                                If buscotipopago(rs!codfp) = 2 Then
                                    punto6 = punto6 + rs!importe 'EFECTIVO
                                End If
                            End If
                        Else
                            If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                punto7 = punto7 + rs!importe
                            Else
                                punto8 = punto8 + rs!importe
                            End If
                        End If
                    End If
                    rs.MoveNext
                    If rs.EOF Then Exit Do
                Loop
            Wend
'        End If
        punto9 = punto9 + punto1 + punto8 - punto4 - punto7
        punto4 = punto4 * -1
        punto7 = punto7 * -1
        rs.Close
        Set rs = Nothing
        
    Else
        If optef3 = True Then
            If optsolofvo = True Then
                'ACA VAN SOLO EFECTIVO
                rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and caja > 0 and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    While Not rs.EOF
                        diant = rs!fecha
                        Total = 0
                        Do While rs!fecha = diant
                            Select Case rs!tipo
                                Case "E":   tipo = "Efectivo"
                                Case "C":   tipo = "Ch. 3º"
                                Case "P":   tipo = "Ch. Propio"
                                Case "T":   tipo = "Transferencia"
                                Case "G":   tipo = "Gasto B."
                                Case "D":   tipo = "Créd. B."
                            End Select
                            If rs!Ing_egr = "I" Then
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            Else
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            End If
                            
                            If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                                punto1 = punto1 + rs!importe
                                If Left(rs!concepto, 6) = "FacCdo" Then
                                    punto2 = punto2 + rs!importe
                                Else
                                    punto3 = punto3 + rs!importe
                                End If
                            Else
                                If rs!TIPODOC = "FAC" Then
                                    punto4 = punto4 + rs!importe
                                    If buscotipopago(rs!codfp) = 1 Then
                                        punto5 = punto5 + rs!importe 'TARJETA
                                    Else
                                        If buscotipopago(rs!codfp) = 2 Then
                                            punto6 = punto6 + rs!importe
                                        End If
                                    End If
                                Else
                                    If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                        punto7 = punto7 + rs!importe
                                    Else
                                        punto8 = punto8 + rs!importe
                                    End If
                                End If
                            End If
                            
                            rs.MoveNext
                            If rs.EOF Then Exit Do
                        Loop
                    Wend
                End If
                punto9 = punto9 + punto1 + punto8 - punto4 - punto7
                punto4 = punto4 * -1
                punto7 = punto7 * -1
                rs.Close
                Set rs = Nothing
            Else
                If optsoloch3 = True Then
                    'ACA VAN SOLO CH 3º
                    rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and caja = 0 and tipo = 'C' and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If Not rs.EOF Then
                        While Not rs.EOF
                            diant = rs!fecha
                            Total = 0
                            Do While rs!fecha = diant
                                Select Case rs!tipo
                                    Case "E":   tipo = "Efectivo"
                                    Case "C":   tipo = "Ch. 3º"
                                    Case "P":   tipo = "Ch. Propio"
                                    Case "T":   tipo = "Transferencia"
                                    Case "G":   tipo = "Gasto B."
                                    Case "D":   tipo = "Créd. B."
                                End Select
                                If rs!Ing_egr = "I" Then
                                    DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                    & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                                Else
                                    DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                    & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                                End If
                                
                                If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                                    punto1 = punto1 + rs!importe
                                    If Left(rs!concepto, 6) = "FacCdo" Then
                                        punto2 = punto2 + rs!importe
                                    Else
                                        punto3 = punto3 + rs!importe
                                    End If
                                Else
                                    If rs!TIPODOC = "FAC" Then
                                        punto4 = punto4 + rs!importe
                                        If buscotipopago(rs!codfp) = 1 Then
                                            punto5 = punto5 + rs!importe 'TARJETA
                                        Else
                                            If buscotipopago(rs!codfp) = 2 Then
                                                punto6 = punto6 + rs!importe
                                            End If
                                        End If
                                    Else
                                        If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                            punto7 = punto7 + rs!importe
                                        Else
                                            punto8 = punto8 + rs!importe
                                        End If
                                    End If
                                End If
                                rs.MoveNext
                                If rs.EOF Then Exit Do
                            Loop
                        Wend
                    End If
                    punto9 = punto9 + punto1 + punto8 - punto4 - punto7
                    punto4 = punto4 * -1
                    punto7 = punto7 * -1
                    rs.Close
                    Set rs = Nothing
                Else
                    'ACA VAN EFEC / CH 3º
                    rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and (caja > 0 or (caja = 0 and tipo = 'C')) and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                    If Not rs.EOF Then
                        While Not rs.EOF
                            diant = rs!fecha
                            Total = 0
                            Do While rs!fecha = diant
                                Select Case rs!tipo
                                    Case "E":   tipo = "Efectivo"
                                    Case "C":   tipo = "Ch. 3º"
                                    Case "P":   tipo = "Ch. Propio"
                                    Case "T":   tipo = "Transferencia"
                                    Case "G":   tipo = "Gasto B."
                                    Case "D":   tipo = "Créd. B."
                                End Select
                                If rs!Ing_egr = "I" Then
                                    DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                    & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                                Else
                                    DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                    & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                                End If
                                
                                If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                                    punto1 = punto1 + rs!importe
                                    If Left(rs!concepto, 6) = "FacCdo" Then
                                        punto2 = punto2 + rs!importe
                                    Else
                                        punto3 = punto3 + rs!importe
                                    End If
                                Else
                                    If rs!TIPODOC = "FAC" Then
                                        punto4 = punto4 + rs!importe
                                        If buscotipopago(rs!codfp) = 1 Then
                                            punto5 = punto5 + rs!importe 'TARJETA
                                        Else
                                            If buscotipopago(rs!codfp) = 2 Then
                                                punto6 = punto6 + rs!importe
                                            End If
                                        End If
                                    Else
                                        If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                            punto7 = punto7 + rs!importe
                                        Else
                                            punto8 = punto8 + rs!importe
                                        End If
                                    End If
                                End If
                                rs.MoveNext
                                If rs.EOF Then Exit Do
                            Loop
                        Wend
                    End If
                    punto9 = punto9 + punto1 + punto8 - punto4 - punto7
                    punto4 = punto4 * -1
                    punto7 = punto7 * -1
                    rs.Close
                    Set rs = Nothing
                End If
            End If
        Else
            If opttodos = True Then
                'ACA VAN TODOS
                rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    While Not rs.EOF
                        diant = rs!fecha
                        Total = 0
                        Do While rs!fecha = diant
                            Select Case rs!tipo
                                Case "E":   tipo = "Efectivo"
                                Case "C":   tipo = "Ch. 3º"
                                Case "P":   tipo = "Ch. Propio"
                                Case "T":   tipo = "Transferencia"
                                Case "G":   tipo = "Gasto B."
                                Case "D":   tipo = "Créd. B."
                            End Select
                            If rs!Ing_egr = "I" Then
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            Else
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            End If
                            
                            If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                                punto1 = punto1 + rs!importe
                                If Left(rs!concepto, 6) = "FacCdo" Then
                                    punto2 = punto2 + rs!importe
                                Else
                                    punto3 = punto3 + rs!importe
                                End If
                            Else
                                If rs!TIPODOC = "FAC" Then
                                    punto4 = punto4 + rs!importe
                                    If buscotipopago(rs!codfp) = 1 Then
                                        punto5 = punto5 + rs!importe 'TARJETA
                                    Else
                                        If buscotipopago(rs!codfp) = 2 Then
                                            punto6 = punto6 + rs!importe
                                        End If
                                    End If
                                Else
                                    If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                        punto7 = punto7 + rs!importe
                                    Else
                                        punto8 = punto8 + rs!importe
                                    End If
                                End If
                            End If
                            rs.MoveNext
                            If rs.EOF Then Exit Do
                        Loop
                    Wend
                End If
                punto9 = punto9 + punto1 + punto8 - punto4 - punto7
                punto4 = punto4 * -1
                punto7 = punto7 * -1
                rs.Close
                Set rs = Nothing
                
            Else
                'ACA VAN SOLO BANCOS
                rs.Open "select * from movicaja where fecha " & ssBetween(dtfechad, dtfechah) & " and caja = 0 and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then
                    While Not rs.EOF
                        diant = rs!fecha
                        Total = 0
                        Do While rs!fecha = diant
                            Select Case rs!tipo
                                Case "E":   tipo = "Efectivo"
                                Case "C":   tipo = "Ch. 3º"
                                Case "P":   tipo = "Ch. Propio"
                                Case "T":   tipo = "Transferencia"
                                Case "G":   tipo = "Gasto B."
                                Case "D":   tipo = "Créd. B."
                            End Select
                            If rs!Ing_egr = "I" Then
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            Else
                                DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                                & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                            End If
                            
                            If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                                punto1 = punto1 + rs!importe
                                If Left(rs!concepto, 6) = "FacCdo" Then
                                    punto2 = punto2 + rs!importe
                                Else
                                    punto3 = punto3 + rs!importe
                                End If
                            Else
                                If rs!TIPODOC = "FAC" Then
                                    punto4 = punto4 + rs!importe
                                    If buscotipopago(rs!codfp) = 1 Then
                                        punto5 = punto5 + rs!importe 'TARJETA
                                    Else
                                        If buscotipopago(rs!codfp) = 2 Then
                                            punto6 = punto6 + rs!importe
                                        End If
                                    End If
                                Else
                                    If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                        punto7 = punto7 + rs!importe
                                    Else
                                        punto8 = punto8 + rs!importe
                                    End If
                                End If
                            End If
                            rs.MoveNext
                            If rs.EOF Then Exit Do
                        Loop
                    Wend
                End If
                punto9 = punto9 + punto1 + punto8 - punto4 - punto7
                punto4 = punto4 * -1
                punto7 = punto7 * -1
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    
    STR = "select * from " & msCajaTemp & " order by fecha"
    RptLisMovCajas.lblFecha = Date
    RptLisMovCajas.lblTitulo = "Listado de Mov. de Cajas del " & CStr(dtfechad) & " al " & CStr(dtfechah)
    RptLisMovCajas.lblsaldo = TraigosaldoAnterior()
    RptLisMovCajas.lblfechasaldo = dtfechad
    
    RptLisMovCajas.lbl11 = punto1
    RptLisMovCajas.lbl22 = punto2
    RptLisMovCajas.lbl33 = punto3
    RptLisMovCajas.lbl44 = punto4
    RptLisMovCajas.lbl55 = punto5
    RptLisMovCajas.lbl66 = punto6
    RptLisMovCajas.lbl77 = punto7
    RptLisMovCajas.lbl88 = punto8
    RptLisMovCajas.lbl99 = punto9
    
    RptLisMovCajas.Data.Connection = DataEnvironment1.Sistema
    RptLisMovCajas.PageSettings.Orientation = ddOLandscape
    RptLisMovCajas.Data.Source = STR
       
    RptLisMovCajas.Show

End Sub

Function buscotipopago(tipo As Long) As Long
    Dim rs1 As New ADODB.Recordset

    If optbuscocaja = True Then
    
    Else
        If optef3 = True Then
            If optsolofvo = True Then
            
            Else
                If optsoloch3 = True Then

                Else
                    'ACA VAN EFEC / CH 3º
                    
                End If
            End If
        Else
            If opttodos = True Then

            Else
                'ACA VAN SOLO BANCOS
                
            End If
        End If
    End If

    rs1.Open "select * from formaspago where codigo = " & tipo & " and  activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF Then
        If rs1!tarjeta = True Then
            buscotipopago = 1
        Else
            If rs1!Efectivo = True Then
                buscotipopago = 2
            Else
                buscotipopago = 3
            End If
        End If
    End If
    rs1.Close
    Set rs1 = Nothing
End Function

Function TraigosaldoAnterior() As Double
    Dim rs1 As New ADODB.Recordset
    Dim tot As Double
    
    Dim tI As Double, tE As Double
'    tI = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1 and ing_egr = 'I' and fecha < " & ssFecha(dtfechad)))
'    tE = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1 and ing_egr = 'E' and fecha < " & ssFecha(dtfechad)))
    
    
'    tot = 0
'    rs1.Open "select * from movicaja where fecha < " & ssFecha(dtfechad) & " and  activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'    While Not rs1.EOF
'        If rs1!Ing_egr = "I" Then
'            tot = tot + rs1!Importe
'        Else
'            tot = tot - rs1!Importe
'        End If
'        rs1.MoveNext
'    Wend
'    rs1.Close
'    Set rs1 = Nothing
'
'    TraigosaldoAnterior = tot

    
    TraigosaldoAnterior = s2n(sumaimportes("I", s2n(txtcodcaja)) - sumaimportes("E", s2n(txtcodcaja)))


End Function
Private Function sumaimportes(Ing_egr As String, caja As Long)
    sumaimportes = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1 and ing_egr = '" & Ing_egr & "' and  caja = '" & caja & "' and fecha < " & ssFecha(dtfechad)))
End Function

Private Sub habilitocaja(habilito As Boolean)
    lblcaja.Visible = habilito
    txtcodcaja.Visible = habilito
    cmbcaja.Visible = habilito
    txtCaja.Visible = habilito
End Sub


Private Sub Form_Load()
    dtfechad = Date
    dtfechah = Date
End Sub

''Private Sub Form_Unload(cancel As Integer)
''    If msCajaTemp > "" Then
''        TablaTempBorrar msCajaTemp
''        msCajaTemp = ""
''    End If
''End Sub

Private Sub optbuscocaja_Click()
    habilitocaja (True)
    LimpioPrimero (False)
    LimpioSegundo (False)
End Sub

Private Sub optef3_Click()
    LimpioSegundo (True)
    optsolofvo = False
    optsoloch3 = False
    opttodosef3 = False
End Sub

Private Sub opttodas_Click()
    habilitocaja (False)
    LimpioPrimero (True)
    optef3 = False
    opttodos = False
    optsolobancos = False
End Sub

Private Sub txtcaja_GotFocus()
    txtCaja.SelStart = 0
    txtCaja.SelLength = Len(txtCaja.Text)
End Sub

Private Sub txtcodcaja_GotFocus()
    txtcodcaja.SelStart = 0
    txtcodcaja.SelLength = Len(txtcodcaja.Text)
End Sub
