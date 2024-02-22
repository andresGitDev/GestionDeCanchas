VERSION 5.00
Begin VB.Form FrmLisPosicionIva 
   BackColor       =   &H00745134&
   Caption         =   "Listado de Posicion de Iva"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   5355
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
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
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
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1035
      Width           =   975
   End
   Begin VB.ComboBox cmbmes 
      Height          =   315
      ItemData        =   "FrmLisPosicionIva.frx":0000
      Left            =   870
      List            =   "FrmLisPosicionIva.frx":0028
      TabIndex        =   1
      Top             =   270
      Width           =   1935
   End
   Begin VB.ComboBox cmbaño 
      Height          =   315
      ItemData        =   "FrmLisPosicionIva.frx":0091
      Left            =   3765
      List            =   "FrmLisPosicionIva.frx":00AA
      TabIndex        =   0
      Top             =   270
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00FFFFFF&
      Height          =   870
      Left            =   75
      Top             =   60
      Width           =   5235
   End
   Begin VB.Label Label1 
      BackColor       =   &H00745134&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   270
      TabIndex        =   6
      Top             =   300
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00745134&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3300
      TabIndex        =   5
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "FrmLisPosicionIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
'On Error GoTo UFAlistado

Dim rs As New ADODB.Recordset
Dim NetoAv As Double
Dim NetoBv As Double
Dim Totalv As Double
Dim Netov As Double
Dim Iva21v As Double
Dim Iva27v As Double
Dim Iva10v As Double
Dim Retv As Double
Dim Exentov As Double
Dim Otrosv As Double
Dim NetoAc As Double
Dim NetoBc As Double
Dim Totalc As Double
Dim Netoc As Double
Dim Iva21c As Double
Dim Iva27c As Double
Dim Iva10c As Double
Dim Retc As Double
Dim Exentoc As Double
Dim Otrosc As Double

    If cmbaño.ListIndex <> -1 And cmbmes.ListIndex <> -1 Then
        rs.Open "select * from facturaventa where month(fecha)=" & (cmbmes.ListIndex + 1) & " and " & _
        "year(fecha)=" & Trim(cmbaño.Text) & " and activo=1 and (tipodoc='NCA' or tipodoc='NCB' or tipodoc='FAA' or tipodoc='FAB' or tipodoc='NDB' or tipodoc='NDA') order by nrofactura", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Do While Not rs.EOF
'            If rs!razonsocial = "CLUB ATLETICO RIVER PLATE ASOC CIVIL Y D" Then
'            MsgBox "as"
'            End If
                Select Case Trim(rs!TIPODOC)
                    Case "FAA", "NDA"
                        Totalv = Totalv + rs!Total
                        NetoAv = NetoAv + rs!Neto
                        Netov = Netov + rs!Neto
                        If rs!PORCENTAJEiva = 0.21 And rs!tipoiva <> 8 Then   ' 8 es el Exento Ley
                               Iva21v = Iva21v + s2n(rs!Iva)
                        End If
                        If rs!PORCENTAJEiva = 0.105 Then
                                Iva10v = Iva10v + s2n(rs!Iva)
                        End If
                        If rs!PORCENTAJEiva = 0.27 Then
                                Iva27v = Iva27v + s2n(rs!Iva)
                        End If
                        
                    Case "FAB", "NDB"
                        Totalv = Totalv + rs!Total
                        NetoBv = NetoBv + rs!Neto
                        Netov = Netov + rs!Neto
                        
                        If rs!PORCENTAJEiva = 0.21 Then
                            Iva21v = Iva21v + s2n(rs!Iva)
                        End If
                        
                        If rs!PORCENTAJEiva = 0.105 Then
                            Iva10v = Iva10v + s2n(rs!Iva)
                        End If
                        
                        If rs!PORCENTAJEiva = 0.27 Then
                                Iva27v = Iva27v + s2n(rs!Iva)
                        End If
                        
                    Case "NCA"
                        Netov = Netov - rs!Neto
                        Totalv = Totalv - rs!Total
                        NetoAv = NetoAv - rs!Neto
                        If rs!PORCENTAJEiva = 0.21 Then
                            Iva21v = Iva21v - s2n(rs!Iva)
                        End If
                         If rs!PORCENTAJEiva = 0.105 Then
                            Iva10v = Iva10v - s2n(rs!Iva)
                         End If
                        If rs!PORCENTAJEiva = 0.27 Then
                           Iva27v = Iva27v - s2n(rs!Iva)
                        End If
                        
                    Case "NCB"
                        Netov = Netov - rs!Neto
                        Totalv = Totalv - rs!Total
                        NetoBv = NetoBv - rs!Neto
                        
                        If rs!PORCENTAJEiva = 0.21 Then
                            Iva21v = Iva21v - s2n(rs!Iva)
                        End If
                        
                        If rs!PORCENTAJEiva = 0.105 Then
                           Iva10v = Iva10v - s2n(rs!Iva)
                        End If
                        
                        If rs!PORCENTAJEiva = 0.105 Then
                            Iva27v = Iva27v - s2n(rs!Iva)
                        End If
                    End Select
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
        ' RETENCIONES
        rs.Open "select sum(total) as tot from facturaventa where month(fecha)=" & (cmbmes.ListIndex + 1) & " and year(fecha)=" & Trim(cmbaño.Text) & " and tipodoc='RET' and activo=1", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not IsNull(rs!tot) Then
            Retv = Retv + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptPosicionIva.lbltitulo = "POSICION DE IVA - PERIODO " & Trim(cmbmes.Text) & "/" & Trim(cmbaño.Text)
        
        'TOTALES DE VTA
        RptPosicionIva.lbl10vta = Format(Iva10v, "standard")
        RptPosicionIva.lbl21vta = Format(Iva21v, "standard")
        RptPosicionIva.lbl27vta = Format(Iva27v, "standard")
        RptPosicionIva.lblexentovta = Format(Exentov, "standard")
        RptPosicionIva.lblnetoavta = Format(NetoAv, "standard")
        RptPosicionIva.lblnetobvta = Format(NetoBv, "standard")
        RptPosicionIva.lblnetovta = Format(Netov, "standard")
        RptPosicionIva.lblotrosvta = Format(Otrosv, "standard")
        RptPosicionIva.lblretvta = Format(Retv, "standard")
        RptPosicionIva.lbltotvta = Format(Totalv, "standard")
        RptPosicionIva.lbltotalvta = Format(Iva10v + Iva21v + Iva27v - Retv, "standard")
        
        
        'COMPRAS
        rs.Open "select compras.*,Ivas.letra FROM COMPRAS INNER JOIN Ivas ON COMPRAS.TIPOIVA = Ivas.codigo where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Trim(cmbaño.Text) & " and compras.activo=1 and(tipodoc='N/C' or tipodoc='N/D' or tipodoc='FAC') ", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic      'and ivas.letra='A'
        If Not rs.EOF Then
            Do While Not rs.EOF
                Select Case Trim(rs!TIPODOC)
                    Case "FAC", "N/D"
                        Totalc = Totalc + rs!Total
                        Netoc = Netoc + rs!Neto
                        Iva21c = Iva21c + rs!iva_21
                        Iva10c = Iva10c + rs!iva_10
                        Iva27c = Iva27c + rs!iva_27
                        Exentoc = Exentoc + rs!EXENTO
                        Otrosc = Otrosc + rs!ibcapital + rs!ibprovincia + rs!imp_int + rs!ret_gan + rs!der_est
                        Retc = Retc + rs!percepc + rs!iva_9
                    Case "N/C"
                        Netoc = Netoc - rs!Neto
                        Totalc = Totalc - rs!Total
                        Iva21c = Iva21c - rs!iva_21
                        Iva10c = Iva10c - rs!iva_10
                        Iva27c = Iva27c - rs!iva_27
                        Exentoc = Exentoc - rs!EXENTO
                        Otrosc = Otrosc - rs!ibcapital - rs!ibprovincia - rs!imp_int - rs!ret_gan - rs!der_est
                        Retc = Retc - rs!percepc - rs!iva_9
                End Select
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select transcom.*,Ivas.letra FROM transcom INNER JOIN Ivas ON transcom.TIPOIVA = Ivas.codigo  where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Trim(cmbaño.Text) & " and transcom.activo=1 and(tipodoc='N/C' or tipodoc='N/D' or tipodoc='FAC') ", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic        'and ivas.letra='A'
        If Not rs.EOF Then
            Do While Not rs.EOF
                Select Case Trim(rs!TIPODOC)
                    Case "FAC", "N/D"
                        Totalc = Totalc + rs!Total
                        Netoc = Netoc + rs!Neto
                        Iva21c = Iva21c + rs!iva_21
                        Iva10c = Iva10c + rs!iva_10
                        Iva27c = Iva27c + rs!iva_27
                        Exentoc = Exentoc + rs!EXENTO
                        Otrosc = Otrosc + rs!ibcapital + rs!ibprovincia + rs!imp_int + rs!ret_gan + rs!der_est
                        Retc = Retc + rs!percepc + rs!iva_9
                    Case "N/C"
                        Netoc = Netoc - rs!Neto
                        Totalc = Totalc - rs!Total
                        Iva21c = Iva21c - rs!iva_21
                        Iva10c = Iva10c - rs!iva_10
                        Iva27c = Iva27c - rs!iva_27
                        Exentoc = Exentoc - rs!EXENTO
                        Otrosc = Otrosc - rs!ibcapital - rs!ibprovincia - rs!imp_int - rs!ret_gan - rs!der_est
                        Retc = Retc - rs!percepc - rs!iva_9
                End Select
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
        RptPosicionIva.lbl10comp = Format(Iva10c, "standard")
        RptPosicionIva.lbl21comp = Format(Iva21c, "standard")
        RptPosicionIva.lbl27comp = Format(Iva27c, "standard")
        RptPosicionIva.lblexentocomp = Format(Exentoc, "standard")
        RptPosicionIva.lblnetoacomp = Format(NetoAc, "standard")
        RptPosicionIva.lblnetobcomp = Format(NetoBc, "standard")
        RptPosicionIva.lblnetocomp = Format(Netoc, "standard")
        RptPosicionIva.lblotroscomp = Format(Otrosc, "standard")
        RptPosicionIva.lblretcomp = Format(Retc, "standard")
        RptPosicionIva.lbltotcomp = Format(Totalc, "standard")
        RptPosicionIva.lbltotalcomp = Format(Iva10c + Iva21c + Iva27c, "standard")
        RptPosicionIva.lblposvta = Format((Iva10v + Iva21v + Iva27v - Retv) - (Iva10c + Iva21c + Iva27c), "standard")
        RptPosicionIva.lblfecha = Date
        RptPosicionIva.Show
    Else
        MsgBox "Debe completar año y mes", 48, "Atención"
    End If
    
fin:
    Exit Sub
UFAlistado:
    MsgBox "err en listado"
    Resume fin
End Sub

Private Sub cmdcancelar_Click()
    cmbmes.ListIndex = -1
    cmbaño.ListIndex = -1
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbaño.ListIndex = 5
End Sub

