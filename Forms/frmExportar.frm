VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExportar 
   Caption         =   "Exportar"
   ClientHeight    =   3585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   Icon            =   "frmExportar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraIibb 
      Caption         =   "Tipo de Iibb"
      Height          =   900
      Left            =   120
      TabIndex        =   10
      Top             =   1170
      Width           =   6090
      Begin VB.OptionButton Optioniibb3 
         Caption         =   "Retenciones Bancarias"
         Height          =   375
         Left            =   4020
         TabIndex        =   13
         Top             =   330
         Width           =   1950
      End
      Begin VB.OptionButton Optioniibb2 
         Caption         =   "Percepciones"
         Height          =   375
         Left            =   2085
         TabIndex        =   12
         Top             =   315
         Width           =   1350
      End
      Begin VB.OptionButton Optioniibb1 
         Caption         =   "Retenciones"
         Height          =   375
         Left            =   225
         TabIndex        =   11
         Top             =   315
         Value           =   -1  'True
         Width           =   1350
      End
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   360
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Fecha desde"
      Top             =   2550
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   635
      _Version        =   393216
      Format          =   62324737
      CurrentDate     =   40519
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   330
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "Ubicacion de la exportacion, cambie si requiere otra ubicacion."
      Top             =   2175
      Width           =   6105
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   900
      Left            =   4095
      Picture         =   "frmExportar.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2535
      Width           =   2040
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elija tipo de exportacion"
      Height          =   1020
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   6105
      Begin VB.OptionButton Option5 
         Caption         =   "Suss"
         Height          =   315
         Left            =   4980
         TabIndex        =   5
         Top             =   390
         Width           =   1005
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Sicore"
         Height          =   315
         Left            =   3690
         TabIndex        =   4
         Top             =   390
         Width           =   1005
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Iva"
         Height          =   315
         Left            =   2655
         TabIndex        =   3
         Top             =   375
         Width           =   1020
      End
      Begin VB.OptionButton Option2 
         Caption         =   "IIBB"
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Top             =   375
         Width           =   1005
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ganancia"
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   360
      Left            =   2055
      TabIndex        =   9
      ToolTipText     =   "Fecha hasta"
      Top             =   2550
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   635
      _Version        =   393216
      Format          =   62324737
      CurrentDate     =   40519
   End
End
Attribute VB_Name = "frmExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExportar_Click()
If Option1 Then exGanancia
'If Option2 Then exIibb
'If Option3 Then exIva
'If Option4 Then exSicore
'If Option5 Then exSuss
End Sub

Private Function exGanancia()
'ID_Cuenta_R_RET_GAN_RR2784 = 168
Dim rsRecibos As New ADODB.Recordset
Dim i As Long, x As Long, yy, rContador As Long, tmp
Dim rCuitAgente As String, rFecha As String, rCodigoReg As String, rImporte As String, rCertificado As String

    rsRecibos.Open "select * from recibosretenciones where  (fecha >=" & ssFecha(dtDesde) & " and fecha<=" & ssFecha(dtHasta) & ") and idcuentasparam=" & ID_Cuenta_R_RET_GAN_RR2784, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If rsRecibos.EOF And rsRecibos.BOF Then
        MsgBox "No hay retenciones en este periodo.", vbInformation, "Informe"
        Exit Function
    End If
    
    Open txtUbicacion For Output As #1
    
    With rsRecibos
        rContador = 0
        .MoveFirst
        
        'rCuitAgente = Format(Replace(VerParametro(BS_CUIT_EMPRESA), "-", ""), "00000000000")
        
        For i = 0 To .RecordCount - 1
            rFecha = "01/01/1900"
            rCuitAgente = "00000000000"
            rCodigoReg = "078"
            rImporte = "000000000000.00"
            rCertificado = "000000000000"
                
                
            rFecha = rsRecibos!Fecha
            tmp = obtenerDeSQL("select cliente from recibos where iddoc=" & s2n(rsRecibos!iddoc))
            
            rCuitAgente = Format(Replace(sSinNull(obtenerDeSQL("select cuit from clientes where codigo=" & tmp)), "-", ""), "00000000000")
            'rCodigoReg = "000"
            rImporte = Format(s2n(rsRecibos!Importe), "000000000000.00")
            rCertificado = Format(s2n(rsRecibos!numero), "000000000000")
            
            If s2n(rImporte) > 0 Then
                Print #1, rCuitAgente & rFecha & rCodigoReg & rImporte & rCertificado
                rContador = rContador + 1
            End If
            .MoveNext
        Next
    End With
    Set rsRecibos = Nothing
    If rContador = 0 Then
        MsgBox "No se encontraron registros para este periodo. (" & qTipo & ")", vbSystemModal, "Informe"
    Else
        MsgBox "Se genero el archivo con " & rContador & " registro/s.(" & qTipo & ")", vbSystemModal, "Informe"
    End If
    Close #1

End Function

Private Function exIibb()
    'rec
    'ID_Cuenta_V_Perc_IB_ProvBsAs = 166
    'ID_Cuenta_R_IIBB = 170
    'ID_Cuenta_R_IIBB_Prov = 183
    
Dim rsDatos As New ADODB.Recordset
Dim i As Long, x As Long, yy, rContador As Long, rConsulta As String

Dim rCodigoJ As String, rCuitAgente As String, rFecha As String, rSucursal As String, rConstancia As String
Dim rTipoDoc As String, rLetra As String, rNroDoc As String, rImporte As String

    If Optioniibb1 Then rConsulta = "select 'R' as tipodoc,numero as certificado,* from recibosretenciones where  (fecha >=" & ssFecha(dtDesde) & " and fecha<=" & ssFecha(dtHasta) & ") and idcuentasparam in (" & ID_Cuenta_R_IIBB & "," & ID_Cuenta_R_IIBB_Prov & ")"
    If Optioniibb2 Then
        rConsulta = "select iibb.fecha,'F' as tipodoc,iibb.importe,p.codiibb  from iibbjurisdiccion iibb iiner join provincias p on iibb.codjur=p.codigo where  (iibb.fecha >=" & ssFecha(dtDesde) & " and iibb.fecha<=" & ssFecha(dtHasta) & ") " _
                & " union " _
                & " "
    End If
    If Optioniibb3 Then rConsulta = "select * from recibosretenciones where  (fecha >=" & ssFecha(dtDesde) & " and fecha<=" & ssFecha(dtHasta) & ") and idcuentasparam=" & ID_Cuenta_R_RET_GAN_RR2784
    
    rsDatos.Open rConsulta, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    
    
    If rsDatos.EOF And rsDatos.BOF Then
        MsgBox "No hay retenciones en este periodo.", vbInformation, "Informe"
        Exit Function
    End If
    
    Open txtUbicacion For Output As #1
    
    With rsDatos
        rContador = 0
        .MoveFirst
        rCuitAgente = "00000000000"
        rCuitAgente = Format(Replace(VerParametro(BS_CUIT_EMPRESA), "-", ""), "00000000000")
        
        For i = 0 To .RecordCount - 1
            rFecha = "01/01/1900"
            rImporte = "000000000000.00"
            rConstancia = "000000000000"
            rCodigoJ = "000"
            rSucursal = "000"
            rTipoDoc = "000"
            
            rFecha = !Fecha
            rImporte = Format(s2n(!Importe), "000000000000.00")
            rConstancia = Format(s2n(!numero), "000000000000")
            
            If s2n(rImporte) > 0 Then
                Print #1, rCuitAgente & rFecha & rImporte & rConstancia
                rContador = rContador + 1
            End If
            .MoveNext
        Next
    End With
    Set rsDatos = Nothing
    If rContador = 0 Then
        MsgBox "No se encontraron registros para este periodo. (" & qTipo & ")", vbSystemModal, "Informe"
    Else
        MsgBox "Se genero el archivo con " & rContador & " registro/s.(" & qTipo & ")", vbSystemModal, "Informe"
    End If
    Close #1
    
End Function

Private Function exIva()
    'op
    'ID_Cuenta_P_RET_GAN_3ros = 164
    'ID_Cuenta_P_RET_IB_3ros = 165
    'ID_Cuenta_C_IB_PROV = 174
    'ID_Cuenta_C_IB_CAP = 163
    'ID_Cuenta_C_RET_GAN_CPRA = 162
    'ID_Cuenta_C_RET_IVA_CPRA = 173
    'ID_Cuenta_C_RG3337 = 175
    'ID_Cuenta_C_RG3431 = 176
    'rec
    'ID_Cuenta_V_Perc_IB_ProvBsAs = 166
    'ID_Cuenta_R_RetSegSoc = 171
    'ID_Cuenta_R_IIBB = 170
    'ID_Cuenta_R_IIBB_Prov = 183
    'ID_Cuenta_R_BONOS_CredFiscal = 169
    'ID_Cuenta_R_RET_GAN_RR2784 = 168
    'ID_Cuenta_R_RET_IVA_RG3125 = 167
    'ID_Cuenta_R_RET_Reparo = 186
End Function

Private Function exSicore()

End Function

Private Function exSuss()

End Function

Private Sub dtDesde_Change()
qUbicacion
End Sub

Private Sub dtDesde_Click()
qUbicacion
End Sub

Private Sub dtHasta_Change()
qUbicacion
End Sub

Private Sub dtHasta_Click()
qUbicacion
End Sub

Private Sub Form_Load()
dtDesde = CDate("01/01/" & Year(Date))
dtHasta = Date
txtUbicacion = ""
qUbicacion
End Sub

Private Function qFechas() As String
Dim xfecha1 As String, xfecha2 As String
xfecha1 = Replace(dtDesde, "/", "-")
xfecha2 = Replace(dtHasta, "/", "-")

qFechas = "_" & xfecha1 & "#" & xfecha2 & "_"

End Function

Private Function qTipo() As String
Dim xtipo As String
If Option1 Then xtipo = Option1.caption
If Option2 Then
    xtipo = Option2.caption
    If Optioniibb1 Then xtipo = xtipo & "#" & Optioniibb1.caption
    If Optioniibb2 Then xtipo = xtipo & "#" & Optioniibb2.caption
    If Optioniibb3 Then xtipo = xtipo & "#" & Optioniibb3.caption
End If
If Option3 Then xtipo = Option3.caption
If Option4 Then xtipo = Option4.caption
If Option5 Then xtipo = Option5.caption
xtipo = xtipo & "-"
qTipo = xtipo
End Function

Private Function qUbicacion()
txtUbicacion = "C:\" & qTipo & qFechas & ".txt"
'If Option1 Then
If Option2 Then
    fraIibb.Visible = True
Else
    fraIibb.Visible = False
End If
'If Option3 Then
'If Option4 Then
'If Option5 Then
End Function

Private Sub Option1_Click()
qUbicacion
End Sub

Private Sub Option2_Click()
qUbicacion
End Sub

Private Sub Option3_Click()
qUbicacion
End Sub

Private Sub Option4_Click()
qUbicacion
End Sub

Private Sub Option5_Click()
qUbicacion
End Sub

Private Sub Optioniibb1_Click()
qUbicacion
End Sub

Private Sub Optioniibb2_Click()
qUbicacion
End Sub

Private Sub Optioniibb3_Click()
qUbicacion
End Sub
