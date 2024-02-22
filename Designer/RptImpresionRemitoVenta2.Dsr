VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionRemitoVenta2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   13550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13980
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24659
   _ExtentY        =   23901
   SectionData     =   "RptImpresionRemitoVenta2.dsx":0000
End
Attribute VB_Name = "RptImpresionRemitoVenta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tot As Double

'Private Sub Detail_AfterPrint()
'    Dim rs As New ADODB.Recordset
'
'    rs.Open "select precio from remitoportedetalle where numero=" & frmRemitoPorte.TxtRemitoNumero & " and producto='" & Trim(txtdescripcion) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'    If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!precio) Or IsEmpty(rs!precio) Then
'        total = 0
'    Else
'        total = rs!precio
'    End If
'    Set rs = Nothing
'End Sub

Private Sub Detail_BeforePrint()
'    Dim rs As New ADODB.Recordset
        
'    rs.Open "select precio from remitoportedetalle where numero=" & remiCarta & " and producto='" & Trim(txtdescripcion) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'    If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!precio) Or IsEmpty(rs!precio) Then
'        total = 0
'    Else
'        total = rs!precio
'        tot = tot + total
'    End If
'    Set rs = Nothing
    
    Dim r, rFactor As Double, rCargar As Double

    Set r = Nothing
    r = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.uparte=f.ufcodigo where p.codigo=" & ssTexto(txtDescripcion)) 'sstexto(dPRODUCTO))
    If IsNull(r) Or IsEmpty(r) Then
        rFactor = 1
    Else
        rFactor = r
    End If
    rCargar = rFactor * txtCantidad 'dCantidad
    Total = rCargar
    
End Sub

Private Sub pagefooter_beforeprint()
    Dim rs As New ADODB.Recordset
    
'    rs.Open "select (cantidad*precio) as valor from remitoportedetalle where numero=" & remiCarta & " and producto='" & Trim(txtdescripcion) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    rs.Open "select (cantidad*precio) as valor from remitoportedetalle where numero=" & remiCarta, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'    If IsNull(rs!valor) Or IsEmpty(rs!valor) Then
'        valor = 0 'tot
'    Else
'        valor = rs!valor
'    End If
    While Not rs.EOF
        Valor = Valor + s2n(nSinNull(rs!Valor))
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub

'Private Sub PageHeader_BeforePrint()
'    If Trim(LblIva) = "INSCRIPTO" Then
'        Label26.Visible = True
'        Label25.Visible = False
'    Else
'        Label25.Visible = True
'        Label26.Visible = False
'    End If
'End Sub

