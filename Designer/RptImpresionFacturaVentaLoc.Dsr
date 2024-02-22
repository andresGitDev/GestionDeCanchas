VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionFacturaVentaLoc 
   Caption         =   "ActiveReport1"
   ClientHeight    =   14790
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   18960
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   33443
   _ExtentY        =   26088
   SectionData     =   "RptImpresionFacturaVentaLoc.dsx":0000
End
Attribute VB_Name = "RptImpresionFacturaVentaLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
    If Me.txtCantidad = 0 Then
        Me.txtCantidad = ""
        Me.txtprecio = ""
        Me.txttotal = ""
    End If
    If Not mayoracero(Me.txtprecio) Then txtprecio = ""
    If Not mayoracero(Me.txttotal) Then Me.txttotal = txttotal '""
End Sub

Private Function mayoracero(que) As Boolean
    On Error Resume Next ' si el cdbl de basura da error, devuelve false
    mayoracero = (CDbl(que) > 0)
End Function
