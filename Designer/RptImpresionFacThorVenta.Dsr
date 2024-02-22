VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionFactThorVenta 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8740
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   12110
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   21361
   _ExtentY        =   15416
   SectionData     =   "RptImpresionFacThorVenta.dsx":0000
End
Attribute VB_Name = "RptImpresionFactThorVenta"
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
    If Not mayoracero(Me.txttotal) Then txttotal = ""
End Sub

Private Function mayoracero(que) As Boolean
    On Error Resume Next ' si el cdbl de basura da error, devuelve false
    mayoracero = (CDbl(que) > 0)
End Function

