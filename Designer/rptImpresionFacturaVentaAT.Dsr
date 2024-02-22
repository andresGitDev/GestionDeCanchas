VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptImpresionFacturaVentaAT 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12620
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22260
   _ExtentY        =   21414
   SectionData     =   "rptImpresionFacturaVentaAT.dsx":0000
End
Attribute VB_Name = "rptImpresionFacturaVentaAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim item As Long

Private Sub Detail_BeforePrint()
    If Me.txtCantidad = 0 Then
        Me.txtCantidad = ""
        Me.txtprecio = ""
        Me.txttotal = ""
    End If
    If Not mayoracero(Me.txtprecio) Then txtprecio = ""
    If Not mayoracero(Me.txttotal) Then Me.txttotal = txttotal '""
    If Left(Me.txtDescripcion, 4) <> "REF." Then
        item = item + 1
        ite = item
    Else
        ite = ""
    End If
End Sub

Private Function mayoracero(que) As Boolean
    On Error Resume Next ' si el cdbl de basura da error, devuelve false
    mayoracero = (CDbl(que) > 0)
End Function


Private Sub pagefooter_beforeprint()
    Dim cant As Long
    cant = 1
    If Left(txttotalfinal, 2) = "U$" Then cant = 3
    If Right(Me.txttotalfinal, Len(Me.txttotalfinal) - cant) = 0 Then
        Me.txtneto2 = "--"
        Me.txtneto = "--"
        Me.txtsub2 = "--"
        Me.txtsub = "--"
        Me.txtivains2 = "--"
        Me.txtivains = "--"
        Me.txtIvaP2 = "--"
        Me.txtIvaP = "--"
        lblimp.Visible = True
        Label3.Visible = False
        Me.txttotalfinal = "--"
    Else
        lblimp.Visible = True
        Label3.Visible = True
    End If
End Sub

