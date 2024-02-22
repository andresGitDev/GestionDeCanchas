VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptimpresionEtiquetaPro 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   21325
   _ExtentY        =   21308
   SectionData     =   "rptimpresionEtiquetaPro.dsx":0000
End
Attribute VB_Name = "rptimpresionEtiquetaPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
'se ejecuta antes de mostrar
    If Trim(Field3.Text) = "" And Trim(Field4.Text) = "" Then
        Shape2.Visible = False
        Field3.Visible = False
        Field4.Visible = False
        Label6.Visible = False
        Label5.Visible = False
        Field2.Visible = False
        Barcode2.Visible = False
        Line2.Visible = False
        Barcode4.Visible = False
        Label9.Visible = False
    End If
End Sub

