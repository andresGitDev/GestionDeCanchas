VERSION 5.00
Begin VB.Form frmPresuVista 
   Caption         =   "Seleccion de Vista"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancel 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton mail 
      Caption         =   "E-Mail"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton pdf 
      Caption         =   "PDF"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Word 
      Caption         =   "Word"
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Imprimir 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmPresuVista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public np As Double

Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Imprimir_Click()
    ImprimirPedido2 CDbl(np)
End Sub

Private Sub mail_Click()
    frmPresuMail.Show vbModal
End Sub

Private Sub pdf_Click()
    Dim ruta As String
    
    ImprimirPedido2 CDbl(np), False
    
    
    If ExisteArch("C:\UDC Output Files\Presu.pdf") > 0 Then
        Kill ("C:\UDC Output Files\Presu.pdf")
    End If
    If ExisteArch("C:\UDC Output Files\Presu.bdf") > 0 Then
        Kill ("C:\UDC Output Files\Presu.bdf")
    End If
    
    RptPedidoCliente2.Printer.DeviceName = "Universal Document Converter" '"cutepdf writer" '"adobe pdf" '"bullzip pdf printer" '
    '.Printer.Port = "PDF995port"
    RptPedidoCliente2.documentName = "Presu.bdf"  'con esto pongo el nombre del archivo por defecto.
    ruta = "C:\UDC Output Files\" & "Presu.pdf" '.documentName 'App.Path
    RptPedidoCliente2.Printer.ToPage = 0
    'no lo uso ya que se configura en la impresora
    RptPedidoCliente2.Printer.FileName = "C:\UDC Output Files\" & RptPedidoCliente2.documentName 'ruta & "\" & .documentName
    '.PageSettings.PaperSize = 9
    RptPedidoCliente2.PrintReport False
        
        
    cierra = True
    Unload rptOrdenCompra
End Sub

Private Sub Word_Click()

    Dim MSWord As New Word.Application
    Dim documento As Word.Document
    Dim Parrafo As Paragraph
    
    ImprimirPedido2 CDbl(np), False
    
    Set documento = MSWord.Documents.Add
    Set Parrafo = documento.Paragraphs.Add
    
    Parrafo.Range.InsertAfter RptPedidoCliente2.documentName
    
    MSWord.Visible = True
    
    cierra = True
    Unload rptOrdenCompra
    
End Sub
