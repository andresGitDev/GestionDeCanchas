VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImpresoras 
   Caption         =   "RevisaImpresora"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "frmImpresoras.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpEtiquetas 
      Caption         =   "Etiquetas"
      Height          =   345
      Left            =   210
      TabIndex        =   5
      Top             =   1680
      Width           =   915
   End
   Begin VB.CommandButton cmdImpCetificado 
      Caption         =   "Certificado"
      Height          =   345
      Left            =   195
      TabIndex        =   4
      Top             =   1185
      Width           =   915
   End
   Begin VB.CommandButton cmdImpFacturas 
      Caption         =   "Facturas"
      Height          =   345
      Left            =   165
      TabIndex        =   3
      Top             =   690
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1335
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtImpCertificado 
      Height          =   390
      Left            =   1290
      TabIndex        =   2
      Top             =   1155
      Width           =   6015
   End
   Begin VB.TextBox txtImpEtiquetas 
      Height          =   390
      Left            =   1260
      TabIndex        =   1
      Top             =   1665
      Width           =   6030
   End
   Begin VB.TextBox txtImpFactura 
      Height          =   390
      Left            =   1290
      TabIndex        =   0
      Top             =   645
      Width           =   6000
   End
End
Attribute VB_Name = "frmImpresoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function traerNombre()
    CommonDialog1.ShowPrinter

    traerNombre = CommonDialog1.PrinterDefault
End Function

Private Sub cmdImpFacturas_Click()
    txtImpFactura = traerNombre
End Sub

Private Sub tttt()

        
'        For x = 0 To Printers.Count - 1
'        List1.AddItem str(x) + " " + Printers(x).DeviceName
'        Next
'        Set Printer = Printers(2)
'        Printer.Print "holA"
'        Printer.EndDoc
'
'        Printer.TrackDefault = True
'        Printer.Print "hola2"
'        Printer.EndDoc
        

End Sub

