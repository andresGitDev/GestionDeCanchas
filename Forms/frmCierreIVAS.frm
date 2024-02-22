VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierreIVAS 
   Caption         =   "Cierre periodos iva"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "frmCierreIVAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   4395
      TabIndex        =   5
      Top             =   195
      Width           =   1335
   End
   Begin VB.Frame fraCierre 
      Height          =   1575
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   4260
      Begin MSComCtl2.DTPicker dtCierreCompras 
         Height          =   330
         Left            =   1890
         TabIndex        =   3
         Top             =   435
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         _Version        =   393216
         Format          =   81330177
         CurrentDate     =   40239
      End
      Begin MSComCtl2.DTPicker dtCierreVentas 
         Height          =   330
         Left            =   1890
         TabIndex        =   4
         Top             =   840
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   582
         _Version        =   393216
         Format          =   81330177
         CurrentDate     =   40239
      End
      Begin VB.Label Label2 
         Caption         =   "Cierre Iva VENTAS"
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   900
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre Iva COMPRAS"
         Height          =   255
         Left            =   255
         TabIndex        =   1
         Top             =   495
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCierreIVAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCierreCompras As Date
Private mCierreVentas As Date

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True
End Sub

Private Sub Form_Load()
    CargoCierres
End Sub

Private Sub CargoCierres()
    mCierreCompras = VerDatoEmpresa("CierreCompras") 'obtenerParametro("CierreIvaCompras")
    mCierreVentas = VerDatoEmpresa("CierreVentas")
    dtCierreCompras = mCierreCompras
    dtCierreVentas = mCierreVentas
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo ufaChe
    
    If dtCierreVentas <> mCierreVentas Or _
        dtCierreCompras <> mCierreCompras Then
        If MsgBox("¿Cambiar cierre?", vbInformation + vbYesNo) = vbYes Then
            
            If dtCierreVentas <> mCierreVentas Then
                DataEnvironment1.Sistema.Execute _
                " update datosempresa set CierreVentas = " & ssFecha(dtCierreVentas)
            End If
                
            If dtCierreCompras <> mCierreCompras Then
                DataEnvironment1.Sistema.Execute _
                " update datosempresa set CierreCompras = " & ssFecha(dtCierreCompras)
            End If
            
            MsgBox "Guardado", vbInformation
        End If
    End If
ufaChe:
    
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub
