VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmNumeracion 
   Caption         =   "Parametros sistema"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10875
   Icon            =   "frmNumeracion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid griPercIIBB 
      Height          =   2190
      Left            =   120
      TabIndex        =   27
      Top             =   1095
      Width           =   3465
      _cx             =   6112
      _cy             =   3863
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox txtNoMod 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   2280
      Width           =   1395
   End
   Begin VB.TextBox txtNoMod 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   1800
      Width           =   1395
   End
   Begin VB.TextBox txtNoMod 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox txtNoMod 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   840
      Width           =   1395
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   7
      Left            =   5985
      TabIndex        =   17
      Top             =   4230
      Visible         =   0   'False
      Width           =   1515
   End
   Begin Gestion.ucBotonera ucMenu 
      Height          =   1605
      Left            =   30
      TabIndex        =   16
      Top             =   4770
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   2831
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   6
      Left            =   5985
      TabIndex        =   12
      Top             =   3750
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   5
      Left            =   5985
      TabIndex        =   10
      Top             =   3270
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   4
      Left            =   5985
      TabIndex        =   8
      Top             =   2790
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   3
      Left            =   5985
      TabIndex        =   7
      Top             =   2310
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   2
      Left            =   5985
      TabIndex        =   5
      Top             =   1830
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   1
      Left            =   5985
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Height          =   375
      Index           =   0
      Left            =   5985
      TabIndex        =   1
      Top             =   870
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Coeficiente Perc IIBB"
      Height          =   390
      Left            =   195
      TabIndex        =   28
      Top             =   765
      Width           =   3300
   End
   Begin VB.Label lblNoMod 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      Height          =   255
      Index           =   3
      Left            =   7725
      TabIndex        =   26
      Top             =   2340
      Width           =   1395
   End
   Begin VB.Label lblNoMod 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      Height          =   255
      Index           =   2
      Left            =   7665
      TabIndex        =   24
      Top             =   1860
      Width           =   1395
   End
   Begin VB.Label lblNoMod 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      Height          =   255
      Index           =   1
      Left            =   7665
      TabIndex        =   21
      Top             =   1380
      Width           =   1395
   End
   Begin VB.Label lblNoMod 
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      Height          =   255
      Index           =   0
      Left            =   7665
      TabIndex        =   19
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   7
      Left            =   4305
      TabIndex        =   18
      Top             =   4230
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimo Numero Impreso:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   330
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "no visibles, se habilitan por prg"
      Height          =   375
      Left            =   9060
      TabIndex        =   14
      Top             =   315
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   6
      Left            =   4305
      TabIndex        =   13
      Top             =   3750
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   5
      Left            =   4305
      TabIndex        =   11
      Top             =   3270
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   4
      Left            =   4305
      TabIndex        =   9
      Top             =   2790
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   3
      Left            =   4365
      TabIndex        =   6
      Top             =   2310
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   2
      Left            =   4365
      TabIndex        =   4
      Top             =   1830
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   1
      Left            =   4305
      TabIndex        =   2
      Top             =   1350
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      Height          =   375
      Index           =   0
      Left            =   4305
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmNumeracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Lito Explicit  14/10/4

' Tabla parametros y campos configurads en ModuloLisistema
'
Private g As LiGrilla
Private gCODI As Long
Private gCATE As Long
Private gCOEF As Long

Private mIndice As Long
Private mArr() As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    mIndice = 0
'    carga "Factura A : ", CAMPO_BS_NroFACTURA_A
'    carga "Factura B : ", CAMPO_BS_NroFACTURA_B
'    carga "Remito : ", CAMPO_BS_NroREMITO
'    carga "Orden Pago : ", CAMPO_BS_OrdenPago
    carga "Aj Cred Prov :", CAMPO_BS_APC
    carga "Aj Deb  Prov :", CAMPO_BS_APD
    carga "Base Perc IIBB", CAMPO_BS_BaseIIBB
    
'    carga "Ejercicio :", CAMPO_BS_EJERCICIO
'    carga "Recibo : ",
'    carga "Orden Compra : ",
    
    cargaNoMod
    inigrilla
    ucMenu.init False, False, True, False, False
    ucMenu.BuscarOK
End Sub

Private Sub carga(titu, campo)
    On Error Resume Next
    ReDim Preserve mArr(mIndice + 1)
    
    Lbl(mIndice) = titu
    mArr(mIndice) = campo
    txt(mIndice) = obtenerParametro(campo)
    Lbl(mIndice).Visible = True
    txt(mIndice).Visible = True
    
    mIndice = mIndice + 1
    
    '---
    
    cargaGrilla
    
End Sub

Private Sub cargaNoMod()
    lblNoMod(0).caption = "Factura A"
    txtNoMod(0).Text = nSinNull(obtenerDeSQL("select max(NroFactura) from FacturaVenta where (tipoDoc = 'FAA' or tipoDoc = 'NCA' or tipoDoc = 'NDA') "))
    
    lblNoMod(1).caption = "Factura B"
    txtNoMod(1).Text = nSinNull(obtenerDeSQL("select max(NroFactura) from FacturaVenta where (tipoDoc = 'FAB' or tipoDoc = 'NCB' or tipoDoc = 'NDB') "))
    
    lblNoMod(2).caption = "Remito "
    txtNoMod(2).Text = nSinNull(obtenerDeSQL("select max(numero ) from RemitoVenta "))

    lblNoMod(3).caption = "O.Pago/a Cuenta"
    txtNoMod(3).Text = nuevoCodigoOP() - 1 'obtenerDeSQL("select max(nro ) from Rec_comp")

End Sub
Private Sub inigrilla()
    Set g = New LiGrilla
    With g
        .init griPercIIBB
        gCODI = .AddCol("Codigo", "H")
        gCATE = .AddCol("Categoria                          ")
        gCOEF = .AddCol("Coeficiente", "N", 4)
    End With
    cargaGrilla
End Sub
Private Sub cargaGrilla()
    Dim rs As New ADODB.Recordset, i As Long
    With rs
        rs.Open "select * from ivas where activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        g.Borrar
        While Not .EOF
            i = g.addRow()
            g.tx i, gCODI, !codigo
            g.tx i, gCATE, !DESCRIPCION
            g.tx i, gCOEF, !CoefPercIIBB
            .MoveNext
        Wend
    End With
End Sub
Private Sub GrabaGrilla()
    Dim i As Long
    For i = 1 To g.rows - 1
        If g.tx(i, gCOEF) = "-null-" Then
            g.tx i, gCOEF, "0"
        End If
        DataEnvironment1.Sistema.Execute "update ivas set CoefPercIIBB = " & x2s(g.tx(i, gCOEF)) & " where codigo = '" & x2s(g.tx(i, gCODI)) & "' "
    Next i
End Sub

'--------------------- MENU -----------------------------
Private Sub ucMenu_Aceptar()
    Dim i As Long
    If confirma("Cambiar Numeracion: Seguro ? ") Then
        For i = 0 To mIndice - 1
            CambiarParametroN mArr(i), s2n(txt(i))
        Next i
        GrabaGrilla
        MsgBox "Grabado"
        ucMenu.AceptarOk
        ucMenu.BuscarOK 'truch
    End If
End Sub
Private Sub ucMenu_Cancelar()
        ucMenu.BuscarOK 'truch
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    Dim i As Long
    For i = 0 To mIndice - 1
        txt(i).enabled = sino
    Next i
    griPercIIBB.enabled = sino
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
'--------------------------------------------------
