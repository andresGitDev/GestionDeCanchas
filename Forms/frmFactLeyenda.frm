VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmFactLeyenda 
   Caption         =   "Leyenda de factura"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6855
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2325
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6615
         _cx             =   11668
         _cy             =   4101
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFactLeyenda.frx":0000
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
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdAddCtasCompras 
         Height          =   315
         Left            =   6000
         Picture         =   "frmFactLeyenda.frx":003D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   345
      End
      Begin VB.CommandButton cmdDelCtasCompras 
         Height          =   315
         Left            =   6375
         Picture         =   "frmFactLeyenda.frx":05C7
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   345
      End
      Begin Gestion.ucCoDe ucTexto 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         CodigoWidth     =   1000
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmFactLeyenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NroFac As Long

Private Sub aceptar_Click()
    'Call aceptar
    leye = 1
    Me.Hide
End Sub

Public Sub AceptarLeyenda(codigo As Long)
    Dim i As Long
    Dim COD As Long
    
    i = 1
    COD = codigo 'nuevoCodigo("FacturaVenta", "codigo")
    If frmFacturaVenta.txtCodigo <> COD Then
        frmFacturaVenta.txtCodigo = COD
        Label1.caption = COD
    End If
    While grilla.rows > i
        DataEnvironment1.Sistema.Execute "insert into facturaventaLeyenda (fac,leyenda,activo) values (" & Trim(Label1.caption) & ",'" & Trim(grilla.TextMatrix(i, 1)) & "',1)"
        i = i + 1
    Wend
    leye = 1
    Unload Me
End Sub

Private Sub Cancelar_Click()
    leye = 0
    Unload Me
End Sub

Private Sub cmdAddCtasCompras_Click()
    grilla.AddItem ucTexto.codigo & Chr(9) & ucTexto.DESCRIPCION
End Sub

Private Sub cmdDelCtasCompras_Click()
    grilla.RemoveItem (grilla.Row)
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    ini
    'ucTexto.ini "Select dbo.rtf2txt(descripcion) from texto Where id= '###'", _
                        "Select id as [ Id    ], dbo.rtf2txt(Descripcion) as [ Descripcion                                                ] From texto Where ACTIVO = 1", False
                        
    ucTexto.ini "Select leyenda from factleyenda Where id= '###'", _
                        "Select id as [ Id    ], leyenda as [ Descripcion                                                ] From factleyenda ", False
                        
    rs.Open "select * from facturaventaleyenda where fac=" & NroFac, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    While Not rs.EOF
        grilla.AddItem rs!ID & Chr(9) & rs!leyenda
        rs.MoveNext
    Wend
End Sub

Private Sub ini()
    grilla.cols = 0
    grilla.cols = 2
    grilla.rows = 1
    grilla.ColHidden(0) = True
    grilla.ColWidth(0) = 2000
    grilla.ColWidth(1) = 6000
    grilla.TextMatrix(0, 0) = "id"
    grilla.TextMatrix(0, 1) = "Descripcion"
    'grilla.Editable = True
    grilla.Editable = flexEDKbdMouse
End Sub
