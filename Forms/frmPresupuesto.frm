VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPresupuesto 
   Caption         =   "Ingreso de presupuesto"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTF 
      Height          =   330
      Left            =   9075
      TabIndex        =   59
      Top             =   7320
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPresupuesto.frx":0000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10800
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmPresupuesto.frx":0082
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      ToolTipText     =   "Borrar Item"
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10260
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmPresupuesto.frx":038C
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboTipoP 
      Height          =   315
      Left            =   4080
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtdireccionentrega 
      Height          =   285
      Left            =   975
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2490
      Width           =   6255
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2130
      Width           =   3135
   End
   Begin VB.ComboBox cmbMoneda 
      Height          =   315
      Left            =   8775
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3090
      Width           =   1575
   End
   Begin VB.TextBox txttel 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9915
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   650
      Width           =   1575
   End
   Begin VB.TextBox txtNro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1755
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   210
      Width           =   1815
   End
   Begin VB.ComboBox cmbformapago 
      Height          =   315
      Left            =   9315
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   2130
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4650
      Width           =   975
   End
   Begin VB.CommandButton cmdotro 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10455
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmPresupuesto.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4410
      Width           =   495
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1770
      Width           =   4635
   End
   Begin VB.TextBox Txtcontacto 
      Height          =   285
      Left            =   8295
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ComboBox cmbTransporte 
      Height          =   315
      Left            =   5415
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   2130
      Width           =   1815
   End
   Begin VB.TextBox txtnropedidocli 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9315
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtobs 
      Height          =   555
      Left            =   975
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2850
      Width           =   6255
   End
   Begin VB.TextBox txtprecio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7755
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4590
      Width           =   975
   End
   Begin VB.ComboBox cmbvendedor 
      Height          =   315
      Left            =   9360
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1750
      Width           =   1815
   End
   Begin VB.CheckBox chkPropio 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo Propio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3750
      Width           =   1755
   End
   Begin VB.CommandButton cmdBorraItem 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   10995
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmPresupuesto.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Borrar Item"
      Top             =   4410
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPresupuesto.frx":0CAA
      Left            =   5040
      List            =   "frmPresupuesto.frx":0CBA
      TabIndex        =   0
      Top             =   4575
      Width           =   1575
   End
   Begin VSFlex7LCtl.VSFlexGrid grillaproductos 
      Height          =   2235
      Left            =   195
      TabIndex        =   1
      Top             =   5010
      Width           =   11295
      _cx             =   19923
      _cy             =   3942
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPresupuesto.frx":0CE2
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
      Editable        =   1
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
   Begin Gestion.ucCuit uCuit 
      Height          =   315
      Left            =   7995
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   1275
      _extentx        =   2249
      _extenty        =   556
   End
   Begin Gestion.ucCoDe uCliente 
      Height          =   315
      Left            =   975
      TabIndex        =   3
      Top             =   1010
      Width           =   6075
      _extentx        =   10716
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   315
      Left            =   1215
      TabIndex        =   4
      Top             =   4170
      Width           =   7515
      _extentx        =   13256
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   300
      Left            =   9795
      TabIndex        =   22
      Top             =   150
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   76808193
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker dtfechaentrega 
      Height          =   300
      Left            =   8955
      TabIndex        =   23
      Top             =   4590
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   76808193
      CurrentDate     =   38098
   End
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   1545
      Left            =   0
      TabIndex        =   24
      Top             =   7515
      Width           =   11745
      _extentx        =   20717
      _extenty        =   2725
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Pedido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9000
         TabIndex        =   26
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label lblTotalPedi 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10500
         TabIndex        =   25
         Top             =   60
         Width           =   1095
      End
   End
   Begin Gestion.ucCoDe uEmisor 
      Height          =   315
      Left            =   960
      TabIndex        =   50
      Top             =   630
      Width           =   6075
      _extentx        =   10716
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin Gestion.ucCoDe uContacto 
      Height          =   315
      Left            =   960
      TabIndex        =   52
      Top             =   1360
      Width           =   6075
      _extentx        =   10716
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   1395
      Left            =   240
      TabIndex        =   56
      Top             =   8160
      Visible         =   0   'False
      Width           =   11295
      _cx             =   19923
      _cy             =   2461
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPresupuesto.frx":0DD4
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
   Begin Gestion.ucCoDe leyenda 
      Height          =   315
      Left            =   1260
      TabIndex        =   57
      Top             =   7680
      Visible         =   0   'False
      Width           =   7515
      _extentx        =   13256
      _extenty        =   556
      codigowidth     =   1000
   End
   Begin VB.Label Label22 
      Caption         =   "Leyenda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   58
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Contacto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   53
      Top             =   1360
      Width           =   1095
   End
   Begin VB.Label Label20 
      Caption         =   "Emisor :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   51
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Grupo de Producto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   49
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "Fecha Entrega :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8955
      TabIndex        =   47
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   735
      Left            =   7755
      Top             =   2850
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Cuit :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7515
      TabIndex        =   46
      Top             =   690
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Entrega:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   45
      Top             =   2490
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Localidad:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   2130
      Width           =   1035
   End
   Begin VB.Label Label10 
      Caption         =   "Moneda :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7875
      TabIndex        =   43
      Top             =   3090
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   7290
      Left            =   75
      Top             =   105
      Width           =   11535
   End
   Begin VB.Label Label8 
      Caption         =   "Tel:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9435
      TabIndex        =   42
      Top             =   690
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Nro Presupuesto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   41
      Top             =   210
      Width           =   1680
   End
   Begin VB.Label Label6 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   40
      Top             =   1010
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Forma de Pago :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7755
      TabIndex        =   39
      Top             =   2130
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9015
      TabIndex        =   38
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Producto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   195
      TabIndex        =   37
      Top             =   4170
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   195
      X2              =   11475
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   195
      TabIndex        =   36
      Top             =   4650
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   35
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "Contacto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7455
      TabIndex        =   34
      Top             =   1050
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Transporte :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4335
      TabIndex        =   33
      Top             =   2130
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "NroPedidoCliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7515
      TabIndex        =   32
      Top             =   2490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "Obs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   135
      TabIndex        =   31
      Top             =   2910
      Width           =   1575
   End
   Begin VB.Label lblunidad 
      Height          =   255
      Left            =   3675
      TabIndex        =   30
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label lblPrecio 
      Caption         =   "Precio :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6915
      TabIndex        =   29
      Top             =   4650
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Vendedor :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   28
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label lblEntrgado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2295
      TabIndex        =   27
      Top             =   4650
      Width           =   1155
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const gITEM = 0
Const gCANTIDAD = 1
Const gPRODUCTO = 2
Const gDESCRIPCION = 3
Const gPRECIO = 4
Const gFENTREGA = 5
Const gESTADO = 6
Const gSaldo = 7
Const gSALDOFAC = 8
Const gTOTAL = 9

  
Private sCampoExistencia As String
Private sTextosDetalle As String

Private Sub cboTipoP_LostFocus()
    set_uProd
End Sub

Private Sub chkPropio_Click()
    set_uProd
End Sub

Private Sub cmdBorraItem_Click()
    On Error Resume Next
    Dim i As Long
    i = grillaproductos.Row
    If i > -1 Then
        DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & txtNro.Text & " and codigo=" & i
        If s2n(grillaproductos.TextMatrix(i, gSALDOFAC)) < s2n(grillaproductos.TextMatrix(i, gCANTIDAD)) Then
            If confirma("producto parcialmente entregado, elimina ?") Then
                grillaproductos.RemoveItem (grillaproductos.Row)
            End If
        Else
            grillaproductos.RemoveItem (grillaproductos.Row)
            DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & Trim(frmPresupuesto.txtNro) & " and codigo=" & frmPresupuesto.grillaproductos.Row
        End If
    End If
    
    chkPropioEnabled True
    CalculaTotal
End Sub

Private Sub cmdotro_Click()
    
    If Trim(cboTipoP.Text) = "TEXTOS" Or Trim(cboTipoP.Text) = "TITULOS" Then
        AgregoTexto
    Else
        AgregoProd
    End If
    
End Sub
Private Function AgregoTexto() As Boolean
    Dim i As Long
    With grillaproductos
        .AddItem ""
        'If .rows = 1 Then .rows = 2
        'If .rows > 2 Or Trim$(.TextMatrix(0, gPRODUCTO)) > "" Then .rows = .rows + 1
        i = .rows - 1
        
        .TextMatrix(i, gITEM) = ""
        .TextMatrix(i, gCANTIDAD) = ""
        .TextMatrix(i, gPRODUCTO) = uProd.codigo
        .TextMatrix(i, gDESCRIPCION) = uProd.DESCRIPCION
        .TextMatrix(i, gPRECIO) = ""
        .TextMatrix(i, gFENTREGA) = ""
        .TextMatrix(i, gESTADO) = ""
        .TextMatrix(i, gSaldo) = ""
        .TextMatrix(i, gSALDOFAC) = ""
        
        If Trim(uProd.codigo) = "SUBTOTAL" Then
            .TextMatrix(i, gTOTAL) = SUByTOT(i)
        ElseIf Trim(uProd.codigo) = "TOTAL" Then
            .TextMatrix(i, gTOTAL) = SUByTOT
        Else
            .TextMatrix(i, gITEM) = uProd.DESCRIPCION
            .TextMatrix(i, gCANTIDAD) = uProd.DESCRIPCION
            .TextMatrix(i, gPRODUCTO) = uProd.DESCRIPCION
            .TextMatrix(i, gDESCRIPCION) = uProd.DESCRIPCION
            .TextMatrix(i, gPRECIO) = uProd.DESCRIPCION
            .TextMatrix(i, gFENTREGA) = uProd.DESCRIPCION
            .TextMatrix(i, gESTADO) = uProd.DESCRIPCION
            .TextMatrix(i, gSaldo) = uProd.DESCRIPCION
            .TextMatrix(i, gSALDOFAC) = uProd.DESCRIPCION
            .TextMatrix(i, gTOTAL) = uProd.DESCRIPCION
            
            .MergeCells = flexMergeFree
            .MergeRow(i) = True
            .Row = i
            .Col = 1
            .CellAlignment = flexAlignLeftCenter
        End If
    End With
    uProd.clear
    txtcantidad = "0"
    lblEntrgado = s2n(0)
    txtprecio = "0"
    dtfechaentrega.Value = Date
    uProd.SetFocus
    
End Function
Private Function SUByTOT(Optional Row As Long = 0) As Double
    Dim i As Long
    Dim Valor As Double
    Valor = 0
    If Row = 0 Then 'calcula total
        i = 0
        While i < grillaproductos.rows
            If grillaproductos.TextMatrix(i, gITEM) <> "" And IsNumeric(grillaproductos.TextMatrix(i, gITEM)) Then
                Valor = Valor + grillaproductos.TextMatrix(i, gTOTAL)
            End If
            i = i + 1
        Wend
    Else 'calcula subtotales
        i = Row - 1
        Do While i > -1
            If Trim(grillaproductos.TextMatrix(i, gDESCRIPCION)) <> "SUB TOTAL" Then
                If grillaproductos.TextMatrix(i, gITEM) <> "" Then
                    If IsNumeric(grillaproductos.TextMatrix(i, gITEM)) Then
                        Valor = Valor + grillaproductos.TextMatrix(i, gTOTAL)
                    End If
                End If
            Else
                Exit Do
            End If
            i = i - 1
        Loop
    End If
    SUByTOT = Valor
End Function
Private Function AgregoProd() As Boolean
    Dim i As Long
    Dim j As Long
    Dim Titulo As String
    Dim Ubica As Long
    Dim AuxAlias As String
    Dim a, aFactor As Double, aCargar As Double
    Dim rs As New ADODB.Recordset, codigomio As String, ssql As String, nuevoSaldo As Double, tengo As Double, queveo As Double, ssmsg As String, mjStock As Long
    Dim Item As Long
    
    AgregoProd = True
    
    If s2n(txtcantidad) <= 0 Then
        che "Especificar cantidad Mayor o igual a 0"
        AgregoProd = False
        Exit Function
    End If
    If s2n(txtcantidad) - s2n(lblEntrgado) < 0 Then
        che "No puede poner cantidad menor a lo ya entregado"
        AgregoProd = False
        Exit Function
    End If
    If Trim(uProd.DESCRIPCION) = "" Then
        che "Falta producto"
        AgregoProd = False
        Exit Function
    End If
    
    AuxAlias = (uProd.codigo)
    mjStock = ManejaStock(AuxAlias)
    
    ssql = "select codigo, componente, cantidad from formulas where activo = 1 and codigo = '" & AuxAlias & "'"
    rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    nuevoSaldo = (s2n(txtcantidad) - s2n(lblEntrgado))
    queveo = QueHay(AuxAlias, True)
    
    With grillaproductos
        .AddItem ""
        'If .rows = 1 Then .rows = 2
        'If .rows > 2 Or Trim$(.TextMatrix(0, gPRODUCTO)) > "" Then .rows = .rows + 1
        i = .rows - 1
        
        aFactor = 1
        aCargar = aFactor * s2n(txtcantidad)
                
        j = 0
        Item = 0
        Ubica = 0
        While j < .rows
            If .TextMatrix(j, gITEM) <> "" Then
                If IsNumeric(.TextMatrix(j, gITEM)) Then
                    If .TextMatrix(j, gITEM) = "1" Then Ubica = j
                    Item = .TextMatrix(j, gITEM)
                End If
            End If
            j = j + 1
        Wend
        If Ubica = 0 And .rows > 1 Then
            Ubica = .rows
        End If
                
        .TextMatrix(i, gITEM) = Item + 1
        .TextMatrix(i, gPRODUCTO) = uProd.codigo
        .TextMatrix(i, gDESCRIPCION) = uProd.DESCRIPCION
        .TextMatrix(i, gCANTIDAD) = s2n(aCargar)
        .TextMatrix(i, gPRECIO) = s2n(txtprecio, 4)
        .TextMatrix(i, gFENTREGA) = dtfechaentrega.Value
        .TextMatrix(i, gESTADO) = ESTADO_ADEUDADO
        .TextMatrix(i, gSaldo) = nuevoSaldo
        .TextMatrix(i, gSALDOFAC) = nuevoSaldo
        .TextMatrix(i, gTOTAL) = s2n(.TextMatrix(i, gCANTIDAD) * .TextMatrix(i, gPRECIO))
        
        If Item = 0 Then
            Titulo = "ITEM                CANT                                                                                                                                  UNITARIO " & IIf(Trim(cmbMoneda.Text) = "Pesos", "$    ", "U$S") & "         TOTAL " & IIf(Trim(cmbMoneda.Text) = "Pesos", "$    ", "U$S")
            .AddItem Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo & Chr(9) & Titulo, IIf(Ubica = 0, 0, Ubica - 1)
            .MergeCells = flexMergeFree
            .MergeRow(i) = True
        End If
        
    End With
    uProd.clear
    txtcantidad = "0"
    lblEntrgado = s2n(0)
    txtprecio = "0"
    dtfechaentrega.Value = Date
    uProd.SetFocus
    chkPropioEnabled True
    CalculaTotal
End Function

Private Sub Combo1_Click()
    If Combo1.Text = "Lista 1" Then
        txtprecio = obtenerDeSQL("select precio from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 2" Then
        txtprecio = obtenerDeSQL("select precio2 from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 3" Then
        txtprecio = obtenerDeSQL("select precio3 from producto where codigo='" & uProd.codigo & "'")
    ElseIf Combo1.Text = "Lista 4" Then
        txtprecio = obtenerDeSQL("select precio4 from producto where codigo='" & uProd.codigo & "'")
    End If
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_Load()
   
    CargaCombo cmbMoneda, "monedas", "descripcion", "codigo", ""
    CargaCombo cmbTransporte, "transportes", "descripcion", "codigo", ""
    CargaCombo cmbformapago, "formaspago", "descripcion", "codigo", ""
    CargaCombo cmbvendedor, "usuarios", "descripcion", "codigo", ""
    CargaCombo cboTipoP, "gruposproducto", "descripcion", "codigo", ""
    InicioGrilla
    dtFecha = Date
    
    HabilitoTxt False

    ucMenu.init True, True, True, True, True, "select * from pedidos_clientes2 where activo = 1 order by numero", DataEnvironment1.Sistema, True
    ucMenu.MsgConfirmaEliminar = "Desea Eliminar este pedido?"
    ucMenu.MsgConfirmaSalir = "Cerrar formulario ?"

    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    uEmisor.ini "select apellido+' '+nombre as descripcion from emisor where id = '###'", "select id as [ Codigo ],apellido as [ Apellido                   ], nombre as [ Nombre                        ] from emisor where activo = 1", False
    uContacto.ini "select apellido+' '+nombre as descripcion from contacto where id = '###'", "select id as [ Codigo ],apellido as [ Apellido                 ], nombre as [ Nombre                        ] from contacto where activo = 1", False ' and cliente=" & uCliente.codigo, False
    leyenda.ini "select dbo.rtf2txt(titulo) from texto where id = '###'", "select id as [ Codigo ], dbo.rtf2txt(titulo) as [ Titulo                        ] from texto where activo = 1", False
    set_uProd
    
    GeneraExistenciaCalculada
    If gEMPR_FormulaEsVirtual Then
        sCampoExistencia = " ExistenciaCalculada "
    Else
        sCampoExistencia = " Existencia "
    End If
    ArmoDetalle
End Sub

Private Function ArmoDetalle()
Dim rsTextos As New ADODB.Recordset, i As Long
With rsTextos
    .Open "select dbo.rtf2txt(titulo) as titulo from texto order by id", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
        sTextosDetalle = ""
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If i = 0 Then
                sTextosDetalle = !Titulo
            Else
                sTextosDetalle = sTextosDetalle & "|" & !Titulo
            End If
            .MoveNext
        Next
    End If
End With
Set rsTextos = Nothing

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Sub InicioGrilla()
    With grillaproductos
        .clear
        '.FormatString = "^Codigo                                    |<Descripcion                                                                        |>Cantidad |    Precio      |^Fecha Entrega  | Estado | Saldo   | Saldo-prueba- | Formula  "
        .rows = 0
        .cols = 10 '8 '7 '9 ????
        .ColHidden(gESTADO) = True
        .ColHidden(gSaldo) = True
        .ColHidden(gSALDOFAC) = True
        .ColHidden(gFENTREGA) = True
        .ColHidden(gPRODUCTO) = True
        .ColAlignment(gITEM) = flexAlignLeftCenter
        .ColAlignment(gTOTAL) = flexAlignLeftCenter
        .ColAlignment(gPRECIO) = flexAlignLeftCenter
    End With
End Sub

Private Sub grillaProductos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = gCANTIDAD Or Col = gPRECIO Then CalculaTotal
End Sub

Private Sub grillaproductos_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
If Col = gDESCRIPCION Then
    grillaproductos.ColComboList(gDESCRIPCION) = ""
End If
End Sub

Private Sub grillaProductos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim resu As Long
If KeyCode = 46 Then
    If grillaproductos.Row > 0 Then grillaproductos.RemoveItem (grillaproductos.Row)
End If
If KeyCode = 45 Then
    grillaproductos.AddItem ""
'    grillaproductos.Row = grillaproductos.rows - 1
'    grillaproductos.ColComboList(gDESCRIPCION) = sTextosDetalle
'    grillaproductos.cell(flexcpFontBold, grillaproductos.Row, gDESCRIPCION) = True
End If
If KeyCode = 112 Then 'esto es el F1
    'la tabla es ItemPedidoCliente2Texto
    frmTexto2.Show
    frmTexto2.Linea = grillaproductos.Row
    resu = s2n(obtenerDeSQL("select id from ItemPedidoCliente2Texto where pedido=" & txtNro.Text & " and codigo=" & grillaproductos.Row))
    If resu > 0 Then
        frmTexto2.txtDesc = obtenerDeSQL("select descripcion from ItemPedidoCliente2Texto where pedido=" & txtNro.Text & " and codigo=" & grillaproductos.Row)
    Else
        frmTexto2.txtDesc = grillaproductos.TextMatrix(grillaproductos.Row, gDESCRIPCION)
    End If
End If
End Sub

Private Sub grillaproductos_DblClick()
    On Error Resume Next
    Dim i As Long
    
    If ucMenu.estado <> ucbEditando Then Exit Sub
    
    With grillaproductos
        i = .Row
        If i = 0 Then Exit Sub
        
        DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & txtNro.Text & " and codigo=" & i
        
        If Trim$(uProd.DESCRIPCION) > "" Then
            If s2n(lblEntrgado) > 0 Then
                che "Hay un item cargado en la linea de edicion parcialmente entregado, no puedo sobreescribirlo"
                Exit Sub
            End If
            If Not confirma("Hay un item cargado en la linea de edicion, lo sobreescribe ?") Then
                Exit Sub
            End If
        End If
        
        uProd.codigo = Trim(.TextMatrix(i, gPRODUCTO))
        
        txtcantidad = .TextMatrix(i, gCANTIDAD)
        txtprecio = .TextMatrix(i, gPRECIO)
        dtfechaentrega = .TextMatrix(i, gFENTREGA)
        lblEntrgado = s2n(.TextMatrix(i, gCANTIDAD) - .TextMatrix(i, gSALDOFAC))
        DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & Trim(frmPresupuesto.txtNro) & " and codigo=" & frmPresupuesto.grillaproductos.Row
        .RemoveItem (grillaproductos.Row)
    
    End With
    CalculaTotal
End Sub

Private Sub CalculaTotal()
    Dim i As Long, tot As Double
    With grillaproductos
        For i = 1 To .rows - 1
            tot = tot + s2n(.TextMatrix(i, gCANTIDAD)) * s2n(.TextMatrix(i, gPRECIO), 4)
        Next i
    End With
    lblTotalPedi = s2n(tot)
End Sub

Private Sub txtCantidad_LostFocus()
    If Trim(cboTipoP.Text) = "TEXTOS" Or Trim(cboTipoP.Text) = "TITULOS" Then
    Else
        If Val(txtcantidad) <= 0 Then
            MsgBox "La cantidad debe ser mayor a 0", 48, "Atencion"
        End If
    End If
End Sub

Private Sub LimpioTxt()
    On Error Resume Next
    
    If txtNro <> "" Then
        DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & txtNro.Text
    End If
    
    lblEntrgado = s2n(0)
    FrmBorrarTxt Me
    uCliente.codigo = 0
    uEmisor.clear
    uContacto.clear
    uCuit.Text = ""
    uProd.clear
    
    txtNro = ""
    txttel = ""
    txtdireccion = ""
    txtlocalidad = ""
    Txtcontacto = ""
    dtFecha.Value = Date
    dtfechaentrega.Value = Date
    cmbTransporte.ListIndex = 0
    cmbformapago.ListIndex = 0
    cmbvendedor.ListIndex = 0
    txtdireccionentrega = ""
    txtnropedidocli = ""
    txtobs = ""
    cmbMoneda.ListIndex = 2
    chkPropio.Value = vbChecked
    txtcantidad = "0"
    txtprecio = "0.00"
    
    InicioGrilla
    CalculaTotal
End Sub

Private Sub HabilitoTxt(habilito As Boolean)
    Dim bloqueo
    bloqueo = Not habilito
    
    txtNro.Locked = bloqueo
    uCliente.enabled = habilito
    uProd.enabled = habilito
    Txtcontacto.Locked = bloqueo
    dtFecha.enabled = Not bloqueo
    dtfechaentrega.enabled = Not bloqueo
    cmbTransporte.Locked = bloqueo
    cmbformapago.Locked = bloqueo
    cmbvendedor.Locked = bloqueo
    chkPropioEnabled (Not bloqueo)
    txtdireccionentrega.Locked = bloqueo
    txtnropedidocli.Locked = bloqueo
    txtobs.Locked = bloqueo
    cmbMoneda.Locked = bloqueo
    txtcantidad.Locked = bloqueo
    txtprecio.Locked = bloqueo
    cmbMoneda.enabled = habilito
    cmdBorraItem.enabled = habilito
    cmdotro.enabled = habilito
End Sub

Private Sub CargoDatosCliente()
    On Error Resume Next
    Dim tmp
    
    tmp = obtenerDeSQL("select codigo, direccion, localidad, contacto, direccion_comercial, telefono_comercial, formapago, transporte, vendedor , cuit from clientes where codigo = " & uCliente.codigo)
    
    txtdireccion = sSinNull(tmp(1))
    txtlocalidad = sSinNull(tmp(2))
    Txtcontacto = sSinNull(tmp(3))
    txtdireccionentrega = sSinNull(tmp(4))
    txttel = sSinNull(tmp(5))
    cmbformapago.ListIndex = BuscarenComboS(cmbformapago, ObtenerDescripcion("formaspago", nSinNull(tmp(6))))
    cmbTransporte.ListIndex = BuscarenComboS(cmbTransporte, ObtenerDescripcion("transportes", nSinNull(tmp(7))))
    cmbvendedor.ListIndex = BuscarenComboS(cmbvendedor, ObtenerDescripcion("usuarios", nSinNull(tmp(8))))
    uCuit.Text = sSinNull(tmp(9))
    
End Sub
Private Sub txtprecio_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txttel_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtDireccion_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtlocalidad_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcontacto_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtdireccionentrega_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtnropedidocli_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtobs_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcodprod_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtDescripcion_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub
Private Sub txtprecio_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Function ConCodigoPropio() As Boolean
    ConCodigoPropio = (chkPropio.Value = vbChecked)
End Function


Private Function altaItems() As Boolean
On Error GoTo aimal
altaItems = True
    Dim fechaEntrega As Date, estado As String, i, COD As String, cant As Double, costo As Double, saldo As Double, formula As String, pDescripcion As String
    Dim Item As Long
    Dim Total As Double
    Dim InsItem As String, TipoItem As String
    Dim Linea As Long
    'SE ELIMINA TODOS LOS ITEMS
    InsItem = "delete from ItemPedidoCliente2 where pedido = " & x2s(txtNro)
    DataEnvironment1.Sistema.Execute InsItem
    
    'SE CARGA DE VUELTA TODOS LOS ITEM
    For i = 0 To grillaproductos.rows - 1
        grillaproductos.Row = i
        grillaproductos.Col = gPRODUCTO '0
        If ConCodigoPropio() Then
            If Trim(grillaproductos.Text) = Trim(grillaproductos.TextMatrix(i, gITEM)) Then
                COD = ""
            Else
                COD = grillaproductos.Text
            End If
        Else
            If Trim(grillaproductos.Text) = Trim(grillaproductos.TextMatrix(i, gITEM)) Then
                COD = ""
            Else
                COD = obtenerDeSQL("select producto from relacion_producto_cliente where productocliente = '" & grillaproductos.Text & "'")
            End If
        End If
        
        pDescripcion = grillaproductos.TextMatrix(i, gDESCRIPCION)
        
        grillaproductos.Col = gCANTIDAD
        cant = IIf((Trim(grillaproductos.TextMatrix(i, gITEM)) <> "" And IsNumeric(grillaproductos.TextMatrix(i, gITEM))), CDbl(s2n(grillaproductos.Text)), 0)
        
        grillaproductos.Col = gPRECIO
        costo = IIf((Trim(grillaproductos.TextMatrix(i, gITEM)) <> "" And IsNumeric(grillaproductos.TextMatrix(i, gITEM))), CDbl(s2n(grillaproductos.Text)), 0)
        
        grillaproductos.Col = gFENTREGA
        'fechaEntrega = IIf(grillaproductos.Text = "", Date, grillaproductos.Text)
        fechaEntrega = IIf((Trim(grillaproductos.TextMatrix(i, gITEM)) <> "" And IsNumeric(grillaproductos.TextMatrix(i, gITEM)) And Not grillaproductos.Text = ""), grillaproductos.Text, Date)
        
        If Trim(grillaproductos.TextMatrix(i, gITEM)) <> "" And IsNumeric(grillaproductos.TextMatrix(i, gITEM)) Then
            grillaproductos.Col = gESTADO
            estado = grillaproductos.Text
            
            grillaproductos.Col = gSaldo
            saldo = s2n(grillaproductos.Text)
            
            grillaproductos.Col = gITEM
            Item = s2n(grillaproductos.Text)
        Else
            estado = ESTADO_ENTREGADO
            saldo = 0
            Item = 0
        End If
        
        grillaproductos.Col = gTOTAL
        Total = IIf(IsNumeric(grillaproductos.Text), s2n(grillaproductos.Text), 0)
        
        If Trim(COD) = "" Then
            TipoItem = "Otro"
        Else
            TipoItem = "Det"
        End If
        
        If obtenerDeSQL("select id from itempedidocliente2texto where pedido=" & txtNro & " and codigo=" & i) <> "" Then
            InsItem = "INSERT INTO ITEMPEDIDOCLIENTE2 (PEDIDO,PRODUCTO, DESCRIPCION,CANTIDAD, FACTURAR,SALDO,ESTADO,PRECIO,FORMULA, FECHAENTREGA,TIPOITEM,ITEM,TOTAL) " _
                    & " VALUES (" & x2s(txtNro) & "," & ssTexto(COD) & "," & ssTexto(obtenerDeSQL("select descripcion from itempedidocliente2texto where pedido=" & txtNro & " and codigo=" & i)) & "," & x2s(cant) & "," & x2s(cant) & ", " & x2s(saldo) & "," & ssTexto(estado) & "," & x2s(costo) & "," & ssTexto(formula) & "," & ssFecha(fechaEntrega) & "," & ssTexto(TipoItem) & "," & Item & "," & x2s(Total) & ")"
        Else
            InsItem = "INSERT INTO ITEMPEDIDOCLIENTE2 (PEDIDO,PRODUCTO, DESCRIPCION,CANTIDAD, FACTURAR,SALDO,ESTADO,PRECIO,FORMULA, FECHAENTREGA,TIPOITEM,ITEM,TOTAL) " _
                    & " VALUES (" & x2s(txtNro) & "," & ssTexto(COD) & "," & ssTexto(pDescripcion) & "," & x2s(cant) & "," & x2s(cant) & ", " & x2s(saldo) & "," & ssTexto(estado) & "," & x2s(costo) & "," & ssTexto(formula) & "," & ssFecha(fechaEntrega) & "," & ssTexto(TipoItem) & "," & Item & "," & x2s(Total) & ")"
        End If
        DataEnvironment1.Sistema.Execute InsItem
        DataEnvironment1.Sistema.Execute "delete from ItemPedidoCliente2Texto where pedido=" & txtNro.Text & " and codigo=" & i
    Next i
Exit Function
aimal:
altaItems = False
End Function

Private Sub leerPedido(cual)
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset, ssql As String, tra As String
    
    Dim i As Long
    
    LimpioTxt
    
    ssql = "select * from pedidos_clientes2 where numero = " & cual & " "
    
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        txtNro = Format(!numero, "00000")
        uCliente.codigo = !cliente
        CargoDatosCliente
        uEmisor.codigo = nSinNull(!Emisor)
        uContacto.codigo = nSinNull(!contacto)
        dtFecha = !Fecha
        cmbformapago = ObtenerDescripcion("formasPago", !pago)
        cmbvendedor = ObtenerDescripcion("usuarios", !Vendedor)
        txtnropedidocli = !pedido_cli
        chkPropio.Value = IIf(!CODIGOPROPIO, vbChecked, vbUnchecked)
        tra = sSinNull(ObtenerDescripcion("transportes", !Transporte))
        
        txtobs = !observaciones
    
        .Close
        
'        ssql = "select * from itemPedidoCliente2 where pedido = " & Val(txtNro) & " order by codigo"
        ssql = "select producto,dbo.rtf2txt(descripcion) as descripcion,cantidad,precio,fechaentrega,estado,saldo,facturar,tipoitem,item,total from itemPedidoCliente2 where pedido = " & Val(txtNro) & " order by codigo"
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        i = 0
        While Not .EOF
            
            grillaproductos.rows = i + 1
            If s2n(!Item) > 0 Then 'cargo solo productos
                grillaproductos.TextMatrix(i, gPRODUCTO) = IIf(ConCodigoPropio(), !Producto, obtenerDeSQL("select productoCliente from Relacion_Producto_Cliente where producto = '" & !Producto & "' and cliente = " & uCliente.codigo))
                grillaproductos.TextMatrix(i, gITEM) = !Item
                grillaproductos.TextMatrix(i, gDESCRIPCION) = !DESCRIPCION
                grillaproductos.TextMatrix(i, gCANTIDAD) = !Cantidad
                grillaproductos.TextMatrix(i, gPRECIO) = !precio
                grillaproductos.TextMatrix(i, gFENTREGA) = !fechaEntrega
                grillaproductos.TextMatrix(i, gESTADO) = !estado
                grillaproductos.TextMatrix(i, gSaldo) = !saldo
                grillaproductos.TextMatrix(i, gSALDOFAC) = !facturar
                grillaproductos.TextMatrix(i, gTOTAL) = !Total
            Else 'cargo todo lo que es texto
                If Trim(!Producto) = "SUBTOTAL" Or Trim(!Producto) = "TOTAL" Or (Trim(!Producto) = "" And Trim(!DESCRIPCION) = "") Then
                    grillaproductos.TextMatrix(i, gITEM) = ""
                    grillaproductos.TextMatrix(i, gCANTIDAD) = ""
                    grillaproductos.TextMatrix(i, gPRODUCTO) = Trim(!Producto)
                    grillaproductos.TextMatrix(i, gDESCRIPCION) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gPRECIO) = ""
                    grillaproductos.TextMatrix(i, gFENTREGA) = ""
                    grillaproductos.TextMatrix(i, gESTADO) = ""
                    grillaproductos.TextMatrix(i, gSaldo) = ""
                    grillaproductos.TextMatrix(i, gSALDOFAC) = ""
                    grillaproductos.TextMatrix(i, gTOTAL) = IIf((Trim(!Producto) = "" And Trim(!DESCRIPCION) = ""), "", !Total)
                Else
                    grillaproductos.TextMatrix(i, gITEM) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gCANTIDAD) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gPRODUCTO) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gDESCRIPCION) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gPRECIO) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gFENTREGA) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gESTADO) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gSaldo) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gSALDOFAC) = Trim(!DESCRIPCION)
                    grillaproductos.TextMatrix(i, gTOTAL) = Trim(!DESCRIPCION)
                    
                    grillaproductos.MergeCells = flexMergeFree
                    grillaproductos.MergeRow(i) = True
                    grillaproductos.Row = i
                    grillaproductos.Col = 1
                    grillaproductos.CellAlignment = flexAlignLeftCenter
                End If
            End If
            
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
    If Trim(tra) > "" Then cmbTransporte = tra
    Set rs = Nothing
    
    HabilitoTxt False
    CalculaTotal
End Sub

Private Sub chkPropioEnabled(que As Boolean)
    chkPropio.enabled = grillaproductos.rows < 2 And que
End Sub

Private Function buscaRelProdClie(cliente As String)
    Dim ssql As String
    ssql = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
        & " from producto  " _
        & " inner join relacion_Producto_Cliente " _
        & " on producto.codigo = relacion_Producto_cliente.producto " _
        & " where cliente = " & cliente _
        & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
        & " order by producto"
    frmBuscar.MostrarSql (ssql)
End Function

Private Function FaltanCosas() As Boolean
    Dim i As Long
    
    FaltanCosas = True
        
    If uEmisor.codigo = 0 Then
        MsgBox "Falta cargar el emisor"
        uEmisor.SetFocus
        Exit Function
    End If
    
    If uCliente.codigo = 0 Then
        If uContacto.codigo = 0 Then
            MsgBox "Falta cargar el contacto"
            uContacto.SetFocus
            Exit Function
        End If
    End If
    
'    If uCliente.codigo = 0 Then
'        MsgBox "Falta cargar el cliente"
'        uCliente.SetFocus
'        Exit Function
'    End If
        
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Function
    End If
    
    With grillaproductos
        If .rows < 2 Then
            che "faltan datos en grilla"
            Exit Function
        End If
        
    End With

    FaltanCosas = False
End Function

Private Sub set_uProd()
    Dim sqlbuscar As String, sqldesc As String
    Dim rs As New ADODB.Recordset
    Dim whe As String
    
    If cboTipoP.Text <> "" Then
        rs.Open "select * from gruposproducto where descripcion='" & Trim(cboTipoP.Text) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!codigo) Or IsEmpty(rs!codigo) Then
            whe = ""
        Else
            If Trim(cboTipoP.Text) = "TEXTOS" Or Trim(cboTipoP.Text) = "TITULOS" Then
                whe = " and grupo='" & Trim(rs!codigo) & "'"
            Else
                whe = " and producto.grupo='" & Trim(rs!codigo) & "' "
            End If
        End If
    Else
        whe = ""
    End If

    If Trim(cboTipoP.Text) = "TEXTOS" Or Trim(cboTipoP.Text) = "TITULOS" Then
        sqldesc = "select dbo.rtf2txt(descripcion) as descripcion from texto where dbo.rtf2txt(titulo) = '###' "
        sqlbuscar = "select dbo.rtf2txt(titulo) as [ Titulo                 ], dbo.rtf2txt(descripcion) as [ Descripcion                                                 ] from texto where activo = 1 " & whe & " order by dbo.rtf2txt(titulo) "
    Else
        If ConCodigoPropio() Then    'propio
            sqldesc = "select descripcion from producto where codigo = '###' "
           sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 " & whe & " order by codigo "
        Else    'relCliente
            sqldesc = "select descripcion from producto  " _
                & " inner join relacion_Producto_Cliente " _
                & " on producto.codigo = relacion_Producto_cliente.producto " _
                & " where cliente = " & uCliente.codigo & " and productoCliente = '###'"
            sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
                & " from producto  " _
                & " inner join relacion_Producto_Cliente " _
                & " on producto.codigo = relacion_Producto_cliente.producto " _
                & " where cliente = " & uCliente.codigo _
                & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " & whe & "" _
                & " order by producto"
        End If
    End If
    uProd.ini sqldesc, sqlbuscar, True
    Set rs = Nothing
End Sub

Private Sub ucMenu_Imprimir()
    frmPresuVista.np = s2n(txtNro, 0)
    frmPresuVista.Show vbModal
    'ImprimirPedido2 s2n(txtNro, 0)
End Sub

Private Sub uContacto_cambio(codigo As Variant)
    Dim clie As Long
    'uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    clie = obtenerDeSQL("select cliente from contacto where id=" & uContacto.codigo)
    If clie <> 0 Then
        uCliente.codigo = clie
    Else
        uCliente.codigo = 0
    End If
End Sub

Private Sub uProd_cambio(codigo As Variant)
    txtprecio = precioProducto(CStr(codigo), ConCodigoPropio(), uCliente.codigo)
    Combo1.ListIndex = 0
End Sub
Private Sub uCliente_cambio(codigo As Variant)
    
    CargoDatosCliente
    set_uProd
    If uCliente.codigo = 0 Then
        If uContacto.codigo = 0 Then
            uContacto.ini "select apellido+' '+nombre as descripcion from contacto where id = '###'", "select id as [ Codigo ],apellido as [ Apellido                 ], nombre as [ Nombre                        ] from contacto where activo = 1 ", False
        End If
    Else
        If uContacto.codigo <> 0 Then
            If uCliente.codigo <> obtenerDeSQL("select cliente from contacto where id=" & uContacto.codigo) Then
                uContacto.ini "select apellido+' '+nombre as descripcion from contacto where id = '###'", "select id as [ Codigo ],apellido as [ Apellido                 ], nombre as [ Nombre                        ] from contacto where activo = 1 and cliente=" & codigo, False
            End If
        Else
            uContacto.ini "select apellido+' '+nombre as descripcion from contacto where id = '###'", "select id as [ Codigo ],apellido as [ Apellido                 ], nombre as [ Nombre                        ] from contacto where activo = 1 and cliente=" & codigo, False
        End If
    End If
End Sub

Private Function QueHay(quePro As String, Optional RestoReservado As Boolean)
    Dim hay As Double, reserva As Double, kk
    quePro = VerProductoMio(quePro, ConCodigoPropio())
    
    If RestoReservado Then
        QueHay = s2n(obtenerDeSQL("select (" & sCampoExistencia & " - ReservaCalculada) as quequeda from producto  where activo = 1 and codigo = '" & quePro & "' "))
    Else
        QueHay = s2n(obtenerDeSQL("select " & sCampoExistencia & " from producto  where activo = 1 and codigo = '" & quePro & "' "))
    End If
   
End Function

Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim np As Long
    Dim moneda As Long
    
    If FaltanCosas() Then Exit Sub
    
    moneda = obtenerDeSQL("select codigo from monedas where descripcion=" & ssTexto(cmbMoneda.Text)) 'BuscarenComboS(cmbMoneda, cmbMoneda.Text)
    cmbMoneda.ListIndex = BuscarenComboS(cmbMoneda, Const_PESOS)
    
    np = nuevoCodigo("Pedidos_Clientes2", "Numero")
    txtNro = np
    
    '**************************************
    DE_BeginTrans
    
        If ABMPedidoCliente("A", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtobs, 0, dtFecha, UsuarioSistema!codigo, uEmisor.codigo, uContacto.codigo, moneda) = False Then GoTo ufaErr
        If altaItems = False Then GoTo ufaErr
        'guardar leyendas
    
    DE_CommitTrans
    '**************************************'soy puto
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    frmPresuVista.np = CDbl(np)
    frmPresuVista.Show vbModal
'    ImprimirPedido2 CDbl(np)
    ucMenu.AceptarOk '"Numero = " & Trim(txtNro)

Exit Sub
ufaErr:
    DE_RollbackTrans
    MsgBox "Error al guardar presupuesto", vbCritical
End Sub

 Private Sub ucMenu_AceptarModi()
    Dim moneda As Long
    
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
                                                                                                                                                   
    If FaltanCosas() Then Exit Sub
    
    moneda = obtenerDeSQL("select codigo from monedas where descripcion=" & ssTexto(cmbMoneda.Text))

    DE_BeginTrans
    If ABMPedidoCliente("M", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtobs, 0, dtFecha, UsuarioSistema!codigo, 0, 0, moneda) = False Then GoTo ufaErr
    If altaItems = False Then GoTo ufaErr
    grabaBitacora "M", s2n(txtNro), "PedidosClientes"
    DE_CommitTrans
    
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk "Numero = " & Trim(txtNro)
Exit Sub
ufaErr:
    DE_RollbackTrans
    MsgBox "Error al guardar el pedido.", vbCritical
End Sub
Private Sub ucMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub ucMenu_Buscar()
    Dim s As String
    s = " select p.numero, p.Pedido_cli as PedidoClie, isnull(c.descripcion,'') as [ Cliente                              ],isnull(o.nombre+' '+o.apellido,'') as [ Contacto                  ], p.fecha as [ Fecha      ], max(i.saldo) as pendiente, max(i.facturar) as [PendientePrueba] " & _
        " from pedidos_clientes2 as p " & _
        " left outer join clientes as c on c.codigo = p.cliente " & _
        " inner join itempedidocliente2 as i on i.pedido = p.numero " & _
        " left outer join contacto as o on p.contacto=o.id " & _
        " where p.activo = 1 " & _
        " group by numero, Pedido_cli, c.descripcion, fecha,o.nombre,o.apellido " & _
        " order by numero desc "
    
    If frmBuscar.MostrarSql(s) = "" Then Exit Sub
    leerPedido Val(frmBuscar.resultado(1))
    ucMenu.BuscarOK "numero = " & txtNro
End Sub
Private Sub ucMenu_BuscarYa(que As Variant)
    If Not IsEmpty(obtenerDeSQL("select numero from pedidos_clientes2 where pedidos_clientes2.activo = 1 and numero = " & s2n(que))) Then
        leerPedido s2n(que)
        ucMenu.BuscarOK "numero = " & txtNro
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim sp
    sp = obtenerDeSQL("select * from remitoventadetalle d inner join remitoventa r on d.numero=r.numero where r.cancelado=0 and r.anulado=0 and d.pedido=" & s2n(txtNro))
    If IsNull(sp) Or IsEmpty(sp) Then
        Set sp = Nothing
        sp = obtenerDeSQL("select * from facturaventadetalle d inner join facturaventa f on d.nrofactura=f.nrofactura where d.nropedido=" & s2n(txtNro))
        If IsNull(sp) Or IsEmpty(sp) Then
        Else
            GoTo o
        End If
    Else
o:
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If
    
    ABMPedidoCliente "B", s2n(txtNro), uCliente.codigo, ObtenerCodigo("Formaspago", Trim(cmbformapago.Text)), ObtenerCodigo("usuarios", Trim(cmbvendedor.Text)), Trim(txtnropedidocli), ConCodigoPropio(), ObtenerCodigo("transportes", cmbTransporte), txtobs, 0, Date, UsuarioSistema!codigo, 0, 0, 0
    grabaBitacora "B", s2n(txtNro), "Pedidos_Clientes2"
    ucMenu.EliminarOK
fin:
    Exit Sub

ufaErr:
    ufa "error al eliminar", Me.Name & " " & txtNro ', Err
    Resume fin
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt sino
End Sub
Private Sub ucMenu_Nuevo()
    txtNro = Format(nuevoCodigo("Pedidos_Clientes2", "Numero"), "00000")
    chkPropio.Value = vbChecked
    chkPropio.enabled = True
    uCliente.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Private Sub ucMenu_SeMovio()
    On Error Resume Next
    leerPedido ucMenu.rs!numero
End Sub

Public Function ABMPedidoCliente(pOPE As String, pNRO As Long, pCLIENTE As Long, pPAGO As Long, pVENDEDOR As Long, pPEDCLI As String, pPROPIO As Long, pTRANSPORTE As Long, pObs As String, pCANC As Long, pFecha As Date, pUsu As Long, Emisor As Long, contacto As Long, moneda As Long) As Boolean
On Error GoTo pcmal
Dim iud As String
ABMPedidoCliente = True
Select Case pOPE
    Case "A"
        iud = "INSERT INTO PEDIDOS_CLIENTES2 ( NUMERO,CLIENTE,FECHA,PAGO,VENDEDOR,PEDIDO_CLI,CODIGOPROPIO,TRANSPORTE,OBSERVACIONES,CANCELADO, FECHA_ALTA, USUARIO_ALTA,  ACTIVO,emisor,contacto,moneda) " _
            & " VALUES ( " & pNRO & "," & pCLIENTE & "," & ssFecha(pFecha) & "," & pPAGO & "," & pVENDEDOR & "," & ssTexto(pPEDCLI) & "," & pPROPIO & "," & pTRANSPORTE & "," & ssTexto(pObs) & "," & pCANC & "," & ssFecha(Date) & "," & pUsu & ", 1," & Emisor & "," & contacto & "," & moneda & ")"
        DataEnvironment1.Sistema.Execute iud
    Case "M"
        iud = " UPDATE PEDIDOS_CLIENTES2 " _
        & " SET CLIENTE=" & pCLIENTE & ",PAGO=" & pPAGO & ",VENDEDOR=" & pVENDEDOR & ",PEDIDO_CLI=" & ssTexto(pPEDCLI) & ",CODIGOPROPIO=" & pPROPIO & ",TRANSPORTE=" & pTRANSPORTE & ",OBSERVACIONES=" & ssTexto(pObs) & ",CANCELADO=" & pCANC & ",emisor=" & Emisor & ",contacto=" & contacto & " WHERE NUMERO=" & pNRO
        DataEnvironment1.Sistema.Execute iud
        

    Case "B":
        iud = "UPDATE PEDIDOS_CLIENTES2  SET ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA= " & pUsu & " WHERE NUMERO=" & pNRO
        DataEnvironment1.Sistema.Execute iud
End Select
Exit Function
pcmal:
ABMPedidoCliente = False
End Function

