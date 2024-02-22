VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmFacturaVenta 
   Caption         =   "Factura Venta"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   14865
   Icon            =   "frmFacturaVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14865
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabDetalle 
      Height          =   4725
      Left            =   15
      TabIndex        =   25
      Top             =   2955
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8334
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmFacturaVenta.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSubtotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label18(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblIIBB"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDescuento"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtIva"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtNeto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label19"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label18(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label28"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label29"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "grilla"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fraOptStock"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPIVA"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPdescuento"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "fraEditDetalle"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSeguro"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtFlete"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdBorrarItem"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Origen"
      TabPicture(1)   =   "frmFacturaVenta.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grillaOrigen"
      Tab(1).Control(1)=   "cmdOrigen"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Series"
      TabPicture(2)   =   "frmFacturaVenta.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkSinSeries"
      Tab(2).Control(1)=   "cmdLlenaSerie"
      Tab(2).Control(2)=   "grillaSeries"
      Tab(2).Control(3)=   "Label23"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton cmdBorrarItem 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   11100
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmFacturaVenta.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   116
         TabStop         =   0   'False
         ToolTipText     =   "Borrar Item"
         Top             =   1245
         Width           =   645
      End
      Begin VB.TextBox txtFlete 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12900
         TabIndex        =   101
         Text            =   "0"
         Top             =   1635
         Width           =   1740
      End
      Begin VB.TextBox txtSeguro 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12900
         TabIndex        =   100
         Text            =   "0"
         Top             =   2010
         Width           =   1740
      End
      Begin VB.CheckBox chkSinSeries 
         Caption         =   "Sin Series  (Series en Remito)"
         Height          =   315
         Left            =   -68400
         TabIndex        =   88
         Top             =   780
         Width           =   4215
      End
      Begin VB.CommandButton cmdLlenaSerie 
         Caption         =   "Llenar Serie"
         Height          =   315
         Left            =   -74400
         TabIndex        =   87
         ToolTipText     =   "Seleccione filas a llenar"
         Top             =   780
         Width           =   1575
      End
      Begin VB.CommandButton cmdOrigen 
         Caption         =   "Traer Items Pendientes"
         Height          =   315
         Left            =   -74640
         TabIndex        =   73
         Top             =   480
         Width           =   1875
      End
      Begin VB.Frame fraEditDetalle 
         Height          =   870
         Left            =   105
         TabIndex        =   69
         Top             =   330
         Width           =   12735
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   10560
            TabIndex        =   20
            Top             =   165
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarItem 
            BackColor       =   &H00E0E0E0&
            Height          =   660
            Left            =   11880
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmFacturaVenta.frx":11E8
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Ingresar item"
            Top             =   150
            Width           =   660
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   10560
            TabIndex        =   21
            Top             =   510
            Width           =   1095
         End
         Begin VB.TextBox txtIvaProducto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   9315
            TabIndex        =   22
            Text            =   "21"
            Top             =   525
            Width           =   495
         End
         Begin Gestion.ucCoDe uProd 
            Height          =   315
            Left            =   60
            TabIndex        =   19
            Top             =   180
            Width           =   9840
            _extentx        =   14949
            _extenty        =   556
            codigowidth     =   1000
         End
         Begin VB.Label Label10 
            Caption         =   "Cant. :"
            Height          =   255
            Left            =   10005
            TabIndex        =   72
            Top             =   225
            Width           =   555
         End
         Begin VB.Label Label11 
            Caption         =   "Precio:"
            Height          =   240
            Left            =   10005
            TabIndex        =   71
            Top             =   555
            Width           =   555
         End
         Begin VB.Label lblIvaProducto 
            Caption         =   "IVA Producto:"
            Height          =   225
            Left            =   8265
            TabIndex        =   70
            Top             =   555
            Width           =   1035
         End
      End
      Begin VB.TextBox txtPdescuento 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12210
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2745
         Width           =   675
      End
      Begin VB.TextBox txtPIVA 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12210
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3900
         Width           =   675
      End
      Begin VB.Frame fraOptStock 
         Caption         =   " Actualiza Stock "
         Height          =   870
         Left            =   12900
         TabIndex        =   57
         Top             =   330
         Width           =   1755
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   2
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1005
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   1
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   525
            Width           =   315
         End
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   0
            Left            =   990
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   225
            Value           =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox TxtRemitoNumero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   180
            TabIndex        =   58
            Top             =   1140
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Genera Remito"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   1005
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Si"
            Height          =   195
            Index           =   1
            Left            =   645
            TabIndex        =   63
            Top             =   540
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "No"
            Height          =   195
            Index           =   3
            Left            =   645
            TabIndex        =   62
            Top             =   240
            Width           =   675
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grillaOrigen 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   66
         Top             =   960
         Width           =   9135
         _cx             =   16113
         _cy             =   5927
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
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   3375
         Left            =   120
         TabIndex        =   67
         Top             =   1260
         Width           =   10965
         _cx             =   19341
         _cy             =   5953
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
      Begin VSFlex7LCtl.VSFlexGrid grillaSeries 
         Height          =   3120
         Left            =   -74640
         TabIndex        =   89
         Top             =   1260
         Width           =   8595
         _cx             =   15161
         _cy             =   5503
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
         FormatString    =   $"frmFacturaVenta.frx":1AB2
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
      Begin VB.Label Label29 
         Caption         =   "Flete:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11850
         TabIndex        =   103
         Top             =   1650
         Width           =   585
      End
      Begin VB.Label Label28 
         Caption         =   "Seguro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11850
         TabIndex        =   102
         Top             =   2055
         Width           =   720
      End
      Begin VB.Label Label23 
         Caption         =   "Puede hacer 'Doble Clic' en el campo  Nro.Serie"
         Height          =   495
         Left            =   -72540
         TabIndex        =   90
         Top             =   600
         Width           =   2235
      End
      Begin VB.Label Label16 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11865
         TabIndex        =   85
         Top             =   4350
         Width           =   660
      End
      Begin VB.Label Label17 
         Caption         =   "Iva:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11865
         TabIndex        =   84
         Top             =   3945
         Width           =   330
      End
      Begin VB.Label Label18 
         Caption         =   "Neto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   11865
         TabIndex        =   83
         Top             =   3195
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Descuento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11850
         TabIndex        =   82
         Top             =   2430
         Width           =   1020
      End
      Begin VB.Label txtNeto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12915
         TabIndex        =   81
         Top             =   3135
         Width           =   1740
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12915
         TabIndex        =   80
         Top             =   4275
         Width           =   1740
      End
      Begin VB.Label txtIva 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12915
         TabIndex        =   79
         Top             =   3900
         Width           =   1740
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12915
         TabIndex        =   78
         Top             =   2745
         Width           =   1740
      End
      Begin VB.Label lblIIBB 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12915
         TabIndex        =   77
         Top             =   3525
         Width           =   1740
      End
      Begin VB.Label Label18 
         Caption         =   "IIBB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   11865
         TabIndex        =   76
         Top             =   3600
         Width           =   765
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12900
         TabIndex        =   75
         Top             =   1245
         Width           =   1740
      End
      Begin VB.Label Label22 
         Caption         =   "Sub Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11850
         TabIndex        =   74
         Top             =   1305
         Width           =   960
      End
   End
   Begin VB.Frame fraCabecera 
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   45
      TabIndex        =   26
      Top             =   0
      Width           =   14745
      Begin VB.TextBox txtCodReferencia 
         Height          =   315
         Left            =   4440
         TabIndex        =   115
         Text            =   "0"
         Top             =   450
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CheckBox chkResaltar 
         Alignment       =   1  'Right Justify
         Caption         =   "Resaltar Alquileres"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   7380
         TabIndex        =   114
         Top             =   2445
         Width           =   1725
      End
      Begin VB.TextBox txtCotiLeyen 
         Height          =   285
         Left            =   13680
         TabIndex        =   112
         Top             =   2625
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox ChkVaLeye 
         Alignment       =   1  'Right Justify
         Caption         =   "Leyenda Fija"
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   11775
         TabIndex        =   110
         Top             =   2850
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtNroCliente 
         Height          =   320
         Left            =   12840
         MaxLength       =   13
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1395
      End
      Begin VB.ComboBox cmbPunto 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   2205
         Width           =   2235
      End
      Begin VB.ComboBox cboIncoterms 
         Height          =   315
         Left            =   4650
         TabIndex        =   104
         Text            =   "cboIncoterms"
         Top             =   2220
         Width           =   1860
      End
      Begin VB.TextBox txtPermisoEmbarque 
         Height          =   300
         Left            =   1380
         TabIndex        =   97
         Top             =   2580
         Visible         =   0   'False
         Width           =   5130
      End
      Begin VB.TextBox Orden 
         Height          =   320
         Left            =   12840
         MaxLength       =   13
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1155
         Width           =   1395
      End
      Begin VB.TextBox Remito 
         Height          =   320
         Left            =   12840
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox txtCotizacion 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   12840
         TabIndex        =   46
         Top             =   1905
         Width           =   1395
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   12840
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   1530
         Width           =   1395
      End
      Begin VB.CommandButton cmdPedidosPendientes 
         Caption         =   "Pedido Pend"
         Height          =   315
         Left            =   7395
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1905
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.OptionButton optCuentaContado 
         Caption         =   "Contado"
         Height          =   300
         Index           =   1
         Left            =   2325
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   90
         Width           =   930
      End
      Begin VB.OptionButton optCuentaContado 
         Caption         =   "Cta Cte"
         Height          =   300
         Index           =   0
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   90
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.CommandButton cmdRemitosPendientes 
         Caption         =   "Remito Pend"
         Height          =   315
         Left            =   8715
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1905
         Width           =   1305
      End
      Begin VB.CommandButton cmdCliente 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   810
         Width           =   375
      End
      Begin VB.TextBox txtNroFacturaRef 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         TabIndex        =   8
         Top             =   450
         Width           =   1230
      End
      Begin VB.TextBox txtTipoDocRef 
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   450
         Width           =   495
      End
      Begin VB.ComboBox cmbDeposito 
         Height          =   315
         ItemData        =   "frmFacturaVenta.frx":1B89
         Left            =   12840
         List            =   "frmFacturaVenta.frx":1B8B
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   405
         Width           =   1395
      End
      Begin VB.CheckBox chkPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Propio"
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   7380
         TabIndex        =   38
         Top             =   2205
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.ComboBox cmbVendedor 
         Height          =   315
         Left            =   7410
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2610
      End
      Begin VB.ComboBox cmbTipoIva 
         Height          =   315
         Left            =   7410
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   810
         Width           =   2610
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         ItemData        =   "frmFacturaVenta.frx":1B8D
         Left            =   7410
         List            =   "frmFacturaVenta.frx":1B8F
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   450
         Width           =   2610
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   465
         Width           =   495
      End
      Begin VB.ComboBox cmbProvincia 
         Height          =   315
         Left            =   4635
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1860
      End
      Begin VB.TextBox txtDireccion 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5115
      End
      Begin VB.TextBox txtLocalidad 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2235
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   10
         Top             =   810
         Width           =   975
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   2775
         TabIndex        =   12
         Top             =   810
         Width           =   3735
      End
      Begin VB.TextBox TxtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   465
         Width           =   1305
      End
      Begin Gestion.ucCuit ucCuit 
         Height          =   315
         Left            =   7410
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1155
         Width           =   2610
         _extentx        =   2355
         _extenty        =   556
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   7410
         TabIndex        =   2
         Top             =   105
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   179503105
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtVencimiento 
         Height          =   315
         Left            =   10005
         TabIndex        =   3
         Top             =   105
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   179503105
         CurrentDate     =   38229
      End
      Begin VB.Label Label20 
         Caption         =   "Cuit:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6660
         TabIndex        =   118
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label33 
         Caption         =   "Cotizacion Leyenda :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11805
         TabIndex        =   111
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label32 
         Caption         =   "Nro Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11805
         TabIndex        =   109
         Top             =   2280
         Width           =   1755
      End
      Begin VB.Label Label31 
         Caption         =   "P de Venta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   120
         TabIndex        =   107
         Top             =   2220
         Width           =   1080
      End
      Begin VB.Label Label30 
         Caption         =   "Incoterms:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   3660
         TabIndex        =   105
         Top             =   2235
         Width           =   945
      End
      Begin VB.Label Label27 
         Caption         =   "Embarque:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   120
         TabIndex        =   99
         Top             =   2580
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblCAE2 
         Height          =   225
         Left            =   11475
         TabIndex        =   98
         Top             =   3135
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADJUNTAR FACTURAS:"
         Height          =   270
         Left            =   1380
         TabIndex        =   96
         Top             =   1905
         Visible         =   0   'False
         Width           =   5130
      End
      Begin VB.Label Label25 
         Caption         =   "Orden:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11805
         TabIndex        =   94
         Top             =   1185
         Width           =   1755
      End
      Begin VB.Label Label24 
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11820
         TabIndex        =   92
         Top             =   825
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Cotizacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11790
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11805
         TabIndex        =   47
         Top             =   1545
         Width           =   900
      End
      Begin VB.Label lblExterior 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXTERIOR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3255
         TabIndex        =   44
         Top             =   90
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label txtCodigo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12840
         TabIndex        =   4
         Top             =   60
         Width           =   1380
      End
      Begin VB.Label lblRef 
         Caption         =   "Asociada:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3450
         TabIndex        =   41
         Top             =   465
         Width           =   1335
      End
      Begin VB.Label lblDepot 
         Caption         =   "Deposito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11820
         TabIndex        =   40
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label15 
         Caption         =   "Vendor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6675
         TabIndex        =   37
         Top             =   1575
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6675
         TabIndex        =   36
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Vencimiento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   8790
         TabIndex        =   35
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   11820
         TabIndex        =   34
         Top             =   90
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "Iva:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6675
         TabIndex        =   33
         Top             =   870
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "Provincia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3675
         TabIndex        =   32
         Top             =   1590
         Width           =   915
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   105
         TabIndex        =   31
         Top             =   1575
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Direccion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   105
         TabIndex        =   30
         Top             =   1200
         Width           =   915
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6660
         TabIndex        =   29
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   90
         TabIndex        =   28
         Top             =   495
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   840
         Width           =   915
      End
   End
   Begin Gestion.ucBotonera ucBoton 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      Height          =   1635
      Left            =   0
      TabIndex        =   24
      Top             =   7800
      Width           =   14865
      _extentx        =   21061
      _extenty        =   2884
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      Begin VB.TextBox lblCAE 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7950
         TabIndex        =   120
         Top             =   1095
         Width           =   3750
      End
      Begin VB.CommandButton cmdAsignarCAE 
         Caption         =   "Asignar CAE"
         Height          =   420
         Left            =   7965
         TabIndex        =   119
         Top             =   645
         Width           =   3735
      End
      Begin VB.CommandButton cmdContado 
         Caption         =   "Ingreso Contado"
         Height          =   525
         Left            =   7965
         TabIndex        =   117
         Top             =   90
         Width           =   1815
      End
      Begin VB.CommandButton cmdRelacionar 
         Caption         =   "Relacionar Factura"
         Height          =   765
         Left            =   11895
         Picture         =   "frmFacturaVenta.frx":1B91
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   645
         Width           =   2895
      End
      Begin VB.CommandButton cmbingresar 
         Caption         =   "Centro de Costos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9795
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   1875
      End
      Begin VB.Frame fraBuscar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   195
         TabIndex        =   49
         Top             =   120
         Width           =   6255
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact B"
            Height          =   255
            Index           =   1
            Left            =   1425
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   30
            Width           =   705
         End
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact A"
            Height          =   255
            Index           =   0
            Left            =   645
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   30
            Width           =   735
         End
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact E"
            Height          =   255
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   30
            Width           =   705
         End
         Begin Gestion.ucEntreFechas ucFechas 
            Height          =   360
            Left            =   3480
            TabIndex        =   51
            Top             =   -15
            Width           =   2655
            _extentx        =   4683
            _extenty        =   635
         End
         Begin VB.Label Label12 
            Caption         =   "Buscar:"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   55
            Top             =   45
            Width           =   555
         End
         Begin VB.Label Label12 
            Caption         =   "Entre:"
            Height          =   195
            Index           =   0
            Left            =   3030
            TabIndex        =   54
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.Label lblFacturaB 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura B"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   525
         Left            =   11895
         TabIndex        =   56
         Top             =   90
         Visible         =   0   'False
         Width           =   2865
      End
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   95
      Text            =   "frmFacturaVenta.frx":245B
      Top             =   6540
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmFacturaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BS_PED_A_FAC_PREGUNTA_CANT = False
Private Const BS_RESPETAR_PRECIO_OC = True
Private midDoc As Long
Private ctasProducto As String

Public nroaux As Double

Private mClienteConIva As Boolean
Private mPropio As String
Private mValoresOK As Boolean
Private mNuevo As Boolean
'Private mContado As Boolean

'tipo factura, llamo con .mostrar() desde menu
Private mFac As fv_FacVentaSobre
Private mFAE As Boolean
'
Public Enum fv_FacVentaSobre
    FacturaVenta_Remito
    FacturaVenta_Pedido
    FacturaVenta_Libre
    FacturaVenta_NCreditoDevolucion
    FacturaVenta_Exterior
    FacturaVenta_NCredito
    FacturaVenta_NDebito
    FacturaVenta_NDebitoPorCheque
    FacturaVenta_NDebitoPorChequeE
    FacturaVenta_ticket
End Enum

Private Enum ActuStock
    ActuStock_NO = 0
    ActuStock_SI = 1
    ActuStock_RE = 2
End Enum

'Tabla facturas ventas , pero para busq, porq uso el SP
Private Const mTablaFV = "facturaventa"
Private Const StringCONTADO = "CONTADO"

Private NPregunta As Boolean
Private NRespuesta As Boolean


Private WithEvents g As LiGrilla, WithEvents gO As LiGrilla, gS As LiGrilla
Attribute g.VB_VarHelpID = -1
Attribute gO.VB_VarHelpID = -1
Private WithEvents cliente As LiCodigo ' cliente
Attribute cliente.VB_VarHelpID = -1
'Private WithEvents p As LiCodigo ' producto

'Detalle
Private gCANT   As Long
'Private gALIA   As Long
Private gprod   As Long
Private gDESC   As Long
Private gPUNI   As Long ' precio uni
Private gPTOT   As Long ' precio tot
Private gNPED   As Long ' pedido
Private gNREM   As Long ' remito
Private gFORM   As Long ' formula
Private gITEM   As Long ' item pedido o remito detalle
Private gIVA    As Long
Private gCTA    As Long
Private gCTAd    As Long
Private gNPCL   As Long ' Nro Pedido clie

'Origen Pedido-Remito
Private gO_PROD As Long
Private gO_DESC As Long
Private gO_CANT As Long
Private gO_NPED As Long ' nro pedido
Private gO_NREM As Long ' nro remito
Private gO_PREC As Long ' precio
Private gO_PEND As Long
Private gO_PROP As Long
Private gO_ITEM As Long ' item remito o pedido detalle
Private g0_NPCL As Long ' Nro Pedido clie
Private gO_IVA  As Long
'Private gO_ITRE As Long ' item remito detalle
'Private gO_ITPE As Long ' item pedido detalle


'Series
Private gS_ITEM As Long
Private gS_PROD As Long
Private gS_NSER As Long
Private gS_CONS As Long
Private gS_HIDD As Long
'

Public Sub mostrar(FacturarSobre As fv_FacVentaSobre, Optional FacturaExterior As Boolean = False, Optional Ver As Boolean = True)
    mFac = FacturarSobre
    mFAE = FacturaExterior
    lblExterior.Visible = mFAE
    
    If mFAE Then
        Label27.Visible = True
        txtPermisoEmbarque.Visible = True
        cmdAsignarCAE.Visible = False
        lblCAE.Visible = False
        txtFlete.Visible = True
        txtSeguro.Visible = True
    Else
        txtPermisoEmbarque.Visible = False
        cmdAsignarCAE.Visible = False
        lblCAE.Visible = False
        txtFlete.Visible = False
        txtSeguro.Visible = False
    End If
    
    Select Case mFac
    Case FacturaVenta_Libre, FacturaVenta_NDebito
        If mFac = FacturaVenta_NDebito Then
            verCampos False, True, True, 1500, 3000, True, "Nota de Debito", False, ActuStock_NO, True, False, False
        Else
            verCampos False, True, True, 1500, 3000, True, "Factura Venta", True, IIf(gEMPR_EmiteFacturaConRemito, ActuStock_RE, ActuStock_SI), False, False, False
        End If
        If gEMPR_EmiteFacturaConRemito Then
            optStock(ActuStock_RE).Value = True ' hace remito
        Else
            optStock(ActuStock_SI).Value = True ' mod stock
        End If
        optStock(ActuStock_NO).Value = True
        fraOptStock.Visible = True 'False
    
    Case FacturaVenta_Pedido
        verCampos True, False, True, 1500, 3000, True, "Factura Venta - SOBRE PEDIDO", True, IIf(gEMPR_EmiteFacturaConRemito, ActuStock_RE, ActuStock_SI), False, False, True
        'optStock(ActuStock_RE).Value = True '
        
'        If gEMPR_EmiteFacturaConRemito Then
'            optStock(ActuStock_RE).Value = True ' hace remito
'        Else
            optStock(ActuStock_NO).Value = True ' mod stock
'        End If
    
    Case FacturaVenta_Remito
        'verCampos True, False, False, 480, 4000, False, "Factura Venta - SOBRE REMITO", False, ActuStock_NO, False, True, False
        verCampos True, True, True, 1500, 3000, False, "Factura Venta - SOBRE REMITO", False, ActuStock_NO, False, True, False

        optStock(ActuStock_NO).Value = True ' Sin modif stock
        fraOptStock.Visible = True 'False
        TabDetalle.TabVisible(2) = False
        
    Case FacturaVenta_NCreditoDevolucion, FacturaVenta_NCredito
        If mFac = FacturaVenta_NCreditoDevolucion Then
            verCampos False, True, True, 1500, 3000, True, "Nota de Credito POR DEVOLUCION", True, ActuStock_SI, True, False, False
        ElseIf mFac = FacturaVenta_NCredito Then
            verCampos False, True, True, 1500, 3000, True, "Nota de Credito", False, ActuStock_NO, False, False, False
        End If
        optCuentaContado.item(0).Value = True
        cmdContado.enabled = False
        optCuentaContado.item(0).enabled = False
        optCuentaContado.item(1).enabled = False
        cmdContado.Visible = False
        
'        optStock(ActuStock_SI).Value = True
        optStock(ActuStock_NO).Value = True
        
        optStock(ActuStock_RE).Visible = False
        TxtRemitoNumero.Visible = False
        Label1(ActuStock_RE).Visible = False
    
    Case Else                 'assert
        ufa "Prg", ".mostrar() " & Me.Name ', Err
    End Select
    If Ver = True Then
        Me.Show
    End If
End Sub

Private Function DeboModificarStock() As Boolean
    'DeboModificarStock = (chkModStock.Value = vbChecked)
    DeboModificarStock = (optStock(ActuStock_SI).Value Or optStock(ActuStock_RE).Value)
End Function

Private Sub chkPropio_LostFocus()
    set_uProd
End Sub

'Private Sub chkModStock_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    TabDetalle.TabVisible(2) = DeboModificarStock()
'End Sub

Private Sub cliente_cambio(codigo) ' As Integer)
    If ON_ERROR_HABILITADO Then On Error GoTo epa_ERR
    
    Dim ac As Variant, a As String, n As Long
    If codigo = 0 Then
        FrmBorrarTxt Me
    Else
        ac = obtenerDeSQL("select codigo, direccion, localidad, provincia, cuit, iva, FormaPago, Vendedor, Descuento1,nrocliente from clientes where codigo = " & codigo)
        txtdireccion = sSinNull(ac(1))
        txtLocalidad = sSinNull(ac(2))
        a = sSinNull(ac(3))
        CmbProvincia = ObtenerDescripcionS("provincias", a)
        ucCuit.Text = sSinNull(ac(4))
 
        cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, ac(6))
        CalcularVencimiento
        
        cmbvendedor.ListIndex = BuscarEnCombo(cmbvendedor, ac(7))
        
        n = ac(5)
        cmbTipoIva.enabled = False
        cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, n)
        lblFacturaB.Visible = "B" = sSinNull(obtenerDeSQL("Select letra from ivas where codigo = " & n))
        
        'txtPIVA = s2n(obtenerDeSQL("select porcentaje from PorcentajesIva where activo = 1 and iva = " & ComboCodigo(cmbTipoIva)) * 100)
        mClienteConIva = Not mFAE And (0 < s2n(obtenerDeSQL("select porcentaje from PorcentajesIva where activo = 1 and iva = " & n)))
        txtPdescuento = ac(8) ' * 100)
        txtNroCliente = sSinNull(ac(9))
        
        If ucBoton.estado = ucbEditando Then
            TxtNroFactura = ""
            BuscoNroYTipo True  'False
        End If
        
        
        
        set_uProd
    End If
    
    g.Borrar
    gO.Borrar
    mPropio = ""
'   gS.Borrar
fin:
    Exit Sub
epa_ERR:
    ufa "", "error leyendo codigo traido por c_cambio Fac Venta " & codigo ', Err
    Resume fin
End Sub


Private Sub cmbFormaPago_LostFocus()
    CalcularVencimiento
End Sub
Private Sub cmbFormaPago_Validate(cancel As Boolean)
    CalcularVencimiento
End Sub

Private Sub CalcularVencimiento()
    On Error GoTo ufa
    If Not dtVencimiento.enabled Then Exit Sub
    
    
    Dim sql As String
    Dim rsFormaP As New ADODB.Recordset
    
    sql = "Select dias from FormasPago WHERE codigo =" & cmbformapago.ItemData(cmbformapago.ListIndex)
    rsFormaP.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    dtVencimiento.Value = dtFecha + rsFormaP!Dias
ufa:
End Sub

Private Sub cmbingresar_Click()
    If txttotal = "" Then
        MsgBox "Debe ingresar algun producto en la factura"
        Exit Sub
    End If
    If txtDescuento = "" Then
        txtDescuento = 0
    End If
    
'    FrmCostosYContable.txtimporte.Enabled = False
    FrmCostosYContable.cargar.enabled = False
'    FrmCostosYContable.txtcuentacod.Enabled = False
'    FrmCostosYContable.cmbcuenta.Enabled = False
'    FrmCostosYContable.txtcuenta.Enabled = False
    FrmCostosYContable.txtconc.enabled = False
    FrmCostosYContable.txtvalor.enabled = False
    FrmCostosYContable.cmdcargar.enabled = False
    FrmCostosYContable.cmbeliminofila.enabled = False
'    FrmCostosYContable.txtTotal.Enabled = False
'    FrmCostosYContable.Grilla.Enabled = False
        
    FrmCostosYContable.CargarImputacion s2n(txtneto) + s2n(txtIva) - s2n(txtDescuento), s2n(txttotal), 2
    'FrmCostosYContable.txtimptotal = txtimporte
    'FrmCostosYContable.txtimporte = txtNeto
    FrmCostosYContable.txtimporte = ""
    FrmCostosYContable.cargar = ""
    FrmCostosYContable.txtcuentacod = ""
    FrmCostosYContable.txtcuenta = ""
    FrmCostosYContable.txtconc = ""
    FrmCostosYContable.txtvalor = ""
    FrmCostosYContable.txttotal = ""
'    FrmCostosYContable.grilla.Clear
    
    FrmCostosYContable.Tag = Me.Name
    vieneDE = Me.Name
    FrmCostosYContable.Show
End Sub

Private Sub cmbPunto_Change()
cmbPunto_Click
End Sub

Private Sub cmbPunto_Click()
Dim andtipo As String, sPunto As String
Dim Tipo As String
    If ucBoton.estado <> ucbMostrando Then
        If txtTipoDoc = "" Then Exit Sub
        If ucBoton.estado = 0 Then Exit Sub
        andtipo = " tipodoc=" & ssTexto(txtTipoDoc)
        sPunto = PuntoVentaTipo(cmbPunto.ListIndex)
        Tipo = sPunto
        sPunto = sSinNull(obtenerDeSQL("select puntoventa from documentoscae where tipo=" & ssTexto(txtTipoDoc) & " and tipopunto=" & ssTexto(sPunto)))
        TxtNroFactura = s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) + 1
        BuscoNroYTipo True
        If Trim(Tipo) = "WS" Then
            TxtNroFactura.Locked = True
        Else
            TxtNroFactura.Locked = False
        End If
    End If
End Sub

Private Sub cmdAgregarItem_Click()
    Dim r As Long, pco, pde ', ivaprod As Double
    Dim PU As Double, PIVA As Double
    
'    pco = CodigoDeAlias(UProd.codigo)
    pco = uProd.codigo
    pde = Trim(uProd.DESCRIPCION)
    
    If nroaux = 0 Then
        nroaux = s2n(txtIvaProducto.Text, 4) / 100
    End If

'    txtDescProducto = Trim(txtDescProducto)
''    If s2n(txtCantidad) = 0 Or txtCodProducto = "" Or txtDescProducto = "" Then Exit Sub
    
'    If pco > "" And (s2n(txtCantidad, 4) = 0 Or s2n(txtPrecio, 4) = 0) Then
'        che "Falta precio/cantidad"
'        Exit Sub
'    End If
    If pco = "" And pde = "" Then
        'che "pone algo..."
        Exit Sub
    End If
    If s2n(txtprecio, 4) > 0 And s2n(txtCantidad, 4) = 0 Then
        che "hay precio pero falta especificar cantidad"
        Exit Sub
    End If
    
    Dim dTipo As String
    dTipo = UCase(txtTipoDoc)
    PIVA = 0
    If InStr(dTipo, "B") Then
        If s2n(txtIvaProducto, 4) > 0 Then
'            PU = s2n(1 + s2n(s2n(txtIvaProducto, 4) / 100), 4)
            PU = s2n(s2n(txtprecio, 4) * s2n(1 + s2n(s2n(txtIvaProducto, 4) / 100, 4), 4), 3)
            PIVA = s2n(txtIvaProducto, 4)
        Else
            PU = s2n(txtprecio, 4)
        End If
    Else
        PU = s2n(txtprecio, 4)
        PIVA = s2n(txtIvaProducto, 4)
    End If
    MetoEnGrilla pco, pde, s2n(txtCantidad, 4), PU, 0, 0, 0, s2n(PIVA), , False     ', s2n(txtConsignacion)
    
    txtCantidad = ""
'    txtCodProducto = ""
'    txtDescProducto = ""
    uProd.codigo = ""
    txtprecio = ""
    txtIvaProducto = ""
    'txtConsignacion = ""
    'txtCodProducto.SetFocus
    uProd.SetFocus
    
    chkPropioEnabled True  ' permite habilitar, ...
End Sub

Private Sub cmdAsignarCAE_Click()
Dim doc As New FacturaElectronica
    If siFCAE(Trim(txtTipoDoc), PuntoVentaTipo(cmbPunto.ListIndex)) Then
        If doc.EmisionFacturaElectronica(s2n(txtCodigo), txtPermisoEmbarque) Then
            If Right(Trim(txtTipoDoc), 1) = "E" Then
                ImprimirComprobanteFE s2n(txtCodigo)
            Else
                ImprimirComprobanteFE2 s2n(txtCodigo)
            End If
        Else
            If MsgBox("GENERAR FACTURA EN PDF", vbYesNo + vbInformation) = vbYes Then
                If Right(Trim(txtTipoDoc), 1) = "E" Then
                    ImprimirComprobanteFE s2n(txtCodigo)
                Else
                    ImprimirComprobanteFE2 s2n(txtCodigo)
                End If
            End If
        End If
    Else
        Dim eCUIT As String, eCODFACTURA As String, ePUNTOVENTA As String, eCAE As String, eBARRA As String, eFECHAVENCE As String
        Dim IDFactura As Long
        If cmbPunto.ListIndex = 1 Then
            If lblCAE > "" Then
                GoTo IMPRIMIR_FE
            Else
                lblCAE = sSinNull(obtenerDeSQL("SELECT CAE FROM FACTURAVENTA WHERE CODIGO=" & s2n(txtCodigo)))
                If lblCAE > "" Then
                    GoTo IMPRIMIR_FE
                End If
            End If
            
            If MsgBox("¿Desea cargar CAE manualmente?", vbYesNo + vbInformation) = vbYes Then
                IDFactura = txtCodigo
                eCUIT = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
                eCUIT = Format(eCUIT, "00000000000")
                ePUNTOVENTA = Format(Trim(obtenerDeSQL("select puntoventa from documentoscae where tipopunto='OL' and tipo=" & ssTexto(txtTipoDoc))), "0000")
                eCODFACTURA = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='OL' and tipo=" & ssTexto(txtTipoDoc) & " and puntoventa=" & ssTexto(ePUNTOVENTA))), "00")
                eCAE = InputBox("INGRESE NUMERO DE CAE (14 DIGITOS)...", "CAE OTORGADO ONLINE")
                While Len(eCAE) < 14
                    eCAE = InputBox("INGRESE NUMERO DE CAE (14 DIGITOS)...", "CAE OTORGADO ONLINE")
                Wend
                eFECHAVENCE = InputBox("INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                While Len(eFECHAVENCE) < 10
                    eFECHAVENCE = InputBox("INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                Wend
                While CDate(eFECHAVENCE) < dtFecha
                    eFECHAVENCE = InputBox("FECHA MENOR AL COMPROBANTE" & Chr(13) & "INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                Wend
                eFECHAVENCE = afipFecha(CDate(eFECHAVENCE))
                
                
                'bBARRA = bCuit & bCodFactura & bPuntoVenta & bCAE & bFechaCAE & "8" 'MOMENTAMEAMENTE VA 8 FIJO HASTA QUE HABERIGUE DE DONDE SALE
                eBARRA = eCUIT & eCODFACTURA & ePUNTOVENTA & eCAE & eFECHAVENCE
                eBARRA = eBARRA & CodVerificador(eBARRA)
                
    '            .FE.FERespuestaDetalleFecha_vto
                DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(eBARRA) & ",caev=" & ssFecha(aFecha(eFECHAVENCE)) & ", cae=" & ssTexto(eCAE) & " where codigo=" & IDFactura
IMPRIMIR_FE:
                If MsgBox("GENERAR FACTURA EN PDF", vbYesNo + vbInformation) = vbYes Then
                    If Right(Trim(txtTipoDoc), 1) = "E" Then
                        ImprimirComprobanteFE s2n(txtCodigo)
                    Else
                        ImprimirComprobanteFE2 s2n(txtCodigo)
                    End If
                End If
            End If
        End If
    End If
End Sub


'Private Sub cmdAyudaProducto_Click()
'    Dim r As Variant
'    r = AyudaProducto(cliente.Codigo, Propio())
'    If r = "" Then Exit Sub
'
'    txtCodProducto = frmBuscar.resultado(1)
'    txtDescProducto = frmBuscar.resultado(2)
'End Sub


Private Sub cmdBorrarItem_Click()
    If g.Row > 0 Then g.delRow (g.Row)
    RevisarTotales
End Sub

Private Sub cmdContado_Click()
    'cmbFormaPago.ListIndex = BuscarEnCombo(cmbFormaPago, 1)
'    mContado = True
    optCuentaContado.item(1).Value = True
    If FaltaAlgo() Then Exit Sub
    mValoresOK = frmValores.mostrar(s2n(txttotal))
End Sub


Private Sub cmdLlenaSerie_Click()
    Dim i As Long, Row As Long
    Dim s As String, n As Long       ' current cell
    Dim ss As String, ns As Long     ' the longer one

    ss = ""
    n = 0

    With grillaSeries ' uso propiedades que no exporte

        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, gS_NSER)

            If Len(s) > Len(ss) Then
                ss = s
                ns = Len(ss)
            End If
        Next i

        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, gS_NSER)
            n = Len(s)
            If n < ns And n > 0 Then .TextMatrix(Row, gS_NSER) = Left(ss, ns - n) & s
        Next i
    End With
End Sub

Private Sub cmdOrigen_Click()
    CargaOrigen
End Sub


Private Sub cmdPedidosPendientes_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAcmdPend
    
    Dim s As String, i As Long, rs As New ADODB.Recordset, resu As String
    Dim prod, desc, cant, prec, pedi, remi, item, prop, r As Long

    s = " Select distinct " _
      & " numero  as Numero, pedido_cli as NroPedidoCliente, cliente " _
      & " from Pedidos_Clientes inner join ItemPedidoCliente " _
      & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
      & " where facturar > 0 and Pedidos_Clientes.activo = 1"

    If cliente.codigo > 0 Then s = s & " and cliente = " & cliente.codigo
   
    resu = frmBuscar.MostrarSql(s)
    If resu = "" Then Exit Sub
    
    cliente.codigo = s2n(frmBuscar.resultado(3))
    g.Borrar
    TabDetalle.Tab = 0

     s = " Select " _
       & " ItemPedidoCliente.codigo as cod, numero as nPedi, producto, cantidad, precio, facturar, CodigoPropio as codPropio, 0 as nRemi" _
       & " from Pedidos_Clientes inner join ItemPedidoCliente " _
       & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
       & " where numero = " & resu
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        If rs!cantidad = 0 Or rs!facturar > 0 Then
            prod = VerProductoCliente(rs!producto, rs!codPropio, cliente.codigo)
            desc = ObtenerDescripcionS("producto", rs!producto)
'            iva = s2n(obtenerDeSQL("select iva from producto where codigo = '" & ssStr(rs!codigo, True) & "' ") * 100)
            cant = rs!facturar

            
            'If BS_RESPETAR_PRECIO_OC Then
            prec = rs!precio
            'End If
            
            pedi = rs!nPedi
            item = rs!COD
            'prop = rs!codPropio
            chkPropio.Value = IIf(rs!codPropio, vbChecked, vbUnchecked)
            set_uProd
            MetoEnGrilla prod, desc, cant, prec, pedi, 0, item, 0, , False
            
            chkPropio.Value = IIf(rs!codPropio, vbChecked, vbUnchecked)
        End If
        rs.MoveNext
    Wend
fin:
    Set rs = Nothing
    relojito False
    Exit Sub
UFAcmdPend:

    ufa "Err leyendo pendientes", "cmdPeridoPend"
    Resume fin
End Sub

Private Sub cmdRelacionar_Click()
Dim EnQueEj As Long, QueIddoc As Long, re
Dim QueTiene As Long
    If s2n(txtCodigo) Then
        EnQueEj = nSinNull(obtenerDeSQL("select ejercicio from ejercicio where fechafin>=" & ssFecha(dtFecha) & " and fechainicio<=" & ssFecha(dtFecha)))
        re = frmBuscar.MostrarSql("select IDASIENTO,NROASIENTO,CONCEPTO,FECHA from  asientos where activo=1 and ejercicio=" & EnQueEj, , , , "", "Anulada", False)
        If re = "" Then Exit Sub
        QueIddoc = nSinNull(obtenerDeSQL("select iddoc from facturaventa where codigo=" & txtCodigo))
        QueTiene = nSinNull(obtenerDeSQL("select iddoc from asientos where idasiento=" & s2n(re)))
        If QueTiene > 0 Then
            If MsgBox("El asiento ya tiene una ralacion. Desea reemplazarla?", vbYesNo + vbInformation) = vbYes Then
                DataEnvironment1.Sistema.Execute "update asientos set iddoc=" & QueIddoc & " where idasiento=" & s2n(re)
                MsgBox "Reemplazado...", vbInformation
            End If
        Else
            DataEnvironment1.Sistema.Execute "update asientos set iddoc=" & QueIddoc & " where idasiento=" & s2n(re)
            MsgBox "Guardado...", vbInformation
        End If
    End If
End Sub

Private Sub cmdRemitosPendientes_Click()
    Dim s ', re

    ' formula = '' : se factura el item separado
    ' formula = 'V' : se factura solo el item virtual, (formula, compuesta por items reales )
    ' formula <> '' : componentes de una formula, son reales, pero se factura solo la formula
    If cliente.codigo = 0 Then
      
        s = " Select distinct " _
         & " RemitoVenta.numero as Remito, cliente as [ Cliente ], clientes.descripcion as [ Nombre                                  ]  " _
         & " from RemitoVenta inner join RemitoVentaDetalle " _
         & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
         & " inner join clientes " _
         & " on cliente = clientes.codigo" _
         & " where facturar > 0 " _
         & " and RemitoVenta.Anulado = 0 " _
         & " and RemitoVenta.Cancelado = 0 " _
         & " and RemitoVentaDetalle.Cancelado = 0 "
        If gEMPR_FormulaEsVirtual Then s = s & " and (formula = '' or formula = 'V')"
        
    Else
      s = " Select distinct " _
        & " RemitoVenta.numero as Remito, cliente as [ Cliente ], clientes.descripcion as [ Nombre                                    ]  " _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " inner join clientes " _
        & " on cliente = clientes.codigo" _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
        If gEMPR_FormulaEsVirtual Then s = s & " and (formula = '' or formula = 'V')"
    End If

    With frmBuscar
        If frmBuscar.MostrarSql(s) > "" Then
            If cliente.codigo = 0 Then cliente.codigo = s2n(.resultado(2))
            CargaDiscriminada s2n(.resultado(2)), s2n(.resultado(1))
        End If
    End With
    
    TabDetalle.Tab = 0
End Sub

Private Sub CargaDiscriminada(clie As Long, remi As Long)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaDISCRI
    Dim rs As New ADODB.Recordset, s As String, i As Long
    
    
'    s = " Select " _
'        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi, producto, cantidad, precio, facturar, codPropio, 0 as nPedi" _
'        & " from RemitoVenta inner join RemitoVentaDetalle " _
'        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
'        & " and RemitoVenta.Anulado = 0 " _
'        & " and RemitoVenta.Cancelado = 0 " _
'        & " and RemitoVentaDetalle.Cancelado = 0 " _
'        & " and cliente = " & clie
        
    s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi, producto, cantidad, precio, facturar, codPropio, 0 as nPedi" _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & clie
        
        
        
'    If gEMPR_FormulaEsVirtual Then s = s & " and (formula = '' or formula = 'V')"
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    Dim dTipo As String
    dTipo = UCase(txtTipoDoc)
    With rs
        If .EOF Then
            ufa "err al cargar remitos", "CargaDiscriminada no trajo items " & clie & " " & remi & Me.Name ', 0
        Else
            If InStr(dTipo, "B") Then
                Dim Valor As Double
                Valor = InputBox("Ingrese el iva correspondiente (ej. 10,5) :" & Chr(13) & "En caso de utilizar decimales utilice coma y no punto.", "ATENCION")
                If s2n(Valor) = 0 Then Valor = 21
            End If
            While Not .EOF
                ' si es = remi q eligio en el frmBuscar , lo meto en la grilla 1
                ' si no, lo meto en la segunda por las dudas
                                
                chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)
                Valor = s2n(obtenerDeSQL("select iva from producto where codigo=" & ssTexto(!producto)) * 100, 4)
                                
                If !nRemi = remi Then
                    Dim PU2 As Double
                                        
                    If InStr(dTipo, "B") Then
                        PU2 = s2n(!precio * s2n((Valor / 100) + 1, 4), 4)
                    Else
                        PU2 = !precio
                    End If
                    MetoEnGrilla VerProductoCliente(!producto, !codPropio, CLng(clie)), ObtenerDescripcionS("producto", !producto), !facturar, PU2, !nPedi, !nRemi, !COD, Valor, , False
                Else
                    i = gO.addRow
                    gO.tx i, gO_ITEM, !COD
                    gO.tx i, gO_NPED, !nPedi
                    gO.tx i, gO_NREM, !nRemi
                    gO.tx i, gO_PROD, VerProductoCliente(!producto, !codPropio, CLng(clie))
                    gO.tx i, gO_DESC, ObtenerDescripcionS("producto", !producto)
                    gO.tx i, gO_CANT, !cantidad
                    gO.tx i, gO_PEND, !facturar
                    gO.tx i, gO_PREC, !precio
                    gO.tx i, gO_PROP, !codPropio
                End If
                .MoveNext
            Wend
        End If
    End With
fin:
    Set rs = Nothing
    Exit Sub
ufaDISCRI:
    ufa "err cargando remito", " CargaDiscri  " & clie & " " & remi & Me.Name ', Err
    Resume fin
End Sub


Private Sub Form_Activate()
'    Dim fact As Integer
    
    
'    If ucBoton.Estado = ucbEditando And mNuevo Then
''        fact = s2n(TxtNroFactura)
'        BuscoNroYTipo False
''        If fact <> 0 And fact <> s2n(TxtNroFactura) Then
''            che "Se Cambio Numeracion"
''        End If
'    End If
    
    SubimeSi800x600
End Sub

Private Sub Form_Load()
    Dim rsEjercicio As New ADODB.Recordset
    Dim rsFormaP As New ADODB.Recordset
    Dim sql As String
    Dim ctmp
    Dim punto As String
    Dim sTipo As String
    
    ctmp = obtenerDeSQL("select (ctasmayorista + ',' + ctasminorista) as ctas from datosempresa")
    ctmp = Replace(ctmp, "#", "")
    ctasProducto = Replace(ctmp, ",", "|")
    
'    uTipoVenta.UsoCuenta = 3  'Fac Venta
'    TabDetalle.TabVisible(2) = gEMPR_Maneja_series
    
    comboSql CmbProvincia, "select descripcion from provincias where activo = 1"
'   comboSql cmbFormaPago, "select descripcion,codigo,dias from formasPago where activo = 1 order by descripcion"
    
    sql = "select descripcion,codigo,dias from formasPago where activo = 1 order by dias, codigo"
    rsFormaP.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    Do While Not rsFormaP.EOF
      cmbformapago.AddItem rsFormaP!DESCRIPCION
      cmbformapago.ItemData(cmbformapago.NewIndex) = rsFormaP!codigo
      rsFormaP.MoveNext
    Loop
    rsFormaP.Close
    Set rsFormaP = Nothing
    
    
    sql = "select * from doc_incoterms where activo = 1 "
    With rsFormaP
        .Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            Do While Not .EOF
              cboIncoterms.AddItem !DESCRIPCION
              cboIncoterms.ItemData(cboIncoterms.NewIndex) = !COD
              .MoveNext
            Loop
    End With
    Set rsFormaP = Nothing
    cboIncoterms.ListIndex = 0
    
    cmbPunto.AddItem "PRE-IMPRESA"
    cmbPunto.AddItem "ONLINE"
    cmbPunto.AddItem "WEBSERVICE"
    cmbPunto.AddItem "WEBSERVICE 2"
    
        
    punto = sSinNull(obtenerDeSQL("select puntodefecto from puntodefecto where usuario=" & UsuarioActual()))
    If punto = "" Then
        cmbPunto.ListIndex = 2
    Else
        sTipo = obtenerDeSQL("select tipopunto from documentoscae where puntoventa=" & ssTexto(punto) & " and tipo  in('FAA','FAB','NCA','NCB')")
        If sTipo = "PI" Then
            cmbPunto.ListIndex = 0
        ElseIf sTipo = "OL" Then
            cmbPunto.ListIndex = 1
        ElseIf sTipo = "WS" Then
            cmbPunto.ListIndex = 2
        ElseIf sTipo = "WS2" Then
            cmbPunto.ListIndex = 3
        Else
            cmbPunto.ListIndex = 2
        End If
        
    End If
        
        
    comboSql cboMoneda, "select descripcion, codigo from monedas order by codigo"
    comboSql cmbTipoIva, "select descripcion, codigo from ivas where activo = 1"
    comboSql cmbvendedor, "select descripcion, codigo  from usuarios where activo = 1 order by descripcion"
    comboArray cmbDeposito, Array("Deposito Central", "Deposito 1", "Deposito 2", "Deposito 3", "Deposito4"), Array(0, 1, 2, 3, 4)
'    uCuenta.ini "Select     "
    dtFecha = Date
    
    dtVencimiento = Date
    chkResaltar.Value = 0
    
    inigrilla 'grillas
    iniCliente 'clientes
    iniBotonera 'botonera
    
'    dtDesde = "1/1/" & Year(Date)
'    dtHasta = Date

    rsEjercicio.Open "SELECT * From Ejercicio WHERE activo =1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
   
    ucFechas.ini CDate(rsEjercicio!FechaInicio), CDate(rsEjercicio!FechaFin), ucefHorizontal, ucefFormatoSqlServer
    rsEjercicio.Close
    set_uProd
    
    optBuscarTipo.item(0).Value = "1" ' True
    
    If gEMPR_EmiteFacturaConRemito Then
        optStock.item(0).Value = "1" 'True '1
    Else
        optStock.item(0).Value = "1" 'True
    End If
    
    TabDetalle.Tab = 0

    If mFAE Then
        cboMoneda.ListIndex = 1
        cmbPunto.ListIndex = 2
    End If
    GRILLA.Editable = flexEDKbdMouse
    NPregunta = False
    NRespuesta = False
End Sub

Private Function FaltaCabecera() As Boolean
    'FaltaCabecera = (cliente.descripcion = "" Or cmbTipoIva = "" Or cmbFormaPago = "" Or ucCuit.text = "")
    FaltaCabecera = False
    
    If cliente.DESCRIPCION = "" Then
        che "falta cliente"
        cmbCliente.SetFocus
        FaltaCabecera = True
        Exit Function
    End If
    If cmbTipoIva = "" Then
        che "falta tipo iva"
        'cmbTipoIva.SetFocus
        FaltaCabecera = True
        Exit Function
    End If
    'If ucCuit.Text = "" Then
    If ucCuit = "" Then
        'If ComboCodigo(cmbTipoIva) <> TIPOIVA_ConsumidorFinal Then ' consummidor final
        'If TipoFormVenta(ComboCodigo(cmbTipoIva)) = "A" Then
            che "falta cuit"
            ucCuit.SetFocus
            FaltaCabecera = True
            Exit Function
        'End If
    End If
    If Trim(cmbformapago) = "" Then
        che "falta forma pago"
        cmbformapago.SetFocus
        FaltaCabecera = True
        Exit Function
    End If
    If CmbProvincia.Text = "" Then
        MsgBox "Falta ingresar la provincia.", , "ATENCION"
        FaltaCabecera = True
        Exit Function
    End If
    
End Function

Private Function FaltaGrilla() As Boolean
    'FaltaGrilla = (g.rows = 1 Or g.suma(gCANT) = 0 Or g.suma(gPUNI) = 0)
    FaltaGrilla = (g.rows = 1 Or g.suma(gCANT) = 0)
    
    If FaltaGrilla Then
        TabDetalle.Tab = 0
        GRILLA.SetFocus
''    Else
''        If uTipoVenta.Diferencia <> 0 Then
'''            che "Faltan imputaciones"
''            TabDetalle.Tab = 3
''            FaltaGrilla = True
''        End If
    End If
End Function

Private Function FaltaSeries() As Boolean
    Dim r As Integer, i As Long, ns As String

    If Not DeboModificarStock() Then Exit Function

    FaltaSeries = False
    r = gS.rows

    ' serie vacio
    If r > 1 And gS.buscar(gS_NSER, "") > 0 And chkSinSeries.Value <> vbChecked Then
        If TabDetalle.TabVisible(2) Then
        
            TabDetalle.Tab = 2
        
            MsgBox "Falta numero de serie"
            grillaSeries.SetFocus
            grillaSeries.Select gS.PrimerVacio(gS_NSER), gS_NSER
    
            FaltaSeries = True
            Exit Function
        End If
    End If

    'serie repetido
    If r > 1 Then
        For i = 1 To r - 2
            ns = gS.tx(i, gS_NSER)
            If ns <> "" And gS.buscar(gS_NSER, ns, i + 1) > 0 Then
                TabDetalle.Tab = 2
                grillaSeries.SetFocus
                grillaSeries.Select i, gS_NSER, gS.buscar(gS_NSER, ns, i + 1), gS_NSER
                MsgBox "Numero Serie repetido"
                'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                FaltaSeries = True
                Exit Function
            End If
        Next i
    End If
End Function

Private Sub BorrarCampos()
    On Error Resume Next
'    txtCodigo = ""
    TxtNroFactura = ""
    cliente.codigo = 0
    txtCodigo = ""
    ucCuit.Text = ""
    dtFecha = Date
    txtCotiLeyen.Text = ""
    ChkVaLeye.Value = 0
    
    FrmBorrarTxt Me
    txtneto = ""
    txtIva = ""
    txttotal = ""
    txtDescuento = ""
    lblSubTotal = ""
    g.Borrar
    gO.Borrar
    txtCotizacion = ""
    nroaux = 0
    Orden.Text = ""
    Remito.Text = ""
    chkResaltar.Value = 0
    NPregunta = False
    NRespuesta = False
    'ComboCodigo(
'''    gS.Borrar
End Sub
Private Sub HabilitarEdicion(habilitar As Boolean)
    fraCabecera.enabled = habilitar
    TabDetalle.enabled = habilitar
End Sub
Private Sub Resetear()
    mValoresOK = False
    'mContado = False
    optCuentaContado.item(0).Value = True
    BorrarCampos
    HabilitarEdicion False
    midDoc = 0
    mPropio = ""
'    uTipoVenta.Borrar
    cmbingresar.enabled = False
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    Set gO = New LiGrilla
    Set gS = New LiGrilla
    
    g.init GRILLA, 4
    gO.init grillaOrigen
    gS.init grillaSeries
    
    gCANT = g.AddCol(" Cantidad    ", "N", 4)
'    gALIA = g.AddCol(" Alias         ", "S")
    gprod = g.AddCol(" Producto            ")
    gDESC = g.AddCol(" Descripcion                                ", "S")
    gPUNI = g.AddCol(" P.Unitario    ", "N", 4)
    gPTOT = g.AddCol(" P.Total       ", "9", 2)
    gNPED = g.AddCol(" Pedido      ", "H") ', IIf(mFac = FacturaVenta_Pedido, "-", "H"))
    gNPCL = g.AddCol(" Pedido Clie ", "H") ', IIf(mFac = FacturaVenta_Pedido, "-", "H"))
    gNREM = g.AddCol(" Remito        ", IIf(mFac = FacturaVenta_Remito, "-", "H"))
    gFORM = g.AddCol(" Formula              ", "H")
    gITEM = g.AddCol(" itemPedidoRemito", "H")
    gIVA = g.AddCol(" P.IVA           ", "N", 2)
    gCTA = g.AddCol(" P.CTA           ", "S")
    GRILLA.ColComboList(gCTA) = ctasProducto
    gCTAd = g.AddCol(" P.CTA.DESCRIPCION           ", "S")
    
    gO_NPED = gO.AddCol(" Pedido     ", IIf(mFac = FacturaVenta_Pedido, "-", "H")) ' oculto si no es pedido
    g0_NPCL = gO.AddCol(" Pedido Clie ", IIf(mFac = FacturaVenta_Pedido, "-", "H")) ' oculto si no es pedido
    gO_NREM = gO.AddCol(" Remito      ", IIf(mFac = FacturaVenta_Remito, "-", "H")) ' oculto si no es   remito
'    gO_ITPE = gO.AddCol(" cod it pe ", "H")
'    gO_ITRE = gO.AddCol(" cod it re ", "H")
    gO_ITEM = gO.AddCol("it pe/re ", "H")
    gO_CANT = gO.AddCol(" Cantidad    ")
    gO_PEND = gO.AddCol(" Pendiente Facturar")
    gO_PROD = gO.AddCol(" Producto             ")
    gO_DESC = gO.AddCol(" Descripcion                              ")
    gO_PROP = gO.AddCol(" propio ", "H")
    
    
    If REMITO_CON_PRECIO Then
        gO_PREC = gO.AddCol(" Precio         ")
    Else
        gO_PREC = gO.AddCol(" Precio         ", "H")
    End If
    
    gS_ITEM = gS.AddCol(" Item ", "H")
    gS_PROD = gS.AddCol(" Producto                ")
    gS_NSER = gS.AddCol(" Num Serie          ", "S")
'    gS_CONS = gS.addCol(" Consig ")
    gS_HIDD = gS.AddCol(" h ", "H")
    grillaSeries.SelectionMode = flexSelectionListBox
    
End Sub

Private Sub iniCliente()
    Set cliente = New LiCodigo
    cliente.init cmbCliente, txtCodCliente, "Clientes", , , cmdCliente, "activo = 1 and categoria>1", True
    cliente.EditaDescripcion = True
End Sub
Private Sub iniBotonera()
'    Dim sMov As String
'    sMov = "select codigo from clientes where activo = 1"
    ucBoton.init True, True, False, True, True, , , False ', sMov, daTaenvironment1.Sistema
    ucBoton.CaptionEliminar = "Anular"
End Sub

Private Sub Grilla_CellChanged(ByVal Row As Long, ByVal Col As Long)
If ON_ERROR_HABILITADO Then On Error GoTo error
If GRILLA.TextMatrix(Row, gCTA) > "" Then
    GRILLA.TextMatrix(Row, gCTAd) = sSinNull(obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(GRILLA.TextMatrix(Row, gCTA))))
End If
error:
End Sub

Private Sub Grilla_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
'If Grilla.TextMatrix(Row, gCTA) > "" Then
    GRILLA.TextMatrix(Row, gCTAd) = sSinNull(obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(GRILLA.TextMatrix(Row, gCTA))))
'End If
End Sub

Private Sub Grilla_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
GRILLA.TextMatrix(Row, gCTAd) = sSinNull(obtenerDeSQL("select descripcion from cuentas where cuenta=" & ssTexto(GRILLA.TextMatrix(Row, gCTA))))
End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If g.Row > 0 Then g.delRow (g.Row)
End If
If KeyCode = 45 Then
    g.Row = g.addRow()
End If
End Sub

Private Sub grillaSeries_dblClick()
    Dim r As Long, prod As String, resu As String
    
    r = gS.Row
    If r < 1 Then Exit Sub
    prod = VerProductoMio(gS.tx(r, gS_PROD), Propio())
    If prod = "" Then Exit Sub
    
    ''''resu = Buscar_SeriesEnStock(prod)
    resu = SerieStockRepetida(prod)
    
    If resu > "" Then gS.tx r, gS_NSER, resu
End Sub


Private Sub optStock_Click(Index As Integer)
'    TabDetalle.TabVisible(2) = DeboModificarStock()
    TxtRemitoNumero.Visible = optStock(2).Value
End Sub

Private Sub CargaDatos()
'        cmbDeposito.ListIndex = s2n(!deposito)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset, i As Long, z As Double
    Dim varias As String
    Dim cheq As Boolean
    Dim sTipoPunto As String
    
    TabDetalle.Tab = 0
    With rs
        .Open "select *,Codigo, TipoDoc, NroFactura, Cliente, Fecha, " _
            & " neto, iva, PorcentajeIva, Total, Descuento, Moneda, cotizacion, ActualizaStock, Deposito, iddoc, iibb,_control_ve as orden,_docum_ve as remi,vendedor,variasfac,ND_xchequerechazado as cheq " _
            & " from FacturaVenta where codigo = " & txtCodigo _
            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        
        z = !cotizacion
        If z = 0 Then z = 1
        txtCotizacion = z
                
        txtIva = s2n(s2n(!Iva) / z)
        txtPdescuento.Text = (s2n(!Descuento, 4) * 100) / z
        txtneto = s2n(s2n(!Neto) / z)
        txttotal = s2n(s2n(!Total) / z)
        lblIIBB = s2n(!IIBB) / z
        txtPIVA = s2n(s2n(!PorcentajeIva, 4) * 100)
        midDoc = nSinNull(!iddoc)
        lblSubTotal = (!Neto / z) + (!Descuento / z)
        cliente.codigo = !cliente
        Remito.Text = Trim(sSinNull(!remi))
        Orden.Text = Trim(sSinNull(!Orden))
        txtNroCliente.Text = Trim(sSinNull(!NroCliente))
        varias = sSinNull(!variasfac)
        If IsNull(!Provincia) Or IsEmpty(!Provincia) Or Trim(!Provincia) = "" Then
        Else
            CmbProvincia = ObtenerDescripcionS("provincias", !Provincia)
        End If
        If !Vendedor >= 0 Then
            cmbvendedor.Text = obtenerDeSQL("select descripcion from usuarios where codigo=" & !Vendedor)
        Else
            cmbvendedor.Text = obtenerDeSQL("select descripcion from usuarios where codigo=0")
        End If
        
        txtPermisoEmbarque = sSinNull(rs!permisoembarque)
        lblCAE = sSinNull(!CAE)
        cboIncoterms.Text = sSinNull(!incoterms)
        
        sTipoPunto = obtenerDeSQL("select tipopunto from documentoscae where puntoventa=" & ssTexto(!PuntoVenta) & " and tipo=" & ssTexto(!TIPODOC))
        If sTipoPunto = "PI" Then
            cmbPunto.ListIndex = 0
        ElseIf sTipoPunto = "OL" Then
            cmbPunto.ListIndex = 1
        ElseIf sTipoPunto = "WS" Then
            cmbPunto.ListIndex = 2
        ElseIf sTipoPunto = "WS2" Then
            cmbPunto.ListIndex = 3
        Else
            cmbPunto.ListIndex = 0
        End If
        
        
        cboMoneda.ListIndex = BuscarEnCombo(cboMoneda, !moneda)
        
        'chkModStock.Value = IIf(!ActualizaStock, vbChecked, vbUnchecked)
'        If !actualizaStock Then
'            If gEMPR_EmiteFacturaConRemito Then
'                optStock.Item(ActuStock_RE) = True
'            Else
'                optStock.Item(ActuStock_SI) = True
'            End If
'        Else
            optStock.item(ActuStock_NO) = True
'        End If
        cheq = rs!cheq
        
        .Close
        .Open "select cantidad, codpropio, producto, Descripcion, formula, precioUnitario, PrecioTotal, NroRemito,_iva as iva from FacturaVentaDetalle where producto<>'1' and codigoFactura = " & txtCodigo & " order by id"
        g.Borrar
        While Not .EOF
            i = g.addRow()
            chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)
            If cheq = True And s2n(s2n(!PrecioUnitario, 4) / z) > 0 And s2n(!cantidad) = 0 Then
                g.tx i, gCANT, 1
            Else
                g.tx i, gCANT, s2n(!cantidad)
            End If
'            g.tx i, gALIA, AliasDeCodigo(!producto)
            g.tx i, gprod, VerProductoCliente(sSinNull(!producto), !codPropio, cliente.codigo)
            'g.tx i, gprod, sSinNull(!producto)
            'g.tx i, gDESC, ObtenerDescripcionS("producto", !Producto)
            g.tx i, gDESC, sSinNull(!DESCRIPCION)
            g.tx i, gPUNI, s2n(s2n(!PrecioUnitario, 4) / z)
            g.tx i, gPTOT, s2n(!PrecioTotal, 4) / z
            g.tx i, gFORM, sSinNull(!formula)
            g.tx i, gNREM, s2n(!NroRemito)
            g.tx i, gIVA, s2n(!Iva)
            
            'agregado trucho? ' cargo la descripcion del producto para las q grabamos sin descripcion
            If sSinNull(!DESCRIPCION) = "" And sSinNull(!producto) > "" Then g.tx i, gDESC, ObtenerDescripcionS("producto", !producto)
            
            .MoveNext
        Wend
        
        Dim Fact
        Dim a As Long
        a = 0
        Label26.caption = "ADJUNTAR FACTURAS:"
        If sSinNull(varias) <> "" Then
            Fact = Split(Replace(varias, "#", ""), ",")
            For i = 0 To UBound(Fact)
                Label26.caption = Label26.caption & Fact(i) & ","
                If a < Fact(i) Then a = Fact(i)
            Next
            Label26.Visible = True
            If TxtNroFactura < a Then
                txtIva = 0
                txtPdescuento.Text = 0
                txtneto = 0
                txttotal = 0
                lblIIBB = 0
                txtPIVA = 0
                lblSubTotal = 0
            Else
                .Close
                .Open "select Codigo, TipoDoc, NroFactura, Cliente, Fecha, " _
                    & " neto, iva, PorcentajeIva, Total, Descuento, Moneda, cotizacion, ActualizaStock, Deposito, iddoc, iibb,_control_ve as orden,_docum_ve as remi,vendedor,variasfac " _
                    & " from FacturaVenta where codigo = " & txtCodigo _
                    , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                z = !cotizacion
                If z = 0 Then z = 1
                txtCotizacion = z
                        
                txtIva = s2n(s2n(!Iva) / z)
                txtPdescuento.Text = (s2n(!Descuento, 4) * 100) / z
                txtneto = s2n(s2n(!Neto) / z)
                txttotal = s2n(s2n(!Total) / z)
                lblIIBB = s2n(s2n(!IIBB) / z)
                txtPIVA = s2n(s2n(!PorcentajeIva, 4) * 100)
                lblSubTotal = (!Neto / z) + (!Descuento / z)
            End If
        Else
            Label26.Visible = False
        End If
        
        
    End With
    GoTo fin
ufaErr:
    ufa "err leyendo datos", Me.Name & txtCodigo ', Err
fin:
    Set rs = Nothing
End Sub


Private Sub CargaOrigen()
    Dim s As String, i As Long, rs As New ADODB.Recordset
    
    gO.Borrar
    Select Case mFac ' alias nPedi nRemi trae num remito o pedido
    Case FacturaVenta_Pedido
    
    s = " Select " _
        & " ItemPedidoCliente.codigo as cod, numero as nPedi, pedido_cli as nPediCli, producto, cantidad, precio, facturar, CodigoPropio as codPropio, 0 as nRemi" _
        & " from Pedidos_Clientes inner join ItemPedidoCliente " _
        & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
        & " where facturar > 0 " _
        & " and formula  = '' " _
        & " and cliente = " & cliente.codigo _
        & " and Pedidos_clientes.activo = 1 " _
        & " order by npedi, cod"
        
    Case FacturaVenta_Remito
      s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi, producto, cantidad, precio, facturar, codPropio, 0 as nPedi, 0 as nPediCli" _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and formula = '' " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
        ' And activo = 1
    
    Case FacturaVenta_Remito
      s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi,  producto, cantidad, precio, facturar, codPropio, 0 as nPedi, 0 as nPediCli " _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
     
    Case Else
        Exit Sub
    End Select
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        i = gO.addRow
        gO.tx i, gO_ITEM, rs!COD
        gO.tx i, gO_NPED, rs!nPedi
        gO.tx i, g0_NPCL, rs!nPedicli
        gO.tx i, gO_NREM, rs!nRemi
        gO.tx i, gO_PROD, VerProductoCliente(rs!producto, rs!codPropio, cliente.codigo)
        gO.tx i, gO_DESC, ObtenerDescripcionS("producto", rs!producto)
        gO.tx i, gO_CANT, rs!cantidad
        gO.tx i, gO_PEND, rs!facturar
        gO.tx i, gO_PREC, rs!precio
        gO.tx i, gO_PROP, rs!codPropio
                
        rs.MoveNext
    Wend
    
    Set rs = Nothing
End Sub


Private Function Propio()
    Propio = (chkPropio.Value = vbChecked)
End Function

Private Sub MetoEnGrilla(prod, desc, cant, prec, pedi, remi, item, iva_prod As Double, Optional ivaprod, Optional conDesglose As Boolean) ', cons)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim i As Long, rs As New ADODB.Recordset, ssql As String, codigomio As String, hay As Double, suma As Double ', ivaprod As Double

   
    codigomio = VerProductoMio(prod, Propio())
    
    
'    If CodigoMio > "" Then  ' debo sacar el iva de tabla producto, o ponerlo a mano.
'    End If
    If cant <> 0 Then
        If mClienteConIva Then
            If IsMissing(iva_prod) Then
                iva_prod = obtenerDeSQL("select iva from producto where codigo = '" & codigomio & "'") * 100
                
'                If ivaprod = 0 Then Stop
                
            End If
            
            suma = g.suma(gPUNI)
            If suma = 0 Then
                txtPIVA = iva_prod
            Else
                
                If s2n(iva_prod) <> s2n(txtPIVA) Then
                    If UCase(txtTipoDoc) = "FAB" Then
                    Else
                        'che "Producto con distinto IVA al cargado"
                        'Exit Sub
                    End If
                End If
            End If
        Else
            'ivaprod = 0
            txtPIVA = "0"
        End If
    End If

    
    i = g.addRow()
    g.tx i, gprod, prod
    g.tx i, gDESC, desc
    g.tx i, gCANT, cant
    g.tx i, gPUNI, prec
    g.tx i, gPTOT, prec * cant
    g.tx i, gNPED, pedi
    g.tx i, gNPCL, sSinNull(obtenerDeSQL("select pedido_cli from Pedidos_Clientes where numero = " & pedi))
    g.tx i, gNREM, remi
    g.tx i, gITEM, item
    g.tx i, gIVA, iva_prod
'    g.tx i, gCONS, cons
   
    ' verifico stock
'    MsgFalta prod, cant
    
    ' verifico stock si no es virtual
    If Not gEMPR_FormulaEsVirtual Then MsgFalta prod, cant
    
    MetoGrillaSeries codigomio, s2n(cant) ', s2n(cons)
    
'    If gEMPR_FormulaEsVirtual Then
    ssql = "select codigo, componente, cantidad from formulas where activo = 1 and codigo = '" & codigomio & "'"
    If conDesglose Then
        With rs
            .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

            While Not .EOF
                MsgFalta !Componente, !cantidad * cant

                i = g.addRow()
                g.tx i, gprod, VerProductoCliente(!Componente, Propio(), cliente.codigo)
                g.tx i, gDESC, "    -" & ObtenerDescripcionS("Producto", !Componente)
                g.tx i, gCANT, cant * !cantidad
                g.tx i, gNPED, pedi
                g.tx i, gPUNI, 0
                g.tx i, gPTOT, 0
                g.tx i, gFORM, prod
                g.tx i, gITEM, item
                g.tx i, gIVA, 0
                'g.tx i, gCONS, cons
                MetoGrillaSeries !codigo, cant * !cantidad ', cons
                .MoveNext
            Wend
        End With
    End If
    
fin:
    RevisarTotales
    Exit Sub
ufaErr:
    ufa "err al poner en grilla", Me.Name ', Err
    Resume fin
End Sub

Private Sub MetoGrillaSeries(prod As String, ByVal cant As Long) ', ByVal consig As long)
    On Error GoTo ERR_FIN
    Dim i As Long, r As Long ', c As long

    If Not DeboModificarStock() Then Exit Sub

    If ProductoConSerie(prod) Then
        For i = 1 To cant
            r = gS.addRow()
            gS.tx r, gS_PROD, prod
        Next i
    End If
    
fin:
    Exit Sub
ERR_FIN:
    ufa "err en series ", Me.Name ', Err
    Resume fin
End Sub

Private Sub MsgFalta(CodProd, canti)
    Dim hay As Double
    
    If Not DeboModificarStock() Then Exit Sub
    If CodProd = "" Then Exit Sub
    
    If mFac = FacturaVenta_Remito Or mFac = FacturaVenta_NCreditoDevolucion Then Exit Sub
    hay = HayProducto(VerProductoMio(CodProd, Propio()), cmbDeposito.ItemData(cmbDeposito.ListIndex))
    If hay < canti Then
        MsgBox " Stock para " & VerProductoCliente(CStr(CodProd), Propio(), cliente.codigo) & ", " & cmbDeposito.Text & "  : " & hay & vbCrLf & vbCrLf & " requeridos : " & canti
    End If
End Sub

Private Sub chkPropioEnabled(que As Boolean)
    chkPropio.enabled = GRILLA.rows < 2 And que
End Sub


Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim rs As New ADODB.Recordset
   
    With g
        .tx Row, gPTOT, s2n(.tx(Row, gCANT), 4) * s2n(.tx(Row, gPUNI), 4)
    End With
    
    If Trim(txtTipoDoc.Text) = "NCA" Or Trim(txtTipoDoc.Text) = "NDA" Then
        rs.Open "select * from facturaventadetalle where codigofactura=" & txtCodigo, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True And rs.BOF = True Then
        Else
            If rs!PrecioTotal = 0 Then
            Else
                RevisarTotales
            End If
        End If
        Set rs = Nothing
    Else
        RevisarTotales
    End If
End Sub

Private Sub gO_DblClick()
    'Static sPropio As String
    Dim prod, desc, cant, prec, pedi, remi, item, prop, r As Long, PIVA As Double
    Dim cantParaPasar As Double
    
    
    With gO
        r = .Row
        If r > 0 Then
            prod = .tx(r, gO_PROD)
            desc = .tx(r, gO_DESC)
            cant = .tx(r, gO_PEND)
            prec = IIf(BS_RESPETAR_PRECIO_OC, .tx(r, gO_PREC), 0)
            pedi = .tx(r, gO_NPED)
            remi = .tx(r, gO_NREM)
            item = .tx(r, gO_ITEM)
            prop = .tx(r, gO_PROP)
            PIVA = s2n(nSinNull(obtenerDeSQL("select iva from producto where codigo=" & ssTexto(prod))) * 100)

            If mPropio = "" Then
                mPropio = CStr(prop)
                chkPropio.Value = IIf((prop = "1"), vbChecked, vbUnchecked)
            Else
                If mPropio <> CStr(prop) Then
'                    MsgBox "No puedo mezclar codigo Propio con Codigo Cliente"
'                    Exit Sub
                    If prop = "True" Then
                        prod = VerProductoCliente(CStr(prod), False, cliente.codigo)
                    Else
                        prod = VerProductoMio(prod, True)
                    End If
'                    If prod = "" Then
'                        che "err al buscar codigo"
'                        Exit Sub
'                    End If
                End If
            End If
                        
            If s2n(prec) = 0 Then prec = precioProducto(CStr(prod), Propio(), cliente.codigo)
            If cant > 1 And BS_PED_A_FAC_PREGUNTA_CANT Then
                cant = s2n(InputBox("Cantidad : ", prod & " A facturar", cant))
            End If
            
            Dim dTipo As String
            Dim PU2 As Double
            dTipo = UCase(txtTipoDoc)
                    
            If InStr(dTipo, "B") Then
                PU2 = s2n(prec * (1 + s2n(PIVA / 100)), 4)
            Else
                PU2 = prec
            End If
                    
            MetoEnGrilla prod, desc, cant, PU2, pedi, remi, item, PIVA, , False
            
            .delRow r
        End If
    End With
End Sub

Private Sub TabDetalle_Click(PreviousTab As Integer)
    Dim i As Long, j As Long, prod As String, cant As Long, cons

    If TabDetalle.Tab <> 2 Or PreviousTab = 2 Then Exit Sub

    With gS
        'borrosinserie
        i = 1
        While i < .rows
            If .tx(i, gS_NSER) = "" Then
                .delRow i
                i = i - 1
            End If
            i = i + 1
        Wend

        'borro marcas
        For i = 1 To .rows - 1
            grillaSeries.TextMatrix(i, gS_HIDD) = ""
        Next i

        'marco o agrego en grilla series
        For i = 1 To g.rows - 1
            prod = Trim(g.tx(i, gprod))
            cant = s2n(g.tx(i, gCANT))
            'cons = s2n(g.tx(i, gCONS))
            If ProductoConSerie(prod, Propio()) Then
                For j = 1 To cant
                    If marcoG3(prod) Then ', cons) Then
    '                    cons = cons - 1
                    End If
                Next j
            End If
        Next i

        'borro no marcadas
        i = 1
        While i < .rows
            If .tx(i, gS_HIDD) = "" Then
                .delRow i
                i = i - 1
            End If
            i = i + 1
        Wend
    End With
End Sub

Private Function marcoG3(codi) ', ByVal cons) '
    Dim i As Long
    
    For i = 1 To gS.rows - 1
        If gS.tx(i, gS_PROD) = codi And gS.tx(i, gS_HIDD) = "" Then
            grillaSeries.TextMatrix(i, gS_HIDD) = "X"
            Exit Function
        End If
    Next i
    i = gS.addRow()
    grillaSeries.TextMatrix(i, gS_PROD) = codi
    grillaSeries.TextMatrix(i, gS_HIDD) = "X"
'    grillaSeries.TextMatrix(i, g3CONS) = IIf(cons > 0, "-1", "0")
   
'    marcoG3 = (g3.tx(i, g3CONS) = "-1")
End Function
Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtCodCliente_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtCotiLeyen_LostFocus()
    txtCotiLeyen.Text = s2n(txtCotiLeyen, 2, True)
End Sub

Private Sub txtDireccion_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtFlete_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtFlete_LostFocus()
RevisarTotales
End Sub

Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub

Private Sub txtSeguro_LostFocus()
RevisarTotales
End Sub

Private Sub txtIvaProducto_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtlocalidad_GotFocus()
    PintoFocoActivo
End Sub

Private Sub TxtNroFactura_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtNroFacturaRef_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtPIVA_LostFocus()
    txtPIVA = n2r(s2n(txtPIVA))
    RevisarTotales
End Sub

Private Sub txtprecio_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtTipoDoc_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtTipoDocRef_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtPdescuento_LostFocus()
    txtPdescuento = n2r(s2n(txtPdescuento))
    RevisarTotales
End Sub

Private Function FaltaAlgo() As Boolean
    FaltaAlgo = True
    
'    If txtCodigo <> "" Or TxtNroFactura <> "" Then
'        che "codigo ya grabado"
'        Exit Function
'    End If
    If Not PuedoVentas(dtFecha) Then
        'msg en funcion
        Exit Function
    End If


    If ComboCodigo(cboMoneda) > 1 And s2n(txtCotizacion) = 0 Then
        che "Falta cotizacion"
        Exit Function
    End If
    If FaltaCabecera() Then
'        che "Faltan datos en el formulario"
        Exit Function
    End If
    If HayProdEnEdicion(uProd.DESCRIPCION) Then
        uProd.SetFocus
        Exit Function
    End If
    If FaltaGrilla() Then
        che "Faltan datos en la grilla"
        Exit Function
    End If
    If Trim$(txtPIVA) = "" Then
        txtPIVA = InputBox("Falta Iva de la factura" & vbCrLf & "si no tiene ingrese 0")
        If Trim$(txtPIVA) > "" Then txtPIVA = s2n(txtPIVA)
        Exit Function
    End If
    If FaltaSeries() Then
        Exit Function
    End If
    
    FaltaAlgo = False
End Function


Private Function EmitirRemito() As Long
    Dim tmp, formula As String, num As Long, sucursal As Long, depot As Long, i As Long, mjstk As Long, produ As String
    
    sucursal = s2n(obtenerDeSQL("select sucursal from datos"))
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    num = s2n(TxtRemitoNumero)
    
    tmp = obtenerDeSQL("select cliente from RemitoVenta where Numero = " & s2n(TxtRemitoNumero))
    If Not IsEmpty(tmp) Then
        che "Numero Remito ya grabado, para cliente " & tmp
        Exit Function
    End If
    
    '*******************************************************************
    ' quiero Transaccion !  , y/o quiero hacer tabla temp y 1 solo stored
'    GrabarBs_Num CAMPO_BS_NroREMITO, Num
    '
    'daTaenvironment1.dbo_abmRemitoVenta "A", s2n(TxtRemitoNumero), cliente.Codigo, (dtFecha),  0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtobs(0), txtobs(1), txtobs(2), txtobs(3)
    'DataEnvironment1.dbo_abmRemitoVenta "A", num, cliente.codigo, (dtFecha), 0, 0, depot, Propio(), "", "", "", ""
    frmRemitoVenta.ABMRemitoVenta "A", num, cliente.codigo, (dtFecha), 0, 0, depot, Propio(), "", "", "", "", "", 0
    '
    For i = 1 To g.rows - 1 'items
'        cantConsign = s2n(g.tx(i, gCONS))
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gprod), Propio())
        mjstk = ManejaStock(produ)
        frmRemitoVenta.ABMRVDetalle "A", num, produ _
            , s2n(g.tx(i, gCANT)), s2n(g.tx(i, gPUNI)), s2n(g.tx(i, gNPED)) _
            , depot, 0, formula, mjstk, 0 'cantConsign, formula
    Next i
    
    ' quiero Transaccion, y/o quiero hacer tabla temp y 1 solo stored
    '*******************************************************************

'    MsgBox "Remito " & num & " grabado"
    EmitirRemito = s2n(TxtRemitoNumero)
End Function
Private Function CalculoLineas(Texto As Object, Linea As Long) As Long
    Dim Bloque As String
    'Numero de caracteres = NumC
    'Numero de Bloques = NumB
    Dim NumC, NumB As Integer
    NumC = Len(Texto.Text)
    If NumC > Linea Then
        NumB = NumC \ Linea
        For i = 0 To NumB
            Texto.SelStart = (Linea * i)
            Texto.SelLength = Linea
            Bloque = Texto.SelText
            'Printer.Print Bloque
            'Debug.Print Bloque
            'Debug.Print i & " * " & Len(RTrim(mpppp(i))) & " *** " & RTrim(mpppp(i))
        Next i
    Else
        'Printer.Print Texto.Text
        Debug.Print Texto.Text
        i = 1
    End If
    Printer.EndDoc
    CalculoLineas = i
End Function


' **********************************************
Private Sub ucBoton_AceptarAlta()
Dim doc As New FacturaElectronica, EsGranEmpresa As Boolean
Dim MontoImponible As Double
    logFacturacion "Boton Aceptar", "", txtTipoDoc & " " & TxtNroFactura
   If ON_ERROR_HABILITADO Then On Error GoTo UFAalta

    
    Dim NroRemito As Long, tmp, tmpfec As Date, QuieroLeyenda As Boolean
    Dim sAssert As String, x As Long, i As Long, a As Long, h As Long
    Dim varias As Boolean, lineas As Long, Caracter As Long, str As String
    Dim sPunto As String, sTipo As String
    logFacturacion "Validacion cotizacion", "", txtTipoDoc & " " & TxtNroFactura
    If s2n(txtCotizacion) = 0 Then txtCotizacion = 1
    logFacturacion "Verificacion cuit gran empresa", "", txtTipoDoc & " " & TxtNroFactura
    EsGranEmpresa = doc.EstaEnPadron(ucCuit.Text, MontoImponible)
    logFacturacion "Fin Verificacion cuit gran empresa", "", txtTipoDoc & " " & TxtNroFactura
    If EsGranEmpresa And s2n(txttotal * s2n(txtCotizacion)) >= MontoImponible Then
        logFacturacion "Validacion monto gran empresa superado", "", txtTipoDoc & " " & TxtNroFactura
        If Trim(txtTipoDoc) <> "FEA" And Trim(txtTipoDoc) <> "CEA" And Trim(txtTipoDoc) <> "DEA" Then
            logFacturacion "Validacion comprobante FEA,CEA,DEA", "", txtTipoDoc & " " & TxtNroFactura
            If PuntoVentaTipo(cmbPunto.ListIndex) = "WS" Then
                logFacturacion "Validacion punto venta WS", "", txtTipoDoc & " " & TxtNroFactura
                If MsgBox("El importe del comprobante no puede superar los $" & MontoImponible & " para este cliente.¿Desea emitir una factura de credito?", vbInformation + vbYesNo, "ATENCION") = vbYes Then
                    logFacturacion "Validacion comprobante mensaje supera monto graba nota de credito", "", txtTipoDoc & " " & TxtNroFactura
                    logFacturacion "Busco nro y tipo comprobante por cliente GE", "", txtTipoDoc & " " & TxtNroFactura
                    BuscoNroYTipo True, , EsGranEmpresa
                    logFacturacion "Fin busco nro y tipo comprobante", "", txtTipoDoc & " " & TxtNroFactura
                Else
                    logFacturacion "No graba por superar monto FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
                    Exit Sub
                End If
            End If
        End If
    ElseIf s2n(lblSubTotal) < MontoImponible Then
        logFacturacion "Validacion monto gran empresa NO superado", "", txtTipoDoc & " " & TxtNroFactura
        EsGranEmpresa = False
        logFacturacion "Se indica que no es GE", "", txtTipoDoc & " " & TxtNroFactura
        logFacturacion "Busco nro y tipo comprobante por cliente NO es GE", "", txtTipoDoc & " " & TxtNroFactura
        BuscoNroYTipo True, , EsGranEmpresa
        logFacturacion "Fin busco nro y tipo comprobante", "", txtTipoDoc & " " & TxtNroFactura
    End If
    
    logFacturacion "Verificacion lineas de la factura", "", txtTipoDoc & " " & TxtNroFactura
    lineas = 22 '25
    Caracter = 75
    i = 1
    a = 0
    varias = False
    While i < GRILLA.rows
        Text1.Text = ""
        Text1.Text = GRILLA.TextMatrix(i, 2) 'Text1.Text & grilla.TextMatrix(i, 2) & Chr(13)
'        i = i + 1
    'Wend
'    i = 1
'    While i < grilla.rows
        logFacturacion "Calculo de lineas", "", txtTipoDoc & " " & TxtNroFactura
        a = a + CalculoLineas(Text1, Caracter)
        i = i + 1
    Wend
    logFacturacion "Fin Verificacion lineas de la factura", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Verificacion punto de venta", "", txtTipoDoc & " " & TxtNroFactura
    sTipo = PuntoVentaTipo(cmbPunto.ListIndex) 'con esto veo tipopunto
    If Trim(sTipo) = "WS" Or Trim(sTipo) = "OL" Or Trim(sTipo) = "WS2" Then
    Else
        If a >= lineas Then
            If MsgBox("Las lineas superan el espacio de impresion." & Chr(13) & "Desea grabarla en varias?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
                varias = True
                logFacturacion "Supera lineas graba igual", "", txtTipoDoc & " " & TxtNroFactura
            Else
                logFacturacion "Supera lineas no graba FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
                Exit Sub
            End If
        End If
    End If
    logFacturacion "Fin Verificacion punto de venta", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Verificacion fecha se puede por ejercicio", "", txtTipoDoc & " " & TxtNroFactura
    If TrabaIva(dtFecha.Value) Then
        MsgBox "La fecha del comprobante esta dentro de las fechas trabadas para emision," & Chr(13) & "verifiquelo con su contadora.", , "ATENCION"
        logFacturacion "fecha trabada FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
        Exit Sub
    End If
    logFacturacion "Fin Verificacion fecha se puede por ejercicio", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Verificacion falta algo", "", txtTipoDoc & " " & TxtNroFactura
    If FaltaAlgo() Then Exit Sub
    logFacturacion "Fin Verificacion falta algo", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Verificacion es contado", "", txtTipoDoc & " " & TxtNroFactura
    If EsContado() And Not mValoresOK Then
        mValoresOK = frmValores.mostrar(s2n(txttotal))
        logFacturacion "Carga de valores por efectivo", "", txtTipoDoc & " " & TxtNroFactura
        If Not mValoresOK Then
            logFacturacion "No se carga valores FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
            Exit Sub
        End If
    End If
    logFacturacion "Fin Verificacion es contado", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Mensaje imprime leyenda", "", txtTipoDoc & " " & TxtNroFactura
    If mFAE Then
        If confirma("Imprime leyenda? ") Then
            QuieroLeyenda = True
            frmLeyendaFAE.Show vbModal
        End If
    Else
        'esto es solo para el alta, si entro aca es porque la reimprimi
        If confirma("Imprime leyenda? ") Then
            QuieroLeyenda = True
aca:
            leye = 0 'variable global
            If varias = True Then
                frmFactLeyenda.NroFac = s2n(txtCodigo + (ItemHoja(a, lineas, False, Caracter)) - 1)
                frmFactLeyenda.Label1.caption = s2n(txtCodigo + (ItemHoja(a, lineas, False, Caracter)) - 1)
            Else
                frmFactLeyenda.NroFac = s2n(txtCodigo)
                frmFactLeyenda.Label1.caption = s2n(txtCodigo)
            End If
            frmFactLeyenda.Show vbModal
            If leye = 1 Then
                'MsgBox "La leyenda se cargo con exito.", , "ATENCION"
            ElseIf leye = 0 Then
                If MsgBox("Esta seguro de cancelar la leyenda?", vbYesNo + vbQuestion, "ATENCION") = vbNo Then
                    GoTo aca
                End If
            End If
        End If
    End If
    logFacturacion "Fin Mensaje imprime leyenda", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Verificacion si existe punto venta", "", txtTipoDoc & " " & TxtNroFactura
    sTipo = PuntoVentaTipo(cmbPunto.ListIndex)
    sPunto = sSinNull(obtenerDeSQL("select puntoventa from documentoscae where tipo=" & ssTexto(txtTipoDoc) & " and tipopunto=" & ssTexto(sTipo)))
    If sPunto = "" Then
        MsgBox "No existe punto de venta para " & txtTipoDoc & "-" & sTipo
        logFacturacion "No existe punto de venta FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
        Exit Sub
    End If
    logFacturacion "Fin Verificacion si existe punto venta", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Busco nro y tipo comprobante", "", txtTipoDoc & " " & TxtNroFactura
    If Not BuscoNroYTipo(True, , EsGranEmpresa) Then
        MsgBox "Se actualizara el numero de factura, para que pueda grabar.", , "ATENCION"
        TxtNroFactura = TxtNroFactura + 1
        logFacturacion "Se actualiza Nro factura FIN ACEPTAR", "", txtTipoDoc & " " & TxtNroFactura
        Exit Sub   ' de nuevo, por las dudas (si fuera multiusuario habria q meter mas control aun)
    End If
    logFacturacion "Fin Busco nro y tipo comprobante", "", txtTipoDoc & " " & TxtNroFactura
    
    If confirma("Factura: " & vbCrLf & vbCrLf & "Tipo :  " & txtTipoDoc & vbCrLf & "Nro :   " & TxtNroFactura & vbCrLf & "Confirma factura ?") Then
   
        If varias = False Then
            logFacturacion "Inicia guardado una hoja de factura", "", txtTipoDoc & " " & TxtNroFactura
            If GrabaFactura() Then          ' graba tamb valores contado
                
                frmFactLeyenda.AceptarLeyenda txtCodigo
                
                If gEMPR_ConSistContable Then
    '
                    If FrmCostosYContable.grillacostos.rows > 1 And FrmCostosYContable.grillacostos.TextMatrix(1, 1) > "" Then
                        sAssert = " dbo_INGCENTROCOSTOS "
                        
                        'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                        For x = 1 To FrmCostosYContable.grillacostos.rows - 1
                            DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(FrmCostosYContable.grillacostos.TextMatrix(x, 0)), _
                            dtFecha, txtTipoDoc, val(TxtNroFactura), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3), 3), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) + s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, "", FrmCostosYContable.grillacostos.TextMatrix(x, 4), 0
                                                                
                        Next
                        FrmCostosYContable.LimpioControles
                        FrmCostosYContable.InicioGrillaCostos
                        Unload FrmCostosYContable
                    End If
                End If
                logFacturacion "Inicia guardado remito de factura", "", txtTipoDoc & " " & TxtNroFactura
                If VaConRemito Then
                    NroRemito = EmitirRemito()
                    logFacturacion "Fin de remito", "", txtTipoDoc & " " & TxtNroFactura
                Else
                    logFacturacion "No va con remito", "", txtTipoDoc & " " & TxtNroFactura
                End If
    ''            DE_CommitTrans
                TabDetalle.Tab = 0
                
                
           'CAE
            'If mFAE Then
                logFacturacion "Inicia Asignacion de cae", "", txtTipoDoc & " " & TxtNroFactura
                If siFCAE(Trim(txtTipoDoc), PuntoVentaTipo(cmbPunto.ListIndex)) Then
                    logFacturacion "Corresponde Asignar CAE", "", txtTipoDoc & " " & TxtNroFactura
                    If MsgBox("¿Asignar CAE?", vbYesNo + vbInformation) = vbYes Then
                        logFacturacion "Confirma Asignar CAE", "", txtTipoDoc & " " & TxtNroFactura
                        If doc.EmisionFacturaElectronica(s2n(txtCodigo), txtPermisoEmbarque) Then
                            logFacturacion "Correcto Asignar CAE", "", txtTipoDoc & " " & TxtNroFactura
                            logFacturacion "Impresion comprobante", "", txtTipoDoc & " " & TxtNroFactura
                            If Right(Trim(txtTipoDoc), 1) = "E" Then
                                ImprimirComprobanteFE s2n(txtCodigo)
                            Else
                                ImprimirComprobanteFE2 s2n(txtCodigo)
                            End If
                            logFacturacion "Fin Impresion comprobante", "", txtTipoDoc & " " & TxtNroFactura
                        Else
                            logFacturacion "NO Correcto Asignar CAE", "", txtTipoDoc & " " & TxtNroFactura
                            If MsgBox("GENERAR FACTURA EN PDF", vbYesNo + vbInformation) = vbYes Then
                                logFacturacion "Impresion comprobante", "", txtTipoDoc & " " & TxtNroFactura
                                If Right(Trim(txtTipoDoc), 1) = "E" Then
                                    ImprimirComprobanteFE s2n(txtCodigo)
                                Else
                                    ImprimirComprobanteFE2 s2n(txtCodigo)
                                End If
                                logFacturacion "Fin Impresion comprobante", "", txtTipoDoc & " " & TxtNroFactura
                            End If
                        End If
                    Else
                        logFacturacion "NO Acepto Asignar CAE", "", txtTipoDoc & " " & TxtNroFactura
                    End If
                Else
                    logFacturacion "NO Corresponde Asignar CAE o no Acepto", "", txtTipoDoc & " " & TxtNroFactura
                    Dim eCUIT As String, eCODFACTURA As String, ePUNTOVENTA As String, eCAE As String, eBARRA As String, eFECHAVENCE As String
                    Dim IDFactura As Long
                    If cmbPunto.ListIndex = 1 Then
                        If lblCAE > "" Then
                            GoTo IMPRIMIR_FE
                        Else
                            lblCAE = sSinNull(obtenerDeSQL("SELECT CAE FROM FACTURAVENTA WHERE CODIGO=" & s2n(txtCodigo)))
                            If lblCAE > "" Then
                                GoTo IMPRIMIR_FE
                            End If
                        End If
                        logFacturacion "Asignar CAE Manual", "", txtTipoDoc & " " & TxtNroFactura
                        If MsgBox("¿Desea cargar CAE manualmente?", vbYesNo + vbInformation) = vbYes Then
                            logFacturacion "Si Asignar CAE Manual", "", txtTipoDoc & " " & TxtNroFactura
                            IDFactura = txtCodigo
                            eCUIT = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
                            eCUIT = Format(eCUIT, "00000000000")
                            ePUNTOVENTA = Format(Trim(obtenerDeSQL("select puntoventa from documentoscae where tipopunto='OL' and tipo=" & ssTexto(txtTipoDoc))), "0000")
                            eCODFACTURA = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='OL' and tipo=" & ssTexto(txtTipoDoc) & " and puntoventa=" & ssTexto(ePUNTOVENTA))), "00")
                            eCAE = InputBox("INGRESE NUMERO DE CAE (14 DIGITOS)...", "CAE OTORGADO ONLINE")
                            While Len(eCAE) < 14
                                eCAE = InputBox("INGRESE NUMERO DE CAE (14 DIGITOS)...", "CAE OTORGADO ONLINE")
                            Wend
                            eFECHAVENCE = InputBox("INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                            While Len(eFECHAVENCE) < 10
                                eFECHAVENCE = InputBox("INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                            Wend
                            While CDate(eFECHAVENCE) < dtFecha
                                eFECHAVENCE = InputBox("FECHA MENOR AL COMPROBANTE" & Chr(13) & "INGRESE FECHA DE VENCIMIENTO DE CAE " & Chr(13) & "<<NO OMITA EL FORMATO(DD/MM/AAAA)>>...", "CAE OTORGADO ONLINE")
                            Wend
                            eFECHAVENCE = afipFecha(CDate(eFECHAVENCE))
                            'bBARRA = bCuit & bCodFactura & bPuntoVenta & bCAE & bFechaCAE & "8" 'MOMENTAMEAMENTE VA 8 FIJO HASTA QUE HABERIGUE DE DONDE SALE
                            eBARRA = eCUIT & eCODFACTURA & ePUNTOVENTA & eCAE & eFECHAVENCE
                            eBARRA = eBARRA & CodVerificador(eBARRA)
                            
                '            .FE.FERespuestaDetalleFecha_vto
                            DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(eBARRA) & ",caev=" & ssFecha(aFecha(eFECHAVENCE)) & ", cae=" & ssTexto(eCAE) & " where codigo=" & IDFactura
                            
IMPRIMIR_FE:
                            logFacturacion "Impresion de comprobante", "", txtTipoDoc & " " & TxtNroFactura
                            If MsgBox("GENERAR FACTURA EN PDF", vbYesNo + vbInformation) = vbYes Then
                                If Right(Trim(txtTipoDoc), 1) = "E" Then
                                    ImprimirComprobanteFE s2n(txtCodigo)
                                Else
                                    ImprimirComprobanteFE2 s2n(txtCodigo)
                                End If
                            End If
                            logFacturacion "Fin Impresion de comprobante", "", txtTipoDoc & " " & TxtNroFactura

                        End If
                    End If
                End If
          
                
                ucBoton.AceptarOk

            End If
        Else
            logFacturacion "Inicia guardado de mas de una hoja de factura", "", txtTipoDoc & " " & TxtNroFactura
            If GrabaFactura2(ItemHoja(a, lineas, False, Caracter), sPunto) Then       ' graba tamb valores contado
                If gEMPR_ConSistContable Then
                    If FrmCostosYContable.grillacostos.rows > 1 And FrmCostosYContable.grillacostos.TextMatrix(1, 1) > "" Then
                        sAssert = " dbo_INGCENTROCOSTOS "
                        
                        'ALTA A LOS DETALLES (MATRIZ) DE CENTRO DE COSTOS
                        For x = 1 To FrmCostosYContable.grillacostos.rows - 1
                            DataEnvironment1.dbo_INGCENTROCOSTOS "A", val(FrmCostosYContable.grillacostos.TextMatrix(x, 0)), _
                            dtFecha, txtTipoDoc, val(TxtNroFactura), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3), 3), s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 2)) + s2n(FrmCostosYContable.grillacostos.TextMatrix(x, 3)), Date, 0, UsuarioSistema!codigo, 0, 1, "", FrmCostosYContable.grillacostos.TextMatrix(x, 4), 0
                                                                
                        Next
                        FrmCostosYContable.LimpioControles
                        FrmCostosYContable.InicioGrillaCostos
                        Unload FrmCostosYContable
                    End If
                End If
                
                If VaConRemito Then NroRemito = EmitirRemito()
                TabDetalle.Tab = 0
                If gEMPR_idEmpresa = 11 Then
                    ImprimirComprobThor (s2n(txtCodigo))
                ElseIf gEMPR_idEmpresa = 6 And (Trim(txtTipoDoc) = "FAE" Or Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A) Then
                    ImprimirAMRAT (s2n(txtCodigo))  '******** esto esta listo solo tienen q avisar para usarlo
                ElseIf gEMPR_idEmpresa = 4 And (Trim(txtTipoDoc) = "FAB" Or Trim(txtTipoDoc) = "FAA" Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_B Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_B Or Trim(txtTipoDoc) = TipoDoc_NCREDITO_A Or Trim(txtTipoDoc) = TipoDoc_NDEBITO_A) Then
                    
                    i = 1
                    h = ItemHoja(a, lineas, False, Caracter)
                    While i <= h
                        ImprimirAMRAT (s2n(txtCodigo + i - 1)), True
                        'If MsgBox("Desea imprimir el triplicado?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
                        '    ImprimirAMRAT (s2n(txtCodigo + i - 1)), True
                        'End If
                        i = i + 1
                        If i <= h Then
                            MsgBox "Se mostrara la siguiente factura para imprimir.", , "ATENCION"
                        End If
                    Wend
                ElseIf gEMPR_idEmpresa = 1 Then
                    ImprimirComprobanteLOC (s2n(txtCodigo))
                Else
                    ImprimirComprobante (s2n(txtCodigo)), False, QuieroLeyenda
                End If
                ucBoton.AceptarOk
                logFacturacion "Fin guardado de mas de una hoja", "", txtTipoDoc & " " & TxtNroFactura
            End If
        End If
    Else
        'borro la leyenda guardada.
        str = "delete from facturaventaleyenda where fac=" & s2n(txtCodigo + (ItemHoja(a, lineas, False, Caracter)) - 1)
        DataEnvironment1.Sistema.Execute str
    End If
    logFacturacion "FIN Aceptar", "", txtTipoDoc & " " & TxtNroFactura
fin:
    Exit Sub
UFAalta:
''    DE_RollbackTrans
    ufa "Fallo el alta", ""
    Resume fin
UFAinprime:
    ufa "falla en impresion", ""
    Resume fin
End Sub

Private Function ItemHoja(a As Long, b As Long, item As Boolean, Caracter As Long, Optional hoja As Long) As Long
    'si item es true, devuelve la cantidad de item, en el caso de que sea false, devuelve las hojas
    Dim C As Double
    Dim i As Long
    Dim Linea As Long
    Dim Hojas As Long
    
    Linea = 0
    If item = False Then 'hojas
'        C = a / b
'        If C > CLng(a / b) Then
'            Redondeo = CLng(a / b) + 1
'        Else
'            Redondeo = CLng(a / b)
'        End If
        Hojas = 1
        i = 1
        While i < GRILLA.rows
            Text1.Text = ""
            Text1.Text = GRILLA.TextMatrix(i, 2)
            Linea = Linea + CalculoLineas(Text1, Caracter)
            If Linea > b Then
                Linea = CalculoLineas(Text1, Caracter)
                Hojas = Hojas + 1
            End If
            i = i + 1
        Wend
        ItemHoja = Hojas
    Else 'items
        Hojas = 1
        i = 1
        Do While i < GRILLA.rows
            Text1.Text = ""
            Text1.Text = GRILLA.TextMatrix(i, 2)
            Linea = Linea + CalculoLineas(Text1, Caracter)
            If Linea > b Then
                Linea = CalculoLineas(Text1, Caracter)
                Hojas = Hojas + 1
            End If
            If Hojas > hoja Then
                Exit Do
            End If
            i = i + 1
        Loop
        ItemHoja = i - 1
    End If
End Function
Private Sub ucBoton_BorrarControles()
    Resetear
End Sub

Private Sub ucBoton_Buscar()
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    Dim re As Variant, WhereTipo As String, WhereFecha As String
    Dim CodRef As Long
    're = frmBuscar.mostrarSql("select codigo, NroFactura, Cliente, Fecha  from " & mTablaFV & " where fecha " & ssBetween(dtDesde, dtHasta) & " order by fecha desc ")
    
    
'    WhereTipo = IIf(optBuscarTipo.Item(0).Value, " (TipoDoc = 'FAA' or TipoDoc = 'NCA' or TipoDoc = 'NDA') ", "(TipoDoc = 'FAB')")
    If optBuscarTipo.item(0).Value Then
        WhereTipo = " (TipoDoc = 'FAA' or TipoDoc = 'NCA' or TipoDoc = 'NDA' or tipodoc='FEA' or tipodoc='CEA' or tipodoc='DEA' ) "
    ElseIf optBuscarTipo.item(1).Value Then
        WhereTipo = "(TipoDoc = 'FAB' or TipoDoc = 'NCB' or TipoDoc = 'NDB' or tipodoc='FEB' or tipodoc='CEB' or tipodoc='DEB' ) " ' ahora lo levanta el frm de ND x chq rechaz
    Else
        WhereTipo = " (TipoDoc = 'FAE' or TipoDoc = 'NCE' or TipoDoc = 'NDE') "
    End If
    
    WhereFecha = "fecha " & ucFechas.ssBetween()
    
    With frmBuscar

        re = .MostrarSql("select f.Codigo as Codigo, TipoDoc, NroFactura, Cliente, c.descripcion as [ Nombre                        ], Fecha as [Fecha ], f.activo as Anulada, Remito  from " & mTablaFV & " as f left join clientes as c on c.codigo = f.cliente where " & WhereTipo & " and " & WhereFecha & " order by NroFactura desc ", , , , "", "Anulada", False)
        
        
        If re = "" Then Exit Sub
        txtCodigo = .resultado(1)
        txtTipoDoc = .resultado(2)
        TxtNroFactura = .resultado(3)
        cliente.codigo = .resultado(4)
        dtFecha = .resultado(6)
        TxtRemitoNumero = .resultado(8)
        
        CodRef = s2n(obtenerDeSQL("select codFactura from facturaventa where codigo=" & .resultado(1)))
        txtCodReferencia = CodRef
        txtTipoDocRef = sSinNull(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & CodRef))
        txtNroFacturaRef = s2n(obtenerDeSQL("select nrofactura from facturaventa where codigo=" & CodRef))
        
        mFAE = (Trim(.resultado(2)) = "FAE")
        lblExterior.Visible = mFAE
        
        Dim sPuntoVenta As String, sTipoPunto As String
        sPuntoVenta = sSinNull(obtenerDeSQL("select puntoventa from facturaventa where codigo=" & s2n(txtCodigo)))
        sTipoPunto = sSinNull(obtenerDeSQL("select tipopunto from documentoscae where tipo=" & ssTexto(txtTipoDoc) & " and puntoventa=" & ssTexto(sPuntoVenta)))
        
        txtCotiLeyen.Text = s2n(sSinNull(obtenerDeSQL("select CotizacionLeyenda from facturaventa where codigo=" & s2n(txtCodigo))), 2, True)
        If (obtenerDeSQL("select VaLeyendaCotizacion from facturaventa where codigo=" & s2n(txtCodigo))) = True Then
            ChkVaLeye.Value = 1
        Else
            ChkVaLeye.Value = 0
        End If
        
        If (obtenerDeSQL("select resaltar from facturaventa where codigo=" & s2n(txtCodigo))) = True Then
            chkResaltar.Value = 1
        Else
            chkResaltar.Value = 0
        End If
        
        If sTipoPunto = "OL" Or sTipoPunto = "WS" Then
            cmdAsignarCAE.Visible = True
            lblCAE.Visible = True
            txtFlete.Visible = True
            txtSeguro.Visible = True
        Else
            If siFCAE(.resultado(2), sTipoPunto) Then
                cmdAsignarCAE.Visible = True
                lblCAE.Visible = True
                txtFlete.Visible = True
                txtSeguro.Visible = True
            End If
        End If
    End With
    CargaDatos
    
    gO.Borrar
    ucBoton.BuscarOK
fin:
End Sub
Private Sub ucBoton_Eliminar()
    'esto es para dejar eliminar o no!!
    If Not permiteEliminar Then 'son los permisos ESPECIALES de los usuarios
        MsgBox "No tiene permiso para poder Eliminar."
        Exit Sub
    End If

    If AnularFacturaVenta(s2n(txtCodigo), midDoc) Then
        che "Factura Anulada"
        ucBoton.EliminarOK
    End If
End Sub
Private Sub ucBoton_HabilitarEdicion(sino As Boolean)
    HabilitarEdicion sino
End Sub

Private Sub ucBoton_Imprimir()
    Dim QuieroLeyenda As Boolean
    Dim ley As Boolean
    Dim activo As Boolean


    
    If s2n(txtCodigo) > 0 Then
    
        activo = obtenerDeSQL("select activo from facturaventa where codigo=" & s2n(txtCodigo))
        If activo = False Then
            MsgBox "No se puede imprimir ya que se encuentra anulada."
            Exit Sub
        End If
    
        ley = False
    
        If mFAE Then
            If confirma("Imprime leyenda? ") Then
                QuieroLeyenda = True
                frmLeyendaFAE.Show vbModal
            End If
        Else
        End If
        
        If Right(Trim(txtTipoDoc), 1) = "E" Then
            ImprimirComprobanteFE s2n(txtCodigo)
        Else
            ImprimirComprobanteFE2 s2n(txtCodigo)
        End If

        If VaConRemito Then
            ImprimirRemitoVenta (TxtRemitoNumero)
        End If
    Else
        ufa "No tengo codigo para imprimir", "boton imprimir " & Me.Name ', Err
    End If
    
End Sub

Private Sub ucBoton_Modificar()
    mNuevo = False
End Sub
Private Sub ucBoton_Nuevo()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaNew
    Dim asse As String
    mostrar mFac, mFAE, False
    
    mNuevo = True
    TabDetalle.Tab = 0
    If mFac = FacturaVenta_NCreditoDevolucion Then
        BuscarFacturaReferencia
    End If
    
    If mFac = FacturaVenta_NDebito Then
        BuscarFacturaReferencia 'True
    End If
    
    cmbFormaPago_LostFocus
    cmbingresar.enabled = True
fin:
    Exit Sub
ufaNew:
    ufa "err nuevo() " & "asse ", "nc devol"
    Resume fin
End Sub
Private Sub ucBoton_Salir()
    Unload Me
End Sub
' **********************************************

Private Sub Form_KeyPress(KeyAscii As Integer) ' con Frm.KeyPreView = true
    FrmKeyPress KeyAscii, True, False, True
End Sub

Private Sub Form_Terminate()
    Set g = Nothing
    Set gO = Nothing
    Set cliente = Nothing
End Sub

Private Function GrabaFactura() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim i As Long, bol As Boolean, j As Long
    Dim k As Long
    Dim asse As String ' assert
    Dim cant As Double, prod As String, formu As String, puni, plis As Double, pedi As Long, ptot, remi As Long, item As Long, depot, desc As String, RemitoJunto As Long, alia As String, piva10 As Double, piva21 As Double, PIVA As Double, tmpiva As Double, valorIva As Double
    Dim codTipoDoc, Serie, intBajaStock As Long
    Dim Total As Double, saldo As Double, bContado As Long, bNCxDevol As Long
    Dim iddoc As Long, asieVenta As New Asiento, cuco As String
    Dim TextoAsientoComprobante As String, sPunto As String, sTipo As String
    Dim z As Double ' COTIZACION
    Dim intereses As Double, NroFac As Long, fDecimales
    Dim doc As New FacturaElectronica
    Dim CodFact As Long
    
    fDecimales = 4
    
    z = s2n(txtCotizacion, 4)
    If z = 0 Then z = 1
    
'    If nroaux = 0 Then
'        nroaux = s2n(0.21)
'    End If

    GrabaFactura = False
    logFacturacion "Debo modifcar stock", "", txtTipoDoc & " " & TxtNroFactura
    intBajaStock = IIf(DeboModificarStock(), 1, 0)
    
    logFacturacion "Que deposito", "", txtTipoDoc & " " & TxtNroFactura
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    
    logFacturacion "Numero de remito manual", "", txtTipoDoc & " " & TxtNroFactura
    If VaConRemito() Then RemitoJunto = s2n(TxtRemitoNumero)
    

    logFacturacion "Asiento cabecera", "", txtTipoDoc & " " & TxtNroFactura
    If mFac = FacturaVenta_NCreditoDevolucion Or mFac = FacturaVenta_NCredito Then
        TextoAsientoComprobante = "NC " & TxtNroFactura
        'asieVenta.Nuevo "NC Devolucion  Cliente " & cliente.descripcion, dtfecha, "NCv"
        asieVenta.nuevo "NC " & cliente.DESCRIPCION, dtFecha, "NCv"
    Else
        If mFac = FacturaVenta_NDebito Then
        'N.Debito venta
            TextoAsientoComprobante = "N.Debito venta " & TxtNroFactura
            asieVenta.nuevo "NB " & cliente.DESCRIPCION, dtFecha, "NBv"
        Else
            TextoAsientoComprobante = "FAV " & TxtNroFactura
            asieVenta.nuevo "FV " & cliente.DESCRIPCION, dtFecha, "FAV"
        End If
    End If
    
    logFacturacion "Grabado de referencia si es nota de credito", "", txtTipoDoc & " " & TxtNroFactura
    If mFac = FacturaVenta_NCreditoDevolucion Then
        If s2n(txtCodReferencia.Text) = 0 Then
            'ojo no chequea punto de venta
            CodFact = s2n(obtenerDeSQL("select codigo from facturaventa where tipodoc='" & Trim(txtTipoDocRef) & "' and nrofactura=" & txtNroFacturaRef))
        Else
            CodFact = txtCodReferencia.Text
        End If
        If Trim(txtTipoDoc) = "CEA" Then
            If doc.EstaAprobada(CodFact) Then
                MsgBox "No se puede generar una nota de credito electronica. La Factura esta aprobada explicitamente tras 15 dias de hacerla.", vbCritical
                Exit Function
                
            End If
            If s2n(txttotal) > s2n(doc.SaldoFacturaCredito(CodFact)) Then
                MsgBox "La nota de credito electronica no puede superar el saldo a usar de la factura de credito. Saldo para credito: $ " & s2n(doc.SaldoFacturaCredito(CodFact), , True), vbCritical
                Exit Function
            End If
        End If
    Else
        If s2n(txtCodReferencia.Text) > 0 Then
            'esto es para cargar la referencia del presupuesto a la factura
            CodFact = s2n(txtCodReferencia.Text)
        Else
            CodFact = 0
        End If
    End If
    
    
DE_BeginTrans

    logFacturacion "Total , saldo, contado, es nota,tipo y punto de factura", "", txtTipoDoc & " " & TxtNroFactura
    Total = s2n(txttotal, fDecimales)
    saldo = IIf(EsContado(), 0, Total)
    bContado = optCuentaContado.item(1).Value
    bNCxDevol = (mFac = FacturaVenta_NCreditoDevolucion)
    sTipo = PuntoVentaTipo(cmbPunto.ListIndex)
    sPunto = nSinNull(obtenerDeSQL("select puntoventa from documentoscae where tipopunto=" & ssTexto(sTipo) & " and tipo=" & ssTexto(txtTipoDoc)))
    
    logFacturacion "Nuevo Iddoc", "", txtTipoDoc & " " & TxtNroFactura
    iddoc = NuevoDocumento(txtTipoDoc, TxtNroFactura, 0, 0, 0, 0, sPunto)
    
    logFacturacion "Asiento acumula en IIBB", "", txtTipoDoc & " " & TxtNroFactura
    If bNCxDevol Or mFac = FacturaVenta_NCredito Then
        'asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(txtIva) * z, 0   ', TextoAsientoComprobante
        asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(lblIIBB, fDecimales) * z, 0
    Else
        'Asiento HABER IVA
        'asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva) * z ', TextoAsientoComprobante
        asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(lblIIBB, fDecimales) * z
    End If
    
    
    NroFac = TxtNroFactura.Text
    'obtenerDato("provincias", "'" & cmbProvincia & "'", "codigo")
    asse = "Graba Cabecera: dbo_FV"
    logFacturacion "Graba cabecera", "", txtTipoDoc & " " & TxtNroFactura
    If ABMFacturaVenta("A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura _
        , dtFecha, dtVencimiento, ComboCodigo(cmbformapago), bContado _
        , cliente.codigo, cliente.DESCRIPCION, obtenerDeSQL("select codigo from provincias where descripcion = '" & CmbProvincia & "' ") _
        , ucCuit.Text, ComboCodigo(cmbTipoIva), ComboCodigo(cmbvendedor) _
        , s2n(txtneto) * z, (s2n(txtPIVA) / 100), s2n(txtIva) * z, Total * z, saldo * z, (s2n(txtPdescuento) / 100) * z, 0, RemitoJunto _
        , UsuarioActual(), Date, s2n(txtCotizacion, 4), ComboCodigo(cboMoneda), s2n(lblIIBB), DeboModificarStock(), ComboCodigo(cmbDeposito), bNCxDevol, 0, 0, 0, iddoc, Remito, Orden, s2n(txtCotiLeyen), ChkVaLeye.Value, chkResaltar.Value) = False Then GoTo UfaGraba
    
    logFacturacion "Actualiza cabecera campos vacios porciva", "", txtTipoDoc & " " & TxtNroFactura
    DataEnvironment1.Sistema.Execute "update facturaventa set porciva = " & x2s(nroaux) & " where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura & " and fecha=" & ssFecha(dtFecha) & " and iddoc=" & iddoc
    logFacturacion "Actualiza cabecera campos vacios permisoembarque,incoterms,puntoventa,nrocliente", "", txtTipoDoc & " " & TxtNroFactura
    DataEnvironment1.Sistema.Execute "update facturaventa set permisoembarque=" & ssTexto(txtPermisoEmbarque) & ",incoterms=" & ssTexto(cboIncoterms.Text) & ",puntoventa=" & ssTexto(sPunto) & ",nrocliente=" & ssTexto(txtNroCliente) & " where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura & " and iddoc=" & iddoc
    
    Dim cPaisFAE As String
    cPaisFAE = Trim(sSinNull(obtenerDeSQL("select paiswsfex from clientes where codigo=" & cliente.codigo)))
    logFacturacion "Actualiza cabecera campos vacios paisfae", "", txtTipoDoc & " " & TxtNroFactura
    DataEnvironment1.Sistema.Execute "update facturaventa set paisfae=" & ssTexto(cPaisFAE) & " where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura & " and iddoc=" & iddoc
    
    logFacturacion "Graba detalle", "", txtTipoDoc & " " & TxtNroFactura
    asse = "GrabaDetalle"
    For i = 1 To g.rows - 1
        If g.tx(i, gDESC) = "" Then i = i + 1
        asse = "GrabaDet: calc grilla: prod,form,puni,ptot"
        cant = s2n(g.tx(i, gCANT))
        'If cant = "" Then cant = 0
        prod = VerProductoMio(g.tx(i, gprod), Propio())
        If prod = "" Then prod = "-"
        formu = "" ' VerProductoMio(g.tx(i, gFORM), Propio())
        puni = s2n(g.tx(i, gPUNI))
        'If puni = "" Then puni = 0
        ptot = s2n(g.tx(i, gPTOT))
        'If ptot = "" Then ptot = 0
        'alia = g.tx(i, gALIA)
        
        asse = "GrabaDet: calc grilla: p lis"
        If prod = "-" Then
            plis = 0
        Else
            plis = s2n(obtenerDeSQL("select precio from producto where codigo = '" & VerProductoMio(prod, Propio()) & "'"))
        End If
        asse = "GrabaDet: calc grilla: pedi,remi,item,desc"
        pedi = s2n(g.tx(i, gNPED))
        remi = s2n(g.tx(i, gNREM))
        item = s2n(g.tx(i, gITEM))
        desc = Trim(g.tx(i, gDESC))
        PIVA = s2n(g.tx(i, gIVA))
        
        asse = "GrabaDet: calc grilla: cuentaproducto"
        
        'MODIFICACION INTELIGENTE: POR PRODUCTO!!! aguante cuco
'        cuco = CuentaProducto(prod)
        
        asse = "GrabaDetalle: Graba SP"
        If ABMFVDetalle("A", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, cant, Propio(), prod, desc, formu, puni * z, ptot * z, plis, pedi, remi, item, iddoc, PIVA) = False Then GoTo UfaGraba
        
        asse = "GrabaDetalle: mod stock "
        If DeboModificarStock() And Not EsProductoVirtual(prod) Then
            If mFac = FacturaVenta_NCreditoDevolucion Then          'devolucion,
                DataEnvironment1.dbo_SumaStock prod, cant, depot    'suma stock.
            Else                                                    'factura,
                If Not VaConRemito() Then                           'pero sin remito,
                    If cant >= 0 Then 'esto es por si pone una bonificacion
                        If (puni > 0 And ptot > 0) Then
                            DataEnvironment1.dbo_SumaStock prod, -cant, depot   'resta stock.
                        End If
                    End If
                End If
            End If
        End If
        
''       *************************************** PREFIERO CUCO
''        If bNCxDevol Then
''            'Asiento debe
''            asieVenta.AcumularItem cuco, ptot, 0
''        Else
''            'Asiento HABER
''            asieVenta.AcumularItem cuco, 0, ptot
''        End If
''       *************************************** PREFIERO CUCO
    Next i
    logFacturacion "Fin Graba detalle", "", txtTipoDoc & " " & TxtNroFactura
    'series
    If DeboModificarStock() Then
        For i = 1 To gS.rows - 1 'series
            Serie = gS.tx(i, gS_NSER)
            'consig = (grillaSeries.Cell(flexcpChecked, i, gs_CONS) = flexChecked)
        
            If Serie <> "" Then
                ' nuevoCodigo("series") autonumerico
                codTipoDoc = ObtenerCodigo("TipoComprobantesGrales", txtTipoDoc)
                'DataEnvironment1.dbo_SERIE "A", 0, VerProductoMio(gS.tx(i, gS_PROD), Propio()), serie, codTipoDoc, TxtNroFactura, 0, 0, "", 0, Date, UsuarioActual(), 0, 0
                DataEnvironment1.dbo_abmSERIEs "A", 0, VerProductoMio(gS.tx(i, gS_PROD), Propio()), Serie, codTipoDoc, TxtNroFactura, 0, 0, "", 0, dtFecha, True, Date, UsuarioActual()
            End If
        Next i
    End If
        
    ' Campo descuento lo usan para intereses
    logFacturacion "Asientos intereses", "", txtTipoDoc & " " & TxtNroFactura
    intereses = IIf(s2n(txtDescuento) < 0, Abs(txtDescuento), 0)
    
    'Asie nto
    Dim tiene_c, CUENTA_VENTAS As String, CUENTA_V_VENTAS As String, valor_NETO As Double, aux_prod As Long
    
    logFacturacion "Asientos acumula cuenta del cliente", "", txtTipoDoc & " " & TxtNroFactura
    tiene_c = obtenerDeSQL("select tiene_cuenta from clientes where codigo = " & txtCodCliente)
        If tiene_c = 1 Then
            CUENTA_VENTAS = obtenerDeSQL("select cuenta from clientes where codigo = " & txtCodCliente)
        Else
            'CUENTA_VENTAS = CuentaParam(ID_Cuenta_V_VENTAS)
            CUENTA_VENTAS = CuentaParam(ID_Cuenta_V_DEUDxVENTAS)
        End If
        valor_NETO = s2n(lblSubTotal.caption, fDecimales)
        
        
        piva21 = 0
        piva10 = 0
        logFacturacion "Calculo de valores de iva", "", txtTipoDoc & " " & TxtNroFactura
        With GRILLA
            For i = 1 To .rows - 1
                tmpiva = s2n(0 + (s2n(.TextMatrix(i, gIVA), 4) / 100), 4)
                
                If lblFacturaB.Visible Then
                    valorIva = s2n(s2n(s2n(.TextMatrix(i, gPTOT)) * tmpiva, fDecimales) / (1 + tmpiva), fDecimales)
                Else
                    valorIva = s2n(s2n(.TextMatrix(i, gPTOT)) * tmpiva, 2) 'fDecimales
                End If
                If x2s(.TextMatrix(i, gIVA)) = "21" Then
                    piva21 = piva21 + valorIva
                ElseIf x2s(.TextMatrix(i, gIVA)) = "10.5" Then
                    piva10 = piva10 + valorIva
                End If
            Next
            piva10 = s2n(piva10)
            piva21 = s2n(piva21)
        End With
        logFacturacion "Fin Calculo de valores de iva", "", txtTipoDoc & " " & TxtNroFactura
        Dim rGril As Double
        logFacturacion "Asiento iva segun calculos", "", txtTipoDoc & " " & TxtNroFactura
        If bNCxDevol Or mFac = FacturaVenta_NCredito Then
            If piva21 > 0 Then
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(piva21, fDecimales) * z, 0
            End If
            If piva10 > 0 Then
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS105), s2n(piva10, fDecimales) * z, 0
            End If
        Else
            If piva21 > 0 Then
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(piva21, fDecimales) * z
            End If
            If piva10 > 0 Then
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS105), 0, s2n(piva10, fDecimales) * z
            End If
        End If
    If EsContado() Then
        logFacturacion "Asiento si es contado fin", "", txtTipoDoc & " " & TxtNroFactura
        GrabaValores asieVenta, iddoc
        asieVenta.AgregarItem CUENTA_VENTAS, 0, (s2n(txtneto, fDecimales)) * z - intereses * z  ', TextoAsientoComprobante
        asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, s2n(intereses, fDecimales) * 1
        
    Else
        logFacturacion "Asiento no es contado", "", txtTipoDoc & " " & TxtNroFactura
        If bNCxDevol Or mFac = FacturaVenta_NCredito Then
            logFacturacion "Asiento es nota de credito", "", txtTipoDoc & " " & TxtNroFactura
            If mFAE Then
                logFacturacion "Asiento es exportacion", "", txtTipoDoc & " " & TxtNroFactura
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), 0, s2n(Total, fDecimales) * z, TextoAsientoComprobante
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS_EXT), (s2n(txtneto, fDecimales)) * z, 0   ', TextoAsientoComprobante
            Else
                logFacturacion "Asiento no es exportacion", "", txtTipoDoc & " " & TxtNroFactura
                If mFac = FacturaVenta_NCredito Or mFac = FacturaVenta_NCreditoDevolucion Then
                    logFacturacion "Asiento es nota de credito", "", txtTipoDoc & " " & TxtNroFactura
                    If lblFacturaB.Visible = True Then
                        'asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n((piva21 + piva10) * z, fDecimales), 0
                        logFacturacion "Asiento es de tipo B calculo neto", "", txtTipoDoc & " " & TxtNroFactura
                        valor_NETO = s2n(valor_NETO - piva21 - piva10, fDecimales)
                    End If
                    
                    logFacturacion "Asiento cuenta ventas", "", txtTipoDoc & " " & TxtNroFactura
                    asieVenta.AcumularItem CUENTA_VENTAS, 0, s2n(Total, fDecimales) * z, TextoAsientoComprobante
                    
                    logFacturacion "Asiento cuenta por productos", "", txtTipoDoc & " " & TxtNroFactura
                    For aux_prod = 1 To GRILLA.rows - 1
                        CUENTA_V_VENTAS = ""
                        If GRILLA.TextMatrix(aux_prod, gCTA) > "" Then
                            CUENTA_V_VENTAS = GRILLA.TextMatrix(aux_prod, gCTA)
                            'valor_NETO = valor_NETO - grilla.TextMatrix(aux_prod, 4)
                            'asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(grilla.TextMatrix(aux_prod, 4)) * z - intereses * z            ', TextoAsientoComprobante
                            If lblFacturaB.Visible Then
                                rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100, 4))))
                                valor_NETO = valor_NETO - rGril
                            Else
                                rGril = GRILLA.TextMatrix(aux_prod, 4)
                                valor_NETO = valor_NETO - rGril
                            End If
                            asieVenta.AgregarItem CUENTA_V_VENTAS, s2n(rGril, fDecimales) * z - intereses * z, 0
                                                
                        Else
                            tiene_c = obtenerDeSQL("select tiene_cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                            If tiene_c = 1 Then
                                CUENTA_V_VENTAS = obtenerDeSQL("select cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                                
                                If lblFacturaB.Visible Then
                                    'rGril = s2n(Grilla.TextMatrix(aux_prod, gIVA) / 100)
                                    rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100, 4))))
                                    valor_NETO = valor_NETO - rGril
                                Else
                                    rGril = GRILLA.TextMatrix(aux_prod, 4)
                                    valor_NETO = valor_NETO - rGril
                                End If
                                asieVenta.AgregarItem CUENTA_V_VENTAS, s2n(rGril) * z - intereses * z, 0
                            End If
                        End If
                    Next
                    logFacturacion "Asiento fin cuenta de productos", "", txtTipoDoc & " " & TxtNroFactura
                    If s2n(valor_NETO) > 0 Then
                        logFacturacion "Asiento quedo algo de neto acumula en contra cuenta de ventas", "", txtTipoDoc & " " & TxtNroFactura
                        asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_VENTAS), s2n(valor_NETO, fDecimales) * z - intereses * z, 0
                    End If
                    logFacturacion "Asiento descuento e intereses", "", txtTipoDoc & " " & TxtNroFactura
                    asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_DESCUENTO), -(s2n(txtDescuento, fDecimales)) * z - intereses * z, 0, TextoAsientoComprobante
                    asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), s2n(intereses * z, fDecimales), 0
                End If
            End If
        Else
            logFacturacion "Asiento es factura", "", txtTipoDoc & " " & TxtNroFactura
            If mFAE Then
                logFacturacion "Asiento es exportacion fin", "", txtTipoDoc & " " & TxtNroFactura
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), s2n(Total * z, fDecimales), 0, TextoAsientoComprobante
                asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_VENTAS_EXT), 0, (s2n(txtneto, fDecimales)) * z - intereses * z   ', TextoAsientoComprobante
                asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, s2n(intereses * z, fDecimales)
            Else
                logFacturacion "Asiento no es exportacion", "", txtTipoDoc & " " & TxtNroFactura
                If lblFacturaB.Visible = True Then
                    logFacturacion "Asiento es de tipo B calcula neto", "", txtTipoDoc & " " & TxtNroFactura
                    'asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n((piva21 + piva10) * z, fDecimales)
                    valor_NETO = s2n(valor_NETO - piva21 - piva10)
                End If
                
                logFacturacion "Asiento acumula cuenta ventas", "", txtTipoDoc & " " & TxtNroFactura
                asieVenta.AcumularItem CUENTA_VENTAS, s2n(Total * z, fDecimales), 0, TextoAsientoComprobante
                
                logFacturacion "Asiento cuenta de productos", "", txtTipoDoc & " " & TxtNroFactura
                For aux_prod = 1 To GRILLA.rows - 1
                    CUENTA_V_VENTAS = ""
                    If GRILLA.TextMatrix(aux_prod, gCTA) > "" Then
                        CUENTA_V_VENTAS = GRILLA.TextMatrix(aux_prod, gCTA)
                        'valor_NETO = valor_NETO - grilla.TextMatrix(aux_prod, 4)
                        'asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(grilla.TextMatrix(aux_prod, 4)) * z - intereses * z            ', TextoAsientoComprobante
                        If lblFacturaB.Visible Then
                            rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100, 4))))
                            valor_NETO = valor_NETO - rGril
                        Else
                            rGril = GRILLA.TextMatrix(aux_prod, 4)
                            valor_NETO = valor_NETO - rGril
                        End If
                        asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(rGril, fDecimales) * z - intereses * z           ', TextoAsientoComprobante
                                            
                    Else
                        tiene_c = obtenerDeSQL("select tiene_cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                        If tiene_c = 1 Then
                            CUENTA_V_VENTAS = obtenerDeSQL("select cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                            
                            If lblFacturaB.Visible Then
                                rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100, 4))))
                                valor_NETO = valor_NETO - rGril
                            Else
                                rGril = GRILLA.TextMatrix(aux_prod, 4)
                                valor_NETO = valor_NETO - rGril
                            End If
                            asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(rGril, fDecimales) * z - intereses * z           ', TextoAsientoComprobante
                        End If
                    End If
                Next
                logFacturacion "Asiento fin cuenta productos", "", txtTipoDoc & " " & TxtNroFactura
                If s2n(valor_NETO) > 0 Then
                    logFacturacion "Asiento si sobra neto acumula en contra cuenta de ventas", "", txtTipoDoc & " " & TxtNroFactura
                    asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_VENTAS), 0, s2n(valor_NETO, fDecimales) * z - intereses * z         ', TextoAsientoComprobante
                End If
                logFacturacion "Asiento descuento e intereses", "", txtTipoDoc & " " & TxtNroFactura
                asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_DESCUENTO), 0, -(s2n(txtDescuento, fDecimales)) * z - intereses * z, TextoAsientoComprobante
                asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, s2n(intereses * z, fDecimales)
                
            End If
        End If
    End If
    
    logFacturacion "Asiento validacion parametro si graba", "", txtTipoDoc & " " & TxtNroFactura
    If siAsiento("AsientosVentas") Then
        logFacturacion "Asiento si graba", "", txtTipoDoc & " " & TxtNroFactura
        If asieVenta.Grabar(iddoc) = 0 Then
            logFacturacion "Asiento fallo al grabar", "", txtTipoDoc & " " & TxtNroFactura
            If MsgBox("Desea continuar???...", vbInformation + vbYesNo) = vbNo Then
                logFacturacion "Asiento grabar no continua, rollback", "", txtTipoDoc & " " & TxtNroFactura
                DE_RollbackTrans
        '        ufa "Err al grabar asiento ", Me.Name & " - " '& sAssert
                Exit Function
            End If
            logFacturacion "Asiento continua igual", "", txtTipoDoc & " " & TxtNroFactura
        End If
        
    End If
    logFacturacion "Asiento fin grabar", "", txtTipoDoc & " " & TxtNroFactura
    logFacturacion "si hay remito actualizo factura", "", txtTipoDoc & " " & TxtNroFactura
        If remi > 0 Then
            DataEnvironment1.Sistema.Execute "update remitoventa set factura = " & NroFac & " where numero=" & remi
            
            DataEnvironment1.Sistema.Execute "update facturaventa set remito = " & remi & " where nrofactura=" & NroFac
        End If
    
    
    logFacturacion "si hay referencia actualizo factura", "", txtTipoDoc & " " & TxtNroFactura
        If s2n(txtCodReferencia) > 0 Then
            'DEJO SOLO UPDATE SOBRE COMPROBANTE GUARDADO ASI QUEDA RELACIONES TOTALES, NC>FACTURA Y ND>FACTURA Y FACTURA SIN RELACION
            'DataEnvironment1.Sistema.Execute "update facturaventa set codfactura=" & s2n(txtCodigo) & "  where codigo=" & s2n(txtCodReferencia)
            DataEnvironment1.Sistema.Execute "update facturaventa set codfactura=" & s2n(txtCodReferencia) & "  where codigo=" & s2n(txtCodigo)
        End If

    
DE_CommitTrans
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura = True
    logFacturacion "FIN GRABAR FACTURA", "", txtTipoDoc & " " & TxtNroFactura
fin:
    Exit Function
    
UfaGraba:
    DE_RollbackTrans
    logFacturacion "ERROR GRABAR FACTURA", "", txtTipoDoc & " " & TxtNroFactura
    ufa "Err al grabar ", " grabaFV() - " & asse & " " & prod ', Err
    'Resume fin
End Function

Private Function GrabaFactura2(CantFac As Long, Optional sPunto As String = "0001") As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim i As Long, bol As Boolean, j As Long
    Dim k As Long
    Dim m As Long
    Dim asse As String ' assert
    Dim cant As Double, prod As String, formu As String, puni, plis As Double, pedi As Long, ptot, remi As Long, item As Long, depot, desc As String, RemitoJunto As Long, alia As String, piva10 As Double, piva21 As Double, PIVA As Double, tmpiva As Double, valorIva As Double
    Dim codTipoDoc, Serie, intBajaStock As Long
    Dim Total As Double, saldo As Double, bContado As Long, bNCxDevol As Long
    Dim iddoc As Long, asieVenta As New Asiento, cuco As String
    Dim TextoAsientoComprobante As String
    Dim z As Double ' COTIZACION
    Dim intereses As Double
    Dim NroFac As Long
    Dim cantItem As Long
    Dim n As Long
    Dim p As Long
    Dim Nro As String
    
    
    z = s2n(txtCotizacion, 4)
    If z = 0 Then z = 1

    GrabaFactura2 = False
    intBajaStock = IIf(DeboModificarStock(), 1, 0)
    
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    
    If VaConRemito() Then RemitoJunto = s2n(TxtRemitoNumero)
    m = 1
    n = 1
    p = 1
    i = 1
    While i <= CantFac
        If Nro <> "" Then Nro = Trim(Nro) & ","
        Nro = Trim(Nro) & "#" & Trim(TxtNroFactura + i - 1) & "#"
        i = i + 1
    Wend
    
'*** una transaccion aqui ..... *********************
    
DE_BeginTrans
    
    While m <= CantFac
    
        If mFac = FacturaVenta_NCreditoDevolucion Then
            TextoAsientoComprobante = "NC " & TxtNroFactura + m - 1
            asieVenta.nuevo "NC " & cliente.DESCRIPCION, dtFecha, "NCv"
        Else
            TextoAsientoComprobante = "FAV " & TxtNroFactura + m - 1
            asieVenta.nuevo "FV " & cliente.DESCRIPCION, dtFecha, "FAV"
        End If
        
    
        asse = "Graba Cabecera"
        Total = IIf((m = CantFac), s2n(txttotal, 2), 0)
        saldo = IIf((m = CantFac), IIf(EsContado(), 0, Total), 0)
        bContado = optCuentaContado.item(1).Value
        bNCxDevol = (mFac = FacturaVenta_NCreditoDevolucion)
        
        iddoc = NuevoDocumento(txtTipoDoc, TxtNroFactura + m - 1, 0, 0)
        
        If m = CantFac Then
            If bNCxDevol Then
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(txtIva) * z, 0 ', TextoAsientoComprobante
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(lblIIBB) * z, 0
            Else
                'Asiento HABER IVA
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(txtIva) * z ', TextoAsientoComprobante
                asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(lblIIBB) * z
            End If
        End If
        
        
        NroFac = TxtNroFactura.Text + m - 1
        asse = "Graba Cabecera: dbo_FV"
        If m = CantFac Then
            If ABMFacturaVenta("A", s2n(txtCodigo + m - 1), txtTipoDoc, TxtNroFactura + m - 1 _
                , dtFecha, dtVencimiento, ComboCodigo(cmbformapago), bContado _
                , cliente.codigo, cliente.DESCRIPCION, obtenerDeSQL("select codigo from provincias where descripcion = '" & CmbProvincia & "' ") _
                , ucCuit.Text, ComboCodigo(cmbTipoIva), ComboCodigo(cmbvendedor) _
                , s2n(txtneto) * z, (s2n(txtPIVA) / 100), s2n(txtIva) * z, Total * z, saldo * z, (s2n(txtPdescuento) / 100) * z, 0, RemitoJunto _
                , UsuarioActual(), Date, s2n(txtCotizacion, 4), ComboCodigo(cboMoneda), s2n(lblIIBB), DeboModificarStock(), ComboCodigo(cmbDeposito), bNCxDevol, 0, 0, 0, iddoc, Remito, Orden) = False Then GoTo UfaGraba
            DataEnvironment1.Sistema.Execute "update facturaventa set variasfac='" & Trim(Nro) & "',puntoventa='" & Trim(sPunto) & "' where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & Trim(TxtNroFactura + m - 1)
        Else
            If ABMFacturaVenta("A", s2n(txtCodigo + m - 1), txtTipoDoc, TxtNroFactura + m - 1 _
                , dtFecha, dtVencimiento, ComboCodigo(cmbformapago), bContado _
                , cliente.codigo, cliente.DESCRIPCION, obtenerDeSQL("select codigo from provincias where descripcion = '" & CmbProvincia & "' ") _
                , ucCuit.Text, ComboCodigo(cmbTipoIva), ComboCodigo(cmbvendedor) _
                , 0, (s2n(txtPIVA) / 100), 0, 0, 0, 0, 0, RemitoJunto _
                , UsuarioActual(), Date, s2n(txtCotizacion, 4), ComboCodigo(cboMoneda), 0, DeboModificarStock(), ComboCodigo(cmbDeposito), bNCxDevol, 0, 0, 0, iddoc, Remito, Orden) = False Then GoTo UfaGraba
            DataEnvironment1.Sistema.Execute "update facturaventa set variasfac='" & Trim(Nro) & "',puntoventa='" & Trim(sPunto) & "' where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & Trim(TxtNroFactura + m - 1)
        End If
        
        DataEnvironment1.Sistema.Execute "update facturaventa set porciva = " & x2s(nroaux) & " where tipodoc='" & Trim(txtTipoDoc) & "' and nrofactura=" & TxtNroFactura + m - 1 & " and fecha=" & ssFecha(dtFecha)
        
        asse = "GrabaDetalle"
        cantItem = ItemHoja(0, 22, True, 75, m) 'antes era 25 lineas de item ahora 22
        For i = n To cantItem 'g.rows - 1
            If g.tx(i, gDESC) = "" Then i = i + 1
            asse = "GrabaDet: calc grilla: prod,form,puni,ptot"
            cant = s2n(g.tx(i, gCANT))
            prod = VerProductoMio(g.tx(i, gprod), Propio())
            If prod = "" Then prod = "-"
            formu = ""
            puni = s2n(g.tx(i, gPUNI))
            ptot = s2n(g.tx(i, gPTOT))
            asse = "GrabaDet: calc grilla: p lis"
            If prod = "-" Then
                plis = 0
            Else
                plis = s2n(obtenerDeSQL("select precio from producto where codigo = '" & VerProductoMio(prod, Propio()) & "'"))
            End If
            asse = "GrabaDet: calc grilla: pedi,remi,item,desc"
            pedi = s2n(g.tx(i, gNPED))
            remi = s2n(g.tx(i, gNREM))
            item = s2n(g.tx(i, gITEM))
            desc = Trim(g.tx(i, gDESC))
            PIVA = s2n(g.tx(i, gIVA))
            
            asse = "GrabaDet: calc grilla: cuentaproducto"
            
            'MODIFICACION INTELIGENTE: POR PRODUCTO!!! aguante cuco
    '        cuco = CuentaProducto(prod)
            
            asse = "GrabaDetalle: Graba SP"
            If ABMFVDetalle("A", s2n(txtCodigo + m - 1), txtTipoDoc, TxtNroFactura + m - 1, cant, Propio(), prod, desc, formu, puni * z, ptot * z, plis, pedi, remi, item, iddoc, PIVA) = False Then GoTo UfaGraba
            
            asse = "GrabaDetalle: mod stock "
            If DeboModificarStock() And Not EsProductoVirtual(prod) Then
                If mFac = FacturaVenta_NCreditoDevolucion Then          'devolucion,
                    DataEnvironment1.dbo_SumaStock prod, cant, depot    'suma stock.
                Else                                                    'factura,
                    If Not VaConRemito() Then                           'pero sin remito,
                        If cant >= 0 Then 'esto es por si pone una bonificacion
                            If (puni > 0 And ptot > 0) Then
                                DataEnvironment1.dbo_SumaStock prod, -cant, depot   'resta stock.
                            End If
                        End If
                    End If
                End If
            End If
            
        Next i
        n = i
        
        'series
        If m = CantFac Then
            If DeboModificarStock() Then
                For i = 1 To gS.rows - 1 'series
                    Serie = gS.tx(i, gS_NSER)
                
                    If Serie <> "" Then
                        codTipoDoc = ObtenerCodigo("TipoComprobantesGrales", txtTipoDoc)
                        DataEnvironment1.dbo_abmSERIEs "A", 0, VerProductoMio(gS.tx(i, gS_PROD), Propio()), Serie, codTipoDoc, TxtNroFactura + m - 1, 0, 0, "", 0, dtFecha, True, Date, UsuarioActual()
                    End If
                Next i
            End If
        End If
            
        ' Campo descuento lo usan para intereses
        intereses = IIf(s2n(txtDescuento) < 0, Abs(txtDescuento), 0)
        
        'Asie nto
        Dim tiene_c, CUENTA_VENTAS As String, CUENTA_V_VENTAS As String, valor_NETO As Double, aux_prod As Long
        
        tiene_c = obtenerDeSQL("select tiene_cuenta from clientes where codigo = " & txtCodCliente)
            If tiene_c = 1 Then
                CUENTA_VENTAS = obtenerDeSQL("select cuenta from clientes where codigo = " & txtCodCliente)
            Else
                CUENTA_VENTAS = CuentaParam(ID_Cuenta_V_DEUDxVENTAS)
            End If
            valor_NETO = IIf((m = CantFac), s2n(lblSubTotal.caption), 0)
            
            
            piva21 = 0
            piva10 = 0
            With GRILLA
                For i = p To cantItem '.rows - 1
                    tmpiva = s2n(0 + (s2n(.TextMatrix(i, gIVA), 4) / 100), 4)
                    valorIva = s2n(s2n(s2n(.TextMatrix(i, gPTOT)) * tmpiva, 4) / (1 + tmpiva), 4)
                    If .TextMatrix(i, gIVA) = "21" Then
                        piva21 = piva21 + valorIva
                    ElseIf .TextMatrix(i, gIVA) = "10.5" Then
                        piva10 = piva10 + valorIva
                    End If
                Next
            End With
            
            
        If EsContado() Then
            GrabaValores asieVenta, iddoc
            asieVenta.AgregarItem CUENTA_VENTAS, 0, (s2n(txtneto)) * z - intereses * z   ', TextoAsientoComprobante
            asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, intereses * 1
            
        Else
            If bNCxDevol Then
                If mFAE Then
                    asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), 0, Total * z, TextoAsientoComprobante
                    asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS_EXT), (s2n(txtneto)) * z, 0    ', TextoAsientoComprobante
                Else
                    asieVenta.AcumularItem CUENTA_VENTAS, 0, Total * z, TextoAsientoComprobante
                    
                    asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), (s2n(txtneto)) * z, 0  ', TextoAsientoComprobante
                End If
            Else
                If mFAE Then
                    asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS_EXT), Total * z, 0, TextoAsientoComprobante
                    asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_VENTAS_EXT), 0, (s2n(txtneto)) * z - intereses * z    ', TextoAsientoComprobante
                    asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, intereses * 1
                Else
                    If m = CantFac Then
                        If lblFacturaB.Visible = True Then
                            asieVenta.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n((piva21 + piva10) * z)
                            valor_NETO = s2n(valor_NETO - piva21 - piva10)
                        End If
                        
                        asieVenta.AcumularItem CUENTA_VENTAS, Total * z, 0, TextoAsientoComprobante
                        Dim rGril As Double
                        For aux_prod = 1 To GRILLA.rows - 1
                            CUENTA_V_VENTAS = ""
                            If GRILLA.TextMatrix(aux_prod, gCTA) > "" Then
                                CUENTA_V_VENTAS = GRILLA.TextMatrix(aux_prod, gCTA)
                                If lblFacturaB.Visible Then
                                    rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100))))
                                    valor_NETO = valor_NETO - rGril
                                Else
                                    rGril = GRILLA.TextMatrix(aux_prod, 4)
                                    valor_NETO = valor_NETO - rGril
                                End If
                                asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(rGril) * z - intereses * z            ', TextoAsientoComprobante
                            
                                'CUENTA_V_VENTAS = grilla.TextMatrix(aux_prod, gCTA)
                                'valor_NETO = valor_NETO - grilla.TextMatrix(aux_prod, 4)
                                'asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(grilla.TextMatrix(aux_prod, 4)) * z - intereses * z            ', TextoAsientoComprobante
                            Else
                                tiene_c = obtenerDeSQL("select tiene_cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                                If tiene_c = 1 Then
                                    CUENTA_V_VENTAS = obtenerDeSQL("select cuenta from producto where codigo = '" & GRILLA.TextMatrix(aux_prod, 1) & "'")
                                    If lblFacturaB.Visible Then
                                        rGril = s2n(GRILLA.TextMatrix(aux_prod, 4) / (1 + (s2n(GRILLA.TextMatrix(aux_prod, gIVA) / 100))))
                                        valor_NETO = valor_NETO - rGril
                                    Else
                                        rGril = GRILLA.TextMatrix(aux_prod, 4)
                                        valor_NETO = valor_NETO - rGril
                                    End If
                                    asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(rGril) * z - intereses * z            ', TextoAsientoComprobante
                                
                                    'CUENTA_V_VENTAS = obtenerDeSQL("select cuenta from producto where codigo = '" & grilla.TextMatrix(aux_prod, 1) & "'")
                                    'valor_NETO = valor_NETO - grilla.TextMatrix(aux_prod, 4)
                                    'asieVenta.AgregarItem CUENTA_V_VENTAS, 0, s2n(grilla.TextMatrix(aux_prod, 4)) * z - intereses * z            ', TextoAsientoComprobante
                                End If
                            End If
                        Next
                        If s2n(valor_NETO) > 0 Then
                            asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_VENTAS), 0, (valor_NETO) * z - intereses * z          ', TextoAsientoComprobante
                        End If
                        asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_DESCUENTO), 0, -(s2n(txtDescuento)) * z - intereses * z, TextoAsientoComprobante
                        asieVenta.AgregarItem CuentaParam(ID_Cuenta_V_INTERESES), 0, intereses * 1
                        
                    End If
                    
                End If
            End If
        End If
        
        If siAsiento("AsientosVentas") Then asieVenta.Grabar iddoc
            If remi > 0 Then
                DataEnvironment1.Sistema.Execute "update remitoventa set factura = " & NroFac & " where numero=" & remi
                
                DataEnvironment1.Sistema.Execute "update facturaventa set remito = " & remi & " where nrofactura=" & NroFac
            End If
            
        m = m + 1
        p = n
    Wend

    
DE_CommitTrans
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura2 = True
    
fin:
    Exit Function
    
UfaGraba:
    DE_RollbackTrans
    ufa "Err al grabar ", " grabaFV() - " & asse & " " & prod ', Err
    'Resume fin
End Function

'Private Function CuentaProducto(pro As String) As String
'    Dim tempo
'    tempo = sSinNull(obtenerDeSQL("select cuenta from producto where codigo = '" & pro & "' "))
'    'If tempo = "" Then tempo = CuentaParam(ID_CuentasParam_PROD_GENERICO)
'    CuentaProducto = tempo
'End Function


Private Sub RevisarTotales()
    Dim subtotal As Double, descu As Double, Neto As Double
    Dim subtot2 As Double
    Dim dTipo As String
    dTipo = UCase(txtTipoDoc)
    
'esta linea se comento solo para amr ya que ellos laburan con netos y no con totales
'pero en el caso de usar solo totales descomentar y ya esta.
    If InStr(dTipo, "B") Then txtPIVA = 0
    
    If gEMPR_idEmpresa = 2 Then
        subtotal = s2n(g.suma(gPTOT), 4)
    Else
        subtotal = s2n(g.suma(gPTOT), 2)
    End If
    subtotal = s2n(g.suma(gPTOT), 2)
    subtot2 = subtotal + s2n(txtFlete) + s2n(txtSeguro)
    subtotal = subtot2
    
    descu = s2n(subtotal, 4) * (s2n(txtPdescuento, 4) / 100)
    Neto = s2n(subtotal - descu, 4)
    
    lblSubTotal = s2n(subtotal, 4)
    txtDescuento = s2n(descu, 2)
    txtneto = Neto
    'ACA ME QUEDE
    
    lblIIBB = CalcPercIIBB(Neto, cliente.codigo)
    
    
    txtIva = Calciva 's2n(Neto * (s2n(txtPIVA, 4) / 100), 2)
    
    txttotal = s2n(Neto + s2n(txtIva, 4), 2) + s2n(lblIIBB)
    
'    uTipoVenta.Total_a_Imputar = s2n(s2n(txtneto, 4) - s2n(txtdescuento, 4), 2)
End Sub

Private Function Calciva() As Double
    Dim i As Long, Acumula As Double, qiva As Double
    Dim dif As Boolean
    Dim Iva As Double
    
    Calciva = 0
    If lblFacturaB.Visible Then Exit Function
    Acumula = 0
    dif = False
    Iva = 0
    With g
        For i = 1 To .rows - 1
            If s2n(.TextMatrix(i, gIVA)) <> 0 Then
                If dif = False And Iva > 0 And Iva <> s2n(.TextMatrix(i, gIVA)) Then dif = True
                Iva = s2n(.TextMatrix(i, gIVA))
            End If
            qiva = s2n(s2n(.TextMatrix(i, gPTOT)) * s2n(.TextMatrix(i, gIVA)) / 100)
            Acumula = Acumula + qiva
        Next
        If dif = True Then txtPIVA = 0
    End With
    Calciva = Acumula
End Function

Private Sub verCampos(verOrigen As Boolean, habChk As Boolean, habFraEditDet As Boolean, gTop As Long, gHe As Long, verDep As Boolean, capt As String, verchkBajaStk As Boolean, chkBajaStkValue, verFacturaReferencia As Boolean, verBotonRemito As Boolean, verBotonPedido As Boolean)
    TabDetalle.TabVisible(1) = verOrigen
    TabDetalle.TabVisible(2) = verOrigen 'series
    
    'chkPropio.Enabled = habChk
    chkPropio.enabled = True  ' creo que falla con migrados
    
    fraEditDetalle.Visible = habFraEditDet
    GRILLA.Top = gTop
    GRILLA.Height = gHe
    lblDepot.Visible = verDep
    cmbDeposito.Visible = verDep
    Me.caption = capt
    'chkModStock.Visible = verchkBajaStk
    'chkModStock.Value = chkBajaStkValue
    
    fraOptStock.Visible = verchkBajaStk
    
    optStock.item(chkBajaStkValue).Value = True
    
    lblref.Visible = verFacturaReferencia
    txtTipoDocRef.Visible = verFacturaReferencia
    txtNroFacturaRef.Visible = verFacturaReferencia
    cmdRemitosPendientes.Visible = verBotonRemito
    cmdPedidosPendientes.Visible = verBotonPedido
End Sub


Private Sub BuscarFacturaReferencia(Optional SoloFacturasCredito As Boolean = False)
    If ON_ERROR_HABILITADO Then On Error GoTo UfaErrFRef
    Dim s As String, re As String, aRe As Variant, i As Long
    Dim rs As New ADODB.Recordset
    Dim asse As String, z As Double
    Dim tempo
    's = "select FacturaVenta.codigo, TipoDoc, NroFactura, Fecha, descripcion from FacturaVenta inner join Clientes on FacturaVenta.Cliente = Clientes.Codigo where FacturaVenta.activo = 1 and  (tipodoc = '" & TipoDoc_FACTURA_A & "' or tipoDoc = '" & TipoDoc_FACTURA_B & "') and fecha  " & ssBetween(dtDesde, dtHasta) & " order by FacturaVenta.codigo desc "
    If SoloFacturasCredito Then
        s = "select FacturaVenta.codigo, TipoDoc, NroFactura, Fecha,Clientes.Codigo as Cod, Descripcion, contado,moneda,cotizacion from FacturaVenta inner join Clientes on FacturaVenta.Cliente = Clientes.Codigo where FacturaVenta.activo = 1 and  (TIPODOC='" & TipoDoc_FACTURA_CREDITO_A & "' OR TIPODOC='" & TipoDoc_FACTURA_CREDITO_B & "' OR TIPODOC='" & TipoDoc_FACTURA_CREDITO_C & "' ) and fecha  " & ucFechas.ssBetween() & " order by FacturaVenta.codigo desc "
    Else
        s = "select FacturaVenta.codigo, TipoDoc, NroFactura, Fecha,Clientes.Codigo as Cod, Descripcion, contado,moneda,cotizacion from FacturaVenta inner join Clientes on FacturaVenta.Cliente = Clientes.Codigo where FacturaVenta.activo = 1 and  (tipodoc = '" & TipoDoc_FACTURA_A & "' or tipoDoc = '" & TipoDoc_FACTURA_B & "'  OR TIPODOC='" & TipoDoc_FACTURA_CREDITO_A & "' OR TIPODOC='" & TipoDoc_FACTURA_CREDITO_B & "' OR TIPODOC='" & TipoDoc_FACTURA_CREDITO_C & "' ) and fecha  " & ucFechas.ssBetween() & " order by FacturaVenta.codigo desc "
    End If
    
    With frmBuscar
        asse = "por buscar"
        re = .MostrarSql(s, , "Credito Sobre Factura:", , StringCONTADO, "Cta Cte")
        If re <> "" Then
            txtCodReferencia = .resultado(1)
            txtTipoDocRef = .resultado(2)
            asse = "nro ref"
            txtNroFacturaRef = .resultado(3)
            asse = "clie Desc"
            cliente.codigo = .resultado(5)            ' datos cliente cambian solos
            asse = "optCont"
            optCuentaContado.item(1).Value = (.resultado(7) = StringCONTADO)
            tempo = obtenerDeSQL("select iva, porcentajeiva, descuento from facturaventa where codigo = " & .resultado(1))
            txtIva = s2n(tempo(0))
            txtPIVA = s2n(tempo(1), 4) * 100
            txtPdescuento = s2n(tempo(2), 4) * 100
            cboMoneda.ListIndex = BuscarEnCombo(cboMoneda, .resultado(8))
            z = s2n(.resultado(9))
            If z = 0 Then z = 1
            txtCotizacion = z
            
            'cargar items
''            aRe = obtenerDeSQL("select Producto, Cantidad, codPropio, precioUnitario, descripcion from FacturaVentaDetalle where CodigoFactura = " & re)
            
            With rs
                .Open "select *,_iva as iva from FacturaVentaDetalle where CodigoFactura = " & re, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                g.Borrar
                While Not .EOF
                    asse = "it propio"
                    chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)

                   ' MetoEnGrilla aRe(0), aRe(4), aRe(1), aRe(3), 0, 0, 0, , False
                    asse = " it grilla " & !producto
                    MetoEnGrilla VerProductoCliente(!producto, Propio(), cliente.codigo), !DESCRIPCION, !cantidad, !PrecioUnitario / z, 0, 0, 0, s2n(!Iva), False, False
                    .MoveNext
                Wend
            End With
        End If
    End With
fin:
    Set rs = Nothing
    Exit Sub
UfaErrFRef:
    ufa "Err al buscar factura referencia", re & " - BuscRef:" & asse
    Resume fin
End Sub

Private Function EsContado()
    'EsContado = (ComboCodigo(cmbFormaPago) = FormaPago_CONTADO) And (Not mFac = FacturaVenta_NCreditoDevolucion)
    EsContado = (optCuentaContado.item(1).Value) And (Not mFac = FacturaVenta_NCreditoDevolucion)
End Function


Private Sub GrabaValores(asien As Asiento, iddoc As Long)
    Dim nMoviC, tdoc, nDoc, chCta, chCod, chMonto As Double
    Dim DetConcepto, MoviConcepto, i As Long
    Dim z As Double
    
'            Dim CantCheques, i As Long,
'            Dim chNro,  chCod, chFech As Date, chPT,  bco
'
'            '    lblCodigo = obtenerParametro(CAMPO_BS_CodFactura_VENTA) + 1
'            CantCheques = g.PrimerVacio(gMONTO) - 1
'           'MoviConcepto = "Recibo 00000973  Cliente 2028   NASTIQUE OSCAR ALF"
    
    MoviConcepto = "Fact   " & Format(TxtNroFactura, "00000000") & "Cliente " & Format(cliente.codigo, "0000") & "   " & cliente.DESCRIPCION
    DetConcepto = "Fact   " & Format(TxtNroFactura, "00000000") & "  Cliente " & Format(cliente.codigo, "0000")
    tdoc = txtTipoDoc ' " FAC "
    nDoc = s2n(TxtNroFactura)
    z = s2n(txtCotizacion, 4)
    If z = 0 Then z = 1

'
'            ' --- Transaccion aqui -----------  ini
'
    With frmValores
            'Efectivo
'            If s2n(txtEfectivo) > 0 Then
        If .Efectivo > 0 Then
            'MoviCaja
            
            nMoviC = nuevoCodigo("movicaja", "movimiento")
            DataEnvironment1.dbo_MOVICAJASdoc "A", .caja, nMoviC, cliente.codigo, MoviCaja_EFECTIVO, MoviCaja_INGRESO, .Efectivo, MoviConcepto, Date, .Cuenta, 0, z, tdoc, nDoc, Date, UsuarioActual(), 0, iddoc
'''''            'detMovCaja
'''''            DataEnvironment1.dbo_DETMOVCAJAS "A", nMoviC, .Efectivo, cliente.codigo, .cuenta, DetConcepto, "FAC" 'tdoc?
            
            'ASIENTO Efectivo en el HABER
            asien.AcumularItem .Cuenta, .Efectivo, 0
        End If
     
        'CHEQUES
        For i = 1 To .ChCant
'                chNro = g.tx(i, gNROCH)
            chMonto = .chMonto(i)
            chCta = CuentaParam(ID_Cuenta_M_CH_CARTERA) 'obtenerParametro("Cta_Caja")
'                chFech = CDate(g.tx(i, gFECHA))
'                chPT = g.tx(i, gPT)
'                bco = s2n(g.tx(i, gBANCC))
            '
            chCod = nuevoCodigo("cheques", "NroInt")

            'ChequesTerceros
            DataEnvironment1.dbo_INGCHEQUESTERCEROS "A" _
                , chCod, .chFecha(i), .chNumero(i), chMonto, nDoc, tdoc, dtFecha, Date _
                , Cheque_CARTERA, .chCodBanco(i), .chPT(i), cliente.codigo _
                , Date, UsuarioActual(), iddoc, 0

            'movi
            nMoviC = nuevoCodigo("movicaja", "movimiento")
            'MoviCaja
            DataEnvironment1.dbo_MOVICAJASdoc "A", 0, nMoviC, cliente.codigo, MoviCaja_CHEQUE, MoviCaja_INGRESO _
                , .chMonto(i), MoviConcepto, Date, chCta, 0, 1, tdoc, nDoc _
                , Date, UsuarioActual, chCod, iddoc
'''''            'DetMovCaja
'''''            DataEnvironment1.dbo_DETMOVCAJAS "A", nMoviC, chMonto, cliente.codigo, 0, DetConcepto, "RA"
            
            'ASiento cheque
            asien.AcumularItem chCta, chMonto, 0
        Next i
''        If .ChCant > 0 Then
''            che "Operacion concluida" & vbCrLf & "Puede anotar los codigos internos de los cheques"
''        End If
    End With
    
End Sub


Private Function BuscoNroYTipo(VerificoFecha As Boolean, Optional presupuesto As Boolean = False, Optional esFacturaDeCredito As Boolean = False) As Boolean
'OPCION PRESUPUESTO NO USA BACIGALUPPI
    Dim tmpfec, letra As String
    BuscoNroYTipo = True
    Dim ss As String, andtipo As String, tmp
    Dim sTipo As String, sPunto As String
    logFacturacion "Busco Nro y tipo", "", txtTipoDoc & " " & TxtNroFactura
    If presupuesto Then
        sTipo = PuntoVentaTipo(5)
        logFacturacion "Es presupuesto", "", txtTipoDoc & " " & TxtNroFactura
    Else
        sTipo = PuntoVentaTipo(cmbPunto.ListIndex)
        logFacturacion "Es factura", "", txtTipoDoc & " " & TxtNroFactura
    End If
    
    If mFAE Then
        letra = "E"
        logFacturacion "Es tipo E", "", txtTipoDoc & " " & TxtNroFactura
    Else
        letra = TipoFormVenta(ComboCodigo(cmbTipoIva))
        logFacturacion "No es tipo E", "", txtTipoDoc & " " & TxtNroFactura
    End If
    
    txtCodigo = nuevoCodigo("FacturaVenta", "codigo")
    logFacturacion "Validacion 15 dias", "", txtTipoDoc & " " & TxtNroFactura
    If Trim(txtTipoDocRef) = "FEA" Or Trim(txtTipoDocRef) = "FEB" Or Trim(txtTipoDocRef) = "FEC" Then
        If s2n(Date - obtenerDeSQL("select fecha from facturaventa where codigo=" & s2n(txtCodReferencia.Text))) >= 15 Then
            If NPregunta = False Then
                If MsgBox("Han transcurrido más de 15 días corridos desde la emisión de este comprobante. ¿Desea emitir una (NCx o NDx) según corresponda?", vbQuestion + vbYesNo, "ATENCION") = vbYes Then
                    esFacturaDeCredito = False
                    NPregunta = True
                    NRespuesta = True
                Else
                    esFacturaDeCredito = True
                    NPregunta = True
                    NRespuesta = False
                End If
            Else
                esFacturaDeCredito = Not NRespuesta
            End If
        Else
            esFacturaDeCredito = True
        End If
        logFacturacion "Validacion 30 dias correspondio", "", txtTipoDoc & " " & TxtNroFactura
    End If
    logFacturacion "Fin Validacion 30 dias", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Validacion letras del comprobante", "", txtTipoDoc & " " & TxtNroFactura
    If esFacturaDeCredito Then
        If mFac = FacturaVenta_NCredito Or mFac = FacturaVenta_NCreditoDevolucion Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_CREDITO_ELECTRONICO_A
                andtipo = " TipoDoc = '" & TipoDoc_CREDITO_ELECTRONICO_A & "' "
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_CREDITO_ELECTRONICO_B
                andtipo = " TipoDoc = '" & TipoDoc_CREDITO_ELECTRONICO_B & "' "
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_CREDITO_ELECTRONICO_C
                andtipo = " TipoDoc = '" & TipoDoc_CREDITO_ELECTRONICO_C & "' "
            End If
        ElseIf mFac = FacturaVenta_NDebito Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_DEBITO_ELECTRONICO_A
                andtipo = " TipoDoc = '" & TipoDoc_DEBITO_ELECTRONICO_A & "' "
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_DEBITO_ELECTRONICO_B
                andtipo = " TipoDoc = '" & TipoDoc_DEBITO_ELECTRONICO_B & "' "
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_DEBITO_ELECTRONICO_C
                andtipo = " TipoDoc = '" & TipoDoc_DEBITO_ELECTRONICO_C & "' "
            End If
        ElseIf mFac = FacturaVenta_Libre Or mFac = FacturaVenta_Pedido Or mFac = FacturaVenta_Remito Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_FACTURA_CREDITO_A
                andtipo = " TipoDoc = '" & TipoDoc_FACTURA_CREDITO_A & "' "
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_FACTURA_CREDITO_B
                andtipo = " TipoDoc = '" & TipoDoc_FACTURA_CREDITO_B & "' "
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_FACTURA_CREDITO_A
                andtipo = " TipoDoc = '" & TipoDoc_FACTURA_CREDITO_C & "' "
            End If
        Else
            MsgBox "No se encontro letra para documento para Tipo Iva :" & cmbTipoIva, vbCritical
            BuscoNroYTipo = False
            Exit Function
        End If
    
    Else
        If mFac = FacturaVenta_NCredito Or mFac = FacturaVenta_NCreditoDevolucion Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_NCREDITO_A
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_A & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_A & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_A & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NCREDITO_A & "' "
                End If
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_NCREDITO_B
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_B & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_B & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_B & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NCREDITO_B & "' "
                End If
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_NCREDITO_C
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_C & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_C & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_C & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NCREDITO_C & "' "
                End If
            ElseIf letra = "E" Then
                txtTipoDoc = TipoDoc_NCREDITO_E
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_E & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_E & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_E & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NCREDITO_E & "' "
                End If
            'ElseIf letra = "C" Then
            '    txtTipoDoc = TipoDoc_NCREDITO_C
            '    If sTipo = "PI" Then
            '        andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_C & "'  OR TipoDoc = '" & TipoDoc_NCREDITO_E & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_E & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_E & "' )"
            '    Else
            '        andtipo = " TipoDoc = '" & TipoDoc_NCREDITO_C & "' "
            '    End If
            End If
        ElseIf mFac = FacturaVenta_NDebito Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_NDEBITO_A
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_A & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_A & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_A & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NDEBITO_A & "' "
                End If
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_NDEBITO_B
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_B & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_B & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_B & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NDEBITO_B & "' "
                End If
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_NDEBITO_C
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_C & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_C & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_C & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NDEBITO_C & "' "
                End If
            ElseIf letra = "E" Then
                txtTipoDoc = TipoDoc_NDEBITO_E
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_E & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_E & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_E & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_NDEBITO_E & "' "
                End If
            End If
        ElseIf mFac = FacturaVenta_Libre Or mFac = FacturaVenta_Pedido Or mFac = FacturaVenta_Remito Then
            If letra = "A" Then
                txtTipoDoc = TipoDoc_FACTURA_A
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_A & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_A & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_A & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_FACTURA_A & "' "
                End If
            ElseIf letra = "B" Then
                txtTipoDoc = TipoDoc_FACTURA_B
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_B & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_B & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_B & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_FACTURA_B & "' "
                End If
            ElseIf letra = "C" Then
                txtTipoDoc = TipoDoc_FACTURA_C
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_C & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_C & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_C & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_FACTURA_C & "' "
                End If
            ElseIf letra = "E" Then
                txtTipoDoc = TipoDoc_FACTURA_E
                If sTipo = "PI" Or sTipo = "PI2" Then
                    andtipo = " (TipoDoc = '" & TipoDoc_NCREDITO_E & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_E & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_E & "' )"
                Else
                    andtipo = " TipoDoc = '" & TipoDoc_FACTURA_E & "' "
                End If
            'ElseIf letra = "C" Then
            '    txtTipoDoc = TipoDoc_FACTURA_C
            '    If sTipo = "PI" Then
            '        andtipo = " ( TipoDoc = '" & TipoDoc_FACTURA_C & "' OR TipoDoc = '" & TipoDoc_NCREDITO_E & "' OR " & " TipoDoc = '" & TipoDoc_NDEBITO_E & "' " & " OR " & " TipoDoc = '" & TipoDoc_FACTURA_E & "' )"
            '    Else
            '        andtipo = " TipoDoc = '" & TipoDoc_FACTURA_C & "' "
            '    End If
            End If
        ElseIf mFac = FacturaVenta_ticket Then
            txtTipoDoc = TipoDoc_TICKET
            andtipo = " TipoDoc = '" & TipoDoc_TICKET & "' "
        Else
            MsgBox "No se encontro letra para documento para Tipo Iva :" & cmbTipoIva, vbCritical
            BuscoNroYTipo = False
            Exit Function
        End If
    End If
    logFacturacion "Fin Validacion letras del comprobante", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Validacion punto de venta", "", txtTipoDoc & " " & TxtNroFactura
    sPunto = sSinNull(obtenerDeSQL("select puntoventa from documentoscae where tipo=" & ssTexto(txtTipoDoc) & " and tipopunto=" & ssTexto(sTipo)))
    If sPunto = "" Then
        MsgBox "No existe punto de venta para " & txtTipoDoc & "-" & sTipo
        BuscoNroYTipo = False
        logFacturacion "Validacion punto de venta no existe", "", txtTipoDoc & " " & TxtNroFactura
        Exit Function
    End If
    logFacturacion "Fin Validacion punto de venta", "", txtTipoDoc & " " & TxtNroFactura

    logFacturacion "Validacion numero de comprobante", "", txtTipoDoc & " " & TxtNroFactura
    If Trim(TxtNroFactura) = "" Then
        TxtNroFactura = s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) + 1
    Else
        If s2n(TxtNroFactura) <= s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) Then
            logFacturacion "Validacion numero de comprobante es menor al que corresponde", "", txtTipoDoc & " " & TxtNroFactura
            MsgBox "El numero de factura es MENOR que la ultima ingresada, se actualizara la numeracion.", vbInformation, "ATENCION"
            TxtNroFactura = s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) + 1
        ElseIf s2n(TxtNroFactura) > s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) + 1 Then
            logFacturacion "Validacion numero de comprobante es mayor al que corresponde", "", txtTipoDoc & " " & TxtNroFactura
            MsgBox "El numero de factura es varias veces MAYOR que la ultima ingresada, se actualizara la numeracion.", vbInformation, "ATENCION"
            TxtNroFactura = s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where " & andtipo & " and puntoventa=" & ssTexto(sPunto))) + 1
        End If
    End If
    logFacturacion "Fin Validacion numero de comprobante", "", txtTipoDoc & " " & TxtNroFactura
    
    logFacturacion "Validacion Nro Fecha Tipo y Punto", "", txtTipoDoc & " " & TxtNroFactura
    BuscoNroYTipo = RevisaNroYFechaOk("FacturaVenta", "NroFactura", "Fecha", s2n(TxtNroFactura, 0), dtFecha, andtipo, True, sPunto)
    logFacturacion "Fin Validacion Nro Fecha Tipo y Punto", "", txtTipoDoc & " " & TxtNroFactura
    If BuscoNroYTipo = False Then
        TxtNroFactura = s2n(obtenerDeSQL("select max(NroFactura) from FacturaVenta where  TipoDoc = " & ssTexto(txtTipoDoc) & " and puntoventa= " & ssTexto(sPunto))) + 1
        MsgBox "Se ha modificado el nro de factura,ahora puede grabar."
    End If
    
    If Not BuscoNroYTipo Then ucBoton.SetFocus
    
End Function




Private Sub set_uProd() ' lo copie de pedido cliente
    Dim sqlbuscar As String, sqldesc As String

    If Propio() Then    'propio
        sqldesc = "select descripcion from producto where codigo = '###' "
        sqlbuscar = "select codigo as [ Codigo                 ],  descripcion as [ Descripcion                                                 ] from producto where activo = 1 and facturable=1 order by codigo "
'        sqldesc = "select descripcion from producto where alias = '###' "
'        sqlbuscar = "select alias as [ Alias                   ], codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    Else    'relCliente
        sqldesc = "select descripcion from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & cliente.codigo & " and productoCliente = '###'"
        sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
            & " from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & cliente.codigo _
            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 and producto.facturable=1 " _
            & " order by producto"
    End If
    uProd.ini sqldesc, sqlbuscar, True
    uProd.EditaDescripcion = True
End Sub

'Private Sub uCuenta_cambio(codigo As Variant)
'
'End Sub

Private Sub uProd_cambio(codigo As Variant)
'    On Error Resume Next ' por las dudas
'    Dim preci As Double
'    If propio() Then
'        'preci = s2n(obtenerDeSQL("select precio from producto where codigo = '" & uProd.codigo & "'"))
'        preci = s2n(obtenerDeSQL("select precio from producto where codigo = '" & UProd.codigo & "'"))
'    Else
'        preci = s2n(obtenerDeSQL("select precio from relacion_producto_cliente where  productocliente = '" & UProd.codigo & "'"))
'    End If
    txtprecio = precioProducto(CStr(codigo), Propio(), cliente.codigo)
'    If preci <> 0 Then txtprecio = preci
    
    
    Dim tmpIvaProd As Variant ', tmpIvaClie
     
     tmpIvaProd = obtenerDeSQL("select iva from producto where codigo = '" & VerProductoMio(uProd.codigo, Propio()) & "'") '* 100
    'tmpivaclie = obtenerdesql("select porcentaje from porcentajesIva where activo = 1 and iva = " &
'    tmpIvaClie = obtenerDeSQL("SELECT porcentaje FROM PorcentajesIva INNER JOIN Clientes ON PorcentajesIva.iva = Clientes.iva WHERE clientes.activo = 1 and PorcentajesIva.activo = 1 and Clientes.codigo = " & cliente.Codigo)
    If mClienteConIva Then
        txtIvaProducto.enabled = (IsEmpty(tmpIvaProd))
        txtIvaProducto = (tmpIvaProd * 100)
        
        If (IsEmpty(tmpIvaProd)) Then
            txtIvaProducto.enabled = True
            txtIvaProducto = txtPIVA
        Else
            txtIvaProducto.enabled = False
            txtIvaProducto = (tmpIvaProd * 100)
        End If
    Else
        txtIvaProducto.enabled = False
        txtIvaProducto = "0"
    End If
    
End Sub

Private Function VaConRemito() As Boolean
    VaConRemito = optStock.item(ActuStock_RE).Value
End Function

Public Function ABMFacturaVenta(fOpe As String, fcodigo As Long, fTipoDoc As String, fNroFactura As Long, FFECHA As Date, fVencim As Date, fFormaPago As Long, fContado As Long, fCodCliente As Long, fRazonSoc As String, fProvincia As String, FCUIT As String, fTipoIva As Long, fVendedor As Long, fneto As Double, fPorcIva As Double, fIva As Double, ftotal As Double, fsaldo As Double, fDescuento As Double, fPedido As Long, fRemito As Long, fUsuario As Long, fHoy As Date, fCotizacion As Double, fMONEDA As Long, fIIBB As Double, fModiStock As Long, fDeposito As Long, fNCxDevolucion As Long, fNDxChRechaz As Long, fMotivoAjuste As Long, fNoGrav As Double, fIdDoc As Long, Optional fRemi As String = "", Optional fOrden As String = "", Optional fCotiLeye As Double = 0, Optional fVaLeye As Integer = 0, Optional fResalta As Integer = 0) As Boolean
On Error GoTo fmal
Dim idf As String
ABMFacturaVenta = True
Select Case fOpe
    Case "A":
        idf = " insert into FacturaVenta (Codigo, TipoDoc, NroFactura, Fecha, Vencimiento, FormaPago, Cliente, RazonSocial, provincia, cuit, TipoIva, Vendedor, Neto, PorcentajeIva, Iva, Total, Descuento, Saldo, contado, Pedido, Remito, Fecha_Alta, Usuario_Alta, Activo, COTIZACION, MONEDA, IIBB, actualizaStock, deposito,NC_xDevolucion, ND_xChequeRechazado, NoGrav, MotivoAjuste , iddoc,_control_ve,_docum_ve,CotizacionLeyenda,VaLeyendaCotizacion,resaltar ) " _
            & " values (" & fcodigo & "," & ssTexto(fTipoDoc) & "," & fNroFactura & "," & ssFecha(FFECHA) & "," & ssFecha(fVencim) & "," & fFormaPago & "," & fCodCliente & "," & ssTexto(fRazonSoc) & "," & ssTexto(fProvincia) & "," & ssTexto(FCUIT) & "," & fTipoIva & "," & fVendedor & "," & x2s(fneto) & "," & x2s(fPorcIva) & "," & x2s(fIva) & "," & x2s(ftotal) & "," & x2s(fDescuento) _
            & "," & x2s(fsaldo) & "," & fContado & "," & fPedido & "," & fRemito & "," & ssFecha(Date) & "," & fUsuario & ",1," & x2s(fCotizacion) & "," & fMONEDA & "," & x2s(fIIBB) & "," & fModiStock & "," & fDeposito & "," & fNCxDevolucion & "," & fNDxChRechaz & "," & fNoGrav & "," & fMotivoAjuste & "," & fIdDoc & "," & ssTexto(fOrden) & "," & ssTexto(fRemi) & "," & x2s(fCotiLeye) & "," & fVaLeye & "," & fResalta & ")"
        DataEnvironment1.Sistema.Execute idf
    Case "B":
        idf = " Update FacturaVenta   set activo = 0, fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & fUsuario _
            & " Where codigo = " & fcodigo
        DataEnvironment1.Sistema.Execute idf
    
End Select

Exit Function
fmal:
ABMFacturaVenta = False
End Function

Public Function ABMFVDetalle(vOpe As String, vCodFactura As Long, vTipoDoc As String, vNroFactura As Long, vCant As Double, vCodPropio As Long, vProducto As String, vDescripcion As String, vFormula As String, vPrecUni As Double, vPrecTot As Double, vPrecLista As Double, vNroPedido As Long, vNroRemito As Long, vItem As Long, vIdDoc As Long, vProdIva As Double, Optional vBajaStock As Long = 0, Optional vDeposito As Long = 0) As Boolean
On Error GoTo fvdmal
Dim idd As String
Dim f, fFactor As Double, fCargar As Double

Set f = Nothing
f = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(vProducto))
If IsNull(f) Or IsEmpty(f) Then
    fFactor = 1
Else
    fFactor = f
End If
fCargar = fFactor * vCant


ABMFVDetalle = True
Select Case vOpe
    Case "A":
        idd = "    insert into FacturaVentaDetalle (CodigoFactura, TipoDoc, NroFactura,        Cantidad, CodPropio, Producto, Formula, Descripcion,        PrecioUnitario, PrecioTotal, PrecioLista,        NroPedido , NroRemito, item_P_R, iddoc,_iva        )" _
        & " values(" & vCodFactura & "," & ssTexto(vTipoDoc) & "," & vNroFactura & "," & x2s(vCant) & "," & vCodPropio & "," & ssTexto(vProducto) & "," & ssTexto(vFormula) & "," & ssTexto(vDescripcion) & "," & x2s(vPrecUni) & "," & x2s(vPrecTot) _
        & "," & x2s(vPrecLista) & "," & vNroPedido & "," & vNroRemito & "," & vItem & "," & vIdDoc & "," & x2s(vProdIva) & ")"
        DataEnvironment1.Sistema.Execute idd
        
        If vNroRemito > 0 Then
            idd = " Update RemitoVentaDetalle " _
                & " Set facturar = facturar - " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
        End If
        If vNroPedido > 0 Then
            idd = " Update ItemPedidoCliente " _
                & " Set facturar = facturar - " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
        End If
        If vNroRemito = 0 And vBajaStock = 1 Then
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
        

    Case "B":
        If vNroRemito > 0 Then
             idd = " Update RemitoVentaDetalle " _
                & " Set facturar = facturar + " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
        End If
        If vNroPedido > 0 Then
            idd = " Update ItemPedidoCliente " _
                & " Set facturar = facturar + " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
        End If
        If vNroRemito = 0 And vBajaStock = 1 Then
        'If vBajaStock = 1 Then
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia + " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 + " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 + " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 + " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 + " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
End Select
Exit Function
fvdmal:
ABMFVDetalle = False
End Function

Private Sub uProd_LostFocus()
    Dim tmpIvaProd As Variant
    Dim dTipo As String
    dTipo = UCase(txtTipoDoc)
     
    tmpIvaProd = obtenerDeSQL("select iva from producto where codigo = '" & VerProductoMio(uProd.codigo, Propio()) & "'") '* 100
    
    If InStr(dTipo, "B") Then
        txtIvaProducto.enabled = True '(IsEmpty(tmpIvaProd))
    Else
        txtIvaProducto.enabled = (IsEmpty(tmpIvaProd))
    End If
    
End Sub

