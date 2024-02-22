VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRemitoVenta 
   Caption         =   "Mercaderia en Transito Salida"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   10545
   Icon            =   "frmRemitoVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tabRemito 
      Height          =   5160
      Left            =   60
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2280
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9102
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Items"
      TabPicture(0)   =   "frmRemitoVenta.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEdicion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmRemitoVenta.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPedidos"
      Tab(1).Control(1)=   "grilla2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Numeros de Serie"
      TabPicture(2)   =   "frmRemitoVenta.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(1)=   "lblErrorSeries"
      Tab(2).Control(2)=   "cmdLlenaSerie"
      Tab(2).Control(3)=   "chkSinSeries"
      Tab(2).Control(4)=   "grillaSeries"
      Tab(2).ControlCount=   5
      Begin VB.Frame fraEdicion 
         Height          =   1155
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   9915
         Begin VB.TextBox txtBarra 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1020
            TabIndex        =   15
            Top             =   660
            Width           =   1755
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   3660
            TabIndex        =   16
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtProductoDescripcion 
            Height          =   320
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   180
            Width           =   6375
         End
         Begin VB.TextBox txtProductoCodigo 
            Height          =   320
            Left            =   1020
            TabIndex        =   14
            Top             =   180
            Width           =   1815
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   7500
            TabIndex        =   18
            Top             =   660
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   8760
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmRemitoVenta.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdBorrar 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   9300
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmRemitoVenta.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "Borrar Item"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtConsignacion 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   5880
            TabIndex        =   17
            Top             =   660
            Width           =   915
         End
         Begin VB.CommandButton cmdAyuda 
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
            Left            =   2940
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   180
            Width           =   375
         End
         Begin VB.Label Label6 
            Caption         =   "C.Barra:"
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
            Left            =   120
            TabIndex        =   50
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Cantidad "
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
            Left            =   2820
            TabIndex        =   49
            Top             =   660
            Width           =   855
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   60
            TabIndex        =   48
            Top             =   240
            Width           =   1335
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   6780
            TabIndex        =   47
            Top             =   660
            Width           =   675
         End
         Begin VB.Label Label12 
            Caption         =   "Consignacion:"
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
            Left            =   4620
            TabIndex        =   46
            Top             =   660
            Width           =   1275
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   3990
         Left            =   -74760
         TabIndex        =   38
         Top             =   900
         Width           =   9435
         _cx             =   16642
         _cy             =   7038
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
         Height          =   3825
         Left            =   -74700
         TabIndex        =   37
         Top             =   1020
         Width           =   9375
         _cx             =   16536
         _cy             =   6747
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
      Begin VB.CheckBox chkSinSeries 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin Series"
         Height          =   315
         Left            =   -67140
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdLlenaSerie 
         Caption         =   "Llenar Serie"
         Height          =   315
         Left            =   -71520
         TabIndex        =   33
         ToolTipText     =   "Seleccione filas a llenar"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdPedidos 
         Caption         =   "Pedidos"
         Height          =   315
         Left            =   -74760
         TabIndex        =   31
         Top             =   540
         Width           =   1155
      End
      Begin VB.Frame fraDetalle 
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   180
         TabIndex        =   30
         Top             =   480
         Width           =   10215
         Begin VSFlex7LCtl.VSFlexGrid grilla 
            Height          =   3405
            Left            =   60
            TabIndex        =   39
            Top             =   1155
            Width           =   9735
            _cx             =   17171
            _cy             =   6006
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
      End
      Begin VB.Label lblErrorSeries 
         Caption         =   "--------------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -69840
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label10 
         Caption         =   "Puede hacer 'Doble Clic' en el campo  Nro.Serie"
         Height          =   495
         Left            =   -74280
         TabIndex        =   32
         Top             =   540
         Width           =   2235
      End
   End
   Begin VB.Frame fraCabecera 
      Height          =   2235
      Left            =   60
      TabIndex        =   24
      Top             =   0
      Width           =   10395
      Begin VB.TextBox txtClie 
         Height          =   320
         Left            =   1260
         TabIndex        =   56
         ToolTipText     =   "Presione la flecha hacia abajo para mostar coincidencias."
         Top             =   1020
         Width           =   4695
      End
      Begin VB.ComboBox cmbPunto 
         Height          =   315
         Left            =   4740
         TabIndex        =   52
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   3
         Left            =   9660
         TabIndex        =   13
         Top             =   1860
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   2
         Left            =   9120
         TabIndex        =   12
         Top             =   1860
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtobs 
         Height          =   320
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   1860
         Width           =   4935
      End
      Begin VB.TextBox txtobs 
         Height          =   320
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   1500
         Width           =   4875
      End
      Begin VB.CommandButton cmdCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?"
         Height          =   315
         Left            =   7140
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Propio "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8280
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbDeposito 
         Height          =   315
         ItemData        =   "frmRemitoVenta.frx":0F32
         Left            =   8220
         List            =   "frmRemitoVenta.frx":0F34
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.TextBox TxtRemitoNumero 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1260
         TabIndex        =   1
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdAyudaPedidos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?"
         Height          =   315
         Left            =   2580
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtPedido 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1260
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbTransporte 
         Height          =   315
         Left            =   4740
         TabIndex        =   6
         Top             =   600
         Width           =   2115
      End
      Begin VB.ComboBox cmbClienteNombre 
         Height          =   315
         Left            =   7620
         TabIndex        =   9
         Top             =   1020
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtClienteCodigo 
         Height          =   320
         Left            =   6660
         TabIndex        =   7
         Top             =   1020
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   8220
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70189057
         CurrentDate     =   38126
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8760
         TabIndex        =   55
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Cod.:"
         Height          =   255
         Left            =   8160
         TabIndex        =   54
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "P.Remito:"
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
         Left            =   3480
         TabIndex        =   53
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label13 
         Caption         =   "Empleado"
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
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   40
         Top             =   1850
         Width           =   795
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   35
         Top             =   660
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Pedido:"
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
         Left            =   480
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Transporte:"
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
         Left            =   3480
         TabIndex        =   28
         Top             =   600
         Width           =   1035
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         TabIndex        =   27
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "NroRemito :"
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
         Left            =   180
         TabIndex        =   26
         Top             =   180
         Width           =   1215
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
         Left            =   7260
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
   End
   Begin Gestion.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   7440
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   2778
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucEntreFechas ucBetween 
         Height          =   315
         Left            =   60
         TabIndex        =   42
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
      End
      Begin VB.Label lblTotalRV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   9060
         TabIndex        =   44
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total :"
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
         Left            =   7980
         TabIndex        =   43
         Top             =   60
         Width           =   975
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMenu2 
         Caption         =   "menu2"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmRemitoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '26/11/4
Private g As LiGrilla, g2 As LiGrilla, g3 As LiGrilla
Private WithEvents cliente As LiCodigo
Attribute cliente.VB_VarHelpID = -1
Private mUltimoPedido As Long
Private mNuevo As Boolean

Private gCODI As Long
Private gDESC As Long
Private gCANT As Long
Private gPEDI As Long
Private gPREC As Long
Private gFORM As Long
Private gCONS As Long

Private g2CODI As Long
Private g2DESC As Long
Private g2CANT As Long
Private g2PEDI As Long
Private g2PREC As Long
Private g2FORM As Long

Private g3ITEM As Long
Private g3PROD As Long
Private g3DESC As Long
Private g3NSER As Long
Private g3HIDD As Long
Private g3CONS As Long
Private g3ALTA As Long

Private Const CTE_SERIE_AGREGAR = "Registrar"


'Private Sub cliente_cambio(codigo As Variant)
'    Dim nc As Long, np As Long, i As Long
'    nc = cliente.codigo
'    np = s2n(txtPedido)
'
'    If nc = 0 Then
'        g.Borrar
'        g2.Borrar
'    Else
'        If np > 0 Then
'            g2.Borrar
'            Carga1 nc, np
'        Else
'            g.Borrar
'            Carga2 nc, 0
'        End If
'        i = s2n(obtenerDato("clientes", cliente.codigo, "transporte"))
'        If i > 0 Then cmbTransporte = ObtenerDescripcion("transportes", i)
'    End If
'End Sub

Private Sub cmbPunto_Click()
    Dim punto As String
    punto = obtenerDeSQL("select punto from puntoremito where descripcion='" & Trim(cmbPunto.Text) & "'")
    Label16.caption = s2n(obtenerDeSQL("select max(codigo) as mas from remitoventa ")) + 1
    TxtRemitoNumero.Text = s2n(obtenerDeSQL("select max(numero) from remitoventa where puntoventa='" & Trim(punto) & "'")) + 1
End Sub

Private Sub cmdAgregar_Click()
    Dim r As Long
    Dim pco As String
    
    If Not IsNumeric(txtCantidad) Or txtProductoCodigo = "" Or txtProductoDescripcion = "" Then Exit Sub
    If s2n(txtConsignacion) > s2n(txtCantidad) Then
        txtConsignacion.SetFocus
        txtConsignacion.SelStart = 0
        txtConsignacion.SelLength = Len(txtConsignacion.Text)
        Exit Sub
    End If
    
    pco = Trim(txtProductoCodigo)
    MetoEnGrilla pco, txtProductoDescripcion, s2n(txtCantidad, 0), "", s2n(txtprecio, 4), s2n(txtConsignacion, 4), ""
    txtCantidad = ""
    txtProductoCodigo = ""
    txtProductoDescripcion = ""
    txtprecio = ""
    txtConsignacion = ""
    txtBarra.Text = ""
    frmImagen.Imagen.Cls
    txtProductoCodigo.SetFocus
    
    chkPropioEnabled True ' permite habilitar, ...
    CalculaTotal
End Sub

Private Sub cmdAyuda_Click()
    Dim Ubicacion As String
    
    frmBuscarProducto
    If frmBuscar.resultado() = "" Then Exit Sub
    
    txtProductoCodigo = frmBuscar.resultado(1)
    txtProductoDescripcion = frmBuscar.resultado(2)
    txtBarra = sSinNull(obtenerDeSQL("select codigobarra from producto where activo = 1 and codigo = '" & Trim(txtProductoCodigo) & "'"))
    Ubicacion = sSinNull(obtenerDeSQL("select grafico from producto where codigo='" & Trim(txtProductoCodigo.Text) & "'"))
    If Trim(txtProductoCodigo.Text) <> "" And Ubicacion <> "" Then
        frmImagen.Imagen.PaintPicture LoadPicture(Ubicacion), 0, 0, frmImagen.Imagen.ScaleWidth, frmImagen.Imagen.ScaleHeight
    Else
        frmImagen.Imagen.Cls
    End If
    
    'Imagen.PaintPicture LoadPicture("C:\WINDOWS\Pompas.bmp"), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
End Sub


Private Sub cmdAyudaPedidos_Click()
    Dim ssql As String ', i As Long
    
    With frmBuscar
        ssql = " SELECT DISTINCT numero as [ Numero    ], fecha as [ Fecha           ], Cliente, Clientes.Descripcion as [ Nombre Cliente     ], CodigoPropio " _
            & " FROM ItemPedidoCliente " _
            & " INNER JOIN Pedidos_Clientes ON ItemPedidoCliente.PEDIDO = Pedidos_Clientes.numero " _
            & " INNER join Clientes on Pedidos_Clientes.cliente = clientes.codigo"
       
        If cliente.codigo = 0 Then
            ssql = ssql & " where Pedidos_Clientes.activo = 1 and ItemPedidoCliente.saldo > 0 order by Pedidos_Clientes.numero desc"
        Else
            ssql = ssql & " where Pedidos_Clientes.activo = 1 and ItemPedidoCliente.saldo > 0 and cliente = " & cliente.codigo & " order by Pedidos_Clientes.numero desc"
        End If
        If .MostrarSql(ssql, , , , "SI", "  ") = "" Then Exit Sub

        txtPedido = .resultado(1)
        cliente.codigo = .resultado(3)
        chkPropio = IIf(.resultado(4) = 1, vbChecked, vbUnchecked)
        
        CargoTransporteDePedido txtPedido
        Carga1 s2n(cliente.codigo), s2n(txtPedido)
        tabRemito.Tab = 0
    End With
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    If g.Row > 0 Then
        If MsgBox("¿Desea Actualizar Stock?", vbYesNo) = vbYes Then
            ABMRVDetalle "B", s2n(TxtRemitoNumero), g.tx(g.Row, gCODI), s2n(g.tx(g.Row, gCANT), 4), s2n(g.tx(g.Row, gPREC), 4), s2n(g.tx(g.Row, gPEDI), 4), 0, s2n(g.tx(g.Row, gCONS)), g.tx(g.Row, gFORM), 1, Label16.caption
        End If
        grilla.RemoveItem (g.Row)
    End If
    chkPropioEnabled True
    CalculaTotal
End Sub

Private Sub cmdPedidos_Click()
    Carga2 s2n(cliente.codigo), s2n(txtPedido)
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
Dim rsEjercicio As New ADODB.Recordset
    inigrilla
    Set cliente = New LiCodigo
    cliente.init cmbClienteNombre, txtClienteCodigo, "Clientes", False, False, cmdCliente, "activo = 1", True
    
    ucMenu.init True, True, True, True, True ',
'    ucMenu.MsgConfirmaCancelar = "¿ Desea abandonar la edicion ? "
'    ucMenu.MsgConfirmaSalir = "Cerrar Formulario ? "
'    ucMenu.MsgConfirmaEliminar = "Anula este Remito ?"
'    ucMenu.CaptionEliminar = "Anular"
    
    chkPropio.Value = vbChecked
    comboSql cmbTransporte, "select descripcion from transportes where activo = 1"
    comboArray cmbDeposito, Array("Deposito Central", "Deposito 1", "Deposito 2", "Deposito 3", "Deposito4"), Array(0, 1, 2, 3, 4)
    comboSql cmbPunto, "select descripcion from puntoremito where activo = 1"
    
    rsEjercicio.Open "SELECT * From Ejercicio WHERE activo =1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    ucBetween.ini rsEjercicio!FechaInicio, rsEjercicio!FechaFin
    dtFecha = Date
    
    lblPrecio.Visible = REMITO_CON_PRECIO
    txtprecio.Visible = REMITO_CON_PRECIO
    
    tabRemito.TabVisible(2) = gEMPR_Maneja_series
    tabRemito.TabVisible(1) = 0
End Sub


Private Sub Form_Resize()
'    encajar ucBetween, Me, , 0, ucMenu.Height
    'encajar tabRemito, Me, 2280, 60, 100 + ucMenu.Height + ucBetween.Height + 120, 120
    encajar tabRemito, Me, 2280, 60, ucMenu.Height, 120
    encajar fraDetalle, tabRemito, 360, 100, 100, 100
    encajar grilla, fraDetalle, 1200, 50, 50, 50
    encajar grilla2, tabRemito, 1080, 60, 120, 300  '260
    encajar grillaSeries, tabRemito, 1080, 60, 120, 300  ' 260
End Sub


Private Sub Form_Terminate()
    Set g2 = Nothing
    Set g = Nothing
    Set g3 = Nothing
    Set cliente = Nothing
End Sub


Private Sub Grilla_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    If Not mNuevo Then Exit Sub
    If (Col = gPREC Or Col = gCANT) And txtClienteCodigo.Text <> "" Then
        If grilla.TextMatrix(Row, gPEDI) <> "" Then
            sql = "select cantidad, saldo  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where cliente = " & txtClienteCodigo.Text & " and activo = 1 and saldo>0 and pedido=" & grilla.TextMatrix(Row, gPEDI) & " and producto='" & grilla.TextMatrix(Row, gCODI) & "'"
            rs.Open sql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If rs.EOF = True And rs.BOF = True Then
                CalculaTotal
            Else
                rs.MoveFirst
                While Not rs.EOF
                    If CDbl(g.TextMatrix(Row, gCANT)) <= CDbl(rs!saldo) Then
                        CalculaTotal
                        Exit Sub
                    Else
                        rs.MoveNext
                        If rs.EOF Then
                            MsgBox "La cantidad ingresada supera la del pedido.", vbExclamation, "Advertencia"
                            rs.MovePrevious
                            grilla.TextMatrix(Row, gCANT) = rs!saldo
                            Exit Sub
                        End If
                        
                    End If
                Wend
            End If
        End If
    End If
End Sub

Private Sub grilla_DblClick()
    If ucMenu.estado <> ucbEditando Then Exit Sub

    chkPropioEnabled True
    If g.Col = gCODI And g.tx(g.Row, gPEDI) = "" Then
        frmBuscarProducto
        If frmBuscar.resultado() <> "" Then
            grilla.TextMatrix(g.Row, gCODI) = frmBuscar.resultado(1)
            grilla.TextMatrix(g.Row, gDESC) = frmBuscar.resultado(2)
        End If
    End If
End Sub


Private Sub grilla2_DblClick()
    Dim r   As Long
    
    If ucMenu.estado <> ucbEditando Then Exit Sub
    
    r = g2.Row
    If r > 0 Then
        MetoEnGrilla _
            g2.tx(r, g2CODI) _
            , g2.tx(r, g2DESC) _
            , g2.tx(r, g2CANT) _
            , g2.tx(r, g2PEDI) _
            , g2.tx(r, g2PREC) _
            , "" _
            , g2.tx(r, g2FORM)
        g2.delRow (r)
    End If
End Sub

Private Sub CargaRemito(cual, cod)
    If ON_ERROR_HABILITADO Then On Error GoTo E_UFA
    Dim rs As New ADODB.Recordset, i As Long, desc As Variant, prod As String
    
    With rs
        .Open "select * from remitoVenta where numero = " & cual & " and codigo=" & cod, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        TxtRemitoNumero = !numero
        Label16.caption = !codigo
        'cliente.codigo = !cliente
        txtClie.Text = !descriclie
        'cmbClienteNombre = ObtenerDescripcion("clientes", !Cliente)
        dtFecha = !Fecha
        cmbDeposito.ListIndex = s2n(!DEPOSITO)
        txtObs(0) = sSinNull(!obs1)
        txtObs(1) = sSinNull(!obs2)
        txtObs(2) = sSinNull(!Obs3)
        txtObs(3) = sSinNull(!obs4)
        cmbTransporte = ObtenerDescripcion("transportes", s2n(!Transporte))
        cmbPunto.Text = obtenerDeSQL("select descripcion from puntoremito where punto='" & Trim(!PuntoVenta) & "'")
        .Close
    
        .Open "select * from RemitoVentaDetalle where CANCELADO=0 AND numero = " & cual & " and codremito=" & cod, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        g.rows = 1
        While Not .EOF
            i = g.addRow()
            
            desc = obtenerDeSQL("select descripcion from producto where codigo = '" & Trim(!producto) & "' and activo = 1 ")
            If IsNull(desc) Then desc = ""
            
            'grilla.TextMatrix(i, gCODI) = VerProductoCliente(!producto, Propio(), cliente.codigo)
            grilla.TextMatrix(i, gCODI) = VerProductoCliente(!producto, Propio(), 0)
            grilla.TextMatrix(i, gDESC) = desc
            grilla.TextMatrix(i, gCANT) = !cantidad
            grilla.TextMatrix(i, gPEDI) = !Pedido
            grilla.TextMatrix(i, gPREC) = s2n(!precio, 4)
            grilla.TextMatrix(i, gFORM) = !formula
            .MoveNext
        Wend
        .Close
        
        .Open "select producto, serie, consignacion from series where nroComprobante = " & cual & " and Comprobante = " & TipoComprobante_REMITOVENTA, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        g3.Borrar
        While Not .EOF
            i = g3.addRow()
            prod = VerProductoCliente(!producto, Propio(), cliente.codigo)
            grillaSeries.TextMatrix(i, g3PROD) = prod
            grillaSeries.TextMatrix(i, g3DESC) = ProductoDescripcion(prod)
            grillaSeries.TextMatrix(i, g3NSER) = !Serie
            grillaSeries.TextMatrix(i, g3CONS) = !consignacion
            .MoveNext
        Wend
        .Close
    End With
    CalculaTotal
E_UFA:
    Set rs = Nothing
End Sub


Private Function FaltaCabecera() As Boolean
    FaltaCabecera = (cliente.codigo = 0) ' Or txtClienteCodigo = "")
End Function


Private Function FaltaGrilla() As Boolean
    FaltaGrilla = (g.suma(gCANT) = 0) ' Or g.Buscar(gCODI, "")
End Function



Private Sub Carga1(NCliente, NPedido)
    If ON_ERROR_HABILITADO Then On Error GoTo UFAcarga1
    
    Dim rs As New ADODB.Recordset, ssql As String, i As Long
    Dim can As Double, pre As Double, formul As String

    If NCliente + NPedido = 0 Then Exit Sub
    
    g.rows = 1
    'sSql = "select producto, cantidad, pedido, precio, CodigoPropio, pago  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where cliente = " & nCliente & " and activo = 1 and estado <> '" & ESTADO_ENTREGADO & "' And pedido = " & nPedido
    ssql = "select producto, cantidad, pedido, precio, CodigoPropio, pago, cliente, saldo, formula  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where activo = 1 and saldo >0 "
    If NCliente > 0 Then ssql = ssql & " and cliente = " & NCliente
    If NPedido > 0 Then ssql = ssql & " and pedido = " & NPedido
    
    With rs
        
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then cliente.codigo = !cliente
        g.Borrar
        While Not .EOF
            chkPropio = IIf(!CODIGOPROPIO, vbChecked, vbUnchecked)
            'can = s2n(!cantidad)
            can = s2n(!saldo)
            pre = s2n(!precio)
            formul = sSinNull(!formula)
            MetoEnGrilla VerProductoCliente(!producto, Propio(), cliente.codigo), ObtenerDescripcionS("producto", !producto), can, NPedido, pre, 0, formul
            .MoveNext
        Wend
    End With
    tabRemito.Tab = 0

fin:
    Set rs = Nothing
    Exit Sub
UFAcarga1:
    ufa "Err al cargar", Me.Name & " carga1" & NCliente & "   " & NPedido ', Err
    Resume fin
End Sub


Private Sub Carga2(NCliente, NPedido)
    Dim rs As New ADODB.Recordset, ssql As String, i As Long

    g2.rows = 1
    ssql = "select producto, cantidad, pedido, precio, CodigoPropio, pago, saldo, formula  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where cliente = " & NCliente & " and activo = 1 "
    If NPedido > 0 Then
        'sSql = " and estado <> '" & ESTADO_ENTREGADO & "' And pedido = " & nPedido
        ssql = ssql & " and saldo >0 and pedido = " & NPedido
    Else
        'sSql = " and estado <> '" & ESTADO_ENTREGADO & "'" ' and pedido = " & nPedido
        ssql = ssql & " and saldo >0 " 'and pedido = " & nPedido
    End If
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            chkPropio = IIf(!CODIGOPROPIO, vbChecked, vbUnchecked)
            i = g2.addRow()
            grilla2.TextMatrix(i, g2CODI) = VerProductoCliente(!producto, Propio(), CLng(NCliente))
            grilla2.TextMatrix(i, g2DESC) = ObtenerDescripcionS("producto", !producto)
            grilla2.TextMatrix(i, g2CANT) = !saldo 'cantidad
            grilla2.TextMatrix(i, g2PEDI) = !Pedido
            grilla2.TextMatrix(i, g2PREC) = !precio
            grilla2.TextMatrix(i, g2FORM) = !formula
            .MoveNext
        Wend
    End With
    'tabRemito.Tab = 1
    Set rs = Nothing
End Sub


Private Function Propio() As Boolean
    Propio = (chkPropio.Value = vbChecked)
End Function


Private Sub BorrarCampos()
    FrmBorrarCbo Me
    FrmBorrarTxt Me
    'cmbTransporte = ""
    'TxtRemitoNumero = ""
    
    'cliente.codigo = 0
    chkSinSeries.Value = 0
    g.Borrar
    g2.Borrar
    g3.Borrar
    tabRemito.Tab = 0
    mUltimoPedido = 0
    frmImagen.Imagen.Cls
End Sub


Private Sub HabilitarEdicion(habilitar As Boolean)
    fraCabecera.enabled = habilitar
    fraDetalle.enabled = habilitar
    tabRemito.enabled = True 'habilitar
    fraEdicion.enabled = habilitar
    cmdLlenaSerie.enabled = habilitar
    cmdPedidos.enabled = habilitar
    
    grilla.Editable = IIf(habilitar, flexEDKbdMouse, flexEDNone)
    grillaSeries.Editable = IIf(habilitar, flexEDKbdMouse, flexEDNone)
    grilla2.Editable = IIf(habilitar, flexEDKbdMouse, flexEDNone)

End Sub



Private Sub inigrilla()
    Set g = New LiGrilla
    Set g2 = New LiGrilla
    Set g3 = New LiGrilla
   
    g.init grilla
    g2.init grilla2
    g3.init grillaSeries
    
    gCODI = g.AddCol("  Producto                      ")
    gDESC = g.AddCol("  Descripcion                                                      ")
    gCANT = g.AddCol(" C.Total ", "N", 0)
    gPEDI = g.AddCol(" NroPedido ")
    gCONS = g.AddCol(" C.Consgn ", "N", 0)
    gFORM = g.AddCol(" Producto Base ")
    g.Borrar
    
    g2CODI = g2.AddCol("  Producto                    ")
    g2DESC = g2.AddCol(" Descripcion                                                      ")
    g2CANT = g2.AddCol(" Cantidad  ", "9", 0)
    g2PEDI = g2.AddCol(" Nro Pedido ")
    g2FORM = g2.AddCol("Formula ")
    g2.Borrar
    
    g3ITEM = g3.AddCol("  -  ", "A")
    g3PROD = g3.AddCol(" Producto                      ")
    g3DESC = g3.AddCol(" Descripcion                                                      ")
    g3NSER = g3.AddCol(" Numero de Serie            ", "S") ' editable
    g3CONS = g3.AddCol("Consignacion", "K")
    g3HIDD = g3.AddCol("", "H")
    g3ALTA = g3.AddCol("                ")
    g3.Borrar
    grillaSeries.SelectionMode = flexSelectionListBox
    
    If REMITO_CON_PRECIO Then                 ' $OLO un Usuario muestra lo$ precio$
        gPREC = g.AddCol(" Precio      ", "N", 4)
        g2PREC = g2.AddCol(" Precio      ")
    Else                                ' Oculto precio
        gPREC = g.AddCol(" Precio      ", "H")
        g2PREC = g2.AddCol(" Precio      ", "H")
    End If
    
End Sub


Private Sub chkPropioEnabled(que As Boolean)
    chkPropio.enabled = grilla.rows < 2 And que
End Sub

Private Sub CalculaTotal()
    Dim i As Long, tot As Double
    
    If g.cols < gPREC Then Exit Sub
        With g
            For i = 1 To .rows - 1
                tot = tot + s2n(.TextMatrix(i, gCANT), 4) * s2n(.TextMatrix(i, gPREC), 4)
            Next i
        End With
    lblTotalRV = n2r(tot, 2)
End Sub


Private Sub MetoEnGrilla(codi, desc, cant, pedi, prec, cons, formula As String)
    If ON_ERROR_HABILITADO Then On Error GoTo ErrMETO
    
    Dim i As Long, ssql As String, codigomio As String, frmul As String
    Dim VirtualConFormula As String

    codigomio = VerProductoMio(codi, Propio())
    
    If EsProductoVirtual(codigomio) Then frmul = CHAR_PROD_VIRTUAL Else frmul = formula 'VirtualConFormula = CHAR_PROD_VIRTUAL
    
    i = g.addRow()
    grilla.TextMatrix(i, gCODI) = codi
    grilla.TextMatrix(i, gDESC) = desc
    grilla.TextMatrix(i, gCANT) = cant
    grilla.TextMatrix(i, gPEDI) = pedi
    grilla.TextMatrix(i, gPREC) = prec 'oculta a veces
    grilla.TextMatrix(i, gCONS) = cons
    grilla.TextMatrix(i, gFORM) = frmul 'VirtualConFormula
    
    If ProductoConSerie(codigomio, Propio()) Then
        MetoGrillaSeries codigomio, s2n(cant), s2n(cons)
    End If
    
    If s2n(pedi) > 0 Then CargoTransporteDePedido pedi
    
fin:
    Exit Sub
ErrMETO:
    ufa "", "metoEnGrilla " & Me.Name ', Err
    Resume fin
End Sub


Private Sub MetoGrillaSeries(prod As String, ByVal cant As Long, ByVal consig As Long)
    If ON_ERROR_HABILITADO Then On Error GoTo ERR_FIN
    Dim i As Long, r As Long ', c As long
   
    'If obtenerDato("producto", prod, "serie") Then
    If ProductoConSerie(prod, Propio()) Then
        'asserts saludables
        If g3.rows > 100 Then
            ufa " Demasiados items para num de serie ", "Remito Venta" ', Err
            Exit Sub
        End If
        If cant < 1 Then
            ufa "", "Cantidad para num serie < 1 " & Me.Name ', Err
            Exit Sub
        End If
        
        For i = 1 To cant
            r = g3.addRow()
            grillaSeries.TextMatrix(r, g3PROD) = prod
            grillaSeries.TextMatrix(r, g3DESC) = ProductoDescripcion(prod)
            If consig > 0 Then
                grillaSeries.TextMatrix(r, g3CONS) = flexChecked
                consig = consig - 1
            End If
        Next i
    End If
    
    GoTo fin

ERR_FIN:
    ufa "err en series ", Me.Name ', Err
fin:
End Sub

Private Sub grillaSeries_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    

    
    If Col = g3NSER Then
        For i = 1 To g3.rows - 1
            If (i <> Row) And (g3.tx(i, g3PROD) = g3.tx(Row, g3PROD)) And (Trim(g3.tx(i, g3NSER)) <> "") Then
                If g3.tx(i, g3NSER) = g3.tx(Row, g3NSER) Then
                    che " Series de filas " & i & " y " & Row & " iguales"
                End If
            End If
        Next i
    End If
End Sub

Private Sub grillaSeries_dblClick()
    Dim r As Long, prod As String, resu As String
    
    If ucMenu.estado <> ucbEditando Then Exit Sub
    
    r = g3.Row
    If r < 1 Then Exit Sub
    prod = VerProductoMio(g3.tx(r, g3PROD), Propio())
    
    resu = Buscar_SeriesEnStock(prod)
    If resu > "" Then grillaSeries.TextMatrix(r, g3NSER) = resu
End Sub


Private Sub mnuMenu2_Click(Index As Integer)
    Dim str As String
    
    str = mnuMenu2(Index).caption
    If str <> "" Then
        txtClie = str
    End If
End Sub

Private Sub tabRemito_Click(PreviousTab As Integer)
    If tabRemito.Tab = 2 And PreviousTab <> 2 Then LlenoGrillaSeries
End Sub
    
Private Sub LlenoGrillaSeries()
    Dim i As Long, j As Long, prod As String, cant As Long, cons

        'borrosinserie
        i = 1
        While i < g3.rows
            If g3.tx(i, g3NSER) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend

        'borro marcas
        For i = 1 To g3.rows - 1
            grillaSeries.TextMatrix(i, g3HIDD) = ""
        Next i

        'marco o agrego en grilla series
        For i = 1 To g.rows - 1
            prod = Trim(g.tx(i, gCODI))
            cant = s2n(g.tx(i, gCANT))
            cons = s2n(g.tx(i, gCONS))
            If ProductoConSerie(prod, Propio()) Then
                For j = 1 To cant
                    If marcoG3(prod, cons) Then
                        cons = cons - 1
                    End If
                Next j
            End If
        Next i

        'borro no marcadas
        i = 1
        While i < g3.rows
            If g3.tx(i, g3HIDD) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend
End Sub


Private Function marcoG3(codi, ByVal cons) '
    Dim i As Long
    
    For i = 1 To g3.rows - 1
        If g3.tx(i, g3PROD) = codi And g3.tx(i, g3HIDD) = "" Then
            grillaSeries.TextMatrix(i, g3HIDD) = "X"
            Exit Function
        End If
    Next i
    i = g3.addRow()
    grillaSeries.TextMatrix(i, g3PROD) = codi
    grillaSeries.TextMatrix(i, g3DESC) = ProductoDescripcion(codi)
    grillaSeries.TextMatrix(i, g3HIDD) = "X"
    grillaSeries.TextMatrix(i, g3CONS) = IIf(cons > 0, "-1", "0")
    
    marcoG3 = (g3.tx(i, g3CONS) = "-1")
End Function


Private Sub cmdLlenaSerie_Click()
    Dim i As Long, Row As Long
    Dim s As String, n As Long       ' current cell
    Dim ss As String, ns As Long     ' the longer one
    
    ss = ""
    n = 0
    
    With grillaSeries ' uso propiedades que no exporte

        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, g3NSER)
            
            If Len(s) > Len(ss) Then
                ss = s
                ns = Len(ss)
            End If
        Next i
    
        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, g3NSER)
            n = Len(s)
            If n < ns And n > 0 Then .TextMatrix(Row, g3NSER) = Left(ss, ns - n) & s
        Next i
    End With

End Sub


'pasar a gral
Private Sub GrabarBs_Num(campo As String, dato)
    DataEnvironment1.Sistema.Execute "update BS set " & campo & " = " & dato
End Sub
'
Private Function LeerBS_Num(campo As String)
    LeerBS_Num = obtenerDeSQL("select " & campo & " from BS ")
End Function


Private Function FaltaSeries() As Boolean
    Dim r As Long, i As Long, ns As String
    Dim seri As String, prod As String
    
    FaltaSeries = False
    lblErrorSeries.Visible = False
    
    If chkSinSeries.Value = vbChecked Then Exit Function
    
    LlenoGrillaSeries
    r = g3.rows
    
    'vacio
    If r > 1 And g3.buscar(g3NSER, "") > 0 Then
        
        tabRemito.Tab = 2
        grillaSeries.SetFocus
        grillaSeries.Select g3.PrimerVacio(g3NSER), g3NSER
        
        FaltaSeries = True
        Exit Function
    End If
    
    'existe serie ?
    For i = 1 To r - 1
        seri = g3.tx(i, g3NSER)
        prod = g3.tx(i, g3PROD)
        If Not SerieEnStock(seri, prod) Then
            If g3.tx(i, g3ALTA) <> CTE_SERIE_AGREGAR Then
                che "No figura en stock " & vbCrLf & prod & vbCrLf & seri
                
                tabRemito.Tab = 2
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER
                If confirma("Desea registrarlo ahora") Then
                    g3.tx i, g3ALTA, CTE_SERIE_AGREGAR
                Else
                    FaltaSeries = True
                    Exit Function
                End If
            End If
        End If
    Next i
     
    If r > 1 Then
        For i = 1 To r - 2
            ns = g3.tx(i, g3NSER)
            If ns <> "" And g3.buscar(g3NSER, ns, i + 1) > 0 Then
                tabRemito.Tab = 2
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER, g3.buscar(g3NSER, ns, i + 1), g3NSER
                
                'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                FaltaSeries = True
                Exit Function
            End If
        Next i
    End If

End Function

Private Sub txtBarra_LostFocus()
    Dim desc
    
    If txtBarra = "" Then Exit Sub
    
    desc = obtenerDeSQL("select codigo from producto where activo = 1 and codigobarra = '" & Trim(txtBarra) & "'")
    If desc = "" Then
        MsgBox "El codigo de barra no existe para ningun producto.", , "ATENCION"
        txtBarra = ""
    Else
        txtProductoCodigo = desc
    End If
    
    desc = obtenerDeSQL("select descripcion from producto where activo = 1 and codigo = '" & Trim(txtProductoCodigo) & "'")
    If desc = "" Then txtProductoCodigo.SetFocus
    txtProductoDescripcion = desc
End Sub

Private Sub txtCantidad_LostFocus()
    Dim cant As Double

    cant = s2n(obtenerDeSQL("select stockmin from producto where codigo='" & Trim(txtProductoCodigo.Text) & "'"))
    If cant > s2n(cant - s2n(txtCantidad)) And cant <> 0 Then
        MsgBox "El stock esta por debajo del minimo establecido.", , "ATENCION"
    End If
End Sub

'Private Sub txtClie_Change()
'    Dim rs As New ADODB.Recordset
'    Dim sql As String
'    Dim i As Long
'    Dim j As Long
'
'    If ucMenu.estado = ucbEditando Then
'        If Len(txtClie.Text) >= 3 Then
'            sql = "select descriclie from remitoventa where anulado=0 and descriClie like '%" & Trim(txtClie) & "%'"
'            rs.Open sql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'
'            j = 1
'            While j < mnuMenu2.Count
'                Unload mnuMenu2(j)
'                j = j + 1
'            Wend
'            i = 1
'            If (rs.EOF = True And rs.BOF) Or IsNull(rs!descriclie) Or IsEmpty(rs!descriclie) Then
'                mnuMenu2(0).caption = " "
'            Else
'                mnuMenu2(0).caption = rs!descriclie
'                rs.MoveNext
'            End If
'            While Not rs.EOF
'                'mnuMenu(0).caption = "Menu"
'                'Load mnuMenu(1)
'                'mnuMenu(1).caption = "Salir"
'
'                Load mnuMenu2(i)
'                mnuMenu2(i).caption = rs!descriclie
'
'                rs.MoveNext
'            Wend
'            If mnuMenu2(0).caption <> " " Then
'                PopupMenu mnuMenu, , 1500, 1350
'            End If
'        End If
'    End If
'End Sub

Private Sub txtClie_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim i As Long
    Dim j As Long
    
    If ucMenu.estado = ucbEditando Then
        If KeyCode = 40 Then
            'If Len(txtClie.Text) >= 3 Then
                sql = "select descriclie from remitoventa where anulado=0 and descriClie like '%" & Trim(txtClie) & "%'"
                rs.Open sql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                j = 1
                While j < mnuMenu2.Count
                    Unload mnuMenu2(j)
                    j = j + 1
                Wend
                i = 1
                If (rs.EOF = True And rs.BOF) Or IsNull(rs!descriclie) Or IsEmpty(rs!descriclie) Then
                    mnuMenu2(0).caption = " "
                Else
                    mnuMenu2(0).caption = rs!descriclie
                    rs.MoveNext
                End If
                While Not rs.EOF
                    'mnuMenu(0).caption = "Menu"
                    'Load mnuMenu(1)
                    'mnuMenu(1).caption = "Salir"
                    
                    Load mnuMenu2(i)
                    mnuMenu2(i).caption = rs!descriclie
                    
                    rs.MoveNext
                Wend
                If mnuMenu2(0).caption <> " " Then
                    PopupMenu mnuMenu, , 1500, 1350
                End If
            'End If
        End If
    End If
End Sub

Private Sub txtPedido_LostFocus()
    If s2n(txtPedido) <> mUltimoPedido Then
        g2.Borrar
        Carga1 0, s2n(txtPedido)
        mUltimoPedido = s2n(txtPedido)
    End If
End Sub

Private Sub txtProductoCodigo_LostFocus()
    Dim desc
    
    If txtProductoCodigo = "" Then Exit Sub
    'desc = obtenerDeSQL("select descripcion from producto where activo = 1 and alias = '" & Trim(txtProductoCodigo) & "'")
    desc = obtenerDeSQL("select descripcion from producto where activo = 1 and codigo = '" & Trim(txtProductoCodigo) & "'")
    If desc = "" Then txtProductoCodigo.SetFocus
    txtProductoDescripcion = desc
    txtBarra = ""
End Sub


Private Function frmBuscarProducto()
'If cliente.codigo = 0 Then Exit Function
  
    frmBuscar.MostrarSql ("SELECT codigo AS [Codigo               ],descripcion AS [Descripcion                                                    ] FROM producto WHERE activo=1 and facturable=0")
    frmBuscarProducto = frmBuscar.resultado()
End Function

Private Sub CargoTransporteDePedido(quepedido)
    Dim i As Long
    i = s2n(obtenerDeSQL("select transporte from Pedidos_Clientes where activo = 1 and  numero = " & s2n(quepedido)))
    If i > 0 Then cmbTransporte = ObtenerDescripcion("transportes", i)
End Sub

Private Function BuscoNroYTipo() As Boolean
    Dim ss As String
    Dim maxr, maxfe, minr, minfe
    Dim punto As String
    BuscoNroYTipo = True
    
    punto = obtenerDeSQL("select punto from puntoremito where descripcion='" & Trim(cmbPunto.Text) & "'")
    BuscoNroYTipo = RevisaNroYFechaOk("RemitoVenta", "Numero", "fecha", s2n(TxtRemitoNumero, 0), dtFecha, "", False, punto)
End Function

Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim num As Long, Serie As String
    Dim i As Long, tmp, asse As String, produ As String, mstk As Long
    Dim sucursal As Long, consig As Boolean, cantConsign As Double, formula As String, depot As Long
    Dim punto As String
    
    CalculaTotal
    
    punto = obtenerDeSQL("select punto from puntoremito where descripcion='" & Trim(cmbPunto.Text) & "'")
    num = s2n(TxtRemitoNumero)
    
    If Not BuscoNroYTipo() Then
        Exit Sub
    End If
    'sin cabeza
    asse = "faltacabecera"
    If FaltaGrilla() Or txtClie.Text = "" Then 'FaltaCabecera() Or
        MsgBox "Faltan datos en el formulario"
        Exit Sub
    End If
    'prod colgado
    If HayProdEnEdicion(txtProductoDescripcion) Then Exit Sub
    'sin series
    asse = "faltaseries"
    If FaltaSeries() Then
        Exit Sub
    End If
    'ya grabado
    asse = "ya grabada"
    tmp = obtenerDeSQL("select cliente from RemitoVenta where Numero = " & s2n(TxtRemitoNumero) & " and puntoventa='" & Trim(punto) & "'")
    If Not IsEmpty(tmp) Then
        che "Numero Remito ya grabado, para cliente " & tmp
        Exit Sub
    End If
    'Controles------------
    
    asse = "obtencion suc y depot"
    sucursal = s2n(obtenerDeSQL("select sucursal from datos"))
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
        
    '*******************************************************************
    'Transaccion !
    DE_BeginTrans
    
'    GrabarBs_Num CAMPO_BS_NroREMITO, num
    '
    asse = "RV"
    'DataEnvironment1.dbo_abmRemitoVenta "A", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3)
    'If ABMRemitoVenta("A", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3), punto, Label16.caption, txtClie) = False Then GoTo ufaErr
    If ABMRemitoVenta("A", num, 0, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3), punto, Label16.caption, txtClie) = False Then GoTo ufaErr
    asse = "RVdetalle"
    For i = 1 To g.rows - 1 'items
        cantConsign = s2n(g.tx(i, gCONS))
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gCODI), Propio())
        mstk = 1
        'Abs(ManejaStock(produ))
        If ABMRVDetalle("A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk, Label16.caption) = False Then GoTo ufaErr
        'DataEnvironment1.dbo_abmRemitoVentaDetalle "A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk
    Next i
    asse = "RVseries"
    For i = 1 To g3.rows - 1 'series
        Serie = g3.tx(i, g3NSER)
        consig = (grillaSeries.cell(flexcpChecked, i, g3CONS) = flexChecked)
        If Serie <> "" Then
           'DataEnvironment1.dbo_SERIE "A", 0, VerProductoMio(g3.tx(i, g3PROD), Propio()), serie, TipoComprobante_REMITOVENTA, num, sucursal, 0, "", consig, CLng(Date), UsuarioActual(), 0, 0
            DataEnvironment1.dbo_abmSERIEs "A", 0, VerProductoMio(g3.tx(i, g3PROD), Propio()), Serie, TipoComprobante_REMITOVENTA, num, sucursal, 0, "", consig, dtFecha, 1, Date, UsuarioActual()
        End If
    Next i
    
    DE_CommitTrans
    ucMenu.AceptarOk
    
    ' quiero Transaccion, y/o quiero hacer tabla temp y 1 solo stored
    '*******************************************************************
    asse = "grabado, fallo impresion"
'    ImprimirRemitoVenta num
    MsgBox "Remito " & num & " grabado"
        
    If gEMPR_idEmpresa = 6 Or gEMPR_idEmpresa = 4 Then
        ImprimirRemitoVentaAT num, Label16.caption
    Else
        ImprimirRemitoVenta num
    End If
    
    
Exit Sub
ufaErr:
    DE_RollbackTrans
    ufa "Err al grabar: " & asse, Me.Name & " " & num    ', Err
End Sub
Private Sub ucMenu_AceptarModi() '****preparado el 16/5/07****por raul
    If ON_ERROR_HABILITADO Then On Error GoTo rv_err
    Dim num As Long, Serie As String
    Dim i As Long, asse As String, produ As String, mstk As Long
    Dim tmp
    Dim tmp2
    Dim sucursal As Long, consig As Boolean, cantConsign As Double, formula As String, depot As Long
    Dim punto As String
    Dim cod As Long
    
    CalculaTotal
    

    num = s2n(TxtRemitoNumero)
    cod = s2n(Label16.caption)
    punto = obtenerDeSQL("select punto from puntoremito where descripcion='" & Trim(cmbPunto.Text) & "'")
    
    If FaltaCabecera() Or FaltaGrilla() Then ' verifica si falta cabecera
        MsgBox "Faltan datos en el formulario"
        Exit Sub
    End If
    
    tmp2 = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where codremito = " & cod)
    If tmp2 > 0 Then
          che "No puedo anular, remito con mercaderia en transito devuelta."
          Exit Sub
    End If
    
    If HayProdEnEdicion(txtProductoDescripcion) Then Exit Sub
    
    If FaltaSeries() Then
        Exit Sub
    End If
    
    tmp = obtenerDeSQL("select numero, fecha from RemitoVenta where Numero = " & s2n(TxtRemitoNumero) & " and puntoventa='" & punto & "'")
    If Not IsEmpty(tmp) Then
        If MsgBox("Remito de Venta Numero : " & tmp(0) & " (Fecha : " & tmp(1) & "), ya existe." & Chr(13) & "¿Desea actualizar los datos existentes en el remito?.", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    sucursal = s2n(obtenerDeSQL("select sucursal from datos"))
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
        
    DE_BeginTrans
    
    'DataEnvironment1.dbo_abmRemitoVenta "M", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3)
    ABMRemitoVenta "M", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3), punto, Label16.caption, txtClie
    
    For i = 1 To g.rows - 1 'carga en remito_venta_detalle todos los items
        cantConsign = s2n(g.tx(i, gCONS))
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gCODI), Propio())
        mstk = 1
        
        ABMRVDetalle "M", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), 0, depot, cantConsign, "V", mstk, Label16.caption
        'DataEnvironment1.dbo_abmRemitoVentaDetalle "A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk
    Next i
    
   
    DE_CommitTrans
    ucMenu.AceptarOk
    
    ImprimirRemitoVentaAT num, Label16.caption
    MsgBox "Remito " & num & " guardado.", vbInformation
    

    Exit Sub
rv_err:
    DE_RollbackTrans
    ufa "Error al grabar: " & asse, Me.Name & " " & num    ', Err

    
End Sub
Private Sub ucMenu_BorrarControles()
    BorrarCampos
End Sub
Private Sub ucMenu_Buscar()
    mNuevo = False
    'frmBuscar.MostrarSql "select r.Numero,r.puntoventa [Punto], r.Cliente,c.descripcion as [ Descripcion                 ], r.Fecha as [ Fecha   ], r.Factura, r.Cancelado, r.Anulado,r.codigo  from RemitoVenta r inner join clientes c on c.codigo=r.cliente where r.fecha " & ucBetween.ssBetween & " order by r.numero desc", , , , "SI", ""
    frmBuscar.MostrarSql "select r.Numero,r.puntoventa [Punto], r.descriClie as [ Cliente                       ], r.Fecha as [ Fecha   ], r.Factura, r.Cancelado, r.Anulado,r.codigo  from RemitoVenta r where r.fecha " & ucBetween.ssBetween & " order by r.numero desc", , , , "SI", ""
    If frmBuscar.resultado() <> "" Then
        CargaRemito frmBuscar.resultado(1), frmBuscar.resultado(8)
        ucMenu.BuscarOK
        g2.Borrar
        tabRemito.Tab = 0
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ERR_ELIM
    Dim i As Long, num As Long, prod As String, cant As Double, pedi As Long, formu As String
    Dim tmp, tmpFac, tx As String, mjstk As Long
    Dim cod As Long
    
    'Controlo----------------
    ' sin numero
    If TxtRemitoNumero = "" Then Exit Sub
    num = s2n(TxtRemitoNumero)
    cod = s2n(Label16.caption)
    'ya anulado
    'If obtenerDeSQL("select Anulado from RemitoVenta where numero = " & num) = True Then
    If obtenerDeSQL("select Anulado from RemitoVenta where codigo = " & cod) = True Then
        MsgBox "Remito ya anulado"
        Exit Sub
    End If

    'ya facturado
    'tmp = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where numero = " & num)
    tmp = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where codremito = " & cod)
    If tmp > 0 Then
        'tmp = obtenerdesql ("select NroFactura from FacturaVenta
          'che "No puedo anular, remito con factura " '& vbCrLf & tmpFac(0) & " " & tmpFac(1)
          che "No puedo anular, remito con mercaderia en transito devuelta."
          Exit Sub
    End If
    
    Dim sp
    sp = obtenerDeSQL("select * from facturaventadetalle d inner join facturaventa f on d.nrofactura=f.nrofactura where f.activo=1 and d.nroremito=" & s2n(num))
    If IsNull(sp) Or IsEmpty(sp) Then
    Else
        MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
        Exit Sub
    End If
   
    If confirma("Anular remito " & TxtRemitoNumero) Then
        tx = InputBox("Motivo ")
        'If Trim(tx) = "" Then Exit Sub
       
        DE_BeginTrans
            'detalle
            ABMRemitoVenta "B", num, 0, 0, 0, 0, 0, 0, "", "", "", tx, "", Label16.caption
            For i = 1 To g.rows - 1
                prod = VerProductoMio(g.tx(i, gCODI), Propio())
                mjstk = ManejaStock(prod)
                cant = s2n(g.tx(i, gCANT))
                pedi = s2n(g.tx(i, gPEDI))
                formu = IIf(EsProductoVirtual(prod), CHAR_PROD_VIRTUAL, "")
                ABMRVDetalle "B", num, prod, cant, 0, pedi, 0, 0, formu, mjstk, Label16.caption
                'DataEnvironment1.dbo_abmRemitoVentaDetalle "B", num, prod, cant, 0, pedi, 0, 0, formu, mjstk
            Next i
            
            'DataEnvironment1.dbo_abmRemitoVenta "B", num, 0, 0, 0, 0, 0, 0, "", "", "", tx
            'Series baja en SP ' DataEnvironment1.Sistema.Execute "update series set activo = 0 where
        DE_CommitTrans
        
        MsgBox "Remito anulado.", vbInformation
        ucMenu.EliminarOK
    End If
    GoTo fin
   
ERR_ELIM:
    DE_RollbackTrans
    
    MsgBox "Error intantando anular remito venta", vbCritical
fin:
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitarEdicion sino
End Sub
Private Sub ucMenu_Imprimir()
    If gEMPR_idEmpresa = 6 Or gEMPR_idEmpresa = 4 Then
        ImprimirRemitoVentaAT2 s2n(TxtRemitoNumero), Label16.caption
    Else
        ImprimirRemitoVenta (s2n(TxtRemitoNumero))
    End If
    
'    ImprimirRemitoVenta (s2n(TxtRemitoNumero))
    ' aqui
End Sub

Private Sub ucMenu_Modificar()
    Dim tmp2
    Dim cod As Long
    
    cod = s2n(Label16.caption)
    tmp2 = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where codremito = " & cod)
    If tmp2 > 0 Then
          che "No puedo anular, remito con mercaderia en transito devuelta."
          Exit Sub
    End If
    
    mNuevo = False
    TxtRemitoNumero.Locked = True
End Sub
Private Sub ucMenu_Nuevo()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
    mNuevo = True
    TxtRemitoNumero = nuevoCodigo("remitoVenta", "numero") ' obtenerDeSQL("select max(numero ) from RemitoVenta ") + 1 ' LeerBS_Num(CAMPO_BS_NroREMITO) + 1
    Label16.caption = nuevoCodigo("remitoVenta", "codigo")
'    TxtRemitoNumero.SetFocus
    'nuevoCodigo ("RemitoVenta","Numero")
    
    '**************************
    Dim punto As String
    punto = obtenerDeSQL("select punto from puntoremito where descripcion='" & Trim(cmbPunto.Text) & "'")
    Label16.caption = s2n(obtenerDeSQL("select max(codigo) as mas from remitoventa ")) + 1
    TxtRemitoNumero.Text = s2n(obtenerDeSQL("select max(numero) from remitoventa where puntoventa='" & Trim(punto) & "'")) + 1
    '*******************************************
fin:
    Exit Sub
ufa:
    TxtRemitoNumero = "1"
    Resume fin
End Sub
Private Sub ucMenu_SALIR()
    Unload frmImagen
    Unload Me
End Sub

Public Function ABMRemitoVenta(rOPE As String, rNumero As Long, rCliente As Long, rFecha As Date, rFactura As Long, rTransporte As Long, rDEPOSITO As Long, rCodPropio As Long, rObs1 As String, rObs2 As String, rObs3 As String, rMotivo As String, punto As String, codigo As Long, Optional descri As String = "") As Boolean
On Error GoTo rvmal
Dim iudr As String
ABMRemitoVenta = True
    Select Case rOPE
        Case "A":
            iudr = "INSERT INTO REMITOVENTA(NUMERO,CLIENTE,FECHA,FACTURA,TRANSPORTE,CANCELADO,ANULADO,DEPOSITO,CODPROPIO,Obs1 , Obs2, Obs3, obs4,puntoventa,descriClie) " _
                & " VALUES( " & rNumero & "," & rCliente & "," & ssFecha(rFecha) & "," & rFactura & "," & rTransporte & ", 0, 0," & rDEPOSITO & "," & rCodPropio & "," & ssTexto(rObs1) & "," & ssTexto(rObs2) & "," & ssTexto(rObs3) & "," & ssTexto(rMotivo) & ",'" & Trim(punto) & "','" & Trim(descri) & "')"
            DataEnvironment1.Sistema.Execute iudr
' if rope = 'M'
'    begin
'        Update RemitoVenta
'        set cliente = rcliente, fecha = rfecha, factura = rfactura, transporte = rtransporte
'        where numero = rnumero
'        delete from RemitoVentaDetalle
'        where numero = rnumero
'    End

        Case "B":
            iudr = "Update SERIES  Set activo = 0   Where NroComprobante = " & rNumero & " And COMPROBANTE = 5"
            DataEnvironment1.Sistema.Execute iudr
            iudr = "Update RemitoVenta  set anulado = 1, cancelado = 1  where codigo=" & codigo
            DataEnvironment1.Sistema.Execute iudr
    End Select
Exit Function
rvmal:
ABMRemitoVenta = False

End Function

Public Function ABMRVDetalle(dOpe As String, dNumero As Long, dPRODUCTO As String, dCantidad As Double, dPrecio As Double, dPedido As Long, dDeposito As Long, dConsign As Double, dFormula As String, dManejaStock As Long, codigo As Long) As Boolean
On Error GoTo rdmal
Dim iudd As String
Dim r, rFactor As Double, rCargar As Double
Dim Valor As Long
Dim Alma As Integer

Set r = Nothing
r = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(dPRODUCTO))
If IsNull(r) Or IsEmpty(r) Then
    rFactor = 1
Else
    rFactor = r
End If
rCargar = rFactor * dCantidad

Alma = s2n(obtenerDeSQL("select almacen from producto where codigo='" & Trim(dPRODUCTO) & "'"))

ABMRVDetalle = True
Select Case dOpe
    Case "A":
        iudd = "INSERT INTO REMITOVENTADETALLE (NUMERO,PRODUCTO,CANTIDAD,FACTURAR,PEDIDO,PRECIO,CANCELADO,CONSIGNACION,FORMULA,codremito) " _
            & " values (" & dNumero & "," & ssTexto(dPRODUCTO) & "," & x2s(dCantidad) & "," & x2s(dCantidad) & "," & dPedido & "," & x2s(dPrecio) & ", 0," & x2s(dConsign) & "," & ssTexto(dFormula) & "," & codigo & ")"
        DataEnvironment1.Sistema.Execute iudd
        If dFormula <> "V" And dManejaStock = 1 Then
            If dDeposito = 0 Then
                iudd = " Update producto  Set existencia = existencia - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 1 Or Alma = 1 Then
                iudd = " Update producto  Set dep1 = dep1 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 2 Or Alma = 2 Then
                iudd = " Update producto  Set dep2 = dep2 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 3 Or Alma = 3 Then
                iudd = " Update producto  Set dep3 = dep3 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 4 Or Alma = 4 Then
                iudd = " Update producto  Set dep4 = dep4 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
        End If
    
        If dPedido > 0 Then
            iudd = " Update itemPedidoCliente  set saldo = saldo- (" & x2s(dCantidad) & ")" _
            & " Where pedido = " & dPedido & " And producto =" & ssTexto(dPRODUCTO)
            DataEnvironment1.Sistema.Execute iudd
        End If
        
    Case "M":
        Valor = s2n(obtenerDeSQL("select codigo from remitoventadetalle WHERE PRODUCTO=" & ssTexto(dPRODUCTO) & " AND NUMERO=" & dNumero & " and codremito=" & codigo))
        If Valor > 0 Then
            iudd = "UPDATE REMITOVENTADETALLE SET " _
                & " CANTIDAD=" & x2s(dCantidad) & ",FACTURAR=" & x2s(dCantidad) & ",PRECIO= " & x2s(dPrecio) & "" _
                & " WHERE PRODUCTO=" & ssTexto(dPRODUCTO) & " AND NUMERO=" & dNumero & " and codremito=" & codigo
            DataEnvironment1.Sistema.Execute iudd
        Else
            iudd = "INSERT INTO REMITOVENTADETALLE (NUMERO,PRODUCTO,CANTIDAD,FACTURAR,PEDIDO,PRECIO,CANCELADO,CONSIGNACION,FORMULA,codremito) " _
                & " values (" & dNumero & "," & ssTexto(dPRODUCTO) & "," & x2s(dCantidad) & "," & x2s(dCantidad) & "," & dPedido & "," & x2s(dPrecio) & ", 0," & x2s(dConsign) & "," & ssTexto(dFormula) & "," & codigo & ")"
            DataEnvironment1.Sistema.Execute iudd
        End If
        If dFormula <> "V" And dManejaStock = 1 Then
            If dDeposito = 0 Then
                iudd = " Update producto  Set existencia = existencia - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 1 Or Alma = 1 Then
                iudd = " Update producto  Set dep1 = dep1 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 2 Or Alma = 2 Then
                iudd = " Update producto  Set dep2 = dep2 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 3 Or Alma = 3 Then
                iudd = " Update producto  Set dep3 = dep3 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 4 Or Alma = 4 Then
                iudd = " Update producto  Set dep4 = dep4 - (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
        End If
    
        If dPedido > 0 Then
            iudd = " Update itemPedidoCliente  set saldo = saldo- (" & x2s(dCantidad) & ")" _
            & " Where pedido = " & dPedido & " And producto =" & ssTexto(dPRODUCTO)
            DataEnvironment1.Sistema.Execute iudd
        End If
    Case "B":
        'ya esta cancelado la cabecera no hace falta el detalle
        iudd = " Update remitoVentaDetalle Set cancelado = 1 " _
        & " where numero = " & dNumero & " and codremito=" & codigo & " and producto =" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute iudd
        
        If dFormula <> "V" And dManejaStock = 1 Then
            If dDeposito = 0 Then
                iudd = " Update producto Set existencia = existencia + (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 1 Or Alma = 1 Then
                iudd = " Update producto Set dep1 = dep1 + (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 2 Or Alma = 2 Then
                iudd = " Update producto Set dep2 = dep2 + (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 3 Or Alma = 3 Then
                iudd = " Update producto Set dep3 = dep3 + (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
            If dDeposito = 4 Or Alma = 4 Then
                iudd = " Update producto Set dep4 = dep4 + (" & x2s(rCargar) & ")" _
                & " Where codigo = " & ssTexto(dPRODUCTO)
                DataEnvironment1.Sistema.Execute iudd
            End If
        End If
        
        If dPedido > 0 Then
            iudd = " Update itemPedidoCliente set saldo = saldo + (" & x2s(dCantidad) & ")" _
            & " Where pedido = " & dPedido & " And producto= " & ssTexto(dPRODUCTO)
            DataEnvironment1.Sistema.Execute iudd
        End If
End Select
Exit Function
rdmal:
ABMRVDetalle = False
End Function



