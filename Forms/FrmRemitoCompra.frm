VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRemitoCompra 
   Caption         =   "Remitos compras"
   ClientHeight    =   8670
   ClientLeft      =   165
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "FrmRemitoCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera uMenu 
      Height          =   1470
      Left            =   1155
      TabIndex        =   32
      Top             =   7065
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2593
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.Frame fraCabecera 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton cmdProveedor 
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
         Left            =   2505
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1020
         Width           =   375
      End
      Begin VB.ComboBox cmbDeposito 
         Height          =   315
         ItemData        =   "FrmRemitoCompra.frx":08CA
         Left            =   7740
         List            =   "FrmRemitoCompra.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtNroRemito 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtProvCodigo 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1335
         TabIndex        =   1
         Top             =   1020
         Width           =   1185
      End
      Begin VB.ComboBox cmbProvNombre 
         Height          =   315
         Left            =   2880
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1035
         Width           =   3630
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtOrden 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   8280
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAyudaOrden 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   9540
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmRemitoCompra.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   930
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   8220
         TabIndex        =   3
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38126
      End
      Begin VB.Label lblIddoc 
         Caption         =   "0"
         Height          =   270
         Left            =   5085
         TabIndex        =   33
         Top             =   195
         Width           =   1485
      End
      Begin VB.Label Label8 
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
         Left            =   6780
         TabIndex        =   27
         Top             =   660
         Visible         =   0   'False
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
         Left            =   6960
         TabIndex        =   26
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Remito Prov:"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label1 
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
         Left            =   480
         TabIndex        =   24
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor :"
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
         TabIndex        =   23
         Top             =   1005
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "Orden de compra:"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   1020
         Width           =   1635
      End
   End
   Begin TabDlg.SSTab tabRemito 
      Height          =   5475
      Left            =   -15
      TabIndex        =   7
      Top             =   1500
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   9657
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Items"
      TabPicture(0)   =   "FrmRemitoCompra.frx":1910
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "FrmRemitoCompra.frx":192C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grilla2"
      Tab(1).Control(1)=   "cmdPedidos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Numeros de Serie"
      TabPicture(2)   =   "FrmRemitoCompra.frx":1948
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkSinSeries"
      Tab(2).Control(1)=   "cmdLlenaSerie"
      Tab(2).Control(2)=   "grillaSeries"
      Tab(2).Control(3)=   "lblHelpSerie"
      Tab(2).Control(4)=   "lblErrorSeries"
      Tab(2).ControlCount=   5
      Begin VB.CheckBox chkSinSeries 
         Alignment       =   1  'Right Justify
         Caption         =   "Sin Series"
         Height          =   315
         Left            =   -66420
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdLlenaSerie 
         Caption         =   "Llenar Serie"
         Height          =   315
         Left            =   -71460
         TabIndex        =   35
         ToolTipText     =   "Seleccione Rango de Filas a llenar"
         Top             =   480
         Width           =   1575
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   4155
         Left            =   -74760
         TabIndex        =   29
         Top             =   1080
         Width           =   9255
         _cx             =   16325
         _cy             =   7329
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
      Begin VB.Frame fraDetalle 
         BorderStyle     =   0  'None
         Height          =   4995
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   9915
         Begin VB.TextBox txtBarra 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1320
            TabIndex        =   12
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtAutorizo 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1305
            TabIndex        =   14
            Top             =   1050
            Width           =   4050
         End
         Begin VSFlex7LCtl.VSFlexGrid grilla 
            Height          =   3465
            Left            =   120
            TabIndex        =   28
            Top             =   1470
            Width           =   9555
            _cx             =   16854
            _cy             =   6112
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
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   4200
            TabIndex        =   13
            Top             =   570
            Width           =   1215
         End
         Begin VB.CommandButton cmdAyudaProducto 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   2970
            MaskColor       =   &H00E0E0E0&
            MouseIcon       =   "FrmRemitoCompra.frx":1964
            Picture         =   "FrmRemitoCompra.frx":365E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   195
            Width           =   420
         End
         Begin VB.TextBox txtProductoDescripcion 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   255
            Width           =   5070
         End
         Begin VB.TextBox txtProductoCodigo 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1320
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdAgregarItem 
            BackColor       =   &H00E0E0E0&
            Height          =   405
            Left            =   8835
            MaskColor       =   &H00E0E0E0&
            Picture         =   "FrmRemitoCompra.frx":5358
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   105
            Width           =   420
         End
         Begin VB.CommandButton cmdBorrarItem 
            Height          =   405
            Left            =   9270
            MaskColor       =   &H00E0E0E0&
            Picture         =   "FrmRemitoCompra.frx":566A
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Borrar Item"
            Top             =   105
            Width           =   420
         End
         Begin VB.Label Label10 
            Caption         =   "C.Barra :"
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
            Left            =   240
            TabIndex        =   39
            Top             =   675
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Autorizo :"
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
            Left            =   315
            TabIndex        =   31
            Top             =   1080
            Width           =   945
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00800000&
            BorderWidth     =   2
            Height          =   900
            Left            =   120
            Top             =   120
            Width           =   8655
         End
         Begin VB.Label Label7 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   3090
            TabIndex        =   19
            Top             =   645
            Width           =   1335
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
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdPedidos 
         Caption         =   "Pedidos"
         Height          =   315
         Left            =   -74760
         TabIndex        =   8
         Top             =   540
         Width           =   1155
      End
      Begin VSFlex7LCtl.VSFlexGrid grillaSeries 
         Height          =   4275
         Left            =   -74760
         TabIndex        =   34
         Top             =   960
         Width           =   9555
         _cx             =   16854
         _cy             =   7541
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
      Begin VB.Label lblHelpSerie 
         Caption         =   "Para llenar serie"
         Height          =   315
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   3075
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
         Left            =   -69720
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmRemitoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Proveedor As LiCodigo
Attribute Proveedor.VB_VarHelpID = -1
Private g As LiGrilla, g2 As LiGrilla, g3 As LiGrilla
Private mUltimaOrden As Long

Private gCODI As Long
Private gDESC As Long
Private gCANT As Long
Private gORDE As Long
Private gOCPR As Long
Private gCONS As Long
Private gOCit As Long
Private gOCca As Long
'Private gFORM As Long ' no hay formula en RC

Private g2CODI As Long
Private g2DESC As Long
Private g2CANT As Long
Private g2ORDE As Long
Private g2OCPR As Long
Private g2OCit As Long
'Private g2OCca As Long
'

Private g3ITEM As Long
Private g3PROD As Long
Private g3DESC As Long
Private g3NSER As Long
Private g3CONS As Long
Private g3HIDD As Long

Private Sub cmdAgregarItem_Click()
    On Error Resume Next
    If Not IsNumeric(txtCantidad) Or txtProductoCodigo = "" Or txtProductoDescripcion = "" Then Exit Sub
    MetoEnGrilla Trim(txtProductoCodigo), txtProductoDescripcion, s2n(txtCantidad), "", 0, "", 0
    
    txtCantidad = ""
    txtProductoCodigo = ""
    txtProductoDescripcion = ""
    txtBarra = ""
    txtProductoCodigo.SetFocus
End Sub


Private Sub cmdAyudaOrden_Click()
    Dim s As String
    
    'solo OC con items saldo >0
    s = "select distinct OrdenesDeCompras.Codigo, Fecha as [ Fecha    ], Proveedor from ItemOrdenCompra inner join OrdenesDeCompras on OrdenesDeCompras.codigo = itemOrdenCompra.ordencompra "
    If Proveedor.codigo = 0 Then
        s = s & " where activo = 1  and saldo >0 order by OrdenesDeCompras.codigo desc"
    Else
        s = s & " where activo = 1 and saldo > 0 and proveedor = " & Proveedor.codigo & " order by OrdenesDeCompras.codigo desc"
    End If
    
    With frmBuscar
        If .MostrarSql(s) = "" Then Exit Sub
        
        txtOrden = .resultado(1)
        Proveedor.codigo = .resultado(3)
        Carga1 s2n(Proveedor.codigo), s2n(txtOrden)
        
        tabRemito.Tab = 0
        txtNroRemito.SetFocus
    End With
End Sub

Private Sub cmdAyudaProducto_Click()
'    If Proveedor.Codigo = 0 Then Exit Sub
  
    If gEMPR_FormulaEsVirtual Then
        frmBuscar.MostrarSql ("select codigo as [ Codigo                    ], descripcion [ Descripcion                                                             ] from producto  where activo = 1 and formula = 0 order by codigo ")
    Else
        frmBuscar.MostrarCodigoDescripcionActivo "Producto"
    End If
    If frmBuscar.resultado = "" Then Exit Sub
    txtProductoCodigo = frmBuscar.resultado(1)
    txtProductoDescripcion = frmBuscar.resultado(2)
End Sub

Private Sub cmdBorrarItem_Click()
    On Error Resume Next
    If g.Row > 0 Then grilla.RemoveItem (g.Row)
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

Private Sub cmdPedidos_Click()
    Carga2 s2n(Proveedor.codigo), s2n(txtOrden)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

'con KeyPreView = true
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub


Private Sub Form_Load()
    Dim sBuscar As String
    
    inigrilla
    Set Proveedor = New LiCodigo
    Proveedor.init cmbProvNombre, txtProvCodigo, "Prov", False, False, cmdProveedor, "activo = 1"
    comboArray cmbDeposito, Array("Deposito Central", "Deposito 1", "Deposito 2", "Deposito 3", "Deposito4"), Array(0, 1, 2, 3, 4)
    dtFecha = Date
'    uMenu.init True, True, True, True, True ', , daTaenvironment1.Sistema
    uMenu.init True, True, True, False, True   ' No modifica aun, prefiero q no este el boton
    lblHelpSerie.caption = ""
    tabRemito.TabVisible(2) = gEMPR_Maneja_series
End Sub
Private Sub Form_Resize()
'    encajar uMenu, Me,  ,  120, 120
    encajar tabRemito, Me, 1920, 60, 100 + uMenu.Height + 120, 120
    encajar fraDetalle, tabRemito, 360, 100, 100, 100
    encajar grilla, fraDetalle, 1740, 50, 50, 60
    encajar grilla2, tabRemito, 1080, 120, 120, 240
    encajar grillaSeries, tabRemito, 1080, 120, 120, 240
End Sub
Private Sub Form_Terminate()
    Set g2 = Nothing
    Set g = Nothing
    Set g3 = Nothing
End Sub

Private Sub grilla2_DblClick()
    Dim r   As Long
    
    r = g2.Row
    If r > 0 Then
        MetoEnGrilla _
              g2.tx(r, g2CODI) _
            , g2.tx(r, g2DESC) _
            , g2.tx(r, g2CANT) _
            , g2.tx(r, g2ORDE), 0 _
            , g2.tx(r, g2OCPR) _
            , g2.tx(r, g2OCit)
        g2.delRow (r)
    End If
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
    Dim i As Long
    Dim Col As Long
    Dim Row As Long
    
    Row = grillaSeries.RowSel
    Col = grillaSeries.ColSel
    If gEMPR_Sucursal = 1 Then
        If Col = g3NSER Then
            'MsgBox "estoy en col: " & col & " y lin: " & row 'esto es solo para mi!!!
            If grillaSeries.TextMatrix(Row, 1) <> "" Then
                frmBuscar.MostrarSql ("select c.nroserie as [ Nro de Serie                 ], p.descripcion [ Descripcion                                                             ],p.codigobarra as [ Codigo de barra de Producto] from producto p inner join codigobarras c on p.codigobarra=c.nroproducto where p.codigo='" & grillaSeries.TextMatrix(Row, 1) & "' and utilizado=0 and c.activo = 1 and p.formula = 0 order by c.nroserie ")
            Else
                MsgBox "No hay producto seleccionado."
                Exit Sub
            End If
            
            If frmBuscar.resultado = "" Then Exit Sub
            grillaSeries.TextMatrix(Row, g3NSER) = frmBuscar.resultado(1)
            
            For i = 1 To g3.rows - 1
                If (i <> Row) And (g3.tx(i, g3PROD) = g3.tx(Row, g3PROD)) And (Trim(g3.tx(i, g3NSER)) <> "") Then
                    If g3.tx(i, g3NSER) = g3.tx(Row, g3NSER) Then
                        che " Series de filas " & i & " y " & Row & " iguales"
                    End If
                End If
            Next i
        End If
    End If
End Sub

Private Sub proveedor_cambio(codigo As Variant)
    Dim np As Long, no As Long
    np = Proveedor.codigo
    no = s2n(txtOrden)
    If np = 0 Then
        g.Borrar
        g2.Borrar
    Else
        If no > 0 Then
            g2.Borrar
            Carga1 np, no
        Else
            g.Borrar
            Carga2 np, 0
        End If
    End If
End Sub

Private Sub CargaOrden(cual)
    If ON_ERROR_HABILITADO Then On Error GoTo E_UFA
    
    Dim rs As New ADODB.Recordset, i As Long, ocProv As String
    
    With rs
        .Open "select * from ordenesdecompras where numero = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
 '       TxtRemitoNumero = !numero
        txtOrden = !numero
       
        Proveedor.codigo = !provededor
        dtFecha = !Fecha
        ocProv = !ordenproveedor
        .Close
    
        .Open "select * from ItemOrdenCompra where numero = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        g.rows = 1
        While Not .EOF
            i = g.addRow()
            g.tx i, gCODI, !producto
            g.tx i, gDESC, ObtenerDescripcionS("producto", !producto)
            g.tx i, gCANT, !cantidad
            g.tx i, gORDE, !ordenCompra
            g.tx i, gOCPR, ocProv
            g.tx i, gOCit, !itemOcDetalle
            .MoveNext
        Wend
    End With
    
E_UFA:
    Set rs = Nothing
End Sub


Private Function FaltaCabecera() As Boolean
    FaltaCabecera = True
    If Proveedor.codigo = 0 Then
        che "Falta proveedor"
        txtProvCodigo.SetFocus
        Exit Function
    End If
    If Trim(txtNroRemito) = "" Then
        che "Falta Nro Remito Proveedor"
        txtNroRemito.SetFocus
        Exit Function
    End If
    FaltaCabecera = False
End Function


Private Function FaltaGrilla() As Boolean
    Dim i As Long, x As Double
    
    x = (g.suma(gCANT) = 0)
    If x Then
        che "no hay cantidades en los items"
        FaltaGrilla = True
    End If
    For i = 1 To g.rows - 1
        If s2n(g.tx(i, gOCit)) > 0 Then
            If s2n(g.tx(i, gCANT)) > s2n(g.tx(i, gOCca)) And Trim(txtAutorizo) = "" Then
                txtAutorizo.enabled = True
                che "Se ingreso una cantidad mayor al pedido" & vbCrLf & "corregir o  registrar autorizacion"
                FaltaGrilla = True
                Exit Function
            End If
        End If
    Next i
    
End Function

Private Sub Carga1(nProveedor, nOrden)
    If ON_ERROR_HABILITADO Then On Error GoTo ErrCarga1
    Dim rs As New ADODB.Recordset, ssql As String, i As Long

    If nProveedor + nOrden = 0 Then Exit Sub
    
    g.rows = 1
    ssql = "select proveedor, producto, saldo, ordencompra, OrdenProveedor, Itemordencompra.codigo as itOcDetalle from Itemordencompra inner join ordenesdecompras on ordenesdecompras.codigo = Itemordencompra.ordencompra where activo = 1 and saldo >0 "
    
    If nProveedor > 0 Then ssql = ssql & " and proveedor = " & nProveedor
    If nOrden > 0 Then ssql = ssql & " and ordencompra = " & nOrden
    
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        If Not .EOF Then Proveedor.codigo = !Proveedor
        g.Borrar
        While Not .EOF
            i = g.addRow()

            g.tx i, gCODI, !producto
            g.tx i, gDESC, ObtenerDescripcionS("producto", !producto)
            g.tx i, gCANT, !saldo
            g.tx i, gORDE, !ordenCompra
            g.tx i, gOCPR, !ordenproveedor
            g.tx i, gOCit, !itOcDetalle
            g.tx i, gOCca, !saldo
            .MoveNext
        Wend
    End With
    tabRemito.Tab = 0

fin:
    Set rs = Nothing
    Exit Sub
ErrCarga1:
    ufa "", " carga1 " & Me.Name ', Err
    Resume fin
End Sub

Private Sub Carga2(nProveedor, nOrden)
    Dim rs As New ADODB.Recordset, ssql As String, i As Long

    g2.rows = 1
    
    ssql = "select producto, saldo, ordencompra, OrdenProveedor,Itemordencompra.codigo as itCodigo from Itemordencompra inner join ordenesdecompras on ordenesdecompras.codigo = Itemordencompra.ordencompra where proveedor = " & nProveedor & " and activo = 1 and saldo >0 "
    If nOrden > 0 Then ssql = ssql & " and ordencompra = " & nOrden
    ssql = ssql & " order by ordencompra, producto "
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            i = g2.addRow()
            g2.tx i, g2CODI, !producto
            g2.tx i, g2DESC, ObtenerDescripcionS("producto", !producto)
            g2.tx i, g2CANT, !saldo
            g2.tx i, g2ORDE, !ordenCompra
            g2.tx i, g2OCPR, !ordenproveedor
            g2.tx i, g2OCit, !itCodigo
            'g2.tx i,gocca,
            .MoveNext
        Wend
    End With
    tabRemito.Tab = 1
    Set rs = Nothing
End Sub

Private Sub BorrarCampos()
    FrmBorrarTxt Me
    FrmBorrarCbo Me
    Proveedor.codigo = 0
    mUltimaOrden = 0

    g.Borrar
    g2.Borrar
    g3.Borrar
End Sub

Private Sub HabilitarEdicion(habilitar As Boolean)
    fraCabecera.enabled = habilitar
'    tabRemito.enabled = habilitar

    fraDetalle.enabled = habilitar
    
'    grilla2.Enabled = habilitar
    grilla2.Editable = IIf(habilitar, flexEDKbdMouse, flexEDNone)
    cmdPedidos.enabled = habilitar
    
    grillaSeries.Editable = IIf(habilitar, flexEDKbdMouse, flexEDNone)
    cmdLlenaSerie.enabled = habilitar
    chkSinSeries.enabled = habilitar
    
End Sub

Private Sub Resetear()
    BorrarCampos
    HabilitarEdicion False
    tabRemito.Tab = 0
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    Set g2 = New LiGrilla
    Set g3 = New LiGrilla
   
    g.init grilla
    g2.init grilla2
    g3.init grillaSeries
    
    gCODI = g.AddCol("  Codigo               ") ', "S")
    gDESC = g.AddCol("  Descripcion                                                               ")
    gCANT = g.AddCol(" Cantidad ", "N")
    gORDE = g.AddCol(" O/C      ")
    gOCPR = g.AddCol(" O/C Prov ")
    gCONS = g.AddCol(" C.Consgn ", "N")
'    gFORM = g.addCol(" Producto Base ")
    gOCca = g.AddCol("cantidadEnOc") ', "H")
    gOCit = g.AddCol("codigoItemOC") ', "H")
    g.Borrar
    
    g2CODI = g2.AddCol("  Codigo             ")
    g2DESC = g2.AddCol(" Descripcion                                                              ")
    g2CANT = g2.AddCol(" Cantidad Pendiente", "9")
    g2ORDE = g2.AddCol(" O/C     ")
    g2OCPR = g2.AddCol(" O/C Prov ")
    g2OCit = g2.AddCol("codigoItemOC") ', "H")
    
    g2.Borrar
    
    g3ITEM = g3.AddCol(" Item ")
    g3PROD = g3.AddCol(" Producto                  ")
    g3DESC = g3.AddCol(" Descripcion                                                            ")
    g3NSER = g3.AddCol(" Numero de Serie           ", "S") ' editable
    g3CONS = g3.AddCol("Consignacion", "K")
    g3HIDD = g3.AddCol("h", "H")
    g3.Borrar
    grillaSeries.SelectionMode = flexSelectionListBox

End Sub

Private Sub MetoEnGrilla(codi, desc, cant, Orden, cons, ocProv, itOcDetalle)
'    Dim i As Long, rs As New ADODB.Recordset, sSql As String, CodigoMio As String
    Dim i As Long
'    ssql = "select codigo, componente, cantidad from formulas where activo = 1 and codigo = '" & CodigoMio & "'"
    i = g.addRow()
    g.tx i, gCODI, codi
    g.tx i, gDESC, desc
    g.tx i, gCANT, cant
    g.tx i, gORDE, Orden
    g.tx i, gOCPR, ocProv
    g.tx i, gCONS, cons
    g.tx i, gOCca, cant
    g.tx i, gOCit, itOcDetalle

    MetoGrillaSeries CStr(codi), cant, cons, desc
End Sub

Private Sub MetoGrillaSeries(prod As String, ByVal cant As Long, ByVal consig As Long, ByVal descri As String)
    On Error GoTo ERR_FIN
    Dim i As Long, r As Long
    Dim rs As New ADODB.Recordset
    Dim Serie As String
    Dim error As Long
    
    error = 0
                
    If Not ProductoConSerie(prod) Then Exit Sub
                
'    If obtenerDato("producto", prod, "serie") Then
    'asserts saludables
    If g3.rows > 99 Then
        ufa " Demasiados items para num de serie ", "Remito Compra" ', Err
        Exit Sub
    End If
    If cant < 1 Then
        ufa "", "Cantidad para num serie < 1  Remito Compra" ', Err
        Exit Sub
    End If
    '
    
    For i = 1 To cant
        
        If gEMPR_Sucursal = 1 Then
            rs.Open "select p.codigobarra,p.descripcion,min(c.nroserie) as serie from producto p inner join codigobarras c on p.codigobarra=c.nroproducto where p.codigo='" & prod & "' and utilizado=0 and c.activo=1 and c.nroserie>'" & Serie & "' group by p.codigobarra,p.descripcion", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            
            If rs.EOF = True And rs.BOF = True Then
            Else
                Serie = rs!Serie
            End If
        End If
        
        r = g3.addRow()
        grillaSeries.TextMatrix(r, g3PROD) = prod
        g3.tx r, g3DESC, descri
        
        If gEMPR_Sucursal = 1 Then
            If rs.EOF = True And rs.BOF = True Then
                'MsgBox "No se ha encontrado ningun nro de serie sin usar para el producto " & descri & "."
                error = 1
            Else
                If IsNull(rs!Serie) Then
                    'MsgBox "No se ha encontrado ningun nro de serie sin usar para el producto " & descri & "."
                    error = 1
                Else
                    grillaSeries.TextMatrix(r, g3NSER) = rs!Serie
                End If
            End If
            Set rs = Nothing
        End If
                
        If consig > 0 Then
            grillaSeries.TextMatrix(r, g3CONS) = flexChecked
            consig = consig - 1
        End If
    Next i
'    End If

    If error = 1 Then
        MsgBox "No se ha encontrado nro de serie sin usar para algun producto." & Chr(13) & Chr(13) & "Verifique antes de aceptar la carga"
    End If
    
    GoTo fin
ERR_FIN:
    ufa "err en series ", "Remito compra" ', Err
fin:
End Sub


'''''Private Sub MetoGrillaSeries(prod As String, ByVal cant As Integer, ByVal consig As Integer, ByVal descri As String)
'''''    If ON_ERROR_HABILITADO Then On Error GoTo ERR_FIN
'''''    Dim i As Long, r As Long
'''''
'''''    If Not ProductoConSerie(prod) Then Exit Sub
'''''
'''''    If g3.rows > 99 Then
'''''        ufa " Demasiados items para num de serie ", "Remito Compra" ', Err
'''''        Exit Sub
'''''    End If
'''''    If cant < 1 Then
'''''        ufa "", "Cantidad para num serie < 1  Remito Compra" ', Err
'''''        Exit Sub
'''''    End If
'''''    '
'''''
'''''    For i = 1 To cant
'''''        r = g3.addRow()
'''''        grillaSeries.TextMatrix(r, g3PROD) = prod
'''''        g3.tx r, g3DESC, descri
'''''        If consig > 0 Then
'''''            grillaSeries.TextMatrix(r, g3CONS) = flexChecked
'''''            consig = consig - 1
'''''        End If
'''''    Next i
'''''
'''''
'''''    GoTo fin
'''''ERR_FIN:
'''''    ufa "err en series ", "Remito compra" ', Err
'''''fin:
'''''End Sub
'

Private Function GrabaRemito(Ope As String) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim codi As Long, i As Long
    Dim prod As String, cant As Double, depo As Long, orde As Long, cons As Double
    Dim Serie As String
    Dim consig As Boolean, sucursal, depot As Long, itOc As Long
    Dim asse 'assert
    Dim Aux As New ADODB.Recordset
    'codigo autonumerico
    
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    
    ' quiero Transaccion, y/o quiero hacer tabla temp y 1 solo stored
    DE_BeginTrans
    
        
    
    asse = "cabecera"
    'RC Cabecera
    'If Ope = "A" Then
    '    DataEnvironment1.dbo_abmRemitoCompra OPE, 0, Proveedor.codigo, dtFecha, txtNroRemito, depot, Date, UsuarioActual(), Trim(txtAutorizo)
    'Else
    '    If OPE = "M" Then
    '        DataEnvironment1.dbo_abmRemitoCompra OPE, txtCodigo, Proveedor.codigo, dtFecha, txtNroRemito, depot, Date, UsuarioActual(), Trim(txtAutorizo)
    '    End If
    'End If
    
    
    
    ABMRemitoCompra Ope, s2n(txtCodigo), Proveedor.codigo, dtFecha, txtNroRemito, depot, UsuarioActual(), Trim(txtAutorizo)
    
    codi = obtenerDeSQL("select codigo from RemitoCompra where activo=1 and NroRemito = '" & txtNroRemito & "' and proveedor = " & Proveedor.codigo)
    If codi = 0 Then
        MsgBox "Error al grabar el remiro.", vbCritical
        Exit Function
    End If
    txtCodigo = codi
    asse = "detalle"
    'RC Detalle
    'depo = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    DataEnvironment1.Sistema.Execute "DELETE FROM RemitoCompraDetalle WHERE CodigoRemito=" & s2n(txtCodigo)
    For i = 1 To g.rows - 1 ' y si verifico vacios aca ?
        prod = g.tx(i, gCODI)
        cant = s2n(g.tx(i, gCANT))
        cons = s2n(g.tx(i, gCONS))
        orde = s2n(g.tx(i, gORDE)) 's2n(txtOrden)
        itOc = s2n(g.tx(i, gOCit))
        'DataEnvironment1.dbo_abmRemitoCompraDetalle codi, depot, prod, CANT, orde, CANT, cons, s2n(itOc)
        ABMRCDetalle codi, depot, prod, cant, orde, cant, cons, s2n(itOc)
    Next i

    sucursal = nSinNull(obtenerDeSQL("select sucursal from datos"))
    
    For i = 1 To g3.rows - 1 'series
        Serie = g3.tx(i, g3NSER)
        consig = (grillaSeries.cell(flexcpChecked, i, g3CONS) = flexChecked)
        
        If Serie <> "" Then
            prod = g3.tx(i, g3PROD)
            'DataEnvironment1.dbo_SERIE Ope, 0, prod, serie, TipoComprobante_REMITOCOMPRA, codi, sucursal, 0, "", consig, CLng(Date), UsuarioActual(), 0, 0
            DataEnvironment1.dbo_abmSERIEs Ope, 0, prod, Serie, TipoComprobante_REMITOCOMPRA, codi, sucursal, 0, "", consig, dtFecha, False, Date, UsuarioActual()
            Aux.Open "select * from producto where codigo = '" & prod & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If Aux.EOF And Aux.BOF Then
            Else
                If gEMPR_idEmpresa = 6 Or gEMPR_idEmpresa = 11 Or gEMPR_idEmpresa = 3 Or gEMPR_idEmpresa = 10 Then
                Else
                    Aux.MoveFirst
                
                    DataEnvironment1.Sistema.Execute "update codigobarras set utilizado=1 where nroproducto='" & Aux!CODIGOBARRA & "' and nroserie='" & Serie & "'"
                End If
            End If
            Set Aux = Nothing
        End If
    Next i

'''''    For i = 1 To g3.rows - 1 'series
'''''        Serie = g3.tx(i, g3NSER)
'''''        consig = (grillaSeries.cell(flexcpChecked, i, g3CONS) = flexChecked)
'''''
'''''        If Serie <> "" Then
'''''            prod = g3.tx(i, g3PROD)
'''''            'DataEnvironment1.dbo_SERIE Ope, 0, prod, serie, TipoComprobante_REMITOCOMPRA, codi, sucursal, 0, "", consig, CLng(Date), UsuarioActual(), 0, 0
'''''            DataEnvironment1.dbo_abmSERIEs Ope, 0, prod, Serie, TipoComprobante_REMITOCOMPRA, codi, sucursal, 0, "", consig, dtFecha, False, Date, UsuarioActual()
'''''        End If
'''''    Next i
    
    DE_CommitTrans
    
    If Ope = "A" Then
        MsgBox "Remito Compra " & codi & " guardado."
    Else
        If Ope = "M" Then
            MsgBox "Remito Compra " & codi & " actualizado."
        End If
    End If
    
    GrabaRemito = True
    
    
fin:
    Exit Function
UfaGraba:
    DE_RollbackTrans
    ufa "Error al grabar", "Graba RemitoCompra - " & asse ', Err
    MsgBox "Verifique si fueron bien cargados los numero de serie.", vbInformation
    Resume fin
End Function

Private Sub CargaRemito(cual)
    If ON_ERROR_HABILITADO Then On Error GoTo E_UFA
    Dim rs As New ADODB.Recordset, i As Long, desc As Variant
    
    With rs
        .Open "select * from remitocompra where codigo = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        txtCodigo = !codigo
        txtNroRemito = !NroRemito
        
        Proveedor.codigo = !Proveedor

        txtOrden = s2n(!ordenCompra)
        dtFecha = !Fecha
        cmbDeposito.ListIndex = s2n(!DEPOSITO)
        txtAutorizo = !AUTORIZO
    
        .Close
    
        .Open "select * from RemitoCompraDetalle where codigoRemito = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        .MoveFirst
        g.rows = 1
        While Not .EOF
            i = g.addRow()

            desc = obtenerDeSQL("select descripcion from producto where codigo = '" & !producto & "' and activo = 1 ")
            If IsNull(desc) Then desc = ""
            
            grilla.TextMatrix(i, gCODI) = !producto
            grilla.TextMatrix(i, gDESC) = desc
            grilla.TextMatrix(i, gCANT) = !cantidad
            grilla.TextMatrix(i, gORDE) = !ordenCompra
            g.tx i, gOCPR, obtenerDato("OrdenesDeCompras", !ordenCompra, "OrdenProveedor")
            g.tx i, gOCit, !itemOcDetalle
            .MoveNext
        Wend
        .Close
        .Open "select producto, serie, consignacion from series where nroComprobante = " & cual & " and Comprobante = " & TipoComprobante_REMITOCOMPRA, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        g3.Borrar
        While Not .EOF
            i = g3.addRow()
            grillaSeries.TextMatrix(i, g3PROD) = !producto ' VerProductoCliente(!producto)
            grillaSeries.TextMatrix(i, g3DESC) = obtenerDeSQL("select descripcion from producto where codigo = '" & !producto & "' and activo = 1 ")
            grillaSeries.TextMatrix(i, g3NSER) = !Serie
            grillaSeries.TextMatrix(i, g3CONS) = !consignacion
            .MoveNext
        Wend
        .Close
    End With
    tabRemito.Tab = 0
    GoTo fin

E_UFA:
    Set rs = Nothing
    ufa "Error al cargar remito", Me.Name ', Err
fin:
    Set rs = Nothing
End Sub


Private Sub tabRemito_Click(PreviousTab As Integer)
    If tabRemito.Tab = 2 And PreviousTab <> 2 Then
        LlenoGrillaSeries
    End If
End Sub

Private Sub LlenoGrillaSeries()
    Dim i As Long, j As Long, prod As String, cant As Long, cons, descri As String

        
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
            descri = (g.tx(i, gDESC))
            If ProductoConSerie(prod) Then
                For j = 1 To cant
                    If marcoG3(prod, cons, descri) Then
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

Private Sub txtBarra_LostFocus()
    On Error Resume Next
    
    If txtBarra = "" Then Exit Sub
    txtProductoCodigo = sSinNull(obtenerDeSQL("select codigo from producto where activo = 1 and codigobarra = '" & Trim(txtBarra) & "' "))
        
    If txtProductoCodigo = "" Then
        MsgBox "El codigo de barras no existe para ningun producto.", , "ATENCION"
        txtProductoDescripcion = ""
        txtBarra = ""
        txtProductoCodigo.SetFocus
        Exit Sub
    End If
    txtProductoDescripcion = sSinNull(obtenerDeSQL("select descripcion from producto where activo = 1 and codigo = '" & Trim(txtProductoCodigo) & "' "))
    If txtProductoDescripcion = "" Then
        txtProductoCodigo = ""
        txtProductoCodigo.SetFocus
    Else
        txtCantidad.SetFocus
    End If
End Sub

Private Sub txtOrden_LostFocus()
    If s2n(txtOrden) <> mUltimaOrden Then
        g2.Borrar
        Carga1 0, s2n(txtOrden)
        mUltimaOrden = s2n(txtOrden)
    End If
End Sub

Private Sub txtProductoCodigo_LostFocus()
    On Error Resume Next
    If txtProductoCodigo = "" Then Exit Sub
    txtProductoDescripcion = sSinNull(obtenerDeSQL("select descripcion from producto where activo = 1 and codigo = '" & Trim(txtProductoCodigo) & "' "))
    If txtProductoDescripcion = "" Then
        txtProductoCodigo = ""
        txtProductoCodigo.SetFocus
    Else
        txtCantidad.SetFocus
    End If
End Sub

' **********************************************
Private Sub uMenu_AceptarAlta()
    Dim tmp
    If HayProdEnEdicion(txtProductoDescripcion) Then Exit Sub
    If FaltaCabecera() Or FaltaSeries() Then Exit Sub
    If FaltaGrilla() Then Exit Sub
    
    tmp = obtenerDeSQL("select codigo, fecha from RemitoCompra where activo=1 and NroRemito = '" & txtNroRemito & "' and Proveedor = " & Proveedor.codigo)
    If Not IsEmpty(tmp) Then
        If tmp(1) > 0 Then
            lMsg Array("Remito " & tmp(1), txtNroRemito & " " & tmp(2) & " prov " & Proveedor.codigo, "existe con fecha " & tmp(1) & tmp(2))
            Exit Sub
        End If
    End If
    If GrabaRemito("A") Then uMenu.AceptarOk
End Sub

Private Sub uMenu_AceptarModi()
    Dim tmp
    If HayProdEnEdicion(txtProductoDescripcion) Then Exit Sub
    If FaltaCabecera() Or FaltaSeries() Then Exit Sub
    If FaltaGrilla() Then Exit Sub
    
    tmp = obtenerDeSQL("select codigo, fecha from RemitoCompra where activo=1 and NroRemito = '" & txtNroRemito & "' and Proveedor = " & Proveedor.codigo)
    If Not IsEmpty(tmp) Then
        If tmp(1) > 0 Then
            If MsgBox("Remito de Compra Nro : " & txtNroRemito & ", ya cargado  (Fecha : " & tmp(1) & ")." & Chr(13) & "¿Desea actualizar los datos del remito existente?.", vbYesNo) = vbNo Then
                'lMsg Array("Remito " & tmp(1), txtNroRemito & " " & tmp(2) & " Proveedor " & Proveedor.codigo, "ya existe con fecha " & tmp(1) & tmp(2))
                'muy pero muy mal esto
                Exit Sub
            End If
        End If
    End If
    If GrabaRemito("M") Then uMenu.AceptarOk
    
End Sub

Private Sub uMenu_BorrarControles()
    Resetear
End Sub
Private Sub uMenu_Buscar()
    Dim re As String
    re = frmBuscar.MostrarSql("select codigo, nroremito,proveedor, fecha  from RemitoCompra where activo = 1order by codigo desc")
    If re = "" Then Exit Sub
    CargaRemito frmBuscar.resultado()
    g2.Borrar
    uMenu.BuscarOK
End Sub
Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    
    Dim depot As Long, prod As String, cant As Double, orde As Long, codi As Long, i As Long, itOc As Long
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    Dim a
    
    If txtCodigo = "" Then Exit Sub
    
    a = obtenerDeSQL("select sum(d.cantidad),sum(d.cantidad- d.cantidad_a_facturar) from remitocompradetalle d inner join remitocompra r on d.codigoremito=r.codigo where r.codigo=" & s2n(txtCodigo))
    If a(0) = a(1) Then
        MsgBox "El remito fue utilizado, no se puede eliminar.", vbCritical
        Exit Sub
    ElseIf a(1) = 0 Then
    ElseIf a(1) < a(0) Then
        MsgBox "El remito esta siendo utilizado, no se puede eliminar.", vbCritical
        Exit Sub
    End If
    
    If confirma("Borrar remito " & txtNroRemito) Then
        codi = s2n(txtCodigo)
        
        DE_BeginTrans
            'baja detalle; cant , deposito, oc
            ABMRemitoCompra "B", s2n(txtCodigo), 0, Date, 0, 0, 0, ""
            For i = 1 To g.rows - 1
                prod = g.tx(i, gCODI)
                cant = s2n(g.tx(i, gCANT))
    '            cons = s2n(g.tx(i, gCONS))
                orde = s2n(g.tx(i, gORDE)) 's2n(txtOrden)
                itOc = s2n(g.tx(i, gOCit))
                ABMRCDetalle codi, depot, prod, -cant, orde, -cant, 0, itOc     ' cons
            Next i
            
            'series
            DataEnvironment1.Sistema.Execute "delete from series where comprobante = " & TipoComprobante_REMITOCOMPRA & " and NroComprobante = " & codi
            
            
        DE_CommitTrans
        
        MsgBox "eliminado"
        uMenu.EliminarOK
    End If
    
    
    GoTo fin
UFAelim:
    DE_RollbackTrans
    ufa "err eliminando.", Me.Name & " cod = " & txtCodigo ', Err
fin:
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitarEdicion sino
End Sub

Private Sub uMenu_Modificar()
'    tabRemito.enabled = False
estadoModificado (True)
End Sub

Public Sub estadoModificado(Valor As Boolean) 'agregado de urgencia 17/5/07
    cmbProvNombre.Locked = Valor
    cmdProveedor.enabled = Not Valor '
    txtProvCodigo.Locked = Valor
    txtOrden.Locked = Valor
    txtNroRemito.Locked = Valor
End Sub

Private Sub uMenu_Nuevo()
    estadoModificado (False) 'agregado de urgencia 17/5/07
    'txtNroRemito.SetFocus
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub

Public Function ABMRemitoCompra(rOPE As String, rCOD As Long, rPROV As Long, rFECHREMITO As Date, rREMITO As String, rDEPOSITO As Long, rUSUARIO As Long, rAUTORIZO As String) As Boolean
On Error GoTo rmal
ABMRemitoCompra = True
Dim iud As String

Select Case rOPE
    Case "A":
        iud = "INSERT INTO RemitoCompra (Proveedor , fecha, NroRemito, DEPOSITO, AUTORIZO,FECHA_ALTA, USUARIO_ALTA, ACTIVO) " _
                & " VALUES (" & rPROV & "," & ssFecha(rFECHREMITO) & "," & ssTexto(rREMITO) & "," & rDEPOSITO & "," & ssTexto(rAUTORIZO) & "," & ssFecha(Date) & "," & rUSUARIO & ", 1)"
        DataEnvironment1.Sistema.Execute iud
    Case "M":
        iud = "UPDATE RemitoCompra SET " _
            & "(Proveedor=" & rPROV & " , fecha=" & ssFecha(rFECHREMITO) & ", NroRemito=" & ssTexto(rREMITO) & ", DEPOSITO=" & rDEPOSITO & ", AUTORIZO=" & ssTexto(rAUTORIZO) & ") "
        DataEnvironment1.Sistema.Execute iud
    Case "B":
        iud = " Update RemitoCompra  " _
                & " SET ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA=" & rUSUARIO _
                & " WHERE CODIGO=" & rCOD
        DataEnvironment1.Sistema.Execute iud
        iud = "  DELETE FROM RemitoCompraDetalle  WHERE codigoRemito=" & rCOD
        DataEnvironment1.Sistema.Execute iud
End Select
Exit Function
rmal:
ABMRemitoCompra = False
End Function

Public Function ABMRCDetalle(dREMITO As Long, dDEP As Long, dPRODUCTO As String, dCANT As Double, dORDEN As Long, dCANTFAC As Double, dCONSIG As Double, ditemOcDetalle As Long) As Boolean
On Error GoTo dmal
ABMRCDetalle = True
Dim ID As String, e, pFactor As Double, pCargar As Double
Dim Alma As Integer
    
    ID = "INSERT INTO RemitoCompraDetalle (CodigoRemito, PRODUCTO, CANTIDAD, CANTIDAD_A_FACTURAR,ORDENCOMPRA,CANTIDAD_CONSIGNADA, itemOcDetalle) " _
         & " VALUES (" & dREMITO & "," & ssTexto(dPRODUCTO) & "," & x2s(dCANT) & "," & x2s(dCANTFAC) & "," & dORDEN & "," & x2s(dCONSIG) & "," & ditemOcDetalle & ")"
    DataEnvironment1.Sistema.Execute ID
    
    If ditemOcDetalle > 0 Then
        ID = " Update ItemOrdenCompra " _
            & " SET saldo=saldo-(" & x2s(dCANT) & ")" _
            & " where codigo = " & ditemOcDetalle
        DataEnvironment1.Sistema.Execute ID
    End If
    
    e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(dPRODUCTO))
    If IsNull(e) Or IsEmpty(e) Then
        pFactor = 0
    Else
        pFactor = e
    End If
    pCargar = pFactor * dCANT
    
    Alma = s2n(obtenerDeSQL("Select almacen from producto where codigo='" & Trim(dPRODUCTO) & "'"))
    
    If dDEP = 0 Then
        ID = "UPDATE PRODUCTO SET EXISTENCIA=EXISTENCIA+" & x2s(pCargar) & " WHERE CODIGO=" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute ID
    ElseIf dDEP = 1 Or Alma = 1 Then
        ID = "UPDATE PRODUCTO SET DEP1=DEP1+" & x2s(pCargar) & " WHERE CODIGO=" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute ID
    ElseIf dDEP = 2 Or Alma = 2 Then
        ID = "UPDATE PRODUCTO SET DEP2=DEP2+" & x2s(pCargar) & " WHERE CODIGO=" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute ID
    ElseIf dDEP = 3 Or Alma = 3 Then
        ID = "UPDATE PRODUCTO SET DEP3=DEP3+" & x2s(pCargar) & "WHERE CODIGO=" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute ID
    ElseIf dDEP = 4 Or Alma = 4 Then
        ID = "UPDATE PRODUCTO SET DEP4=DEP4+" & x2s(pCargar) & "WHERE CODIGO=" & ssTexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute ID
    End If
Exit Function
dmal:
ABMRCDetalle = False
End Function

Private Function marcoG3(codi, ByVal cons, ByVal descri As String)
    Dim i As Long
    
    For i = 1 To g3.rows - 1
        If g3.tx(i, g3PROD) = codi And g3.tx(i, g3HIDD) = "" Then
            grillaSeries.TextMatrix(i, g3HIDD) = "X"
            Exit Function
        End If
    Next i
    i = g3.addRow()
    grillaSeries.TextMatrix(i, g3PROD) = codi
    g3.tx i, g3DESC, descri
    grillaSeries.TextMatrix(i, g3HIDD) = "X"
    grillaSeries.TextMatrix(i, g3CONS) = IIf(cons > 0, "-1", "0")
    
    marcoG3 = (g3.tx(i, g3CONS) = "-1")
End Function

Private Function FaltaSeries() As Boolean
    Dim r As Long, i As Long, ns As String, j As Long, pr As String
    
    FaltaSeries = False
    lblErrorSeries.Visible = False
    
    LlenoGrillaSeries
    r = g3.rows
    If r > 1 And g3.buscar(g3NSER, "") > 0 And chkSinSeries.Value <> 1 Then
        tabRemito.Tab = 2
        grillaSeries.SetFocus
        grillaSeries.Select g3.PrimerVacio(g3NSER), g3NSER
        lblErrorSeries.Visible = True
        lblErrorSeries.caption = " Faltan Series"
        FaltaSeries = True
        Exit Function
    End If

    If r > 1 Then
        For i = 1 To r - 2
            ns = g3.tx(i, g3NSER)
            pr = g3.tx(i, g3PROD)
            If ns <> "" Then
                'antes buscaba NS repetidos, ahora solo si son del mismo producto
                For j = i + 1 To r - 1 '
                    If i <> j And pr = g3.tx(j, g3PROD) And ns = g3.tx(j, g3NSER) Then
                        tabRemito.Tab = 2
                        grillaSeries.SetFocus
                        grillaSeries.Select i, g3NSER, j, g3NSER
                        'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                        lblErrorSeries.Visible = True
                        lblErrorSeries.caption = "Serie Repetida: " & i & " y " & g3.buscar(g3NSER, ns, i + 1)
                        FaltaSeries = True
                        Exit Function
                    End If
                Next j
            End If
        Next i
    End If
End Function

