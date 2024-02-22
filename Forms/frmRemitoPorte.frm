VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRemitoPorte 
   Caption         =   "Remito (Carta de Porte)"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   10545
   Icon            =   "frmRemitoPorte.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   10545
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabRemito 
      Height          =   5160
      Left            =   60
      TabIndex        =   21
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
      TabPicture(0)   =   "frmRemitoPorte.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEdicion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmRemitoPorte.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grilla2"
      Tab(1).Control(1)=   "cmdPedidos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Numeros de Serie"
      TabPicture(2)   =   "frmRemitoPorte.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grillaSeries"
      Tab(2).Control(1)=   "chkSinSeries"
      Tab(2).Control(2)=   "cmdLlenaSerie"
      Tab(2).Control(3)=   "lblErrorSeries"
      Tab(2).Control(4)=   "Label10"
      Tab(2).ControlCount=   5
      Begin VB.Frame fraEdicion 
         Height          =   1155
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   9915
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   1020
            TabIndex        =   17
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtProductoDescripcion 
            Height          =   320
            Left            =   3420
            Locked          =   -1  'True
            TabIndex        =   16
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
            Left            =   6180
            TabIndex        =   19
            Top             =   660
            Width           =   1215
         End
         Begin VB.CommandButton cmdAgregar 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   8520
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmRemitoPorte.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   495
         End
         Begin VB.CommandButton cmdBorrar 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   9060
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmRemitoPorte.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   44
            TabStop         =   0   'False
            ToolTipText     =   "Borrar Item"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtConsignacion 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   3720
            TabIndex        =   18
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
            TabIndex        =   15
            Top             =   180
            Width           =   375
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
            Left            =   60
            TabIndex        =   48
            Top             =   660
            Width           =   1095
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
            TabIndex        =   47
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
            Left            =   5340
            TabIndex        =   46
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Label12 
            Caption         =   "En Consignacion:"
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
            Left            =   2100
            TabIndex        =   45
            Top             =   660
            Width           =   1635
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla2 
         Height          =   3990
         Left            =   -74760
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdLlenaSerie 
         Caption         =   "Llenar Serie"
         Height          =   315
         Left            =   -71520
         TabIndex        =   31
         ToolTipText     =   "Seleccione filas a llenar"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdPedidos 
         Caption         =   "Pedidos"
         Height          =   315
         Left            =   -74760
         TabIndex        =   29
         Top             =   540
         Width           =   1155
      End
      Begin VB.Frame fraDetalle 
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   180
         TabIndex        =   28
         Top             =   480
         Width           =   10215
         Begin VSFlex7LCtl.VSFlexGrid grilla 
            Height          =   3405
            Left            =   60
            TabIndex        =   37
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
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label Label10 
         Caption         =   "Puede hacer 'Doble Clic' en el campo  Nro.Serie"
         Height          =   495
         Left            =   -74280
         TabIndex        =   30
         Top             =   540
         Width           =   2235
      End
   End
   Begin VB.Frame fraCabecera 
      Height          =   2235
      Left            =   60
      TabIndex        =   22
      Top             =   0
      Width           =   10395
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   3
         Left            =   5340
         TabIndex        =   13
         Top             =   1860
         Width           =   4935
      End
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   2
         Left            =   840
         TabIndex        =   12
         Top             =   1860
         Width           =   4395
      End
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   1
         Left            =   5340
         TabIndex        =   11
         Top             =   1500
         Width           =   4935
      End
      Begin VB.TextBox txtobs 
         Alignment       =   1  'Right Justify
         Height          =   320
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   1500
         Width           =   4395
      End
      Begin VB.CommandButton cmdCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "?"
         Height          =   315
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1020
         Width           =   375
      End
      Begin VB.CheckBox chkPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Propio "
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8280
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbDeposito 
         Height          =   315
         ItemData        =   "frmRemitoPorte.frx":0F32
         Left            =   4740
         List            =   "frmRemitoPorte.frx":0F34
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
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
         Left            =   3060
         TabIndex        =   9
         Top             =   1020
         Width           =   3795
      End
      Begin VB.TextBox txtClienteCodigo 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1260
         TabIndex        =   7
         Top             =   1020
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   8520
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   81068033
         CurrentDate     =   38126
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
         Left            =   180
         TabIndex        =   38
         Top             =   1560
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
         Left            =   3600
         TabIndex        =   33
         Top             =   300
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
         TabIndex        =   27
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
         Left            =   3540
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "NroCarta :"
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
         TabIndex        =   24
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
         Left            =   7500
         TabIndex        =   23
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
         TabIndex        =   40
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
         Left            =   9000
         TabIndex        =   42
         Top             =   50
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
         TabIndex        =   41
         Top             =   60
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRemitoPorte"
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

Private Sub cliente_cambio(codigo As Variant)
    Dim nc As Long, np As Long, i As Long
    nc = cliente.codigo
    np = s2n(txtPedido)
    
    If nc = 0 Then
        g.Borrar
        g2.Borrar
    Else
        If np > 0 Then
            g2.Borrar
            Carga1 nc, np
        Else
            g.Borrar
            Carga2 nc, 0
        End If
        i = s2n(obtenerDato("clientes", cliente.codigo, "transporte"))
        If i > 0 Then cmbTransporte = ObtenerDescripcion("transportes", i)
    End If
End Sub


Private Sub cmdAgregar_Click()
    Dim r As Long
    Dim pco As String
    
    If Not IsNumeric(txtcantidad) Or txtProductoCodigo = "" Or txtProductoDescripcion = "" Then Exit Sub
    If s2n(txtConsignacion) > s2n(txtcantidad) Then
        txtConsignacion.SetFocus
        txtConsignacion.SelStart = 0
        txtConsignacion.SelLength = Len(txtConsignacion.Text)
        Exit Sub
    End If
    
    pco = Trim(txtProductoCodigo)
    MetoEnGrilla pco, txtProductoDescripcion, s2n(txtcantidad, 0), "", s2n(txtPrecio, 4), s2n(txtConsignacion, 4), ""
    txtcantidad = ""
    txtProductoCodigo = ""
    txtProductoDescripcion = ""
    txtPrecio = ""
    txtConsignacion = ""
    txtProductoCodigo.SetFocus
    
    chkPropioEnabled True ' permite habilitar, ...
    CalculaTotal
End Sub

Private Sub cmdAyuda_Click()
    frmBuscarProducto
    If frmBuscar.resultado() = "" Then Exit Sub
    
    txtProductoCodigo = frmBuscar.resultado(1)
    txtProductoDescripcion = frmBuscar.resultado(2)
End Sub


Private Sub cmdAyudaPedidos_Click2() 'sacar el dos para habilitar
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

Private Sub cmdAyudaPedidos_Click()
    Dim ssql As String ', i As Long
    
    With frmBuscar
        ssql = " SELECT DISTINCT numero as [ Numero    ], fecha as [ Fecha           ], Cliente, Clientes.Descripcion as [ Nombre Cliente     ], CodigoPropio " _
            & " FROM ItemPedidoCliente " _
            & " INNER JOIN Pedidos_Clientes ON ItemPedidoCliente.PEDIDO = Pedidos_Clientes.numero " _
            & " INNER join Clientes on Pedidos_Clientes.cliente = clientes.codigo"
       
        If cliente.codigo = 0 Then
            ssql = ssql & " where Pedidos_Clientes.activo = 1 and ItemPedidoCliente.facturar > 0 order by Pedidos_Clientes.numero desc"
        Else
            ssql = ssql & " where Pedidos_Clientes.activo = 1 and ItemPedidoCliente.facturar > 0 and cliente = " & cliente.codigo & " order by Pedidos_Clientes.numero desc"
        End If
        If .MostrarSql(ssql, , , , "SI", "  ") = "" Then Exit Sub

        txtPedido = .resultado(1)
        cliente.codigo = .resultado(3)
        chkPropio = IIf(.resultado(4) = 1, vbChecked, vbUnchecked)
        

'        i = s2n(obtenerDeSQL("select transporte from Pedidos_Clientes where activo = 1 and  numero = " & txtPedido))
'        If i > 0 Then cmbTransporte = ObtenerDescripcion("transportes", i)
        CargoTransporteDePedido txtPedido

        Carga1 s2n(cliente.codigo), s2n(txtPedido)
        tabRemito.Tab = 0
    End With
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    If g.Row > 0 Then
        'If MsgBox("�Desea Actualizar Stock?", vbYesNo) = vbYes Then
        '    ABMRVDetalle "B", s2n(TxtRemitoNumero), g.tx(g.Row, gCODI), s2n(g.tx(g.Row, gCANT), 4), s2n(g.tx(g.Row, gPREC), 4), s2n(g.tx(g.Row, gPEDI), 4), 0, s2n(g.tx(g.Row, gCONS)), g.tx(g.Row, gFORM), 1
        'End If
        grilla.RemoveItem (g.Row)
    End If
    chkPropioEnabled True
    CalculaTotal
End Sub


Private Sub cmdPedidos_Click()
    'Carga2 s2n(cliente.codigo), s2n(txtPedido)
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
'    ucMenu.MsgConfirmaCancelar = "� Desea abandonar la edicion ? "
'    ucMenu.MsgConfirmaSalir = "Cerrar Formulario ? "
'    ucMenu.MsgConfirmaEliminar = "Anula este Remito ?"
'    ucMenu.CaptionEliminar = "Anular"
    
    chkPropio.Value = vbChecked
    comboSql cmbTransporte, "select descripcion from transportes where activo = 1"
    comboArray cmbDeposito, Array("Deposito Central", "Deposito 1", "Deposito 2", "Deposito 3", "Deposito4"), Array(0, 1, 2, 3, 4)
    
    rsEjercicio.Open "SELECT * From Ejercicio WHERE activo =1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    ucBetween.ini rsEjercicio!FechaInicio, rsEjercicio!FechaFin
    dtFecha = Date
    
    lblPrecio.Visible = REMITO_CON_PRECIO
    txtPrecio.Visible = REMITO_CON_PRECIO
    
    tabRemito.TabVisible(2) = gEMPR_Maneja_series
    tabRemito.TabVisible(1) = False
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
    
    If (Col = gPREC Or Col = gCANT) And txtClienteCodigo.Text <> "" Then
        If grilla.TextMatrix(Row, gPEDI) <> "" Then
            sql = "select cantidad, saldo  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where cliente = " & txtClienteCodigo.Text & " and activo = 1 and saldo>0 and pedido=" & grilla.TextMatrix(Row, gPEDI) & " and producto='" & grilla.TextMatrix(Row, gCODI) & "'"
            rs.Open sql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If rs.EOF = True And rs.BOF = True Then
                CalculaTotal
            Else
                If CDbl(g.TextMatrix(Row, gCANT)) <= CDbl(rs!saldo) Then
                    CalculaTotal
                Else
                    MsgBox "La cantidad ingresada supera la del pedido.", vbExclamation, "Advertencia"
                    grilla.TextMatrix(Row, gCANT) = rs!saldo
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub grilla_DblClick()
    If ucMenu.Estado <> ucbEditando Then Exit Sub

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
    
    If ucMenu.Estado <> ucbEditando Then Exit Sub
    
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

Private Sub CargaRemito(cual)
    If ON_ERROR_HABILITADO Then On Error GoTo E_UFA
    Dim rs As New ADODB.Recordset, i As Long, desc As Variant, prod As String
    
    With rs
        .Open "select * from remitoporte where numero = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        TxtRemitoNumero = !numero
        cliente.codigo = !cliente
        'cmbClienteNombre = ObtenerDescripcion("clientes", !Cliente)
        dtFecha = !fecha
        cmbDeposito.ListIndex = s2n(!DEPOSITO)
        txtobs(0) = sSinNull(!obs1)
        txtobs(1) = sSinNull(!obs2)
        txtobs(2) = sSinNull(!Obs3)
        txtobs(3) = sSinNull(!obs4)
        cmbTransporte = ObtenerDescripcion("transportes", s2n(!Transporte))
        .Close
    
        .Open "select * from RemitoPorteDetalle where CANCELADO=0 AND numero = " & cual, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        g.rows = 1
        While Not .EOF
            i = g.addRow()
            
            desc = obtenerDeSQL("select descripcion from producto where codigo = '" & !producto & "' and activo = 1 ")
            If IsNull(desc) Then desc = ""
            
            grilla.TextMatrix(i, gCODI) = VerProductoCliente(!producto, Propio(), cliente.codigo)
            grilla.TextMatrix(i, gDESC) = desc
            grilla.TextMatrix(i, gCANT) = !cantidad
            grilla.TextMatrix(i, gPEDI) = !PEDIDO
            grilla.TextMatrix(i, gPREC) = s2n(!precio, 4)
            grilla.TextMatrix(i, gFORM) = !formula
            .MoveNext
        Wend
        .Close
        
        '.Open "select producto, serie, consignacion from series where nroComprobante = " & cual & " and Comprobante = " & TipoComprobante_REMITOVENTA, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        'g3.Borrar
        'While Not .EOF
        '    i = g3.addRow()
        '    prod = VerProductoCliente(!producto, Propio(), cliente.codigo)
        '    grillaSeries.TextMatrix(i, g3PROD) = prod
        '    grillaSeries.TextMatrix(i, g3DESC) = ProductoDescripcion(prod)
        '    grillaSeries.TextMatrix(i, g3NSER) = !Serie
        '    grillaSeries.TextMatrix(i, g3CONS) = !consignacion
        '    .MoveNext
        'Wend
        '.Close
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
    ssql = "select producto, cantidad, pedido, precio, CodigoPropio, pago, cliente, saldo, formula  from ItemPedidoCliente inner join pedidos_clientes on pedidos_clientes.numero = ItemPedidoCliente.pedido where activo = 1 "
    If NCliente > 0 Then ssql = ssql & " and cliente = " & NCliente
    If NPedido > 0 Then ssql = ssql & " and pedido = " & NPedido
    
    With rs
        
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then cliente.codigo = !cliente
        g.Borrar
        While Not .EOF
            chkPropio = IIf(!CODIGOPROPIO, vbChecked, vbUnchecked)
            'can = s2n(!cantidad)
            can = s2n(!cantidad) 'saldo
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
            grilla2.TextMatrix(i, g2PEDI) = !PEDIDO
            grilla2.TextMatrix(i, g2PREC) = !precio
            grilla2.TextMatrix(i, g2FORM) = !formula
            .MoveNext
        Wend
    End With
    tabRemito.Tab = 0
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
    
    cliente.codigo = 0
    chkSinSeries.Value = 0
    g.Borrar
    g2.Borrar
    g3.Borrar
    tabRemito.Tab = 0
    mUltimoPedido = 0
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
        gPREC = g.AddCol(" Precio ", "N", 4)
        g2PREC = g2.AddCol(" Precio ")
    Else                                ' Oculto precio
        gPREC = g.AddCol(" Precio ", "H")
        g2PREC = g2.AddCol(" Precio ", "H")
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
    
    If ucMenu.Estado <> ucbEditando Then Exit Sub
    
    r = g3.Row
    If r < 1 Then Exit Sub
    prod = VerProductoMio(g3.tx(r, g3PROD), Propio())
    
    resu = Buscar_SeriesEnStock(prod)
    If resu > "" Then grillaSeries.TextMatrix(r, g3NSER) = resu
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
    If r > 1 And g3.Buscar(g3NSER, "") > 0 Then
        
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
            If ns <> "" And g3.Buscar(g3NSER, ns, i + 1) > 0 Then
                tabRemito.Tab = 2
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER, g3.Buscar(g3NSER, ns, i + 1), g3NSER
                
                'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                FaltaSeries = True
                Exit Function
            End If
        Next i
    End If

End Function


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
End Sub


Private Function frmBuscarProducto()
If cliente.codigo = 0 Then Exit Function
  
    frmBuscar.MostrarSql ("SELECT codigo AS [Codigo               ],descripcion AS [Descripcion                                                    ] FROM producto WHERE activo=1")
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
    BuscoNroYTipo = True
    
    BuscoNroYTipo = RevisaNroYFechaOk("RemitoPorte", "Numero", "fecha", s2n(TxtRemitoNumero, 0), dtFecha, "", False)
End Function

Private Sub ucMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim num As Long, Serie As String
    Dim i As Long, tmp, asse As String, produ As String, mstk As Long
    Dim sucursal As Long, consig As Boolean, cantConsign As Double, formula As String, depot As Long
    
    CalculaTotal
    
    num = s2n(TxtRemitoNumero)
    
    If Not BuscoNroYTipo() Then
        Exit Sub
    End If
    'sin cabeza
    asse = "faltacabecera"
    If FaltaCabecera() Or FaltaGrilla() Then
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
    tmp = obtenerDeSQL("select cliente from Remitoporte where Numero = " & s2n(TxtRemitoNumero))
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
    If ABMRemitoPorte("A", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtobs(0), txtobs(1), txtobs(2), txtobs(3)) = False Then GoTo ufaErr
    asse = "RVdetalle"
    For i = 1 To g.rows - 1 'items
        cantConsign = s2n(g.tx(i, gCONS))
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gCODI), Propio())
        mstk = 1
        'Abs(ManejaStock(produ))
        If ABMRPDetalle("A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk) = False Then GoTo ufaErr
        'DataEnvironment1.dbo_abmRemitoVentaDetalle "A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk
    Next i
'''''    asse = "RVseries"
'''''    For i = 1 To g3.rows - 1 'series
'''''        Serie = g3.tx(i, g3NSER)
'''''        consig = (grillaSeries.cell(flexcpChecked, i, g3CONS) = flexChecked)
'''''        If Serie <> "" Then
'''''           'DataEnvironment1.dbo_SERIE "A", 0, VerProductoMio(g3.tx(i, g3PROD), Propio()), serie, TipoComprobante_REMITOVENTA, num, sucursal, 0, "", consig, CLng(Date), UsuarioActual(), 0, 0
'''''            DataEnvironment1.dbo_abmSERIEs "A", 0, VerProductoMio(g3.tx(i, g3PROD), Propio()), Serie, TipoComprobante_REMITOVENTA, Num, sucursal, 0, "", consig, dtfecha, 1, Date, UsuarioActual()
'''''        End If
'''''    Next i
    
    DE_CommitTrans
    
    
    ' quiero Transaccion, y/o quiero hacer tabla temp y 1 solo stored
    '*******************************************************************
    asse = "grabado, fallo impresion"
    
    ImprimirRemitoPorte num
    MsgBox "Remito " & num & " grabado"
    
    ucMenu.AceptarOk
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
    Dim sucursal As Long, consig As Boolean, cantConsign As Double, formula As String, depot As Long
    
    CalculaTotal
    

    num = s2n(TxtRemitoNumero)
    
    If FaltaCabecera() Or FaltaGrilla() Then ' verifica si falta cabecera
        MsgBox "Faltan datos en el formulario"
        Exit Sub
    End If
    
    If HayProdEnEdicion(txtProductoDescripcion) Then Exit Sub
    
    If FaltaSeries() Then
        Exit Sub
    End If
    
    tmp = obtenerDeSQL("select numero, fecha from RemitoPorte where Numero = " & s2n(TxtRemitoNumero))
    If Not IsEmpty(tmp) Then
        If MsgBox("Remito (Carta de Porte) Numero : " & tmp(0) & " (Fecha : " & tmp(1) & "), ya existe." & Chr(13) & "�Desea actualizar los datos existentes en el remito?.", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    sucursal = s2n(obtenerDeSQL("select sucursal from datos"))
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
        
    DE_BeginTrans
    
    'DataEnvironment1.dbo_abmRemitoVenta "M", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtObs(0), txtObs(1), txtObs(2), txtObs(3)
    ABMRemitoPorte "M", num, cliente.codigo, (dtFecha), 0, ObtenerCodigo("transportes", cmbTransporte), depot, Propio(), txtobs(0), txtobs(1), txtobs(2), txtobs(3)
    
    For i = 1 To g.rows - 1 'carga en remito_venta_detalle todos los items
        cantConsign = s2n(g.tx(i, gCONS))
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gCODI), Propio())
        mstk = 1
        
        ABMRPDetalle "M", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), 0, depot, cantConsign, "V", mstk
        'DataEnvironment1.dbo_abmRemitoVentaDetalle "A", num, produ, s2n(g.tx(i, gCANT), 4), s2n(g.tx(i, gPREC), 4), s2n(g.tx(i, gPEDI), 4), depot, cantConsign, formula, mstk
    Next i
    
   
    DE_CommitTrans
    ucMenu.AceptarOk
    
    ImprimirRemitoPorte num
    MsgBox "Remito Porte" & num & " guardado.", vbInformation
    

    Exit Sub
rv_err:
    DE_RollbackTrans
    ufa "Error al grabar: " & asse, Me.Name & " " & num    ', Err
End Sub
Private Sub ucMenu_BorrarControles()
    BorrarCampos
End Sub
Private Sub ucMenu_Buscar()
    frmBuscar.MostrarSql "select Numero, Cliente, Fecha as [ Fecha   ],  Cancelado, Anulado  from RemitoPorte where fecha " & ucBetween.ssBetween & " order by numero desc", , , , "SI", ""
    If frmBuscar.resultado() <> "" Then
        CargaRemito frmBuscar.resultado(1)
        ucMenu.BuscarOK
        g2.Borrar
        tabRemito.Tab = 0
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ERR_ELIM
    Dim i As Long, num As Long, prod As String, cant As Double, pedi As Long, formu As String
    Dim tmp, tmpFac, tx As String, mjstk As Long
    
    'Controlo----------------
    ' sin numero
    If TxtRemitoNumero = "" Then Exit Sub
    num = s2n(TxtRemitoNumero)
    'ya anulado
    If obtenerDeSQL("select Anulado from RemitoPorte where numero = " & num) = True Then
        MsgBox "Remito ya anulado"
        Exit Sub
    End If

    'ya facturado
    'tmp = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where numero = " & num)
    'If tmp > 0 Then
    '    'tmp = obtenerdesql ("select NroFactura from FacturaVenta
    '      che "No puedo anular, remito con factura " '& vbCrLf & tmpFac(0) & " " & tmpFac(1)
    '      Exit Sub
    'End If
    
    'Dim sp
    'Set sp = Empty 'obtenerDeSQL("select * from facturaventadetalle d inner join facturaventa f on d.nrofactura=f.nrofactura where f.activo=1 and d.nroremito=" & s2n(num))
    'If IsNull(sp) Or IsEmpty(sp) Then
    'Else
    '    MsgBox "No se puede eliminar el comprobante. Esta asociada a otro comprobante", vbCritical
    '    Exit Sub
    'End If
   
    If confirma("Anular remito " & TxtRemitoNumero) Then
        tx = InputBox("Motivo ")
        'If Trim(tx) = "" Then Exit Sub
       
        DE_BeginTrans
            'detalle
            ABMRemitoPorte "B", num, 0, 0, 0, 0, 0, 0, "", "", "", tx
            For i = 1 To g.rows - 1
                prod = VerProductoMio(g.tx(i, gCODI), Propio())
                mjstk = ManejaStock(prod)
                cant = s2n(g.tx(i, gCANT))
                pedi = s2n(g.tx(i, gPEDI))
                formu = IIf(EsProductoVirtual(prod), CHAR_PROD_VIRTUAL, "")
                ABMRPDetalle "B", num, prod, cant, 0, pedi, 0, 0, formu, mjstk
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
    ImprimirRemitoPorte (s2n(TxtRemitoNumero))
    ' aqui
End Sub

Private Sub ucMenu_Modificar()
    mNuevo = False
    TxtRemitoNumero.Locked = True
End Sub
Private Sub ucMenu_Nuevo()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
    mNuevo = True
    TxtRemitoNumero = nuevoCodigo("RemitoPorte", "numero") ' obtenerDeSQL("select max(numero ) from RemitoVenta ") + 1 ' LeerBS_Num(CAMPO_BS_NroREMITO) + 1
'    TxtRemitoNumero.SetFocus
    'nuevoCodigo ("RemitoVenta","Numero")
fin:
    Exit Sub
ufa:
    TxtRemitoNumero = "1"
    Resume fin
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub

Public Function ABMRemitoPorte(rOPE As String, rNumero As Long, rCliente As Long, rFecha As Date, rFactura As Long, rTransporte As Long, rDEPOSITO As Long, rCodPropio As Long, rObs1 As String, rObs2 As String, rObs3 As String, rMotivo As String) As Boolean
On Error GoTo rvmal
Dim iudr As String
ABMRemitoPorte = True
    Select Case rOPE
        Case "A":
            iudr = "INSERT INTO RemitoPorte(NUMERO,CLIENTE,FECHA,FACTURA,TRANSPORTE,CANCELADO,ANULADO,DEPOSITO,CODPROPIO,Obs1 , Obs2, Obs3, obs4) " _
                & " VALUES( " & rNumero & "," & rCliente & "," & ssFecha(rFecha) & "," & rFactura & "," & rTransporte & ", 0, 0," & rDEPOSITO & "," & rCodPropio & "," & sstexto(rObs1) & "," & sstexto(rObs2) & "," & sstexto(rObs3) & "," & sstexto(rMotivo) & ")"
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
            'iudr = "Update SERIES  Set activo = 0   Where NroComprobante = " & rNumero & " And COMPROBANTE = 5"
            'DataEnvironment1.Sistema.Execute iudr
            iudr = "Update RemitoPorte  set anulado = 1, cancelado = 1  where numero = " & rNumero
            DataEnvironment1.Sistema.Execute iudr
    End Select
Exit Function
rvmal:
ABMRemitoPorte = False

End Function

Public Function ABMRPDetalle(dOpe As String, dNumero As Long, dPRODUCTO As String, dCantidad As Double, dPrecio As Double, dPedido As Long, dDeposito As Long, dConsign As Double, dFormula As String, dManejaStock As Long) As Boolean
On Error GoTo rdmal
Dim iudd As String
Dim r, rFactor As Double, rCargar As Double

'Set r = Nothing
'r = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(dPRODUCTO))
'If IsNull(r) Or IsEmpty(r) Then
'    rFactor = 1
'Else
'    rFactor = r
'End If
'rCargar = rFactor * dCantidad

ABMRPDetalle = True
Select Case dOpe
    Case "A":
        iudd = "INSERT INTO RemitoPorteDETALLE (NUMERO,PRODUCTO,CANTIDAD,FACTURAR,PEDIDO,PRECIO,CANCELADO,CONSIGNACION,FORMULA) " _
            & " values (" & dNumero & "," & sstexto(dPRODUCTO) & "," & x2s(dCantidad) & "," & x2s(dCantidad) & "," & dPedido & "," & x2s(dPrecio) & ", 0," & x2s(dConsign) & "," & sstexto(dFormula) & ")"
        DataEnvironment1.Sistema.Execute iudd
        
    Case "M":
        iudd = "UPDATE RemitoPorteDETALLE SET " _
            & " CANTIDAD=" & x2s(dCantidad) & ",FACTURAR=" & x2s(dCantidad) & ",PRECIO= " & x2s(dPrecio) & "" _
            & " WHERE PRODUCTO=" & sstexto(dPRODUCTO) & " AND NUMERO=" & dNumero
        DataEnvironment1.Sistema.Execute iudd
    Case "B":
        'ya esta cancelado la cabecera no hace falta el detalle
        iudd = " Update RemitoPorteDETALLE Set cancelado = 1 " _
        & " where numero = " & dNumero & " and producto =" & sstexto(dPRODUCTO)
        DataEnvironment1.Sistema.Execute iudd
End Select
Exit Function
rdmal:
ABMRPDetalle = False
End Function
