VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form FrmOrdenPago 
   Caption         =   "Orden de Pago"
   ClientHeight    =   8655
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "FrmOrdenPago.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8655
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbingresar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imputaciones Contables"
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
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   2355
   End
   Begin VB.TextBox txtFalta 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10050
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   750
      Width           =   1035
   End
   Begin VB.TextBox txttot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Tag             =   "8"
      Top             =   765
      Width           =   855
   End
   Begin VB.TextBox tot1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   45
      Width           =   855
   End
   Begin VB.TextBox tot2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
   End
   Begin TabDlg.SSTab tabOP 
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comprobantes"
      TabPicture(0)   =   "FrmOrdenPago.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(4)=   "fac"
      Tab(0).Control(5)=   "credito"
      Tab(0).Control(6)=   "rac"
      Tab(0).Control(7)=   "debito"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Forma de Pago"
      TabPicture(1)   =   "FrmOrdenPago.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraPago"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraPago 
         BorderStyle     =   0  'None
         Caption         =   "Retenciones"
         Height          =   5235
         Left            =   75
         TabIndex        =   24
         Top             =   420
         Width           =   11115
         Begin Gestion.ucRetCompras uRetCompras 
            Height          =   780
            Left            =   45
            TabIndex        =   23
            Top             =   75
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   1376
         End
         Begin VB.TextBox txtTotalPago 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9660
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   4860
            Width           =   1035
         End
         Begin VB.TextBox txtTotalRet 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9690
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   450
            Width           =   1035
         End
         Begin VB.TextBox txtefectivo 
            Height          =   285
            Left            =   810
            TabIndex        =   25
            Top             =   1200
            Width           =   1080
         End
         Begin VB.TextBox txttransf 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Top             =   4875
            Width           =   1080
         End
         Begin VB.TextBox txtcuenta 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4890
            TabIndex        =   38
            Tag             =   "2"
            Top             =   4875
            Width           =   2895
         End
         Begin VB.CommandButton cmbcuenta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3990
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   4875
            Width           =   855
         End
         Begin VB.TextBox txtcodcuenta 
            Height          =   285
            Left            =   2970
            TabIndex        =   35
            Top             =   4875
            Width           =   975
         End
         Begin VB.TextBox txtcodcaja 
            Height          =   285
            Left            =   2565
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1185
            Width           =   975
         End
         Begin VB.CommandButton cmbcaja 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caja"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtcaja 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4515
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "2"
            Top             =   1215
            Width           =   2895
         End
         Begin Gestion.ucCheques uCheques 
            Height          =   2715
            Left            =   795
            TabIndex        =   31
            Top             =   1860
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   4789
         End
         Begin VB.Line Line1 
            X1              =   45
            X2              =   10710
            Y1              =   1035
            Y2              =   1035
         End
         Begin VB.Label label26 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Caja"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2160
            TabIndex        =   46
            Top             =   1215
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cheques"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   75
            TabIndex        =   34
            Top             =   2040
            Width           =   915
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Efectivo"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   30
            TabIndex        =   32
            Top             =   1215
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Debito"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   105
            TabIndex        =   30
            Top             =   4890
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2325
            TabIndex        =   28
            Top             =   4875
            Width           =   735
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid debito 
         Height          =   2295
         Left            =   -69600
         TabIndex        =   6
         Top             =   750
         Width           =   5775
         _cx             =   10186
         _cy             =   4048
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
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
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
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid rac 
         Height          =   2295
         Left            =   -74940
         TabIndex        =   7
         Top             =   3435
         Width           =   5175
         _cx             =   9128
         _cy             =   4048
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
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
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
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid credito 
         Height          =   2295
         Left            =   -69600
         TabIndex        =   8
         Top             =   3420
         Width           =   5775
         _cx             =   10186
         _cy             =   4048
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
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmOrdenPago.frx":0902
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
         Editable        =   2
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
      Begin VSFlex7LCtl.VSFlexGrid fac 
         Height          =   2295
         Left            =   -74970
         TabIndex        =   5
         Top             =   735
         Width           =   5175
         _cx             =   9128
         _cy             =   4048
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
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmOrdenPago.frx":098A
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
         Editable        =   2
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
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -73140
         TabIndex        =   22
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "N/D - Aj. por débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   -67860
         TabIndex        =   21
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pagos a Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -73560
         TabIndex        =   20
         Top             =   3090
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "N/C - Aj. por crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   -67860
         TabIndex        =   19
         Top             =   3105
         Width           =   2415
      End
   End
   Begin Gestion.ucCoDe uProv 
      Height          =   315
      Left            =   1635
      TabIndex        =   3
      Top             =   420
      Width           =   5880
      _ExtentX        =   12091
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.TextBox txttipoiva 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3975
      TabIndex        =   13
      Tag             =   "1"
      Top             =   -60
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4275
      TabIndex        =   12
      Top             =   90
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtopago 
      Height          =   285
      Left            =   1635
      TabIndex        =   1
      Top             =   90
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   345
      Left            =   6120
      TabIndex        =   2
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   147718145
      CurrentDate     =   37934
   End
   Begin Gestion.ucBotonera uMenu 
      Cancel          =   -1  'True
      Height          =   1590
      Left            =   1440
      TabIndex        =   0
      Top             =   7080
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   2805
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucFecha ufDesde 
         Height          =   315
         Left            =   4020
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         FechaInit       =   5
      End
      Begin Gestion.ucFecha ufHasta 
         Height          =   315
         Left            =   5100
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         FechaInit       =   4
      End
      Begin VB.OptionButton optQueBusco 
         Caption         =   "Busca OP"
         Height          =   315
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optQueBusco 
         Caption         =   "Busca IMP"
         Height          =   315
         Index           =   1
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buscar Entre:"
         Height          =   315
         Index           =   1
         Left            =   2640
         TabIndex        =   18
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.Label Label7 
      Caption         =   "idDoc:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10020
      TabIndex        =   50
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblIdDoc 
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10605
      TabIndex        =   49
      Top             =   15
      Width           =   660
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Falta Pagar: "
      Height          =   195
      Left            =   10065
      TabIndex        =   48
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7890
      TabIndex        =   44
      Top             =   795
      Width           =   690
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      Visible         =   0   'False
      X1              =   7890
      X2              =   9450
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "FAC - ND"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   7890
      TabIndex        =   42
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NC - FAC"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   7890
      TabIndex        =   41
      Top             =   405
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Orden de Pago Nº"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   780
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5415
      TabIndex        =   9
      Top             =   105
      Width           =   855
   End
End
Attribute VB_Name = "FrmOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 23/3/5
' 17/11/4

'Dim a As Double, busco As String, cargrillas As Boolean

Private stblOrdenPagoTmp As String
Private stblChequesOPtmp As String


''retenciones
''Private WithEvents gRet As LiGrilla
'Private gR_NRO As Long  'nro retencion
'Private gR_FEC As Long  'fech
'Private gR_TIP As Long  'tipo: descripcion RIB RGA
'Private gR_IMP As Long  'importe
'Private gR_CUE As Long  'cuenta
'Private gR_TIC As Long  'codigo
'Private gR_IdC As Long  'idCuentasParam, el unico que se graba
''Private gR_FAC As Long  'numero factura ?
'
Private midDoc As Long

Private Enum queFue
    NADA
    pago
    imputacion
    AMBAS
End Enum

Private Enum busque
    buscoNADA
    buscoOP
    buscoIMP
End Enum
Private mBusco As busque ' OJO debe setearse con datos de frm (qué está cargado) , no con el boton (q puede camb en cualq momento)

Private Sub cmbcaja_Click()
    FrmHelp.Show
    CargarHelp "Cajas", "Codigo", "Descripción", "codigo", "responsable"
    FrmHelp.Tag = Me.Name
    cargar = "Cajas"
End Sub

Private Sub cmbcuenta_Click()
    FrmHelp.Show
    CargarHelpCtasBanc "Ctasbank", "Codigo", "Banco - Nº", "Codigo", "Banco", "Numero", "moneda", "Codigo"
    FrmHelp.Tag = Me.Name
    cargar = "CuentasBank"
End Sub

Private Function es_IMP() As Boolean
Dim Total As Double, sf As Double, sd As Double, sc As Double, sca As Double, x As Long
Total = 0

    'AVERIGUO SI LA SUMA DE FAC Y DEBITOS - LA SUMA DE CREDITOS Y RECIBOS ES IGUAL A CERO
    If fac.TextMatrix(1, 0) <> "" Then
        For x = 1 To fac.rows - 1
            If fac.TextMatrix(x, 3) <> "" Then
                sf = sf + s2n(fac.TextMatrix(x, 3))
            End If
        Next
    End If
    
    If debito.TextMatrix(1, 0) <> "" Then
        For x = 1 To debito.rows - 1
            If debito.TextMatrix(x, 4) <> "" Then
                sd = sd + s2n(debito.TextMatrix(x, 4))
            End If
        Next
    End If
    
    If credito.TextMatrix(1, 0) <> "" Then
        For x = 1 To credito.rows - 1
            If credito.TextMatrix(x, 4) <> "" Then
                sc = sc + s2n(credito.TextMatrix(x, 4))
            End If
        Next
    End If
    
    If rac.TextMatrix(1, 0) <> "" Then
        For x = 1 To rac.rows - 1
            If rac.TextMatrix(x, 3) <> "" Then
                sca = sca + s2n(rac.TextMatrix(x, 3))
            End If
        Next
    End If
    
    Total = s2n((sf + sd) - (sc + sca))
    
    If Total <> 0 Then
        es_IMP = False
    Else
        es_IMP = True
    End If
End Function


Private Function Graba_OP_IMP() As Boolean ' cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaNUEVO
    
    Dim rs As New ADODB.Recordset, x As Long, tmp, asse As String
    Dim Valor As Long, anticip As Long ', fechaop As String, fechadia As String
    'Dim suma As Double,
    'Dim retgan As Double
    Dim i As Long
    Dim sumafac As Double, sumadebito As Double, sumacredito As Double, sumarac As Double, Total As Double
    Dim refe, reim, reti, recu, reid
    Dim str As String
    Dim NroCertifGan As Long, NroCertifIIBB As Long
    Dim z As Double
    Dim RealizoPagoACuenta As Boolean
    z = 1
    RealizoPagoACuenta = False
    
    Dim cueche As Long ' cue ban che pro
    
    If UpROV.codigo = 0 Then
        che "Debe ingresar el código de proveedor"
        Exit Function
    End If
    
    If s2n(txttransf) <> 0 And Trim$(txtcuenta) = "" Then
        che "Falta cuenta para la transferencia"
        Exit Function
    End If
       
    'sindato
    If s2n(tot1) = 0 And s2n(tot2) = 0 Then
        che "Debe ingresar algun valor de comprobantes"
        Exit Function
    End If
    
    If s2n(txtefectivo) > 0 And Trim(txtcaja) = "" Then
        che "falta cuenta efectivo"
        Exit Function
    End If
    
    If ChequeaChq = False Then
        If MsgBox("Desea chequear los cheques antes de continuar?" & Chr(13) & "Tenga en cuenta que si continua pueden duplicarse.", vbQuestion + vbYesNo) = vbYes Then
            Exit Function
        End If
    End If
    
    Dim CuentaProv As String
    Dim Dat
    Dat = obtenerDeSQL("select categ,tiene_cuenta,cuenta from prov where codigo=" & UpROV.codigo)
    If Dat(0) = 3 Then
        If Dat(1) = 1 Then
            CuentaProv = Dat(2)
        Else
            CuentaProv = CuentaParam(ID_Cuenta_C_DEUD_A_PROV)
        End If
    Else
        CuentaProv = CuentaParam(ID_Cuenta_C_DEUD_A_PROV)
    End If
    '*************
   
'    suma = 0
    If s2n(txttot) > 0 Then
    'If 1 = 2 Then 'absurdo para que no entre, deshabilitado
        tmp = s2n(s2n(txtTotalPago) - s2n(txttot))
        If tmp > 0 Then
            If MsgBox("Se esta pagando mas de la cuenta. ¿Desea generar un pago a cuenta?" & vbCrLf & tmp, vbYesNo) = vbYes Then
                RealizoPagoACuenta = True
            Else
                'Exit Function
            End If
        ElseIf tmp < 0 Then
            Exit Function
        End If
    Else
        If s2n(txttot) < 0 Then
            MsgBox "El total de la orden no puede ser negativo"
            Exit Function
        End If
    End If


    If Not uCheques.FechasOk Then Exit Function
    
' -------------------------------------

    'Total = 0
    
'    'AVERIGUO SI LA SUMA DE FAC Y DEBITOS - LA SUMA DE CREDITOS Y RECIBOS ES IGUAL A CERO
'    If fac.TextMatrix(1, 0) <> "" Then
'        For x = 1 To fac.rows - 1
'            If fac.TextMatrix(x, 3) <> "" Then
'                sumafac = sumafac + s2n(fac.TextMatrix(x, 3))
'            End If
'        Next
'    End If
'
'   If debito.TextMatrix(1, 0) <> "" Then
'        For x = 1 To debito.rows - 1
'            If debito.TextMatrix(x, 4) <> "" Then
'                sumadebito = sumadebito + s2n(debito.TextMatrix(x, 4))
'            End If
'        Next
'    End If
'
'    If credito.TextMatrix(1, 0) <> "" Then
'        For x = 1 To credito.rows - 1
'            If credito.TextMatrix(x, 4) <> "" Then
'                sumacredito = sumacredito + s2n(credito.TextMatrix(x, 4))
'            End If
'        Next
'    End If
'
'    If rac.TextMatrix(1, 0) <> "" Then
'        For x = 1 To rac.rows - 1
'            If rac.TextMatrix(x, 3) <> "" Then
'                sumarac = sumarac + s2n(rac.TextMatrix(x, 3))
'            End If
'        Next
'    End If
    
'    Total = s2n((sumafac + sumadebito) - (sumacredito + sumarac))
    
'    If Total <> 0 Then  ' numero O/P
    If es_IMP = False Then '  numero O/P
        If Not RevisaNro("Rec_comp", "Nro", "fecha", s2n(txtopago), "activo = 1") Then
            Exit Function
        End If
        If Not RevisaNro("Compras", "NroDoc", "fecha", s2n(txtopago), "tipodoc = 'RAC' ") Then
            Exit Function
        End If
        If Not RevisaNro("transcom", "NroDoc", "fecha", s2n(txtopago), "tipodoc = 'RAC'") Then
            Exit Function
        End If
    Else                ' numero IMP
        rs.Open "select num_imppro, anticip_pr from bs", DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
        If Not rs.EOF Then
            Valor = rs!NUM_ImpPro + 1
            txtopago = Valor
            Label2.caption = "Imputacion Nº"
            anticip = s2n(rs!anticip_pr) 'trucho s2n() por null
       End If
        rs.Close
        Set rs = Nothing
    End If
    
    
    Dim AsientoCompra As New Asiento, iddoc As Long
    Dim sComprobante As String
    Dim numIntP As Long, numIntT As Long ' nro int cheques P T
    
    
    '********************************************
    'DE_BeginTrans
    
        
    asse = "0-"
    If es_IMP Then
        
        iddoc = NuevoDocumento("IMP", Valor, 0, 0)
        AsientoCompra.nuevo "IMP " & UpROV.DESCRIPCION, Fecha, "IMP"
        
        'IMPPRO
        asse = "01 ingresoImputacion"
        DataEnvironment1.dbo_INGRESOIMPUTACION Fecha, Valor, UpROV.codigo, iddoc, Date, UsuarioSistema!codigo
                
        'INCREMENTO NUM_IMPPRO DE LA TABLA BS
        asse = " 03 dbo_INCREMENTONUMIMPPRO"
        DataEnvironment1.dbo_INCREMENTONUMIMPPRO Valor
        
        Call ProcesoGrillas(Valor, "IMP", AsientoCompra, iddoc)
    
'        AsientoCompra.AcumularItem CuentaParam(ID_CuentasParam_DEUD_A_PROV), (sumafac + sumadebito), 0
 '       AsientoCompra.AcumularItem CuentaParam(ID_CuentasParam_ANTICIP_A_PROV), 0, (sumafac + sumadebito)

    Else
    
        If uRetCompras.retgan > 0 Then NroCertifGan = NuevoNroCertifGan()
        If uRetCompras.retIB > 0 Then NroCertifIIBB = NuevoNroCertifIIBB()
    
        sComprobante = "OP " & txtopago
        'iddoc = NuevoDocumento("O/P", s2n(txtopago), 0, s2n(txtopago))   ' uprov.codigo (nro unico)
        iddoc = NuevoDocumento("O/P", s2n(txtopago), 0, s2n(txtopago), NroCertifGan, NroCertifIIBB) ' uprov.codigo (nro unico)
        
        'ASIENTO CABECERA
        'AsientoCompra.Nuevo "O/Pago " & txtopago & " " & uProv.descripcion, Fecha, "O/P"
        AsientoCompra.nuevo "OP " & UpROV.DESCRIPCION, Fecha, "O/P"
        'ASIENTO DEBE
'        AsientoCompra.AgregarItem CuentaParam(ID_CuentasParam_DEUD_A_PROV), Total, 0
                              
        Dim porciva As Double, maximobanc As Long, maximocaja As Long
        Dim valorcuenta As String, valorcuentacon As String, valorcartera As String, fechapropio As String, valcartera As String
        Dim Neto As Double
   
        If existeOP(txtopago) Then
            che "err: ya existe numero OPago"
            DE_RollbackTrans
            Exit Function
        End If
        
        rs.Open "select * from PorcentajesIva where iva = " & val(txttipoiva) & " order by fecha_baja", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not rs.EOF
            If IsNull(rs!fecha_baja) Then
                porciva = rs!PORCENTAJE
            Else
                porciva = 0
            End If
            rs.MoveNext
        Wend
        rs.Close
'        Set rs = Nothing
                                    
        Neto = s2n(s2n(txttot) / (1 + porciva))
        
        'REC_COMP
        asse = " 4 IngrIMputacion2"
        DataEnvironment1.dbo_INGRESOIMPUTACION2 Fecha, val(txtopago), UpROV.codigo, s2n(txttot), s2n(Neto), uRetCompras.retgan, uRetCompras.retIB, iddoc, Date, UsuarioSistema!codigo, 1, z

        'Asiento Retenciones compra
        'RetGan
        AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_GAN_3ros), 0, uRetCompras.retgan ' retgan  ', sComprobante
        AsientoCompra.AcumularItem CuentaProv, uRetCompras.retgan, 0, sComprobante
        'iibb
        AsientoCompra.AgregarItem CuentaParam(ID_Cuenta_P_RET_IB_3ros), 0, uRetCompras.retIB    ' retgan  ', sComprobante
        AsientoCompra.AcumularItem CuentaProv, uRetCompras.retIB, 0, sComprobante
        
        
        asse = "7 transferencia"
        'SI REALIZO UNA TRANSFERENCIA
        If s2n(txttransf) <> 0 Then ' <> "" And txttransf <> "0" Then
            maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
            asse = "5 dbo_INGCOMPRAMOVIBANC"
            DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", val(txtcodcuenta), "S", "Transf. " & "Prov. " & ObtenerDescripcion("Prov", UpROV.codigo), _
                Fecha, "E", 0, s2n(txttransf), "O/P", val(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z
        End If
                    
        'SI PAGO CON CHEQUES PROPIOS
        'If txtimpcheques <> "" And txtimpcheques <> "0" Then
        asse = "8 cheques"
        If uCheques.Total > 0 Then
            If ExistenPropios Then
                
                maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
                
                With uCheques
                    For x = 1 To .rows
                        If .chPropio(x) Then
                        
                            If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                                If uCheques.chNroInt(x) = 0 Then
                                    numIntP = nuevoCodigo("chq_Comp")
                                    ' cargo por 1ra vez
                                    DataEnvironment1.dbo_INGRESOCHEQUERA numIntP, 0, uCheques.chNumero(x), uCheques.chBancCod(x), uCheques.chCuenta(x), _
                                              0, 0, "", 0, "C", 0, 0, Date, UsuarioSistema!codigo, 0, 0, 1
                                    uCheques.chSetearNroInt x, numIntP
                                End If
                            End If
                        
                        
                        
                            asse = "6 dbo_INGCOMPRAMOVIBANC"
                            'mod lito 20/7/6  cuenta = la del cheque
                            cueche = s2n(obtenerDeSQL("select cuentabancaria from chq_comp where codigo = " & uCheques.chNroInt(x)))
                            
                            DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", cueche, "L", "O/P " & txtopago & "Prov. " & UpROV.DESCRIPCION _
                                , Fecha, "P", .chNroInt(x), .chMonto(x), "O/P", val(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                            'INCREMENTO EL AUTOMATICO DE MOVIBANC
                            maximobanc = maximobanc + 1
                        End If
                    Next
                End With
            End If
        End If
                    
        'SI PAGO CON CHEQUES DE TERCEROS
         If uCheques.Total > 0 Then
            If ExistenTerceros Then
                
                maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
                
                With uCheques
                    For x = 1 To .rows
                        If Not .chPropio(x) Then
                        
                            'aca meto cheque, despues proximo stored lo saco
                            If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                                If uCheques.chNroInt(x) = 0 Then
                                    numIntT = nuevoCodigo("cheques", "nroint")
                                    ' cargo por 1ra vez
                                    DataEnvironment1.dbo_INGCHEQUESTERCEROS "A", numIntT, .chFecha(x), .chNumero(x), .chMonto(x), val(txtopago), "O/P", Fecha, .chFecha(x), "C", .chBancCod(x), " ", 0, Date, UsuarioActual(), iddoc, iddoc
                                    uCheques.chSetearNroInt x, numIntT
                                End If
                            End If
                        
                            'INGCOMPRAMOVIBANC es igual al STORE del INGCHEQUEMOVIBANC
                            'mod lito 20/7/6  cuenta = 0, no debe aparecer como mov bancario
                            'raul 23/10/07 cuenta=cuenta del cheue, si debe apararecer el mov como que salio el cheque
                            Dim D_Q_C
                            D_Q_C = obtenerDeSQL("select dep_cuenta from cheques where nroint =" & .chNroInt(x))
                            If D_Q_C = 0 Then
                                If MsgBox("El cheque " & .chNroInt(x) & " no tiene cuenta de deposito." & Chr(13) & "Por favor indique una a continuacion, gracias.", vbInformation + vbYesNo) = vbYes Then
                                    D_Q_C = frmBuscar.MostrarSql("select c.codigo as [CODIGO], c.banco as [BANCO - Nº],b.descripcion as  [NOMBRE  ],c.numero as [CUENTA - Nº] from ctasbank c inner join bancosgrales b on c.banco=b.codigo where c.activo=1", , "Cuentas bancarias", " - ")
                                    DataEnvironment1.Sistema.Execute "update cheques set dep_cuenta = " & D_Q_C & " where nroint = " & .chNroInt(x)
                                Else
                                    D_Q_C = 0
                                End If
                            End If
                            asse = "7 dbo_INGCOMPRAMOVIBANC"
                            DataEnvironment1.dbo_INGCOMPRAMOVIBANC "A", D_Q_C, "S", "O/P " & txtopago & "Prov. " & UpROV.DESCRIPCION _
                                , Fecha, "C", .chNroInt(x), .chMonto(x), "O/P", val(txtopago), maximobanc, iddoc, Date, UsuarioSistema!codigo, z
                            'INCREMENTO EL AUTOMATICO DE MOVIBANC
                            maximobanc = maximobanc + 1
                        End If
                    Next
                End With
            End If
        End If
                                
        
        'ACA EMPIEZA LAS ALTAS A MOVICAJA
        'SI PAGO EN EFECTIVO
        If s2n(txtefectivo) <> 0 Then ' txtefectivo <> "" And txtefectivo <> "0" Then
        
            maximocaja = nuevoCodigo("movicaja", "movimiento")
            valorcuenta = verCuentaContableCaja(val(txtcodcaja))

            asse = "8a) asiento efectivo"
            'haber EFECTIVO
            AsientoCompra.AgregarItem valorcuenta, 0, s2n(txtefectivo)  ', sComprobante
            AsientoCompra.AcumularItem CuentaProv, s2n(txtefectivo), 0, sComprobante
            sComprobante = ""
            
            asse = "8 dbo_INGCOMPRAMOVICAJA"
            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", val(txtcodcaja), maximocaja, "E", "E", s2n(txtefectivo), "O/P " & txtopago & "Prov. " & UpROV.codigo, _
                Fecha, 0, UpROV.codigo, "O/P", val(txtopago), valorcuenta, 0, _
                iddoc, Date, UsuarioSistema!codigo, z
        
        End If
                    
        'SI REALIZO UNA TRANSFERENCIA
        If s2n(txttransf) <> 0 Then ' "" And txttransf <> "0" Then
        
            maximocaja = nuevoCodigo("movicaja", "movimiento")
        
            rs.Open "select cuenta_con from Ctasbank where codigo = " & val(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
            If Not rs.EOF Then
                valorcuentacon = rs!cuenta_con
            Else
                valorcuentacon = ""
            End If
            rs.Close
'            Set rs = Nothing
                            
            maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
                            
            asse = "9 dbo_INGCOMPRAMOVICAJA"
            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "T", "E", s2n(txttransf), "O/P " & txtopago & "Prov. " & UpROV.codigo, _
                Fecha, 0, UpROV.codigo, "O/P", val(txtopago), valorcuentacon, maximobanc, _
                iddoc, Date, UsuarioSistema!codigo, z
                    
                'haber  TRANSFERENCIA
                AsientoCompra.AcumularItem obtenerDeSQL("select  cuenta_con from ctasbank where activo = 1 and codigo = '" & x2s(s2n(txtcodcuenta)) & "' "), 0, s2n(txttransf)
                AsientoCompra.AcumularItem CuentaProv, s2n(txttransf), 0, sComprobante
                sComprobante = ""
        End If
                   
                   
        'SI PAGO CON CHEQUES PROPIOS
        'If txtimpcheques <> "" And txtimpcheques <> "0" Then
        If uCheques.Total > 0 Then
            If ExistenPropios Then
                                                                        
                maximocaja = nuevoCodigo("movicaja", "movimiento")
                                                                                        
                rs.Open "select cuenta_con from Ctasbank where codigo = " & val(txtcodcuenta) & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
                If Not rs.EOF Then
                    valorcuentacon = rs!cuenta_con
                Else
                    valorcuentacon = ""
                End If
                rs.Close


                maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
                
                
                With uCheques
                    For x = 1 To .rows
                        If .chPropio(x) Then
                            asse = "10 dbo_INGCOMPRAMOVICAJA"
                            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "P", "E", .chMonto(x), "O/P " & txtopago & "Prov. " & UpROV.codigo _
                                , Fecha, .chNroInt(x), UpROV.codigo, "O/P", val(txtopago), valorcuentacon, maximobanc _
                                , iddoc, Date, UsuarioSistema!codigo, z
    
                            'fechapropio = Month(CDate(FrmCheques.grillapropios.TextMatrix(x, 3))) & "/" & Day(CDate(FrmCheques.grillapropios.TextMatrix(x, 3))) & "/" & Year(CDate(FrmCheques.grillapropios.TextMatrix(x, 3)))
    
                            asse = "11 dbo_INGCOMPRACHEQUEPROPIO"
                            DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "A", .chNroInt(x), .chFecha(x), .chMonto(x) _
                                , val(txtopago), "REC", UpROV.codigo, "T", .chFecha(x), Fecha, Date, UsuarioSistema!codigo, 0, 0, 1, z, 1
    
                            'INCREMENTO EL AUTOMATICO DE MOVIBANC
                            maximobanc = maximobanc + 1
                            
                            
                            'haber CHEQUE PROPIO
                            AsientoCompra.AcumularItem uCheques.chCuenta(x), 0, uCheques.chMonto(x), "ch " & uCheques.chNumero(x)
                            AsientoCompra.AcumularItem CuentaProv, uCheques.chMonto(x), 0, sComprobante
                            sComprobante = ""
                        End If
                    Next
                End With
            End If
        End If
                   
                   
        'SI PAGO CON CHEQUES TERCEROS
        'If txtimpcheques <> "" And txtimpcheques <> "0" Then
        If uCheques.Total > 0 Then
            If ExistenTerceros Then
                
                maximocaja = nuevoCodigo("movicaja", "movimiento")
                
                valcartera = CuentaParam(ID_Cuenta_M_CH_CARTERA)
                
                maximobanc = nuevoCodigo("MOVIBANC", "movbanco")
                
                With uCheques
                    For x = 1 To .rows
                        If Not .chPropio(x) Then
                            asse = "13  dbo_INGCOMPRAMOVICAJA"
                            DataEnvironment1.dbo_INGCOMPRAMOVICAJA "A", 0, maximocaja, "C", "E", .chMonto(x), "O/P " & txtopago & "Prov. " & UpROV.codigo _
                                , Fecha, .chNroInt(x), UpROV.codigo, "O/P", val(txtopago), valcartera, maximobanc _
                                , iddoc, Date, UsuarioSistema!codigo, z
                            
                            asse = "14  dbo_INGCOMPRACHEQUETERCEROS"
                            DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "A", .chNroInt(x), 0, "", UpROV.codigo, val(txtopago), 0 _
                                , Fecha, "T", "O/P", Date, UsuarioSistema!codigo, 0, 0, 1, 1, z
                            
                            'INCREMENTO EL AUTOMATICO DE MOVIBANC
                            maximobanc = maximobanc + 1
                            
                                'Haber Cheques 3ros
                                AsientoCompra.AcumularItem CuentaParam(ID_Cuenta_M_CH_CARTERA), 0, uCheques.chMonto(x), "ch " & uCheques.chNumero(x)
                                AsientoCompra.AcumularItem CuentaProv, uCheques.chMonto(x), 0, sComprobante
                                sComprobante = ""
                        End If
                    Next
                End With

            End If
        End If
        
        ProcesoGrillas s2n(txtopago), "REC", AsientoCompra, iddoc
                         
'        ImprimirOrden
        ' asiento  '  VA ACA , solo cuando es O/P, no IMP
                       
                       
        asse = "16 incrnumpago"
        DataEnvironment1.dbo_INCREMENTONUMOPAGO val(txtopago)
  
        FrmCostosYContable.LimpioControles
        
    End If
    
        
       If siAsiento("AsientosPagos") Then AsientoCompra.Grabar iddoc
    
    'DE_CommitTrans
    midDoc = iddoc
    Graba_OP_IMP = True
    '********************************************
    
'On Error GoTo UfaOpImprimeNuevo
    
    asse = "17 a imprimir "
    ImprimirOrden
    
    MsgBox "Operación Realizada con éxito", vbOKOnly
    'Graba_OP_IMP = True
    

    If RealizoPagoACuenta Then
        Dim Riddoc As String, Pacuenta As Long
        Dim AsientoPAC As New Asiento, impPac As Double, TextoAsientoPAC As String
        impPac = s2n(tmp)
        Pacuenta = nSinNull(obtenerDeSQL("select max(nrodoc) from transcom where tipodoc='RAC'")) + 1
        Riddoc = NuevoDocumento("RAC", s2n(Pacuenta, 0), UpROV.codigo, s2n(Pacuenta, 0))
        
        'TextoAsientoPAC = "RAC " & Pacuenta
        'AsientoPAC.Nuevo "Pago " & uProv.DESCRIPCION, fecha, "PAC"
        'AsientoPAC.AgregarItem CuentaParam(ID_Cuenta_P_ANTICIP_A_PROV), s2n(impPac), 0, TextoAsientoPAC
        
        DataEnvironment1.dbo_INGTRANSCOM "A", Fecha, UpROV.codigo, UpROV.DESCRIPCION, "0", "RAC", Pacuenta _
          , s2n(impPac), s2n(impPac), s2n(impPac), 0, 0, 0, 0, 0, Date, UsuarioActual, Date, 0, 1, Riddoc, 0, 0, 0, 0, 0, _
          "0", "0", "0", "0"
        
        'txtopago = nuevoCodigoOP 'nSinNull(obtenerDeSQL("select max(nrodoc) from transcom where tipodoc='RAC'")) + 1
        
    End If
    
fin:
    Set rs = Nothing
    Exit Function
UfaNUEVO:
    DE_RollbackTrans
    uCheques.resetNroIntPropios
            
    ufa "Err al grabar", "Nueva OP/IMP " & " OP : " & txtopago & " Assertion : " & asse ', Err
    Resume fin
UfaOpImprimeNuevo:
    ufa "Err al imprimir", "Nueva OP/IMP " & " OP : " & txtopago & " Assertion : " & asse ', Err
    Resume fin
End Function

Function ChequeaChq() As Boolean
    Dim x As Long
    Dim inter As Long
    
    ChequeaChq = True
    If uCheques.Total > 0 Then
        With uCheques
            For x = 1 To .rows
                If .chPropio(x) Then
                    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                        If .chNroInt(x) = 0 Then
                            inter = s2n(obtenerDeSQL("select codigo from chq_comp where nro = " & .chNumero(x) & " and banco=" & .chBancCod(x)))
                            If inter > 0 Then
                                MsgBox "El cheque Nro." & .chNumero(x) & " existe con interno " & inter & ", por lo que debe seleccionarlo.", , "ATENCION"
                                ChequeaChq = False
                                'Exit Function
                            End If
                        End If
                    End If
                Else
                    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False Then
                        If .chNroInt(x) = 0 Then
                            inter = s2n(obtenerDeSQL("select nroint from cheques where nro = " & .chNumero(x) & " and banco_nro=" & .chBancCod(x)))
                            If inter > 0 Then
                                MsgBox "El cheque Nro." & .chNumero(x) & " existe con interno " & inter & ", por lo que debe seleccionarlo.", , "ATENCION"
                                ChequeaChq = False
                                'Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End With
    End If
End Function

Sub ImprimirOrden()

    '*************************************
    'ACA MANDO LA IMPRESION DE LA ORDEN
    Dim str1 As String, Empresa As String
    Dim r As Long
    Dim str, SqlRetencion, direccion, Localidad As String
    Dim rsTemp As New ADODB.Recordset
    'Dim fecha As Variant
    Dim Fecha As Date, tdoc As String
    Dim rsempresa As New ADODB.Recordset
    If ON_ERROR_HABILITADO Then On Error GoTo UFAimprimir
    
    Dim sfecha As String
    Dim docTipo As String
    
    sfecha = Format(FrmOrdenPago.Fecha, "dd/mm/yyyy")
    'RptOrdenPago.Restart
    
'    rs2.Open "select fij_emp1 from datos", daTaenvironment1.Sistema, adOpenStatic, adLockOptimistic
'    If Not rs2.EOF Then
'        empresa = rs2!FIJ_EMP1
'    End If
'
'    rs2.Close
'    Set rs2 = Nothing
    
    'Probando SubReports
    
    
    
    '----------mod tabla temp----------
    'Cargo las facturas para imprimir
    'daTaenvironment1.Sistema.Execute "delete from ordenpagotemp"
    'daTaenvironment1.Sistema.Execute "delete from chequesordentemp"
'    If stblOrdenPagoTmp = "" Then
        stblOrdenPagoTmp = TablaTempCrear(tt_OrdenPagoTemp)
'    Else
'        daTaenvironment1.Sistema.Execute "delete from " & stblOrdenPagoTmp
'    End If
    
'    If stblChequesOPtmp = "" Then
        stblChequesOPtmp = TablaTempCrear(tt_ChequeOPtmp)
'    Else
'        daTaenvironment1.Sistema.Execute "delete from " & stblChequesOPtmp
'    End If
    '----------mod tabla temp----------
    
    
    For r = 1 To fac.rows - 1
        fac.Row = r
        If Trim(fac.TextMatrix(r, 3)) <> "" Then
            rsTemp.Open "Select fecha from transcom where tipodoc='FAC' and nrodoc=" & fac.TextMatrix(r, 0) & " and codpr = " & UpROV.codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
                rsTemp.Close
                rsTemp.Open "Select fecha from compras where tipodoc='FAC' and nrodoc=" & fac.TextMatrix(r, 0) & " and codpr = " & UpROV.codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            'Else
            '    fecha = #1/1/1900#
            End If
            If rsTemp.EOF Then
                rsTemp.Close
                rsTemp.Open "Select fecha from gastosboletas where tipodoc='BOL' and nrodoc=" & fac.TextMatrix(r, 0) & " and codpr = " & UpROV.codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            'Else
            '    fecha = #1/1/1900#
            End If
            
            Fecha = rsTemp!Fecha
'            DataEnvironment1.Sistema.Execute "insert into " & stblOrdenPagoTmp & " (tipodoc, nrodoc, fecha, saldo, pagado, ultimosaldo) " _
'                & " values ('FAC', " & fac.TextMatrix(r, 0) & ", " & ssFecha(fecha) & ", " & Replace(fac.TextMatrix(r, 2), ", ", ".") & ", " & Replace(s2n(fac.TextMatrix(r, 3)), ",", ".") & "," & Replace(s2n(s2n(fac.TextMatrix(r, 2)) - s2n(fac.TextMatrix(r, 3))), ",", ".") & ")"
            
            If UCase(fac.TextMatrix(r, 1)) = "TOTAL-BOL" Or UCase(fac.TextMatrix(r, 1)) = "SALDO-BOL" Then
                docTipo = "BOL"
            Else
                docTipo = "FAC"
            End If
            
            DataEnvironment1.Sistema.Execute "insert into " & stblOrdenPagoTmp & " (tipodoc, nrodoc, fecha, saldo, pagado, ultimosaldo) " _
                & " values ('" & docTipo & "', " & fac.TextMatrix(r, 0) & ", " & ssFecha(Fecha) & ", " & x2s(fac.TextMatrix(r, 2)) & ", " & x2s(s2n(fac.TextMatrix(r, 3))) & ", " & x2s(s2n(s2n(fac.TextMatrix(r, 2)) - s2n(fac.TextMatrix(r, 3)))) & ")"
            
            rsTemp.Close
            Set rsTemp = Nothing
        End If
    Next r
    
    'Cargo los debitos para imprimir
    For r = 1 To debito.rows - 1
        debito.Row = r
        If Trim(debito.TextMatrix(r, 4)) <> "" Then
            rsTemp.Open "Select fecha,tipodoc from transcom where (tipodoc='N/D' or tipodoc='APD') and nrodoc=" & debito.TextMatrix(r, 1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
                rsTemp.Close
                rsTemp.Open "Select fecha,tipodoc from compras where (tipodoc='N/D' or tipodoc='APD') and nrodoc=" & debito.TextMatrix(r, 1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'                fecha = rsTemp!fecha
'            Else
'                fecha = #1/1/1900#
            End If
            Fecha = rsTemp!Fecha
            DataEnvironment1.Sistema.Execute "insert into " & stblOrdenPagoTmp & " (tipodoc,nrodoc,fecha,saldo,pagado,ultimosaldo) " _
                & " values ('" & rsTemp!TIPODOC & "'," & debito.TextMatrix(r, 1) & ", " & ssFecha(Fecha) & ", " & x2s(debito.TextMatrix(r, 3)) & "," & x2s(s2n(debito.TextMatrix(r, 4))) & "," & x2s(s2n(s2n(debito.TextMatrix(r, 3)) - s2n(debito.TextMatrix(r, 4)))) & ")"
            rsTemp.Close
            Set rsTemp = Nothing
        End If
    Next r
    
    'Cargo los creditos para imprimir
    For r = 1 To credito.rows - 1
        credito.Row = r
        If Trim(credito.TextMatrix(r, 4)) <> "" Then
            rsTemp.Open "Select fecha,tipodoc from transcom where (tipodoc='N/C' or tipodoc='APC') and nrodoc=" & credito.TextMatrix(r, 1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
'                fecha = rsTemp!fecha
                rsTemp.Close
                rsTemp.Open "Select fecha, tipodoc from compras where (tipodoc='N/C' or tipodoc='APC') and nrodoc=" & credito.TextMatrix(r, 1), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'            Else
'                fecha = 1 / 1 / 1900
            End If
            tdoc = rsTemp!TIPODOC
            Fecha = rsTemp!Fecha
            DataEnvironment1.Sistema.Execute "insert into " & stblOrdenPagoTmp & " (tipodoc, nrodoc, fecha, saldo, pagado,ultimosaldo) values ( '" & tdoc & "'," & credito.TextMatrix(r, 1) & "," & ssFecha(CDate(Fecha)) & "," & Replace(credito.TextMatrix(r, 3), ",", ".") & "," & Replace(s2n(credito.TextMatrix(r, 4)), ",", ".") & "," & x2s(s2n(s2n(credito.TextMatrix(r, 3)) - s2n(credito.TextMatrix(r, 4)))) & ")"
            rsTemp.Close
            Set rsTemp = Nothing
        End If
    Next r
    
    'Cargo los pagos a cuenta para imprimir
    For r = 1 To rac.rows - 1
        rac.Row = r
        If Trim(rac.TextMatrix(r, 3)) <> "" Then
            rsTemp.Open "Select fecha from transcom where tipodoc='RAC' and nrodoc=" & rac.TextMatrix(r, 0), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rsTemp.EOF Then
                rsTemp.Close
                rsTemp.Open "Select fecha from compras where tipodoc='RAC' and nrodoc=" & rac.TextMatrix(r, 0), DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'            Else
'                fecha = #1/1/1900#
            End If
            
'            tdoc = rsTemp!TIPODOC
            Fecha = rsTemp!Fecha
            DataEnvironment1.Sistema.Execute "insert into " & stblOrdenPagoTmp & " (tipodoc, nrodoc, fecha, saldo, pagado, ultimosaldo ) values ( 'RAC'," & rac.TextMatrix(r, 0) & "," & ssFecha(Fecha) & "," & Replace(rac.TextMatrix(r, 2), ",", ".") & "," & Replace(s2n(rac.TextMatrix(r, 3)), ",", ".") & "," & Replace(s2n(s2n(rac.TextMatrix(r, 2)) - s2n(rac.TextMatrix(r, 3))), ",", ".") & ")"
            rsTemp.Close
            Set rsTemp = Nothing
        End If
    Next r
    
'''    'Cargo los Cheques
'''    For r = 1 To grilla.rows - 1
'''        grilla.row = r
'''        If Trim(grilla.TextMatrix(r, 3)) <> "" Then
'''            daTaenvironment1.Sistema.Execute "insert into " & stblChequesOPtmp & " (nroint,banco,cheque,importe,fecha,propio)values('" & grilla.TextMatrix(r, 1) & "','" & grilla.TextMatrix(r, 2) & "','" & grilla.TextMatrix(r, 3) & "'," & grilla.TextMatrix(r, 4) & "," & ssFecha(CDate(grilla.TextMatrix(r, 5))) & ",'" & IIf(grilla.TextMatrix(r, 0) = "Propio", "P", "T") & "')"
'''        End If
'''    Next r
    'Cargo los Cheques
    With uCheques
    For r = 1 To .rows
        DataEnvironment1.Sistema.Execute "insert into " & stblChequesOPtmp _
            & " (nroint, banco, cheque, importe, fecha, propio) values( " _
            & .chNroInt(r) & ", '" & .chBancDes(r) & "', '" & .chNumero(r) & "', " & x2s(.chMonto(r)) & ", " & ssFecha(.chFecha(r)) & ", '" & IIf(.chPropio(r), "P", "T") & "')"
    Next r
    End With

    
    str = "select * from " & stblChequesOPtmp
    RptDetalledeValores.data1.Connection = DataEnvironment1.Sistema
    RptDetalledeValores.data1.Source = str
    
    str1 = "select * from " & stblOrdenPagoTmp
    'rsempresa.Open "select nombrelogofull froLblRetGananciam datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    'rsempresa.Close
    'Set rsempresa = Nothing
    
    RptOrdenPago.LblRetGanancia = Format(uRetCompras.retgan, "#,##0.00")
    RptOrdenPago.LblretIB = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPago.Data.Connection = DataEnvironment1.Sistema
    RptOrdenPago.Data.Source = str1
    RptOrdenPago.lblfecha = sfecha 'Date
    If txttot <> 0 Then
      RptOrdenPago.lblTitulo = "ORDEN DE PAGO Nº " & txtopago
    Else
      RptOrdenPago.lblTitulo = "IMPUTACION Nº " & txtopago
    End If
    RptOrdenPago.lblproveedor = "A la orden de " & UpROV.DESCRIPCION
    RptOrdenPago.lblvalor = "Por la cantidad de pesos " & NroEnLetras(txttot)
    RptOrdenPago.lblefectivo = Format(s2n(txtefectivo), "#,##0.00")
'    RptOrdenPago.lblpie.caption = FrmPrincipal.lblNombreEmpresa.caption
    RptOrdenPago.lblcheques = Format(uCheques.Total, "#,##0.00")
    RptOrdenPago.lbltransf = Format(s2n(txttransf), "#,##0.00")
    RptOrdenPago.SubReport1.object = RptDetalledeValores
    With RptOrdenPago
        .fieFecha.DataField = "FECHA"
        .fieTipoDoc.DataField = "TIPODOC"
        .fieNroDoc.DataField = "NRODOC"
        .fieSaldo.DataField = "SALDO"
        .fiePagado.DataField = "PAGADO"
        .fieNvoSaldo.DataField = "ULTIMOSALDO"
    End With
    direccion = sSinNull(obtenerDeSQL("select direccion from prov where codigo = " & UpROV.codigo & " "))
    Localidad = sSinNull(obtenerDeSQL("select localidad from prov where codigo = " & UpROV.codigo & " "))
 If uRetCompras.retgan > 0 Then
    RptOrdenPagoConstRet_IG.DataImp_Ganancia.Connection = DataEnvironment1.Sistema
    RptOrdenPagoConstRet_IG.DataImp_Ganancia.Source = str1
    RptOrdenPagoConstRet_IG.lblfecha = sfecha 'date
    RptOrdenPagoConstRet_IG.LblRegimen_IG = uRetCompras.IG_Tipo
    RptOrdenPagoConstRet_IG.txtProveedor = UpROV.DESCRIPCION
    RptOrdenPagoConstRet_IG.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
    RptOrdenPagoConstRet_IG.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConstRet_IG.RG_PagosTotalMes = s2n(uRetCompras.RG_PagosTotalMes, 2, True)
    RptOrdenPagoConstRet_IG.retgan = s2n(uRetCompras.retgan, 2, True)
    RptOrdenPagoConstRet_IG.retganEnPesos = enletras(uRetCompras.retgan)
    RptOrdenPagoConstRet_IG.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
    RptOrdenPagoConstRet_IG.Txtop = Format(txtopago, "00000000")
    
    RptOrdenPagoConsRet_IG_calculo.lblfecha = sfecha 'Date
    RptOrdenPagoConsRet_IG_calculo.txtProveedor = UpROV.DESCRIPCION
    RptOrdenPagoConsRet_IG_calculo.TxtDomicilioProv = direccion & "   " & Localidad
    RptOrdenPagoConsRet_IG_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
    RptOrdenPagoConsRet_IG_calculo.RG_PagosTotalMes = s2n(uRetCompras.RG_PagosTotalMes, 2, True)
    RptOrdenPagoConsRet_IG_calculo.RG_MinimoNoImponible = uRetCompras.RG_MinimoNoImponible
    RptOrdenPagoConsRet_IG_calculo.RG_TxtFormula = uRetCompras.RG_TxtFormula '**************************
    RptOrdenPagoConsRet_IG_calculo.retgan = s2n(uRetCompras.retgan, 2, True)
    RptOrdenPagoConsRet_IG_calculo.RG_PagosAnterioresMes = s2n(uRetCompras.RG_PagosAnterioresMes, 2, True)
    RptOrdenPagoConsRet_IG_calculo.RG_PagosRetAnteriores = s2n(uRetCompras.RG_PagosRetAnteriores, 2, True)
    
    RptOrdenPagoConsRet_IG_calculo.NroCertificado = Format(VerNroCertifGan(midDoc), "0001-00000000")
    RptOrdenPagoConsRet_IG_calculo.LblRetGanPesos = enletras(uRetCompras.retgan)
    RptOrdenPagoConsRet_IG_calculo.Pago_Fecha = Format(Abs(CDbl(uRetCompras.RG_PagosAnterioresMes) - CDbl(uRetCompras.RG_PagosTotalMes)), "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.Total_Imponible = Format(CDbl(uRetCompras.RG_PagosRetAnteriores) + CDbl(uRetCompras.retgan), "#,##0.00")
    RptOrdenPagoConsRet_IG_calculo.Printer.Copies = 1
    If txttot <> 0 Then
       RptOrdenPagoConstRet_IG.Printer.Copies = 2
      Else
       RptOrdenPagoConstRet_IG.Printer.Copies = 1
    End If
    RptOrdenPagoConstRet_IG.Restart
    RptOrdenPagoConsRet_IG_calculo.Restart
    
    If PREVIEW_IMPRESIONES Then
        RptOrdenPagoConstRet_IG.Image1.Picture = FrmPrincipal.imgLogoSimple
        RptOrdenPagoConstRet_IG.Label1 = VerParametro(BS_DIRECCION_EMPRESA)
        RptOrdenPagoConstRet_IG.Label2 = VerParametro(BS_CUIT_EMPRESA)
        RptOrdenPagoConstRet_IG.Printer.PaperSize = vbPRPSA4 'hoja A4=9
        RptOrdenPagoConstRet_IG.Show
        RptOrdenPagoConsRet_IG_calculo.Image1.Picture = FrmPrincipal.imgLogoSimple
        RptOrdenPagoConsRet_IG_calculo.Label2 = VerParametro(BS_DIRECCION_EMPRESA)
        RptOrdenPagoConsRet_IG_calculo.Label3 = VerParametro(BS_CUIT_EMPRESA)
        RptOrdenPagoConsRet_IG_calculo.Printer.PaperSize = vbPRPSA4 'hoja A4=9
        RptOrdenPagoConsRet_IG_calculo.Show
    Else
        RptOrdenPagoConstRet_IG.PrintReport False
        RptOrdenPagoConsRet_IG_calculo.PrintReport False
    End If
 End If
 If uRetCompras.retIB > 0 Then
    RptOrdenPagoConstRet_IB.DataImp_IB.Connection = DataEnvironment1.Sistema
    RptOrdenPagoConstRet_IB.DataImp_IB.Source = str1
    RptOrdenPagoConstRet_IB.lblfecha = sfecha 'Date
    RptOrdenPagoConstRet_IB.LblRegimen_IIBB = uRetCompras.IB_Tipo
    RptOrdenPagoConstRet_IB.txtProveedor = UpROV.DESCRIPCION
    RptOrdenPagoConstRet_IB.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConstRet_IB.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
    RptOrdenPagoConstRet_IB.txtNroIIBB = obtenerDeSQL("select NumIIBB from prov where codigo = " & UpROV.codigo & " ")
    RptOrdenPagoConstRet_IB.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
    RptOrdenPagoConstRet_IB.retgan = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConstRet_IB.retganEnPesos = enletras(uRetCompras.retIB)
    RptOrdenPagoConstRet_IB.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
    RptOrdenPagoConstRet_IB.Txtop = Format(txtopago, "00000000")
    
    RptOrdenPagoConsRet_IB_calculo.lblfecha = sfecha 'Date
    RptOrdenPagoConsRet_IB_calculo.txtProveedor = UpROV.DESCRIPCION
    RptOrdenPagoConsRet_IB_calculo.TxtDomicilioProv = direccion & "    " & Localidad
    RptOrdenPagoConsRet_IB_calculo.txtCuit = obtenerDeSQL("select cuit from prov where codigo = " & UpROV.codigo & " ")
    RptOrdenPagoConsRet_IB_calculo.RG_PagosTotalMes = Format(uRetCompras.RG_PagosTotalMes, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.IB_TxtFormula = uRetCompras.IB_TxtFormula
    RptOrdenPagoConsRet_IB_calculo.retIB = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.IB_base = Format(uRetCompras.IB_base, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.retIB1 = Format(uRetCompras.retIB, "#,##0.00")
    RptOrdenPagoConsRet_IB_calculo.NroCertificado = Format(VerNroCertifIIBB(midDoc), "0001-00000000")
    RptOrdenPagoConsRet_IB_calculo.LblRetIIBBPesos = enletras(uRetCompras.retIB)
    RptOrdenPagoConsRet_IB_calculo.Printer.Copies = 1
    If txttot <> 0 Then
       RptOrdenPagoConstRet_IB.Printer.Copies = 2
      Else
       RptOrdenPagoConstRet_IB.Printer.Copies = 1
    End If
    RptOrdenPagoConstRet_IB.Restart
    RptOrdenPagoConsRet_IB_calculo.Restart
    If PREVIEW_IMPRESIONES Then
        RptOrdenPagoConstRet_IB.Image1.Picture = FrmPrincipal.imgLogoSimple
        RptOrdenPagoConstRet_IB.Label2 = VerParametro(BS_DIRECCION_EMPRESA)
        RptOrdenPagoConstRet_IB.Label3 = VerParametro(BS_CUIT_EMPRESA)
        RptOrdenPagoConstRet_IB.Printer.PaperSize = vbPRPSA4 'hoja A4=9
        RptOrdenPagoConstRet_IB.Show
        RptOrdenPagoConsRet_IB_calculo.Image1.Picture = FrmPrincipal.imgLogoSimple
        RptOrdenPagoConsRet_IB_calculo.Label2 = VerParametro(BS_DIRECCION_EMPRESA)
        RptOrdenPagoConsRet_IB_calculo.Label3 = VerParametro(BS_CUIT_EMPRESA)
        RptOrdenPagoConsRet_IB_calculo.Printer.PaperSize = vbPRPSA4 'hoja A4=9
        RptOrdenPagoConsRet_IB_calculo.Show
    Else
        RptOrdenPagoConstRet_IB.PrintReport False
        RptOrdenPagoConsRet_IB_calculo.PrintReport False
    End If
 End If
 
    If txttot = 0 Then
      RptOrdenPago.Printer.Copies = 1
     Else
      RptOrdenPago.Printer.Copies = Abs(VerParametro(BS_PRINTCOPIASOP))
    End If
    RptOrdenPago.Restart
    
   
    If PREVIEW_IMPRESIONES Then
        RptOrdenPago.Printer.PaperSize = vbPRPSA4 'hoja A4=9
        RptOrdenPago.Show
    Else
        RptOrdenPago.PrintReport False
    End If
    
fin:
    Exit Sub
UFAimprimir:
    ufa "error en la impresión", Me.Name
    Resume fin:
'Print uRetCompras.retgan
' 58,3
'Print uRetCompras.RG_Coeficiente
' 0,02
'Print uRetCompras.RG_MinimoNoImponible
'12000
'Print uRetCompras.RG_PagosTotalMes
'14915
'Print uRetCompras.RG_TxtFormula
' 2915 * 2%
'Print uRetCompras.IB_TxtFormula
'14915 *  %1,75
'Print uRetCompras.IB_Base
'400
'**************************************************
'Print uRetCompras.IB_Coef
' 1,75
'Print uRetCompras.retIB
' 261,01
End Sub



Private Sub ProcesoGrillas(Valor As Long, Tipo As String, asi As Asiento, iddoc As Long)
    Dim x As Long
    Dim esboleta As Boolean, cade As String, codbol As Long, signo As String
    'SI HAY ALGO PAGO EN FAC
    If fac.TextMatrix(1, 0) <> "" Then
        For x = 1 To fac.rows - 1
            'TRANSCOM Y COMPRAS
            If fac.TextMatrix(x, 3) <> "" Then
'                daTaenvironment1.dbo_INGRESOALISTADOTEMP "FAC", fac.TextMatrix(x, 0), obtengofecha("FAC", fac.TextMatrix(x, 0)), s2n(fac.TextMatrix(x, 2)), s2n(fac.TextMatrix(x, 3)), s2n(fac.TextMatrix(x, 2)) - s2n(fac.TextMatrix(x, 3))
                If UCase(fac.TextMatrix(x, 1)) = "TOTAL-BOL" Or UCase(fac.TextMatrix(x, 1)) = "SALDO-BOL" Then
                    esboleta = True
                Else
                    esboleta = False
                End If
                
                If esboleta Then
                    codbol = obtenerDeSQL("select id from gastosboletas where nrodoc=" & s2n(fac.TextMatrix(x, 0)) & " and codpr=" & UpROV.codigo)
                    If s2n(fac.TextMatrix(x, 3)) < 0 Then
                        signo = ""
                    Else
                        signo = ""
                    End If
                    cade = "update GastosBoletas set saldo=saldo-(" & x2s(fac.TextMatrix(x, 3)) & ") where id =" & codbol
                    DataEnvironment1.Sistema.Execute cade
                Else
                    If (s2n(fac.TextMatrix(x, 2)) - s2n(fac.TextMatrix(x, 3))) = 0 Then
                        DataEnvironment1.dbo_MODIFICOSALDOYPASOREG UpROV.codigo, "FAC", fac.TextMatrix(x, 0)
                    Else
                        DataEnvironment1.dbo_MODIFICOSALDOTRANS UpROV.codigo, "FAC", val(fac.TextMatrix(x, 0)), (s2n(fac.TextMatrix(x, 2)) - s2n(fac.TextMatrix(x, 3)))
                    End If
                End If
                'RELFNR_C
                If esboleta Then
                    DataEnvironment1.dbo_INGRESORELIMPUT UpROV.codigo, Tipo, Valor, val(fac.TextMatrix(x, 0)), "BOL", s2n(fac.TextMatrix(x, 3)), s2n(fac.TextMatrix(x, 2)), iddoc
                Else
                    DataEnvironment1.dbo_INGRESORELIMPUT UpROV.codigo, Tipo, Valor, val(fac.TextMatrix(x, 0)), "FAC", s2n(fac.TextMatrix(x, 3)), s2n(fac.TextMatrix(x, 2)), iddoc
                End If
                
'                'asiento DEBE
'                asi.AcumularItem CuentaParam(ID_CuentasParam_DEUD_A_PROV), s2n(fac.TextMatrix(X, 3)), 0
            End If
        Next
    End If

    'SI HAY ALGO PAGO EN DEBITOS
    If debito.TextMatrix(1, 0) <> "" Then
        For x = 1 To debito.rows - 1
            'TRANSCOM Y COMPRAS
            If debito.TextMatrix(x, 4) <> "" Then
'                daTaenvironment1.dbo_INGRESOALISTADOTEMP debito.TextMatrix(x, 0), debito.TextMatrix(x, 1), obtengofecha(debito.TextMatrix(x, 0), debito.TextMatrix(x, 1)), s2n(debito.TextMatrix(x, 3)), s2n(debito.TextMatrix(x, 4)), s2n(debito.TextMatrix(x, 3)) - s2n(debito.TextMatrix(x, 4))
                If (s2n(debito.TextMatrix(x, 3)) - s2n(debito.TextMatrix(x, 4))) = 0 Then
                    DataEnvironment1.dbo_MODIFICOSALDOYPASOREG UpROV.codigo, debito.TextMatrix(x, 0), val(debito.TextMatrix(x, 1))
                Else
                    DataEnvironment1.dbo_MODIFICOSALDOTRANS UpROV.codigo, debito.TextMatrix(x, 0), val(debito.TextMatrix(x, 1)), (s2n(debito.TextMatrix(x, 3)) - s2n(debito.TextMatrix(x, 4)))
                End If
                
                'RELFNR_C
                DataEnvironment1.dbo_INGRESORELIMPUT UpROV.codigo, Tipo, Valor, val(debito.TextMatrix(x, 1)), debito.TextMatrix(x, 0), s2n(debito.TextMatrix(x, 4)), s2n(debito.TextMatrix(x, 3)), iddoc
            End If
        Next
    End If

    'SI HAY ALGO PAGO EN CREDITOS
    If credito.TextMatrix(1, 0) <> "" Then
        For x = 1 To credito.rows - 1
            'TRANSCOM Y COMPRAS
            If credito.TextMatrix(x, 4) <> "" Then
'                daTaenvironment1.dbo_INGRESOALISTADOTEMP credito.TextMatrix(x, 0), credito.TextMatrix(x, 1), obtengofecha(credito.TextMatrix(x, 0), credito.TextMatrix(x, 1)), s2n(credito.TextMatrix(x, 3)) * -1, s2n(credito.TextMatrix(x, 4)), s2n(credito.TextMatrix(x, 3)) - s2n(credito.TextMatrix(x, 4))
                If s2n(credito.TextMatrix(x, 3)) - s2n(credito.TextMatrix(x, 4)) = 0 Then
                    DataEnvironment1.dbo_MODIFICOSALDOYPASOREG UpROV.codigo, credito.TextMatrix(x, 0), val(credito.TextMatrix(x, 1))
                Else
                    DataEnvironment1.dbo_MODIFICOSALDOTRANS UpROV.codigo, credito.TextMatrix(x, 0), val(credito.TextMatrix(x, 1)), (s2n(credito.TextMatrix(x, 3)) - s2n(credito.TextMatrix(x, 4)))
                End If
                
                'RELFNR_C
                DataEnvironment1.dbo_INGRESORELIMPUT UpROV.codigo, Tipo, Valor, val(credito.TextMatrix(x, 1)), credito.TextMatrix(x, 0), s2n(credito.TextMatrix(x, 4)), s2n(credito.TextMatrix(x, 3)), iddoc
            End If
        Next
    End If

    'SI HAY ALGO PAGO EN RAC
    If rac.TextMatrix(1, 0) <> "" Then
        For x = 1 To rac.rows - 1
            'TRANSCOM Y COMPRAS
            If rac.TextMatrix(x, 3) <> "" Then
                'daTaenvironment1.dbo_INGRESOALISTADOTEMP "RAC", rac.TextMatrix(x, 0), obtengofecha("RAC", rac.TextMatrix(x, 0)), s2n(rac.TextMatrix(x, 2)), s2n(rac.TextMatrix(x, 3)), s2n(rac.TextMatrix(x, 2)) - s2n(rac.TextMatrix(x, 3))
                If (s2n(rac.TextMatrix(x, 2)) - s2n(rac.TextMatrix(x, 3))) = 0 Then
                    DataEnvironment1.dbo_MODIFICOSALDOYPASOREG UpROV.codigo, "RAC", val(rac.TextMatrix(x, 0))
                Else
                    DataEnvironment1.dbo_MODIFICOSALDOTRANS UpROV.codigo, "RAC", val(rac.TextMatrix(x, 0)), (s2n(rac.TextMatrix(x, 2)) - s2n(rac.TextMatrix(x, 3)))
                End If
                
                'RELFNR_C
                DataEnvironment1.dbo_INGRESORELIMPUT UpROV.codigo, Tipo, Valor, val(rac.TextMatrix(x, 0)), "RAC", s2n(rac.TextMatrix(x, 3)), s2n(rac.TextMatrix(x, 2)), iddoc
                
                'asiento
                asi.AcumularItem CuentaParam(ID_Cuenta_P_ANTICIP_A_PROV), 0, s2n(rac.TextMatrix(x, 3))  ', "RAC " & rac.TextMatrix(x, 0)
                asi.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(rac.TextMatrix(x, 3)), 0, "OP " & txtopago
            End If
        Next
    End If
End Sub

Function obtengofecha(Tipo, Nro) As Date
    Dim rs As New ADODB.Recordset

    rs.Open "select fecha from Transcom where tipodoc='" & Tipo & "' and nrodoc=" & Nro & "", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rs.EOF Then
        obtengofecha = rs!Fecha
    Else
        obtengofecha = 0
    End If
    rs.Close
    Set rs = Nothing
    
End Function

    

'
'Private Sub Iniciogrillacheques()
'    With grilla
'        .Clear
'        .ColWidth(1) = 1500
'        .TextMatrix(0, 0) = "Propio / Tercero"
'        .TextMatrix(0, 1) = "Nro. Interno"
'        .TextMatrix(0, 2) = "Banco"
'        .TextMatrix(0, 3) = "Nro. Cheque"
'        .TextMatrix(0, 4) = "Importe"
'        .TextMatrix(0, 5) = "Fecha"
'        .rows = 2
'    End With
'End Sub
Sub CargoOrden() '(tipo)
    Dim Tipo As String
    Dim rs As New ADODB.Recordset
    Dim re As Variant, sql As String
    Dim i As Long
    
    uCheques.Borrar
'    gRet.Borrar
    LimpioControles
    
    If mBusco = buscoIMP Then
        sql = "select I.fecha as [Fecha            ], I.nro as [Nro. Orden], I.codpr as [Proveedor],p.Descripcion as [Descripcion                         ],p.Cuit as [      Cuit       ], I.iddoc, 0 as [_H_rg] from imppro I inner join prov p on p.codigo=I.codpr where I.activo = 1 and (I.fecha between " & ufDesde.ConvertFecha & " AND " & ufHasta.ConvertFecha & ")  order by I.nro desc "
        Tipo = "IMP"
    ElseIf mBusco = buscoOP Then
        sql = "select I.fecha as [Fecha            ], I.nro as [Nro. Orden], I.codpr as [Proveedor],p.Descripcion as [Descripcion                         ],p.Cuit as [      Cuit       ], I.iddoc , I.RetGanPago as [_H_rg ], I.ibPago as [_H_rib ] from rec_comp I inner join prov p on p.codigo=I.codpr where I.activo = 1 and (I.fecha between " & ufDesde.ConvertFecha & " AND " & ufHasta.ConvertFecha & ")  order by I.nro desc "
        Tipo = "REC"
    Else
        ufa "err prg: tipo busqueda", "cargoorden() OrdenPago" ', Err
    End If
    If frmBuscar.MostrarSql(sql) > "" Then
'        Iniciogrillacheques
        Fecha = frmBuscar.resultado(1)
        txtopago = frmBuscar.resultado(2)
        UpROV.codigo = s2n(frmBuscar.resultado(3))
        'txtprov = ObtenerDescripcion("Prov", Val(txtcodprov))
        midDoc = nSinNull(frmBuscar.resultado(6))
        lblIDDOC = midDoc
        uRetCompras.retgan = s2n(frmBuscar.resultado(7))
        uRetCompras.retIB = s2n(frmBuscar.resultado(8))
        
        If frmBuscar.resultado(1) <> "" Then
            Limpiogrillas
            rs.Open "select * from relfnr_c where iddoc=" & lblIDDOC & " and prov = " & val(frmBuscar.resultado(3)) & " and ndoc = " & val(frmBuscar.resultado(2)) & " and tdoc = '" & Tipo & "' order by ndoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                Select Case rs!tfac
                    Case "FAC":
                        If fac.TextMatrix(1, 0) = "" Then
                            fac.Col = 0
                            fac.TextMatrix(1, 0) = rs!Fact
                            fac.Col = 1
                            If rs!Impor = rs!totaldocu Then
                                fac.TextMatrix(1, 1) = "TOTAL"
                            Else
                                fac.TextMatrix(1, 1) = "SALDO"
                            End If
                            fac.Col = 2
                            fac.TextMatrix(1, 2) = rs!totaldocu
                            fac.Col = 3
                            fac.TextMatrix(1, 3) = rs!Impor
                        Else
                            fac.AddItem rs!Fact & Chr(9) & IIf(rs!Impor = rs!totaldocu, "TOTAL", "SALDO") & Chr(9) & rs!totaldocu & Chr(9) & rs!Impor
                        End If
                    Case "BOL":
                        If fac.TextMatrix(1, 0) = "" Then
                            fac.Col = 0
                            fac.TextMatrix(1, 0) = rs!Fact
                            fac.Col = 1
                            If rs!Impor = rs!totaldocu Then
                                fac.TextMatrix(1, 1) = "TOTAL-BOL"
                            Else
                                fac.TextMatrix(1, 1) = "SALDO-BOL"
                            End If
                            fac.Col = 2
                            fac.TextMatrix(1, 2) = rs!totaldocu
                            fac.Col = 3
                            fac.TextMatrix(1, 3) = rs!Impor
                        Else
                            fac.AddItem rs!Fact & Chr(9) & IIf(rs!Impor = rs!totaldocu, "TOTAL-BOL", "SALDO-BOL") & Chr(9) & rs!totaldocu & Chr(9) & rs!Impor
                        End If
                    Case "N/D", "APD":
                        If debito.TextMatrix(1, 0) = "" Then
                            debito.Col = 0
                            debito.TextMatrix(1, 0) = rs!tfac
                            debito.Col = 1
                            debito.TextMatrix(1, 1) = rs!Fact
                            debito.Col = 2
                            If rs!Impor = rs!totaldocu Then
                                debito.TextMatrix(1, 2) = "TOTAL"
                            Else
                                debito.TextMatrix(1, 2) = "SALDO"
                            End If
                            debito.Col = 3
                            debito.TextMatrix(1, 3) = rs!totaldocu
                            debito.Col = 4
                            debito.TextMatrix(1, 4) = rs!Impor
                        Else
                            debito.AddItem rs!tfac & Chr(9) & rs!Fact & Chr(9) & IIf(rs!Impor = rs!totaldocu, "TOTAL", "SALDO") & Chr(9) & rs!totaldocu & Chr(9) & rs!Impor
                        End If
                    Case "N/C", "APC":
                        If credito.TextMatrix(1, 0) = "" Then
                            credito.Col = 0
                            credito.TextMatrix(1, 0) = rs!tfac
                            credito.Col = 1
                            credito.TextMatrix(1, 1) = rs!Fact
                            credito.Col = 2
                            If rs!Impor = rs!totaldocu Then
                                credito.TextMatrix(1, 2) = "TOTAL"
                            Else
                                credito.TextMatrix(1, 2) = "SALDO"
                            End If
                            credito.Col = 3
                            credito.TextMatrix(1, 3) = rs!totaldocu
                            credito.Col = 4
                            credito.TextMatrix(1, 4) = rs!Impor
                        Else
                            credito.AddItem rs!tfac & Chr(9) & rs!Fact & Chr(9) & IIf(rs!Impor = rs!totaldocu, "TOTAL", "SALDO") & Chr(9) & rs!totaldocu & Chr(9) & rs!Impor
                        End If
                    Case "RAC":
                        If rac.TextMatrix(1, 0) = "" Then
                            rac.Col = 0
                            rac.TextMatrix(1, 0) = rs!Fact
                            rac.Col = 1
                            If rs!Impor = rs!totaldocu Then
                                rac.TextMatrix(1, 1) = "TOTAL"
                            Else
                                rac.TextMatrix(1, 1) = "SALDO"
                            End If
                            rac.Col = 2
                            rac.TextMatrix(1, 2) = rs!totaldocu
                            rac.Col = 3
                            rac.TextMatrix(1, 3) = rs!Impor
                        Else
                            rac.AddItem rs!Fact & Chr(9) & IIf(rs!Impor = rs!totaldocu, "TOTAL", "SALDO") & Chr(9) & rs!totaldocu & Chr(9) & rs!Impor
                        End If
                End Select
                rs.MoveNext
            Wend
            rs.Close
'            Set rs = Nothing
        End If
        
        If Tipo = "REC" Then
            
            'ACA TRAIGO LA FORMA DEL PAGO DE LA ORDEN
'            rs.Open "select * from movicaja where cli_prov = " & uProv.Codigo & " and tipodoc = 'O/P' and nrodoc = " & Val(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
            'rs.Open "select * from movicaja where tipodoc = 'O/P' and nrodoc = " & Val(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            rs.Open "select * from movicaja where codprov = " & uProv.Codigo & " and tipodoc = 'O/P' and nrodoc = " & Val(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            
            '*********************************************
            'el nro O/P es interno, no se puede repetir asi q ya no pregunto por proveedor...
            'OJO en los casos Factura proveedor, no aca !!!
            rs.Open "select * from movicaja where tipodoc = 'O/P' and nrodoc = " & val(txtopago) & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            'Dim i As Long
            'ii = 1
            While Not rs.EOF
                
                Select Case rs!Tipo
                    Case "E":
                        txtefectivo = rs!Importe
                        txtcodcaja = rs!caja
                        txtcaja = ObtenerDescripcionCajas("Cajas", rs!caja)
                    Case "T":
                        txttransf = s2n(obtenerDeSQL("select importe from movibanc where operacion='S' and iddoc=" & midDoc))
                        'txtcodcuenta = ObtenerCuenta("Movicaja", rsmov!NroDoc, rsmov!CODPR)
                        txtcodcuenta = s2n(obtenerDeSQL("select cuenta from movibanc where OPERACION='S' and iddoc=" & midDoc))
                        txtcuenta = ObtenerDescripcionCuentas("ctasbank", s2n(txtcodcuenta))
    
                        'txttransf = rs!Importe
                        'txtcodcuenta = rs!Cuenta
'                        txtcuenta = ObtenerDescripcion("Ctasbank", rs!Cuenta)
                    Case "P":
                        'If txtimpcheques <> "" Then
                        '    txtimpcheques = s2n(txtimpcheques) + rs!Importe
                        'Else
                        '    txtimpcheques = rs!Importe
                        'End If
                        
                        
                        
'''                        If grilla.rows = 2 Then
'''                            grilla.row = 1
'''                            If Trim(grilla.Text) = "" Then
'''                                grilla.TextMatrix(1, 0) = "Propio"
'''                                grilla.TextMatrix(1, 1) = rs!interno
'''                                grilla.TextMatrix(1, 2) = buscodesbanco(rs!interno) '"ACA TENGO QUE BUSCAR EL BANCO POR EL CODIGO DE BANCO QUE LO SACO DE CHQ_COMP"
'''                                grilla.TextMatrix(1, 3) = busconumcheque(rs!interno) '"ACA TENGO QUE BUSCAR EL Nº DE CHEQUE QUE LO SACO DE CHQ_COMP"
'''                                grilla.TextMatrix(1, 4) = rs!importe
'''                                grilla.TextMatrix(1, 5) = rs!fecha
'''                            Else
'''                                grilla.AddItem "Propio" & Chr(9) & rs!interno & Chr(9) & buscodesbanco(rs!interno) & Chr(9) & busconumcheque(rs!interno) & Chr(9) & rs!importe & Chr(9) & rs!fecha
'''                            End If
'''                        Else
'''                            grilla.AddItem "Propio" & Chr(9) & rs!interno & Chr(9) & buscodesbanco(rs!interno) & Chr(9) & busconumcheque(rs!interno) & Chr(9) & rs!importe & Chr(9) & rs!fecha
'''                        End If
                        uCheques.metoCheque uCheques.rows + 1, rs!interno, "P"
                    Case "C":
                        'If txtimpcheques <> "" Then
                        '    txtimpcheques = s2n(txtimpcheques) + rs!Importe
                        'Else
                        '    txtimpcheques = rs!Importe
                        'End If
'''                        If grilla.rows = 2 Then
'''                            grilla.row = 1
'''                            If Trim(grilla.Text) = "" Then
'''                                grilla.TextMatrix(1, 0) = "Propio"
'''                                grilla.TextMatrix(1, 1) = rs!interno
'''                                grilla.TextMatrix(1, 2) = buscodesbanco(rs!interno) '"ACA TENGO QUE BUSCAR EL BANCO POR EL CODIGO DE BANCO QUE LO SACO DE CHQ_COMP"
'''                                grilla.TextMatrix(1, 3) = busconumcheque(rs!interno) '"ACA TENGO QUE BUSCAR EL Nº DE CHEQUE QUE LO SACO DE CHQ_COMP"
'''                                grilla.TextMatrix(1, 4) = rs!importe
'''                                grilla.TextMatrix(1, 5) = rs!fecha
'''                            Else
'''                                grilla.AddItem "Tercero" & Chr(9) & rs!interno & Chr(9) & buscodesbanco(rs!interno) & Chr(9) & busconumcheque(rs!interno) & Chr(9) & rs!importe & Chr(9) & rs!fecha
'''                            End If
'''                        Else
'''                            grilla.AddItem "Tercero" & Chr(9) & rs!interno & Chr(9) & buscodesbanco(rs!interno) & Chr(9) & busconumcheque(rs!interno) & Chr(9) & rs!importe & Chr(9) & rs!fecha
'''                        End If
                    uCheques.metoCheque uCheques.rows + 1, rs!interno, "T"
                End Select
             '   ii = ii + 1
                rs.MoveNext
            Wend
            
'            uRetCompras.retgan = rs!retGanPago
'            uRetCompras.RetIb = rs!RetIb
            
            rs.Close
            
'            Dim tempo
'            tempo = obtenerdesql("select retganpago, retib from rec_comp
            
            'i = 0
'''            rs.Open "select * from ComprasRetenciones where iddoc = '" & midDoc & "' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'''            While Not rs.EOF
'''                'i = i + 1
'''                i = gRet.addRow()
'''                gRet.tx i, gR_IdC, rs!idCuentasParam
'''                gRet.tx i, gR_IMP, rs!Importe
'''                gRet.tx i, gR_CUE, rs!cuenta
'''                gRet.tx i, gR_TIP, sSinNull(obtenerDeSQL("select Descripcion from CuentasParam where id = '" & rs!idCuentasParam & "' "))
'''                rs.MoveNext
'''            Wend
'''            Set rs = Nothing

        End If
        
               
                
        FormadePago (True)
        uMenu.BuscarOK
    Else
        ' no busque...
    End If
    'assert
    'If s2n(txtimpcheques) <> uCheques.Total Then
    '    ufa "Err datos cheques no coinciden ", txtopago
    'End If
End Sub

Function buscodesbanco(interno As Long) As String
    Dim rs As New ADODB.Recordset

    rs.Open "select bancosgrales.descripcion from Chq_comp inner join bancosgrales on chq_comp.banco = bancosgrales.codigo where chq_comp.codigo = " & interno & "", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rs.EOF Then
        buscodesbanco = rs!DESCRIPCION
    Else
        buscodesbanco = ""
    End If
    rs.Close
    Set rs = Nothing
End Function

Function busconumcheque(interno As Long) As Long
    Dim rs As New ADODB.Recordset

    rs.Open "select nro from Chq_comp where chq_comp.codigo = " & interno & "", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rs.EOF Then
        busconumcheque = rs!Nro
    Else
        busconumcheque = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub cmbingresar_Click()
    If txttot <> "" Then
        If (s2n(s2nt(uCheques.Total) + s2nt(txtefectivo) + s2nt(txttransf) + s2nt(uRetCompras.TotalRet)) <> s2n(txttot)) Then
            MsgBox "No coinciden los totales en la forma de pago con el importe total a pagar"
            Exit Sub
        End If
    End If

    FrmCostosYContable.LimpioControles
    FrmCostosYContable.InicioGrilla
    FrmCostosYContable.txtimporte = txttot
    FrmCostosYContable.habilitogrillaCostos (True)
    FrmCostosYContable.txtimptotal = txttot.Text
    FrmCostosYContable.Show
End Sub

'Private Sub cmdBorraRet_Click()
'    If gRet.row() > 0 Then gRet.delRow gRet.row()
'End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_Load()
'    Me.Top = 100
'    Me.Left = 1500
'    SubimeSi800x600
'    gRetIni
    UpROV.ini "Select Descripcion from prov where activo = 1 and codigo = ### ", "Select codigo as [ Proveedor ],cuit as [ Cuit             ], descripcion as [ Nombre                                                                   ] from prov where categ<>2 and activo = 1 order by codigo", False
    uMenu.init True, True, False, True, True
    mBusco = buscoNADA
'    cmbingresar.Visible = gEMPR_ConSistContable
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

'Private Sub gRetIni()
'    Dim sr As String
'    Set gRet = New LiGrilla
'    gRet.init gRetenciones
'
'    sr = strComboGrilla("select descripcion from cuentasparam where usocuenta = '" & ID_UsoCuenta_RETCOM & "'")
'    With gRet
'        gR_NRO = .AddCol(" Nro Ret     ", "H") '"N", 0)*************************
'        gR_FEC = .AddCol(" Fecha Ret   ", "H") '"D")*************************
'        gR_TIP = .AddCol(" Tipo Ret                     ", "B", sr)
'        gR_IMP = .AddCol(" Importe Ret  ", "N", 2)
'        gR_CUE = .AddCol(" Cuenta        ")
'        gR_TIC = .AddCol(" cod ret ", "H") '********************************
'        gR_IdC = .AddCol(" idCuentaParam", "H") '******************************
''        gR_FAC = .AddCol(" Nro Factura ", "H") ' --------------
'    End With
'End Sub
'Private Function RevisarReten(cual, row As Long) As Boolean
'    Dim tempo
'    With gRet
'        tempo = obtenerDeSQL("select codigo, cuenta, id from cuentasparam where descripcion = '" & cual & "' and activo = 1 ") 'and usoCuenta = '" & ID_UsoCuenta_RETCOM & "'")
'        If IsEmpty(tempo) Then
'            .tx row, gR_CUE, ""
'            .tx row, gR_TIC, 0
'            .tx row, gR_IdC, 0
'            RevisarReten = False
'        Else
'            .tx row, gR_CUE, tempo(1)
'            .tx row, gR_TIC, tempo(0)
'            .tx row, gR_IdC, tempo(2)
'            RevisarReten = True
'        End If
'    End With
'End Function


Private Sub fac_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = (Col <> 3)
End Sub

''Private Sub gRet_cambio(ByVal row As Long, ByVal col As Long, txt As String)
''    With gRet
''        If row = .rows - 1 Then .rows = .rows + 1
''
''        If col = gR_NRO Then
''            If s2n(.tx(row, col)) = 0 Then
''                'borro la parte de retencion
''                .tx row, gR_NRO, ""
''                .tx row, gR_CUE, ""
''                .tx row, gR_FEC, ""
''                .tx row, gR_IMP, ""
''                .tx row, gR_TIC, ""
''                .tx row, gR_TIP, ""
''                .tx row, gR_IdC, ""
''            End If
''        End If
''        If col = gR_IMP Then
''            'lblSumaRet = s2n(gRet.suma(gR_IMP))
''            recalpago
''        End If
'''        If col = gR_IdC Then
''
''
''
''
''    End With
''
'''    Select Case col
'''    Case gCUENT
'''        g.tx row, gCDESC, sSinNull(CuentaDescripcion(txt, False, False))
'''    Case gMONTO
'''        recalpago
'''    End Select
''End Sub
'Private Sub recalcFalta()
''    lblSumaCheques = n2r(g.suma(gMONTO))
''    lblSumaRet = n2r(gRet.suma(gR_IMP))
'''    lblFaltaPagar = n2r(s2n(txtTotalRecibo) - s2n(txtefectivo) - s2n(lblSumaCheques) - s2n(lblSumaRet))
''
''    txtTotalRet = s2n(gRet.suma(gR_IMP))
''
''    txtTotalPago = s2n(txtefectivo) + s2n(txtTotalRet) + uCheques.Total
'End Sub

Private Sub recalpago()
    txtTotalRet = uRetCompras.TotalRet 's2n(txtRetGanPago) + s2n(txtRetIIBBPago) 'gRet.suma(gR_IMP)
    'cheques
    txtTotalPago = sumaPagos() 's2n(txtTotalRet) + s2n(txtefectivo) + s2n(uCheques.Total) + s2n(txttransf)
    txtFalta = s2n(s2n(txttot) - s2n(txtTotalPago))
End Sub


Private Sub lblIdDoc_Click()
    frmAsientoManual.mostrar (s2n(lblIDDOC))
End Sub

'Private Sub gRet_Validar(ByVal row As Long, ByVal col As Long, cancel As Boolean)
'    Select Case col
''     Case gR_NRO
'     Case gR_TIP
'        If Not RevisarReten(gRet.EditText, row) Then cancel = True
''     Case gR_IMP
''        If s2n(gRet.EditText) > s2n(gRet.tx(row, gF_SAL)) Then cancel = True
'    End Select
'
'End Sub

Private Sub rac_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = (Col <> 3)
End Sub
Private Sub credito_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = (Col <> 4)
End Sub
Private Sub debito_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = (Col <> 4)
End Sub


Private Sub debito_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = Not VacioONumero(debito.EditText)
End Sub
Private Sub credito_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = Not VacioONumero(credito.EditText)
End Sub
Private Sub rac_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = Not VacioONumero(rac.EditText)
End Sub
Private Sub fac_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = Not VacioONumero(fac.EditText)
End Sub
Private Function VacioONumero(s As String)
    VacioONumero = (IsNumeric(s) Or s = "")
End Function


Private Sub credito_DblClick()
    credito.TextMatrix(credito.Row, 4) = credito.TextMatrix(credito.Row, 3)
    recalculo
End Sub
Private Sub debito_DblClick()
    debito.TextMatrix(debito.Row, 4) = debito.TextMatrix(debito.Row, 3)
    recalculo
End Sub
Private Sub fac_DblClick()
    fac.TextMatrix(fac.Row, 3) = fac.TextMatrix(fac.Row, 2)
    recalculo
End Sub
Private Sub rac_DblClick()
    rac.TextMatrix(rac.Row, 3) = rac.TextMatrix(rac.Row, 2)
    recalculo
End Sub
'
Private Sub fac_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If fac.Text <> "" And s2n(fac.Text) <= s2n(fac.TextMatrix(fac.Row, 2)) Then
        recalculo
    Else
        MsgBox "No debe superar el valor del comprobante.", vbCritical, "ATENCION"
        fac.TextMatrix(fac.Row, 3) = 0
    End If
End Sub
Private Sub rac_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If rac.Text <> "" And s2n(rac.Text) <= s2n(rac.TextMatrix(rac.Row, 2)) Then
        recalculo
    Else
        MsgBox "No debe superar el valor del comprobante.", vbCritical, "ATENCION"
        rac.TextMatrix(rac.Row, 3) = 0
    End If
End Sub
Private Sub credito_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If credito.Text <> "" And s2n(credito.Text) <= s2n(credito.TextMatrix(credito.Row, 3)) Then
        recalculo
    Else
        MsgBox "No debe superar el valor del comprobante.", vbCritical, "ATENCION"
        credito.TextMatrix(credito.Row, 4) = 0
    End If
End Sub
Private Sub debito_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If debito.Text <> "" And s2n(debito.Text) <= s2n(debito.TextMatrix(debito.Row, 3)) Then
        recalculo
    Else
        MsgBox "No debe superar el valor del comprobante.", vbCritical, "ATENCION"
        debito.TextMatrix(debito.Row, 4) = 0
    End If
End Sub
'
Private Sub recalculo()
    CargoTotales
    'cmbforma.Enabled = (s2n(txttot) > 0)
    tabOP.TabEnabled(1) = True '(s2n(txttot) > 0)
End Sub



Public Sub CargarDatos()

    Dim rs As New ADODB.Recordset
    Dim codigo As String

    codigo = Trim(Me.Tag)
    
    If cargar = "CuentasBank" Then
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcuenta = rs!codigo
            txtcuenta = ObtenerDescripcionBancos("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
    End If
    
   If cargar = "Cajas" Then
        rs.Open "select * from Cajas where codigo = " & val(txtcodcaja) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcodcaja = rs!codigo
            txtcaja = rs!responsable
        End If
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub tabOP_Click(PreviousTab As Integer)
    On Error Resume Next
    
    If tabOP.Tab = 1 And s2n(txtefectivo) = 0 Then
        FormadePago (True)
        'txtefectivo.SetFocus
        uRetCompras.SetFocus
    End If
End Sub


Private Sub txtEfectivo_GotFocus()
'    TxtEfectivo = s2n(TxtEfectivo) + s2n(txtFalta) 'txttot
    txtefectivo = nuevoMonto(txtefectivo, s2n(txttot), sumaPagos())
    txtefectivo.SelStart = 0
    txtefectivo.SelLength = Len(txtefectivo.Text)
End Sub

Private Sub txtefectivo_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtEfectivo_LostFocus()
    On Error Resume Next
    Dim efecZero As Boolean, efecTope As Boolean
    Dim efec As Double, tot As Double
    
    efec = s2n(txtefectivo)
    tot = s2n(txttot)
    
    txtefectivo = efec
    efecZero = (efec = 0)
    efecTope = (efec = tot)
    
'    If efec > tot Then
'        MsgBox "El importe en efectivo no puede superar al importe de la orden de pago"
'        txtefectivo.SetFocus
'        Exit Sub
''    Else
''        txtimpcheques = s2n(tot) - s2n(txtefectivo)
'    End If
    
    cmbcaja.enabled = Not efecZero
    txtcodcaja.enabled = Not efecZero
    txtcaja.enabled = Not efecZero
    

    'txtimpcheques.Enabled = Not efecTope:
    uCheques.enabled = Not efecTope
    txttransf.enabled = Not efecTope
    txtcodcuenta.enabled = Not efecTope
    cmbcuenta.enabled = Not efecTope
    cmbingresar.enabled = gEMPR_ConSistContable ' habilito si tiene s contable
'    If s2n(txtEfectivo) = s2n(txttot) Then uMenu.SetFocus
    recalpago
    
End Sub

Private Sub txtcodcaja_GotFocus()
    If Trim$(txtcodcaja) = "" Then txtcodcaja = "1"
    PintoFocoActivo
End Sub
Private Sub txtcodcaja_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub

Private Sub txtcodcaja_LostFocus()
    If txtcodcaja <> "" Then
        txtcaja = ObtenerDescripcionCajas("Cajas", val(txtcodcaja))
        If txtcaja = "" Then
            MsgBox "Código de caja incorrecto"
'            txtcodcaja = "0"
            txtcodcaja.SetFocus
        Else
            cargar = "Cajas"
            CargarDatos
        End If
    End If
End Sub

'Private Sub txtimpcheques_GotFocus()
'    txtimpcheques.SelStart = 0
'    txtimpcheques.SelLength = Len(txtimpcheques.Text)
'End Sub
'
'Private Sub txtimpcheques_KeyPress(KeyAscii As Integer)
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'End Sub

'Private Sub txtimpcheques_LostFocus()
'
'    txtimpcheques = s2n(txtimpcheques)
'
'''    If txtimpcheques <> "" And txtimpcheques <> "0" Then
'''        If FrmCheques.txttotal = "0" Or s2n(txtimpcheques) <> s2n(FrmCheques.txttotal) Then
'''            FrmCheques.txttotalcheques = txtimpcheques
'''            FrmCheques.Mostrar Me
'''        End If
'''    End If
'''
'''    If Val(txtimpcheques) = 0 Then
'''        FrmCheques.grilla.Clear
'''        FrmCheques.txttotal = "0"
'''        FrmCheques.Limpiogrillas
'''        FrmCheques.InicioGrilla
'''    End If
'''
'''    If txtimpcheques = "" Then
'''        txtimpcheques = "0"
'''    End If
    
'''    If s2n(txtimpcheques) + s2n(txtEfectivo) + s2n(txttransf) = txttot Then
'''        txtEfectivo.Enabled = False
'''        txtcodcaja.Enabled = False
'''        cmbcaja.Enabled = False
'''        txttransf.Enabled = False
'''        txtcodcuenta.Enabled = False
'''        cmbcuenta.Enabled = False
'''        If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''    Else
'''        If s2n(txtimpcheques) < s2n(txttot) Then
'''            txtEfectivo.Enabled = True
'''            txtcodcaja.Enabled = True
'''            cmbcaja.Enabled = True
'''            txttransf.Enabled = True
'''            txtcodcuenta.Enabled = True
'''            cmbcuenta.Enabled = True
'''            If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''        Else
'''            MsgBox "El importe en cheques no puede superar al importe de la orden de pago"
'''            Exit Sub
'''        End If
'''    End If
'    txttransf = s2n(s2n(txttot) - s2n(txtefectivo) - s2n(txtimpcheques))
'End Sub

Private Sub txtopago_LostFocus()
    If Not IsEmpty(obtenerDeSQL("select Nro from Rec_Comp where activo = 1 and Nro = " & s2n(txtopago, 0))) Then
        che "Orden de Pago " & txtopago & " ya existe"
    End If
End Sub

'Private Sub txtreten_GotFocus()
''''    txtreten.SelStart = 0
''''    txtreten.SelLength = Len(txttransf.Text)
'    PintoFocoActivo
'End Sub
'
'Private Sub txtreten_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub

'''Private Sub txtreten_LostFocus()
'''
''''''    If txtreten = "" Then
''''''        txtreten = "0"
''''''    End If
'''    txtreten = s2n(txtreten)
'''
'''    If txtreten = txttot Then
''''        txtefectivo.Enabled = False
''''        txtcodcaja.Enabled = False
''''        cmbcaja.Enabled = False
''''        txtimpcheques.Enabled = False
'''        If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''    Else
'''        If s2n(txtreten) < s2n(txttot) Then
''''            txtefectivo.Enabled = True
''''            txtcodcaja.Enabled = True
''''            cmbcaja.Enabled = True
''''            txtimpcheques.Enabled = True
'''            If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''        Else
'''            MsgBox "El importe de retención no puede superar al importe de la orden de pago"
'''            Exit Sub
'''        End If
'''    End If
'''
'''End Sub

Private Sub txttransf_GotFocus()
'''    txttransf.SelStart = 0
'''    txttransf.SelLength = Len(txttransf.Text)
    txttransf = nuevoMonto(txttransf, txttot, sumaPagos())
    PintoFocoActivo
End Sub

Private Sub txttransf_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txttransf_LostFocus()
        txttransf = s2n(txttransf)
        recalpago
        cmbingresar.enabled = gEMPR_ConSistContable ' habilito si tiene s contable
'''        If txttransf <> "" And txttransf <> "0" Then
'''            txtcodcuenta.Enabled = True
'''            cmbcuenta.Enabled = True
'''        End If
'''
'''        If txttransf = "" Then
'''            txttransf = "0"
'''        End If
'''
'''        If txttransf = txttot Then
'''            txtefectivo.Enabled = False
'''            txtcodcaja.Enabled = False
'''            cmbcaja.Enabled = False
'''            txtimpcheques.Enabled = False
'''            If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''        Else
'''            If s2n(txttransf) < s2n(txttot) Then
'''                txtefectivo.Enabled = True
'''                txtcodcaja.Enabled = True
'''                cmbcaja.Enabled = True
'''                txtimpcheques.Enabled = True
'''                If cmbingresar.Enabled = False Then cmbingresar.Enabled = True
'''            Else
'''                MsgBox "El importe en deposito no puede superar al importe de la orden de pago"
'''                Exit Sub
'''            End If
'''        End If
End Sub

Private Sub txtcodcuenta_GotFocus()
'    txtcodcuenta.SelStart = 0
'    txtcodcuenta.SelLength = Len(txtcodcuenta.Text)
    PintoFocoActivo
End Sub

Private Sub txtcodcuenta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtcodcuenta_LostFocus()
    On Error Resume Next
    If txtcodcuenta <> "" Then
        Dim rs As New ADODB.Recordset
        
        rs.Open "select * from Ctasbank where codigo = " & val(txtcodcuenta) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            txtcuenta = ObtenerDescripcion("BancosGrales", rs!Banco) & " - " & rs!numero
        End If
        rs.Close
        Set rs = Nothing
        
        If txtcuenta = "" Then
            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = "0"
            txtcodcuenta.SetFocus
        Else
            cargar = "CuentasBank"
            CargarDatos
        End If
'    Else
''        If txtcodcuenta <> "" Then
'            MsgBox "Codigo de cuenta incorrecto"
'            txtcodcuenta = ""
'            txtcodcuenta.SetFocus
''        End If
    End If
End Sub

Private Sub LimpioControles()
    midDoc = 0
    
    FrmBorrarTxt Me
    lblIDDOC = ""
    
    
 '   txtopago = ""
    Fecha = Date
    txtcodcaja = 1
    txtcaja = ObtenerDescripcionCajas("Cajas", 1)
'    txttotal = ""
    cargar = ""
    'txtcodprov = ""
    'txtprov = ""
    UpROV.codigo = 0
    
'    txtefectivo = "0"
'    txtimpcheques = "0"
'    txttransf = "0"
'    txtreten = "0"
'    txtcodcaja = ""
'    txtcaja = ""
'    txtcodcuenta = ""
'    txtcuenta = ""
'    txttot = "0"
'    tot1 = "0"
'    tot2 = "0"
    Limpiogrillas
    
    uCheques.Borrar
    uRetCompras.retgan = 0
    uRetCompras.retIB = 0
    midDoc = 0
    tabOP.Tab = 0
End Sub

Private Sub HabilitoControles(habilito As Boolean)
    Fecha.enabled = habilito
    fac.Editable = habilito
    debito.Editable = habilito
    credito.Editable = habilito
    rac.Editable = habilito
    uCheques.enabled = habilito
'    gRetenciones.Enabled = habilito
'    txtRetGanPago.Enabled = habilito
'    txtRetIIBBPago.Enabled = habilito
    uRetCompras.enabled = habilito
    fraPago.enabled = habilito
End Sub

Private Sub Limpiogrillas()
    fac.clear
    debito.clear
    credito.clear
    rac.clear
'    grilla.Clear
    InicioGrillas
End Sub

Sub InicioGrillas()
    fac.clear
'    fac.ColWidth(2) = 1500
    fac.TextMatrix(0, 0) = "Nº Comprob."
    fac.TextMatrix(0, 1) = "Pago"
    fac.TextMatrix(0, 2) = "Importe"
    fac.TextMatrix(0, 3) = "A Pagar"
    fac.TextMatrix(0, 4) = "A Ret"
    fac.rows = 2
    
    debito.clear
'    debito.ColWidth(2) = 1500
    debito.TextMatrix(0, 0) = "Tipo Comp."
    debito.TextMatrix(0, 1) = "Nº Comp."
    debito.TextMatrix(0, 2) = "Pago"
    debito.TextMatrix(0, 3) = "Importe"
    debito.TextMatrix(0, 4) = "A Pagar"
    debito.rows = 2
    
    credito.clear
'    credito.ColWidth(2) = 1500
    credito.TextMatrix(0, 0) = "Tipo Comp."
    credito.TextMatrix(0, 1) = "Nº Comp."
    credito.TextMatrix(0, 2) = "Pago"
    credito.TextMatrix(0, 3) = "Importe"
    credito.TextMatrix(0, 4) = "A Pagar"
    credito.TextMatrix(0, 5) = "Neto"
    credito.rows = 2
    
    rac.clear
'    rac.ColWidth(2) = 1500
    rac.TextMatrix(0, 0) = "Nº Comprob."
    rac.TextMatrix(0, 1) = "Pago"
    rac.TextMatrix(0, 2) = "Importe"
    rac.TextMatrix(0, 3) = "A Pagar"
    rac.rows = 2
    
'    Iniciogrillacheques
End Sub

Private Sub CargoGrillas()
    Dim rs As New ADODB.Recordset, facRET As Double, yaPagado As Double, sQue As Double

    Limpiogrillas

    rs.Open "select * from Transcom where codpr = " & UpROV.codigo & " and saldo <> 0 and tipodoc = 'FAC' and activo = 1 order by nroDoc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    facRET = 0
    yaPagado = 0
    While Not rs.EOF
        If rs!saldo = rs!Total Then
            If fac.TextMatrix(1, 0) = "" Then
                fac.TextMatrix(1, 0) = rs!NroDoc
                fac.TextMatrix(1, 1) = "TOTAL"
                fac.TextMatrix(1, 2) = s2n(rs!saldo)
                fac.TextMatrix(1, 4) = IIf(s2n(rs!Neto) = 0, s2n(rs!EXENTO), s2n(rs!Neto))
            Else
                fac.AddItem rs!NroDoc & Chr(9) & "TOTAL" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9) & IIf(s2n(rs!Neto) = 0, s2n(rs!EXENTO), s2n(rs!Neto))
            End If
        Else
            yaPagado = rs!Total - rs!saldo
            sQue = IIf(s2n(rs!Neto) = 0, s2n(rs!EXENTO), s2n(rs!Neto))
            If yaPagado >= sQue Then
                facRET = 0
            Else
                facRET = sQue - yaPagado 'lo dejo para ver que hace
                'pero enzo quiere que la retencion se haga la primera y no en cualquier pago
                'ahora quiere que si el primer pago fue imputacion se calcule ret
                If PagoAnterior(rs!NroDoc, UpROV.codigo) = imputacion Then
                    facRET = rs!saldo / 1.21
                Else
                    facRET = 0
                End If
            End If
            If fac.TextMatrix(1, 0) = "" Then
                fac.TextMatrix(1, 0) = rs!NroDoc
                fac.TextMatrix(1, 1) = "SALDO"
                fac.TextMatrix(1, 2) = s2n(rs!saldo)
                fac.TextMatrix(1, 4) = s2n(facRET)
            Else
                fac.AddItem rs!NroDoc & Chr(9) & "SALDO" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9) & s2n(facRET)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
''***************boletas
    rs.Open "select * from Gastosboletas where codpr = " & UpROV.codigo & " and saldo <> 0 and tipodoc = 'BOL' and activo = 1 order by nroDoc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    facRET = 0
    yaPagado = 0
    While Not rs.EOF
        If rs!saldo = rs!Total Then
            If fac.TextMatrix(1, 0) = "" Then
                fac.TextMatrix(1, 0) = rs!NroDoc
                fac.TextMatrix(1, 1) = "TOTAL-BOL"
                fac.TextMatrix(1, 2) = s2n(rs!saldo)
                fac.TextMatrix(1, 4) = 0 'IIf(s2n(rs!Neto) = 0, s2n(rs!EXENTO), s2n(rs!Neto))
            Else
                'fac.AddItem rs!NroDoc & Chr(9) & "TOTAL-BOL" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9) & IIf(s2n(rs!Neto) = 0, s2n(rs!EXENTO), s2n(rs!Neto))
                fac.AddItem rs!NroDoc & Chr(9) & "TOTAL-BOL" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9)
            End If
        Else
            If fac.TextMatrix(1, 0) = "" Then
                fac.TextMatrix(1, 0) = rs!NroDoc
                fac.TextMatrix(1, 1) = "SALDO-BOL"
                fac.TextMatrix(1, 2) = s2n(rs!saldo)
                fac.TextMatrix(1, 4) = s2n(facRET)
            Else
                fac.AddItem rs!NroDoc & Chr(9) & "SALDO-BOL" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9) & s2n(facRET)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
''*********************fin boletas
    
    rs.Open "select * from Transcom where codpr = " & UpROV.codigo & " and saldo <> 0 and (tipodoc = 'N/D' or tipodoc = 'APD') and activo = 1 order by NroDoc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    While Not rs.EOF
        If rs!saldo = rs!Total Then
            If debito.TextMatrix(1, 0) = "" Then
                debito.TextMatrix(1, 0) = rs!TIPODOC
                debito.TextMatrix(1, 1) = rs!NroDoc
                debito.TextMatrix(1, 2) = "TOTAL"
                debito.TextMatrix(1, 3) = s2n(rs!saldo)
            Else
                debito.AddItem rs!TIPODOC & Chr(9) & rs!NroDoc & Chr(9) & "TOTAL" & Chr(9) & s2n(rs!saldo)
            End If
        Else
            If debito.TextMatrix(1, 0) = "" Then
                debito.TextMatrix(1, 0) = rs!TIPODOC
                debito.TextMatrix(1, 1) = rs!NroDoc
                debito.TextMatrix(1, 2) = "SALDO"
                debito.TextMatrix(1, 3) = s2n(rs!saldo)
            Else
                debito.AddItem rs!TIPODOC & Chr(9) & rs!NroDoc & Chr(9) & "SALDO" & Chr(9) & s2n(rs!saldo)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    rs.Open "select * from Transcom where codpr = " & UpROV.codigo & " and saldo <> 0 and (tipodoc = 'N/C' or tipodoc = 'APC') and activo = 1 order by NroDoc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    While Not rs.EOF
        If rs!saldo = rs!Total Then
            If credito.TextMatrix(1, 0) = "" Then
                credito.TextMatrix(1, 0) = rs!TIPODOC
                credito.TextMatrix(1, 1) = rs!NroDoc
                credito.TextMatrix(1, 2) = "TOTAL"
                credito.TextMatrix(1, 3) = s2n(rs!saldo)
            Else
                credito.AddItem rs!TIPODOC & Chr(9) & rs!NroDoc & Chr(9) & "TOTAL" & Chr(9) & s2n(rs!saldo)
            End If
        Else
            If credito.TextMatrix(1, 0) = "" Then
                credito.TextMatrix(1, 0) = rs!TIPODOC
                credito.TextMatrix(1, 1) = rs!NroDoc
                credito.TextMatrix(1, 2) = "SALDO"
                credito.TextMatrix(1, 3) = s2n(rs!saldo)
                credito.TextMatrix(1, 5) = s2n(rs!Neto)
            Else
                credito.AddItem rs!TIPODOC & Chr(9) & rs!NroDoc & Chr(9) & "SALDO" & Chr(9) & s2n(rs!saldo) & Chr(9) & Chr(9) & s2n(rs!Neto)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
    
    rs.Open "select * from Transcom where  codpr = " & UpROV.codigo & " and saldo <> 0 and tipodoc = 'RAC' and activo = 1 order by NroDoc", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    While Not rs.EOF
        If rs!saldo = rs!Total Then
            If rac.TextMatrix(1, 0) = "" Then
                rac.TextMatrix(1, 0) = rs!NroDoc
                rac.TextMatrix(1, 1) = "TOTAL"
                rac.TextMatrix(1, 2) = s2n(rs!saldo)
            Else
                rac.AddItem rs!NroDoc & Chr(9) & "TOTAL" & Chr(9) & s2n(rs!saldo)
            End If
        Else
            If rac.TextMatrix(1, 0) = "" Then
                rac.TextMatrix(1, 0) = rs!NroDoc
                rac.TextMatrix(1, 1) = "SALDO"
                rac.TextMatrix(1, 2) = s2n(rs!saldo)
            Else
                rac.AddItem rs!NroDoc & Chr(9) & "SALDO" & Chr(9) & s2n(rs!saldo)
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
'    HabilitoGrillas (True)
End Sub

Sub FormadePago(habilito As Boolean)
    
    fraPago.Visible = habilito
    
    txtefectivo.Visible = habilito
'    txtimpcheques.Visible = habilito
    txttransf.Visible = habilito
    txtcodcaja.Visible = habilito
    cmbcaja.Visible = habilito
    cmbcuenta.Visible = habilito
    txtcaja.Visible = habilito
    txtcodcuenta.Visible = habilito
    txtcuenta.Visible = habilito
    'txtreten.Visible = habilito
    Label4.Visible = habilito
    Label5.Visible = habilito
    Label6.Visible = habilito
    'Label7.Visible = habilito
    Label8.Visible = habilito
    Label26.Visible = habilito
'    Line3.Visible = habilito
'    grilla.Visible = habilito
    uCheques.Visible = habilito
'    gRetenciones.Visible = habilito
'    Line1.Visible = habilito
'    Label28.Visible = habilito
'    cmbingresar.Visible = habilito And gEMPR_ConSistContable
End Sub
Private Sub debito_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub credito_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub rac_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub fac_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        If KeyAscii <> 13 Then
            KeyAscii = 0
        End If
    End If
End Sub



Private Sub CargoTotales()
    Dim x As Long
    Dim suma1 As Double, suma2 As Double, sumoNeto As Double, sumoNetoC As Double, FueIMP As Boolean

    suma1 = 0
    suma2 = 0
    sumoNeto = 0
    sumoNetoC = 0
    For x = 1 To fac.rows - 1
        If fac.TextMatrix(x, 3) <> "" Then
            suma1 = suma1 + s2n(fac.TextMatrix(x, 3))
            sumoNeto = sumoNeto + s2n(fac.TextMatrix(x, 4))
        End If
    Next

    
    For x = 1 To debito.rows - 1
        If debito.TextMatrix(x, 4) <> "" Then suma1 = suma1 + s2n(debito.TextMatrix(x, 4))
    Next
    
    For x = 1 To credito.rows - 1
        If credito.TextMatrix(x, 4) <> "" Then
            suma2 = suma2 + s2n(credito.TextMatrix(x, 4))
            sumoNetoC = sumoNetoC + s2n(credito.TextMatrix(x, 5))
        End If
    Next
    
    For x = 1 To rac.rows - 1
        If rac.TextMatrix(x, 3) <> "" Then suma2 = suma2 + s2n(rac.TextMatrix(x, 3))
    Next
    
    tot1 = suma1
    tot2 = suma2
    txttot = s2n(suma1 - suma2)
    
    If es_IMP = False Then
        cargoRetGanCompra sumoNeto - sumoNetoC, s2n(txttot)
    Else
        cargoRetGanCompra 0, 0
    End If
    recalpago
 
End Sub

Private Function PagoAnterior(doc As Long, prov As Long) As queFue
Dim rsPago As New ADODB.Recordset
Dim cIMP As Long, cPAG As Long, i As Long
cIMP = 0: cPAG = 0
rsPago.Open "select tdoc from relfnr_c where tfac='FAC' and fact=" & doc & " and prov=" & prov, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsPago
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If UCase(!tdoc) = "IMP" Then
                cIMP = cIMP + 1
            ElseIf UCase(!tdoc) = "REC" Then
                cPAG = cPAG + 1
            End If
            .MoveNext
        Next
    End If
End With
Set rsPago = Nothing
If cPAG = 0 And cIMP <> 0 Then
    PagoAnterior = imputacion
ElseIf cIMP = 0 And cPAG <> 0 Then
    PagoAnterior = pago
ElseIf (cIMP + cPAG) >= 2 Then
    PagoAnterior = AMBAS
ElseIf cIMP = 0 And cPAG = 0 Then
    PagoAnterior = NADA
End If
End Function

Private Function cargoRetGanCompra(sNetoIIBB As Double, sNetoG As Double)
'    Dim i As Long, s
'    i = gRet.Buscar(gR_IdC, ID_CuentasParam_RET_GAN_CPRA)
'    If i = 0 Then
'        i = gRet.PrimerVacio(gR_TIP) 'addRow()
'        s = CuentaParamDesc(ID_CuentasParam_RET_GAN_CPRA)
'        gRet.tx i, gR_TIP, s
'        RevisarReten s, i
'    End If
'
'    'gRet.tx i, gR_IMP,
    'txtRetGanPago = CalculaRetGan(uProv.codigo, s2n(txttot), fecha)
'    txtTotalRet = uRetCompras.Calcular(uProv.codigo, s2n(txttot) / (1 + ProvCoefIVA(uProv.codigo)), fecha)
    txtTotalRet = uRetCompras.Calcular(UpROV.codigo, s2n(sNetoIIBB), s2n(sNetoG), Fecha)
End Function

Private Function nChequesDebitados() As Long
    On Error Resume Next
    nChequesDebitados = obtenerDeSQL(" select count(codigo) as nn from Chq_Comp  " _
        & " where tipoDoc = 'O/P' and nrodoc = " & s2n(txtopago) & " and activo = 1 and estado = 'B' ")
End Function
                 
 Private Function ObtenerImporte(TipoFac As String, nFac As Long) As Double ' , tipoDoc As String) As Double
    On Error Resume Next
    Dim ssql As String, docIMP_o_REC As String
    
    Select Case mBusco
    Case buscoOP:  docIMP_o_REC = "REC"
    Case buscoIMP: docIMP_o_REC = "IMP"
    Case Else:     ufa "err prg", "case REC IMP" & Me.Name ', 0
    End Select
    
    ssql = "select impor from relfnr_c where prov = " & UpROV.codigo _
          & " and tfac = '" & TipoFac _
          & "' and fact = " & nFac _
          & " and tdoc = '" & docIMP_o_REC _
          & "' and ndoc = " & s2n(txtopago)
    ObtenerImporte = s2n(obtenerDeSQL(ssql))
    If ObtenerImporte = 0 Then MsgBox "importe 0 en " & TipoFac & " " & nFac
 End Function

Private Function eliminaOP_IMP() As Boolean 'cmdeliminar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim

    Dim NumOPAGO As Long, Importe As Variant, CodProv As Long
    Dim asse As String, x As Long, ssql As String, supplierTable As String
    Dim rs As New ADODB.Recordset
        
    ' Hay cheques debitados ?
    If nChequesDebitados() > 0 Then
        che "No se puede dar de baja " & vbCrLf & nChequesDebitados & " cheques propios ya fueron debitados"
        Exit Function
    End If
    supplierTable = IIf(mBusco = buscoIMP, "IMPPRO", "REC_COMP")
    If obtenerDeSQL("select activo from " & supplierTable & " where iddoc=" & lblIDDOC) = False Then
        MsgBox "La " & IIf(mBusco = buscoIMP, "imputacion", " orden de pago") & " ya esta eliminada.", , "ATENCION"
        Exit Function
    End If
    
    NumOPAGO = s2n(txtopago)
    CodProv = UpROV.codigo
   
'------- ************* TRANSACCION AQUI ************** ------------------
    DE_BeginTrans
                                   
    'ACA DOY DE BAJA A LOS ENCABEZADOS EN IMPPRO O BIEN EN REC_COMP
    asse = "00  --Baja imppro o rec_comp"
    If mBusco = buscoIMP Then '  "IMP" Then
        DataEnvironment1.dbo_DOYDEBAJAIMPPRO Date, NumOPAGO, CodProv
    Else '"O/P"
        DataEnvironment1.dbo_DOYDEBAJARECCOMP Date, NumOPAGO, CodProv
    End If
            
    'LAS BUSQUEDAS QUE HAGO A CONTINUACION SON A PARTIR DE LAS GRILLAS YA CARGADAS DE ESA ORDEN,
    'ES DECIR, TODOS LOS COMPROBANTES QUE FUERON PAGADOS YA SEA EN SU TOTALIDAD O NO
    
    ' loop cada grilla
    '    paso compras a transcom,                         ' de COMPRAS a TRANSCOM,  si no ta en compras, q me importa, ya estaria en transcom
    '    busco importe en relfnr_c, actualizo saldos, borro relfnr_c
    
    'BUSCO EN FAC
    Dim cadd As String
    asse = "10 FAC - "
    If fac.TextMatrix(1, 0) <> "" Then
        For x = 1 To fac.rows - 1
            If UCase(fac.TextMatrix(x, 1)) = "TOTAL-BOL" Or UCase(fac.TextMatrix(x, 1)) = "SALDO-BOL" Then
                cadd = "update gastosboletas set saldo=saldo+(" & x2s(fac.TextMatrix(x, 3)) & ") where codpr=" & UpROV.codigo & " and nrodoc=" & fac.TextMatrix(x, 0)
                DataEnvironment1.Sistema.Execute cadd
            Else
                DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, "FAC", val(fac.TextMatrix(x, 0))
                Importe = ObtenerImporte("FAC", s2n(fac.TextMatrix(x, 0)))
                If s2n(fac.TextMatrix(x, 2)) < Importe Then Importe = s2n(fac.TextMatrix(x, 2))
                If Importe > 0 Then ' sumo al saldo lo q hay en relfnr_c
                    DataEnvironment1.dbo_sumoSALDOTRANS CodProv, "FAC", val(fac.TextMatrix(x, 0)), Importe
                    'DataEnvironment1.dbo_DOYDEBAJARELFNRC CodProv, "FAC", Val(fac.TextMatrix(x, 0)), NumOPAGO
                End If
            End If
        Next
    End If
        
    'BUSCO EN DEBITO
    asse = "11 DEBITO - "
    If debito.TextMatrix(1, 0) <> "" Then
        For x = 1 To debito.rows - 1
            DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, debito.TextMatrix(x, 0), s2n(debito.TextMatrix(x, 1))
            Importe = ObtenerImporte(debito.TextMatrix(x, 0), s2n(debito.TextMatrix(x, 1)))
            If Importe > 0 Then ' sumo al saldo lo q hay en relfnr_c
                DataEnvironment1.dbo_sumoSALDOTRANS CodProv, debito.TextMatrix(x, 0), val(debito.TextMatrix(x, 1)), Importe
                'DataEnvironment1.dbo_DOYDEBAJARELFNRC CodProv, debito.TextMatrix(x, 0), Val(debito.TextMatrix(x, 1)), NumOPAGO
            End If
        Next
    End If
            
    'BUSCO EN CREDITO
    asse = "12 CREDITO - "
    If credito.TextMatrix(1, 0) <> "" Then
        For x = 1 To credito.rows - 1
            DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, credito.TextMatrix(x, 0), val(credito.TextMatrix(x, 1))
            Importe = ObtenerImporte(credito.TextMatrix(x, 0), s2n(credito.TextMatrix(x, 1)))
            If Importe > 0 Then ' sumo al saldo lo q hay en relfnr_c
                DataEnvironment1.dbo_sumoSALDOTRANS CodProv, credito.TextMatrix(x, 0), credito.TextMatrix(x, 1), Importe
                'DataEnvironment1.dbo_DOYDEBAJARELFNRC CodProv, credito.TextMatrix(x, 0), Val(credito.TextMatrix(x, 1)), NumOPAGO
            End If
        Next
    End If
            
    'BUSCO EN RAC
    asse = "13 Rac - "
    If rac.TextMatrix(1, 0) <> "" Then
        For x = 1 To rac.rows - 1
            DataEnvironment1.dbo_DECOMPRASATRANSCOM CodProv, "RAC", rac.TextMatrix(x, 0)
            Importe = ObtenerImporte("RAC", s2n(rac.TextMatrix(x, 0)))
            If Importe > 0 Then ' sumo al saldo lo q hay en relfnr_c
                DataEnvironment1.dbo_sumoSALDOTRANS CodProv, "RAC", rac.TextMatrix(x, 0), Importe
                'DataEnvironment1.dbo_DOYDEBAJARELFNRC CodProv, "RAC", Val(rac.TextMatrix(x, 0)), NumOPAGO
            End If
        Next
    End If
               
    asse = "20 MOVIBANC    "
    rs.Open "select * from Movibanc where fecha = " & ssFecha(Fecha) & " and tipdoc = 'O/P' and nrodoc = " & NumOPAGO & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.dbo_INGCOMPRAMOVIBANC "B", 0, "", "", _
            0, "", 0, 0, "", 0, rs!MovBanco, midDoc, Date, UsuarioSistema!codigo, 1
        rs.MoveNext
    Wend
    rs.Close
    
    asse = "21 Movi Caja..."
    rs.Open "select * from Movicaja where tipodoc = 'O/P' and nrodoc = " & NumOPAGO & " and activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.dbo_INGCOMPRAMOVICAJA "B", 0, rs!movimiento, "", "", 0, "", _
            0, 0, 0, "", 0, "", 0, midDoc, Date, UsuarioSistema!codigo, 1
        rs.MoveNext
    Wend
    rs.Close
            
    asse = "30 Cheques... "
    DataEnvironment1.dbo_INGCOMPRACHEQUEPROPIO "B", 0, 0, 0, NumOPAGO, "REC", 0, "", 0, 0, 0, 0, Fecha, UsuarioSistema!codigo, 0, 1, 0
    DataEnvironment1.dbo_INGCOMPRACHEQUETERCEROS "B", 0, 0, "", CodProv, NumOPAGO, 0, _
        0, "", "O/P", 0, 0, Fecha, UsuarioSistema!codigo, 0, 1, 1

'    DataEnvironment1.dbo_INGCOMPRASDETALLE "B", 0, 0, 0, "", "", "", CodProv, "O/P", NumOPAGO, 0
    
    'retenciones
    DataEnvironment1.Sistema.Execute "delete from ComprasRetenciones where iddoc = '" & midDoc & "' "
    
    
    If midDoc <> 0 Then
        'borr doc y asiento
        If Not BorroDocumento(midDoc) Then 'AsientoBaja_idDoc(mIdDoc) Then
            ufa "err no se pudo borrar doc OP", "middoc " & midDoc
            DE_RollbackTrans
            Exit Function
        End If
    End If
    
            
    grabaBitacora "B", NumOPAGO, "OPago..."
    
    DE_CommitTrans
'------- ************* TRANSACCION hasta AQUI ************** ------------------

    eliminaOP_IMP = True

fin:
    Set rs = Nothing
    Exit Function
UFAelim:
    DE_RollbackTrans
    ufa "Err al eliminar ", Me.Name & asse ', Err
    Resume fin
End Function

Function ExistenTerceros() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If Not uCheques.chPropio(x) Then
            ExistenTerceros = True
            Exit Function
        End If
    Next x
End Function
Function ExistenPropios() As Boolean
    Dim x As Long
    For x = 1 To uCheques.rows
        If uCheques.chPropio(x) Then
            ExistenPropios = True
            Exit Function
        End If
    Next x
End Function
'''Function existenPropios() As Boolean
'''    Dim x As Integer
'''    Dim cont As Integer
'''
'''    For x = 1 To FrmCheques.grillapropios.rows - 1
'''        If FrmCheques.grillapropios.TextMatrix(x, 5) <> "" Then
'''            cont = cont + 1
'''        End If
'''    Next
'''    If cont > 0 Then
'''        existenPropios = True
'''    Else
'''        existenPropios = False
'''    End If
'''End Function
'''
'''
'''Function existenTerceros() As Boolean
'''    Dim x As Integer, cont As Integer
'''
'''    For x = 1 To FrmCheques.grillaterceros.rows - 1
'''        If FrmCheques.grillaterceros.TextMatrix(x, 5) <> "" Then
'''            cont = cont + 1
'''        End If
'''    Next
'''
'''    If cont > 0 Then
'''        existenTerceros = True
'''    Else
'''        existenTerceros = False
'''    End If
'''End Function

Private Sub uCheques_cambio()
    recalpago
End Sub
Private Sub uCheques_LostFocus()
    recalpago
    cmbingresar.enabled = gEMPR_ConSistContable ' habilito si tiene s contable
End Sub
Private Sub uRetCompras_cambio(Total As Double)
    recalpago
End Sub
Private Sub uRetCompras_LostFocus()
    recalpago
End Sub


Private Sub sumo()
    Dim TotalND, TotalNC, Total As Double
    Dim i As Long
    For i = 1 To fac.rows - 1
       TotalND = TotalND + s2n(fac.TextMatrix(i, 3))
    Next
    For i = 1 To debito.rows - 1
       TotalND = TotalND + s2n(debito.TextMatrix(i, 4))
    Next
    For i = 1 To rac.rows - 1
       TotalNC = TotalNC + s2n(rac.TextMatrix(i, 3))
    Next
    For i = 1 To credito.rows - 1
       TotalNC = TotalNC + s2n(credito.TextMatrix(i, 4))
    Next
    tot1.Text = TotalND
    tot2.Text = TotalNC
    txttot = s2n(tot1 - tot2) '+ s2n(txtRetGan)
End Sub

Private Function nuevoMonto(cuanto, Total, sumPagos)
    cuanto = s2n(cuanto)
    
    nuevoMonto = cuanto + (s2n(Total) - sumPagos)
    recalpago
End Function
Private Function sumaPagos()
    sumaPagos = s2n(txtTotalRet) + s2n(txtefectivo) + uCheques.Total + s2n(txttransf)
End Function

' ******************* MENU *************************
'
Private Sub uMenu_AceptarAlta()
    If Graba_OP_IMP() Then
        uMenu.AceptarOk
    End If
End Sub
Private Sub uMenu_BorrarControles()
    LimpioControles
    Limpiogrillas
    uCheques.Borrar
    
    FrmCostosYContable.LimpioControles
    
    mBusco = IIf(optQueBusco(0), buscoOP, buscoIMP)
End Sub
Private Sub uMenu_Buscar()
    mBusco = IIf(optQueBusco(0), buscoOP, buscoIMP)
    CargoOrden
    sumo
End Sub

Private Sub uMenu_eliminar()
    If eliminaOP_IMP() Then
        uMenu.EliminarOK
        che "La " & IIf(mBusco = buscoIMP, "imputacion", " orden de pago") & " se ha anulado correctamente"
    End If
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoControles sino
    'If Not SiNo Then FormadePago False ' solo deshabilita, no habilita
'    HabilitoGrillas SiNo
    UpROV.enabled = sino
End Sub
Private Sub uMenu_Imprimir()
    ImprimirOrden
End Sub

Private Sub uMenu_Nuevo()
    txtopago = nuevoCodigoOP()
    
    Label2.caption = "Orden de Pago Nº"
    'fecha.SetFocus
'    cargrillas = True
    UpROV.enabled = True
    txtefectivo = 0
'    gRet.rows = 2
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub

'******************************************************

Private Sub uProv_cambio(codigo As Variant)
    Limpiogrillas
    If codigo > 0 And uMenu.estado = ucbEditando Then
        txttipoiva = ObtenerIvaProv("Prov", s2n(codigo))
        CargoGrillas
        ver_IIBBc
    End If
End Sub

Private Function ver_IIBBc()
    Dim Propio As Boolean
    Dim TIVA As Long
    TIVA = obtenerDeSQL("select tipoiva from prov where codigo=" & UpROV.codigo)
    If TIVA = 7 Then
    Else

    
    Propio = obtenerDeSQL("select  conretiibbper from Prov where codigo = " & s2n(UpROV.codigo))
    If Propio = True Then
        If MsgBox("El proveedor que selecciono tiene Retencion de IIBB personal." & Chr(13) & Chr(13) & "¿Desea utilizarlo?", vbCritical + vbYesNo, "Alvertencia") = vbYes Then
            uRetCompras.tieneIIBB = True
        Else
            uRetCompras.tieneIIBB = False
        End If
    Else
        uRetCompras.tieneIIBB = False
    End If
    
    Propio = obtenerDeSQL("select  conretganper from Prov where codigo = " & s2n(UpROV.codigo))
    If Propio = True Then
        If MsgBox("El proveedor que selecciono tiene Retencion de Ganancias personal." & Chr(13) & Chr(13) & "¿Desea utilizarlo?", vbCritical + vbYesNo, "Alvertencia") = vbYes Then
            uRetCompras.tieneGAN = True
        Else
            uRetCompras.tieneGAN = False
        End If
    Else
        uRetCompras.tieneGAN = False
    End If
        
    End If
End Function

'1/11/4
'   rpt amir
'   txtEfectivo_LostFocis
'17/11/4
'   fix amir : s2n en inresar,
'   amir: codrpt
'19/1/5
'   hace Busqueda restringida
'   Busq, orden x numero desc
'   FIX GRAVE no grababa neto en rec_comp
'   fix fechabaja string -puede haber otros?-
'   Fix botones mal sincronizados - ucMenu
'20/1/5
'   REHICE todo eliminacion OP, pero deje la retorcida filosofia original, lee grilla cargadas, busca importes en database
'   Grabacion/Eliminacion tiene control de errores. Ufa.
'9/3/5
'   edicion grillas: FIX dblClic, afteredit, recalcula siempre, controla columna
'   validacion  para editar x nro col
'23/3/5
'   mod por pablo, impresion---
'   falla: obtenerimporte fallaba con nro grandes, pasa a long
'   grillacheque no se borraba al cargar  nueva op
'   QUEDA PEND frmcheques, cheque propio usado queda en grilla, pero con monto hasta q se recarga form.
'   parametro sist contable
'30/3/5
'   quito el horrendo frmCheques
'   fix varios grabacion cheques, habilitacion controles... que todavia esta medio medio
'15/4/5
'    SubimeSi800x600
'11/10/5
'   redondeo total porq daba dif  9 E-13 y fallaba al grabar
'1/6/5
'   fix numero OP funcionNuevoCodigoOP
'16/11/5
'   campo retgan agregado
'

