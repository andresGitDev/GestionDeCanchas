VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmPermisosdeAcceso 
   Caption         =   "Permisos de Usuarios"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "FrmPermisosdeAcceso.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid GrillaPermisos 
      Height          =   3135
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   10335
      _cx             =   18230
      _cy             =   5530
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmPermisosdeAcceso.frx":08CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPermisosdeAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Dim rs As New ADODB.Recordset
Dim i As Long
Dim c As Long
Dim r As Long

Private Sub cmdAceptar_Click()

Dim c As Long
Dim r As Long

End Sub

Private Sub cmdsalir_Click()
    rs.Close
    Set rs = Nothing
    Unload Me
End Sub
Private Sub Form_Load()
    If rs.State = 0 Then
        rs.Open "select * from usuarios where activo=1 order by descripcion", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    End If
    If Not rs.EOF Then
        grillapermisos.rows = 1
        grillapermisos.cols = rs.Fields.Count - 14
        grillapermisos.Row = 0
        For i = 1 To rs.Fields.Count - 15
                grillapermisos.Col = i
                grillapermisos.ColWidth(i) = 2000
                grillapermisos.Text = rs.Fields(i + 9).Name
        Next i
        grillapermisos.rows = 2
        r = 1
        Do While Not rs.EOF
            
            For c = 0 To grillapermisos.cols - 1
                    grillapermisos.Col = c
                    grillapermisos.Row = r
                    grillapermisos.ColAlignment(0) = flexAlignLeftCenter
                    If c = 0 Then
                         
                        grillapermisos.Text = rs!DESCRIPCION
                    Else
                        grillapermisos.ColDataType(c) = flexDTBoolean
                        If Not IsNull(rs.Fields(c + 9).Value) Then
                            grillapermisos.Text = rs.Fields(c + 9).Value
                        Else
                            grillapermisos.Text = False
                        End If
                    End If
            
            Next c
            rs.MoveNext
            grillapermisos.rows = grillapermisos.rows + 1
            r = r + 1
        Loop
    End If
End Sub
