VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmColector 
   Caption         =   "Colector"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   Icon            =   "frmColector.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   855
      Top             =   5805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLeer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Leer excel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2655
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4620
      Width           =   1140
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4620
      Width           =   1365
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   465
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4620
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1545
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4620
      Width           =   975
   End
   Begin VB.CommandButton cmdTransferir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir"
      Height          =   450
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   285
      Width           =   1230
   End
   Begin VB.TextBox txtCarpetaRaizDbf 
      Height          =   345
      Left            =   2835
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   4725
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3315
      Left            =   135
      TabIndex        =   2
      Top             =   1065
      Width           =   8925
      _cx             =   15743
      _cy             =   5847
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmColector.frx":08CA
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
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   660
      Left            =   345
      Top             =   4500
      Width           =   8640
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   855
      Left            =   360
      Top             =   105
      Width           =   8595
   End
End
Attribute VB_Name = "frmColector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim Fs As New FileSystemObject
Private g As LiGrilla

Private gFECHA As Long
Private gFAC_INI As Long
Private gFAC_FINAL As Long
Private gCLIENTE As Long

Private sTmpFacturas As String
Private sTmpClientes As String
Private cDireccion As String
Private cn As New ADODB.Connection


Private Sub cmdcancelar_Click()
    g.Borrar
    cmdGuardar.enabled = False
    cmdLeer.enabled = False
    txtCarpetaRaizDbf.Text = ""
End Sub

Private Sub cmdguardar_Click()
    Dim sql As String
    Dim ide As Long
    Dim util As Long
    Dim i As Long
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    i = 1
    If Not grilla.rows > 1 Then
        MsgBox "Debe buscar los datos de un excel para guardar."
        Exit Sub
    End If
    
    
    While grilla.rows > i
        With grilla
            
            rs.Open "Select max(id) as cod from codigobarras", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rs.BOF = True And rs.EOF = True Then
                ide = 1
            Else
                If IsNull(rs!cod) Then
                    ide = 1
                Else
                    ide = rs!cod + 1
                End If
            End If
            Set rs = Nothing
            
            rs2.Open "Select * from producto where codigobarra='" & .TextMatrix(i, 0) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rs2.EOF = True And rs2.BOF = True Then
                MsgBox "El codigo de barra " & .TextMatrix(i, 0) & " no tiene una asignacion a un producto."
                datobarra = .TextMatrix(i, 0)
                frmCodbarra.Show
                Exit Sub
            Else
            End If
            
            rs.Open "Select * from codigobarras where nroproducto='" & .TextMatrix(i, 0) & "' and nroserie='" & .TextMatrix(i, 1) & "' and activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If rs.EOF = True And rs.BOF = True Then
                sql = "Insert Into codigobarras ( nroproducto, nroserie, utilizado,activo) " & _
                    "Values ('" & .TextMatrix(i, 0) & "','" & .TextMatrix(i, 1) & "'," & util & ",1)"
                DataEnvironment1.Sistema.Execute sql
                MsgBox "Se ha realizado con exito la carga."
            Else
                MsgBox "El nro de serie " & .TextMatrix(i, 1) & " del producto " & rs2!DESCRIPCION & " ya existe y no se podra agregar ni modificar." & Chr(13) & Chr(13) & Chr(13) & "Presione aceptar para continuar grabando..."
            End If
            
            Set rs = Nothing
            Set rs2 = Nothing
        End With
        i = i + 1
    Wend
        
    g.Borrar
    txtCarpetaRaizDbf.Text = ""
    cmdLeer.enabled = False
    cmdGuardar.enabled = False
End Sub

Private Sub cmdLeer_Click()
'Private Sub Form_Load()
    Inicio
    Caracteristicas
    FormatearTabla
    LlenadoDeTabla
    cmdGuardar.enabled = True
'End Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Finalizar
End Sub


'Private Sub mnuSalir_Click()
'ApliNudos.Quit
'End
'End Sub



Private Sub cmdTransferir_Click()
    ' Establecer CancelError a True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler

    CommonDialog1.Filter = "Archivos de planilla de calculo(*.xls)|*.xls|Todos los archivos (*.*)|*.*|"
    CommonDialog1.ShowOpen
    txtCarpetaRaizDbf = CommonDialog1.FileName
    
        'MsgBox "Debe seleccionar solo archivos de EXCEL"
        'CommonDialog1.ShowOpen
        'txtCarpetaRaizDbf.Text = ""
    cmdLeer.enabled = True
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    'txtCarpetaRaizDbf = ""
    Exit Sub
End Sub
Private Sub cmdsalir_Click()
    Set ApliNudos = CreateObject("Excel.Application")
    ApliNudos.Quit
    Unload Me
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    Set g = New LiGrilla
    g.init grilla
    
    gFECHA = g.AddCol("     Nro de Producto     ")
    gFAC_INI = g.AddCol("       Nro de Serie       ")
    gFAC_FINAL = g.AddCol("borrar", "H")
    gCLIENTE = g.AddCol("       borrar         ", "H")

    txtCarpetaRaizDbf.Text = "C:\"
End Sub





