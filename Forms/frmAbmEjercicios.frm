VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAbmEjercicios 
   Caption         =   "Ejercicios"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "DESTRABAR"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TRABAR"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1545
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   8400
      _extentx        =   14817
      _extenty        =   2725
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VSFlex7LCtl.VSFlexGrid gEjercicios 
      Height          =   2805
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   8235
      _cx             =   14526
      _cy             =   4948
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
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   3720
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   77070337
      CurrentDate     =   38126
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de traba actual:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "-"
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
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha de la traba al iva: "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbmEjercicios.frx":0000
      Height          =   885
      Left            =   2250
      TabIndex        =   3
      Top             =   -30
      Visible         =   0   'False
      Width           =   6075
   End
   Begin VB.Label Label2 
      Caption         =   "Ejercicios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmAbmEjercicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents gej As LiGrilla
Attribute gej.VB_VarHelpID = -1

Private gID As Long
Private gEJER As Long
Private gDESC As Long
Private gFINI As Long
Private gFFIN As Long
Private gACTI As Long
Private gCERR As Long

Private Sub Command1_Click()
    Dim sql As String
    Dim fecha As Date
    
    sql = "update datosempresa set fechatraba=" & ssFecha(dtFecha.Value) & " where idempresa=" & gEMPR_idEmpresa
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se ha actualizado correctamente.", , "ATENCION"
    
    
    fecha = obtenerDeSQL("select fechatraba from datosempresa where idempresa=" & gEMPR_idEmpresa)
    Label4.caption = IIf(fecha = "01/01/1900", "-", fecha)
    If fecha = "01/01/1900" Then
        Command2.enabled = False
    Else
        Command2.enabled = True
    End If
    
End Sub

Private Sub Command2_Click()
    Dim sql As String
    Dim fecha As Date
    
    sql = "update datosempresa set fechatraba='19000101' where idempresa=" & gEMPR_idEmpresa
    DataEnvironment1.Sistema.Execute sql
    
    MsgBox "Se ha destrabado correctamente.", , "ATENCION"
    
    
    fecha = obtenerDeSQL("select fechatraba from datosempresa where idempresa=" & gEMPR_idEmpresa)
    Label4.caption = IIf(fecha = "01/01/1900", "-", fecha)
    If fecha = "01/01/1900" Then
        Command2.enabled = False
    Else
        Command2.enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    Dim fecha As Date
    
    inigrilla
    inimenu
    cargaGrilla
    dtFecha.Value = Date
    fecha = obtenerDeSQL("select fechatraba from datosempresa where idempresa=" & gEMPR_idEmpresa)
    Label4.caption = IIf(fecha = "01/01/1900", "-", fecha)
    If fecha = "01/01/1900" Then
        Command2.enabled = False
    Else
        Command2.enabled = True
    End If
End Sub
Private Sub inigrilla()
    Set gej = New LiGrilla
    
    With gej
        .init gEjercicios, 0
        gID = .AddCol("id", "H")
        gEJER = .AddCol(" Nro  ", "N", 0)
        gDESC = .AddCol(" Denominacion ", "S")
        gFINI = .AddCol(" Inicio    ", "D")
        gFFIN = .AddCol(" Fin       ", "D")
        gACTI = .AddCol(" Activo    ", "K")
        gCERR = .AddCol(" Cerrado    ", "K")
        .rows = 1
    End With
End Sub
Private Sub inimenu()
    uMenu.init False, False, True, False, False
    uMenu.AceptarBorraControles = False
    uMenu.BuscarOK
End Sub
Private Sub cargaGrilla()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCarga
    Dim rs As New ADODB.Recordset, i As Long
    
    With rs
        gej.Borrar
        .Open "select * from ejercicio ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            i = gej.addRow()
            gej.tx i, gID, !idejercicio
            gej.tx i, gEJER, !ejercicio
            gej.tx i, gDESC, !denominacion
            gej.tx i, gFINI, !FechaInicio
            gej.tx i, gFFIN, !FechaFin
            gej.tx i, gACTI, !activo
            gej.tx i, gCERR, !cerrado
            .MoveNext
            
        Wend
        .Close
        gej.addRow
    End With

fin:
    Set rs = Nothing
    Exit Sub
UfaCarga:
    ufa "err cargando ejercicios o parametros cuentas ", ""
End Sub

Private Sub Form_Resize()
    Anclar gEjercicios, Me, anclarLadosTodos
End Sub

Private Sub gej_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim i As Long
    If Col = gACTI Then
        If gej.tk(Row, gACTI) Then
            For i = 1 To gej.rows - 1
                If i <> Row Then
                    If gej.tk(i, gACTI) Then gej.tk i, gACTI, False
                End If
            Next i
        End If
    End If
End Sub

Private Sub uMenu_Aceptar()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    Dim i As Long, su As String, si As String
    Dim ID As Long, ejer As Long, desc As String, fini As Date, ffin As Date, acti As Long, cerr As Long, cuen As String
    
    With gej
        For i = 1 To .rows - 1
            ID = s2n(.tx(i, gID))
            ejer = s2n(.tx(i, gEJER))
            desc = .tx(i, gDESC)
            fini = .td(i, gFINI)
            ffin = .td(i, gFFIN)
            acti = .tk(i, gACTI)
            cerr = .tk(i, gCERR)
            su = "update ejercicio set ejercicio = " & x2s(ejer) & ", denominacion = '" & desc & "',  fechainicio = " & ssFecha(fini) & ", fechafin = " & ssFecha(ffin) & " , activo = " & x2s(acti) & ", cerrado = " & x2s(cerr) & " where idejercicio = " & x2s(ID)
            si = "insert into ejercicio (ejercicio, denominacion, fechainicio, fechafin, activo, cerrado) values ( " & x2s(ejer) & ", '" & desc & "', " & ssFecha(fini) & ", " & ssFecha(ffin) & ", " & x2s(acti) & ", " & x2s(cerr) & " )"
            If ID > 0 Then
                DataEnvironment1.Sistema.Execute su
            ElseIf ejer > 0 Then
                DataEnvironment1.Sistema.Execute si
            End If
        Next i
    End With
    uMenu.AceptarOk
    cargaGrilla
    che "Grabado"
fin:
    Exit Sub
UfaGraba:
    ufa "err al grabar ejercicio", "loop " & i & " de " & gej.rows - 1
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    gEjercicios.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
Private Sub uMenu_SeMovio()
    cargaGrilla
End Sub
