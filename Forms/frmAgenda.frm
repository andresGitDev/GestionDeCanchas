VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAgenda 
   Caption         =   "Agenda"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   Icon            =   "frmAgenda.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMas 
      Height          =   315
      Left            =   11790
      Picture         =   "frmAgenda.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2385
      Width           =   405
   End
   Begin RichTextLib.RichTextBox txtObs 
      Height          =   1305
      Left            =   3795
      TabIndex        =   18
      Top             =   9480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2302
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmAgenda.frx":2284
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   510
      Left            =   10950
      Picture         =   "frmAgenda.frx":2308
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   135
      Width           =   1305
   End
   Begin VB.TextBox txtReferencia 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   9660
      TabIndex        =   16
      Text            =   "0"
      Top             =   90
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   795
      Left            =   10950
      Picture         =   "frmAgenda.frx":334A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1470
      Width           =   1305
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   795
      Left            =   10950
      Picture         =   "frmAgenda.frx":3C14
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   660
      Width           =   1305
   End
   Begin VSFlex7LCtl.VSFlexGrid gAgenda 
      Height          =   7980
      Left            =   105
      TabIndex        =   7
      Top             =   2775
      Width           =   3510
      _cx             =   6191
      _cy             =   14076
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   16711680
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
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
      Cols            =   0
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
   Begin VSFlex7LCtl.VSFlexGrid gTelefonos 
      Height          =   6720
      Left            =   3810
      TabIndex        =   6
      Top             =   2760
      Width           =   8400
      _cx             =   14817
      _cy             =   11853
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   16711680
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
      Cols            =   0
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
   Begin VB.TextBox txtProvincia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   5745
      TabIndex        =   5
      Top             =   1830
      Width           =   5130
   End
   Begin VB.TextBox txtEmpresa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   1830
      Width           =   5130
   End
   Begin VB.TextBox txtLocalidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   5745
      TabIndex        =   3
      Top             =   1140
      Width           =   5130
   End
   Begin VB.TextBox txtDomicilio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   5730
      TabIndex        =   2
      Top             =   435
      Width           =   5130
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   1125
      Width           =   5130
   End
   Begin VB.TextBox txtBusco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   405
      Width           =   5130
   End
   Begin VB.Label lblEstado 
      Height          =   315
      Left            =   5790
      TabIndex        =   19
      Top             =   2355
      Width           =   5085
   End
   Begin VB.Label Label6 
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5745
      TabIndex        =   13
      Top             =   1590
      Width           =   2370
   End
   Begin VB.Label Label5 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   1590
      Width           =   2370
   End
   Begin VB.Label Label4 
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5760
      TabIndex        =   11
      Top             =   870
      Width           =   2370
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   10
      Top             =   870
      Width           =   2370
   End
   Begin VB.Label Label2 
      Caption         =   "Domicilio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5745
      TabIndex        =   9
      Top             =   150
      Width           =   2370
   End
   Begin VB.Label Label1 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   2370
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const tNOMBRE = 0
Private Const tTELEFONO = 1
Private Const tCODIGO = 2
Private Const tSISTEMAS = 3

Public Enum vbpLinea
    lVACIO
    lSISTEMA
    lLLENO
End Enum

Private Sub cmdeliminar_Click()
If MsgBox("¿Seguro?...", vbExclamation + vbYesNo, "Registro") = vbNo Then
    Exit Sub
End If

If abmAgenda("B", s2n(txtReferencia), "", "", "", "", "", "") = True Then
    MsgBox "Eliminado...", vbInformation, "Registro"
    txtBusco_Change
    cmdnuevo_Click
Else
    MsgBox "No Eliminado...", vbExclamation, "Registro"
End If
End Sub

Private Sub cmdguardar_Click()
Dim aOP As String
    If s2n(txtReferencia) = 0 Then
        aOP = "A"
    Else
        aOP = "M"
    End If
    
    If abmAgenda(aOP, s2n(txtReferencia), txtNombre, txtEmpresa, txtDomicilio, txtLocalidad, txtProvincia, txtObs) = True Then
        If aOP = "A" Then
            MsgBox "Guardado...", vbInformation, "Registro"
        Else
            MsgBox "Actualizado...", vbInformation, "Registro"
        End If
    Else
        If aOP = "A" Then
            MsgBox "No Guardado...", vbExclamation, "Registro"
        Else
            MsgBox "No Actualizado...", vbExclamation, "Registro"
        End If
    End If
End Sub

Public Function abmAgenda(aOPERACION As String, aREFERENCIA As Long, aDESCRIPCION As String, aEMPRESA As String, aDOMICILIO As String, aLOCALIDAD As String, aPROVINCIA As String, aOBS As String) As Boolean
On Error GoTo abmerr
Dim aEXECUTE As String, aCOD As Long
'Select Case aOPERACION
'    Case "A":
'        aCOD = nSinNull(obtenerDeSQL("select max(cod) from agenda")) + 1
'        aEXECUTE = "INSERT INTO AGENDA (COD,NOMBRE,EMPRESA,DOMICILIO,LOCALIDAD,PROVINCIA,TELEFONO,OBSER) VALUES (" & aCOD & "," & ssTexto(aNOMBRE) & "," & ssTexto(aEMPRESA) & "," & ssTexto(aDOMICILIO) & "," & ssTexto(aLOCALIDAD) & "," & ssTexto(aPROVINCIA) & ",0," & ssTexto(aOBS) & ")"
'        DataEnvironment1.Sistema.Execute aEXECUTE
'    Case "M":
'        aEXECUTE = "UPDATE AGENDA SET NOMBRE=" & ssTexto(aNOMBRE) & ",EMPRESA=" & ssTexto(aEMPRESA) & ",DOMICILIO=" & ssTexto(aDOMICILIO) & ",LOCALIDAD=" & ssTexto(aLOCALIDAD) & ",PROVINCIA=" & ssTexto(aPROVINCIA) & ",OBSER=" & ssTexto(aOBS) & " WHERE COD=" & aREFERENCIA
'        DataEnvironment1.Sistema.Execute aEXECUTE
'    Case "B":
'        aEXECUTE = "DELETE FROM AGENDA WHERE COD=" & aREFERENCIA
'        DataEnvironment1.Sistema.Execute aEXECUTE
'        aEXECUTE = "DELETE FROM TELEFONOS WHERE NROREF=" & aREFERENCIA
'        DataEnvironment1.Sistema.Execute aEXECUTE
'End Select
Select Case aOPERACION
    Case "A":
        aCOD = nSinNull(obtenerDeSQL("select max(codigo) from clientes")) + 1
        aEXECUTE = "INSERT INTO CLIENTES (CODIGO,DESCRIPCION,NOMBREFANTASIA,DIRECCION,LOCALIDAD,PROVINCIA,TELEFONO,CONTACTO,CATEGORIA,ACTIVO) VALUES (" & aCOD & "," & ssTexto(aDESCRIPCION) & "," & ssTexto(aEMPRESA) & "," & ssTexto(aDOMICILIO) & "," & ssTexto(aLOCALIDAD) & "," & ssTexto(aPROVINCIA) & ",0," & ssTexto(aOBS) & ",1,1)"
        DataEnvironment1.Sistema.Execute aEXECUTE
    Case "M":
        aEXECUTE = "UPDATE CLIENTES SET DESCRIPCION=" & ssTexto(aDESCRIPCION) & ",NOMBREFANTASIA=" & ssTexto(aEMPRESA) & ",DIRECCION=" & ssTexto(aDOMICILIO) & ",LOCALIDAD=" & ssTexto(aLOCALIDAD) & ",PROVINCIA=" & ssTexto(aPROVINCIA) & ",CONTACTO=" & ssTexto(aOBS) & " WHERE CODIGO=" & aREFERENCIA
        DataEnvironment1.Sistema.Execute aEXECUTE
    Case "B":
        'aEXECUTE = "DELETE FROM AGENDA WHERE COD=" & aREFERENCIA
        aEXECUTE = "UPDATE CLIENTES SET ACTIVO=0 WHERE CODIGO=" & aREFERENCIA
        DataEnvironment1.Sistema.Execute aEXECUTE
        aEXECUTE = "DELETE FROM TELEFONOS WHERE NROREF=" & aREFERENCIA
        DataEnvironment1.Sistema.Execute aEXECUTE
End Select


abmAgenda = True


Exit Function
abmerr:
abmAgenda = False
End Function

Private Sub cmdMas_Click()
    gTelefonos.rows = gTelefonos.rows + 1 'insert
    If gTelefonos.Row = -1 Then Exit Sub
End Sub

Private Sub cmdnuevo_Click()
txtReferencia = 0
txtBusco = ""
txtNombre = ""
txtEmpresa = ""
txtDomicilio = ""
txtLocalidad = ""
txtProvincia = ""
lblEstado = ""
gTelefonos.rows = 0
End Sub

Private Sub Form_Load()
Dim sBusco As String
sBusco = "select DESCRIPCION as TELEFONOS, CODIGO from CLIENTES where ACTIVO=1 AND  DESCRIPCION is not null and descripcion>''" _
        & " union " _
        & "select NOMBREFANTASIA as TELEFONOS , CODIGO from CLIENTES where ACTIVO=1 AND  NOMBREFANTASIA is not null and nombrefantasia>''"
LlenarGrilla gAgenda, sBusco, False

If gAgenda.rows > 1 Then
    gAgenda.ColWidth(0) = 3500
    gAgenda.ColWidth(1) = 0
End If
gTelefonos.Editable = flexEDKbdMouse
gTelefonos.cols = 3
gTelefonos.ColWidth(0) = 0
gTelefonos.rows = 0

End Sub

Private Sub gAgenda_Click()
gAgenda_DblClick
End Sub

Private Sub gAgenda_DblClick()
Dim sRef As String, sTel As String
sRef = s2n(gAgenda.TextMatrix(gAgenda.Row, 1))
sTel = "select NOMBRE , TELEFONO , COD from telefonos where nroref=" & sRef
LlenarGrilla gTelefonos, sTel, False
If gTelefonos.rows > 1 Then
    gTelefonos.ColWidth(0) = 2500
    gTelefonos.ColWidth(1) = 2500
    gTelefonos.ColWidth(2) = 0
Else
    gTelefonos.rows = 0
End If

Dim rsAgenda As New ADODB.Recordset, sAgenda As String
sAgenda = "select * from CLIENTES where CODIGO=" & sRef
rsAgenda.Open sAgenda, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsAgenda
    If .EOF And .BOF Then
    Else
        txtReferencia = sRef
        txtNombre = sSinNull(!DESCRIPCION)
        txtEmpresa = sSinNull(!nombrefantasia)
        txtDomicilio = sSinNull(!direccion)
        txtLocalidad = sSinNull(!Localidad)
        txtProvincia = sSinNull(!Provincia)
        txtObs = sSinNull(!contacto)
        lblEstado = "ESTADO : " & sSinNull(obtenerDeSQL("SELECT DESCRIPCION FROM CATEGCLIE WHERE CODIGO=" & s2n(nSinNull(!Categoria))))
    End If
End With
End Sub


Private Sub gAgenda_KeyDown(KeyCode As Integer, Shift As Integer)
gAgenda_DblClick
End Sub

Private Sub gAgenda_KeyPress(KeyAscii As Integer)
gAgenda_DblClick
End Sub

Private Sub gTelefonos_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow = OldRow Then
        If gTelefonos.rows = 1 Then
            abmTelefono OldRow
        Else
            Exit Sub
        End If
    End If
    If gTelefonos.TextMatrix(OldRow, tCODIGO) = "" And LineaTelefono(OldRow) = lLLENO Then
        abmTelefono OldRow
    ElseIf Not IsNumeric(gTelefonos.TextMatrix(OldRow, tCODIGO)) And LineaTelefono(OldRow) = lSISTEMA Then
        gTelefonos.TextMatrix(OldRow, tCODIGO) = Replace(gTelefonos.TextMatrix(OldRow, tCODIGO), "E", "")
        abmTelefono OldRow
    End If
End Sub

Private Function abmTelefono(Roow As Long, Optional rDelete As Boolean = False)
Dim i As Long, tOpe As String, tEXECUTE
If LineaTelefono(Roow) = lVACIO Then Exit Function
If s2n(txtReferencia) = 0 Then Exit Function
With gTelefonos
    If s2n(.TextMatrix(Roow, tCODIGO)) > 0 Then
        tOpe = "M"
    Else
        tOpe = "A"
    End If
    If rDelete Then tOpe = "B"
        
    
    Select Case tOpe
        Case "A":
            tEXECUTE = "INSERT INTO TELEFONOS (NROREF,NOMBRE,TELEFONO) VALUES " _
                    & "(" & s2n(txtReferencia) & "," & ssTexto(.TextMatrix(Roow, tNOMBRE)) & "," & ssTexto(.TextMatrix(Roow, tTELEFONO)) & ")"
            DataEnvironment1.Sistema.Execute tEXECUTE
        Case "M":
            tEXECUTE = "UPDATE TELEFONOS SET NOMBRE=" & ssTexto(.TextMatrix(Roow, tNOMBRE)) & ",TELEFONO=" & ssTexto(.TextMatrix(Roow, tTELEFONO)) & " WHERE COD=" & s2n(.TextMatrix(Roow, tCODIGO))
            DataEnvironment1.Sistema.Execute tEXECUTE
        Case "B":
            tEXECUTE = "DELETE FROM TELEFONOS WHERE COD=" & s2n(.TextMatrix(Roow, tCODIGO))
            DataEnvironment1.Sistema.Execute tEXECUTE
            
    End Select
    
End With
gAgenda_DblClick
End Function

Private Sub gTelefonos_DblClick()
    If gTelefonos.rows = 0 Then gTelefonos.rows = gTelefonos.rows + 1  'insert
End Sub

Private Sub gTelefonos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If gTelefonos.TextMatrix(gTelefonos.Row, tCODIGO) = "" Then
            abmTelefono gTelefonos.Row
        End If
    End If
End Sub

Private Sub gTelefonos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With gTelefonos
    
    If .TextMatrix(Row, tCODIGO) = "" And ((Col = 1 Or Col = 2) Or (Col >= 4)) Then
        'If Col = 3 And .TextMatrix(row, tNOMBRE) > "" Then .TextMatrix(row, cDescripcion) = obtenerDeSQL("select denom from sconcepto where codcon=" & .TextMatrix(row, cConcep))
        'If .TextMatrix(row, cConcep) > "" Then
        '    .TextMatrix(row, cDescripcion) = obtenerDeSQL("select denom from sconcepto where codcon=" & .TextMatrix(row, cConcep))
        'End If
    ElseIf Not IsNumeric(.TextMatrix(Row, tCODIGO)) Then
    Else
        Cancel = True
    End If
End With
End Sub

Private Sub gTelefonos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If gTelefonos.rows = 0 Then Exit Sub
    If s2n(gTelefonos.TextMatrix(gTelefonos.Row, tCODIGO)) > 0 Then
        abmTelefono gTelefonos.Row, True
        If gTelefonos.rows = 0 Then Exit Sub
    Else
        gTelefonos.RemoveItem gTelefonos.Row
        If gTelefonos.rows = 0 Then Exit Sub
    End If
End If
If KeyCode = 45 Then
    gTelefonos.rows = gTelefonos.rows + 1 'insert
    If gTelefonos.Row = -1 Then Exit Sub
End If
'f2
If gTelefonos.Row > (gTelefonos.rows - 1) Then Exit Sub
If KeyCode = 113 And gTelefonos.TextMatrix(gTelefonos.Row, tCODIGO) > "" Then
    gTelefonos.TextMatrix(gTelefonos.Row, tCODIGO) = "E" & gTelefonos.TextMatrix(gTelefonos.Row, tCODIGO)
    'OPEL = "M"
    With gTelefonos
        .cell(flexcpFontBold, gTelefonos.Row, tCODIGO) = True
        .cell(flexcpFontBold, gTelefonos.Row, tNOMBRE) = True
        .cell(flexcpFontBold, gTelefonos.Row, tTELEFONO) = True
    End With
End If
End Sub

Private Function LineaTelefono(r As Long) As vbpLinea
Dim sis As Boolean, Dat As Boolean
With gTelefonos
    If r = 0 Then
        If gTelefonos.rows = 1 Then
        Else
            LineaTelefono = lVACIO
            Exit Function
        End If
    End If

    sis = False
    If .TextMatrix(r, tCODIGO) > "" Then: sis = True
    Dat = True
    'If .TextMatrix(r, tNOMBRE) = "" Then: dat = False
    If .TextMatrix(r, tTELEFONO) = "" Then: Dat = False
End With
If sis And Dat Then
    LineaTelefono = lSISTEMA
ElseIf sis = False And Dat = False Then
    LineaTelefono = lVACIO
ElseIf sis = False And Dat Then
    LineaTelefono = lLLENO
End If
End Function

Private Sub txtBusco_Change()
Dim sBusco As String
'sBusco = "select Nombre AS TELEFONOS, Cod from agenda where Nombre is not null and nombre like '%" & txtBusco & "%'" _
        & " union " _
        & "select Empresa as TELEFONOS , Cod from agenda where empresa is not null and empresa like '%" & txtBusco & "%'"
sBusco = "select DESCRIPCION AS TELEFONOS, CODIGO from CLIENTES where ACTIVO=1 AND descripcion >'' and DESCRIPCION is not null and DESCRIPCION like '%" & txtBusco & "%'" _
        & " union " _
        & "select NOMBREFANTASIA as TELEFONOS , CODIGO from CLIENTES where ACTIVO=1 and nombrefantasia>'' AND NOMBREFANTASIA is not null and NOMBREFANTASIA like '%" & txtBusco & "%'"
        
LlenarGrilla gAgenda, sBusco, False

If gAgenda.rows > 1 Then
    gAgenda.ColWidth(0) = 3500
    gAgenda.ColWidth(1) = 0
End If
        
End Sub

Private Sub txtBusco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then gAgenda.SetFocus
End Sub
