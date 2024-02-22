VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmActClieProv 
   Caption         =   "Actualizacion de Clientes y Proveedores"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   Icon            =   "frmActClieProv.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10770
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin Gestion.ucXls ucXls1 
      Height          =   900
      Left            =   5880
      TabIndex        =   8
      Top             =   600
      Width           =   1530
      _extentx        =   2699
      _extenty        =   1588
   End
   Begin VB.CommandButton cmdAgrilla 
      Caption         =   "4-Actualizar de grilla"
      Height          =   900
      Left            =   4320
      Picture         =   "frmActClieProv.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   585
      Width           =   1530
   End
   Begin VSFlex7LCtl.VSFlexGrid gDatos 
      Height          =   8115
      Left            =   9090
      TabIndex        =   5
      Top             =   150
      Width           =   5565
      _cx             =   9816
      _cy             =   14314
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmActClieProv.frx":1194
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
   Begin VSFlex7LCtl.VSFlexGrid gResultado 
      Height          =   8085
      Left            =   105
      TabIndex        =   4
      Top             =   2010
      Width           =   8880
      _cx             =   15663
      _cy             =   14261
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
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "3-Actualizar"
      Height          =   915
      Left            =   2910
      Picture         =   "frmActClieProv.frx":122F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   585
      Width           =   1365
   End
   Begin VB.CommandButton cmdSubir 
      Caption         =   "2-Subir a la base"
      Height          =   930
      Left            =   1515
      Picture         =   "frmActClieProv.frx":1AF9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   570
      Width           =   1350
   End
   Begin VB.CommandButton cmdUbicacion 
      Caption         =   "1-Ubicacion"
      Height          =   915
      Left            =   105
      Picture         =   "frmActClieProv.frx":23C3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   585
      Width           =   1365
   End
   Begin VB.TextBox txtUbicacion 
      Height          =   345
      Left            =   105
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   105
      Width           =   7020
   End
   Begin VB.Label lblULTPERIODO 
      Height          =   360
      Left            =   135
      TabIndex        =   9
      Top             =   10200
      Width           =   5610
   End
   Begin VB.Label Label1 
      Caption         =   "Clientes y Proveedores que no se actualizaron (si desea ver los ultimos que no se actualizaron cliclear en el boton 4)"
      Height          =   360
      Left            =   90
      TabIndex        =   6
      Top             =   1620
      Width           =   8700
   End
End
Attribute VB_Name = "frmActClieProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const tt = "ActCliProv"
Private Const T1 = "Cliente"
Private Const t2 = "Proveedor"
Private Const COD = 1
Private Const cc = 4
Private Const CP = 5

Private Sub cmdActualizar_Click()
On Error GoTo mal1
Dim rsdat As New ADODB.Recordset, i As Long, cad As String
Dim tmp01, e
gIni
add_sql "B", "", "", "", "", "", ""
'actualizacion de clientes
With rsdat
    .Open "Select codigo,cuit,descripcion from clientes where puedofacturar=1 and activo=1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tmp01 = obtenerDeSQL("select coefcli_perc from actcliprov where cuit=" & ssTexto(!CUIT))
            If IsNull(tmp01) Or IsEmpty(tmp01) Then
                addgri T1, !codigo, !CUIT, !DESCRIPCION, 0, "-"
                cad = " delete from ClieTipoRetIB_Per where codclie= " & !codigo
                DataEnvironment1.Sistema.Execute cad
                cad = "update clientes set ConPercIIBB=0,conperciibbper=0,conpercganper=0 where codigo=" & !codigo
                DataEnvironment1.Sistema.Execute cad
            Else
                e = obtenerDeSQL(" select idiibbper from ClieTipoRetIB_Per where codclie= " & !codigo)
                If IsNull(e) Or IsEmpty(e) Then
                    cad = "insert into ClieTipoRetIB_Per (Codigo,CodClie,BaseImponible,Coeficiente) " _
                    & " values (1," & !codigo & ",50," & x2s(tmp01 / 100) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    cad = "update clientes set ConPercIIBB=1,conperciibbper=1 where codigo=" & !codigo
                    DataEnvironment1.Sistema.Execute cad
                Else
                    cad = "update ClieTipoRetIB_Per set coeficiente= " & x2s(tmp01 / 100) & " where idiibbper=" & e
                    DataEnvironment1.Sistema.Execute cad
                    cad = "update clientes set ConPercIIBB=1,conperciibbper=1 where codigo=" & !codigo
                    DataEnvironment1.Sistema.Execute cad
                End If
            End If
            .MoveNext
        Next
        
    End If
End With
Set rsdat = Nothing

'actualizacion de proveedores
With rsdat
    .Open "Select codigo,cuit,descripcion from prov where Activo_PR=1 and activo=1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tmp01 = obtenerDeSQL("select coefprov_ret from actcliprov where cuit=" & ssTexto(!CUIT))
            
            Select Case !codigo
                'Case 426, 564, 432, 387, 495, 607, 425, 431, 562, 424: Stop
            End Select
            
            If IsNull(tmp01) Or IsEmpty(tmp01) Then
                addgri t2, !codigo, !CUIT, !DESCRIPCION, "-", 0
                cad = " delete from ProvTipoRetIB_Per where codprov= " & !codigo
                DataEnvironment1.Sistema.Execute cad
                cad = "update prov set RetenerIIBB=0,conretiibbper=0,conretganper=0 where codigo=" & !codigo
                DataEnvironment1.Sistema.Execute cad
            Else
                e = obtenerDeSQL(" select idiibbper from ProvTipoRetIB_Per where codprov= " & !codigo)
                If IsNull(e) Or IsEmpty(e) Then
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (0," & !codigo & ",0,0)"
                    DataEnvironment1.Sistema.Execute cad
                
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (1," & !codigo & ",400," & x2s(tmp01) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (2," & !codigo & ",400," & x2s(tmp01) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (3," & !codigo & ",400," & x2s(tmp01) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    
                    cad = "update prov set RetenerIIBB=1,conretiibbper=1 where codigo=" & !codigo
                    DataEnvironment1.Sistema.Execute cad
                Else
                    cad = "update ProvTipoRetIB_Per set coeficiente= " & x2s(tmp01) & " where codigo<>0 and codprov=" & !codigo
                    DataEnvironment1.Sistema.Execute cad
                    cad = "update prov set RetenerIIBB=1,conretiibbper=1 where codigo=" & !codigo
                    DataEnvironment1.Sistema.Execute cad
                End If
                
            End If
            .MoveNext
        Next
        
    End If
End With
Set rsdat = Nothing

Exit Sub
mal1:
    MsgBox "Error", vbExclamation
End Sub

Private Function addgri(tip, COD, cu, ra, c01, c02)
With gResultado
    .AddItem tip & Chr(9) & COD & Chr(9) & cu & Chr(9) & ra & Chr(9) & c01 & Chr(9) & c02
    add_sql "A", CStr(tip), CStr(COD), CStr(cu), CStr(ra), CStr(c01), CStr(c02)
End With
End Function

Private Function add_sql(Ope As String, dat01 As String, dat02 As String, dat03 As String, dat04 As String, dat05 As String, dat06 As String)
Dim tabla As String, xxp As String
tabla = "NoPadron_XLS"
Ope = UCase(Ope)

Select Case Ope
    Case "B":
        xxp = "DELETE FROM " & tabla
        DataEnvironment1.Sistema.Execute xxp
        xxp = "DBCC CHECKIDENT (" & tabla & ", RESEED, 0)"
        DataEnvironment1.Sistema.Execute xxp
    Case "A":
        xxp = "INSERT INTO " & tabla & " (DAT01,DAT02,DAT03,DAT04,DAT05,DAT06) " _
        & " VALUES (" & ssTexto(dat01) & "," & ssTexto(dat02) & "," & ssTexto(dat03) & "," & ssTexto(ssStr(dat04)) & "," & ssTexto(dat05) & "," & ssTexto(dat06) & ")"
        DataEnvironment1.Sistema.Execute xxp
End Select

End Function

Private Sub cmdAgrilla_Click()
Dim i As Long
Dim cad As String
Dim e
Dim cont As Long
cont = 0
With gResultado
    For i = 1 To .rows - 1
        If .TextMatrix(i, 0) = T1 Then
            If s2n(.TextMatrix(i, cc)) > 0 Then
                e = obtenerDeSQL(" select idiibbper from ClieTipoRetIB_Per where codclie= " & .TextMatrix(i, COD))
                If IsNull(e) Or IsEmpty(e) Then
                    cont = cont + 1
                    cad = "insert into ClieTipoRetIB_Per (Codigo,CodClie,BaseImponible,Coeficiente) " _
                    & " values (1," & .TextMatrix(i, COD) & ",50," & x2s(.TextMatrix(i, cc)) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    cad = "update clientes set conperciibbper=1 where codigo=" & .TextMatrix(i, COD)
                    DataEnvironment1.Sistema.Execute cad
                Else
                    cont = cont + 1
                    cad = "update ClieTipoRetIB_Per set coeficiente= " & x2s(.TextMatrix(i, cc)) & " where idiibbper=" & e
                    DataEnvironment1.Sistema.Execute cad
                End If
            End If
        ElseIf .TextMatrix(i, 0) = t2 Then
            If s2n(.TextMatrix(i, CP)) > 0 Then
                e = obtenerDeSQL(" select idiibbper from ProvTipoRetIB_Per where codprov= " & .TextMatrix(i, COD))
                If IsNull(e) Or IsEmpty(e) Then
                    cont = cont + 1
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (0," & .TextMatrix(i, COD) & ",0,0)"
                    DataEnvironment1.Sistema.Execute cad
                
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (1," & .TextMatrix(i, COD) & ",400," & x2s(.TextMatrix(i, CP)) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (2," & .TextMatrix(i, COD) & ",400," & x2s(.TextMatrix(i, CP)) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    cad = "insert into ProvTipoRetIB_Per (Codigo,CodProv,BaseImponible,Coeficiente) " _
                    & " values (3," & .TextMatrix(i, COD) & ",400," & x2s(.TextMatrix(i, CP)) & ")"
                    DataEnvironment1.Sistema.Execute cad
                    
                    
                    cad = "update prov set conretiibbper=1 where codigo=" & .TextMatrix(i, COD)
                    DataEnvironment1.Sistema.Execute cad
                Else
                    cont = cont + 1
                    cad = "update ProvTipoRetIB_Per set coeficiente= " & x2s(.TextMatrix(i, CP)) & " where codigo<>0 and idiibbper=" & e
                    DataEnvironment1.Sistema.Execute cad
                End If
            End If
        End If
    Next
    MsgBox cont & " Registros Actualizados"
End With
If gResultado.rows > 1 Then
Else
    LlenarGrilla gResultado, "select dat01 as Tipo, dat02 as Codigo,dat03 as Cuit,dat04 as [Razon Social],dat05 as [Coef Cli],dat06 as [Coef Prov] from nopadron_xls", False
End If
End Sub

Private Sub cmdSubir_Click()
Dim tmp03, tmp02
If txtUbicacion = "" Then
    MsgBox "No se indico el archivo.", vbCritical
    Exit Sub
End If

If ExisteArchivo(txtUbicacion) Then
Else
    MsgBox "El archivo indicado no existe.", vbExclamation
    Exit Sub
End If

If Leer_Txt(gDatos, "", True, , , , txtUbicacion) = False Then
    MsgBox "Datos no se subieron a la base", vbCritical
Else
    MsgBox "Datos subidos a la base", vbInformation
End If

End Sub

Private Sub cmdUbicacion_Click()
txtUbicacion = VentanaArchivo(Me, "*.txt", "Archivo TXT de Actualizacion")
End Sub

Public Function Leer_Txt(grilla As Control, ConsultaSQL As String, AjustarAnchos As Boolean, Optional nColCorte, Optional nColSum, Optional llenacomo As LlenarGrillaComo, Optional Ubi As String) As Boolean
On Error GoTo M2
    ' agregado corte, no implementada la suma aun
    'agregado col invisible, alias empieza con "_H_"     ejnombre = "_H_idRegistro"
    Dim rsaux As New ADODB.Recordset
    Dim C As Long
    Dim Encabezado As String
    Dim ConCorte As Boolean, ColCorte As Long, ColTMP()  ' todo para corte
    Dim aPos As Double
'*****************************************
    Dim tmp01, tmp04, tmp05
    Dim txtNombre As String, txtArchivo As String
    Dim ubi2 As String, sql As String
    tmp01 = Split(Ubi, "\")
    tmp04 = Split(tmp01(UBound(tmp01)), ".")
    txtNombre = "TMPTXTDAT" 'tmp04(0)
    txtArchivo = "\" & txtNombre & ".txt"
    
    ubi2 = tmp01(0)
    For C = 1 To UBound(tmp01) - 1
        ubi2 = ubi2 & "\" & tmp01(C)
    Next
    If C = 1 Then ubi2 = ubi2 & "\"
    
    If ConsultaSQL = "" Then
        ConsultaSQL = "'select * from [" & txtNombre & "#txt]'"
    End If
    
    Leer_Txt = True
'Importar mediante la funciÃ³n OPENROWSET
'CREA AUTOMATICAMENTE LOS NOMBRES DE LAS COLUMNAS
ubi2 = App.Path & "\"
 'sql = "SELECT *  INTO ActCliProv  " & _
       "FROM OPENROWSET(" & _
       "'Microsoft.Jet.OLEDB.4.0'," & _
       "'TEXT;Database=" & ubi2 & ";HDR=No;FMT=Delimited(;)'," & _
       ConsultaSQL & ")"
 
'Importar mediante la funciÃ³n OPENDATASOURCE
'EL NOMBRE DE CADA COLUMNA DEBE ESTAR EN EL PRIMER RENGLON
'sql = "SELECT * INTO ActCliProv " & _
      "FROM OPENDATASOURCE(" & _
      "'Microsoft.Jet.OLEDB.4.0'," & _
      "'Data Source=" & ubi2 & ";" & _
        "Extended Properties=""TEXT;HDR=No""')" & _
      "..." & txtnombre & "#txt"
    'CAMPO1;CAMPO2;CAMPO3;CUIT;CAMPO5;CAMPO6;CAMPO7;COEFCLI_PERC;COEFPROV_RET;CAMPO10;CAMPO11
'BLUK INSERT pubs.dbo.plano
'FROM ‘c:\plano.txt’
'WITH
'(DATAFILETYPE = ‘char’, FIELDTERMINATOR = ‘,’ , ROWTERMINATOR = ‘;’)
    
  'sql = "BLUK INSERT tonka.dbo.actcliprov " & _
        "FROM ‘" & ubi2 & txtnombre & ".txt" & "’ " & _
        "WITH (DATAFILETYPE = ‘varchar’, FIELDTERMINATOR = ‘;’, ROWTERMINATOR = ‘\n’)"
    
    If ExisteArchivo(App.Path & txtArchivo) Then Kill App.Path & txtArchivo
    If ExisteArchivo(Ubi) Then
        FileCopy Ubi, App.Path & txtArchivo
    Else
        MsgBox "No se encontro el archivo de Origen.", vbCritical
        Exit Function
    End If
    
    'Dim connTXT As New ADODB.Connection
    'Cadena de conexión
    'en crudo como vienen los datos
    'connTXT.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                  "DBQ=" & App.Path & ";", "", ""
    
    'separados por coma
    'connTXT.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
                     & "Data Source=" & App.Path & ";" _
                    & "Extended Properties='text;FMT=Delimited'"
'*******************************************
    DataEnvironment1.Sistema.Execute "drop table actcliprov"
    
    Dim MMM As String
    Dim baseVB As String
    baseVB = DataEnvironment1.Sistema.Properties("Initial Catalog")
    MMM = DataEnvironment1.Sistema.Properties("Data Source")
    
    If ACTCLIPROV(MMM, baseVB, "ActCliProv", ubi2 & txtNombre & ".txt", txtNombre) Then
    Else
        GoTo M2
    End If
        
    DataEnvironment1.Sistema.Execute "ALTER TABLE ACTCLIPROV ADD CUIT VARCHAR(500) "
    DataEnvironment1.Sistema.Execute "ALTER TABLE ACTCLIPROV ADD COEFCLI_PERC VARCHAR(500) "
    DataEnvironment1.Sistema.Execute "ALTER TABLE ACTCLIPROV ADD COEFPROV_RET VARCHAR(500) "
    DataEnvironment1.Sistema.Execute "UPDATE    ActCliProv SET COEFPROV_RET = F9,COEFCLI_PERC = F8"
    DataEnvironment1.Sistema.Execute "UPDATE    ActCliProv  set CUIT=cast(F4 as bigint)"
    DataEnvironment1.Sistema.Execute "UPDATE    ActCliProv SET CUIT = left(cuit,2) +'-' + LEFT(RIGHT(cuit,9),8) +  '-' + right(cuit,1)"
    
Exit Function
M2:
Leer_Txt = False

End Function

Public Function Leer_Txt2(grilla As Control, ConsultaSQL As String, AjustarAnchos As Boolean, Optional nColCorte, Optional nColSum, Optional llenacomo As LlenarGrillaComo, Optional Ubi As String) As Boolean
    ' agregado corte, no implementada la suma aun
    'agregado col invisible, alias empieza con "_H_"     ejnombre = "_H_idRegistro"
    Dim rsaux As New ADODB.Recordset
    Dim C As Long
    Dim Encabezado As String
    Dim ConCorte As Boolean, ColCorte As Long, ColTMP()  ' todo para corte
    Dim aPos As Double
'*****************************************
    Dim tmp01, tmp04, tmp05
    Dim txtNombre As String, txtArchivo
    tmp01 = Split(Ubi, "\")
    tmp04 = Split(tmp01(UBound(tmp01)), ".")
    txtNombre = "TMPTXTDAT" 'tmp04(0)
    txtArchivo = "\" & txtNombre & ".txt"
    
    If ConsultaSQL = "" Then
        ConsultaSQL = "select * from [" & txtNombre & "#txt]"
    End If
    
    If ExisteArchivo(App.Path & txtArchivo) Then Kill App.Path & txtArchivo
    If ExisteArchivo(Ubi) Then
        FileCopy Ubi, App.Path & txtArchivo
    Else
        MsgBox "No se encontro el archivo de Origen.", vbCritical
        Exit Function
    End If

    Dim connTXT As New ADODB.Connection
    'Cadena de conexión
    'en crudo como vienen los datos
    connTXT.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                  "DBQ=" & App.Path & ";", "", ""
    
    'separados por coma
    'connTXT.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
                     & "Data Source=" & App.Path & ";" _
                    & "Extended Properties='text;FMT=Delimited'"
'*******************************************

    'ColCorte = s2n(nColCorte)
    ConCorte = Not IsMissing(nColCorte) And ColCorte >= 0
    
    If ConsultaSQL <> "" Then
        If llenacomo = llenagResetear Then grilla.clear: grilla.rows = 1
        
        rsaux.Open ConsultaSQL, connTXT, adOpenForwardOnly, adLockReadOnly
        If Not rsaux.EOF Then
        grilla.clear
            With rsaux
                If ColCorte > .Fields.Count Then ConCorte = False
                
                If llenacomo = llenagResetear Then
                    'hago el encabezado de la grilla
                    grilla.FixedCols = 0
                    grilla.cols = 1 '.Fields.Count
                    grilla.Row = 0
                    grilla.TextMatrix(0, 0) = "Datos"
                    grilla.ColWidth(0) = 6000
                
                    grilla.rows = 1
              
                End If
                Dim tmpCuit As String
                'lleno la grilla con los datos de la consulta
                ABMTMP "B", "", "", "", "", "", "", "", 0, 0, "", ""
                While Not .EOF
                    grilla.rows = grilla.rows + 1
                    grilla.TextMatrix(grilla.rows - 1, 0) = .Fields(0) & "," & .Fields(1) & "," & .Fields(2) 'IIf(IsNull(.Fields(c)), "", .Fields(c))
                    tmp05 = Split(grilla.TextMatrix(grilla.rows - 1, 0), ";")
                    tmpCuit = CORTO(CStr(tmp05(3)), 0, Len(CStr(tmp05(3))) - 2) & "-" & CORTO(CStr(tmp05(3)), 2, 1) & "-" & CORTO(CStr(tmp05(3)), Len(CStr(tmp05(3))) - 1, 0)
                    ABMTMP "A", CStr(tmp05(0)), CStr(tmp05(1)), CStr(tmp05(2)), tmpCuit, CStr(tmp05(4)), CStr(tmp05(5)), CStr(tmp05(6)), s2n(tmp05(7)), s2n(tmp05(8)), CStr(tmp05(9)), CStr(tmp05(10))
                    .MoveNext
                Wend
                
            End With
            Leer_Txt2 = True
        Else
            Leer_Txt2 = False
        End If
    End If
    
    Set rsaux = Nothing
End Function

Private Function ABMTMP(tOpe As String, C1 As String, C2 As String, C3 As String, tCUIT As String, C5 As String, C6 As String, C7 As String, tCOEFCLI As Double, tCOEFPRO As Double, C10 As String, C11 As String) As Boolean
Dim cad As String
ABMTMP = True
cad = ""
    Select Case tOpe
        Case "A":
            cad = "INSERT INTO ACTCLIPROV (CAMPO1,CAMPO2,CAMPO3,CUIT,CAMPO5,CAMPO6,CAMPO7,COEFCLI_PERC,COEFPROV_RET,CAMPO10,CAMPO11) " _
            & " VALUES (" & ssTexto(C1) & "," & ssTexto(C2) & "," & ssTexto(C3) & "," & ssTexto(tCUIT) & "," & ssTexto(C5) & "," & ssTexto(C6) & "," & ssTexto(C7) & "," & x2s(tCOEFCLI) & "," & x2s(tCOEFPRO) & "," & ssTexto(C10) & "," & ssTexto(C11) & ")"
            DataEnvironment1.Sistema.Execute cad
        Case "B":
            cad = "DELETE FROM ACTCLIPROV"
            DataEnvironment1.Sistema.Execute cad
            DataEnvironment1.Sistema.Execute "DBCC CHECKIDENT (ACTCLIPROV, RESEED, 0)"
    End Select
Exit Function
MAL:
ABMTMP = False
End Function

Public Function ssTexto(d As String) As String
ssTexto = "'" & Trim(d) & "'"
End Function

Private Function gIni()
With gResultado
    .rows = 1
    .cols = 6
    .TextMatrix(0, 0) = "Tipo"
    .TextMatrix(0, COD) = "Codigo"
    .TextMatrix(0, 2) = "Cuit"
    .TextMatrix(0, 3) = "Razon Social"
    .TextMatrix(0, cc) = "Coef Cli"
    .TextMatrix(0, CP) = "Coef Prov"
    .ColWidth(0) = 900
    .ColWidth(COD) = 800
    .ColWidth(2) = 1500
    .ColWidth(3) = 2800
    .ColWidth(cc) = 800
    .ColWidth(CP) = 800
    .Editable = flexEDKbdMouse
End With

End Function

Private Sub Form_Load()
gDatos.rows = 1
gIni
ucXls1.ini gResultado, "C:\PorActualizar.xls"
lblULTPERIODO = "ULTIMO PERIODO CARGADO : " & Format(obtenerDeSQL("SELECT F2 FROM ACTCLIPROV GROUP BY F2"), "##/##/####")
End Sub

Private Sub gResultado_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With gResultado
    If (Col = cc) Then
        If .TextMatrix(Row, 0) = T1 Then
        Else
            Cancel = True
        End If
    ElseIf (Col = CP) Then
        If .TextMatrix(Row, 0) = t2 Then
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End With
End Sub
