Attribute VB_Name = "modVinculoXl"
' Lito Explicit    ' 10/12/4
Option Explicit


Private oxAp As Object  'que porqueria, perdon.

Private Const FonNor = 10
Private Const FonCpo = 10
Private Const FonTit = 16
'



Public Sub Grilla2Xl(grilla As Object, PathYNombre As String, Optional Titulo As String)
    On Error GoTo UFA_ERR
    Dim nr As Long, nc As Long, r As Long, c As Long, ro As Long, rv As Long
    
    Dim oxSh As Object, ss As String
    
    nr = grilla.rows
    nc = grilla.cols
    Screen.MousePointer = vbHourglass
    
    Set oxSh = CreateObject("Excel.sheet")
    Set oxAp = oxSh.Application
    
    With oxAp
        fon FonTit, True
        If Titulo > "" Then
            .cells(1, 1) = Titulo
            ro = 2
        End If

        rv = 0
        For r = 0 To nr - 1
 '           If r = 0 Then .cells.range(r+1+ro).select fon FonTit, True
            If Not FilaOculta(grilla, r) Then
                rv = rv + 1
                If r = 1 Then fon FonNor, False: ro = ro + 1
                
                For c = 0 To nc - 1
                    ss = grilla.TextMatrix(r, c)
                    .cells(rv + 1 + ro, c + 1) = ss
                Next c
            End If
        Next r
    End With
    
    oxSh.SaveAs PathYNombre
    MsgBox "Informe :" & vbCrLf & "   " & Titulo & vbCrLf & "en" & vbCrLf & "   " & PathYNombre

fin:
    Set oxSh = Nothing
    Set oxAp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
UFA_ERR:
    'ufa
    MsgBox "error"
    Resume fin
End Sub

Private Function FilaOculta(gri As Object, fila) As Boolean
    On Error Resume Next
    FilaOculta = (gri.RowHidden(fila)) '= True)
End Function

Public Sub VinculoXl(ByVal PathYNombre As String, ByVal Titulo As String, Optional rsMaestro As ADODB.Recordset, Optional ByVal Tit0 As String, Optional rsDet0 As ADODB.Recordset, Optional ByVal Tit1 As String, Optional rsDet1 As ADODB.Recordset, Optional ByVal tit2 As String, Optional rsDet2 As ADODB.Recordset, Optional ByVal tit3 As String, Optional rsDet3 As ADODB.Recordset)
  On Error GoTo ufa 'Resume Next
Dim rs As New ADODB.Recordset
  Dim oxSh As Object

  'oXsh.Visible = True
  'oXsh.UserControl = True
  Screen.MousePointer = vbHourglass

  Set oxSh = CreateObject("Excel.sheet")
  Set oxAp = oxSh.Application

  With oxAp
    fon FonTit, True
    .cells(1, 1) = Titulo
    .cells.range("A3").Select
  End With

  PongoRegistro rsMaestro
  PongoDetalle Tit0, rsDet0
  PongoDetalle Tit1, rsDet1
  PongoDetalle tit2, rsDet2
  PongoDetalle tit3, rsDet3

  oxSh.SaveAs PathYNombre
  MsgBox "Informe :" & vbCrLf & "   " & Titulo & vbCrLf & "en" & vbCrLf & "   " & PathYNombre

fin:
  Set oxSh = Nothing
  Set oxAp = Nothing
  Screen.MousePointer = vbDefault
  Exit Sub
ufa:
    che "fallo vinculo excel"
    Resume fin
End Sub

Private Sub PongoRegistro(rr As ADODB.Recordset)
  Dim CC As Long, ii As Long
  
  If rr Is Nothing Then Exit Sub
  
  CC = rr.Fields.Count
  With oxAp
    For ii = 0 To CC - 1
      fon FonCpo, True
      .ActiveCell.FormulaR1C1 = rr.Fields(ii).Name
      .ActiveCell.Offset(0, 1).range("A1").Select
    Next ii
    .ActiveCell.Offset(1, -CC).range("A1").Select
    
    For ii = 0 To CC - 1
      fon FonNor, False
      .ActiveCell.FormulaR1C1 = rr.Fields(ii).Value
      .ActiveCell.Offset(0, 1).range("A1").Select
    Next ii
    .ActiveCell.Offset(2, -CC).range("A1").Select
  End With
  
End Sub

Private Sub PongoDetalle(ByVal tt As String, rr As ADODB.Recordset)
  Dim CC As Long, ii As Long, kk As Long ' kk 4 watch
  
  If rr Is Nothing Then Exit Sub
  'If rr.RecordCount = 0 Then Exit Sub
  If rr.EOF Then Exit Sub
  
  With oxAp
      If tt > "" Then
        .ActiveCell.Offset(1, 0).range("A1").Select
        fon FonCpo, True
        .ActiveCell.FormulaR1C1 = tt
        .ActiveCell.Offset(1, 0).range("A1").Select
      End If
      
      rr.MoveFirst
      CC = rr.Fields.Count
    
    For ii = 0 To CC - 1
      fon FonCpo, True
      .ActiveCell.FormulaR1C1 = rr.Fields(ii).Name
      .ActiveCell.Offset(0, 1).range("A1").Select
    Next ii
    .ActiveCell.Offset(2, -CC).range("A1").Select
      
      While Not rr.EOF
        For ii = 0 To CC - 1
          fon FonNor, False
          .ActiveCell.FormulaR1C1 = rr.Fields(ii).Value
          .ActiveCell.Offset(0, 1).range("A1").Select
        Next ii
        rr.MoveNext
        kk = kk + 1
        .ActiveCell.Offset(1, -CC).range("A1").Select
      Wend
  End With
End Sub

'Public Sub Grilla2Xl(grilla As Object)
'
'End Sub

Private Sub fon(tam, bol)
    With oxAp.Selection.Font
'        .Name = "Arial"
        .Size = tam
        .Bold = bol
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
    End With
End Sub

'17/2/2     hasta 4 rsDetalle, con subtitulos , con hourglass
'2/8/4      Grilla2Excel basico,
'10/12/4    No paso hidden rows
'

