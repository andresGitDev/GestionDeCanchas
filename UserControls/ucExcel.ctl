VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl ucXls 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   FillStyle       =   0  'Solid
   Picture         =   "ucExcel.ctx":0000
   ScaleHeight     =   840
   ScaleWidth      =   780
   Begin VB.CommandButton cmd2Xls 
      Caption         =   "&Excel"
      Height          =   825
      Left            =   -15
      Picture         =   "ucExcel.ctx":5C12
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -15
      Width           =   780
   End
   Begin MSComDlg.CommonDialog dlgXls 
      Left            =   1290
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "xls"
      DialogTitle     =   "Guardar Planilla Como"
      Filter          =   "Archivos Excel (*.xls)|*.xls"
      FilterIndex     =   1
      InitDir         =   "C:\Mis Documentos\"
   End
End
Attribute VB_Name = "ucXls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' metodo ini()
'       obligatorio, usualmente en form_load() : pasar obj grilla
' prop
'   nombrearchivo, titulo dentro del xls, caption,
'   todavia no estan en el bag


'----------------------------------
Private oxAp As Object  'que porqueria, perdon.
Private g As Object     'que porqueria, perdon.
Private mTitulo As String
Private mMsgConfirmacion As String

Private Const FonNor = 10
Private Const FonCpo = 10
Private Const FonTit = 16
'
Public Event Clic(Cancel As Boolean)
'


Public Property Let aMsgConfirmacion(que As String)
    mMsgConfirmacion = que
End Property

Public Property Let aPathYNombre(que As String)
    dlgXls.FileName = que
End Property
Public Property Get aPathYNombre() As String
    aPathYNombre = dlgXls.FileName
End Property

Public Property Let aTitulo(que As String)
    mTitulo = que
End Property
Public Property Let caption(que As String)
    cmd2Xls.caption = que
End Property

Public Property Let enabled(que As Boolean)
    UserControl.enabled = que
End Property
Public Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property

Public Sub ini(grilla As Object, Optional PathYNombre As String, Optional Titulo As String)
    Set g = grilla

    mTitulo = Titulo ' iif(titulo >"" ,Titulo, ????)
    If PathYNombre > "" Then dlgXls.FileName = PathYNombre  'pregunto para no sobreescribir propiedad
End Sub


Private Sub cmd2Xls_Click()
    On Error GoTo ErrHandler
    Dim cancela As Boolean
    
    dlgXls.CancelError = True

    If g.rows = 1 Then Exit Sub
    RaiseEvent Clic(cancela)
    If cancela Then Exit Sub
    
    If mMsgConfirmacion > "" Then
        If Not (MsgBox(mMsgConfirmacion, vbYesNo) = vbYes) Then Exit Sub
    End If
    
    dlgXls.ShowSave
    If dlgXls.FileName > "" Then Grilla2Xl ' g, dlgXls.FileName
    
ErrHandler:
    Exit Sub
End Sub


Private Sub UserControl_Resize()
    cmd2Xls.Height = UserControl.Height
    cmd2Xls.Width = UserControl.Width
End Sub

Private Sub Grilla2Xl() '(Grilla As Object, PathYNombre As String) ', Optional Titulo As String)
    On Error GoTo UFA_ERR
    Dim nr As Long, nc As Long, r As Long, C As Long, ro As Long, rv As Long, cv As Long
    Dim d, m, a
    Dim oxSh As Object, ss As String
    
    nr = g.rows
    nc = g.cols
    Screen.MousePointer = vbHourglass
    
    Set oxSh = CreateObject("Excel.sheet")
    Set oxAp = oxSh.Application
    
    With oxAp
        fon FonTit, True
        If mTitulo > "" Then
            .Cells(1, 1) = mTitulo
            ro = 2
        End If

        rv = 0
        For r = 0 To nr - 1
 '           If r = 0 Then .cells.range(r+1+ro).select fon FonTit, True
            If Not FilaOculta(g, r) Then
                rv = rv + 1
                If r = 1 Then fon FonNor, False: ro = ro + 1
                
                cv = -1
                For C = 0 To nc - 1
                    If Not ColumnaOculta(g, C) Then
                        cv = cv + 1
                        ss = g.TextMatrix(r, C)
                        If IsNumeric(ss) Then
                            .Cells(rv + 1 + ro, cv + 1) = Round(CDbl(ss), 2)
                        ElseIf Len(ss) = 5 Then
                            .Cells(rv + 1 + ro, cv + 1) = Trim(ss)
                        ElseIf IsDate(ss) And Len(ss) = 10 Then
                            d = Day(CDate(ss)) 'Format(Day(CDate(ss)), "00")
                            m = Month(CDate(ss)) 'Format(Month(CDate(ss)), "00")
                            a = Year(CDate(ss)) 'Format(Year(CDate(ss)), "0000")
                            .Cells(rv + 1 + ro, cv + 1) = CDate(d & "/" & m & "/" & a)
                            '.Cells(rv + 1 + ro, cv + 1) = CDate(ss)
                        Else
                            .Cells(rv + 1 + ro, cv + 1) = ss
                        End If
                    End If
                Next C
                
            End If
        Next r
    End With
    
    oxSh.SaveAs dlgXls.FileName 'PathYNombre
    MsgBox "Informe :" & vbCrLf & "   " & mTitulo & vbCrLf & "en" & vbCrLf & "   " & dlgXls.FileName, , "Archivo grabado" 'PathYNombre

fin:
    Set oxSh = Nothing
    Set oxAp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
UFA_ERR:
    'ufa
    MsgBox "Error durante la creacion del archivo.", vbCritical, "Informe"
    Resume fin
End Sub

Private Function FilaOculta(gri As Object, fila) As Boolean
    On Error Resume Next
    FilaOculta = (gri.RowHidden(fila)) '= True)
End Function
Private Function ColumnaOculta(gri As Object, colu) As Boolean
    On Error Resume Next
    ColumnaOculta = (gri.ColHidden(colu)) '= True)
End Function

Private Sub fon(tam, bol)
    With oxAp.selection.Font
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

' 18/5/6  respeta hidden col
