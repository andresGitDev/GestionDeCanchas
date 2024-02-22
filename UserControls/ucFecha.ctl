VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ucFecha 
   BackColor       =   &H80000005&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   1275
   Begin MSMask.MaskEdBox txtFecha 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   476
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "ucFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '   17/4/6
'Lito 17/11/4

'un dtPicker, pero teclas mas inteligentes, sin el calendario
'
Private Const AnioLimite = 1930  ' debe haber alguna dll por ahi

Public Enum ucFechaIni
    ucPrimerDiaAnio
    ucPrimerDiaMes
    ucUltimoDiaAnio
    ucUltimoDiaMes
    ucHoy
    ucUnAnioAtras
End Enum

Private mUltFechaOk As Date
Private mFechaIni As ucFechaIni

'Public Property Get strFecha() As String
'    strFecha = ""
'    If Verificar() Then strFecha = txtFecha
'End Property
'Public Property Let strFecha(cual As String)
'    If Verificar(cual) Then txtFecha = cual
'End Property

Public Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property

Public Property Let enabled(como As Boolean)
    UserControl.enabled = como
    PropertyChanged "Enabled"
End Property

Public Property Get Dia() As Long
    Dia = Day(dtFecha())
End Property
'Public Property Let Dia(cual As Integer)
'
'End Property

Public Property Get Mes() As Long
    Mes = Month(dtFecha())
End Property
'Public Property Let Mes(cual As Integer)
'    if cual >= 1 and cual <= 12 then
'End Property

Public Property Get Anio() As Long
    Anio = Year(dtFecha())
End Property
'Public Property Let Anio(cual As Integer)
'
'End Property



Public Function strFecha(Optional cual) As String
    If IsMissing(cual) Then
        strFecha = ""
        If Verificar() Then strFecha = TxtFecha
    Else
        If Verificar(cual) Then TxtFecha = cual
    End If
End Function
Public Function contenido(Valor As Variant) As Date
    'TxtFecha = dtfecha(valor)
    TxtFecha = Format(Valor, "dd/mm/yy")
End Function
Public Function dtFecha(Optional cual) As Date
    'verificar
    On Error Resume Next
    Dim s As String
    Dim fff As String
   
    If IsMissing(cual) Then
        dtFecha = CDate(TxtFecha)
    Else
'        fff = FormatoFecha()
'        If Left(LCase(fff), 1) = "m" Then
'            s = Format(cual, "mm/dd/yy")
'        ElseIf Left(LCase(fff), 1) = "d" Then
'            s = Format(cual, "dd/mm/yy")
'        Else
'            ufa "", "formato fecha no reconocido " & fff
'        End If
        TxtFecha = ff(CStr(cual))
        mUltFechaOk = cual
    End If
End Function

Public Function NumeralFecha()
    NumeralFecha = "#" & Format(dtFecha(), "mm/dd/yy") & "#"
End Function

Public Function ConvertFecha()
    ConvertFecha = " convert(datetime, '" & Format(dtFecha(), "mm/dd/yy") & "' ,1 ) "
End Function

Public Function ssFecha() As String
    ssFecha = " '" & Format(dtFecha(), "yyyymmdd") & "' "
End Function

Public Sub SetPrimerDiaMes(Optional Mes As Long, Optional Anio As Long)
    If Mes = 0 Then Mes = Month(Date)
    If Anio = 0 Then Anio = Year(Date)
    If UCase(Left(FormatoFecha(), 1)) = "M" Then
        TxtFecha = Format(Mes, "00") & "/01/" & Right(CStr(Anio), 2)
    Else
        TxtFecha = "01/" & Format(Mes, "00") & "/" & Right(CStr(Anio), 2)
    End If
    mUltFechaOk = Me.dtFecha
End Sub
Public Sub setUltDiaMes(Optional Mes As Long, Optional Anio As Long)
    If Mes = 0 Then Mes = Month(Date)
    If Anio = 0 Then Anio = Year(Date)
    TxtFecha = ff(DateSerial(Anio, Mes + 1, 0))
    mUltFechaOk = Me.dtFecha
End Sub


'Public Property Get dtFecha() As Date
'    'verificar
'    On Error Resume Next
'    dtFecha = CDate(txtFecha)
'End Property
'Public Property Let dtFecha(cual As Date)
'    txtFecha = Format(cual, "dd/mm/yy")
'    PropertyChanged "dtFecha"
'End Property

Public Property Let FechaInit(cual As ucFechaIni)
    mFechaIni = cual
    PropertyChanged "FechaInit"
End Property
Public Property Get FechaInit() As ucFechaIni
    FechaInit = mFechaIni
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    mFechaIni = PropBag.ReadProperty("FechaInit")
    ini mFechaIni
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "FechaInit", mFechaIni
End Sub


'********* deprecated ***********
Public Sub ini(Dia As ucFechaIni) '---*/*/*/--- deprecated '---*/*/*/---
    Select Case Dia
     Case ucPrimerDiaAnio:   TxtFecha = ff("01/01/" & Year(Date))
     Case ucPrimerDiaMes:   TxtFecha = ff("01/" & Month(Date) & "/" & Year(Date))
     Case ucUltimoDiaAnio:   TxtFecha = ff("31/12/" & Year(Date))
     Case ucUltimoDiaMes:    TxtFecha = ff(DateSerial(Year(Date), Month(Date) + 1, 0))  'ff("__/" & Month(Date) & "/" & Year(Date))
     Case ucHoy:             TxtFecha = ff(Date)
     Case ucUnAnioAtras:   TxtFecha = ff(Date - 365)
    End Select
    
    mUltFechaOk = Me.dtFecha
End Sub
'-----------------------------------

Private Sub txtFecha_LostFocus()
  
    If Not Verificar() Then
        TxtFecha.SetFocus
    Else
        
        mUltFechaOk = Me.dtFecha
    End If
End Sub


Private Sub txtFecha_GotFocus()
    GotFocusPinto TxtFecha
End Sub

Private Function Verificar()
    Dim t, s
    t = s2n(Right(TxtFecha, 2))
    If t > 30 Then
        s = Left(TxtFecha, 6) & "19" & Right(TxtFecha, 2)
    Else
        s = Left(TxtFecha, 6) & "19" & Right(TxtFecha, 2)
    End If
    
    Verificar = IsDate(s) Or TxtFecha = "  /  /  "
End Function

Private Sub UserControl_Initialize()
    'txtFecha.BackColor = UserControl.BackColor
    'txtFecha.ForeColor = UserControl.ForeColor

End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Dim s
    With TxtFecha
        If Chr(KeyAscii) = "/" Or KeyAscii = 13 Then
            Select Case .SelStart
            Case 0, 1:
                s = Format(val(Mid$(.Text, 1, 2)), "00")
                .Text = s & Right(.Text, 6)
                If s < "01" Or s > "31" Then
                    .SelStart = 0
                    .SelLength = 2
                Else
                    .SelStart = 3
                End If
            Case 3, 4:
                s = Format(val(Mid$(.Text, 4, 2)), "00")
                .Text = Left(.Text, 3) & s & Right(.Text, 3)
                If s < "01" Or s > "12" Then
                    .SelStart = 3
                    .SelLength = 2
                Else
                    .SelStart = 6
                End If
            Case 6, 7, 8:
                s = Format(val(Mid$(.Text, 7, 2)), "00")
                .Text = Left(.Text, 6) & s
                SendKeys "{tab}"
            End Select
            KeyAscii = 0
        ElseIf KeyAscii = 27 Then
            Me.dtFecha mUltFechaOk
            GotFocusPinto TxtFecha
        End If
    End With
End Sub


Private Sub UserControl_Resize()
    TxtFecha.Left = 0
    TxtFecha.Top = 0
    TxtFecha.Height = UserControl.Height
    TxtFecha.Width = UserControl.Width
End Sub

Private Function ff(sf As String) As String
    If UCase(Left(FormatoFecha(), 1)) = "M" Then
        ff = Format(sf, "mm/dd/yy")
    Else
        ff = Format(sf, "dd/mm/yy")
    End If
End Function

'7/4/5 permito selir de edicion con cancelar
'26/4/5 + enum ultDiaMes
'7/10/5 dtfecha yyyy a yy
'17/4/6 ult dia año
'           fecha limite para que no transforme 31/02/06 a 1931
'11/5/6 fix asignacion dtfecha dif configuracion regional
