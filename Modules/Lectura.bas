Attribute VB_Name = "Lectura"
Option Explicit

Public datobarra As String
Public ApliNudos As Excel.Application
Public LibroNudos As Excel.Workbook
Public HojaNudos As Excel.Worksheet
Public RangoNudos As Excel.Range
Public CeldaVacia As Long
Public Columnas As Long
Public Filas As Long
Public i As Long, j As Long
Public Sub Caracteristicas()
Set HojaNudos = LibroNudos.Sheets(1)
Set RangoNudos = HojaNudos.rows(1)
If (RangoNudos.Cells(1, 1) = "") Then
    CeldaVacia = 0
Else
    CeldaVacia = RangoNudos.Find("").Column
End If
Columnas = CeldaVacia
Set RangoNudos = HojaNudos.Columns(1)
If (RangoNudos.Cells(1, 1) = "") Then
    CeldaVacia = 0
Else
    CeldaVacia = RangoNudos.Find("").Row + 1
End If
Filas = CeldaVacia
Set HojaNudos = Nothing
Set RangoNudos = Nothing
End Sub


Public Sub Finalizar()
Set ApliNudos = Nothing
Set LibroNudos = Nothing
Set HojaNudos = Nothing
Set RangoNudos = Nothing
End Sub


Public Sub FormatearTabla()
'With frmLectura.grdNudos
With frmColector.Grilla
    .cols = Columnas
    .rows = Filas / 2
    .Visible = True
End With
End Sub

Public Sub Inicio()
Set ApliNudos = CreateObject("Excel.Application")
'Set LibroNudos = ApliNudos.Workbooks.Open(App.Path & "\Nudos.xls")
Set LibroNudos = ApliNudos.Workbooks.Open(frmColector.txtCarpetaRaizDbf.Text)
End Sub


Public Sub LlenadoDeTabla()
Dim cont As Long
cont = 1
Set HojaNudos = LibroNudos.Worksheets(1)
'For i = 1 To Columnas - 1
'    frmColector.grilla.Col = i
'    For j = 1 To Filas - 1
'        frmColector.grilla.Row = j
'        frmColector.grilla.Text = HojaNudos.Cells(j, i)
'    Next j
'Next i
'*************
i = 0
For j = 1 To Filas '- 1
    If Not frmColector.Grilla.rows = cont Then
        If Not frmColector.Grilla.Row = cont Then
            frmColector.Grilla.Row = cont
        End If
        'For i = 1 To 2
            'frmColector.grilla.Col = i
            'frmColector.grilla.Text = HojaNudos.Cells(j, 1)
            frmColector.Grilla.TextMatrix(cont, i) = HojaNudos.Cells(j, 1)
        'Next i
            If i = 0 Then
                i = 1
            Else
                i = 0
                cont = cont + 1
            End If
    End If
Next j
End Sub

Sub Main()
frmColector.Show
End Sub


'*********19/4/07*****VERIFICADOR DE DATOS******RAUL
Public Function Verificar_Dato(Dat As Variant, Mode As Long) As Variant
Dim datoV As Variant
datoV = ""
    Select Case Mode:
        Case 1: 'caso si es un entero
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Numero"
            Else
                datoV = Dat
            End If
        Case 2: 'caso si es una fecha
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Fecha"
            Else
                datoV = Dat
            End If
        Case 3: 'caso si es un string
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Datos"
            Else
                datoV = Dat
            End If
        Case 4: 'caso si es importe
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Valor"
            Else
                datoV = Dat
            End If
    End Select
Verificar_Dato = datoV
End Function

