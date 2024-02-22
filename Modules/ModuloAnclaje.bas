Attribute VB_Name = "ModuloAnclaje"
Option Explicit
'LitoSoft 9/6/5

' UNA sola funcion, va en el resize del form

'hay que especificarle un anclaje horizontal:    anclarIzquierda,   anclarDerecha,  anclarLadosAncho
' + uno vertical:                                anclarAbajo,       anclarArriba,   anclarLadosAlto

' Maximo, una linea por control
' o pones todo dentro de frames, y anclas frames

'genio

'**************************** EJEMPLO  ***********************
'Private Sub Form_Resize()
'    dim i as long

'    Anclar txtder, Me,     anclarDerecha + anclarArriba     ' se mueve con la pared derecha
'    Anclar Cmand1, Me,     anclarAbajo                      ' mantiene distancia al piso, flota entre los lados izq y der
'    Anclar grilla, Me,     anclarArriba  + anclarLadosAncho ' se estira a lo ancho

''frame con grilla
'    Anclar Frame1, Me,     anclarAbajo   + anclarLadosAncho ' idem
'    Anclar Text11, Frame1, anclarNinguno                    ' flota en el medio
'    Anclar grill2, Frame1, anclarLadosTodos                 ' modo Tupac, se estira a lo largo y a lo ancho
    
''frame con botones
'    Anclar FramB, Me, anclarDerecha + anclarArriba + anclarAbajo
'    For i = 0 To 4
'        Anclar Command3(i), FramB, anclarIzquierda          ' esto si queres q los botones se desparramen, flotan a lo largo
'    Next i

'End Sub
'**************************** EJEMPLO  ***********************

'futuro :

' agregar al enum anclajes elastico,
' asi si hay 2 cuadros grandes uno izq  y otro der,
' anclo fijo los bordes y ancla elastica al centro, asi no se pisan

Public Enum liAnclar
    anclarArriba = 1
    anclarAbajo = 2
    anclarIzquierda = 4
    anclarDerecha = 8

    anclarNinguno = 0       'quizas deberia borrar estos 4, creo que confunden
    anclarLadosAncho = 12   '
    anclarLadosAlto = 3     '
    anclarLadosTodos = 15   '
    
'    anclarFlexArriba = 16
'    anclarFlexAbajo = 32
'    anclarFlexIzquierda = 64
'    anclarFlexDerecha = 128
End Enum

' ----------------------------

Private Type CososAnclables
    ID  As Long
    Top As Long
    Left As Long
    Right As Long
    Bottom As Long
    cfTop  As Double
    cfLeft As Double
    cfBottom As Double
    cfRight As Double
End Type

Private cosos() As CososAnclables, iCosos As Long

Public Function Anclar(que As Object, donde As Object, opciones As liAnclar)
    On Error GoTo ufaChe
    Dim i As Long
    Dim coso As Variant
    For i = 1 To iCosos - 1
        If cosos(i).ID = que.hWnd Then
            encajar que, donde, i, opciones
            Exit Function
        End If
    Next
    agregocoso que, donde
ufaChe:
End Function

Private Sub agregocoso(que As Object, donde As Object)
    On Error Resume Next
    iCosos = iCosos + 1
    ReDim Preserve cosos(iCosos)
    With cosos(iCosos)
        .ID = que.hWnd
        .Bottom = donde.Height - que.Top - que.Height
        .Left = que.Left
        .Right = donde.Width - que.Width - que.Left
        .Top = que.Top
        .cfBottom = propo(.Bottom, donde.Height)
        .cfLeft = propo(que.Left, donde.Width)
        .cfRight = propo(.Right, donde.Width)
        .cfTop = propo(que.Top, donde.Height)
    End With
End Sub
Private Function propo(deque As Long, dedonde As Long) As Double
    On Error Resume Next
    propo = deque / dedonde
End Function

Private Sub encajar(oH As Object, oP As Object, ix As Long, opciones As liAnclar)
    On Error Resume Next
    Dim oPh As Long, oPw As Long
    oPh = oP.Height
    oPw = oP.Width
    
    Select Case opciones And anclarLadosAncho
     Case anclarNinguno
        oH.Left = minZ((oPw - oH.Width) * cosos(ix).Left / (cosos(ix).Left + cosos(ix).Right))
     Case anclarLadosAncho
        oH.Width = (oPw - cosos(ix).Right - cosos(ix).Left)
'    Case anclarIzquierda
     Case anclarDerecha
        oH.Left = minZ((oPw - oH.Width) - cosos(ix).Right)
    End Select
     
    Select Case opciones And anclarLadosAlto
     Case anclarNinguno
        oH.Top = minZ((oPh - oH.Height) * cosos(ix).Top / (cosos(ix).Top + cosos(ix).Bottom))
     Case anclarLadosAlto
        oH.Height = (oPh - cosos(ix).Top - cosos(ix).Bottom)
'    Case anclarArriba
     Case anclarAbajo
        oH.Top = minZ((oPh - oH.Height) - cosos(ix).Bottom)
    End Select
End Sub

Private Function minZ(que)
    minZ = IIf(que < 0, 0, que)
End Function
Public Sub CentrarMe(frmMe As Form) 'DEPRECATED- use propiedad .StartupPosition
    On Error Resume Next
    frmMe.Move (Screen.Width - frmMe.Width) \ 2, (Screen.Height - frmMe.Height) \ 2
End Sub
