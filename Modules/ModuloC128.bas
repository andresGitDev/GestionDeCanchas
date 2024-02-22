Attribute VB_Name = "ModuloC128"
Option Explicit

Public Function Code128$(chaine$)
  'V 1.1.0
  'Param�tres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affich�e avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si param�tre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, mini%, dummy%, tableB As Boolean
  Code128$ = ""
  If Len(chaine$) > 0 Then
  'V�rifier si caract�res valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 198
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    Code128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si int�ressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au d�but ou � la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub testnum
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'D�buter sur table C / Starting with table C
              Code128$ = Chr$(205)
            Else 'Commuter sur table C / Switch to table C
              Code128$ = Code128$ & Chr$(199)
            End If
            tableB = False
          Else
            If i% = 1 Then Code128$ = Chr$(204) 'D�buter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres / We are on table C, try to process 2 digits
          mini% = 2
          GoSub testnum
          If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
            dummy% = Val(Mid$(chaine$, i%, 2))
            dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 100)
            Code128$ = Code128$ & Chr$(dummy%)
            i% = i% + 2
          Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
            Code128$ = Code128$ & Chr$(200)
            tableB = True
          End If
        End If
        If tableB Then
          'Traiter 1 caract�re en table B / Process 1 digit with table B
          Code128$ = Code128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la cl� de contr�le / Calculation of the checksum
      For i% = 1 To Len(Code128$)
        dummy% = Asc(Mid$(Code128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 100)
        If i% = 1 Then checksum& = dummy%
        checksum& = (checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la cl� / Calculation of the checksum ASCII code
      checksum& = IIf(checksum& < 95, checksum& + 32, checksum& + 100)
      'Ajout de la cl� et du STOP / Add the checksum and the STOP
      Code128$ = Code128$ & Chr$(checksum&) & Chr$(206)
    End If
  End If
  Exit Function
testnum:
  'si les mini% caract�res � partir de i% sont num�riques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function

