Attribute VB_Name = "modFonctions"
'Copyright Renaud ILTIS, 2014
'
'renaudiltis@ yahoo.fr
'
'Ce logiciel est un programme informatique servant à piloter une machine de découpe par fil chaud.
'Ce logiciel est régi par la licence CeCILL soumise au droit français et respectant les principes de diffusion des logiciels libres. Vous pouvez
'utiliser, modifier et/ou redistribuer ce programme sous les conditionsde la licence CeCILL telle que diffusée par le CEA, le CNRS et l'INRIA
'sur le site "http://www.cecill.info".
'En contrepartie de l'accessibilité au code source et des droits de copie, de modification et de redistribution accordés par cette licence, il n'est
'offert aux utilisateurs qu'une garantie limitée. Pour les mêmes raisons, seule une responsabilité restreinte pèse sur l'auteur du programme,  le
'titulaire des droits patrimoniaux et les concédants successifs.
'A cet égard  l'attention de l'utilisateur est attirée sur les risques associés au chargement, à l'utilisation, à la modification et/ou au
'développement et à la reproduction du logiciel par l'utilisateur étant donné sa spécificité de logiciel libre, qui peut le rendre complexe à
'manipuler et qui le réserve donc à des développeurs et des professionnels avertis possédant des connaissances informatiques approfondies. Les
'utilisateurs sont donc invités à charger et tester l'adéquation du logiciel à leurs besoins dans des conditions permettant d'assurer la
'sécurité de leurs systèmes et ou de leurs données et, plus généralement, à l'utiliser et l'exploiter dans les mêmes conditions de sécurité.
'Le fait que vous puissiez accéder à cet en-tête signifie que vous avez pris connaissance de la licence CeCILL, et que vous en avez accepté les
'termes.
'
'------
'
'This software is a computer program whose purpose is to control a hot wire foam cutter.
'This software is governed by the CeCILL license under French law and abiding by the rules of distribution of free software.  You can  use,
'modify and/ or redistribute the software under the terms of the CeCILL license as circulated by CEA, CNRS and INRIA at the following URL
'"http://www.cecill.info".
'As a counterpart to the access to the source code and rights to copy, modify and redistribute granted by the license, users are provided only
'with a limited warranty and the software's author, the holder of the economic rights, and the successive licensors have only limited
'liability.
'In this respect, the user's attention is drawn to the risks associated with loading, using, modifying and/or developing or reproducing the
'software by the user in light of its specific status of free software, that may mean that it is complicated to manipulate, and that also
'therefore means that it is reserved for developers and experienced professionals having in-depth computer knowledge. Users are therefore
'encouraged to load and test the software's suitability as regards their requirements in conditions enabling the security of their systems and/or
'data to be ensured and, more generally, to use and operate it in the same conditions as regards security.
'The fact that you are presently reading this means that you have had knowledge of the CeCILL license and that you accept its terms.

Option Explicit

'****** MANIPULATION DES OCTETS <=> IPL5X **********
'Concaténation de 2 à 4 octets en un nombre entier
Public Function ConcatOctets(ByVal OctetL As Byte, ByVal OctetM As Byte, Optional ByVal OctetH As Byte = &H0&, Optional ByVal OctetU As Byte = &H0&) As Long
   ConcatOctets = OctetU * &H1000000 + OctetH * &H10000 + OctetM * &H100& + OctetL
End Function

'Transformation d'un entier en octets pour transfert vers IPL5X
'On considère la décomposition Lower - Medium - Higher - Upper de 16 à 32 bits (donc sur 16 bits, il faut utilier L et M)
Public Function LPart(Entier As Long) As Byte
   LPart = Entier And &HFF&
End Function

Public Function MPart(Entier As Long) As Byte
   MPart = (Entier And &HFF00&) / &H100&
End Function

Public Function HPart(Entier As Long) As Byte
   HPart = (Entier And &HFF0000) / &H10000
End Function

Public Function UPart(Entier As Long) As Byte
   UPart = (Entier And &HFF000000) / &H1000000
End Function

'*******************
'Procédure de temporisation, avec en paramètre le temps en secondes et la possibilité de bloquer le soft :
Public Sub Pause(ByVal DureeEnSecondes As Single, Optional ByVal ProgrammeBloque As Boolean = False)
   Dim t0 As Single
   Dim t As Single

   t0 = Timer  'La fonction timer renvoie le nombre de secondes écoulées depuis minuit ; il faut gérer le passage par 0
   Do
      If ProgrammeBloque = False Then  'par défaut le programme n'est pas bloqué
         DoEvents
      End If
      t = Timer
      If t < t0 Then t = t + 24 * 3600
   Loop While (t0 + DureeEnSecondes) > t
End Sub

'Fonction qui renvoie le répertoire du fichier à partir du chemin complet SANS l'antislash final
Public Function ExtraitRepertoireFichier(ByVal NomComplet As String) As String
   If Right(NomComplet, 1) = "\" Then
       ExtraitRepertoireFichier = NomComplet
   Else
       ExtraitRepertoireFichier = Left(NomComplet, InStrRev(NomComplet, "\") - 1)
   End If
End Function


'Fonction qui renvoie le nom du fichier à partir du chemin complet
Public Function ExtraitNomFichier(ByVal CheminComplet As String) As String
    If InStr(CheminComplet, "\") = 0 Or Right(CheminComplet, 1) = "\" Then
        ExtraitNomFichier = ""
        Exit Function
    End If
    ExtraitNomFichier = Mid(CheminComplet, InStrRev(CheminComplet, "\") + 1)
End Function

'Fonction qui renvoie l'extension du fichier à partir du chemin complet
Public Function ExtraitExtensionFichier(ByVal CheminComplet As String) As String
    Dim Nom As String
    Nom = ExtraitNomFichier(CheminComplet)  'on utilise la fonction ci-dessus
    If InStr(Nom, ".") = 0 Then
        ExtraitExtensionFichier = ""
    Else
        ExtraitExtensionFichier = Mid(Nom, InStrRev(Nom, ".") + 1)
    End If
End Function

'Fonction qui renvoie le nom du fichier sans l'extension à partir du nom avec l'extension (sans le chemin)
Public Function EnleveExtensionFichier(ByVal NomComplet As String) As String
    If InStr(NomComplet, ".") = 0 Then
        EnleveExtensionFichier = NomComplet
    Else
        EnleveExtensionFichier = Mid(NomComplet, 1, InStrRev(NomComplet, ".") - 1)
    End If
End Function

Public Sub VerifierDecimaux(ByRef KeyAscii As Integer)
'Vérification des frapes numériques
   Select Case KeyAscii
    Case 44, 46, 48 To 57, 8, 13     'Autorise la virgule, le point, les touches 0 à 9 et backspace 45 pour le signe moins
    Case Else 'sinon
        KeyAscii = 0       'on annule la touche
    End Select
End Sub

Public Sub VerifierDecimauxRelatifs(ByRef KeyAscii As Integer)
'Vérification des frapes numériques
   Select Case KeyAscii
    Case 44, 45, 46, 48 To 57, 8, 13, 9    'Autorise la virgule, le moins, le point, les touches 0 à 9 et backspace 45 pour le signe moins
    Case Else 'sinon
        KeyAscii = 0       'on annule la touche
    End Select
End Sub

Public Sub VerifierEntiers(ByRef KeyAscii As Integer)
'Vérification des frapes numériques
   Select Case KeyAscii
    Case 48 To 57, 8, 13, 9 '(9=tabulation)
    Case Else 'sinon
        KeyAscii = 0       'on annule la touche
    End Select
End Sub

'*********************************************************
'Désactivation/Activation de tous les contrôles d'un frame
'*********************************************************
Public Sub DesactiverFrame(ByVal NomDeLaForm As Form, ByVal NomDuFrame As Control, ByVal ValeurDeEnabled As Boolean)
   Dim ctl As Control
   For Each ctl In NomDeLaForm.Controls
      If ctl.Container Is NomDuFrame Then
         If Not TypeOf ctl Is Line Then  'les contrôles Line n'ont pas la propriété Enabled
            ctl.Enabled = ValeurDeEnabled
         Else  'si ce sont des lignes, on va jouer sur la couleur
            If ValeurDeEnabled = False Then
               ctl.BorderColor = &H808080
            Else
               ctl.BorderColor = &H0
            End If
         End If
      End If
   Next
End Sub

'**************************************
'**** Transfformation de 2,3 en 2.3 ****
'**************************************
Public Function V2P(strNombre As String) As String
'Virgule2Point change la première virgule rencontrée en point
'la fonction Format peut renvoyer une virgule, donc faire V2P(Format(strxx,"..."))
   Dim i As Integer
   V2P = strNombre
   For i = 1 To Len(strNombre)
      If Mid$(strNombre, i, 1) = "," Then
         Mid(strNombre, i, 1) = "."
         V2P = strNombre
         Exit For
      End If
   Next i
End Function

'----------------------------------------------------------------------
' Conversion String -> Nombre quelque soit le séparateur décimal ou de
' millier contenu dans le texte d'origine.
'----------------------------------------------------------------------
' Comme Val ne marche pas avec les symboles décimaux, et que C*()
' Renvoient une erreur de type si le séparateur n'est pas le bon :
' Appel de la fonction Val après avoir Transfformé les éventuels
' séparateurs en simple point décimal et éliminé les séparateurs
' de milliers.
'----------------------------------------------------------------------
' Exemple d'utilisation :
'Dim Nombre as Long
'Nombre = MyVal(Text1.Text)
Public Function MyVal(Chaine As Variant) As String
    Dim strTmp As String, charTmp As String
    Dim i As Long
    Dim SepDecimal As String, SepMillier As String
    SepDecimal = Format$(0, ".")
    SepMillier = Mid$(Format$(1000, "0 000"), 2, 1)
    strTmp = ""
    If IsNull(Chaine) Then
        MyVal = Null
        Exit Function
    End If
    For i = 1 To Len(Chaine)
        charTmp = Mid(Chaine, i, 1)
        If Asc(charTmp) >= 48 And Asc(charTmp) <= 57 Then
            ' C'est un chiffre, on le traite.
            strTmp = strTmp & charTmp
        ElseIf charTmp = "+" Or charTmp = "-" Then
            ' Le signe...
            strTmp = strTmp & charTmp
        ElseIf charTmp = "," Or charTmp = "." Or charTmp = SepDecimal Then
            ' La virgule, le point ou un autre séparateur, même combat !
            strTmp = strTmp & "."
        ElseIf charTmp = " " Or charTmp = Chr$(160) Or charTmp = SepMillier Then
            ' Séparateur de milliers éliminé
        Else
            ' Fin de la boucle au premier
            ' caractère non numérique comme le fait normalement
            ' la fonction Val().
            Exit For
        End If
    Next
    MyVal = Val(strTmp)
End Function

'------------------------------------------------------
'Vérification qu'une chaîne de caractères est un nombre
'par la méthode des automates
'------------------------------------------------------
'une fonction publique et 4 fonctions privées
'------------------------------------------------------
Public Function MyIsNumeric(S As String) As Boolean
   Dim Etat As Integer
   Dim car As String
   Dim L As Long
   Dim i As Long
   Dim validStates As String
   validStates = "1,4,"
   L = Len(S) ' longueur de la chaine d'entrée
   If L > 0 Then ' si chaine pas vide
      If Mid$(S, L, 1) = "." Or Mid$(S, L, 1) = "," Then
         S = S & "0"     'rajout perso pour autoriser la saisie de 100. qui est remplacé par 100.0
         L = L + 1
      End If
      For i = 1 To L ' puis on avance caractere par caractere
         car = Mid$(S, i, 1)
         Select Case Etat ' et on suit l'automate
         Case 0
            If IsDigit(car) Then
               Etat = 1
            ElseIf IsMinusSign(car) Then
               Etat = 2
            ElseIf IsPlusSign(car) Then
               Etat = 2
            ElseIf IsDot(car) Then
               Etat = 3
            Else
               Etat = 5 ' fini
            End If
         Case 1
            If IsDigit(car) Then
               Etat = 1
            ElseIf IsDot(car) Then
               Etat = 3
            Else
               Etat = 5 ' fini
            End If
         Case 2
            If IsDigit(car) Then
               Etat = 1
            ElseIf IsDot(car) Then
               Etat = 3
            End If
         Case 3
            If IsDigit(car) Then
               Etat = 4
            End If
         Case 4
            If IsDigit(car) Then
               Etat = 4
            Else
               Etat = 5 ' fini
            End If
         Case 5
            ' etat invalide, pas de Transfitions
            ' on peut quitter ici si on veut
         End Select
      Next i
   End If
   If InStr(validStates, Trim$(Str$(Etat)) & ",") > 0 Then
      MyIsNumeric = True
   Else
      MyIsNumeric = False
   End If
End Function

Public Function IsDigit(car As String) As Boolean
   If car >= "0" And car <= "9" Then
      IsDigit = True
   End If
End Function

Public Function IsMinusSign(car As String) As Boolean
   If car = "-" Then
      IsMinusSign = True
   End If
End Function

Public Function IsPlusSign(car As String) As Boolean
   If car = "+" Then
      IsPlusSign = True
   End If
End Function

Public Function IsDot(car As String) As Boolean
   If car = "." Or car = "," Then
      IsDot = True
   End If
End Function

'********************************************************************
'**** Calcul de l'ANGLE d'un SEGMENT par rapport à l'HORIZONTALE ****
'********************************************************************
Public Function AngleSegment(ByVal XA As Single, ByVal YA As Single, ByVal XB As Single, ByVal YB As Single) As Single
   Dim DeltaX As Single, DeltaY As Single
   
   DeltaX = XB - XA
   DeltaY = YB - YA
   If DeltaX = 0 Then   'On traite le cas particulier d'un segment horizontal
      If DeltaY < 0 Then
         AngleSegment = -pi / 2
      Else
         AngleSegment = pi / 2
      End If
   Else              'dans les autres cas
      AngleSegment = Atn(DeltaY / DeltaX)
      If DeltaX < 0 Then
         If DeltaY < 0 Then
            AngleSegment = AngleSegment - pi
         Else
            AngleSegment = AngleSegment + pi
         End If
      End If
   End If
   If AngleSegment < 0 Then
      AngleSegment = AngleSegment + 2 * pi
   End If
End Function

'**********************************************************************
'**** NETTOYAGE des PROFILS (points presque alignés) A L'OUVERTURE ****
'**********************************************************************
Public Function Nettoyage(ByVal Dist As Single) As Long     'renvoie le nombre de points

'Tassement du tableau : suppression des points "presque" alignés
   Dim i As Long, j As Long, k As Long
   Dim lined As Long
   Dim Nbr_Out As Long
   Dim A As Single, B As Single
   Dim d2 As Single, d1 As Single
   Dim epsilon As Single
   
   If UBound(profil) < 3 Then
      Nettoyage = UBound(profil) + 1
      Exit Function
   End If
   epsilon = Dist * Dist  'On élève au carré pour aller plus vite pour les comparaisons
   i = 0
   Nbr_Out = 1
   Do
      j = 1
      Do
         lined = 0
         j = j + 1
         For k = 1 To j - 1
            'Est-ce que tous les points intermédiaires sont alignés ?
            A = profil(i + j).x - profil(i).x
            B = profil(i + j).y - profil(i).y
            d2 = A * A + B * B
            If d2 <= epsilon Then
               ' Les points i and i+j sont trop près
               d1 = (profil(i + k).x - profil(i).x) * (profil(i + k).x - profil(i).x) + _
                  (profil(i + k).y - profil(i).y) * (profil(i + k).y - profil(i).y)
               If d1 <= epsilon Then
                  lined = lined + 1         ' Le point i+k est aussi trop près
               Else
                  Exit For                  ' Le point i+k n'est pas aligné, pas la peine de continuer !
               End If
            Else
               d1 = ((profil(i + k).x - profil(i).x) * A + (profil(i + k).y - profil(i).y) * B) / d2
               If d1 < 0# Then
                  d1 = 0#                   ' La référence est le point i
               Else
                  If d1 > 1# Then d1 = 1#   ' La référence est le point i+j
               End If
               A = profil(i).x + d1 * A - profil(i + k).x
               B = profil(i).y + d1 * B - profil(i + k).y
               If A * A + B * B <= epsilon Then
                  lined = lined + 1         ' point aligné
               Else
                  Exit For                  ' Le point i+k n'est pas aligné, pas besoin de continuer
               End If
            End If
         Next k
      Loop While lined = j - 1 And i + j + 1 <= UBound(profil) 'Boucle jusqu'à ce qu'un point ne soit pas aligné
      i = i + j - 1
      If lined <> j - 1 Then
        profil(Nbr_Out) = profil(i)
        Nbr_Out = Nbr_Out + 1
      End If
   Loop While (i <= UBound(profil) - 2)
   profil(Nbr_Out) = profil(UBound(profil))
   Nbr_Out = Nbr_Out + 1
   'Redimensionnement et réaffichage
   ReDim Preserve profil(0 To Nbr_Out - 1)
   Nettoyage = Nbr_Out
End Function

'**** Récupérer le chemin d'un fichier ****
Public Function GetPathName(ByVal strPath) As String

    Dim intLenPath As Integer
    Dim intSlash As Integer
     
    On Error GoTo GetPathName_Err
     
    intLenPath = Len(strPath)
     
    If intLenPath > 0 Then
        strPath = StrReverse(strPath)
        intSlash = InStr(1, strPath, "\", vbTextCompare)
        If intSlash > 0 Then
            strPath = Right$(strPath, intLenPath - (intSlash - 1))
            strPath = StrReverse(strPath)
            GetPathName = strPath
        ElseIf intSlash = 0 Then
            'No slash found, may not be a valid path name.
            GetPathName = vbNullString
        End If
    ElseIf intLenPath = 0 Then
        GetPathName = vbNullString
    End If
     
GetPathName_Exit:
    Exit Function
     
GetPathName_Err:
    GetPathName = vbNullString
    Resume GetPathName_Exit
End Function

'*************************************************************************************
'********** Procédure d'extraction des séquences à partir du tableau total ***********
'*************************************************************************************
Public Sub ExtractionSequ()
   Dim i As Long, j As Long, k As Long
   Dim NbPoints As Long
   
   '*** Contrôle de l'UNICITE de la Sequ pour extraction si besoin***
   If profil(UBound(profil)).NumSequ <> profil(LBound(profil)).NumSequ Then
      '*** EXTRACTION DES Sequ ***
      ReDim Sequ(1 To 1)   'pour vider la mémoire
      ReDim SequTrace(1 To 1)   'pour vider la mémoire
      i = 1 'numéro de la séquence
      j = 0 'numéro de ligne du tableau Profil()
      Do
         ReDim Preserve Sequ(1 To i)
         ReDim Preserve SequTrace(1 To i)
         k = 1 'numéro du point dans la séquence
         Do
            ReDim Preserve Sequ(i).Point(1 To k)
            ReDim Preserve SequTrace(i).Point(1 To k)
            Numero = profil(j).NumSequ
            Sequ(i).Point(k).x = profil(j).x
            Sequ(i).Point(k).y = profil(j).y
            j = j + 1
            k = k + 1
            If j = UBound(profil) + 1 Then
               Exit Do
            End If
         Loop Until profil(j).NumSequ <> Numero
         Sequ(i).NbPoints = k - 1
         If j = UBound(profil) + 1 Then
            Exit Do
         End If
         i = i + 1
      Loop
   Else
      ReDim Sequ(1 To 1)
      ReDim Sequ(1).Point(1 To UBound(profil) + 1)
      ReDim SequTrace(1 To 1)
      ReDim SequTrace(1).Point(1 To UBound(profil) + 1)
      Sequ(1).NbPoints = UBound(Sequ(1).Point)
      For i = 1 To Sequ(1).NbPoints
         Sequ(1).Point(i).x = profil(i - 1).x
         Sequ(1).Point(i).y = profil(i - 1).y
      Next i
   End If
   NbSequ = UBound(Sequ)
   
   For i = 1 To NbSequ
      Sequ(i).Etat = 0 'initialisation de l'état
      SequTrace(i).Etat = 0 'initialisation de l'état
   Next i
   
   'Nettoyage des séquences ; cf. NettoyageSimple
   For j = 1 To NbSequ
      Call MaxiMiniSequ(Sequ(j))
      With Sequ(j)
         ReDim profil(0 To .NbPoints - 1)
         For i = 1 To .NbPoints
            profil(i - 1).x = .Point(i).x
            profil(i - 1).y = .Point(i).y
         Next i
         If .DeltaX >= .DeltaY Then
            EpsilonNettoyage = .DeltaX / CoeffNett     'CoeffNett défini dans le Form Load
         Else
            EpsilonNettoyage = .DeltaY / CoeffNett
         End If
         NbPoints = Nettoyage(EpsilonNettoyage)     'nettoyage des points trop rapprochés
         .NbPoints = NbPoints
         ReDim .Point(1 To .NbPoints)               'tableau du profil initial, inclus dans le type "Trajet"
         ReDim SequTrace(j).Point(1 To .NbPoints)
         For i = 1 To .NbPoints                      'Transffert des points dans mon type de tableau
            .Point(i).x = profil(i - 1).x
            .Point(i).y = profil(i - 1).y
         Next i
      End With
      Call MaxiMiniSequ(Sequ(j))
   Next j
   'Suppression des points doubles (si point simple dans le dxf, CNCTools envoie deux points avec PU et PD, cf. type)
   For j = 1 To NbSequ
      With Sequ(j)
         If .NbPoints = 2 Then
            If .Point(1).x = .Point(2).x And .Point(1).y = .Point(2).y Then
               ReDim Preserve .Point(1 To 1) 'on vire le doublon
               .NbPoints = 1
            End If
         End If
      End With
   Next j
   Erase profil  'libérer la mémoire
   NumSequSel = 0
End Sub

'****** FONCTION de CALCUL des MAXI/MINI d'une Sequ ******
Public Sub MaxiMiniSequ(ByRef Sequ As Sequ)
   Dim i As Long
   '*** calcul des maxi et mini d'une séquence ***
   With Sequ
      .Xmin = .Point(1).x
      .Xmax = .Point(1).x
      .Ymin = .Point(1).y
      .Ymax = .Point(1).y
      For i = 1 To .NbPoints
         If .Point(i).x <= .Xmin Then .Xmin = .Point(i).x
         If .Point(i).x >= .Xmax Then .Xmax = .Point(i).x
         If .Point(i).y <= .Ymin Then .Ymin = .Point(i).y
         If .Point(i).y >= .Ymax Then .Ymax = .Point(i).y
      Next i
      .DeltaX = Abs(.Xmax - .Xmin)
      .DeltaY = Abs(.Ymax - .Ymin)
   End With
End Sub

'********* fonctions de conversion Pixels <=> Twips ************
Public Function PixelsToTwips(pixels As Integer) As Single
   PixelsToTwips = pixels * Screen.TwipsPerPixelX
End Function

Public Function TwipsToPixels(twips As Integer) As Single
   TwipsToPixels = twips / Screen.TwipsPerPixelX
End Function

'******** Inversion du sens d'une séquence ********
Public Sub InverserSensSequ(ByRef Sequ As Sequ)
   Dim i As Long
   
   With Sequ
      ReDim Preserve .Point(1 To .NbPoints + .NbPoints)
      For i = 1 To .NbPoints
         .Point(.NbPoints + i) = .Point(.NbPoints + 1 - i)
      Next i
      For i = 1 To .NbPoints
         .Point(i) = .Point(i + .NbPoints)
      Next i
      ReDim Preserve .Point(1 To .NbPoints)
   End With
End Sub

'inversion de l'ordre des groupes de séquences qui se suivent dans Transf
Public Sub InverserOrdreTransf()
   Dim i As Long, j As Long, u As Long, L As Long
   Dim flagSection As Boolean
   Dim TransfTemp2() As Sequ
   
   flagSection = False
   ReDim Preserve Transf(1 To NbTransf + 1) 'on ajoute une ligne pour permettre le test sur la dernière séquence
   Transf(NbTransf + 1).Etat = 0
   For i = 1 To NbTransf + 1
      If (Transf(i).Etat And 1) = 1 And flagSection = False Then
         ReDim TransfTemp2(i To i)
         flagSection = True
         TransfTemp2(i) = Transf(i)
      ElseIf (Transf(i).Etat And 1) = 1 And flagSection = True Then
         ReDim Preserve TransfTemp2(LBound(TransfTemp2) To i)
         TransfTemp2(i) = Transf(i)
      ElseIf (Transf(i).Etat And 1) = 0 And flagSection = True Then
         flagSection = False
         'inverser TransfTemp et le réinjecter
         L = LBound(TransfTemp2)
         u = UBound(TransfTemp2)
         For j = 0 To (u - L)
            Transf(L + j) = TransfTemp2(u - j)
         Next j
      End If
   Next i
   ReDim Preserve Transf(1 To NbTransf)  'on vire la ligne en trop
   Erase TransfTemp2  'pour vider la mémoire
End Sub

'inversion de l'ordre de toutes les séquences de Transf
Public Sub InverserOrdreToutTransf()
   Dim i As Long, j As Long, u As Long, L As Long
   Dim TransfTemp2() As Sequ
   
   ReDim TransfTemp2(1 To NbTransf)
   For i = 1 To NbTransf
         TransfTemp2(i) = Transf(i)
   Next i
   'inverser TransfTemp et le réinjecter
   L = LBound(TransfTemp2)
   u = UBound(TransfTemp2)
   For j = 0 To (u - L)
      Transf(L + j) = TransfTemp2(u - j)
   Next j
   Erase TransfTemp2   'pour vider la mémoire
End Sub

'******** Savoir s'il y a une séquence sous le curseur en bas ***********
Public Function SousCursTransf(ByVal x As Single, ByVal y As Single) As Long
   Dim i As Long
   Dim EpsilonX As Single, EpsilonY As Single
   
   SousCursTransf = 0
   If NbTransf > 0 Then    'la recherche n'est valable que s'il y a des séquences transférées
      For i = 1 To NbTransf
         With Transf(i)
            If .Xmin = .Xmax Then      'segment vertical
               EpsilonX = 4
            Else
               EpsilonX = 0
            End If
            If .Ymin = .Ymax Then      'segment horizontal
               EpsilonY = 4
            Else
               EpsilonY = 0
            End If
            If x >= .Xmin - EpsilonX And x <= .Xmax + EpsilonX Then
               If y >= .Ymin - EpsilonY And y <= .Ymax + EpsilonY Then 'on sépare la contrôle suivant X et Y pour aller plus vite
                  SousCursTransf = i
                  Exit For
               End If
            End If
         End With
      Next i
   End If
End Function

'***************************************
'**** INTERSECTION de deux segments ****
'***************************************
Public Function Intersection(ByVal XA As Single, ByVal YA As Single, ByVal XB As Single, ByVal YB As Single, _
      ByVal XC As Single, ByVal YC As Single, ByVal XD As Single, ByVal YD As Single, _
      ByRef XI As Single, ByRef YI As Single) As Long
      'la fonction renvoie 0 si // ou colinéaires, 1 si intersection sur les segments, 2 intersection hors des segments,
      '3 si intersection sur A, 4 si intersection sur B, 5 si intersection sur C, 6 si intersection sur D
   Dim Determinant As Single
   Dim t1 As Single, t2 As Single
   
   Determinant = (XB - XA) * (YC - YD) - (XC - XD) * (YB - YA)
   If Determinant = 0 Then
      Intersection = 0
      Exit Function
   Else
      t1 = ((XC - XA) * (YC - YD) - (XC - XD) * (YC - YA)) / Determinant
      t2 = ((XB - XA) * (YC - YA) - (XC - XA) * (YB - YA)) / Determinant
      If t1 > 1 Or t1 < 0 Or t2 > 1 Or t2 < 0 Then
         Intersection = 2     'intersection hors des segments
         XI = XA + t1 * (XB - XA)
         YI = YA + t1 * (YB - YA)
         Exit Function
      Else
         If t1 = 0 Or t1 = 1 Then
            If t1 = 0 Then
               Intersection = 3  'intersection sur le point A
               XI = XA
               YI = YA
               Exit Function
            Else
               Intersection = 4  'intersection sur le point B
               XI = XB
               YI = YB
               Exit Function
            End If
         
         ElseIf t2 = 0 Or t2 = 1 Then
            If t2 = 0 Then
               Intersection = 5  'intersection sur le point C
               XI = XC
               YI = YC
               Exit Function
            Else
               Intersection = 6  'intersection sur le point D
               XI = XD
               YI = YD
               Exit Function
            End If
         Else
            Intersection = 1  'intersection sur les segments
            XI = XA + t1 * (XB - XA)
            YI = YA + t1 * (YB - YA)
         End If
      End If
   End If
End Function

'******************* Fonctions pour gestion IPL ****************
Public Sub Activer(FormActive As Form)
   Dim ctl As Control
   For Each ctl In FormActive.Controls
    If TypeOf ctl Is CommandButton Or TypeOf ctl Is OptionButton Then
      ctl.Enabled = True
    End If
   Next
End Sub

'Désactivation/activation des boutons de la form
Public Sub Desactiver(FormActive As Form)
   Dim ctl As Control
   For Each ctl In FormActive.Controls
    If TypeOf ctl Is CommandButton Or TypeOf ctl Is OptionButton Then
      ctl.Enabled = False
    End If
   Next
End Sub




'********************************************************************
'**** Ensemble de fonctions permettant la conversion entre bases ****
'********************************************************************
'Exemple d'utilisation : OUT(5) = BaseToBase("00000011", 2, 10)

Public Function DecimalToBase(ByVal nNumber As Long, ByVal DstBase As Long) As String
    Do While (nNumber >= DstBase)
        DecimalToBase = NumberToSymbol((nNumber Mod DstBase), DstBase) & DecimalToBase
        If DstBase >= 36 Then DecimalToBase = "." & DecimalToBase
        nNumber = nNumber \ DstBase
    Loop
    DecimalToBase = NumberToSymbol(nNumber, DstBase) & DecimalToBase
End Function

Public Function NumberToSymbol(ByVal nNumber As Long, ByVal DestBase As Long) As String
    If ((nNumber >= 10) And (nNumber < 36) And DestBase < 36) Then
        NumberToSymbol = Chr(Asc("A") + (nNumber - 10))
    Else
        NumberToSymbol = nNumber
    End If
End Function

Public Function BaseToDecimal(ByVal sNumber As String, ByVal SrcBase As Long) As Long
    Dim i As Integer
    Dim v() As String

    If SrcBase < 36 Then
        For i = 0 To Len(sNumber) - 1
            BaseToDecimal = BaseToDecimal + SymbolToNumber(Mid(sNumber, Len(sNumber) - i, 1)) * (SrcBase ^ i)
        Next
    Else
        v() = Split(sNumber, ".")
        For i = 0 To UBound(v)
            BaseToDecimal = BaseToDecimal + SymbolToNumber(v(UBound(v) - i)) * (SrcBase ^ i)
        Next
    End If
End Function

Public Function SymbolToNumber(ByVal sSymbol As String) As Long
    If Len(sSymbol) = 1 And Asc(UCase(sSymbol)) >= Asc("A") And Asc(UCase(sSymbol)) <= Asc("Z") Then
        SymbolToNumber = (Asc(UCase(sSymbol)) - Asc("A")) + 10
    Else
        SymbolToNumber = CLng(sSymbol)
    End If
End Function

Public Function BaseToBase(ByVal vNumber As Variant, ByVal SrcBase As Long, ByVal DstBase As Long) As Variant
    Dim nDecTemp As Long

    If (SrcBase <> 10) Then
        nDecTemp = BaseToDecimal(vNumber, SrcBase)
    Else
        nDecTemp = vNumber
    End If
    If (DstBase <> 10) Then
        BaseToBase = DecimalToBase(nDecTemp, DstBase)
    Else
        BaseToBase = nDecTemp
    End If
End Function

'**************************************************
'********* Arrondir à l'entier supérieur **********
'**************************************************
' Passez une valeur double dans la function. Par exemple:
' myNum = 11/5; myNum = roundUp(myNum)
Public Function roundUp(myNum As Double) As Double
   roundUp = -Int(-(myNum))
End Function

'************** Procédure de calcul de points de la simulation **************
Public Sub FaireLeFilm(ByRef ListePoints() As PointSimple, ByVal Pas As Single)  'la liste des points sera remplacée par ceux du film
   Dim i As Long, N As Long, Sgm As Segment, Sgms() As Segment
   Dim X1 As Single, X2 As Single, dx As Single, Y1 As Single, Y2 As Single, dy As Single
   Dim L() As Single, A As Long, B As Long, C As Long
   Dim u As Single, k As Single, D As Single
   Dim XA As Single, YA As Single, XB As Single, YB As Single, x As Single, y As Single
   Dim a0 As Long
   Dim PointsDuFilm() As PointSimple
   Dim NbrePointsDuFilm As Long
   
   'Récupération du nombre de segments
   N = UBound(ListePoints) - 1
   
   'Construction et dessin des segments
   ReDim Sgms(0 To N)
   For i = 1 To N
       With Sgm
           .Point1 = ListePoints(i): .Point2 = ListePoints(i + 1)
           X1 = .Point1.x: X2 = .Point2.x: dx = X1 - X2
           Y1 = .Point1.y: Y2 = .Point2.y: dy = Y1 - Y2
           .Longueur = Sqr(dx * dx + dy * dy)
       End With
       Sgms(i) = Sgm
   Next i
   
   'Construction de la fonction de linéarisation
   ReDim L(0 To N + 1)
   L(1) = 0
   For i = 2 To N + 1
       L(i) = L(i - 1) + Sgms(i - 1).Longueur
   Next i
   
   'Calcul et affichage des points
   i = 0: a0 = 1
   NbrePointsDuFilm = 0
   Do
       D = i * Pas
       If D > L(N + 1) Then
           Exit Do
       Else
           'Calcul et affichage du point
           If D = L(N + 1) Then
               A = N: u = Sgms(N).Longueur
           Else
               A = a0: B = N + 1
               Do
                   If B - A = 1 Then
                       Exit Do
                   Else
                       C = Int((A + B) / 2)
                       If D < L(C) Then B = C Else A = C
                   End If
               Loop
           End If
           Sgm = Sgms(A)
           u = D - L(A)
           k = u / Sgm.Longueur
           With ListePoints(A)
               XA = .x: YA = .y
           End With
           With ListePoints(B)
               XB = .x: YB = .y
           End With
           x = XA + k * (XB - XA)
           y = YA + k * (YB - YA)
           
           NbrePointsDuFilm = NbrePointsDuFilm + 1
           ReDim Preserve PointsDuFilm(1 To NbrePointsDuFilm)
           PointsDuFilm(NbrePointsDuFilm).x = x
           PointsDuFilm(NbrePointsDuFilm).y = y
           
           'Itération
           i = i + 1: a0 = A
       End If
   Loop
   'On transfère les résultats dans le tableau source
   ReDim ListePoints(1 To NbrePointsDuFilm)
   For i = 1 To NbrePointsDuFilm
      With ListePoints(i)
         .x = PointsDuFilm(i).x
         .y = PointsDuFilm(i).y
      End With
   Next i
   Erase PointsDuFilm   'on libère la mémoire
   Erase L
   Erase Sgms
End Sub

'********************************************************
'******* Vérification de l'existence d'un dossier *******
'********************************************************
Public Function VerifierExistenceRepertoire(ByVal Chemin As String) As Boolean

    Dim fs As Scripting.FileSystemObject
    Set fs = New Scripting.FileSystemObject  'objet instancié
    ' test si le repertoire existe
    If Not fs.FolderExists(Chemin) Then
      '  création du repertoire si besoin
      VerifierExistenceRepertoire = False
      'soit dit en passant, on pourrait le créer avec   fs.CreateFolder (Chemin)
    Else
      VerifierExistenceRepertoire = True
    End If
    Set fs = Nothing    'desctruction de l'objet créé
End Function

Function VerifierExistenceFichier(ByVal sFileName As String) As Boolean
    On Error Resume Next
    VerifierExistenceFichier = ((GetAttr(sFileName) And vbDirectory) = 0)
End Function

'*******************************
'**** DECALAGE d'un SEGMENT ****
'*******************************
Public Sub Decale(ByRef X1 As Single, ByRef X2 As Single, ByRef Y1 As Single, ByRef Y2 As Single, ByVal Decalage As Single)
   Dim Xtranslation As Single, Ytranslation As Single
   'On traite d'abord les cas des segments horizontaux et verticaux
   'on part sur une base de rotation anti-horaire
   If Y1 = Y2 And X2 > X1 Then
      X1 = X1
      X2 = X2
      Y1 = Y1 + Decalage
      Y2 = Y2 + Decalage
      GoTo Fin
  End If
  If Y1 = Y2 And X2 < X1 Then
      X1 = X1
      X2 = X2
      Y1 = Y1 - Decalage
      Y2 = Y2 - Decalage
      GoTo Fin
  End If
  If X1 = X2 And Y2 > Y1 Then
      X1 = X1 - Decalage
      X2 = X2 - Decalage
      Y1 = Y1
      Y2 = Y2
      GoTo Fin
  End If
  If X1 = X2 And Y2 < Y1 Then
      X1 = X1 + Decalage
      X2 = X2 + Decalage
      Y1 = Y1
      Y2 = Y2
      GoTo Fin
  End If
   'dans les autres cas
  If X1 <> X2 And Y1 <> Y2 Then
      Xtranslation = Decalage * ((Y1 - Y2) / (X2 - X1)) / (Sqr((Y1 - Y2) * (Y1 - Y2) / (X2 - X1) / (X2 - X1) + 1 * 1))
      Ytranslation = Decalage * 1 / (Sqr((Y1 - Y2) * (Y1 - Y2) / (X2 - X1) / (X2 - X1) + 1 * 1))
      If Y2 > Y1 And Xtranslation < 0 Then
          X1 = X1 + Xtranslation
          X2 = X2 + Xtranslation
          Y1 = Y1 + Ytranslation
          Y2 = Y2 + Ytranslation
          GoTo Fin
      End If
      If Y2 < Y1 And Xtranslation < 0 Then
          X1 = X1 - Xtranslation
          X2 = X2 - Xtranslation
          Y1 = Y1 - Ytranslation
          Y2 = Y2 - Ytranslation
          GoTo Fin
      End If
      If Y2 > Y1 And Xtranslation > 0 Then
          X1 = X1 - Xtranslation
          X2 = X2 - Xtranslation
          Y1 = Y1 - Ytranslation
          Y2 = Y2 - Ytranslation
          GoTo Fin
      End If
      If Y2 < Y1 And Xtranslation > 0 Then
          X1 = X1 + Xtranslation
          X2 = X2 + Xtranslation
          Y1 = Y1 + Ytranslation
          Y2 = Y2 + Ytranslation
          GoTo Fin
      End If
  End If
Fin:
End Sub


'***************************************
'**** INTERSECTION de deux segments ****
'***************************************
Public Function ChercherIntersection(ByVal XA As Single, ByVal YA As Single, ByVal XB As Single, ByVal YB As Single, _
      ByVal XC As Single, ByVal YC As Single, ByVal XD As Single, ByVal YD As Single, _
      ByRef XI As Single, ByRef YI As Single) As Long
      'la fonction renvoie 0 si // ou colinéaires, 1 si intersection sur les segments, 2 intersection hors des segments,
      '3 si intersection sur A, 4 si intersection sur B, 5 si intersection sur C, 6 si intersection sur D
   Dim Determinant As Single
   Dim t1 As Single, t2 As Single
   
   Determinant = (XB - XA) * (YC - YD) - (XC - XD) * (YB - YA)
   If Determinant = 0 Then
      ChercherIntersection = 0
      Exit Function
   Else
      t1 = ((XC - XA) * (YC - YD) - (XC - XD) * (YC - YA)) / Determinant
      t2 = ((XB - XA) * (YC - YA) - (XC - XA) * (YB - YA)) / Determinant
      If t1 > 1 Or t1 < 0 Or t2 > 1 Or t2 < 0 Then
         ChercherIntersection = 2     'intersection hors des segments
         XI = XA + t1 * (XB - XA)
         YI = YA + t1 * (YB - YA)
         Exit Function
      Else
         If t1 = 0 Or t1 = 1 Then
            If t1 = 0 Then
               ChercherIntersection = 3  'intersection sur le point A
               XI = XA
               YI = YA
               Exit Function
            Else
               ChercherIntersection = 4  'intersection sur le point B
               XI = XB
               YI = YB
               Exit Function
            End If
         
         ElseIf t2 = 0 Or t2 = 1 Then
            If t2 = 0 Then
               ChercherIntersection = 5  'intersection sur le point C
               XI = XC
               YI = YC
               Exit Function
            Else
               ChercherIntersection = 6  'intersection sur le point D
               XI = XD
               YI = YD
               Exit Function
            End If
         Else
            ChercherIntersection = 1  'intersection sur les segments
            XI = XA + t1 * (XB - XA)
            YI = YA + t1 * (YB - YA)
         End If
      End If
   End If
End Function


Public Function DecalerFil(ByVal Decalage As Single) As Sequ
   Dim i As Long, j As Long, k As Long, m As Long
   Dim XN As Single, YN As Single, XA As Single, YA As Single, XB As Single, YB As Single
   Dim XI As Single, YI As Single
   Dim PointDecale() As PointProfil
   Dim SelDecalee() As PointProfil
   Dim TypeIntersection As Long
   Dim NumeroDebut As Long, NumeroFin As Long
   Dim Delta1 As Single, Delta2 As Single, Delta3 As Single, Delta4 As Single
   
   Erase DecalerFil.Point
   
   With SequDecoupe
      i = 1 'le numéro du segment
      ReDim PointDecale(1 To i)
      Do       'boucle de décalage, on stocke les segments dans "PointDecale"
         Delta1 = Abs(.Point(i + 1).x - .Point(i).x)
         Delta2 = Abs(.Point(i + 1).y - .Point(i).y)
         If Delta1 < 0.01 And Delta2 < 0.01 Then
            MsgBox Message(Corps, 1), vbCritical, Message(Titre, 1)  'deux points consécutifs confondus, décalage impossible
            Exit Function
         End If
         'vecteur normal unitaire
         XN = (.Point(i + 1).y - .Point(i).y) / Sqr((.Point(i + 1).y - .Point(i).y) ^ 2 + (.Point(i).x - .Point(i + 1).x) ^ 2)
         YN = (.Point(i).x - .Point(i + 1).x) / Sqr((.Point(i + 1).y - .Point(i).y) ^ 2 + (.Point(i).x - .Point(i + 1).x) ^ 2)
         ReDim Preserve PointDecale(1 To 2 * i)    'on stocke les deux extrémités de chaque segment décalé
         PointDecale(2 * i - 1).x = .Point(i).x + XN * Decalage
         PointDecale(2 * i - 1).y = .Point(i).y + YN * Decalage
         PointDecale(2 * i).x = .Point(i + 1).x + XN * Decalage
         PointDecale(2 * i).y = .Point(i + 1).y + YN * Decalage
         If i = .NbPoints - 1 Then
            Exit Do
         End If
         i = i + 1
      Loop
   End With
   With DecalerFil
      .NbPoints = SequDecoupe.NbPoints
      ReDim .Point(1 To .NbPoints)
      .Point(1) = PointDecale(1)
      .Point(.NbPoints) = PointDecale(UBound(PointDecale))
      For k = 1 To UBound(PointDecale) - 3 Step 2
         TypeIntersection = ChercherIntersection(PointDecale(k).x, PointDecale(k).y, PointDecale(k + 1).x, PointDecale(k + 1).y, _
                        PointDecale(k + 2).x, PointDecale(k + 2).y, PointDecale(k + 3).x, PointDecale(k + 3).y, XI, YI)
         If TypeIntersection = 0 Then  'segments colinéaires
            .Point(((k + 1) / 2) + 1).x = PointDecale(k + 1).x
            .Point(((k + 1) / 2) + 1).y = PointDecale(k + 1).y
         Else     'sinon, intersection
            .Point(((k + 1) / 2) + 1).x = XI
            .Point(((k + 1) / 2) + 1).y = YI
         End If
      Next k
   End With
   Erase PointDecale
   
End Function

' forcer la position du pointeur de la souris
Public Function SetCursorPosition(Window As Object, xPos As Long, yPos As Long) As Boolean
    Dim x As Long, y As Long
    Dim lRet As Long
    Dim lHandle As Long
    Dim typPoint As POINTAPI
 
    On Error GoTo ErrorHandler
    lHandle = Window.hwnd
    With Screen
        x = CLng(xPos / .TwipsPerPixelX)
        y = CLng(yPos / .TwipsPerPixelY)
    End With
    typPoint.x = x
    typPoint.y = y
    lRet = ClientToScreen(lHandle, typPoint)
    lRet = SetCursorPos(typPoint.x, typPoint.y)
    SetCursorPosition = (lRet <> 0)
    Exit Function
 
ErrorHandler:
    SetCursorPosition = False
    Exit Function
End Function

'************************************
'*** Chargement des fichiers .CPX ***
'************************************
Public Sub LireCPX(ByVal NomDuFichier As String)
   Dim i As Long
   Dim t() As String
   Dim AlignerAGauche As Single, AlignerEnBas As Single, DecalageVersLaDroite As Single

   ReDim Sequ(1 To 2)   '2 séquences dans un .fc
   ReDim SequTrace(1 To 2)   'pour vider la mémoire

   With Sequ(1)
      .NbPoints = Val(LitFichierTypeIni("Emplanture", "NombrePoints", NomDuFichier))
      ReDim .Point(1 To .NbPoints)
      For i = 1 To .NbPoints
         t() = Split(LitFichierTypeIni("Emplanture", Str(i), NomDuFichier), ":")
         .Point(i).x = MyVal(t(0))
         .Point(i).y = MyVal(t(1))
      Next i
   End With
   With Sequ(2)
      .NbPoints = Val(LitFichierTypeIni("Saumon", "NombrePoints", NomDuFichier))
      ReDim .Point(1 To .NbPoints)
      For i = 1 To .NbPoints
         t() = Split(LitFichierTypeIni("Saumon", Str(i), NomDuFichier), ":")
         .Point(i).x = MyVal(t(0))
         .Point(i).y = MyVal(t(1))
      Next i
   End With
   Erase t  'vider la mémoire
   
   NbSequ = 2
   
   ReDim SequTrace(1).Point(1 To Sequ(1).NbPoints)
   ReDim SequTrace(2).Point(1 To Sequ(2).NbPoints)
   
   For i = 1 To NbSequ
      Sequ(i).Etat = 0 'initialisation de l'état
      SequTrace(i).Etat = 0 'initialisation de l'état
      Call MaxiMiniSequ(Sequ(i))
   Next i
   'Les deux séquences sont superposées, il faut décaler la deuxième
   AlignerAGauche = Sequ(2).Xmin - Sequ(1).Xmin
   AlignerEnBas = Sequ(2).Ymin - Sequ(1).Ymin
   DecalageVersLaDroite = 1.1 * (Sequ(1).Xmax - Sequ(1).Xmin)
   For i = 1 To Sequ(2).NbPoints
      With Sequ(2).Point(i)
         .x = .x - AlignerAGauche + DecalageVersLaDroite
         .y = .y - AlignerEnBas
      End With
   Next i
   Call MaxiMiniSequ(Sequ(2))
   
   NumSequSel = 0
   
End Sub

'************************************
'*** Chargement des fichiers .FC ***
'************************************
Public Sub LireProfilsFC(ByVal NomDuFichier As String)  'MonFichier As String)
   'variables issues du code de RPFC
   Dim monEnreg As String
   Dim i As Long, ii As Long, j As Long, k As Long
   Dim EouS As Integer
   Dim strX As String
   Dim strY As String
   Dim monCode As String
   Dim Temp As String
   Dim DecalageVersLaDroite As Single, AlignerAGauche As Single, AlignerEnBas As Single
       
   ReDim Sequ(1 To 2)   '2 séquences dans un .fc
   ReDim SequTrace(1 To 2)   'pour vider la mémoire
   ReDim Sequ(1).Point(1 To 1)
   ReDim Sequ(2).Point(1 To 1)
     
   ' Utilisation des éléments nécessaires du code de RP-FC
   Open NomDuFichier For Input As #1
   Input #1, monEnreg
   While Not EOF(1)
      'On cherche le "="
      i = 1
      While Mid$(monEnreg, i, 1) <> "="
         i = i + 1
      Wend
      If Left$(monEnreg, 11) = "Commentaire" Then
         monCode = "Commentaire"
      Else
         EouS = IIf(Left$(monEnreg, 1) = "E", 1, 2)   'si la première lettre est E, EouS vaut 1, sinon 2
         monCode = Mid$(monEnreg, 2, i - 2)
      End If
      Temp = Replace(Right$(monEnreg, Len(monEnreg) - i), ".", Format$(0#, "#.#"))
      If IsNumeric(monCode) Then  'on ne s'occupe que des profils, on zappe tout le reste
         ' Traitement des coordonnées qui sont de la forme E000001=x;y
         strX = Right$(monEnreg, Len(monEnreg) - i)
         j = 1
         While Mid$(strX, j, 1) <> ";"
             j = j + 1
         Wend
         strY = Right$(monEnreg, Len(strX) - j)
         strX = Left$(strX, j - 1)
         If EouS = 1 Then 'on est à l'emplanture
            If UBound(Sequ(1).Point) = 1 Then
               Sequ(1).Point(1).x = CSng(Replace(strX, ".", Format(0#, "#.#")))
               Sequ(1).Point(1).y = CSng(Replace(strY, ".", Format(0#, "#.#")))
               ReDim Preserve Sequ(1).Point(1 To 2) 'uniquement pour passer la condition du If
               k = 1
            Else
               k = k + 1
               ReDim Preserve Sequ(1).Point(1 To k)
               Sequ(1).Point(k).x = CSng(Replace(strX, ".", Format(0#, "#.#")))
               Sequ(1).Point(k).y = CSng(Replace(strY, ".", Format(0#, "#.#")))
            End If
         ElseIf EouS = 2 Then 'on est au saumon
            If UBound(Sequ(2).Point) = 1 Then
               Sequ(2).Point(1).x = CSng(Replace(strX, ".", Format(0#, "#.#")))
               Sequ(2).Point(1).y = CSng(Replace(strY, ".", Format(0#, "#.#")))
               ReDim Preserve Sequ(2).Point(1 To 2) 'uniquement pour passer la condition du If
               k = 1
            Else
               k = k + 1
               ReDim Preserve Sequ(2).Point(1 To k)
               Sequ(2).Point(k).x = CSng(Replace(strX, ".", Format(0#, "#.#")))
               Sequ(2).Point(k).y = CSng(Replace(strY, ".", Format(0#, "#.#")))
            End If
         End If
      End If
      Do
         Input #1, monEnreg
      Loop While Len(monEnreg) = 0
   Wend
   
Fermeture:
   Close #1
   
   NbSequ = UBound(Sequ)
   Sequ(1).NbPoints = UBound(Sequ(1).Point)
   Sequ(2).NbPoints = UBound(Sequ(2).Point)
   
   ReDim SequTrace(1).Point(1 To Sequ(1).NbPoints)
   ReDim SequTrace(2).Point(1 To Sequ(2).NbPoints)
   
   For i = 1 To NbSequ
      Sequ(i).Etat = 0 'initialisation de l'état
      SequTrace(i).Etat = 0 'initialisation de l'état
      Call MaxiMiniSequ(Sequ(i))
   Next i
   'Les deux séquences sont superposées, il faut décaler la deuxième
   AlignerAGauche = Sequ(2).Xmin - Sequ(1).Xmin
   AlignerEnBas = Sequ(2).Ymin - Sequ(1).Ymin
   DecalageVersLaDroite = 1.1 * (Sequ(1).Xmax - Sequ(1).Xmin)
   For i = 1 To Sequ(2).NbPoints
      With Sequ(2).Point(i)
         .x = .x - AlignerAGauche + DecalageVersLaDroite
         .y = .y - AlignerEnBas
      End With
   Next i
   Call MaxiMiniSequ(Sequ(2))
   
   NumSequSel = 0

End Sub

Public Function DebutDUnNombre(ByVal strLigne As String) As Boolean
   'retourne true si le premier caractère peut être le début d'un nombre
   Select Case Mid(Trim(strLigne), 1, 1) 'on regarde le premier caractère après avoir enlevé les espaces
   Case "-", "+", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
      DebutDUnNombre = True
   Case Else
      DebutDUnNombre = False
   End Select
End Function

'********************************************************************
'**** Calcul de l'ANGLE d'un SEGMENT par rapport à l'HORIZONTALE ****
'********************************************************************
Public Function Angle_Segment(ByVal XA As Single, ByVal YA As Single, ByVal XB As Single, ByVal YB As Single) As Single
   Dim DeltaX As Single, DeltaY As Single
   
   DeltaX = XB - XA
   DeltaY = YB - YA
   If DeltaX = 0 Then   'On traite le cas particulier d'un segment horizontal
      If DeltaY < 0 Then
         Angle_Segment = -pi / 2
      Else
         Angle_Segment = pi / 2
      End If
   Else              'dans les autres cas
      Angle_Segment = Atn(DeltaY / DeltaX)
      If DeltaX < 0 Then
         If DeltaY < 0 Then
            Angle_Segment = Angle_Segment - pi
         Else
            Angle_Segment = Angle_Segment + pi
         End If
      End If
   End If
   If Angle_Segment < 0 Then
      Angle_Segment = Angle_Segment + 2 * pi
   End If
End Function
