Attribute VB_Name = "modDeplacements"
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

'****** TEST DES INTERRUPTEURS ******
Public Function EtatInterOrigine() As Boolean
   Call RAZByteComm  'on met zéro partout
   ByteIPL(1) = &H49 'demande de réception de la salve information
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL = 1 Then     'OK
      If (ByteIPL(8) And &H2) = 0 Then
         EtatInterOrigine = False   'les inters sont tous fermés
      Else
         EtatInterOrigine = True    'un des inters est ouvert
      End If
   End If
End Function

Public Function EtatInterFDC() As Boolean
   Call RAZByteComm  'on met zéro partout
   ByteIPL(1) = &H49 'demande de réception de la salve information
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL = 1 Then  'OK
      If (ByteIPL(8) And &H4) = 0 Then
         EtatInterFDC = False   'les inters sont tous fermés
      Else
         EtatInterFDC = True    'un des inters est ouvert
      End If
   End If
End Function

'****** ENVOI D'UN DEPLACEMENT UNIQUE ******
'à chaque fois il faut initialiser le buffer, donner la source,
'lancer le "buffer go", et faire le fin de découpe à la fin
Public Sub MouvementUniquePas(ByVal XG As Single, ByVal YG As Single, ByVal XD As Single, ByVal YD As Single, _
                           ByVal Vitesse As Single, ByRef NBRPulsesMouvementUnique As Long)

   Dim SalveAEnvoyer As SalveData
   Dim VitesseDebut As Single, VitesseMvt As Single, VitesseFin As Single
   
   'On suppose que la vitesse demandée est dans les spécifications de la machine
   If Vitesse > VMaxSansAcc Then
      VitesseDebut = VMaxSansAcc
      VitesseMvt = Vitesse
      VitesseFin = VitesseDebut
   Else
      VitesseDebut = Vitesse
      VitesseMvt = VitesseDebut
      VitesseFin = VitesseMvt
   End If
   
   DoEvents  'juste avant le mouvement, on regarde si on n'a pas cliqué
   If flagAppuiSTOP = True Then
      Exit Sub
   End If
      
   'Reset buffer + Data source: USB
   Call EnvoiBytes(&H42, &H0)
   Call EnvoiBytes(&H42, &H1)
   'Calcul de la salve Data
   SalveAEnvoyer = CalculSalveDataPas(VitesseDebut, VitesseMvt, VitesseFin, XG, YG, XD, YD)
   'Transfert de la salve de mouvement dans le tableau de communication
   'on récupère le nombre de pulses de la salve
   NBRPulsesMouvementUnique = TransfertSalve2ByteIPL(SalveAEnvoyer)
   'stockage de la demande de mouvement dans le buffer
   ErrIPL = IPL5X_Send(ByteIPL(), 0)
   'c'est un mouvement unique, il faut envoyer l'instruction
   'de fin de découpe dans le buffer (pour pouvoir lire l'état de l'interface je crois)
   Call EnvoiBytes(&H44, &H80)
   'Go buffer : on lance l'exécution des mouvements stockés dans le buffer ce qui met les moteurs on si l'auto on/off a été paramétré
   Call EnvoiBytes(&H42, &H80)
   Do
      DoEvents  ' on regarde si l'utilisateur n'a pas cliqué sur STOP
      If flagAppuiSTOP = True Or flagAppuiStopSansMsgBox = True Then
         Call EnvoiBytes(&H53)   'on envoie le stop pour arrêter le mouvement en cours
      End If
      'On envoi "Information" en boucle jusqu'à ce que le mouvement soit terminé
      Call EnvoiBytes(&H49)
      If (ByteIPL(12) And &H2) = &H2 Then   'step activity stopped = on est à l'arrêt
         Exit Sub          'La fonction appelante va décoder la cause du stop sur ByteIPL(12)
      End If
   Loop
End Sub

'Transfert d'une salve dans le tableau de communication
' la fonction renvoit le nombre de pulse de la salve
Public Function TransfertSalve2ByteIPL(ByRef Salve As SalveData) As Long
   Call RAZByteComm  'on met zéro partout
   With Salve
      ByteIPL(1) = &H44
      ByteIPL(2) = .CMD
      ByteIPL(3) = .NBRL
      ByteIPL(4) = .NBRM
      ByteIPL(5) = .NBRH
      ByteIPL(6) = .NBRU
      TransfertSalve2ByteIPL = ConcatOctets(.NBRL, .NBRM, .NBRH, .NBRU)
      ByteIPL(7) = .S1L
      ByteIPL(8) = .S1M
      ByteIPL(9) = .S1H
      ByteIPL(10) = .S1U
      ByteIPL(11) = .S2L
      ByteIPL(12) = .S2M
      ByteIPL(13) = .S2H
      ByteIPL(14) = .S2U
      ByteIPL(15) = .S3L
      ByteIPL(16) = .S3M
      ByteIPL(17) = .S3H
      ByteIPL(18) = .S3U
      ByteIPL(19) = .S4L
      ByteIPL(20) = .S4M
      ByteIPL(21) = .S4H
      ByteIPL(22) = .S4U
      ByteIPL(23) = .S5L
      ByteIPL(24) = .S5M
      ByteIPL(25) = .S5H
      ByteIPL(26) = .S5U
      ByteIPL(27) = .F_ACC
      ByteIPL(28) = .F_DEC
      ByteIPL(29) = .DECL
      ByteIPL(30) = .DECM
      ByteIPL(31) = .DECH
      ByteIPL(32) = .DECU
   End With

End Function

Public Function DecodeErrDeplacement(ByVal Erreur As Long) As String
'******* Décodage des erreurs de la fonction Deplacement *******
   Select Case Erreur
   Case -300
      DecodeErrDeplacement = "Erreur -300 : la vitesse demandée sort des spécifications de la machine."
   Case -301
      DecodeErrDeplacement = "Erreur -301 : problème lors de l'appel de CalculSalveData par MouvementUnique."
   Case -302
      DecodeErrDeplacement = "Erreur -302 : problème lors d'envoi d'une salve Data dans la fonction MouvementUnique."
   Case -303
      DecodeErrDeplacement = "Erreur -303 : problème lors d'une demande d'Information dans la fonction MouvementUnique."
   Case -304
      DecodeErrDeplacement = "Erreur -304 : problème lors d'une demande de fin de mouvement."
   Case Else
      DecodeErrDeplacement = "Erreur n°" & Str(Erreur) & " non répertoriée sur demande de déplacement."
   End Select
End Function


'******* SALVE DEPLACEMENTS IPL5X ********
'Il faut créer les tableaux des accélérations au chargement du soft (voir TableLIN dans le Form_load):
Public Function CalculSalveDataPas(ByVal VitesseAvAcc As Single, _
            ByVal VitesseAcc As Single, ByVal VitesseApAcc As Single, ByVal PasXG As Long, ByVal PasYG As Long, _
            ByVal PasXD As Long, ByVal PasYD As Long) As SalveData
            'les Frequence, Pente, MmParTour et PasParTour dont on a besoin sont définis en Public dans la form appelante
            'les valeurs PasXX peuvent être négatives
   Dim i As Integer
   Dim IndexFrequence As Integer  'pour utiliser le bon tableau
   
   Dim PENTE As Long
   Dim NbrDec As Long
   Dim NbrPulse As Long
   Dim DistanceXG As Double
   Dim DistanceYG As Double
   Dim DistanceXD As Double
   Dim DistanceYD As Double
   Dim NbrPasXG As Long 'Attention, c'est PasX * 2 (voir Dev_guide)
   Dim NbrPasYG As Long
   Dim NbrPasXD As Long
   Dim NbrPasYD As Long
   Dim temps_s As Double
   Dim K_Acc As Double, K_Dec As Double
   Dim K_Acc_Comp As Double, K_Dec_Comp As Double  'K *600000/freq pour comparaison dans la table des acc
   Dim Index_Acc As Integer, Index_Dec As Integer, Index_Min As Integer, Index_Point As Integer
   Dim Periode_Acc As Integer, Periode_Dec As Integer
   Dim Nbr_Acc As Double, Nbr_Dec As Double
   Dim Nbr_Dec_Salve As Long
   Dim Temps_Acc As Single, Temps_Dec As Single, Temps_Plateau As Single
   
   If Frequence <> 10000 And Frequence <> 20000 And Frequence <> 30000 And Frequence <> 40000 And Frequence <> 50000 Then
      CalculSalveDataPas.CodeErreur = 100 'Fréquence fausse
      Exit Function
   End If
   If PenteAcceleration < 0 Or PenteAcceleration > 15 Then
      CalculSalveDataPas.CodeErreur = 101 'Pente en dehors des limites
      Exit Function
   End If
   If VitesseAvAcc <= 0 Or VitesseAcc <= 0 Or VitesseApAcc <= 0 Then
      CalculSalveDataPas.CodeErreur = 102 'L'une des vitesses est négative ou nulle
      Exit Function
   End If
   
   On Error GoTo Erreur  'si plantage sur erreur non répetoriée
   
   With CalculSalveDataPas 'Initialisation de la salve
      .CMD = 0
      .NBRL = 0
      .NBRM = 0
      .NBRH = 0
      .NBRU = 0
      .S1L = 0
      .S1M = 0
      .S1H = 0
      .S1U = 0
      .S2L = 0
      .S2M = 0
      .S2H = 0
      .S2U = 0
      .S3L = 0
      .S3M = 0
      .S3H = 0
      .S3U = 0
      .S4L = 0
      .S4M = 0
      .S4H = 0
      .S4U = 0
      .S5L = 0
      .S5M = 0
      .S5H = 0
      .S5U = 0
      .F_ACC = 0
      .F_DEC = 0
      .DECL = 0
      .DECM = 0
      .DECH = 0
      .DECU = 0
      .TempsTotal = 0
      .CodeErreur = 0
   End With
   
   ReDim NbrPas(0 To 3)
   PENTE = 2 ^ (PenteAcceleration + 8 + 1)
   IndexFrequence = Frequence / 10000
   DistanceXG = PasXG * MmParTourXG / PasParTourXG
   DistanceXD = PasXD * MmParTourXD / PasParTourXD
   DistanceYG = PasYG * MmParTourYG / PasParTourYG
   DistanceYD = PasYD * MmParTourYD / PasParTourYD
   With CalculSalveDataPas
      .CMD = 0
      If PasXG >= 0 Then .CMD = .CMD + 1    'inversion axe 1
      If PasYG >= 0 Then .CMD = .CMD + 2    'inversion axe 2
      If PasXD >= 0 Then .CMD = .CMD + 4    'inversion axe 3
      If PasYD >= 0 Then .CMD = .CMD + 8    'inversion axe 4
      If VitesseAcc / VitesseAvAcc > 1 Or VitesseAcc / VitesseApAcc > 1 Then
         .CMD = .CMD + 32    'bit 5 : présence de l'accélération (oui/non)
      End If
      'Nbr pulse :
      If Sqr(DistanceXG ^ 2 + DistanceYG ^ 2) / VitesseAcc > Sqr(DistanceXD ^ 2 + DistanceYD ^ 2) / VitesseAcc Then
         temps_s = Sqr(DistanceXG ^ 2 + DistanceYG ^ 2) / VitesseAcc
      Else
         temps_s = Sqr(DistanceXD ^ 2 + DistanceYD ^ 2) / VitesseAcc
      End If
      NbrPulse = CLng(roundUp(temps_s * Frequence))
      .NBRL = LPart(NbrPulse)
      .NBRM = MPart(NbrPulse)
      .NBRH = HPart(NbrPulse)
      .NBRU = UPart(NbrPulse)
      i = 1
      'Step x :
      NbrPasXG = 2 * Abs(PasXG)  'ATTENTION AU COEFF 2 (voir Dev_Guide)
      NbrPasYG = 2 * Abs(PasYG)
      NbrPasXD = 2 * Abs(PasXD)
      NbrPasYD = 2 * Abs(PasYD)
      .S1L = LPart(NbrPasXG)
      .S1M = MPart(NbrPasXG)
      .S1H = HPart(NbrPasXG)
      .S1U = UPart(NbrPasXG)
      .S2L = LPart(NbrPasYG)
      .S2M = MPart(NbrPasYG)
      .S2H = HPart(NbrPasYG)
      .S2U = UPart(NbrPasYG)
      .S3L = LPart(NbrPasXD)
      .S3M = MPart(NbrPasXD)
      .S3H = HPart(NbrPasXD)
      .S3U = UPart(NbrPasXD)
      .S4L = LPart(NbrPasYD)
      .S4M = MPart(NbrPasYD)
      .S4H = HPart(NbrPasYD)
      .S4U = UPart(NbrPasYD)
      .S5L = 0
      .S5M = 0
      .S5H = 0
      .S5U = 0
      'Accélérations :
      If VitesseAcc / VitesseAvAcc > 1 Or VitesseAcc / VitesseApAcc > 1 Then
         K_Acc = VitesseAcc / VitesseAvAcc
         K_Acc_Comp = K_Acc * 6000000 / Frequence
         Index_Acc = 0
         For i = 1 To 255
            If Abs(TableLIN(IndexFrequence, i) - K_Acc_Comp) < Abs(TableLIN(IndexFrequence, Index_Acc) - K_Acc_Comp) Then
               Index_Acc = i
            End If
         Next i
         K_Dec = VitesseAcc / VitesseApAcc
         K_Dec_Comp = K_Dec * 6000000 / Frequence
         Index_Dec = 0
         For i = 1 To 255
            If Abs(TableLIN(IndexFrequence, i) - K_Dec_Comp) < Abs(TableLIN(IndexFrequence, Index_Dec) - K_Dec_Comp) Then
               Index_Dec = i
            End If
         Next i
         If K_Acc > 1 Or K_Dec > 1 Then
            .F_ACC = Index_Acc
            .F_DEC = Index_Dec
         Else
            .F_ACC = 0
            .F_DEC = 0
         End If
         Periode_Acc = TableLIN(IndexFrequence, Index_Acc)
         Periode_Dec = TableLIN(IndexFrequence, Index_Dec)
         Nbr_Acc = 0
         Nbr_Dec = 0
         Index_Min = Index_Acc
         If Index_Dec < Index_Min Then Index_Min = Index_Dec
         
         For i = Index_Min To 254
            If i >= Index_Acc Then Nbr_Acc = Nbr_Acc + PENTE / (2 * TableLIN(IndexFrequence, i))
            If i >= Index_Dec Then Nbr_Dec = Nbr_Dec + PENTE / (2 * TableLIN(IndexFrequence, i))
            If Nbr_Acc + Nbr_Dec > NbrPulse Then Exit For
         Next i
         Index_Point = i
         Nbr_Dec_Salve = NbrPulse - Int(Nbr_Dec)  'à mettre dans la dernière case de la salve
         .DECL = LPart(Nbr_Dec_Salve)
         .DECM = MPart(Nbr_Dec_Salve)
         .DECH = HPart(Nbr_Dec_Salve)
         .DECU = UPart(Nbr_Dec_Salve)
      End If
      Temps_Acc = (Index_Point - Index_Acc) * PENTE / 12000000
      Temps_Dec = (Index_Point - Index_Dec) * PENTE / 12000000
      Temps_Plateau = (NbrPulse - Nbr_Acc - Nbr_Dec) / Frequence
      If Temps_Plateau < 0 Then Temps_Plateau = 0
      
      .TempsTotal = Temps_Acc + Temps_Plateau + Temps_Dec
   End With
   CalculSalveDataPas.CodeErreur = 1 'tout s'est bien passé
   Exit Function
   
Erreur:
   CalculSalveDataPas.CodeErreur = -400 'Plantage du calcul de la salve
   
End Function

Public Function DecodeErrSalveData(ByVal Erreur As Long) As String
'******* Décodage des erreurs de la fonction Deplacement *******
   Select Case Erreur
   Case -100
      DecodeErrSalveData = "Erreur -400 : plantage lors du calcul d'une salve Data."
   Case Else
      DecodeErrSalveData = "Erreur non répertoriée sur calcul d'une salve Data."
   End Select
End Function

Public Sub TraduireEnSalve(SequenceATraduire As SequPas, ByRef TableauSalves() As SalveData)
'******* Traduction des segments de la séquence de découpe en vecteurs de déplacement (salve data)***********
   Dim i As Long
   Dim VitesseDebut As Single, VitesseMvt As Single, VitesseFin As Single
   Dim SequencePas As SequPas
   Dim PasXGSegment As Long, PasYGSegment As Long
   Dim PasXDSegment As Long, PasYDSegment As Long
   
   With SequenceATraduire
      ReDim TableauSalves(0 To .NbPoints - 1)   'la salve zéro servira pour la tempo, puis il y a une salve de moins que le nombre de points
      For i = 1 To .NbPoints - 1
         VitesseDebut = .PointPas(i).Vitesse
         VitesseFin = .PointPas(i + 1).Vitesse
         If .PointPas(i).Acceleration = True Then
            VitesseMvt = VMaxAvecAcc
         Else
            VitesseMvt = .PointPas(i).Vitesse
         End If
         PasXGSegment = .PointPas(i + 1).XPas - .PointPas(i).XPas
         PasYGSegment = .PointPas(i + 1).YPas - .PointPas(i).YPas
         PasXDSegment = PasXGSegment
         PasYDSegment = PasYGSegment
         TableauSalves(i) = CalculSalveDataPas(VitesseDebut, VitesseMvt, VitesseFin, PasXGSegment, PasYGSegment, PasXDSegment, PasYDSegment)
      Next i
   End With
   'à la fin du tableau, on ajoute l'ordre de fin de mouvement (44 80)
   ReDim Preserve TableauSalves(0 To SequenceATraduire.NbPoints)
   TableauSalves(SequenceATraduire.NbPoints).CMD = &H80  'tout le reste est à zéro
End Sub

Public Function SalveAttenteChauffe(ByVal Duree As Integer, ByVal ChauffeFil As Byte) As SalveData
'******* Création d'une salve à mouvement nul pour mise en route de la chauffe ; Durée en s ; Chauffe en %
   Dim NbrePulsesAttente As Long
   
   NbrePulsesAttente = Duree * Frequence
   SalveAttenteChauffe = CalculSalveDataPas(VitesseDecoupe, VitesseDecoupe, VitesseDecoupe, 0, 0, 0, 0)  'on crée une salve à petite vitesse et mouvement nul
   'On modifie la salve pour y activer la chauffe et augmenter sa durée (pour le moment =0 ce qui est invalide)
   With SalveAttenteChauffe
      .CMD = .CMD + 64
      .NBRL = LPart(NbrePulsesAttente)
      .NBRM = MPart(NbrePulsesAttente)
      .NBRH = HPart(NbrePulsesAttente)
      .NBRU = UPart(NbrePulsesAttente)
      .F_ACC = Round(ChauffeFil * 2.55, 0)  'le mouvement n'est pas accéléré
   End With
End Function

'********** Procédure de lancement d'un mouvement de découpe ************
Public Function CalculSalvesDecoupe(SequMouvDecoupe As Sequ, ByVal TempoChauffe As Integer, ByVal ValeurChauffe As Integer) As Single  'renvoie la durée de la découpe si OK, -1 si NOK
   Dim i As Long
   
   On Error GoTo ErreurSalveDecoupe 'si la fonction plante, on renvoie -1
   'on transfère tout dans un tableau identique, mais en pas
   With SequMouvDecoupe
      ReDim SequMouvementPas.PointPas(0 To .NbPoints)
      SequMouvementPas.PointPas(0).XPas = 0  'le point 0 est important pour calculer les pas du segment
      SequMouvementPas.PointPas(0).YPas = Round(CourseY * PasParTourYG / MmParTourYG, 0)
      For i = 1 To .NbPoints
         SequMouvementPas.PointPas(i).XPas = Round(.Point(i).x * PasParTourXG / MmParTourXG, 0)
         SequMouvementPas.PointPas(i).YPas = Round(.Point(i).y * PasParTourYG / MmParTourYG, 0)
         SequMouvementPas.PointPas(i).Vitesse = .Point(i).Vitesse
         SequMouvementPas.PointPas(i).Acceleration = .Point(i).Acceleration
      Next i
      SequMouvementPas.NbPoints = .NbPoints
   End With

   ReDim SalveDecoupe(0 To 0)   'initialisation du tableau des salves, la salve n°0 servira pour la tempo de chauffe
   Call TraduireEnSalve(SequMouvementPas, SalveDecoupe)  'SalveDecoupe est passé byref et modifié par la procédure
   'La procédure TraduireEnSalve ajoute dans le tableau la salve fin 0x44 0x80
   SalveDecoupe(0) = SalveAttenteChauffe(TempoChauffe, ValeurChauffe) 'on intègre la tempo de chauffe et sa mise en route (attention, 0secondes est impossible)
   CalculSalvesDecoupe = 0
   For i = 1 To SequMouvDecoupe.NbPoints - 1
      CalculSalvesDecoupe = CalculSalvesDecoupe + SalveDecoupe(i).TempsTotal
   Next i
   Exit Function
ErreurSalveDecoupe:
   CalculSalvesDecoupe = -1
End Function

'******** calcul de la temporisation de mise en température du fil en fonction de la chauffe *****
Public Function CalculTempoChauffe(ByVal ChauffeFil As Single) As Integer
   Select Case TypeMachine
   Case "MiniCut2d_v1.2"  'sur la nouvelle version de l'interface, c'est plus rapide !!
      CalculTempoChauffe = 2
   Case Else
      If ChauffeFil > 0 And ChauffeFil <= 45 Then
         CalculTempoChauffe = 20
      ElseIf ChauffeFil > 45 And ChauffeFil <= 60 Then
         CalculTempoChauffe = 25
      ElseIf ChauffeFil > 60 And ChauffeFil <= 70 Then
         CalculTempoChauffe = 28
      ElseIf ChauffeFil > 70 And ChauffeFil <= 80 Then
         CalculTempoChauffe = 30
      Else
         CalculTempoChauffe = 35
      End If
   End Select
End Function



