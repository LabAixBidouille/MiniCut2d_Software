Attribute VB_Name = "modLangueFrancais"
'Copyright Renaud ILTIS, 2014
'
'renaudiltis@ yahoo.fr
'
'Ce logiciel est un programme informatique servant � piloter une machine de d�coupe par fil chaud.
'Ce logiciel est r�gi par la licence CeCILL soumise au droit fran�ais et respectant les principes de diffusion des logiciels libres. Vous pouvez
'utiliser, modifier et/ou redistribuer ce programme sous les conditionsde la licence CeCILL telle que diffus�e par le CEA, le CNRS et l'INRIA
'sur le site "http://www.cecill.info".
'En contrepartie de l'accessibilit� au code source et des droits de copie, de modification et de redistribution accord�s par cette licence, il n'est
'offert aux utilisateurs qu'une garantie limit�e. Pour les m�mes raisons, seule une responsabilit� restreinte p�se sur l'auteur du programme,  le
'titulaire des droits patrimoniaux et les conc�dants successifs.
'A cet �gard  l'attention de l'utilisateur est attir�e sur les risques associ�s au chargement, � l'utilisation, � la modification et/ou au
'd�veloppement et � la reproduction du logiciel par l'utilisateur �tant donn� sa sp�cificit� de logiciel libre, qui peut le rendre complexe �
'manipuler et qui le r�serve donc � des d�veloppeurs et des professionnels avertis poss�dant des connaissances informatiques approfondies. Les
'utilisateurs sont donc invit�s � charger et tester l'ad�quation du logiciel � leurs besoins dans des conditions permettant d'assurer la
's�curit� de leurs syst�mes et ou de leurs donn�es et, plus g�n�ralement, � l'utiliser et l'exploiter dans les m�mes conditions de s�curit�.
'Le fait que vous puissiez acc�der � cet en-t�te signifie que vous avez pris connaissance de la licence CeCILL, et que vous en avez accept� les
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

Public Sub LangueFrancais(LangueAUtiliser As String)
   '**** traduction du SplashScreen
   With frmSplashScreen
      .Caption = "MiniCut2d Software - Bienvenue !"
      .lblCliquez = "Cliquez sur le type de machine que vous utilisez"
      .cmdPasDeMachine.Caption = "Si vous n'avez pas de machine, cliquez ici"
   End With
   '**** traduction du A propos ****
   With frmAboutAndSettings
      .Caption = "Param�tres"
      .lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      .lblTitle.Caption = TypeMachine & " Software"
      .lblParametres.Caption = "Course totale X : " & Format(CourseX, "#####") & "mm - Course totale Y : " & Format(CourseY, "#####") & "mm." & _
               vbCrLf & "D�calage inter/origine X : " & Format(MmToOriXG, "#0.0##") & "mm, soit " & Format(NbrPasToOriXG, "#######") & " micropas." & _
               vbCrLf & "D�calage inter/origine YG : " & Format(MmToOriYG, "#0.0##") & "mm, soit " & Format(NbrPasToOriYG, "#######") & " micropas." & _
               vbCrLf & "D�calage inter/origine YD : " & Format(MmToOriYD, "#0.0##") & "mm, soit " & Format(NbrPasToOriYD, "#######") & " micropas."
      .lblTraduction.Caption = "Traduction :" & vbCrLf & vbCrLf & _
                              "Anglais : aiR-C�/Hugh Potter" & vbCrLf & _
                              "Allemand : Charles Wittmer" & vbCrLf & _
                              "Espagnol : Enrique Iglesias"
      .frmParametres.Caption = "Param�tres"
      .frmAPropos.Caption = "A propos"
      .frmModeExpert.Caption = "Fonctionnement Normal / Expert"
      .chkActiverLeChangementDeMode.Caption = "Activer le changement de mode"
      .optNormalExpert(0).Caption = "Mode Normal"
      .optNormalExpert(1).Caption = "Mode Expert"
   End With
   '**** traduction de la form "D�coupe Inactive" *****
   With frmDecoupeInactive
      .lblDecoupeInactive.Caption = "L'acc�s � cette partie du logiciel est impossible" & vbCrLf & "car l'interface USB n'est pas d�tect�e."
   End With
   '**** traduction de la form "Param�tres machine"
   With frmParametres
      .Caption = "Param�tres de la machine"
   End With
   '**** traduction de la form de vectorisation ****
   With frmImpConv
      .frmImporterImage.Caption = "Image"
      .frmRecadrer.Caption = "Recadrer"
      .frmApercu.Caption = "Noir et Blanc"
      .frmLisserTransferer.Caption = "Vectoriser"
      .cmdImporterImage(0).ToolTipText = "Importer..."
      .cmdImporterImage(1).ToolTipText = "Coller"
      .cmdRecadrerLaSelection.ToolTipText = "Recadrer"
      .optApercu(0).ToolTipText = "Convertir en noir et blanc"
      .optApercu(1).ToolTipText = "Annuler"
      .optInterieur(0).ToolTipText = "Supprimer les trac�s int�rieurs"
      .optInterieur(1).ToolTipText = "Conserver les trac�s int�rieurs"
      .cmdContoursLissageTransfert.ToolTipText = "Vectoriser"
      .chkVoirPointsVecto.ToolTipText = "Affichage des points"
      .cmdSauverVecto.ToolTipText = "Sauver en .DXF"
      .cmdQuitterImpConv.ToolTipText = "Valider"
   End With
   '**** traduction de la form principale ****
   With frmMiniCut2d
      .cmdLangue.Picture = frmImages.imgDrapeauFrancais.Picture
      ' *** traductions des contr�les de l'interface ***
      .SSTab1.TabCaption(0) = "Cr�ation"
      .SSTab1.TabCaption(1) = "D�coupe"
      .lblBoiteOutils.Caption = " Boite � outils "
      .lblAlignerAjuster.Caption = " Aligner et ajuster "
      .lblDimensionsBloc.Caption = " Dimensions du bloc "
      .lblTitreChauffe.Caption = " Chauffe "
      .lblTrajet.Caption = " Trajet "
      .lblDecoupe.Caption = " D�coupe "
      .lblFil.Caption = " Fil "
      .lblCadreDecoupe.Caption = " D�coupe "
      .lblStopReprise.Caption = " Stop / Reprise "
      .lblPiloterLeFil.Caption = " Piloter le fil "
      .frameEntreeBloc.Caption = "Entr�e"
      .frameSortieBloc.Caption = "Sortie"
      .frameSimulation.Caption = "Simulation"
      .frameManuel.Caption = "Piloter"
      .frameProcedures.Caption = "Positions"
      .frameDecalage.Caption = "D�caler le fil"
      .frameInformation.Caption = "Information"
      .frameAction.Caption = "Lancer la d�coupe"
      .frameAnnulationStop.Caption = "Annulation - Stop/Reprise"
      .frameChauffeEnCoursDecoupe.Caption = "Chauffe"
      .frameInformationStop.Caption = "Information"
      .frameModifierLaChauffe.Caption = "Modifier la chauffe"
      .frameRetourOrigine.Caption = "Retourner � l'origine"
      .frameTrajetRetour.Caption = "Trajet"
      .frameReprendre.Caption = "Terminer la d�coupe"
      .frameAnnulationReprise.Caption = "Annuler"
      .frameChauffeFil.Caption = "Faire chauffer"
      .frameInformationFil.Caption = "Information"
      .frameFilManuel.Caption = "D�placer"
      .frameAnnulationFil.Caption = "Annuler / Quitter"
      .cmdNouveauProjet.ToolTipText = "Nouveau projet"
      .cmdOuvrirFichierSequ.ToolTipText = "Ouvrir un projet"
      .cmdSauver(0).ToolTipText = "Enregistrer sous..."
      .cmdSauver(1).ToolTipText = "Enregistrer"
      .cmdLangue.ToolTipText = "Changer la langue"
      .cmdSettings.ToolTipText = "Param�tres"
      .cmdRafraichir.ToolTipText = "Actualiser"
      .cmdEffacerFichier.ToolTipText = "Supprimer (corbeille)"
      .cmdImporterProfil.ToolTipText = "Importer un fichier dans la biblioth�que"
      .cmdSimulation.ToolTipText = "Voir le d�placement du fil"
      .optOutils(4).ToolTipText = "Mesurer"
      .optOutils(5).ToolTipText = "Couper un trajet"
      .optOutils(2).ToolTipText = "Etirer"
      .optOutils(1).ToolTipText = "Tourner"
      .optOutils(0).ToolTipText = "D�placer"
      .optOutils(6).ToolTipText = "Modifier le point d'entr�e"
      .cmdUndo(0).ToolTipText = "D�faire"
      .cmdUndo(1).ToolTipText = "Refaire"
      .cmdPoubelle.ToolTipText = "Supprimer"
      .cmdInsererPoint.ToolTipText = "Ins�rer un point"
      .cmdDupliquer.ToolTipText = "Dupliquer"
      .cmdMiroir.ToolTipText = "Miroir"
      .cmdInverser.ToolTipText = "Inverser le sens"
      .cmdAligner(0).ToolTipText = "Aligner en bas"
      .cmdAligner(1).ToolTipText = "Aligner au milieu"
      .cmdAligner(2).ToolTipText = "Aligner en haut"
      .cmdAligner(3).ToolTipText = "Aligner � gauche"
      .cmdAligner(4).ToolTipText = "Aligner au milieu"
      .cmdAligner(5).ToolTipText = "Aligner � droite"
      .chkCentrer(0).ToolTipText = "Mettre � l'�chelle du bloc"
      .chkCentrer(1).ToolTipText = "Centrer horizontalement"
      .chkCentrer(2).ToolTipText = "Centrer verticalement"
      .chkVoirPoints.ToolTipText = "Afficher les points"
      .chkCouleurProfils.ToolTipText = "Alterner la couleur des profils"
      .cmdAgrandirRetrecir.ToolTipText = "Taille de la fen�tre"
      .chkZoomProjet.ToolTipText = "Zoomer sur le bloc"
      .pctZoomInfo.ToolTipText = "Zoom : clic droit + molette ou fl�ches haut et bas"
      .cmdGestionMatiere(0).ToolTipText = "Nouvelle mati�re"
      .cmdGestionMatiere(1).ToolTipText = "Remplacer la valeur de la chauffe"
      .cmdGestionMatiere(2).ToolTipText = "Supprimer la mati�re"
      .cmdDecouper.ToolTipText = "D�couper le projet"
      .cmdDeplacementsManuels.ToolTipText = "Piloter le fil en direct"
      .cmdPlierLePortique.ToolTipText = "Aller en position de rangement"
      .cmdRetourOrigine.ToolTipText = "Ramener le fil � l'origine"
      .optDecalage(1).ToolTipText = "-0.5mm"
      .optDecalage(2).ToolTipText = "0mm"
      .optDecalage(3).ToolTipText = "0.5mm"
      .cmdFaireOrigineAvantDecoupe.ToolTipText = "Lancer la d�coupe"
      .cmdAnnulerDecoupe.ToolTipText = "Annuler la d�coupe"
      .cmdSTOP.ToolTipText = "Arr�t d'urgence !"
      .optTrajetRetour(0).ToolTipText = "En diagonale"
      .optTrajetRetour(1).ToolTipText = "Par la gauche"
      .optTrajetRetour(2).ToolTipText = "Par le haut"
      .cmdLancerRetourApresStop.ToolTipText = "Lancer le retour"
      .cmdStopRetourApresReprise.ToolTipText = "Arr�t d'urgence !"
      .cmdRepriseDecoupe.ToolTipText = "Terminer la d�coupe"
      .cmdAnnulerReprise.ToolTipText = "Annuler tout"
      .optChauffe(0).ToolTipText = "Faire chauffer"
      .optChauffe(1).ToolTipText = "Arr�ter la chauffe"
      .optGoManuel(0).ToolTipText = "Lancer le mouvement"
      .optGoManuel(1).ToolTipText = "Arr�ter le mouvement"
      .cmdAnnulerFil.ToolTipText = "Quitter"
      .optHomeY.ToolTipText = "Origine verticale"
      .optHomeX.ToolTipText = "Origine horizontale"
      .optAnnulerHome.ToolTipText = "Stop"
   End With
   '**** les MessagBox ****
   ReDim Message(1 To 2, 1 To 1)
   'MessageBox n�1
   Message(Corps, 1) = "Deux points cons�cutifs sont confondus." & vbCrLf & "Impossible de d�finir un d�calage."
   Message(Titre, 1) = "Calcul impossible."
   'MessageBox n�2
   ReDim Preserve Message(Corps To Titre, 1 To 2)
   Message(Corps, 2) = "Le r�pertoire \Bibliotheque n'est pas pr�sent, il va �tre cr�� mais il sera vide : � vous de le remplir!"
   Message(Titre, 2) = "Initialisation de la biblioth�que"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 3)
   Message(Corps, 3) = "Le fichier ""MiniCut2d Software.ini"" contient des cl�s qui commencent par ""NbrPasToOri...""." & vbCrLf & _
                        "Ces cl�s ne sont plus valides et vont �tre remplac�es par les " & vbCrLf & _
                        "nouvelles cl�s de type ""MmToOri..."" pour que MiniCut2d Software puisse fonctionner correctement."
   Message(Titre, 3) = "Ancienne version du .ini"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 4)
   Message(Corps, 4) = "La valeur de chauffe m�moris�e avec une mati�re d�passe la valeur maximale possible sur la machine." _
                        & vbCrLf & "La chauffe sera brid�e � la valeur maximale."
   Message(Titre, 4) = "Chauffe trop �lev�e"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 5)
   Message(Corps, 5) = "Le fichier IPL5XComm.dll pr�sent sur votre ordinateur est trop vieux." & vbCrLf & _
                        "MiniCut2d Software ne peut pas fonctionner." & vbCrLf & _
                        "Veuillez t�l�charger la derni�re version et relancer."
   Message(Titre, 5) = "Communication avec la machine impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 6)
   Message(Corps, 6) = "L'initialisation de l'interface USB pose probl�me" & vbCrLf & _
                        "L'acc�s aux fonctions de d�coupe risque d'�tre impossible"
   Message(Titre, 6) = "Probl�me d'initialisation de l'interpolateur."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 7)
   Message(Corps, 7) = "Le fichier IPL5XCom.dll n'a pas �t� trouv�," & vbCrLf & _
                        "veuillez l'installer sur votre ordinateur, � la bonne place." & vbCrLf & _
                        "La d�coupe sera d�sactiv�e."
   Message(Titre, 7) = "Communication avec la machine impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 8)
   Message(Corps, 8) = "Il semblerait que la MiniCut2d ne soit pas en position de rangement." & vbCrLf & _
                        "Voulez-vous lancer la proc�dure de rangement du plateau et du fil avant de quitter le logiciel ?"
   Message(Titre, 8) = "Demande de fermeture de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 9)
   Message(Corps, 9) = "Coupez l'alimentation de la MiniCut2d avant de quitter le programme." & vbCrLf & _
   "V�rifiez �galement que votre projet est sauvegard�." & vbCrLf & vbCrLf & _
   "Confirmez-vous la fermeture de l'application ?"
   Message(Titre, 9) = "Demande de fermeture de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 10)
   Message(Corps, 10) = "Impossible de repr�senter ce fichier."
   Message(Titre, 10) = "Erreur de lecture de fichier"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 11)
   Message(Corps, 11) = "Voulez-vous effectuer une sauvegarde du projet en cours?"
   Message(Titre, 11) = "Nouveau projet"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 12)
   Message(Corps, 12) = "Extension de fichier non valide. MiniCut2d Software sait ouvrir les .mnc, .dxf, .dat, .plt, .eps (et les .txt avec les coordonn�es des points s�par�es par un double-point)."
   Message(Titre, 12) = "Ouverture impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 13)
   Message(Corps, 13) = "Vous avez coup� sur le premier point d'un profil, impossible de couper � cet endroit."
   Message(Titre, 13) = "Op�ration impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 14)
   Message(Corps, 14) = "Vous avez coup� sur le dernier point d'un profil, impossible de couper � cet endroit."
   Message(Titre, 14) = "Op�ration impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 15)
   Message(Corps, 15) = "Veuillez entrer un nombre d�cimal positif ou n�gatif."
   Message(Titre, 15) = "Erreur de saisie"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 16)
   Message(Corps, 16) = "L'op�ration demand�e est impossible, une dimension est �gale � z�ro."
   Message(Titre, 16) = "Calcul impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 17)
   Message(Corps, 17) = "La zone utile de la machine est d�pass�e, le projet est tronqu� automatiquement."
   Message(Titre, 17) = "D�passement des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 18)
   Message(Corps, 18) = "Le nom de cette mati�re existe d�j�, �craser les valeurs?"
   Message(Titre, 18) = "Modification d'une mati�re"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 19)
   Message(Corps, 19) = "Cette ligne ne peut pas �tre effac�e."
   Message(Titre, 19) = "Op�ration impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 20)
   Message(Corps, 20) = "Valider l'effacement de ce mat�riau?"
   Message(Titre, 20) = "Supprimer un mat�riau"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 21)
   Message(Corps, 21) = "Il y a un probl�me : la valeur de chauffe pour la mati�re choisie n'est pas dans les limites pr�vues."
   Message(Titre, 21) = "Valeur incorrecte"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 22)
   Message(Corps, 22) = "Initialisation de la " & TypeMachine
   Message(Titre, 22) = "Interface USB d�tect�e"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 23)
   Message(Corps, 23) = "Ex�cuter la proc�dure de retour � la position de repos ?"
   Message(Titre, 23) = "Validation de s�curit�"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 24)
   Message(Corps, 24) = "Le fil est � la position de repos."
   Message(Titre, 24) = "Retour effectu�"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 25)
   Message(Corps, 25) = "L'op�ration a �t� annul�e."
   Message(Titre, 25) = "Arr�t d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 26)
   Message(Corps, 26) = "La boucle des interrupteurs d'origine est ouverte." & vbCrLf & _
                        "Ce n'est pas normal." & vbCrLf & "Il faut d�gager les interrupteurs en tournant les moteurs � la main" & _
                        vbCrLf & "(couper l'alimentation s'ils r�sistent)."
   Message(Titre, 26) = "Sortie des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 27)
   Message(Corps, 27) = "La boucle des interrupteurs de fin de course est ouverte." & vbCrLf & _
                        "Ce n'est pas normal." & vbCrLf & "Il faut d�gager les interrupteurs en tournant les moteurs � la main" & _
                        vbCrLf & "(couper l'alimentation s'ils r�sistent)."
   Message(Titre, 27) = "Sortie des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 28)
   Message(Corps, 28) = "D�placement en position de rangement?"
   Message(Titre, 28) = "Validation de s�curit�"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 29)
   Message(Corps, 29) = "Position de pliage atteinte."
   Message(Titre, 29) = "Mouvement effectu�"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 30)
   Message(Corps, 30) = "Les interrupteurs de la table courante ne sont pas actifs."
   Message(Titre, 30) = "Impossible de faire la proc�dure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 31)
   Message(Corps, 31) = "Un interrupteur de fin de course est ouvert, la proc�dure ne peut pas commencer." & vbCrLf & _
                        "D�gagez tous les interrupteurs en tournant les tiges filet�es � la main."
   Message(Titre, 31) = "Impossible de faire la proc�dure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 32)
   Message(Corps, 32) = "Un interrupteur d'origine est ouvert, la proc�dure ne peut pas commencer." & vbCrLf & _
                        "D�gagez tous les interrupteurs en tournant les tiges filet�es � la main."
   Message(Titre, 32) = "Impossible de faire la proc�dure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 33)
   Message(Corps, 33) = "Il y a un probl�me, il semblerait qu'il faille plus de 2mm pour sortir des interrupteurs. On va essayer encore..."
   Message(Titre, 33) = "Recherche de l'origine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 34)
   Message(Corps, 34) = "Les interrupteurs sont encore ouverts. La mise � l'origine n'est pas possible, contr�lez la machine."
   Message(Titre, 34) = "Recherche de l'origine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 35)
   Message(Corps, 35) = "Arr�t chauffe et mouvements par appui sur bouton STOP de la MiniCut2d."
   Message(Titre, 35) = "Arr�t d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 36)
   Message(Corps, 36) = "Un probl�me est survenu lors du calcul du temps de d�coupe."
   Message(Titre, 36) = "Annulation"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 37)
   Message(Corps, 37) = "Pas de trac� charg�." & vbCrLf & "Impossible d'acc�der aux param�tres de d�coupe"
   Message(Titre, 37) = "Op�ration impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 38)
   Message(Corps, 38) = "Faut-il remplacer les valeurs m�moris�e pour cette mati�re par les valeurs affich�es?"
   Message(Titre, 38) = "Mati�re"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 39)
   Message(Corps, 39) = "D�coupe termin�e, fil en position de repos, chauffe coup�e."
   Message(Titre, 39) = "MiniCut2d disponible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 40)
   Message(Corps, 40) = "Attention, la chauffe n'est pas coup�e!"
   Message(Titre, 40) = "Probl�me de communication avec la machine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 41)
   Message(Corps, 41) = "Le bouton Stop a �t� appuy� avant le d�but du mouvement." & vbCrLf & "La chauffe est coup�e."
   Message(Titre, 41) = "Arr�t pendant la chauffe"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 42)
   Message(Corps, 42) = "Le bouton STOP a �t� appuy�, op�ration annul�e."
   Message(Titre, 42) = "Arr�t d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 43)
   Message(Corps, 43) = "Attention, le fil n'est pas � sa position de repos."
   Message(Titre, 43) = "Arr�t dans la zone utile"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 44)
   Message(Corps, 44) = "Le fil est � la position de repos."
   Message(Titre, 44) = "Retour effectu�"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 45)
   Message(Corps, 45) = "Un interrupteur est ouvert, la d�coupe ne peut pas commencer." & vbCrLf & _
                        "D�gagez tous les interrupteurs en tournant les tiges filet�es � la main."
   Message(Titre, 45) = "Impossible de faire la d�coupe."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 46)
   Message(Corps, 46) = "Un interrupteur de mise � l'origine s'est ouvert. Ce n'est pas normal, d�gagez la pi�ce et refaites l'origine."
   Message(Titre, 46) = "Arr�t en cours de d�coupe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 47)
   Message(Corps, 47) = "Un interrupteur de fin de course s'est ouvert. Ce n'est pas normal, d�gagez la pi�ce et refaites l'origine."
   Message(Titre, 47) = "Arr�t en cours de d�coupe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 48)
   Message(Corps, 48) = "Le poussoir d'arr�t d'urgence a �t� appuy�, annulation de la d�coupe."
   Message(Titre, 48) = "Arr�t en cours de d�coupe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 49)
   ' Message(Corps, 49) = "Attention,la chauffe courante et la chauffe de la mati�re s�lectionn�e ont des valeurs diff�rentes" & vbCrLf & _
   '                      "Avant de sauvegarder, faut-il remplacer la chauffe m�moris�e pour cette mati�re par la chauffe courante?"
   ' Message(Titre, 49) = "Sauvegarde mati�re utilis�e"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 50)
   Message(Corps, 50) = "Copie annul�e"
   Message(Titre, 50) = "Copie de fichiers"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 51)
   Message(Corps, 51) = "Les fichiers ont �t� copi�s � l'emplacement demand�."
   Message(Titre, 51) = "Copie de fichiers"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 52)
   Message(Corps, 52) = "La zone utile de la machine est d�pass�e, le projet est tronqu� automatiquement."
   Message(Titre, 52) = "D�passement des courses"
   '
   '      ReDim Preserve Message(Corps To Titre, 1 To 37)
   '      Message(Corps, 37) = "Pas de trac� charg�." & vbCrLf & "Impossible d'acc�der aux param�tres de d�coupe"
   '      Message(Titre, 37) = "Op�ration impossible"
   ReDim Preserve Message(Corps To Titre, 1 To 53)
   Message(Corps, 53) = "Pas de trac� dans le bloc."
   Message(Titre, 53) = "Op�ration impossible"
   '      ReDim Preserve Label(1 To 8)
   '      Label(8) = " Pas de contour transf�r� dans le bloc. S�lectionnez un fichier de la " & vbCrLf & "  biblioth�que puis double-cliquez sur un contour ou faites-le glisser ici."
   
   '      ReDim Preserve Message(Corps To Titre, 1 To 36)
   '      Message(Corps, 36) = "Un probl�me est survenu lors du calcul du temps de d�coupe."
   '      Message(Titre, 36) = "Annulation"
   ReDim Preserve Message(Corps To Titre, 1 To 54)
   Message(Corps, 54) = "Op�ration impossible."
   Message(Titre, 54) = "Annulation"
   
   ReDim Preserve Message(Corps To Titre, 1 To 55)
   Message(Corps, 55) = "Clic droit � l'endroit du zoom" & vbCrLf & _
                        "puis molette de la souris" & vbCrLf & _
                        "ou fl�ches haut et bas du clavier."
   Message(Titre, 55) = "Zoom"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 56)
   Message(Corps, 56) = "La valeur de vitesse m�moris�e avec une mati�re d�passe la valeur maximale possible sur la machine." _
                        & vbCrLf & "La vitesse sera brid�e � la valeur maximale."
   Message(Titre, 56) = "Vitesse trop �lev�e"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 57)
   Message(Corps, 57) = "Le nom de cette mati�re est d�j� en m�moire (mode Expert)." & vbCrLf & _
                        "Ecraser les valeurs?"
   Message(Titre, 57) = "Modification d'une mati�re"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 58)
   Message(Corps, 58) = "Il y a un probl�me : la valeur de vitesse pour la mati�re choisie n'est pas dans les limites pr�vues."
   Message(Titre, 58) = "Valeur incorrecte"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 59)
   Message(Corps, 59) = "La vitesse de cette mati�re ne peut pas �tre chang�e. Cr�ez une nouvelle mati�re."
   Message(Titre, 59) = "Modification impossible"
   '

   
   
   '**** Les Labels des outils, informations, avertissements ****
   ReDim Label(1 To 1)
   Label(1) = "Cliquez sur un point d'un trajet pour le couper en deux trajets."
   '
   ReDim Preserve Label(1 To 2)
   Label(2) = "Cliquez sur le nouveau point d'entr�e"
   '
   ReDim Preserve Label(1 To 3)
   Label(3) = "Angle (�) :"
   '
   ReDim Preserve Label(1 To 4)
   Label(4) = "Ctrl=centrer, Shift=proportion, Alt=miroir"
   '
   ReDim Preserve Label(1 To 5)
   Label(5) = "S�lection : "
   '
   ReDim Preserve Label(1 To 6)
   Label(6) = " Pas de contour visible. S�lectionnez un fichier dans la biblioth�que. "
   '
   ReDim Preserve Label(1 To 7)
   Label(7) = " Pas de contour sous le pointeur. Double-cliquez " & vbCrLf & "  sur un contour ou faites le glisser dans le bloc."
   '
   ReDim Preserve Label(1 To 8)
   Label(8) = " Pas de contour transf�r� dans le bloc. S�lectionnez un fichier de la " & vbCrLf & "  biblioth�que puis double-cliquez sur un contour ou faites-le glisser ici."
   '
   ReDim Preserve Label(1 To 9)
   Label(9) = "Suivant X : "  ' il s'agit d'une dimensions suivant X, le mot � traduire est "Suivant"
   '
   ReDim Preserve Label(1 To 10)
   Label(10) = " mm - Suivant Y : "  ' il s'agit d'une dimensions suivant Y
   '
   ReDim Preserve Label(1 To 11)
   Label(11) = "Retour automatique � la position de repos"
   '
   ReDim Preserve Label(1 To 12)
   Label(12) = "Recherche des interrupteurs"
   '
   ReDim Preserve Label(1 To 13)
   Label(13) = "D�calage vers la position de repos"
   '
   ReDim Preserve Label(1 To 14)
   Label(14) = "Position de repos"
   '
   ReDim Preserve Label(1 To 15)
   Label(15) = "Pr�paration pour rangement"
   '
   ReDim Preserve Label(1 To 16)
   Label(16) = "D�placement vers la position de pliage"
   '
   ReDim Preserve Label(1 To 17)
   Label(17) = "Position de repos"
   '
   ReDim Preserve Label(1 To 18)
   Label(18) = "LE FIL SE DEPLACE"
   '
   ReDim Preserve Label(1 To 19)
   Label(19) = "LE FIL CHAUFFE"
   '
   ReDim Preserve Label(1 To 20)
   Label(20) = "MISE EN TEMPERATURE DU FIL"
   '
   ReDim Preserve Label(1 To 21)
   Label(21) = "dont "
   '
   ReDim Preserve Label(1 To 22)
   Label(22) = " s. de mise en temp�rature."
   '
   ReDim Preserve Label(1 To 23)
   Label(23) = "(Chauffe �"
   '
   ReDim Preserve Label(1 To 24)
   Label(24) = "Proc�dure de remise � l'origine"
   '
   ReDim Preserve Label(1 To 25)
   Label(25) = "Arr�t sur le segment n� "
   '
   ReDim Preserve Label(1 To 26)
   Label(26) = "Retour position de repos"
   '
   ReDim Preserve Label(1 To 27)
   Label(27) = "Position de repos"
   '
   ReDim Preserve Label(1 To 28)
   Label(28) = "Mise en temp�rature du fil"
   '
   ReDim Preserve Label(1 To 29)
   Label(29) = "D�coupe segment n�"
   '
   ReDim Preserve Label(1 To 30)
   Label(30) = "Pilotage du fil"
   '
   ReDim Preserve Label(1 To 31)
   Label(31) = "vitesse �"
   '
   ReDim Preserve Label(1 To 32)
   Label(32) = "Dur�e :"
   '
   ReDim Preserve Label(1 To 33)
   Label(33) = "Retour automatique vertical"
   '
   ReDim Preserve Label(1 To 34)
   Label(34) = "Retour automatique horizontal"
   '
   ReDim Preserve Label(1 To 35)
   Label(35) = "Origine verticale atteinte"
   '
   ReDim Preserve Label(1 To 36)
   Label(36) = "Image couleur ou niveaux de gris"
   '
   ReDim Preserve Label(1 To 37)
   Label(37) = "Image noir et blanc"
   '
   ReDim Preserve Label(1 To 38)
   Label(38) = "Origine horizontale atteinte"
   '
End Sub
