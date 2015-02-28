Attribute VB_Name = "modLangueFrancais"
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

Public Sub LangueFrancais(LangueAUtiliser As String)
   '**** traduction du SplashScreen
   With frmSplashScreen
      .Caption = "MiniCut2d Software - Bienvenue !"
      .lblCliquez = "Cliquez sur le type de machine que vous utilisez"
      .cmdPasDeMachine.Caption = "Si vous n'avez pas de machine, cliquez ici"
   End With
   '**** traduction du A propos ****
   With frmAboutAndSettings
      .Caption = "Paramètres"
      .lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      .lblTitle.Caption = TypeMachine & " Software"
      .lblParametres.Caption = "Course totale X : " & Format(CourseX, "#####") & "mm - Course totale Y : " & Format(CourseY, "#####") & "mm." & _
               vbCrLf & "Décalage inter/origine X : " & Format(MmToOriXG, "#0.0##") & "mm, soit " & Format(NbrPasToOriXG, "#######") & " micropas." & _
               vbCrLf & "Décalage inter/origine YG : " & Format(MmToOriYG, "#0.0##") & "mm, soit " & Format(NbrPasToOriYG, "#######") & " micropas." & _
               vbCrLf & "Décalage inter/origine YD : " & Format(MmToOriYD, "#0.0##") & "mm, soit " & Format(NbrPasToOriYD, "#######") & " micropas."
      .lblTraduction.Caption = "Traduction :" & vbCrLf & vbCrLf & _
                              "Anglais : aiR-C²/Hugh Potter" & vbCrLf & _
                              "Allemand : Charles Wittmer" & vbCrLf & _
                              "Espagnol : Enrique Iglesias"
      .frmParametres.Caption = "Paramètres"
      .frmAPropos.Caption = "A propos"
      .frmModeExpert.Caption = "Fonctionnement Normal / Expert"
      .chkActiverLeChangementDeMode.Caption = "Activer le changement de mode"
      .optNormalExpert(0).Caption = "Mode Normal"
      .optNormalExpert(1).Caption = "Mode Expert"
   End With
   '**** traduction de la form "Découpe Inactive" *****
   With frmDecoupeInactive
      .lblDecoupeInactive.Caption = "L'accès à cette partie du logiciel est impossible" & vbCrLf & "car l'interface USB n'est pas détectée."
   End With
   '**** traduction de la form "Paramètres machine"
   With frmParametres
      .Caption = "Paramètres de la machine"
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
      .optInterieur(0).ToolTipText = "Supprimer les tracés intérieurs"
      .optInterieur(1).ToolTipText = "Conserver les tracés intérieurs"
      .cmdContoursLissageTransfert.ToolTipText = "Vectoriser"
      .chkVoirPointsVecto.ToolTipText = "Affichage des points"
      .cmdSauverVecto.ToolTipText = "Sauver en .DXF"
      .cmdQuitterImpConv.ToolTipText = "Valider"
   End With
   '**** traduction de la form principale ****
   With frmMiniCut2d
      .cmdLangue.Picture = frmImages.imgDrapeauFrancais.Picture
      ' *** traductions des contrôles de l'interface ***
      .SSTab1.TabCaption(0) = "Création"
      .SSTab1.TabCaption(1) = "Découpe"
      .lblBoiteOutils.Caption = " Boite à outils "
      .lblAlignerAjuster.Caption = " Aligner et ajuster "
      .lblDimensionsBloc.Caption = " Dimensions du bloc "
      .lblTitreChauffe.Caption = " Chauffe "
      .lblTrajet.Caption = " Trajet "
      .lblDecoupe.Caption = " Découpe "
      .lblFil.Caption = " Fil "
      .lblCadreDecoupe.Caption = " Découpe "
      .lblStopReprise.Caption = " Stop / Reprise "
      .lblPiloterLeFil.Caption = " Piloter le fil "
      .frameEntreeBloc.Caption = "Entrée"
      .frameSortieBloc.Caption = "Sortie"
      .frameSimulation.Caption = "Simulation"
      .frameManuel.Caption = "Piloter"
      .frameProcedures.Caption = "Positions"
      .frameDecalage.Caption = "Décaler le fil"
      .frameInformation.Caption = "Information"
      .frameAction.Caption = "Lancer la découpe"
      .frameAnnulationStop.Caption = "Annulation - Stop/Reprise"
      .frameChauffeEnCoursDecoupe.Caption = "Chauffe"
      .frameInformationStop.Caption = "Information"
      .frameModifierLaChauffe.Caption = "Modifier la chauffe"
      .frameRetourOrigine.Caption = "Retourner à l'origine"
      .frameTrajetRetour.Caption = "Trajet"
      .frameReprendre.Caption = "Terminer la découpe"
      .frameAnnulationReprise.Caption = "Annuler"
      .frameChauffeFil.Caption = "Faire chauffer"
      .frameInformationFil.Caption = "Information"
      .frameFilManuel.Caption = "Déplacer"
      .frameAnnulationFil.Caption = "Annuler / Quitter"
      .cmdNouveauProjet.ToolTipText = "Nouveau projet"
      .cmdOuvrirFichierSequ.ToolTipText = "Ouvrir un projet"
      .cmdSauver(0).ToolTipText = "Enregistrer sous..."
      .cmdSauver(1).ToolTipText = "Enregistrer"
      .cmdLangue.ToolTipText = "Changer la langue"
      .cmdSettings.ToolTipText = "Paramètres"
      .cmdRafraichir.ToolTipText = "Actualiser"
      .cmdEffacerFichier.ToolTipText = "Supprimer (corbeille)"
      .cmdImporterProfil.ToolTipText = "Importer un fichier dans la bibliothèque"
      .cmdSimulation.ToolTipText = "Voir le déplacement du fil"
      .optOutils(4).ToolTipText = "Mesurer"
      .optOutils(5).ToolTipText = "Couper un trajet"
      .optOutils(2).ToolTipText = "Etirer"
      .optOutils(1).ToolTipText = "Tourner"
      .optOutils(0).ToolTipText = "Déplacer"
      .optOutils(6).ToolTipText = "Modifier le point d'entrée"
      .cmdUndo(0).ToolTipText = "Défaire"
      .cmdUndo(1).ToolTipText = "Refaire"
      .cmdPoubelle.ToolTipText = "Supprimer"
      .cmdInsererPoint.ToolTipText = "Insérer un point"
      .cmdDupliquer.ToolTipText = "Dupliquer"
      .cmdMiroir.ToolTipText = "Miroir"
      .cmdInverser.ToolTipText = "Inverser le sens"
      .cmdAligner(0).ToolTipText = "Aligner en bas"
      .cmdAligner(1).ToolTipText = "Aligner au milieu"
      .cmdAligner(2).ToolTipText = "Aligner en haut"
      .cmdAligner(3).ToolTipText = "Aligner à gauche"
      .cmdAligner(4).ToolTipText = "Aligner au milieu"
      .cmdAligner(5).ToolTipText = "Aligner à droite"
      .chkCentrer(0).ToolTipText = "Mettre à l'échelle du bloc"
      .chkCentrer(1).ToolTipText = "Centrer horizontalement"
      .chkCentrer(2).ToolTipText = "Centrer verticalement"
      .chkVoirPoints.ToolTipText = "Afficher les points"
      .chkCouleurProfils.ToolTipText = "Alterner la couleur des profils"
      .cmdAgrandirRetrecir.ToolTipText = "Taille de la fenêtre"
      .chkZoomProjet.ToolTipText = "Zoomer sur le bloc"
      .pctZoomInfo.ToolTipText = "Zoom : clic droit + molette ou flèches haut et bas"
      .cmdGestionMatiere(0).ToolTipText = "Nouvelle matière"
      .cmdGestionMatiere(1).ToolTipText = "Remplacer la valeur de la chauffe"
      .cmdGestionMatiere(2).ToolTipText = "Supprimer la matière"
      .cmdDecouper.ToolTipText = "Découper le projet"
      .cmdDeplacementsManuels.ToolTipText = "Piloter le fil en direct"
      .cmdPlierLePortique.ToolTipText = "Aller en position de rangement"
      .cmdRetourOrigine.ToolTipText = "Ramener le fil à l'origine"
      .optDecalage(1).ToolTipText = "-0.5mm"
      .optDecalage(2).ToolTipText = "0mm"
      .optDecalage(3).ToolTipText = "0.5mm"
      .cmdFaireOrigineAvantDecoupe.ToolTipText = "Lancer la découpe"
      .cmdAnnulerDecoupe.ToolTipText = "Annuler la découpe"
      .cmdSTOP.ToolTipText = "Arrêt d'urgence !"
      .optTrajetRetour(0).ToolTipText = "En diagonale"
      .optTrajetRetour(1).ToolTipText = "Par la gauche"
      .optTrajetRetour(2).ToolTipText = "Par le haut"
      .cmdLancerRetourApresStop.ToolTipText = "Lancer le retour"
      .cmdStopRetourApresReprise.ToolTipText = "Arrêt d'urgence !"
      .cmdRepriseDecoupe.ToolTipText = "Terminer la découpe"
      .cmdAnnulerReprise.ToolTipText = "Annuler tout"
      .optChauffe(0).ToolTipText = "Faire chauffer"
      .optChauffe(1).ToolTipText = "Arrêter la chauffe"
      .optGoManuel(0).ToolTipText = "Lancer le mouvement"
      .optGoManuel(1).ToolTipText = "Arrêter le mouvement"
      .cmdAnnulerFil.ToolTipText = "Quitter"
      .optHomeY.ToolTipText = "Origine verticale"
      .optHomeX.ToolTipText = "Origine horizontale"
      .optAnnulerHome.ToolTipText = "Stop"
   End With
   '**** les MessagBox ****
   ReDim Message(1 To 2, 1 To 1)
   'MessageBox n°1
   Message(Corps, 1) = "Deux points consécutifs sont confondus." & vbCrLf & "Impossible de définir un décalage."
   Message(Titre, 1) = "Calcul impossible."
   'MessageBox n°2
   ReDim Preserve Message(Corps To Titre, 1 To 2)
   Message(Corps, 2) = "Le répertoire \Bibliotheque n'est pas présent, il va être créé mais il sera vide : à vous de le remplir!"
   Message(Titre, 2) = "Initialisation de la bibliothèque"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 3)
   Message(Corps, 3) = "Le fichier ""MiniCut2d Software.ini"" contient des clés qui commencent par ""NbrPasToOri...""." & vbCrLf & _
                        "Ces clés ne sont plus valides et vont être remplacées par les " & vbCrLf & _
                        "nouvelles clés de type ""MmToOri..."" pour que MiniCut2d Software puisse fonctionner correctement."
   Message(Titre, 3) = "Ancienne version du .ini"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 4)
   Message(Corps, 4) = "La valeur de chauffe mémorisée avec une matière dépasse la valeur maximale possible sur la machine." _
                        & vbCrLf & "La chauffe sera bridée à la valeur maximale."
   Message(Titre, 4) = "Chauffe trop élevée"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 5)
   Message(Corps, 5) = "Le fichier IPL5XComm.dll présent sur votre ordinateur est trop vieux." & vbCrLf & _
                        "MiniCut2d Software ne peut pas fonctionner." & vbCrLf & _
                        "Veuillez télécharger la dernière version et relancer."
   Message(Titre, 5) = "Communication avec la machine impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 6)
   Message(Corps, 6) = "L'initialisation de l'interface USB pose problème" & vbCrLf & _
                        "L'accès aux fonctions de découpe risque d'être impossible"
   Message(Titre, 6) = "Problème d'initialisation de l'interpolateur."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 7)
   Message(Corps, 7) = "Le fichier IPL5XCom.dll n'a pas été trouvé," & vbCrLf & _
                        "veuillez l'installer sur votre ordinateur, à la bonne place." & vbCrLf & _
                        "La découpe sera désactivée."
   Message(Titre, 7) = "Communication avec la machine impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 8)
   Message(Corps, 8) = "Il semblerait que la MiniCut2d ne soit pas en position de rangement." & vbCrLf & _
                        "Voulez-vous lancer la procédure de rangement du plateau et du fil avant de quitter le logiciel ?"
   Message(Titre, 8) = "Demande de fermeture de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 9)
   Message(Corps, 9) = "Coupez l'alimentation de la MiniCut2d avant de quitter le programme." & vbCrLf & _
   "Vérifiez également que votre projet est sauvegardé." & vbCrLf & vbCrLf & _
   "Confirmez-vous la fermeture de l'application ?"
   Message(Titre, 9) = "Demande de fermeture de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 10)
   Message(Corps, 10) = "Impossible de représenter ce fichier."
   Message(Titre, 10) = "Erreur de lecture de fichier"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 11)
   Message(Corps, 11) = "Voulez-vous effectuer une sauvegarde du projet en cours?"
   Message(Titre, 11) = "Nouveau projet"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 12)
   Message(Corps, 12) = "Extension de fichier non valide. MiniCut2d Software sait ouvrir les .mnc, .dxf, .dat, .plt, .eps (et les .txt avec les coordonnées des points séparées par un double-point)."
   Message(Titre, 12) = "Ouverture impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 13)
   Message(Corps, 13) = "Vous avez coupé sur le premier point d'un profil, impossible de couper à cet endroit."
   Message(Titre, 13) = "Opération impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 14)
   Message(Corps, 14) = "Vous avez coupé sur le dernier point d'un profil, impossible de couper à cet endroit."
   Message(Titre, 14) = "Opération impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 15)
   Message(Corps, 15) = "Veuillez entrer un nombre décimal positif ou négatif."
   Message(Titre, 15) = "Erreur de saisie"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 16)
   Message(Corps, 16) = "L'opération demandée est impossible, une dimension est égale à zéro."
   Message(Titre, 16) = "Calcul impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 17)
   Message(Corps, 17) = "La zone utile de la machine est dépassée, le projet est tronqué automatiquement."
   Message(Titre, 17) = "Dépassement des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 18)
   Message(Corps, 18) = "Le nom de cette matière existe déjà, écraser les valeurs?"
   Message(Titre, 18) = "Modification d'une matière"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 19)
   Message(Corps, 19) = "Cette ligne ne peut pas être effacée."
   Message(Titre, 19) = "Opération impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 20)
   Message(Corps, 20) = "Valider l'effacement de ce matériau?"
   Message(Titre, 20) = "Supprimer un matériau"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 21)
   Message(Corps, 21) = "Il y a un problème : la valeur de chauffe pour la matière choisie n'est pas dans les limites prévues."
   Message(Titre, 21) = "Valeur incorrecte"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 22)
   Message(Corps, 22) = "Initialisation de la " & TypeMachine
   Message(Titre, 22) = "Interface USB détectée"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 23)
   Message(Corps, 23) = "Exécuter la procédure de retour à la position de repos ?"
   Message(Titre, 23) = "Validation de sécurité"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 24)
   Message(Corps, 24) = "Le fil est à la position de repos."
   Message(Titre, 24) = "Retour effectué"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 25)
   Message(Corps, 25) = "L'opération a été annulée."
   Message(Titre, 25) = "Arrêt d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 26)
   Message(Corps, 26) = "La boucle des interrupteurs d'origine est ouverte." & vbCrLf & _
                        "Ce n'est pas normal." & vbCrLf & "Il faut dégager les interrupteurs en tournant les moteurs à la main" & _
                        vbCrLf & "(couper l'alimentation s'ils résistent)."
   Message(Titre, 26) = "Sortie des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 27)
   Message(Corps, 27) = "La boucle des interrupteurs de fin de course est ouverte." & vbCrLf & _
                        "Ce n'est pas normal." & vbCrLf & "Il faut dégager les interrupteurs en tournant les moteurs à la main" & _
                        vbCrLf & "(couper l'alimentation s'ils résistent)."
   Message(Titre, 27) = "Sortie des courses"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 28)
   Message(Corps, 28) = "Déplacement en position de rangement?"
   Message(Titre, 28) = "Validation de sécurité"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 29)
   Message(Corps, 29) = "Position de pliage atteinte."
   Message(Titre, 29) = "Mouvement effectué"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 30)
   Message(Corps, 30) = "Les interrupteurs de la table courante ne sont pas actifs."
   Message(Titre, 30) = "Impossible de faire la procédure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 31)
   Message(Corps, 31) = "Un interrupteur de fin de course est ouvert, la procédure ne peut pas commencer." & vbCrLf & _
                        "Dégagez tous les interrupteurs en tournant les tiges filetées à la main."
   Message(Titre, 31) = "Impossible de faire la procédure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 32)
   Message(Corps, 32) = "Un interrupteur d'origine est ouvert, la procédure ne peut pas commencer." & vbCrLf & _
                        "Dégagez tous les interrupteurs en tournant les tiges filetées à la main."
   Message(Titre, 32) = "Impossible de faire la procédure."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 33)
   Message(Corps, 33) = "Il y a un problème, il semblerait qu'il faille plus de 2mm pour sortir des interrupteurs. On va essayer encore..."
   Message(Titre, 33) = "Recherche de l'origine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 34)
   Message(Corps, 34) = "Les interrupteurs sont encore ouverts. La mise à l'origine n'est pas possible, contrôlez la machine."
   Message(Titre, 34) = "Recherche de l'origine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 35)
   Message(Corps, 35) = "Arrêt chauffe et mouvements par appui sur bouton STOP de la MiniCut2d."
   Message(Titre, 35) = "Arrêt d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 36)
   Message(Corps, 36) = "Un problème est survenu lors du calcul du temps de découpe."
   Message(Titre, 36) = "Annulation"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 37)
   Message(Corps, 37) = "Pas de tracé chargé." & vbCrLf & "Impossible d'accéder aux paramètres de découpe"
   Message(Titre, 37) = "Opération impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 38)
   Message(Corps, 38) = "Faut-il remplacer les valeurs mémorisée pour cette matière par les valeurs affichées?"
   Message(Titre, 38) = "Matière"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 39)
   Message(Corps, 39) = "Découpe terminée, fil en position de repos, chauffe coupée."
   Message(Titre, 39) = "MiniCut2d disponible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 40)
   Message(Corps, 40) = "Attention, la chauffe n'est pas coupée!"
   Message(Titre, 40) = "Problème de communication avec la machine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 41)
   Message(Corps, 41) = "Le bouton Stop a été appuyé avant le début du mouvement." & vbCrLf & "La chauffe est coupée."
   Message(Titre, 41) = "Arrêt pendant la chauffe"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 42)
   Message(Corps, 42) = "Le bouton STOP a été appuyé, opération annulée."
   Message(Titre, 42) = "Arrêt d'urgence"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 43)
   Message(Corps, 43) = "Attention, le fil n'est pas à sa position de repos."
   Message(Titre, 43) = "Arrêt dans la zone utile"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 44)
   Message(Corps, 44) = "Le fil est à la position de repos."
   Message(Titre, 44) = "Retour effectué"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 45)
   Message(Corps, 45) = "Un interrupteur est ouvert, la découpe ne peut pas commencer." & vbCrLf & _
                        "Dégagez tous les interrupteurs en tournant les tiges filetées à la main."
   Message(Titre, 45) = "Impossible de faire la découpe."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 46)
   Message(Corps, 46) = "Un interrupteur de mise à l'origine s'est ouvert. Ce n'est pas normal, dégagez la pièce et refaites l'origine."
   Message(Titre, 46) = "Arrêt en cours de découpe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 47)
   Message(Corps, 47) = "Un interrupteur de fin de course s'est ouvert. Ce n'est pas normal, dégagez la pièce et refaites l'origine."
   Message(Titre, 47) = "Arrêt en cours de découpe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 48)
   Message(Corps, 48) = "Le poussoir d'arrêt d'urgence a été appuyé, annulation de la découpe."
   Message(Titre, 48) = "Arrêt en cours de découpe!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 49)
   ' Message(Corps, 49) = "Attention,la chauffe courante et la chauffe de la matière sélectionnée ont des valeurs différentes" & vbCrLf & _
   '                      "Avant de sauvegarder, faut-il remplacer la chauffe mémorisée pour cette matière par la chauffe courante?"
   ' Message(Titre, 49) = "Sauvegarde matière utilisée"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 50)
   Message(Corps, 50) = "Copie annulée"
   Message(Titre, 50) = "Copie de fichiers"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 51)
   Message(Corps, 51) = "Les fichiers ont été copiés à l'emplacement demandé."
   Message(Titre, 51) = "Copie de fichiers"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 52)
   Message(Corps, 52) = "La zone utile de la machine est dépassée, le projet est tronqué automatiquement."
   Message(Titre, 52) = "Dépassement des courses"
   '
   '      ReDim Preserve Message(Corps To Titre, 1 To 37)
   '      Message(Corps, 37) = "Pas de tracé chargé." & vbCrLf & "Impossible d'accéder aux paramètres de découpe"
   '      Message(Titre, 37) = "Opération impossible"
   ReDim Preserve Message(Corps To Titre, 1 To 53)
   Message(Corps, 53) = "Pas de tracé dans le bloc."
   Message(Titre, 53) = "Opération impossible"
   '      ReDim Preserve Label(1 To 8)
   '      Label(8) = " Pas de contour transféré dans le bloc. Sélectionnez un fichier de la " & vbCrLf & "  bibliothèque puis double-cliquez sur un contour ou faites-le glisser ici."
   
   '      ReDim Preserve Message(Corps To Titre, 1 To 36)
   '      Message(Corps, 36) = "Un problème est survenu lors du calcul du temps de découpe."
   '      Message(Titre, 36) = "Annulation"
   ReDim Preserve Message(Corps To Titre, 1 To 54)
   Message(Corps, 54) = "Opération impossible."
   Message(Titre, 54) = "Annulation"
   
   ReDim Preserve Message(Corps To Titre, 1 To 55)
   Message(Corps, 55) = "Clic droit à l'endroit du zoom" & vbCrLf & _
                        "puis molette de la souris" & vbCrLf & _
                        "ou flèches haut et bas du clavier."
   Message(Titre, 55) = "Zoom"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 56)
   Message(Corps, 56) = "La valeur de vitesse mémorisée avec une matière dépasse la valeur maximale possible sur la machine." _
                        & vbCrLf & "La vitesse sera bridée à la valeur maximale."
   Message(Titre, 56) = "Vitesse trop élevée"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 57)
   Message(Corps, 57) = "Le nom de cette matière est déjà en mémoire (mode Expert)." & vbCrLf & _
                        "Ecraser les valeurs?"
   Message(Titre, 57) = "Modification d'une matière"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 58)
   Message(Corps, 58) = "Il y a un problème : la valeur de vitesse pour la matière choisie n'est pas dans les limites prévues."
   Message(Titre, 58) = "Valeur incorrecte"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 59)
   Message(Corps, 59) = "La vitesse de cette matière ne peut pas être changée. Créez une nouvelle matière."
   Message(Titre, 59) = "Modification impossible"
   '

   
   
   '**** Les Labels des outils, informations, avertissements ****
   ReDim Label(1 To 1)
   Label(1) = "Cliquez sur un point d'un trajet pour le couper en deux trajets."
   '
   ReDim Preserve Label(1 To 2)
   Label(2) = "Cliquez sur le nouveau point d'entrée"
   '
   ReDim Preserve Label(1 To 3)
   Label(3) = "Angle (°) :"
   '
   ReDim Preserve Label(1 To 4)
   Label(4) = "Ctrl=centrer, Shift=proportion, Alt=miroir"
   '
   ReDim Preserve Label(1 To 5)
   Label(5) = "Sélection : "
   '
   ReDim Preserve Label(1 To 6)
   Label(6) = " Pas de contour visible. Sélectionnez un fichier dans la bibliothèque. "
   '
   ReDim Preserve Label(1 To 7)
   Label(7) = " Pas de contour sous le pointeur. Double-cliquez " & vbCrLf & "  sur un contour ou faites le glisser dans le bloc."
   '
   ReDim Preserve Label(1 To 8)
   Label(8) = " Pas de contour transféré dans le bloc. Sélectionnez un fichier de la " & vbCrLf & "  bibliothèque puis double-cliquez sur un contour ou faites-le glisser ici."
   '
   ReDim Preserve Label(1 To 9)
   Label(9) = "Suivant X : "  ' il s'agit d'une dimensions suivant X, le mot à traduire est "Suivant"
   '
   ReDim Preserve Label(1 To 10)
   Label(10) = " mm - Suivant Y : "  ' il s'agit d'une dimensions suivant Y
   '
   ReDim Preserve Label(1 To 11)
   Label(11) = "Retour automatique à la position de repos"
   '
   ReDim Preserve Label(1 To 12)
   Label(12) = "Recherche des interrupteurs"
   '
   ReDim Preserve Label(1 To 13)
   Label(13) = "Décalage vers la position de repos"
   '
   ReDim Preserve Label(1 To 14)
   Label(14) = "Position de repos"
   '
   ReDim Preserve Label(1 To 15)
   Label(15) = "Préparation pour rangement"
   '
   ReDim Preserve Label(1 To 16)
   Label(16) = "Déplacement vers la position de pliage"
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
   Label(22) = " s. de mise en température."
   '
   ReDim Preserve Label(1 To 23)
   Label(23) = "(Chauffe à"
   '
   ReDim Preserve Label(1 To 24)
   Label(24) = "Procédure de remise à l'origine"
   '
   ReDim Preserve Label(1 To 25)
   Label(25) = "Arrêt sur le segment n° "
   '
   ReDim Preserve Label(1 To 26)
   Label(26) = "Retour position de repos"
   '
   ReDim Preserve Label(1 To 27)
   Label(27) = "Position de repos"
   '
   ReDim Preserve Label(1 To 28)
   Label(28) = "Mise en température du fil"
   '
   ReDim Preserve Label(1 To 29)
   Label(29) = "Découpe segment n°"
   '
   ReDim Preserve Label(1 To 30)
   Label(30) = "Pilotage du fil"
   '
   ReDim Preserve Label(1 To 31)
   Label(31) = "vitesse à"
   '
   ReDim Preserve Label(1 To 32)
   Label(32) = "Durée :"
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
