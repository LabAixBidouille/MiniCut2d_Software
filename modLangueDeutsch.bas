Attribute VB_Name = "modLangueDeutsch"
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

Public Sub LangueDeutsch(LangueAUtiliser As String)
   '**** traduction du SplashScreen
   With frmSplashScreen
      .Caption = "MiniCut2d Software - Willkommen !"
      .lblCliquez = "Klicken Sie auf den Typ der Maschine, die Sie verwenden."
      .cmdPasDeMachine.Caption = "Wenn Sie nicht über eine Maschine, klicken Sie hier"
   End With

   '**** traduction du A propos ****
   With frmAboutAndSettings
      .Caption = "Einstellungen"
      .lblVersion.Caption = "Version" & App.Major & "." & App.Minor & "." & App.Revision
      .lblTitle.Caption = TypeMachine & " Software"
      .lblParametres.Caption = "Gesamt X Kurs : " & Format(CourseX, "#####") & "mm - Gesamt Y Kurs : " & Format(CourseY, "#####") & "mm." & _
               vbCrLf & "Offset X Ursprung : " & Format(MmToOriXG, "#0.0##") & "mm, also " & Format(NbrPasToOriXG, "#######") & " Mikroschritte " & _
               vbCrLf & "Gesamt YL Kurs : " & Format(MmToOriYG, "#0.0##") & "mm, also " & Format(NbrPasToOriYG, "#######") & " Mikroschritte " & _
               vbCrLf & "Gesamt YR Kurs : " & Format(MmToOriYD, "#0.0##") & "mm, also " & Format(NbrPasToOriYD, "#######") & " Mikroschritte "
      .lblTraduction.Caption = "Übersetzung :" & vbCrLf & vbCrLf & _
                              "Englisch : aiR-C²/Hugh Potter" & vbCrLf & _
                              "Deutsch : Charles Wittmer" & vbCrLf & _
                              "Spanisch : Enrique Iglesias"
      .frmParametres.Caption = "Einstellungen"
      .frmAPropos.Caption = "Über"
      .frmModeExpert.Caption = "Betried Normal / Expert"
      .chkActiverLeChangementDeMode.Caption = "Aktivieren"
      .optNormalExpert(0).Caption = "Normal Modus"
      .optNormalExpert(1).Caption = "Experten Modus"
   End With
   '**** traduction de la form "Découpe Inactive" *****
   With frmDecoupeInactive
      .lblDecoupeInactive.Caption = "Der Zugang zu diesem Teil der Software ist nicht möglich" & vbCrLf & "weil die USB-Schnittstelle ist nicht erkannt."
   End With
   '**** traduction de la form "Paramètres machine"
   With frmParametres
      .Caption = "Maschinenparameter"
   End With
   '**** traduction de la form de vectorisation ****
   With frmImpConv
      .frmImporterImage.Caption = "Bild"
      .frmRecadrer.Caption = "Beschneiden"
      .frmApercu.Caption = "Schwartz/Weiss"
      .frmLisserTransferer.Caption = "Vektorisieren"
      .cmdImporterImage(0).ToolTipText = "Importieren..."
      .cmdImporterImage(1).ToolTipText = "Kleben"
      .cmdRecadrerLaSelection.ToolTipText = "Beschneiden"
      .optApercu(0).ToolTipText = "Schwartz/Weiss convertieren"
      .optApercu(1).ToolTipText = "Annullieren"
      .optInterieur(0).ToolTipText = "Innere Trassen entfernen"
      .optInterieur(1).ToolTipText = "Innere Trassen bewahren"
      .cmdContoursLissageTransfert.ToolTipText = "Vektorisieren"
      .chkVoirPointsVecto.ToolTipText = "Punkte anzeigen"
      .cmdSauverVecto.ToolTipText = "In .DXF format speichern"
      .cmdQuitterImpConv.ToolTipText = "Validieren"
   End With
   '**** traduction de la form principale ****
   With frmMiniCut2d
      .cmdLangue.Picture = frmImages.imgDrapeauAllemand.Picture   'le drapeau
      ' *** traductions des contrôles de l'interface ***
      .SSTab1.TabCaption(0) = "Schaffung"
      .SSTab1.TabCaption(1) = "Schneiden"
      .lblBoiteOutils.Caption = " Werkzeugkasten "
      .lblAlignerAjuster.Caption = " Ausrichten und anpassen "
      .lblDimensionsBloc.Caption = " Blockmaße "
      .lblTitreChauffe.Caption = " Heiß "
      .lblTrajet.Caption = " Route "
      .lblDecoupe.Caption = " Schneiden "
      .lblFil.Caption = " Draht "
      .lblCadreDecoupe.Caption = " Schneiden "
      .lblStopReprise.Caption = " Stopp / Fortsetzen "
      .lblPiloterLeFil.Caption = " Draht steuerung "
      .frameEntreeBloc.Caption = "Eingang"
      .frameSortieBloc.Caption = "Aussgang"
      .frameSimulation.Caption = "Simulation"
      .frameManuel.Caption = "Fahren"
      .frameProcedures.Caption = "Positionen"
      .frameDecalage.Caption = "Draht verschieben"
      .frameInformation.Caption = "Information"
      .frameAction.Caption = "Schneiden starten"
      .frameAnnulationStop.Caption = "Stornieren - Stopp / Fortsetzen"
      .frameChauffeEnCoursDecoupe.Caption = "Heizung"
      .frameInformationStop.Caption = "Information"
      .frameModifierLaChauffe.Caption = "Heizung ändern"
      .frameRetourOrigine.Caption = "Zurück an Ursprung"
      .frameTrajetRetour.Caption = "Route"
      .frameReprendre.Caption = "Schneiden enden"
      .frameAnnulationReprise.Caption = "Stornieren"
      .frameChauffeFil.Caption = "Aufheizen"
      .frameInformationFil.Caption = "Information"
      .frameFilManuel.Caption = "Bewehgen"
      .frameAnnulationFil.Caption = "Stornieren / Verlassen"
      .cmdNouveauProjet.ToolTipText = "Neues Projekt"
      .cmdOuvrirFichierSequ.ToolTipText = "Projekt öffnen"
      .cmdSauver(0).ToolTipText = "Speichern unter..."
      .cmdSauver(1).ToolTipText = "Speichern"
      .cmdLangue.ToolTipText = "Sprache ändern"
      .cmdSettings.ToolTipText = "Einstellungen"
      .cmdRafraichir.ToolTipText = "Erfrischen"
      .cmdImporterProfil.ToolTipText = "Importieren einer Datei in die Bibliothek"
      .cmdSimulation.ToolTipText = "Draht bewehgung zeigen"
      .optOutils(4).ToolTipText = "Messen"
      .optOutils(5).ToolTipText = "Route schneiden"
      .optOutils(2).ToolTipText = "Strecken"
      .optOutils(1).ToolTipText = "Drehen"
      .optOutils(0).ToolTipText = "Bewehgen"
      .cmdUndo(0).ToolTipText = "Abmachen"
      .cmdUndo(1).ToolTipText = "Neu erstellen"
      .cmdPoubelle.ToolTipText = "Löschen"
      .cmdInsererPoint.ToolTipText = "Ein Punkt zubringen"
      .cmdDupliquer.ToolTipText = "Duplizieren"
      .cmdMiroir.ToolTipText = "Spiegel"
      .cmdInverser.ToolTipText = "Fahrt umkehren"
      .cmdAligner(0).ToolTipText = "Unten ausrichten"
      .cmdAligner(1).ToolTipText = "Mitte ausrichten"
      .cmdAligner(2).ToolTipText = "Oben ausrichten"
      .cmdAligner(3).ToolTipText = "Links ausrichten"
      .cmdAligner(4).ToolTipText = "Mitte ausrichten"
      .cmdAligner(5).ToolTipText = "RRechts ausrichten"
      .chkCentrer(0).ToolTipText = "Skalierung der Block"
      .chkCentrer(1).ToolTipText = "Horizontal zentrieren"
      .chkCentrer(2).ToolTipText = "Vertikal zentrieren"
      .chkVoirPoints.ToolTipText = "Punkte anzeigen"
      .chkCouleurProfils.ToolTipText = "Profilfarben abwechseln"
      .cmdAgrandirRetrecir.ToolTipText = "Fenster größe"
      .chkZoomProjet.ToolTipText = "Zoom-Block"
      .pctZoomInfo.ToolTipText = "Zoom: rechte Maustaste + Mausrad oder Pfeiltasten"
      .cmdGestionMatiere(0).ToolTipText = "Neue Materie"
      .cmdGestionMatiere(1).ToolTipText = "Ersetzen des  Heizungs wert"
      .cmdGestionMatiere(2).ToolTipText = "Materie löschen"
      .cmdDecouper.ToolTipText = "Projekt schneiden"
      .cmdDeplacementsManuels.ToolTipText = "Handsteuern der Drahtführung"
      .cmdPlierLePortique.ToolTipText = "In Lager position gehen"
      .cmdRetourOrigine.ToolTipText = "Draht in Urspung position bringen"
      .optDecalage(1).ToolTipText = "-0.5mm"
      .optDecalage(2).ToolTipText = "0mm"
      .optDecalage(3).ToolTipText = "0.5mm"
      .cmdFaireOrigineAvantDecoupe.ToolTipText = "Schneiden starten"
      .cmdAnnulerDecoupe.ToolTipText = "Starten stornieren"
      .cmdSTOP.ToolTipText = "Not-Aus !"
      .optTrajetRetour(0).ToolTipText = "Diagonal"
      .optTrajetRetour(1).ToolTipText = "Per links"
      .optTrajetRetour(2).ToolTipText = "Per oben"
      .cmdLancerRetourApresStop.ToolTipText = "Zurück starten"
      .cmdStopRetourApresReprise.ToolTipText = "Not-Aus !"
      .cmdRepriseDecoupe.ToolTipText = "Schneiden enden"
      .cmdAnnulerReprise.ToolTipText = "Alles stornieren"
      .optChauffe(0).ToolTipText = "Heizen"
      .optChauffe(1).ToolTipText = "Heizen ausschalten"
      .optGoManuel(0).ToolTipText = "Bewehgung starten"
      .optGoManuel(1).ToolTipText = "Bewehgung stoppen"
      .cmdAnnulerFil.ToolTipText = "Verlassen"
      .optHomeY.ToolTipText = "Vertikal Ursprungs"
      .optHomeX.ToolTipText = "Horizontal Ursprungs"
      .optAnnulerHome.ToolTipText = "Verlassen"
End With
   '**** les MessagBox ****
   ReDim Message(1 To 2, 1 To 1)
   'MessageBox n°1
   Message(Corps, 1) = "Zwei aufeinander folgende Punkte sind zusammenfallen." & vbCrLf & "Unmöglich ein Offset eizustellen"
   Message(Titre, 1) = "Berechnung unmöglich"
   'MessageBox n°2
   ReDim Preserve Message(Corps To Titre, 1 To 2)
   Message(Corps, 2) = "Das Verzeichnis \Bibliothek ist nicht vorhanden, Sie wird erstellt werden, aber lehr sein : an Ihnen Sie zu füllen."
   Message(Titre, 2) = "Bibliothek Initialisierung"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 3)
   Message(Corps, 3) = "Die Datei ""Minicut2d Software.ini"" enthält Schlüssel Die mit ""NbrPasToOri..."" beginnen." & vbCrLf & _
                        "Diese Schlüssel sind nicht mehr gültig und werden durch " & vbCrLf & _
                        "Neue Schlüssel von Art ""MmToOri"" ersetzt"" so dass Minicut2d Software richtig funktionnieren kann."
   Message(Titre, 3) = "Alte .ini Version"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 4)
   Message(Corps, 4) = "Der Heizwert mit einer gespeicherten Materialien überschreitet dem höchst möglichen Wert der Maschine." _
                        & vbCrLf & "Die Heizung wird auf den maximalen Wert geklemmt werden."
   Message(Titre, 4) = "Heizung zu hoch"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 5)
   Message(Corps, 5) = "Diese IPL5XComm.dll auf Ihrem Computer ist zu alt." & vbCrLf & _
                        "MiniCut2d Software kann nicht funktionnieren." & vbCrLf & _
                        "Bitte laden Sie sich die neueste Version und neu starten."
   Message(Titre, 5) = "Kommunikation mit der Maschine unmöglich."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 6)
   Message(Corps, 6) = "Die Initialisierung der USB-Schnittstelle ist problematisch." & vbCrLf & _
                        "Zugang zu Schneide Funktionen unter Umständen nicht möglich"
   Message(Titre, 6) = "Problem der Initialisierung der Interpolation."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 7)
   Message(Corps, 7) = "Die IPL5XCom.dll Datei wurde nicht gefunden," & vbCrLf & _
                        "Bitte installieren Sie sie auf Ihrem Computer, an der richtigen Stelle." & vbCrLf & _
                        "Das Schneiden wird deaktiviert."
   Message(Titre, 7) = "Kommunikation mit der Maschine unmöglich."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 8)
   Message(Corps, 8) = "Es scheint, dass MiniCut2d nicht in der Ruheposition." & vbCrLf & _
                        "Wollen Sie den Ruhestand Prozess der Ablage und Draht vor dem Verlassen der Software beginnen?"
   Message(Titre, 8) = "Anfrage an MiniCut2d Software zu schließen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 9)
   Message(Corps, 9) = "Schalten Sie den MiniCut2d aus vor dem Verlassen des Programms." & vbCrLf & _
                        "Stellen Sie außerdem sicher, dass Ihr Projekt gespeichert ist." & vbCrLf & vbCrLf & _
                        "Bestätigen Sie die Schließung der Anwendung?"
   Message(Titre, 9) = "Anfrage an MiniCut2d Software zu schließen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 10)
   Message(Corps, 10) = "Unmöglich diese Datei zu repräsentieren."
   Message(Titre, 10) = "Fehler beim Lesen der Datei"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 11)
   Message(Corps, 11) = "Wollen Sie das laufende Projekt speichern?"
   Message(Titre, 11) = "Neues Projekt"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 12)
   Message(Corps, 12) = "Invalide Datei Extension. MiniCut2d Software kann die .mnc, .dxf, .dat, .plt, .eps öffnen (und die .txt-Dateien mit den Koordinaten von Punkten durch einen Doppelpunkt getrennt)."
   Message(Titre, 12) = "Kann nicht geöffnet werden"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 13)
   Message(Corps, 13) = "Sie haben auf den ersten Punkt des Profils abgeschnitten, Sie können hier nicht abschneiden."
   Message(Titre, 13) = "Operation nicht möglich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 14)
   Message(Corps, 14) = "Sie haben auf den letzten Punkt des Profils abgeschnitten, Sie können hier nicht abschneiden."
   Message(Titre, 14) = "Operation nicht möglich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 15)
   Message(Corps, 15) = "Bitte geben Sie eine positive oder negative Dezimalzahl ein."
   Message(Titre, 15) = "Eingabe Fehler"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 16)
   Message(Corps, 16) = "Dieser Vorgang ist nicht möglich,  eine Dimension ist gleich Null."
   Message(Titre, 16) = "Unmögliche Berechnung"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 17)
   Message(Corps, 17) = "Der Arbeitsbereich der Maschine ist überschritten, das Projekt wird automatisch abgeschnitten."
   Message(Titre, 17) = "Kurs überschreitung"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 18)
   Message(Corps, 18) = "Der Name dieses Material ist bereits vorhanden, den Wert  überschreiben?"
   Message(Titre, 18) = "Ändern eines Materials"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 19)
   Message(Corps, 19) = "Diese Linie kann nicht gelöscht werden."
   Message(Titre, 19) = "Operation nicht möglich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 20)
   Message(Corps, 20) = "Das Löschen dieses Materials bestätigen?"
   Message(Titre, 20) = "Material löschen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 21)
   Message(Corps, 21) = "Es gibt ein Problem: der Heizwert für das ausgewählte Material ist nicht innerhalb der Grenzen."
   Message(Titre, 21) = "Ungültiger Wert"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 22)
   Message(Corps, 22) = "Initialisierung der " & TypeMachine
   Message(Titre, 22) = "USB-Schnittstelle erkannt"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 23)
   Message(Corps, 23) = "Ruhestellung Prozedur aktivieren?"
   Message(Titre, 23) = "Sicherheit validieren"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 24)
   Message(Corps, 24) = "Der Draht ist in Ruheposition."
   Message(Titre, 24) = "Rückfahrt beendet"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 25)
   Message(Corps, 25) = "Der Vorgang wurde abgebrochen."
   Message(Titre, 25) = "Not Aus"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 26)
   Message(Corps, 26) = "Die Schleife der Ursprungs Schalter ist geöffnet ." & vbCrLf & _
                        "Das ist nicht normal." & vbCrLf & "Sie müssen die Schalter durch von Hand Drehen der Motoren befreien" & _
                        vbCrLf & "(Das Gerät ausschalten, wenn sie sich wehren)."
   Message(Titre, 26) = "Kurs überausgefahren."
   ReDim Preserve Message(Corps To Titre, 1 To 27)
   Message(Corps, 27) = "Die Schleife der Endschalter ist geöffnet ." & vbCrLf & _
                        "Das ist nicht normal." & vbCrLf & "Sie müssen die Schalter durch von Hand Drehen der Motoren befreien" & _
                        vbCrLf & "(Das Gerät ausschalten, wenn sie sich wehren)."
   Message(Titre, 27) = "Kurs überausgefahren."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 28)
   Message(Corps, 28) = "In Ruhe Position fahren?"
   Message(Titre, 28) = "Sicherheit validieren"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 29)
   Message(Corps, 29) = "Falte Position erreicht."
   Message(Titre, 29) = "Bewehgung beendet"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 30)
   Message(Corps, 30) = "Die Schalter des aktuellen Tisch sind nicht aktiv."
   Message(Titre, 30) = "Unmöglich den Prozess zu fahren."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 31)
   Message(Corps, 31) = "Ein Endschalter ist geöffnet, die Prozedur kann nicht beginnen." & vbCrLf & _
                        "Sie müssen alle Schalter durch von Hand Drehen der Motoren befreien."
   Message(Titre, 31) = "Unmöglich den Prozess zu fahren."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 32)
   Message(Corps, 32) = "Ein Ursprung Schalter ist geöffnet, die Prozedur kann nicht vortfahren." & vbCrLf & _
                        "Sie müssen alle Schalter durch von Hand Drehen der Motoren befreien."
   Message(Titre, 32) = "Unmöglich die Prozedur zu starten"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 33)
   Message(Corps, 33) = "Es ist ein Problem, es scheint so, dass es mehr als 2 mm braucht um den Schalter zu befreien. Lasst uns noch einmal versuchen ..."
   Message(Titre, 33) = "Ursprung suche"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 34)
   Message(Corps, 34) = "Die Schalter sind noch geöffnet. In Ursprung position fahren ist nicht möglich, überprüfen Sie die Maschine."
   Message(Titre, 34) = "Ursprung suche"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 35)
   Message(Corps, 35) = "Heizung und Bewehgungen abbrechen durch drücken der MiniCut2d Not-Aus Taste."
   Message(Titre, 35) = "Not-Aus"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 36)
   Message(Corps, 36) = "Es trat ein Problem ein bei der Berechnung der Schneiden Zeit."
   Message(Titre, 36) = "Annulierung"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 37)
   Message(Corps, 37) = "Keine Route geladen" & vbCrLf & "Unmöglich Schneideparametern zu lesen"
   Message(Titre, 37) = "Operation nicht möglich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 38)
   Message(Corps, 38) = "Sollten die Werte im Speicher von der dargestellten Werte?"
   Message(Titre, 38) = "Materie"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 39)
   Message(Corps, 39) = "Schneiden abgeschlossen Draht in Ruheposition, Heizung aus."
   Message(Titre, 39) = "MiniCut2d verfügbar"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 40)
   Message(Corps, 40) = "Vorsicht, heizung nicht ausgeschaltet!"
   Message(Titre, 40) = "Kommunikation Problem mit der Maschine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 41)
   Message(Corps, 41) = "Die Stopp-Taste wurde vor dem Start der Bewegung gedrückt." & vbCrLf & "Die Heizung ist abgebrochen."
   Message(Titre, 41) = "Stoppen beim Erhitzen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 42)
   Message(Corps, 42) = "Die Stop-Taste wurde gedrückt, Betrieb abgebrochen."
   Message(Titre, 42) = "Not-Aus"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 43)
   Message(Corps, 43) = "Vorsicht, der Draht ist nicht in Ruhe position."
   Message(Titre, 43) = "Stop in dem aktiven Bereich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 44)
   Message(Corps, 44) = "Der Draht ist in Ruheposition."
   Message(Titre, 44) = "Rückfahrt beendet"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 45)
   Message(Corps, 45) = "Ein Schalter ist geöffnet, das Schneiden kann nicht beginnen." & vbCrLf & _
                        "Löschen Sie alle Schalter durch von Hand Drehen der Gewindestangen."
   Message(Titre, 45) = "Schneiden unmöglich."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 46)
   Message(Corps, 46) = "Ein Ursprung Schalter ist eröffnet. Das ist nicht normal, entfernen Sie den Block, und wiederholen Sie die Ursprung position."
   Message(Titre, 46) = "Stoppen während des Schneidens!."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 47)
   Message(Corps, 47) = "Ein Schalter ist eröffnet. Das ist nicht normal, entfernen Sie den Block, und wiederholen Sie die Ursprung position."
   Message(Titre, 47) = "Stoppen während des Schneidens!."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 48)
   Message(Corps, 48) = "Die Not-Aus-Taste wurde gedrückt, Zuschnitt ist abgebrochen."
   Message(Titre, 48) = "Stoppen im Laufe des Zuschneidens! "
   '
'      ReDim Preserve Message(Corps To Titre, 1 To 49)
'      Message(Corps, 49) = "Attention," & vbCrLf & _
'                           "- Die Heizung ist derzeit auf" & ChauffeCourante & "% eingestellt" & vbCrLf & _
'                           "Die Materie" & MatiereUtilisee.Nom & " --- " & MatiereUtilisee.Chauffe & "ist ausgewählt," & vbCrLf & _
'                           "Vor dem Speichern sollte die gespeicherte Temperatur für diese Materie mit der aktuelle ersetzen werden?"
'      Message(Titre, 49) = "Speichern verwendete Material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 50)
   Message(Corps, 50) = "Kopieren abgebrochen."
   Message(Titre, 50) = "Kopieren von Dateien"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 51)
   Message(Corps, 51) = "Die Dateien wurden auf die gewünschte Stelle kopiert."
   Message(Titre, 51) = "Kopieren von Dateien"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 52)
   Message(Corps, 52) = "Der Arbeitsbereich der Maschine ist überschritten, das Projekt wird automatisch abgeschnitten."
   Message(Titre, 52) = "Überschreiten der Kurse"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 53)
   Message(Corps, 53) = "Kein Umriss übertragen auf den Block."
   Message(Titre, 53) = "Operation nicht möglich"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 54)
   Message(Corps, 54) = "Operation nicht möglich."
   Message(Titre, 54) = "Annulierung"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 55)
   Message(Corps, 55) = "Rechts klicken, um zu vergrößern Lage" & vbCrLf & _
                        "dann Mausrad" & vbCrLf & _
                        "oder nach oben und unten Pfeiltasten."
   Message(Titre, 55) = "Zoom"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 56)
   Message(Corps, 56) = "Die Geschwindigkeit mit einer gespeicherten Materialien überschreitet dem höchst möglichen Wert der Maschine." _
                        & vbCrLf & "Die Geschwindigkeit wird auf den maximalen Wert geklemmt werden."
   Message(Titre, 56) = "Geschwindigkeit zu hoch"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 57)
   Message(Corps, 57) = "Der Name dieses Material ist bereits im Speicher (Expert-Modus)." & vbCrLf & _
                        "Überschreiben Wert?"
   Message(Titre, 57) = "Ändern eines Materials"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 58)
   Message(Corps, 58) = "Es gibt ein Problem: die Geschwindigkeit für das ausgewählte Material ist nicht innerhalb der Grenzen."
   Message(Titre, 58) = "Ungültiger Wert"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 59)
   Message(Corps, 59) = "Die Geschwindigkeit dieses Material nicht verändert werden kann. Erstellen Sie ein neues Material."
   Message(Titre, 59) = "Änderung nicht möglich."
   '
   '**** Les Labels des outils, informations, avertissements ****
   ReDim Label(1 To 1)
   Label(1) = "Klicken Sie auf einen Punkt auf einer Route, um ihn in zwei Fahrten zu schneiden."
   '
   ReDim Preserve Label(1 To 2)
   Label(2) = "Auswählen per Klick oder per Frame. Verwenden Sie die Strg-Taste für Mehrfachauswahl."
   '
   ReDim Preserve Label(1 To 3)
   Label(3) = "Winkel (°) :"
   '
   ReDim Preserve Label(1 To 4)
   Label(4) = "Strg=Center, Shift=proportion, Alt=Spiegel"
   '
   ReDim Preserve Label(1 To 5)
   Label(5) = "Auswahl : "
   '
   ReDim Preserve Label(1 To 6)
   Label(6) = " Kein sichtbare Umriss. Wählen Sie eine Datei in der Bibliothek. "
   '
   ReDim Preserve Label(1 To 7)
   Label(7) = " Kein Umriss unter dem Zeiger. Doppelklicken Sie auf " & vbCrLf & " auf ein Kontur oder ziehen Sie ihn auf den Block. "
   '
   ReDim Preserve Label(1 To 8)
   Label(8) = " Kein Umriss übertragen auf den Block. Wählen Sie eine Datei in der " & vbCrLf & " Bibliothek und Doppelklicken Sie auf ein Kontur oder ziehen Sie ihn hier. "
   '
   ReDim Preserve Label(1 To 9)
   Label(9) = "Nächste X : "  ' il s'agit d'une dimensions suivant X, le mot à traduire est "Suivant"
   '
   ReDim Preserve Label(1 To 10)
   Label(10) = " mm - Nächste Y : "  ' il s'agit d'une dimensions suivant Y
   '
   ReDim Preserve Label(1 To 11)
   Label(11) = "Automatische Rückkehr in die Ruheposition"
   '
   ReDim Preserve Label(1 To 12)
   Label(12) = "Auf der Suche nach Schaltern"
   '
   ReDim Preserve Label(1 To 13)
   Label(13) = "In Ruheposition verschiebung"
   '
   ReDim Preserve Label(1 To 14)
   Label(14) = "Ruheposition"
   '
   ReDim Preserve Label(1 To 15)
   Label(15) = "Vorbereitung für die Lagerung"
   '
   ReDim Preserve Label(1 To 16)
   Label(16) = "Bewehgung zur Falteposition"
   '
   ReDim Preserve Label(1 To 17)
   Label(17) = "Ruheposition"
   '
   ReDim Preserve Label(1 To 18)
   Label(18) = "Der Draht bewehgt sich"
   '
   ReDim Preserve Label(1 To 19)
   Label(19) = "Der Draht heitzt auf"
   '
   ReDim Preserve Label(1 To 20)
   Label(20) = "Draht in temperatur setzen"
   '
   ReDim Preserve Label(1 To 21)
   Label(21) = "wovon"
   '
   ReDim Preserve Label(1 To 22)
   Label(22) = "s. in Temperatur setzung des Drahtes"      '
   '
   ReDim Preserve Label(1 To 23)
   Label(23) = "Heizung"
   '
   ReDim Preserve Label(1 To 24)
   Label(24) = "Ursprungposition Prozedur"
   '
   ReDim Preserve Label(1 To 25)
   Label(25) = "Halt auf dem Segment Nr."
   '
   ReDim Preserve Label(1 To 26)
   Label(26) = "Rückfahrt in Ruheposition"
   '
   ReDim Preserve Label(1 To 27)
   Label(27) = "Ruheposition"
   '
   ReDim Preserve Label(1 To 28)
   Label(28) = "Draht in temperatur setzen"
   '
   ReDim Preserve Label(1 To 29)
   Label(29) = "Segment Nr. Schneiden"
   '
   ReDim Preserve Label(1 To 30)
   Label(30) = "Fahren des Drahts"
   '
   ReDim Preserve Label(1 To 31)
   Label(31) = "Geschwindigkeit"
   '
   ReDim Preserve Label(1 To 32)
   Label(32) = "Dauer :"
   '
   ReDim Preserve Label(1 To 33)
   Label(33) = "Automatisch vertikal Rückwerts"
   '
   ReDim Preserve Label(1 To 34)
   Label(34) = "Automatisch horizontal Rückwerts"
   '
   ReDim Preserve Label(1 To 35)
   Label(35) = "Vertikal Ursprung erreicht"
   '
   ReDim Preserve Label(1 To 36)
   Label(36) = "Bild in Farbe oder Graustufen"
   '
   ReDim Preserve Label(1 To 38)
   Label(37) = "Schwartz/Weiss Bild"
   '
   ReDim Preserve Label(1 To 38)
   Label(38) = "Horizontal Ursprung erreicht"
   '

End Sub
