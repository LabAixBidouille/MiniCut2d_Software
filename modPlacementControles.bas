Attribute VB_Name = "modPlacementControles"
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

Public Sub InitialiserControles()
   Dim i As Integer
   
   With frmMiniCut2d
      '********* Gestion des polices : on passe tout en Arial *********
      PoliceNormal.Name = "Arial"
      PoliceGras.Name = "Arial"
      PoliceNormal.Bold = False
      PoliceGras.Bold = True
      PoliceNormal.Size = 9
      PoliceGras.Size = 9
      ' affectation des polices
      Set .Tree.Font = PoliceNormal  'treeview
      Set .SSTab1.Font = PoliceNormal 'onglets
      Set .lblBoiteOutils.Font = PoliceGras  'titres
      Set .lblAlignerAjuster.Font = PoliceGras
      Set .lblDimensionsBloc.Font = PoliceGras
      Set .lblTitreChauffe.Font = PoliceGras
      Set .lblTrajet.Font = PoliceGras
      Set .lblDecoupe.Font = PoliceGras
      Set .lblFil.Font = PoliceGras
      Set .lblCadreDecoupe.Font = PoliceGras
      Set .lblStopReprise.Font = PoliceGras
      Set .lblMm(0).Font = PoliceNormal
      Set .lblMm(1).Font = PoliceNormal
      Set .lblBlocMaxi(0).Font = PoliceNormal
      Set .lblBlocMaxi(1).Font = PoliceNormal
      Set .txtBloc(0).Font = PoliceNormal
      Set .txtBloc(1).Font = PoliceNormal
      Set .lblMesure.Font = PoliceNormal
      Set .lblAvertissementTransf.Font = PoliceNormal
      Set .lblAvertissementSequ.Font = PoliceNormal
      Set .lblAvertissementDecoupe.Font = PoliceNormal
      Set .lblMesures(0).Font = PoliceNormal
      Set .lblMesures(1).Font = PoliceNormal
      Set .lblMetre.Font = PoliceNormal
      Set .txtMesures(0).Font = PoliceNormal
      Set .txtMesures(1).Font = PoliceNormal
      Set .lblDimensionSelection.Font = PoliceNormal
      Set .frameEntreeBloc.Font = PoliceNormal
      Set .frameSortieBloc.Font = PoliceNormal
      Set .frameSimulation.Font = PoliceNormal
      Set .frameProcedures.Font = PoliceNormal
      Set .frameManuel.Font = PoliceNormal
      Set .lblChauffeDecoupe.Font = PoliceGras
      Set .comboMatieres.Font = PoliceNormal
      Set .frameMateriaux.Font = PoliceNormal
      Set .lblDureeDecoupe.Font = PoliceNormal
      
      'on masque le slider de zoom en le mettant à l'arrière-plan
      .sliderZoom.ZOrder 1
      'En tout PREMIER, on MASQUE pctDECOUPE car c'est un CRITERE de test pour activer des événements
      .pctDecoupe.Visible = False
      .cmdSimulation.Visible = False 'le bouton de simulation n'est pas dans SSTab1
      .chkZoomDecoupe.Visible = False
      .SSTab1.Tab = 0
      .txtBloc(0).Enabled = True
      .txtBloc(1).Enabled = True
      'définition des dimensions de la fenêtre de découpe
      .pctDecoupe.Left = .pctSequ.Left
      .pctDecoupe.Top = .pctSequ.Top
      .pctDecoupe.Height = frmMiniCut2d.ScaleHeight - 2
      .pctDecoupe.Width = .pctSequ.Width
      .cmdSauver(1).Enabled = False  'le bouton d'enregistrement simple n'est actif que si le projet a un nom
      .progrbarChauffe.Visible = False
      .cmdSTOP.Visible = False
      .cmdSTOP.Top = .cmdAnnulerDecoupe.Top
      
      'box de masquage pendant la simulation
      .pctMasquage.Left = 0
      .pctMasquage.Top = 0
      .pctMasquage.Width = .SSTab1.Width
      .pctMasquage.Height = .ScaleHeight
      .pctMasquage.Visible = False

      'box de validation des découpes
      .pctValidationDecoupe.Left = 0
      .pctValidationDecoupe.Top = 0
      .pctValidationDecoupe.Width = .SSTab1.Width
      .pctValidationDecoupe.Height = .ScaleHeight
      .pctValidationDecoupe.Visible = False
      Set .frameInformation.Font = PoliceNormal
      Set .frameAction.Font = PoliceNormal
      Set .frameAnnulationStop.Font = PoliceNormal
      Set .frameChauffeEnCoursDecoupe.Font = PoliceNormal
      Set .frameDecalage.Font = PoliceNormal
      Set .frmDecalageExpert.Font = PoliceNormal
      Set .lblDecalageExpert.Font = PoliceGras
      .frmDecalageExpert.Top = .frameDecalage.Top
      .optChauffePendantDecoupe(1).Left = .optChauffePendantDecoupe(0).Left
      .optChauffePendantDecoupe(1).Top = .optChauffePendantDecoupe(0).Top
      .optChauffePendantDecoupe(0).ZOrder

      
      'box de reprise après stop
      .pctRepriseDecoupe.Left = 0
      .pctRepriseDecoupe.Top = 0
      .pctRepriseDecoupe.Width = .SSTab1.Width
      .pctRepriseDecoupe.Height = .ScaleHeight
      .pctRepriseDecoupe.Visible = False
      .cmdStopRetourApresReprise.Visible = False
      .cmdStopRetourApresReprise.Top = .cmdLancerRetourApresStop.Top
      .progrbarRetour.Visible = False
      Set .frameInformationStop.Font = PoliceNormal
      Set .frameReprendre.Font = PoliceNormal
      Set .frameModifierLaChauffe.Font = PoliceNormal
      Set .frameRetourOrigine.Font = PoliceNormal
      Set .frameTrajetRetour.Font = PoliceNormal
      Set .frameAnnulationReprise.Font = PoliceNormal
      Set .lblArretParBoutonStop.Font = PoliceNormal
      Set .lblNumSegmentStop.Font = PoliceNormal
      
      'box fil
      .optGoManuel(1).Left = .optGoManuel(0).Left
      .optGoManuel(1).Top = .optGoManuel(0).Top
      .optGoManuel(0).ZOrder
      .optChauffe(1).Left = .optChauffe(0).Left
      .optChauffe(1).Top = .optChauffe(0).Top
      .optChauffe(0).ZOrder
      .pctFil.Left = 0
      .pctFil.Top = 0
      .pctFil.Width = .SSTab1.Width
      .pctFil.Height = .ScaleHeight
      .progrbarChauffeManu.Visible = False
      .pctFil.Visible = False
      .frameFilManuel.Visible = False
      .frameChauffeFil.Visible = False
      Set .frameFilManuel.Font = PoliceNormal
      Set .frameChauffeFil.Font = PoliceNormal
      Set .frameInformationFil.Font = PoliceNormal
      Set .frameAnnulationFil.Font = PoliceNormal
      Set .lblAvertissementFil.Font = PoliceNormal
      Set .lblAvertissementFil2.Font = PoliceNormal
      Set .lblChauffeManuel.Font = PoliceGras
      .lblChauffeManuel.Caption = "0%"
      'cadres de textes
      .lblMesure.Caption = ""
      .lblMesure.Visible = False
      .lblDimensionSelection.Visible = False
      .lblAvertissementTransf.Caption = ""
      .lblAvertissementTransf.Visible = False
      .lblAvertissementSequ.Caption = ""
      .lblAvertissementSequ.Visible = False
      .lblAvertissementDecoupe.Caption = ""
      .lblAvertissementDecoupe.Visible = False
   End With
   With frmAboutAndSettings
      Set .frmAPropos.Font = PoliceNormal
      Set .frmParametres.Font = PoliceNormal
      Set .lblTitle.Font = PoliceNormal
      Set .lblVersion.Font = PoliceNormal
      Set .cmdOK.Font = PoliceNormal
      Set .lblParametres.Font = PoliceNormal
      Set .lblTraduction.Font = PoliceNormal
      Set .frmModeExpert.Font = PoliceNormal
      Set .chkActiverLeChangementDeMode.Font = PoliceNormal
      Set .optNormalExpert(0).Font = PoliceNormal
      Set .optNormalExpert(1).Font = PoliceNormal
   End With
   With frmParametres
      For i = 0 To .optTypeMachine.UBound
         Set .optTypeMachine(i).Font = PoliceNormal
      Next i
      Set .frmTypeMachine.Font = PoliceNormal
      Set .cmdValiderParametresMachine.Font = PoliceNormal
   End With
   With frmDecoupeInactive
      Set .lblDecoupeInactive.Font = PoliceNormal
      Set .cmdOK.Font = PoliceNormal
   End With
   With frmLangue
      For i = 0 To .optLangue.UBound
         Set .optLangue(i).Font = PoliceNormal
      Next i
   End With
   With frmImpConv
      Set .frmImporterImage.Font = PoliceNormal
      Set .frmApercu.Font = PoliceNormal
      Set .frmRecadrer.Font = PoliceNormal
      Set .frmRecadrer.Font = PoliceNormal
      Set .frmLisserTransferer.Font = PoliceNormal
      Set .cmdQuitterImpConv.Font = PoliceNormal
      Set .lblSensibilite.Font = PoliceGras
      Set .lblLissageVecto.Font = PoliceGras
   End With
   With frmSplashScreen
      .lblCliquez.Font.Name = "Arial"
      .cmdPasDeMachine.Font.Name = "Arial"
   End With
End Sub

'***** Affichage/effacement du réglage de la vitesse en fonction du mode *****
Public Sub AffichageEnFonctionDuModeSoft(ModeSoft As String)
   Select Case ModeSoft
   Case "Expert"
      With frmMiniCut2d
         .frameProcedures.Top = 8360
         .frameManuel.Top = 7260
         .lblFil.Top = 6975
         .cmdDecouper.Top = 6065
         .lblDecoupe.Top = 5700
         .frameSortieBloc.Top = 1720
         .frameEntreeBloc.Top = 805
         .frameSimulation.Top = 805
         .lblTrajet.Top = 500
         .frameMateriaux.Top = 3080
         .lblTitreChauffe.Top = 2800
         .imgChauffe.Top = 270
         .hscChauffeDecoupe.Top = 270
         .lblChauffeDecoupe.Top = 300
         .comboMatieres.Top = 1070
         .cmdGestionMatiere(0).Top = 1540
         .cmdGestionMatiere(1).Top = 1540
         .cmdGestionMatiere(2).Top = 1540
         .frameMateriaux.Height = 2440
         .imgVitesse.Top = 700
         .hscVitesseBDD.Top = 670
         .lblVitesseDecoupe.Top = 700
         .imgVitesse.Visible = True
         .hscVitesseBDD.Visible = True
         .lblVitesseDecoupe.Visible = True
         'cadre pilotage manuel
         .frameFilManuel.Height = 3805
         .frameOriginesManu.Top = 2740
         .frmFleches.Top = 700
         .imgVitesseManuel.Visible = True
         .hscVitesseManuel.Visible = True
         .lblValeurVitesseManuel.Visible = True
         .imgVitesseManuel.Top = 365
         .hscVitesseManuel.Top = 340
         .lblValeurVitesseManuel.Top = 360
         .frameDecalage.Visible = False
         .frmDecalageExpert.Visible = True
         .cmdSettings.Picture = frmImages.imgSettingsExpert.Picture
      End With
   Case Else   'Normal
      With frmMiniCut2d
         .imgVitesse.Visible = False
         .hscVitesseBDD.Visible = False
         .lblVitesseDecoupe.Visible = False
         .frameProcedures.Top = 8310
         .frameManuel.Top = 7155
         .lblFil.Top = 6855
         .cmdDecouper.Top = 5880
         .lblDecoupe.Top = 5520
         .frameSortieBloc.Top = 1740
         .frameEntreeBloc.Top = 825
         .frameSimulation.Top = 825
         .lblTrajet.Top = 510
         .frameMateriaux.Top = 3090
         .lblTitreChauffe.Top = 2820
         .imgChauffe.Top = 270
         .hscChauffeDecoupe.Top = 270
         .lblChauffeDecoupe.Top = 300
         .comboMatieres.Top = 765
         .cmdGestionMatiere(0).Top = 1260
         .cmdGestionMatiere(1).Top = 1260
         .cmdGestionMatiere(2).Top = 1260
         .frameMateriaux.Height = 2210
         'cadre pilotage manuel
         .frameFilManuel.Height = 3375
         .frameOriginesManu.Top = 2300
         .imgVitesseManuel.Visible = False
         .hscVitesseManuel.Visible = False
         .lblValeurVitesseManuel.Visible = False
         .frmFleches.Top = 300
         .frmDecalageExpert.Visible = False
         .frameDecalage.Visible = True
         .cmdSettings.Picture = frmImages.imgSettings.Picture
      End With
   End Select
End Sub
