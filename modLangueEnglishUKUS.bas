Attribute VB_Name = "modLangueEnglishUKUS"
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

Public Sub LangueEnglishUKUS(LangueAUtiliser As String)
   '**** traduction du SplashScreen
   With frmSplashScreen
      .Caption = "MiniCut2d Software - Welcome !"
      .lblCliquez = "Click on the type of machine you are using"
      .cmdPasDeMachine.Caption = "If you do not have a machine, click here"
   End With
   '**** traduction du A propos ****
   With frmAboutAndSettings
      .Caption = "Settings"
      .lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
      .lblTitle.Caption = TypeMachine & " Software"
      .lblParametres.Caption = "Whole course X : " & Format(CourseX, "#####") & "mm – Whole course Y : " & Format(CourseY, "#####") & "mm." & _
               vbCrLf & "Interval home switch/origin X : " & Format(MmToOriXG, "#0.0##") & "mm, or " & Format(NbrPasToOriXG, "######") & " microsteps." & _
               vbCrLf & "Interval home switch/origin Y left : " & Format(MmToOriYG, "#0.0##") & "mm, or " & Format(NbrPasToOriYG, "#######") & " microsteps." & _
               vbCrLf & "Interval home switch/origin Y right : " & Format(MmToOriYD, "#0.0##") & "mm, or " & Format(NbrPasToOriYD, "#######") & " microsteps."
      .lblTraduction.Caption = "Translation :" & vbCrLf & vbCrLf & _
                              "English : aiR-C²/Hugh Potter" & vbCrLf & _
                              "Deutch : Charles Wittmer" & vbCrLf & _
                              "Espanol : Enrique Iglesias"
      .frmParametres.Caption = "Settings"
      .frmAPropos.Caption = "About"
      .frmModeExpert.Caption = "Operation Normal / Expert"
      .chkActiverLeChangementDeMode.Caption = "Activate"
      .optNormalExpert(0).Caption = "Normal mode"
      .optNormalExpert(1).Caption = "Expert mode"
   End With
   '**** traduction de la form "Découpe Inactive" *****
   With frmDecoupeInactive
      .lblDecoupeInactive.Caption = "The acces to this part of the software is impossible" & vbCrLf & "because  the USB interface is not detected."
   End With
   '**** traduction de la form "Paramètres machine"
   With frmParametres
      .Caption = "Machine parameters"
   End With
   '**** traduction de la form principale ****
   With frmMiniCut2d
      If LangueAUtiliser = "USA" Then
         .cmdLangue.Picture = frmImages.imgDrapeauAmericain.Picture  'le drapeau
      Else
         .cmdLangue.Picture = frmImages.imgDrapeauAnglais.Picture 'le drapeau
      End If
      .SSTab1.TabCaption(0) = "Creation"
      .SSTab1.TabCaption(1) = "Cut"
      .lblBoiteOutils.Caption = " Toolbox "
      .lblAlignerAjuster.Caption = " Align and adjust "
      .lblDimensionsBloc.Caption = " Block size "
      .lblTitreChauffe.Caption = " Heating "
      .lblTrajet.Caption = " Path "
      .lblDecoupe.Caption = " Cutting "
      .lblFil.Caption = " Wire "
      .lblCadreDecoupe.Caption = " Cut "
      .lblStopReprise.Caption = " Stop / Resume "
      .lblPiloterLeFil.Caption = " Wire driver "
      .frameEntreeBloc.Caption = "Input"
      .frameSortieBloc.Caption = "Output"
      .frameSimulation.Caption = "Simulation"
      .frameManuel.Caption = "Operate"
      .frameProcedures.Caption = "Positions"
      .frameDecalage.Caption = "Shift the wire"
      .frameInformation.Caption = "Information"
      .frameAction.Caption = "Start cutting"
      .frameAnnulationStop.Caption = "Cancel - Stop / Resume"
      .frameChauffeEnCoursDecoupe.Caption = "Heater"
      .frameInformationStop.Caption = "Information"
      .frameModifierLaChauffe.Caption = "Modify heater"
      .frameRetourOrigine.Caption = "Return to origin"
      .frameTrajetRetour.Caption = "Path" '
      .frameReprendre.Caption = "Finish cutting" 'Finish the cut ?
      .frameAnnulationReprise.Caption = "Cancel"
      .frameChauffeFil.Caption = "Heat up"
      .frameInformationFil.Caption = "Information"
      .frameFilManuel.Caption = "Move"
      .frameAnnulationFil.Caption = "Cancel / End Session"
      .cmdNouveauProjet.ToolTipText = "New project"
      .cmdOuvrirFichierSequ.ToolTipText = "Open a project"
      .cmdSauver(0).ToolTipText = "Save as..."
      .cmdSauver(1).ToolTipText = "Save"
      .cmdLangue.ToolTipText = "Change language"
      .cmdSettings.ToolTipText = "Settings"
      .cmdRafraichir.ToolTipText = "Update"
      .cmdEffacerFichier.ToolTipText = "Delete (trash)"
      .cmdImporterProfil.ToolTipText = "Import a file in the library"
      .cmdSimulation.ToolTipText = "See the movement of the wire"
      .optOutils(4).ToolTipText = "Measure"
      .optOutils(5).ToolTipText = "Cut a path"
      .optOutils(2).ToolTipText = "Stretch"
      .optOutils(1).ToolTipText = "Turn"
      .optOutils(0).ToolTipText = "Move"
      .optOutils(6).ToolTipText = "Change the entry point"
      .cmdUndo(0).ToolTipText = "Undo"
      .cmdUndo(1).ToolTipText = "Redo"
      .cmdPoubelle.ToolTipText = "Delete"
      .cmdInsererPoint.ToolTipText = "Insert a point"
      .cmdDupliquer.ToolTipText = "Duplicate"
      .cmdMiroir.ToolTipText = "Mirror"
      .cmdInverser.ToolTipText = "Reverse direction"
      .cmdAligner(0).ToolTipText = "Align Bottom"
      .cmdAligner(1).ToolTipText = "Align in the middle"
      .cmdAligner(2).ToolTipText = "Align at the top"
      .cmdAligner(3).ToolTipText = "Align to the left"
      .cmdAligner(4).ToolTipText = "Align in the middle"
      .cmdAligner(5).ToolTipText = "Align to the right"
      .chkCentrer(0).ToolTipText = "Block rescaling"
      .chkCentrer(1).ToolTipText = "Center horizontally"
      .chkCentrer(2).ToolTipText = "Center vertically"
      .chkVoirPoints.ToolTipText = "Display points"
      .chkCouleurProfils.ToolTipText = "Alternate profil colors"
      If LangueAUtiliser = "english" Then
         .chkCouleurProfils.ToolTipText = "Alternate profil colours"
      Else  'USA
         .chkCouleurProfils.ToolTipText = "Alternate profil colors"
      End If
      .cmdAgrandirRetrecir.ToolTipText = "Windows size"
      .chkZoomProjet.ToolTipText = "Zoom extends"
      .pctZoomInfo.ToolTipText = "Zoom: right click + mouse wheel or up and down arrows"
      .cmdGestionMatiere(0).ToolTipText = "New material"
      .cmdGestionMatiere(1).ToolTipText = "Replace the value of the heating"
      .cmdGestionMatiere(2).ToolTipText = "Delete the material"
      .cmdDecouper.ToolTipText = "Cut project"
      .cmdDeplacementsManuels.ToolTipText = "Manually control the wire"
      .cmdPlierLePortique.ToolTipText = "Back to the safe position (storage)"
      .cmdRetourOrigine.ToolTipText = "Take the wire back to the origin"
      .optDecalage(1).ToolTipText = "-0.5mm"
      .optDecalage(2).ToolTipText = "0mm"
      .optDecalage(3).ToolTipText = "0.5mm"
      .cmdFaireOrigineAvantDecoupe.ToolTipText = "Start cutting"
      .cmdAnnulerDecoupe.ToolTipText = "Cancel the cut"
      .cmdSTOP.ToolTipText = "Emergency stop!"
      .optTrajetRetour(0).ToolTipText = "Diagonal"
      .optTrajetRetour(1).ToolTipText = "From left"
      .optTrajetRetour(2).ToolTipText = "From top"
      .cmdLancerRetourApresStop.ToolTipText = "Launch the return"
      .cmdStopRetourApresReprise.ToolTipText = "Emergency stop!"
      .cmdRepriseDecoupe.ToolTipText = "Complete cutting"
      .cmdAnnulerReprise.ToolTipText = "Cancel"
      .optChauffe(0).ToolTipText = "Heat up"
      .optChauffe(1).ToolTipText = "Stop heating"
      .optGoManuel(0).ToolTipText = "Throwing motion"
      .optGoManuel(1).ToolTipText = "Stop motion"
      .cmdAnnulerFil.ToolTipText = "Cancel"
      .optHomeY.ToolTipText = "Vertical origin"
      .optHomeX.ToolTipText = "Horizontal origin"
      .optAnnulerHome.ToolTipText = "Stop"
   End With
   '**** les MsgBox ****
   ReDim Message(1 To 2, 1 To 1)
   'MessageBox n°1
   Message(Corps, 1) = "Two successive points coincide." & vbCrLf & "Unable to set an offset."
   Message(Titre, 1) = "Calculation is impossible."
   'MessageBox n°2
   ReDim Preserve Message(Corps To Titre, 1 To 2)
   Message(Corps, 2) = "The directory is not present or library is not present, it will be created, but it will be empty : up to you to complete it!"
   Message(Titre, 2) = "Library initialization"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 3)
   Message(Corps, 3) = "The file ""MiniCut2d Software.ini"" contains keys starting with ""NbrPasToOri...""." & vbCrLf & _
                        "These keys are no longer valid and will be replaced by" & vbCrLf & _
                        "new key type ""MmToOri..."" as for MiniCut2d Software can work corretly"
   Message(Titre, 3) = "Old version of .ini"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 4)
   Message(Corps, 4) = " The heating value with a stored material exceeds the maximum value possible on the machine." _
                        & vbCrLf & "The heater is set to the maximum value."
   Message(Titre, 4) = "Excessive heating"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 5)
   Message(Corps, 5) = "The files IPL5XComm.dll is out of date ." & vbCrLf & _
                        "MiniCut2d Software can not operate." & vbCrLf & _
                        "Please download the latest version and restart."
   Message(Titre, 5) = "Communication impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 6)
   Message(Corps, 6) = "Bad initialization of the USB interface" & vbCrLf & _
                        "Access to cutting functions may be impossible"
   Message(Titre, 6) = "Problem initializing interpolation."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 7)
   Message(Corps, 7) = "The file IPL5XCom.dll was not found," & vbCrLf & _
                        "Please install it at the right place on your computer." & vbCrLf & _
                        "The cut will be disabled.."
   Message(Titre, 7) = "Communication impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 8)
   Message(Corps, 8) = "It appears that the MiniCut2d may not be in the storage position." & vbCrLf & _
                        "Do you want to initiate the procedure for storage tray and wire before exit ?"
   Message(Titre, 8) = "Request to close MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 9)
   Message(Corps, 9) = "Disconnect power suply from the MiniCut2d before exiting the program." & vbCrLf & _
   "Also make sure your project is saved" & vbCrLf & vbCrLf & _
   "EXIT ?"
   Message(Titre, 9) = "Closing MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 10)
   Message(Corps, 10) = "Impossible to represent this file."
   Message(Titre, 10) = "Read error"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 11)
   Message(Corps, 11) = "Do you want save the current project?"
   Message(Titre, 11) = "New project"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 12)
   Message(Corps, 12) = "File extension not valid. MiniCut2d Software can open .mnc, .dxf, .dat, .plt, .eps (and .txt files with the coordinates of points separated by colons)."
   Message(Titre, 12) = "Opening impossible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 13)
   Message(Corps, 13) = "You cut on the first point of the profile, you can not cut here."
   Message(Titre, 13) = "Impossible operation"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 14)
   Message(Corps, 14) = "You have cut on the last point of the profil, impossible to cut here."
   Message(Titre, 14) = "Operation not possible "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 15)
   Message(Corps, 15) = "Please enter a positive or negative decimal number."
   Message(Titre, 15) = "Input Error "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 16)
   Message(Corps, 16) = "The requested operation is not possible, one of the dimensions is zero"
   Message(Titre, 16) = "Computation not possible "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 17)
   Message(Corps, 17) = "The working area of the machine is exceeded, the project is automatically truncated."
   Message(Titre, 17) = "Exceeding machine size"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 18)
   Message(Corps, 18) = "The name of this material already exists, overwrite the values?"
   Message(Titre, 18) = "Changing the characteristic of a material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 19)
   Message(Corps, 19) = "This line can not be erased."
   Message(Titre, 19) = "Impossible operation"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 20)
   Message(Corps, 20) = "Confirm the deletion of this material?"
   Message(Titre, 20) = "Delete a material "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 21)
   Message(Corps, 21) = "There is a problem: the heating value for the selected material is not within the limits set."
   Message(Titre, 21) = "Incorrect value "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 22)
   Message(Corps, 22) = TypeMachine & "’s initialization"
   Message(Titre, 22) = "USB interface detected"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 23)
   Message(Corps, 23) = "Run to idle position?"
   Message(Titre, 23) = "Safety validation"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 24)
   Message(Corps, 24) = " The wire is at home position."
   Message(Titre, 24) = "Return made"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 25)
   Message(Corps, 25) = "The operation was canceled."
   Message(Titre, 25) = "Emergency Stop "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 26)
   Message(Corps, 26) = "The loop of the home switch is open." & vbCrLf & _
                        "This is not right." & vbCrLf & "It is necessary to release the switches by turning the threaded rods manually" & _
                        vbCrLf & "(turn off the power suply if motor resists)."
   Message(Titre, 26) = "Limit switch exceeded"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 27)
   Message(Corps, 27) = "The loop of the limit switch is open." & vbCrLf & _
                        "This is not right." & vbCrLf & "It is necessary to release the switches by turning the threaded rods manually" & _
                        vbCrLf & "(Turn off the power suply if motor resists)."
   Message(Titre, 27) = "Limit switch exceeded "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 28)
   Message(Corps, 28) = "Return to storage position?"
   Message(Titre, 28) = "Safety validation "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 29)
   Message(Corps, 29) = "Storage position reached."
   Message(Titre, 29) = "Movement done"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 30)
   Message(Corps, 30) = "Switches of current table are inactive."
   Message(Titre, 30) = "Can not do this"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 31)
   Message(Corps, 31) = "A limit switch is open, the process can not start." & vbCrLf & _
                        "Release all the switches by manually turning threaded rods."
   Message(Titre, 31) = "Can not do this."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 32)
   Message(Corps, 32) = "A home switch is open, the process can not start.." & vbCrLf & _
                        "Release all the switches by manually turning threaded rods."
   Message(Titre, 32) = "Can not do this."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 33)
   Message(Corps, 33) = "There is a problem,it would appear that more than 2 mm are necessary to release the switches. We'll try again ......"
   Message(Titre, 33) = "Home search"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 34)
   Message(Corps, 34) = "The switches are still open. Going back to the origin is not possible, check the machine."
   Message(Titre, 34) = "Home search"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 35)
   Message(Corps, 35) = "Stop heating and movements by pressing STOP button "
   Message(Titre, 35) = "Emergency stop"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 36)
   Message(Corps, 36) = "A problem occurred while calculating the cutting time."
   Message(Titre, 36) = "Cancel"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 37)
   Message(Corps, 37) = "No route loaded" & vbCrLf & "Unable to access cutting parameters "
   Message(Titre, 37) = "Operation not possible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 38)
   Message(Corps, 38) = "Should the values in memory by the values shown?"
   Message(Titre, 38) = "Material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 39)
   Message(Corps, 39) = "Cutting completed, wire at home position, heating off."
   Message(Titre, 39) = "MiniCut2d ready"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 40)
   Message(Corps, 40) = "Be careful! The heater works!"
   Message(Titre, 40) = "Communication problem with the machine"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 41)
   Message(Corps, 41) = "Stop button has been pressed before the start of the movement." & vbCrLf & " The heater is turned off."
   Message(Titre, 41) = "Stopping during heating"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 42)
   Message(Corps, 42) = "The STOP button was pressed, operation canceled."
   Message(Titre, 42) = "Emergency stop"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 43)
   Message(Corps, 43) = "Caution, the wire is not at home position"
   Message(Titre, 43) = "Stop in the working area"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 44)
   Message(Corps, 44) = "The wire is at home position."
   Message(Titre, 44) = "Return made"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 45)
   Message(Corps, 45) = "A switch is open, the cut can not start." & vbCrLf & _
                        "Clear all switches by turning the threaded rods manually."
   Message(Titre, 45) = "Cannot make the cut."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 46)
   Message(Corps, 46) = "A limit switch is open. This is unusual, clear the workpiece and repeat the homing procedure."
   Message(Titre, 46) = "Stop during cutting!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 47)
   Message(Corps, 47) = "A endstop switch is open. This is unusual, clear the workpiece and repeat the origin procedure."
   Message(Titre, 47) = " Stop during the cutting!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 48)
   Message(Corps, 48) = "The emergency stop button was pressed, cancellation cutting."
   Message(Titre, 48) = "Stop during cutting"
   '***
   ReDim Preserve Message(Corps To Titre, 1 To 49)
'   Message(Corps, 49) = "Caution, the heater is currently set to another value as the preset of the selected material" & vbCrLf & _
'                        "Before saving, should I replace the preset for this material by current value heater?"
'   Message(Titre, 49) = "Save the parameters of the material used"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 50)
   Message(Corps, 50) = "Copy canceled"
   Message(Titre, 50) = "Copying Files"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 51)
   Message(Corps, 51) = "The files have been copied to the requested location"
   Message(Titre, 51) = "Copying Files"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 52)
   Message(Corps, 52) = "The working area of the machine is exceeded, the project is automatically truncated."
   Message(Titre, 52) = "Exceeding machine size"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 53)
   Message(Corps, 53) = "Outline not transferred in the block."
   Message(Titre, 53) = "Operation not possible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 54)
   Message(Corps, 54) = "Operation not possible."
   Message(Titre, 54) = "Cancel"
   
   ReDim Preserve Message(Corps To Titre, 1 To 55)
   Message(Corps, 55) = "Right clic to zoom location" & vbCrLf & _
                        "then mouse wheel" & vbCrLf & _
                        "or up and down arrows keys."
   Message(Titre, 55) = "Zoom"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 56)
   Message(Corps, 56) = " The speed value with a stored material exceeds the maximum value possible on the machine." _
                        & vbCrLf & "The speed is set to the maximum value."
   Message(Titre, 56) = "Excessive speed"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 57)
   Message(Corps, 57) = "The name of this material is already in memory (Expert mode). " & vbCrLf & _
                        "Overwrite values?"
   Message(Titre, 57) = "Changing the characteristic of a material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 58)
   Message(Corps, 58) = "There is a problem: the speed value for the selected material is not within the limits set."
   Message(Titre, 58) = "Incorrect value "
   '
   ReDim Preserve Message(Corps To Titre, 1 To 59)
   Message(Corps, 59) = "The speed of this material can not be changed. Create a new material."
   Message(Titre, 59) = "Impossible operation"
   '

   
   '**** Les Labels des outils, informations, avertissements ****
   ReDim Label(1 To 1)
   Label(1) = "Click on a point in a path to cut it in half."
   '
   ReDim Preserve Label(1 To 2)
   Label(2) = "Clic the new entry point."
   '
   ReDim Preserve Label(1 To 3)
   Label(3) = "Angle (°) :"
   '
   ReDim Preserve Label(1 To 4)
   Label(4) = "Ctrl=center, Shift=scale, Alt=mirror"
   '
   ReDim Preserve Label(1 To 5)
   Label(5) = "Selection : "
   '
   ReDim Preserve Label(1 To 6)
   Label(6) = " No visible outline. Select a file in the library. "
   '
   ReDim Preserve Label(1 To 7)
   Label(7) = " No outline beneath the pointer.. Double-click " & vbCrLf & "  an outline or drag it to the block. "
   '
   ReDim Preserve Label(1 To 8)
   Label(8) = " Outline not transferred in the block. Select a library file " & vbCrLf & "  and double-click on a shape or drag here. "
   '
   ReDim Preserve Label(1 To 9)
   Label(9) = "To X : "  ' il s'agit d'une dimensions suivant X, le mot à traduire est "Next"
   '
   ReDim Preserve Label(1 To 10)
   Label(10) = " mm - To Y : "  ' it is a size relative to Y
   '
   ReDim Preserve Label(1 To 11)
   Label(11) = " Automatic return to the home position "
   '
   ReDim Preserve Label(1 To 12)
   Label(12) = "Search switches"
   '
   ReDim Preserve Label(1 To 13)
   Label(13) = "Shift to the home position"
   '
   ReDim Preserve Label(1 To 14)
   Label(14) = "Home position"
   '
   ReDim Preserve Label(1 To 15)
   Label(15) = "Preparation for storage"
   '
   ReDim Preserve Label(1 To 16)
   Label(16) = "Moving to storage position "
   '
   ReDim Preserve Label(1 To 17)
   Label(17) = "Home position"
   '
   ReDim Preserve Label(1 To 18)
   Label(18) = "THE WIRE IS MOVING"
   '
   ReDim Preserve Label(1 To 19)
   Label(19) = "THE WIRE IS WARMING"
   '
   ReDim Preserve Label(1 To 20)
   Label(20) = "THE WIRE IS GETTING TO TEMPERATURE"
   '
   ReDim Preserve Label(1 To 21)
   Label(21) = "which "
   '
   ReDim Preserve Label(1 To 22)
   Label(22) = " s. of put in temperature."
   '
   ReDim Preserve Label(1 To 23)
   Label(23) = "(heating at"
   '
   ReDim Preserve Label(1 To 24)
   Label(24) = "Going back to origin procedure"
   '
   ReDim Preserve Label(1 To 25)
   Label(25) = "Stop on segment n° "
   '
   ReDim Preserve Label(1 To 26)
   Label(26) = "Return to home position"
   '
   ReDim Preserve Label(1 To 27)
   Label(27) = "Home position"
   '
   ReDim Preserve Label(1 To 28)
   Label(28) = "Heating the wire"
   '
   ReDim Preserve Label(1 To 29)
   Label(29) = "Cutting segment n°"
   '
   ReDim Preserve Label(1 To 30)
   Label(30) = "Control wire"
   '
   ReDim Preserve Label(1 To 31)
   Label(31) = "speed at"
   '
   ReDim Preserve Label(1 To 32)
   Label(32) = "Duration :"
   '

End Sub
