Attribute VB_Name = "modTypeMachine"
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

Public Sub ParametresMachine()
   Select Case TypeMachine
   Case "MiniCut2d_v1.0"
      NomTable = "MC2d v10" '8 caract�res dans IPL5X
      Frequence = 50000
      ChauffeMaxi = 80 '%
      PWMmaxi = Round(255 * ChauffeMaxi / 100, 0) '204 ; 100%=255
      MmParTourXG = 1   'vis de 1mm/tour
      MmParTourYG = 1
      MmParTourXD = 1
      MmParTourYD = 1
      PasParTourMoteurXG = 200
      PasParTourMoteurYG = 200
      PasParTourMoteurXD = 200
      PasParTourMoteurYD = 200
      MicroPasXG = 8
      MicroPasYG = 8
      MicroPasXD = 8
      MicroPasYD = 8
      PasParTourXG = PasParTourMoteurXG * MicroPasXG 'soit 1600 pas/tour
      PasParTourYG = PasParTourMoteurYG * MicroPasYG 'soit 1600 pas/tour
      PasParTourXD = PasParTourMoteurXD * MicroPasXD 'soit 1600 pas/tour
      PasParTourYD = PasParTourMoteurYD * MicroPasYD 'soit 1600 pas/tour
      MargeFil = 2
      MargePlateau = 2
      MargeInterieureX = 3 'on garde des marges lors de l'ajustment au bloc
      MargeInterieureY = 3
      DelaiMoteursOff = 1
      MmToOriXG = 1  'd�calage du fil par rapport aux inters (mm)
      MmToOriYG = 1
      MmToOriXD = 1
      MmToOriYD = 1
      NbrPasToOriXG = CLng(PasParTourXG * MmToOriXG / MmParTourXG)
      NbrPasToOriYG = CLng(PasParTourYG * MmToOriYG / MmParTourYG)
      NbrPasToOriXD = CLng(PasParTourXD * MmToOriXD / MmParTourXD)
      NbrPasToOriYD = CLng(PasParTourYD * MmToOriYD / MmParTourYD)
      'd�finition des axes
      TypeAxeInterpolateur(1) = &H5  'pour que la salve data se calcule correctement, l'interpolateur 1 doit repr�senter XL
      TypeAxeInterpolateur(2) = &H9  'pour que la salve data se calcule correctement, l'interpolateur 2 doit repr�senter YL
      TypeAxeInterpolateur(3) = &H6  'pour que la salve data se calcule correctement, l'interpolateur 3 doit repr�senter XR (ne sera pas utilis�)
      TypeAxeInterpolateur(4) = &HA  'pour que la salve data se calcule correctement, l'interpolateur 4 doit repr�senter YR
      TypeAxeInterpolateur(5) = &H0  'non attribu�
      DefinitionSortie(1) = BaseToBase("10000001", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 001(on fait sortir les steps de l'interpolateur 1)
      DefinitionSortie(2) = BaseToBase("00000001", 2, 10) ' 0(dir) 1(signal normal) 0(na) 0(na) 0(na) 001(on fait sortir les dir de l'interpolateur 1)
      DefinitionSortie(3) = 0   'non attribu�
      DefinitionSortie(4) = 0   'non attribu�
      DefinitionSortie(5) = 0   'non attribu�
      DefinitionSortie(6) = 0   'non attribu�
      DefinitionSortie(7) = BaseToBase("10000010", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 010(on fait sortir les steps de l'interpolateur 2)
      DefinitionSortie(8) = BaseToBase("01000010", 2, 10) ' 0(dir) 0(signal invers�) 0(na) 0(na) 0(na) 010(on fait sortir les dir invers�s de l'interpolateur 2)
      DefinitionSortie(9) = BaseToBase("10000100", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 100(on fait sortir les step de l'interpolateur 4)
      DefinitionSortie(10) = BaseToBase("01000100", 2, 10) ' 0(dir) 0(signal invers�) 0(na) 0(na) 0(na) 100(on fait sortir les dir invers�s de l'interpolateur 4)
     'codage en dur des vitesses et des courses ; l'utilisateur n'y a pas acc�s
      CourseX = 304
      CourseY = 266
      PenteAcceleration = 7
      VMaxSansAcc = 5
      VMaxAvecAcc = 8
      VitesseDecoupe = 4   '� ne pas confondre avec VMaxSansAcc
      VitesseRapide = 8   '� ne pas confondre avec VMaxAvecAcc
      With frmMiniCut2d
         .Caption = "MiniCut2d Software"
         .cmdPlierLePortique.Enabled = True
      End With
   Case "MiniCut2d_v1.2"
      NomTable = "MC2d v12" '8 caract�res dans IPL5X
      Frequence = 50000
      ChauffeMaxi = 70 '%
      PWMmaxi = Round(255 * ChauffeMaxi / 100, 0) '204 ; 100%=255
      MmParTourXG = 1   'vis de 1mm/tour
      MmParTourYG = 1
      MmParTourXD = 1
      MmParTourYD = 1
      PasParTourMoteurXG = 200
      PasParTourMoteurYG = 200
      PasParTourMoteurXD = 200
      PasParTourMoteurYD = 200
      MicroPasXG = 8
      MicroPasYG = 8
      MicroPasXD = 8
      MicroPasYD = 8
      PasParTourXG = PasParTourMoteurXG * MicroPasXG 'soit 1600 pas/tour
      PasParTourYG = PasParTourMoteurYG * MicroPasYG 'soit 1600 pas/tour
      PasParTourXD = PasParTourMoteurXD * MicroPasXD 'soit 1600 pas/tour
      PasParTourYD = PasParTourMoteurYD * MicroPasYD 'soit 1600 pas/tour
      MargeFil = 2
      MargePlateau = 2
      MargeInterieureX = 3 'on garde des marges lors de l'ajustment au bloc
      MargeInterieureY = 3
      DelaiMoteursOff = 1
      MmToOriXG = 1  'd�calage du fil par rapport aux inters (mm)
      MmToOriYG = 1
      MmToOriXD = 1
      MmToOriYD = 1
      NbrPasToOriXG = CLng(PasParTourXG * MmToOriXG / MmParTourXG)
      NbrPasToOriYG = CLng(PasParTourYG * MmToOriYG / MmParTourYG)
      NbrPasToOriXD = CLng(PasParTourXD * MmToOriXD / MmParTourXD)
      NbrPasToOriYD = CLng(PasParTourYD * MmToOriYD / MmParTourYD)
      'd�finition des axes
      TypeAxeInterpolateur(1) = &H5  'pour que la salve data se calcule correctement, l'interpolateur 1 doit repr�senter XL
      TypeAxeInterpolateur(2) = &H9  'pour que la salve data se calcule correctement, l'interpolateur 2 doit repr�senter YL
      TypeAxeInterpolateur(3) = &H6  'pour que la salve data se calcule correctement, l'interpolateur 3 doit repr�senter XR (ne sera pas utilis�)
      TypeAxeInterpolateur(4) = &HA  'pour que la salve data se calcule correctement, l'interpolateur 4 doit repr�senter YR
      TypeAxeInterpolateur(5) = &H0  'non attribu�
      DefinitionSortie(1) = BaseToBase("10000001", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 001(on fait sortir les steps de l'interpolateur 1)
      DefinitionSortie(2) = BaseToBase("00000001", 2, 10) ' 0(dir) 1(signal normal) 0(na) 0(na) 0(na) 001(on fait sortir les dir de l'interpolateur 1)
      DefinitionSortie(3) = 0   'non attribu�
      DefinitionSortie(4) = 0   'non attribu�
      DefinitionSortie(5) = 0   'non attribu�
      DefinitionSortie(6) = 0   'non attribu�
      DefinitionSortie(7) = BaseToBase("10000010", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 010(on fait sortir les steps de l'interpolateur 2)
      DefinitionSortie(8) = BaseToBase("01000010", 2, 10) ' 0(dir) 0(signal invers�) 0(na) 0(na) 0(na) 010(on fait sortir les dir invers�s de l'interpolateur 2)
      DefinitionSortie(9) = BaseToBase("10000100", 2, 10) ' 1(step) 0(signal normal) 0(na) 0(na) 0(na) 100(on fait sortir les step de l'interpolateur 4)
      DefinitionSortie(10) = BaseToBase("01000100", 2, 10) ' 0(dir) 0(signal invers�) 0(na) 0(na) 0(na) 100(on fait sortir les dir invers�s de l'interpolateur 4)
     'codage en dur des vitesses et des courses ; l'utilisateur n'y a pas acc�s
      CourseX = 304
      CourseY = 266
      PenteAcceleration = 7
      VMaxSansAcc = 5
      VMaxAvecAcc = 8
      VitesseDecoupe = 4   '� ne pas confondre avec VMaxSansAcc
      VitesseRapide = 8   '� ne pas confondre avec VMaxAvecAcc
      With frmMiniCut2d
         .Caption = "MiniCut2d Software"
         .cmdPlierLePortique.Enabled = True
      End With
      

   End Select
   'On met le bouton radio de l'�cran cach� des param�tres sur la bonne valeur
   Select Case TypeMachine
   Case "MiniCut2d_v1.0"
      frmParametres.optTypeMachine(0).Value = True
   Case "MiniCut2d_v1.2"
      frmParametres.optTypeMachine(1).Value = True
   End Select

End Sub
