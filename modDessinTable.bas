Attribute VB_Name = "modDessinTable"
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

'module de d�finition du dessin de l'onglet d�coupe et des maxi et mini

Public Sub TableauTraceMachine()
   Dim i As Integer
   
   Select Case TypeMachine
   Case "MiniCut2d_v1.0"
      '**************** RECTANGLES ************** valeurs en dimensions r�elles en mm
      ReDim RECT(1 To 1)
      NbRect = 0
      '*** la section suivante est � r�p�ter pour chaque rectangle ***
      'rectangles de fond du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = -70
         .X2 = 6.8
         .Y2 = 291
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangle sup�rieur du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = 291
         .X2 = 6.8
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangleS support plateau
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 112
         .Y1 = -57
         .X2 = 206
         .Y2 = -6
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 112
         .Y1 = -57
         .X2 = 118
         .Y2 = -6
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 200
         .Y1 = -57
         .X2 = 206
         .Y2 = -6
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangles verticaux du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -56
         .Y1 = -70
         .X2 = -50
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 6.8
         .Y1 = -70
         .X2 = 12.8
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangles du socle
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -69
         .Y1 = -70
         .X2 = 216
         .Y2 = -27
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 210
         .Y1 = -64
         .X2 = 216
         .Y2 = -27
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -69
         .Y1 = -70
         .X2 = 226
         .Y2 = -63
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangle du plateau
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -64.6
         .Y1 = -7
         .X2 = 381.4
         .Y2 = 0
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
Case "MiniCut2d_v1.2"
      '**************** RECTANGLES ************** valeurs en dimensions r�elles en mm
      ReDim RECT(1 To 1)
      NbRect = 0
      '*** la section suivante est � r�p�ter pour chaque rectangle ***
      'rectangles de fond du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = -70
         .X2 = 6.8
         .Y2 = 291
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangle sup�rieur du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = 291
         .X2 = 6.8
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangleS support plateau
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 112
         .Y1 = -57
         .X2 = 206
         .Y2 = -12
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangles verticaux du caisson des Y
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -56
         .Y1 = -70
         .X2 = -50
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 6.8
         .Y1 = -70
         .X2 = 12.8
         .Y2 = 297
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangles du socle
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -69
         .Y1 = -70
         .X2 = 216
         .Y2 = -19
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 11726588
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -69
         .Y1 = -70
         .X2 = 226
         .Y2 = -63
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'rectangle du plateau
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -64.6
         .Y1 = -7
         .X2 = 381.4
         .Y2 = 0
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -64.6
         .Y1 = -14
         .X2 = 381.4
         .Y2 = -7
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
   Case "MaxiCut2d"
      '**************** RECTANGLES **************
      ReDim RECT(1 To 1)
      NbRect = 0
      '*** la section suivante est � r�p�ter pour chaque rectangle ***
      'barre de soutien du plateau m�lamin�
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = -12
         .X2 = 1410
         .Y2 = -52
         .Rempli = True
         .CoulTour = 4276545
         .CoulFond = 11711154
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'plateau m�lamin� 12mm
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = 0
         .X2 = 1410
         .Y2 = -12
         .Rempli = True
         .CoulTour = 8157820
         .CoulFond = 7114419
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'montant gauche
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -90
         .Y1 = 778
         .X2 = -50
         .Y2 = -92
         .Rempli = True
         .CoulTour = 4276545
         .CoulFond = 11711154
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'montant droit
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = 1410
         .Y1 = 778
         .X2 = 1450
         .Y2 = -92
         .Rempli = True
         .CoulTour = 4276545
         .CoulFond = 11711154
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'barre basse
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -90
         .Y1 = -92
         .X2 = 1450
         .Y2 = -132
         .Rempli = True
         .CoulTour = 4276545
         .CoulFond = 11711154
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
      'barre haute
      NbRect = NbRect + 1
      ReDim Preserve RECT(1 To NbRect)
      With RECT(NbRect)
         .X1 = -50
         .Y1 = 778
         .X2 = 1410
         .Y2 = 738
         .Rempli = True
         .CoulTour = 4276545
         .CoulFond = 11711154
         .TypeTrait = vbSolid     'vbDash = tirets ; vbDot = pointill�s ; vbInsideSolid = interieur uni
      End With
   End Select
   
   Call MaxiMiniDecoupeTable  'recherche des extremums de la repr�sentation (ci-dessous)
End Sub

Public Sub MaxiMiniDecoupeTable()
   'D�finition des extr�mes des coordonn�es de repr�sentation de la table
   Dim i As Integer
   MaxiDecoupeX = 0
   MiniDecoupeX = 0
   MaxiDecoupeY = 0
   MiniDecoupeY = 0
   For i = 1 To NbRect
    If RECT(i).X1 >= MaxiDecoupeX Then MaxiDecoupeX = RECT(i).X1
    If RECT(i).X2 >= MaxiDecoupeX Then MaxiDecoupeX = RECT(i).X2
    If RECT(i).X1 <= MiniDecoupeX Then MiniDecoupeX = RECT(i).X1
    If RECT(i).X2 <= MiniDecoupeX Then MiniDecoupeX = RECT(i).X2
    If RECT(i).Y1 >= MaxiDecoupeY Then MaxiDecoupeY = RECT(i).Y1
    If RECT(i).Y2 >= MaxiDecoupeY Then MaxiDecoupeY = RECT(i).Y2
    If RECT(i).Y1 <= MiniDecoupeY Then MiniDecoupeY = RECT(i).Y1
    If RECT(i).Y2 <= MiniDecoupeY Then MiniDecoupeY = RECT(i).Y2
   Next i
'   For i = 1 To NbLignes
'    If Ligne(i).x1 >= MaxiDecoupeX Then MaxiDecoupeX = Ligne(i).x1
'    If Ligne(i).x2 >= MaxiDecoupeX Then MaxiDecoupeX = Ligne(i).x2
'    If Ligne(i).x1 <= MiniDecoupeX Then MiniDecoupeX = Ligne(i).x1
'    If Ligne(i).x2 <= MiniDecoupeX Then MiniDecoupeX = Ligne(i).x2
'    If Ligne(i).y1 >= MaxiDecoupeY Then MaxiDecoupeY = Ligne(i).y1
'    If Ligne(i).y2 >= MaxiDecoupeY Then MaxiDecoupeY = Ligne(i).y2
'    If Ligne(i).y1 <= MiniDecoupeY Then MiniDecoupeY = Ligne(i).y1
'    If Ligne(i).y2 <= MiniDecoupeY Then MiniDecoupeY = Ligne(i).y2
'   Next i
End Sub

