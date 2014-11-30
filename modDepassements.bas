Attribute VB_Name = "modDepassements"
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

'************** Test de dépassement des courses de la machine et recalcul de la découpe *************

Public Sub CalculDepassementCourses(ByRef SequAVerifier As Sequ)
   'calcul des intersections en cas de dépassement des courses
   Dim XA As Single, YA As Single, XB As Single, YB As Single, Xinter As Single, Yinter As Single
   Dim Xentree As Single, Yentree As Single, Xsortie As Single, Ysortie As Single
   Dim StatutIntersection As Integer
   Dim flagSortieDesCourses As Boolean
   
   Dim i As Integer, j As Integer
   
   flagDepassementDecoupe = False
   With SequAVerifier
      'on teste l'intersection des segments avec la limite gauche :
      XA = MargeFil
      YA = MargePlateau
      XB = MargeFil
      YB = CourseY - MargeFil
      i = 1
      Do
         'la fonction renvoie 0 si // ou colinéaires, 1 si intersection sur les segments, 2 intersection hors des segments,
         '3 si intersection sur A, 4 si intersection sur B, 5 si intersection sur C, 6 si intersection sur D
         StatutIntersection = Intersection(XA, YA, XB, YB, .Point(i).x, .Point(i).y, _
                           .Point(i + 1).x, .Point(i + 1).y, Xinter, Yinter)
         If StatutIntersection = 1 Then 'si intersection, on rajoute le point
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = Xinter
            .Point(i + 1).y = Yinter
            .Point(i + 1).Etat = 1        'codage des états : 1 à gauche, 2 en haut, 3 à droite, 4 en bas
            .NbPoints = .NbPoints + 1
         End If
         i = i + 1
         If i = .NbPoints Then Exit Do
      Loop
      'on teste l'intersection des segments avec la limite droite :
      XA = CourseX
      YA = MargePlateau
      XB = CourseX
      YB = CourseY - MargeFil
      i = 1
      Do
         StatutIntersection = Intersection(XA, YA, XB, YB, .Point(i).x, .Point(i).y, _
                           .Point(i + 1).x, .Point(i + 1).y, Xinter, Yinter)
         If StatutIntersection = 1 Then '  si intersection, on rajoute le point
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = Xinter
            .Point(i + 1).y = Yinter
            .Point(i + 1).Etat = 3        'codage des états : 1 à gauche, 2 en haut, 3 à droite, 4 en bas
            .NbPoints = .NbPoints + 1
         End If
         i = i + 1
         If i = .NbPoints Then Exit Do
      Loop
      'on teste l'intersection des segments avec la limite basse :
      XA = MargeFil
      YA = MargePlateau
      XB = CourseX
      YB = MargePlateau
      i = 1
      Do
         StatutIntersection = Intersection(XA, YA, XB, YB, .Point(i).x, .Point(i).y, _
                           .Point(i + 1).x, .Point(i + 1).y, Xinter, Yinter)
         If StatutIntersection = 1 Then 'si intersection, on rajoute le point
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = Xinter
            .Point(i + 1).y = Yinter
            .Point(i + 1).Etat = 4        'codage des états : 1 à gauche, 2 en haut, 3 à droite, 4 en bas
            .NbPoints = .NbPoints + 1
         End If
         i = i + 1
         If i = .NbPoints Then Exit Do
      Loop
      'on teste l'intersection des segments avec la limite haute :
      XA = MargeFil
      YA = CourseY - MargeFil
      XB = CourseX
      YB = CourseY - MargeFil
      i = 1
      Do
         StatutIntersection = Intersection(XA, YA, XB, YB, .Point(i).x, .Point(i).y, _
                           .Point(i + 1).x, .Point(i + 1).y, Xinter, Yinter)
         If StatutIntersection = 1 Then '  si intersection, on rajoute le point
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = Xinter
            .Point(i + 1).y = Yinter
            .Point(i + 1).Etat = 2        'codage des états : 1 à gauche, 2 en haut, 3 à droite, 4 en bas
            .NbPoints = .NbPoints + 1
         End If
         i = i + 1
         If i = .NbPoints Then Exit Do
      Loop

      '***** On supprime les points à l'extérieur des limites de la table *****
      i = 0
      Do
         i = i + 1
         If .Point(i).x < MargeFil Or .Point(i).x > CourseX Or .Point(i).y < 1 Or .Point(i).y > CourseY - MargeFil Then
            flagDepassementDecoupe = True
            For j = i To .NbPoints - 1
               .Point(j) = .Point(j + 1)
            Next j
            ReDim Preserve .Point(1 To .NbPoints - 1)
            .NbPoints = .NbPoints - 1
            i = i - 1
         End If
         If i = .NbPoints Then Exit Do
      Loop
      
      'on gère les angles
      i = 0
      Do
         i = i + 1
         If (.Point(i).Etat = 1 And .Point(i + 1).Etat = 2) Or (.Point(i).Etat = 2 And .Point(i + 1).Etat = 1) Then  'passage GH (gauche <-> haut)
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = MargeFil
            .Point(i + 1).y = CourseY - MargeFil
            .NbPoints = .NbPoints + 1
            i = i + 1
         ElseIf (.Point(i).Etat = 2 And .Point(i + 1).Etat = 3) Or (.Point(i).Etat = 3 And .Point(i + 1).Etat = 2) Then  'passage HD
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = CourseX
            .Point(i + 1).y = CourseY - MargeFil
            .NbPoints = .NbPoints + 1
            i = i + 1
         ElseIf (.Point(i).Etat = 3 And .Point(i + 1).Etat = 4) Or (.Point(i).Etat = 4 And .Point(i + 1).Etat = 3) Then 'passage DB
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = CourseX
            .Point(i + 1).y = 1
            .NbPoints = .NbPoints + 1
            i = i + 1
         ElseIf (.Point(i).Etat = 4 And .Point(i + 1).Etat = 1) Or (.Point(i).Etat = 1 And .Point(i + 1).Etat = 4) Then 'passage BG
            ReDim Preserve .Point(1 To .NbPoints + 1)
            For j = .NbPoints To i + 1 Step -1
               .Point(j + 1) = .Point(j)
            Next j
            .Point(i + 1).x = MargeFil
            .Point(i + 1).y = 1
            .NbPoints = .NbPoints + 1
            i = i + 1
         End If
         If i = .NbPoints - 1 Then Exit Do
      Loop
   End With
End Sub



