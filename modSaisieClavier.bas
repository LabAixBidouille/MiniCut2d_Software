Attribute VB_Name = "modSaisieClavier"
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

'***************************************************************
'***** GESTION de l'AFFICHAGE des txtbox de saisie clavier *****
'***************************************************************

Option Explicit

Public Sub DesactiverMesures()
   If OutilEnCours = Deplacer Then
      Call DesactiverDeplacer
   ElseIf OutilEnCours = Tourner Then
      Call DesactiverTourner
   ElseIf OutilEnCours = Etirer Then
      Call DesactiverEtirer
   ElseIf OutilEnCours = Mesurer Then
      Call DesactiverMesurer
   ElseIf OutilEnCours = CouperProfil Then
      Call DesactiverCouperProfil
   ElseIf OutilEnCours = PointNumero1 Then
      Call DesactiverPointNumero1
   End If
End Sub

Public Sub ActiverMesures()
   If OutilEnCours = Deplacer Then
      Call ActiverDeplacer
   ElseIf OutilEnCours = Tourner Then
      Call ActiverTourner
   ElseIf OutilEnCours = Etirer Then
      Call ActiverEtirer
   ElseIf OutilEnCours = Mesurer Then
      Call ActiverMesurer
   ElseIf OutilEnCours = CouperProfil Then
      Call ActiverCouperProfil
   ElseIf OutilEnCours = PointNumero1 Then
      Call ActiverPointNumero1
   End If
End Sub

Public Sub MasquerMesures()   'tout cacher
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Visible = False
      .lblMesures(1).Visible = False
      .txtMesures(0).Visible = False
      .txtMesures(1).Visible = False
   End With
End Sub

Public Sub DesactiverMesurer()   'lorsque l'on a s�lectionn� l'outil m�tre
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Caption = "0.0 mm"
      .lblMetre.Visible = True
      .lblMetre.Enabled = False
   End With
End Sub

Public Sub ActiverMesurer()   'lorsque l'on a s�lectionn� l'outil m�tre
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Caption = "0.0 mm"
      .lblMetre.Visible = True
      .lblMetre.Enabled = True
   End With
End Sub

Public Sub DesactiverCouperProfil()   'lorsque l'on a s�lectionn� l'outil couper profil
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Caption = Label(1)  'cliquez sur un trajet pour le couper en deux
      .lblMetre.Visible = True
      .lblMetre.Enabled = False
   End With
End Sub

Public Sub ActiverCouperProfil()   'lorsque l'on a s�lectionn� l'outil couper profil
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Left = 165
      .lblMetre.Caption = Label(1)
      .lblMetre.Visible = True
      .lblMetre.Enabled = True
   End With
End Sub

Public Sub DesactiverPointNumero1()   'lorsque l'on a s�lectionn� l'outil couper profil
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Left = 165
      .lblMetre.Caption = Label(2) 's�lection par clic ou cadre (+ ctrl)
      .lblMetre.Visible = True
      .lblMetre.Enabled = False
   End With
End Sub

Public Sub ActiverPointNumero1()   'lorsque l'on a s�lectionn� l'outil couper profil
   Call MasquerMesures
   With frmMiniCut2d
      .lblMetre.Left = 165
      .lblMetre.Caption = Label(2) 's�lection par clic ou cadre (+ ctrl)
      .lblMetre.Visible = True
      .lblMetre.Enabled = True
   End With
End Sub

Public Sub DesactiverDeplacer()
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Caption = "X (mm) :"
      .lblMesures(0).Enabled = False
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = "Y (mm) :"
      .lblMesures(1).Enabled = False
      .lblMesures(1).Visible = True
      .txtMesures(0).Text = Format(0, "##0.0")
      .txtMesures(0).Enabled = False
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = Format(0, "##0.0")
      .txtMesures(1).Enabled = False
      .txtMesures(1).Visible = True
   End With
End Sub

Public Sub ActiverDeplacer()
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Caption = "X (mm) :"
      .lblMesures(0).Enabled = True
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = "Y (mm) :"
      .lblMesures(1).Enabled = True
      .lblMesures(1).Visible = True
      .txtMesures(0).Text = Format(0, "##0.0")
      .txtMesures(0).Enabled = True
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = Format(0, "##0.0")
      .txtMesures(1).Enabled = True
      .txtMesures(1).Visible = True
   End With
End Sub

Public Sub DesactiverTourner()
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Caption = Label(3) ' "Angle (�) :"
      .lblMesures(0).Enabled = False
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = ""
      .lblMesures(1).Enabled = False
      .lblMesures(1).Visible = False
      .txtMesures(0).Text = Format(0, "##0.0")
      .txtMesures(0).Enabled = False
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = ""
      .txtMesures(1).Enabled = False
      .txtMesures(1).Visible = False
   End With
End Sub

Public Sub ActiverTourner()
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Caption = Label(3) ' "Angle (�) :"
      .lblMesures(0).Enabled = True
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = ""
      .lblMesures(1).Enabled = False
      .lblMesures(1).Visible = False
      .txtMesures(0).Text = Format(0, "##0.0")
      .txtMesures(0).Enabled = True
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = ""
      .txtMesures(1).Enabled = False
      .txtMesures(1).Visible = False
   End With
End Sub

Public Sub DesactiverEtirer()
   With frmMiniCut2d
      .lblMetre.Visible = False
      .lblMesures(0).Caption = "X (%) :"
      .lblMesures(0).Enabled = False
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = "Y (%) :"
      .lblMesures(1).Enabled = False
      .lblMesures(1).Visible = True
      .txtMesures(0).Text = Format(100, "##0.0")
      .txtMesures(0).Enabled = False
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = Format(100, "##0.0")
      .txtMesures(1).Enabled = False
      .txtMesures(1).Visible = True
   End With
End Sub

Public Sub ActiverEtirer()
   With frmMiniCut2d
      If NbTransfSel > 0 Then
         .lblMetre.Left = 4500
         .lblMetre.Caption = Label(4)  '  "Ctrl=centrer, Shift=proportion, Alt=miroir"
         .lblMetre.Visible = True
      Else
         .lblMetre.Visible = False
      End If
      .lblMesures(0).Caption = "X (%) :"
      .lblMesures(0).Enabled = True
      .lblMesures(0).Visible = True
      .lblMesures(1).Caption = "Y (%) :"
      .lblMesures(1).Enabled = True
      .lblMesures(1).Visible = True
      .txtMesures(0).Text = Format(100, "##0.0")
      .txtMesures(0).Enabled = True
      .txtMesures(0).Visible = True
      .txtMesures(1).Text = Format(100, "##0.0")
      .txtMesures(1).Enabled = True
      .txtMesures(1).Visible = True
   End With
End Sub

'************************************************
'********* Utilisation de la touche TAB *********
'************************************************
Public Function GetTabState() As Boolean
    GetTabState = False
    If GetKeyState(VK_TAB) And -256 Then
        GetTabState = True
    End If
End Function
