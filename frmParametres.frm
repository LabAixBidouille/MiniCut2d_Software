VERSION 5.00
Begin VB.Form frmParametres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type machine"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdValiderParametresMachine 
      Caption         =   "OK"
      Height          =   450
      Left            =   4320
      TabIndex        =   3
      Top             =   2310
      Width           =   1410
   End
   Begin VB.Frame frmTypeMachine 
      Caption         =   "Type"
      Height          =   1905
      Left            =   270
      TabIndex        =   0
      Top             =   240
      Width           =   5505
      Begin VB.OptionButton optTypeMachine 
         Caption         =   "MiniCut2d - v1.2"
         Height          =   405
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   870
         Width           =   2445
      End
      Begin VB.OptionButton optTypeMachine 
         Caption         =   "MiniCut2d - v1"
         Height          =   405
         Index           =   0
         Left            =   195
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmParametres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdValiderParametresMachine_Click()

   If optTypeMachine(0).Value = True Then
      TypeMachine = "MiniCut2d_v1.0"
   ElseIf optTypeMachine(1).Value = True Then
      TypeMachine = "MiniCut2d_v1.2"
   ElseIf optTypeMachine(2).Value = True Then
      TypeMachine = "MaxiCut2d"
   Else
      Me.Hide
      Exit Sub
   End If
   EcritFichierIni "Machine", "Type", TypeMachine
   Call ParametresMachine  'pour remplacement des variables et changement du titre de la fenêtre du soft
   CoeffBloc = 1
   Call frmMiniCut2d.InitialisationCoursesEtBloc
   Call TableauTraceMachine  ' pour redessin de la table dans l'onglet découpe
   
   If frmMiniCut2d.SSTab1.Tab = 1 Then
      Call frmMiniCut2d.InitialisationDessinDecoupe
   End If
   Call frmMiniCut2d.ZoomAutoToutVoir
   Call frmMiniCut2d.TraceTransf
   'Chargement de la table dans l'interface
   If IPL5X_IsConnected() <> 1 Then 'l'interpolateur n'est pas connecté
      flagTableEcriteDansIPL = False
      frmDecoupeInactive.Show vbModal
   Else  'l'interpolateur est connecté, on vérifie si la table a été mémorisée dedans, sinon on le fait
      If flagTableEcriteDansIPL = False Then
         MsgBox Message(Corps, 22), vbInformation, Message(Titre, 22)  'Initialisation de l'interface
         'Chargement de la table dans l'interface
         Call EcrireTable
         If ErrIPL <> 1 Then 'il s'est produit une erreur
            GoTo Erreur
         Else
            flagTableEcriteDansIPL = True
         End If
      End If
   End If
   Call GestionLangue(strLangue) 'pour modification des textes de la fenêtre "About"
   Me.Hide
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
End Sub

Private Sub Form_Load()
   'On met le bouton radio de l'écran caché des paramètres sur la bonne valeur
   Select Case TypeMachine
   Case "MiniCut2d_v1.0"
      optTypeMachine(0).Value = True
   Case "MiniCut2d_v1.2"
      optTypeMachine(1).Value = True
   Case "MaxiCut2d"
      optTypeMachine(2).Value = True
   End Select
End Sub
