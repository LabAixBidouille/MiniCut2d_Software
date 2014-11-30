VERSION 5.00
Begin VB.Form frmLangue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix de la langue"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optLangue 
      Caption         =   "Espanol"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   540
      TabIndex        =   5
      Top             =   1920
      Width           =   1400
   End
   Begin VB.OptionButton optLangue 
      Caption         =   "Deutsch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   540
      TabIndex        =   4
      Top             =   1500
      Width           =   1400
   End
   Begin VB.CommandButton cmdValiderLangue 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3015
      Width           =   1155
   End
   Begin VB.OptionButton optLangue 
      Caption         =   "UK english"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   540
      TabIndex        =   2
      Top             =   1080
      Width           =   1400
   End
   Begin VB.OptionButton optLangue 
      Caption         =   "US english"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   645
      Width           =   1400
   End
   Begin VB.OptionButton optLangue 
      Caption         =   "Français"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   210
      Value           =   -1  'True
      Width           =   1400
   End
   Begin VB.Image imgDrapeauEspagnol 
      Height          =   150
      Left            =   2160
      Picture         =   "frmLangue.frx":0000
      Top             =   2055
      Width           =   225
   End
   Begin VB.Image imgDrapeauAllemand 
      Height          =   150
      Left            =   2160
      Picture         =   "frmLangue.frx":0512
      Top             =   1635
      Width           =   225
   End
   Begin VB.Image imgDrapeauAmericain 
      Height          =   150
      Left            =   2160
      Picture         =   "frmLangue.frx":0A24
      Top             =   780
      Width           =   225
   End
   Begin VB.Image imgDrapeauAnglais 
      Height          =   150
      Left            =   2160
      Picture         =   "frmLangue.frx":0F36
      Top             =   1215
      Width           =   225
   End
   Begin VB.Image imgDrapeauFrancais 
      Height          =   150
      Left            =   2160
      Picture         =   "frmLangue.frx":1448
      Top             =   345
      Width           =   225
   End
End
Attribute VB_Name = "frmLangue"
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

Private Sub cmdValiderLangue_Click()
   If optLangue(0).Value = True Then
      strLangue = "francais"
   ElseIf optLangue(1).Value = True Then
      strLangue = "USA"
   ElseIf optLangue(2).Value = True Then
      strLangue = "english"
   ElseIf optLangue(3).Value = True Then
      strLangue = "deutsch"
   ElseIf optLangue(4).Value = True Then
      strLangue = "espanol"
   End If
   Call GestionLangue(strLangue)
   EcritFichierIni "Parametres", "Langue", strLangue  'mémorisation de la langue dans le .ini
   Unload frmLangue
End Sub

Private Sub Form_Load()
   Select Case strLangue
   Case "francais"
      optLangue(0).Value = True
   Case "USA"
      optLangue(1).Value = True
   Case "english"
      optLangue(2).Value = True
   Case "deutsch"
      optLangue(3).Value = True
   Case "espanol"
      optLangue(4).Value = True
   End Select
End Sub
