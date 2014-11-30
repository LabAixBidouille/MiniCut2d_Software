VERSION 5.00
Begin VB.Form frmImages 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   Icon            =   "frmImages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgSettings 
      Height          =   255
      Left            =   4140
      Picture         =   "frmImages.frx":151A
      Top             =   1620
      Width           =   270
   End
   Begin VB.Image imgSettingsExpert 
      Height          =   255
      Left            =   4140
      Picture         =   "frmImages.frx":1AFC
      Top             =   825
      Width           =   270
   End
   Begin VB.Image imgDrapeauAmericain 
      Height          =   150
      Left            =   2355
      Picture         =   "frmImages.frx":20DE
      Top             =   2265
      Width           =   225
   End
   Begin VB.Image imgDrapeauItalien 
      Height          =   150
      Left            =   1665
      Picture         =   "frmImages.frx":25F0
      Top             =   2235
      Width           =   225
   End
   Begin VB.Image imgDrapeauAllemand 
      Height          =   150
      Left            =   3435
      Picture         =   "frmImages.frx":2B02
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image imgDrapeauEspagnol 
      Height          =   150
      Left            =   2730
      Picture         =   "frmImages.frx":3014
      Top             =   1740
      Width           =   225
   End
   Begin VB.Image imgDrapeauAnglais 
      Height          =   150
      Left            =   2205
      Picture         =   "frmImages.frx":3526
      Top             =   1875
      Width           =   225
   End
   Begin VB.Image imgDrapeauFrancais 
      Height          =   150
      Left            =   1740
      Picture         =   "frmImages.frx":3A38
      Top             =   1860
      Width           =   225
   End
   Begin VB.Image imgFeuVert 
      Height          =   615
      Left            =   1005
      Picture         =   "frmImages.frx":3F4A
      Top             =   1785
      Width           =   210
   End
   Begin VB.Image imgFeuRouge 
      Height          =   615
      Left            =   270
      Picture         =   "frmImages.frx":46C8
      Top             =   1845
      Width           =   210
   End
   Begin VB.Image imgStopSimulation 
      Height          =   330
      Left            =   1710
      Picture         =   "frmImages.frx":4E46
      Top             =   1155
      Width           =   930
   End
   Begin VB.Image imgSimuler 
      Height          =   330
      Left            =   180
      Picture         =   "frmImages.frx":58C0
      Top             =   1140
      Width           =   930
   End
   Begin VB.Image imgUndoGris 
      Height          =   480
      Left            =   3030
      Picture         =   "frmImages.frx":633A
      Top             =   420
      Width           =   480
   End
   Begin VB.Image imgRedo 
      Height          =   480
      Left            =   2040
      Picture         =   "frmImages.frx":6C04
      Top             =   450
      Width           =   480
   End
   Begin VB.Image imgUndo 
      Height          =   480
      Left            =   1350
      Picture         =   "frmImages.frx":74CE
      Top             =   450
      Width           =   480
   End
   Begin VB.Image imgRetrecir 
      Height          =   360
      Left            =   750
      Picture         =   "frmImages.frx":7D98
      Top             =   480
      Width           =   360
   End
   Begin VB.Image imgAgrandir 
      Height          =   360
      Left            =   150
      Picture         =   "frmImages.frx":8482
      Top             =   510
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Form de stockage des images"
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmImages"
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

