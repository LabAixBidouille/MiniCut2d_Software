VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "MiniCut2d Software - Bienvenue !"
   ClientHeight    =   10260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPasDeMachine 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Si vous n'avez pas de machine, cliquez ici"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9000
      Width           =   5850
   End
   Begin VB.CommandButton cmdTypeMiniCut2d 
      BackColor       =   &H00FFFFFF&
      Height          =   7170
      Index           =   1
      Left            =   7395
      Picture         =   "frmSplashScreen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1395
      Width           =   5850
   End
   Begin VB.CommandButton cmdTypeMiniCut2d 
      BackColor       =   &H00FFFFFF&
      Height          =   7170
      Index           =   0
      Left            =   780
      Picture         =   "frmSplashScreen.frx":A5DA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1395
      Width           =   5850
   End
   Begin VB.PictureBox pctCadreSplash 
      Height          =   10080
      Left            =   90
      ScaleHeight     =   10020
      ScaleWidth      =   13815
      TabIndex        =   4
      Top             =   90
      Width           =   13875
   End
   Begin VB.Label lblCliquez 
      Alignment       =   2  'Center
      Caption         =   "Cliquez sur le type de machine que vous utilisez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1035
      TabIndex        =   2
      Top             =   405
      Width           =   11580
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub cmdPasDeMachine_Click()
   TypeMachine = "MiniCut2d_v1.2"
   EcritFichierIni "Machine", "Type", TypeMachine
   Me.Hide
End Sub

Private Sub cmdTypeMiniCut2d_Click(Index As Integer)
   Select Case Index
   Case 0
      TypeMachine = "MiniCut2d_v1.2"
   Case 1
      TypeMachine = "MiniCut2d_v1.0"
   End Select
   EcritFichierIni "Machine", "Type", TypeMachine
   Me.Hide
End Sub

