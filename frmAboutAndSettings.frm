VERSION 5.00
Begin VB.Form frmAboutAndSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "À propos de MiniCut2D Software"
   ClientHeight    =   5820
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10365
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4017.067
   ScaleMode       =   0  'User
   ScaleWidth      =   9733.271
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmAPropos 
      Caption         =   "A propos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   119
      TabIndex        =   3
      Top             =   180
      Width           =   4545
      Begin VB.PictureBox pctIconeSoft 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   180
         Picture         =   "frmAboutAndSettings.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   405
         Width           =   480
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000C&
         X1              =   282
         X2              =   4215
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   285
         X2              =   4215
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   9
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Titre de l'application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1050
         TabIndex        =   8
         Top             =   390
         Width           =   1635
      End
      Begin VB.Label lblSitesInternet 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "www.minicut2d.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1395
         TabIndex        =   7
         Top             =   1605
         Width           =   1680
      End
      Begin VB.Image imgIconeSoft 
         Height          =   480
         Left            =   3615
         Picture         =   "frmAboutAndSettings.frx":08CA
         Top             =   405
         Width           =   480
      End
      Begin VB.Label lblTraduction 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Traduction :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1650
         TabIndex        =   6
         Top             =   2910
         Width           =   975
      End
      Begin VB.Label lblSitesInternet 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "www.filchaud.com / www.frenchfoam.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   435
         TabIndex        =   5
         Top             =   2040
         Width           =   3420
      End
   End
   Begin VB.Frame frmParametres 
      Caption         =   "Paramètres"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   4845
      TabIndex        =   1
      Top             =   180
      Width           =   5370
      Begin VB.Frame frmModeExpert 
         Caption         =   "Fonctionnement Normal / Expert"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2520
         Left            =   210
         TabIndex        =   10
         Top             =   2055
         Width           =   4980
         Begin VB.OptionButton optNormalExpert 
            Caption         =   "Mode Expert"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   1
            Left            =   795
            TabIndex        =   13
            Top             =   1515
            Width           =   3525
         End
         Begin VB.OptionButton optNormalExpert 
            Caption         =   "Mode Normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   795
            TabIndex        =   12
            Top             =   1065
            Value           =   -1  'True
            Width           =   3165
         End
         Begin VB.CheckBox chkActiverLeChangementDeMode 
            Caption         =   "Activer le changement de mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   285
            TabIndex        =   11
            Top             =   375
            Width           =   4455
         End
      End
      Begin VB.Label lblParametres 
         Caption         =   "Courses, Décalages"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   210
         TabIndex        =   2
         Top             =   375
         Width           =   5100
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4380
      TabIndex        =   0
      Top             =   5250
      Width           =   1605
   End
End
Attribute VB_Name = "frmAboutAndSettings"
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

Private Sub Form_Load()
   ModeSoftTemp = ModeSoft
   flagFenetreChargee = False
   Select Case ModeSoft
   Case "Expert"
      optNormalExpert(1).Value = True
      optNormalExpert(0).Enabled = True
      optNormalExpert(1).Enabled = True
      chkActiverLeChangementDeMode.Value = vbChecked
   Case Else  'mode normal
      optNormalExpert(0).Value = True
      optNormalExpert(0).Enabled = False
      optNormalExpert(1).Enabled = False
      chkActiverLeChangementDeMode.Value = vbUnchecked
   End Select
   flagFenetreChargee = True
End Sub

Private Sub chkActiverLeChangementDeMode_Click()
   If flagFenetreChargee = True Then
      If chkActiverLeChangementDeMode.Value = vbChecked Then
         optNormalExpert(0).Enabled = True
         optNormalExpert(1).Enabled = True
      ElseIf chkActiverLeChangementDeMode.Value = vbUnchecked Then
         optNormalExpert(0).Enabled = False
         optNormalExpert(1).Enabled = False
      End If
   End If
End Sub

Private Sub optNormalExpert_Click(Index As Integer)
   If flagFenetreChargee = True Then
      Select Case Index
      Case 0
      Case 1
         MsgBox "Attention, le mode Expert permet de faire varier la vitesse." & vbCrLf & _
               "Vos paramètres de chauffe devrot être ajustés en conséquence." & vbCrLf & _
               "Ce mode n'est pas destiné aux débutants." & vbCrLf & _
               "En choisissant ce mode vous déclarez êtes suffisamment compétent" & vbCrLf & _
               "pour vous en servir sans dommage pour la machine de découpe."
      End Select
   End If
End Sub

Private Sub cmdOK_Click()
   Dim i As Long
   
   If optNormalExpert(1).Value = True Then
      ModeSoft = "Expert"
      frmMiniCut2d.comboMatieres.Text = frmMiniCut2d.comboMatieres.List(0)
   Else
      ModeSoft = "Normal"
      frmMiniCut2d.comboMatieres.Text = frmMiniCut2d.comboMatieres.List(0)
      frmMiniCut2d.hscVitesseBDD.Value = 40 'en mode normal, on est à 4mm/s, soit 40/10
      frmMiniCut2d.hscVitesseManuel.Value = 40 'en mode normal, on est à 4mm/s, soit 40/10
   End If
   EcritFichierIni "Logiciel", "Mode", ModeSoft
   If ModeSoft <> ModeSoftTemp Then
      Call AffichageEnFonctionDuModeSoft(ModeSoft)
      Call ListerMatieresDuIni  'réinitialisation du combobox
   End If
   Unload Me
End Sub

