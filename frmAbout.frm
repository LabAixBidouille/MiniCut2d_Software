VERSION 5.00
Begin VB.Form frmAbout 
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
         Picture         =   "frmAbout.frx":0000
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
         Picture         =   "frmAbout.frx":08CA
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
            Caption         =   "Fonctionnement Expert"
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
            Caption         =   "Fonctionnement Normal"
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
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Le code du logiciel MiniCut2d Software est mis à disposition selon les termes de la licence
' Creative Commons Attribution Pas d'Utilisation Commerciale Partage à l'Identique 3.0 France.
'Pour voir une copie de cette licence, ouvrez le fichier "Licence_MiniCut2d_Software.txt",
' ou visitez le site internet http://creativecommons.org/licenses/by-nc-sa/3.0/fr/,
' ou écrivez à Creative Commons, 444 Castro Street, Suite 900, Mountain View, California, 94041, USA.

Option Explicit

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

Private Sub cmdOK_Click()
   If optNormalExpert(1).Value = True Then
      ModeSoft = "Expert"
   Else
      ModeSoft = "Normal"
   End If
   EcritFichierIni "Logiciel", "Mode", ModeSoft
   Unload Me
End Sub

Private Sub Form_Load()
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

'*** Si on double-clique sur l'image de droite, on ouvre une fenêtre de paramètres ***
Private Sub imgIconeSoft_DblClick()
   frmParametres.Show vbModal
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
