VERSION 5.00
Begin VB.Form formLangue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choix de la langue"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3135
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optLangue 
      Caption         =   "Italiano"
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
      Top             =   1515
      Width           =   1980
   End
   Begin VB.CommandButton cmdValiderLangue 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2535
      Width           =   1155
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
      Index           =   2
      Left            =   540
      TabIndex        =   2
      Top             =   1080
      Width           =   1980
   End
   Begin VB.OptionButton optLangue 
      Caption         =   "English"
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
      Width           =   1905
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
      Width           =   1950
   End
End
Attribute VB_Name = "formLangue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Select Case strLangue
   Case "francais"
      optLangue(0).value = True
   Case "english"
      optLangue(1).value = True
   Case "deutsch"
      optLangue(2).value = True
   Case "italiano"
      optLangue(3).value = True
   End Select
End Sub

Private Sub cmdValiderLangue_Click()
   If optLangue(0).value = True Then
      strLangue = "francais"
      If Dir(App.Path & "\Aide_Complexes_fr.chm", vbNormal) = "" Then     'si le fichier n'a pas été installé, on prend l'ancien
         strFichierAide = "Aide_Complexes.chm"
      Else
         strFichierAide = "Aide_Complexes_fr.chm"
      End If
   ElseIf optLangue(1).value = True Then
      strLangue = "english"
      If Dir(App.Path & "\Aide_Complexes_en.chm", vbNormal) = "" Then     'si le fichier n'a pas été installé, on prend l'ancien
         strFichierAide = "Aide_Complexes.chm"
      Else
         strFichierAide = "Aide_Complexes_en.chm"
      End If
   ElseIf optLangue(2).value = True Then
      strLangue = "deutsch"
      If Dir(App.Path & "\Aide_Complexes_de.chm", vbNormal) = "" Then     'si le fichier n'a pas été installé, on prend l'ancien
         strFichierAide = "Aide_Complexes.chm"
      Else
         strFichierAide = "Aide_Complexes_de.chm"
      End If
   ElseIf optLangue(3).value = True Then
      strLangue = "italiano"
      If Dir(App.Path & "\Aide_Complexes_it.chm", vbNormal) = "" Then     'si le fichier n'a pas été installé, on prend l'ancien
         strFichierAide = "Aide_Complexes.chm"
      Else
         strFichierAide = "Aide_Complexes_it.chm"
      End If
   End If

   Call GestionLangue(strLangue)
   App.HelpFile = App.Path & "\" & strFichierAide
   
   EcritDansFichierIni "Parametres", "Langue", strLangue, App.Path & "\Complexes.ini"
   EcritDansFichierIni "Parametres", "Fichier_aide", strFichierAide, App.Path & "\Complexes.ini"

   Unload formLangue
   
End Sub

