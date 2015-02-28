VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMiniCut2d 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "MiniCut2d Software"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   15240
   Icon            =   "frmMiniCut2d.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   673
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pctZoomInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   14835
      Picture         =   "frmMiniCut2d.frx":058A
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   151
      Top             =   6930
      Width           =   300
   End
   Begin MSComctlLib.Slider sliderZoom 
      Height          =   1275
      Left            =   12555
      TabIndex        =   150
      Top             =   5475
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2249
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Min             =   -1
      Max             =   1
   End
   Begin VB.PictureBox pctRepriseDecoupe 
      Height          =   9075
      Left            =   10155
      ScaleHeight     =   9015
      ScaleWidth      =   3975
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   585
      Width           =   4035
      Begin VB.Frame frameAnnulationReprise 
         Caption         =   "Annuler"
         ForeColor       =   &H00004000&
         Height          =   1065
         Left            =   180
         TabIndex        =   100
         Top             =   7920
         Width           =   3645
         Begin VB.CommandButton cmdAnnulerReprise 
            BackColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":0BB4
            Style           =   1  'Graphical
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "Annuler tout"
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Frame frameRetourOrigine 
         Caption         =   "Retourner à l'origine"
         ForeColor       =   &H00004000&
         Height          =   2445
         Left            =   180
         TabIndex        =   93
         Top             =   3705
         Width           =   3645
         Begin VB.CommandButton cmdLancerRetourApresStop 
            BackColor       =   &H0080FF80&
            Height          =   615
            Left            =   165
            Picture         =   "frmMiniCut2d.frx":1656
            Style           =   1  'Graphical
            TabIndex        =   98
            TabStop         =   0   'False
            ToolTipText     =   "Lancer le retour"
            Top             =   1635
            Width           =   3315
         End
         Begin VB.Frame frameTrajetRetour 
            Caption         =   "Trajet"
            ForeColor       =   &H00004000&
            Height          =   1125
            Left            =   120
            TabIndex        =   94
            Top             =   345
            Width           =   3405
            Begin VB.OptionButton optTrajetRetour 
               BackColor       =   &H00C0FFC0&
               Height          =   585
               Index           =   2
               Left            =   2370
               Picture         =   "frmMiniCut2d.frx":26BC
               Style           =   1  'Graphical
               TabIndex        =   97
               TabStop         =   0   'False
               ToolTipText     =   "Par la haut"
               Top             =   330
               Width           =   855
            End
            Begin VB.OptionButton optTrajetRetour 
               BackColor       =   &H00C0FFC0&
               Height          =   585
               Index           =   1
               Left            =   1280
               Picture         =   "frmMiniCut2d.frx":2EA2
               Style           =   1  'Graphical
               TabIndex        =   96
               TabStop         =   0   'False
               ToolTipText     =   "Par la gauche"
               Top             =   330
               Width           =   855
            End
            Begin VB.OptionButton optTrajetRetour 
               BackColor       =   &H00C0FFC0&
               Height          =   585
               Index           =   0
               Left            =   180
               Picture         =   "frmMiniCut2d.frx":3688
               Style           =   1  'Graphical
               TabIndex        =   95
               TabStop         =   0   'False
               ToolTipText     =   "En diagonale"
               Top             =   330
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdStopRetourApresReprise 
            BackColor       =   &H0080FF80&
            Height          =   585
            Left            =   165
            Picture         =   "frmMiniCut2d.frx":3E6E
            Style           =   1  'Graphical
            TabIndex        =   99
            TabStop         =   0   'False
            ToolTipText     =   "Arrêt d'urgence!"
            Top             =   2220
            Width           =   3315
         End
      End
      Begin VB.Frame frameModifierLaChauffe 
         Caption         =   "Modifier la chauffe"
         ForeColor       =   &H00004000&
         Height          =   855
         Left            =   180
         TabIndex        =   90
         Top             =   2610
         Width           =   3645
         Begin VB.PictureBox imgChauffe2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   210
            Picture         =   "frmMiniCut2d.frx":48A0
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   126
            Top             =   375
            Width           =   375
         End
         Begin VB.HScrollBar hscChauffeStop 
            Height          =   255
            LargeChange     =   10
            Left            =   735
            Max             =   100
            Min             =   1
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   360
            Value           =   1
            Width           =   1845
         End
         Begin VB.Label lblChauffeStop 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "XXX %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   92
            Top             =   390
            Width           =   570
         End
      End
      Begin VB.Frame frameReprendre 
         Caption         =   "Terminer la découpe"
         ForeColor       =   &H00004000&
         Height          =   1170
         Left            =   180
         TabIndex        =   88
         Top             =   6465
         Width           =   3645
         Begin VB.CommandButton cmdRepriseDecoupe 
            BackColor       =   &H0080FF80&
            Height          =   705
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":4E3A
            Style           =   1  'Graphical
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   "Terminer la découpe"
            Top             =   270
            Width           =   3315
         End
      End
      Begin VB.Frame frameInformationStop 
         Caption         =   "Information"
         ForeColor       =   &H00004000&
         Height          =   1845
         Left            =   180
         TabIndex        =   84
         Top             =   480
         Width           =   3645
         Begin MSComctlLib.ProgressBar progrbarRetour 
            Height          =   255
            Left            =   90
            TabIndex        =   119
            Top             =   1470
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lblNumSegmentStop 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Arrêt sur le segment n° :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   780
            TabIndex        =   86
            Top             =   1110
            Width           =   2025
         End
         Begin VB.Label lblArretParBoutonStop 
            Alignment       =   2  'Center
            Caption         =   "Vous venez de demander l'arrêt de la découpe en appuyant/cliquant sur le bouton STOP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   60
            TabIndex        =   85
            Top             =   300
            Width           =   3465
         End
      End
      Begin VB.Label lblStopReprise 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Stop / Reprise  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1155
         TabIndex        =   87
         Top             =   120
         Width           =   1545
      End
   End
   Begin VB.PictureBox pctValidationDecoupe 
      Height          =   9075
      Left            =   6990
      ScaleHeight     =   9015
      ScaleWidth      =   3975
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1065
      Width           =   4035
      Begin VB.Frame frmDecalageExpert 
         Caption         =   "Décaler le fil (mode Expert)"
         Height          =   1095
         Left            =   180
         TabIndex        =   173
         Top             =   585
         Width           =   3645
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H00004000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   10
            Left            =   3100
            Style           =   1  'Graphical
            TabIndex        =   184
            ToolTipText     =   "+0.7mm (S=1.4mm)"
            Top             =   300
            Width           =   320
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   9
            Left            =   2770
            Style           =   1  'Graphical
            TabIndex        =   183
            ToolTipText     =   "+0.65mm (S=1.3mm)"
            Top             =   300
            Width           =   300
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   8
            Left            =   2450
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "+0.6mm (S=1.2mm)"
            Top             =   300
            Width           =   280
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0000FF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   7
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   181
            ToolTipText     =   "+0.55mm (S=1.1mm)"
            Top             =   300
            Width           =   260
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   6
            Left            =   1890
            Style           =   1  'Graphical
            TabIndex        =   180
            ToolTipText     =   "+0.5mm (S=1mm)"
            Top             =   300
            Width           =   240
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H00004000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   5
            Left            =   100
            Style           =   1  'Graphical
            TabIndex        =   179
            ToolTipText     =   "-0.7mm (S=1.4mm)"
            Top             =   300
            Width           =   320
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H00008000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   4
            Left            =   440
            Style           =   1  'Graphical
            TabIndex        =   178
            ToolTipText     =   "-0.65mm (S=1.3mm)"
            Top             =   300
            Width           =   300
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   3
            Left            =   770
            Style           =   1  'Graphical
            TabIndex        =   177
            ToolTipText     =   "-0.6mm (S=1.2mm)"
            Top             =   300
            Width           =   280
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0000FF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   2
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   176
            ToolTipText     =   "-0.55mm (S=1.1mm)"
            Top             =   300
            Width           =   260
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   1
            Left            =   1360
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "-0.5mm (S=1mm)"
            Top             =   300
            Width           =   240
         End
         Begin VB.OptionButton optDecalageExpert 
            BackColor       =   &H00C0FFC0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   0
            Left            =   1640
            Style           =   1  'Graphical
            TabIndex        =   174
            ToolTipText     =   "0mm"
            Top             =   300
            Value           =   -1  'True
            Width           =   220
         End
         Begin VB.Label lblDecalageExpert 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            Caption         =   "0mm"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   150
            TabIndex        =   185
            Top             =   765
            Width           =   3240
         End
      End
      Begin VB.Frame frameChauffeEnCoursDecoupe 
         Caption         =   "Chauffe"
         ForeColor       =   &H00004000&
         Height          =   1020
         Left            =   195
         TabIndex        =   144
         Top             =   7755
         Width           =   3645
         Begin VB.OptionButton optChauffePendantDecoupe 
            BackColor       =   &H00C0FFC0&
            Height          =   525
            Index           =   0
            Left            =   165
            Picture         =   "frmMiniCut2d.frx":6104
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   330
            Width           =   645
         End
         Begin VB.OptionButton optChauffePendantDecoupe 
            BackColor       =   &H00C0FFC0&
            Height          =   525
            Index           =   1
            Left            =   270
            Picture         =   "frmMiniCut2d.frx":67A6
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   420
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.HScrollBar hscChauffePendantDecoupe 
            Height          =   255
            LargeChange     =   10
            Left            =   1020
            Max             =   100
            Min             =   1
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   495
            Value           =   1
            Width           =   1635
         End
         Begin VB.PictureBox imgChauffe3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2925
            Picture         =   "frmMiniCut2d.frx":6EAC
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   145
            Top             =   225
            Width           =   375
         End
         Begin VB.Label lblChauffePendantDecoupe 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "XXX %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2865
            TabIndex        =   147
            Top             =   540
            Width           =   570
         End
      End
      Begin VB.Frame frameDecalage 
         Caption         =   "Décaler le fil"
         ForeColor       =   &H00004000&
         Height          =   1095
         Left            =   180
         TabIndex        =   103
         Top             =   1125
         Width           =   3645
         Begin VB.OptionButton optDecalage 
            BackColor       =   &H00C0FFC0&
            Height          =   585
            Index           =   3
            Left            =   2475
            Picture         =   "frmMiniCut2d.frx":7446
            Style           =   1  'Graphical
            TabIndex        =   106
            TabStop         =   0   'False
            ToolTipText     =   "0.5"
            Top             =   315
            Width           =   915
         End
         Begin VB.OptionButton optDecalage 
            BackColor       =   &H00C0FFC0&
            Height          =   585
            Index           =   2
            Left            =   1335
            Picture         =   "frmMiniCut2d.frx":7D10
            Style           =   1  'Graphical
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "0mm"
            Top             =   315
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optDecalage 
            BackColor       =   &H00C0FFC0&
            Height          =   585
            Index           =   1
            Left            =   195
            Picture         =   "frmMiniCut2d.frx":85DA
            Style           =   1  'Graphical
            TabIndex        =   104
            TabStop         =   0   'False
            ToolTipText     =   "-0.5mm"
            Top             =   315
            Width           =   915
         End
      End
      Begin VB.Frame frameAnnulationStop 
         Caption         =   "Annulation - Stop/Reprise"
         ForeColor       =   &H00004000&
         Height          =   1065
         Left            =   195
         TabIndex        =   80
         Top             =   6405
         Width           =   3645
         Begin VB.CommandButton cmdAnnulerDecoupe 
            BackColor       =   &H00C0FFC0&
            Height          =   585
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":8EA4
            Style           =   1  'Graphical
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "Annuler la découpe"
            Top             =   300
            Width           =   3315
         End
         Begin VB.CommandButton cmdSTOP 
            BackColor       =   &H0080FF80&
            Height          =   585
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":9946
            Style           =   1  'Graphical
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "Arrêt d'urgence!"
            Top             =   810
            Width           =   3315
         End
      End
      Begin VB.Frame frameAction 
         Caption         =   "Lancer la découpe"
         ForeColor       =   &H00004000&
         Height          =   1335
         Left            =   180
         TabIndex        =   74
         Top             =   4770
         Width           =   3645
         Begin VB.CommandButton cmdFaireOrigineAvantDecoupe 
            BackColor       =   &H0080FF80&
            Height          =   810
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":A378
            Style           =   1  'Graphical
            TabIndex        =   75
            TabStop         =   0   'False
            ToolTipText     =   "Lancer la découpe"
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Frame frameInformation 
         Caption         =   "Information"
         ForeColor       =   &H00004000&
         Height          =   2040
         Left            =   165
         TabIndex        =   70
         Top             =   2430
         Width           =   3645
         Begin MSComctlLib.ProgressBar progrbarChauffe 
            Height          =   255
            Left            =   90
            TabIndex        =   73
            Top             =   1635
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lblAvertissementDecoupe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Avertissement Découpe"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   765
            TabIndex        =   72
            Top             =   1245
            Width           =   2025
         End
         Begin VB.Label lblDureeDecoupe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Durée : xx min. yy s."
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
            Left            =   1020
            TabIndex        =   71
            Top             =   330
            Width           =   1605
         End
      End
      Begin VB.Label lblCadreDecoupe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Découpe "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1230
         TabIndex        =   69
         Top             =   195
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkZoomDecoupe 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":BE42
      Style           =   1  'Graphical
      TabIndex        =   143
      TabStop         =   0   'False
      ToolTipText     =   "Zoom Bloc"
      Top             =   45
      Width           =   495
   End
   Begin VB.CommandButton cmdMiroirFichierSource 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":C46C
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   4290
      Width           =   495
   End
   Begin VB.CommandButton cmdInverserSensSequ 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":CA96
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   3795
      Width           =   495
   End
   Begin VB.CommandButton cmdAfficherSensSequ 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":D0C0
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   3300
      Width           =   495
   End
   Begin VB.CommandButton cmdAgrandirRetrecir 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":D6EA
      Style           =   1  'Graphical
      TabIndex        =   139
      TabStop         =   0   'False
      ToolTipText     =   "Taille de la fenêtre"
      Top             =   5865
      Width           =   495
   End
   Begin VB.CheckBox chkCouleurProfils 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":DDD4
      Style           =   1  'Graphical
      TabIndex        =   138
      TabStop         =   0   'False
      ToolTipText     =   "Alterner la couleur des profils"
      Top             =   5370
      Width           =   495
   End
   Begin VB.CheckBox chkZoomProjet 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":E44E
      Style           =   1  'Graphical
      TabIndex        =   137
      TabStop         =   0   'False
      ToolTipText     =   "Zoom Bloc"
      Top             =   6360
      Width           =   495
   End
   Begin VB.CheckBox chkVoirPoints 
      Height          =   495
      Left            =   14730
      Picture         =   "frmMiniCut2d.frx":EA78
      Style           =   1  'Graphical
      TabIndex        =   136
      TabStop         =   0   'False
      ToolTipText     =   "Afficher les points"
      Top             =   4875
      Width           =   495
   End
   Begin VB.PictureBox pctFil 
      Height          =   9075
      Left            =   4215
      ScaleHeight     =   9015
      ScaleWidth      =   4005
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   435
      Width           =   4065
      Begin VB.Timer TimerTempoChauffeManu 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3315
         Top             =   120
      End
      Begin VB.Frame frameChauffeFil 
         Caption         =   "Faire chauffer"
         ForeColor       =   &H00C00000&
         Height          =   990
         Left            =   180
         TabIndex        =   115
         Top             =   810
         Width           =   3645
         Begin VB.PictureBox PictureChauffe 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   780
            Picture         =   "frmMiniCut2d.frx":F0F2
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   155
            ToolTipText     =   "Chauffe"
            Top             =   285
            Width           =   375
         End
         Begin VB.OptionButton optChauffe 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ON"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   0
            Left            =   2700
            Picture         =   "frmMiniCut2d.frx":F68C
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Faire chauffer"
            Top             =   300
            Width           =   660
         End
         Begin VB.OptionButton optChauffe 
            BackColor       =   &H008080FF&
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   1
            Left            =   2985
            Picture         =   "frmMiniCut2d.frx":FC26
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Arrêter la chauffe"
            Top             =   150
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.HScrollBar hscChauffeFilManuel 
            Height          =   255
            LargeChange     =   10
            Left            =   150
            Max             =   100
            Min             =   1
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   540
            Value           =   1
            Width           =   1515
         End
         Begin VB.Label lblChauffeManuel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "XXX %"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1920
            TabIndex        =   117
            Top             =   570
            Width           =   570
         End
      End
      Begin VB.Frame frameFilManuel 
         Caption         =   "Déplacer"
         ForeColor       =   &H00C00000&
         Height          =   3375
         Left            =   180
         TabIndex        =   114
         Top             =   3645
         Width           =   3645
         Begin VB.OptionButton optGoManuel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Index           =   0
            Left            =   2415
            Picture         =   "frmMiniCut2d.frx":101C0
            Style           =   1  'Graphical
            TabIndex        =   171
            ToolTipText     =   "Lancer le mouvement"
            Top             =   795
            Width           =   705
         End
         Begin VB.OptionButton optGoManuel 
            BackColor       =   &H008080FF&
            Caption         =   "STOP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Index           =   1
            Left            =   2130
            Picture         =   "frmMiniCut2d.frx":1093E
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Arrêter le mouvement"
            Top             =   795
            Width           =   705
         End
         Begin VB.Frame frameOriginesManu 
            Caption         =   "Auto"
            ForeColor       =   &H00FF0000&
            Height          =   915
            Left            =   255
            TabIndex        =   158
            Top             =   2300
            Width           =   3150
            Begin VB.OptionButton optAnnulerHome 
               BackColor       =   &H00FFC0C0&
               Height          =   525
               Left            =   2145
               Picture         =   "frmMiniCut2d.frx":110BC
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Stop"
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton optHomeX 
               BackColor       =   &H00FFC0C0&
               Height          =   525
               Left            =   1185
               Picture         =   "frmMiniCut2d.frx":11846
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Origine horizontale"
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optHomeY 
               BackColor       =   &H00FFC0C0&
               Height          =   525
               Left            =   225
               Picture         =   "frmMiniCut2d.frx":11F70
               Style           =   1  'Graphical
               TabIndex        =   159
               ToolTipText     =   "Origine verticale"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.HScrollBar hscVitesseManuel 
            Height          =   255
            LargeChange     =   5
            Left            =   765
            Max             =   60
            Min             =   1
            TabIndex        =   156
            Top             =   180
            Value           =   1
            Width           =   1455
         End
         Begin VB.Frame frmFleches 
            Caption         =   "Manu"
            ForeColor       =   &H00FF0000&
            Height          =   1920
            Left            =   255
            TabIndex        =   118
            Top             =   300
            Width           =   3150
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   7
               Left            =   1245
               Picture         =   "frmMiniCut2d.frx":12676
               Style           =   1  'Graphical
               TabIndex        =   169
               Top             =   1350
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   6
               Left            =   690
               Picture         =   "frmMiniCut2d.frx":12C88
               Style           =   1  'Graphical
               TabIndex        =   168
               Top             =   1350
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   5
               Left            =   120
               Picture         =   "frmMiniCut2d.frx":132B2
               Style           =   1  'Graphical
               TabIndex        =   167
               Top             =   1350
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   4
               Left            =   1245
               Picture         =   "frmMiniCut2d.frx":138C4
               Style           =   1  'Graphical
               TabIndex        =   166
               Top             =   810
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   3
               Left            =   120
               Picture         =   "frmMiniCut2d.frx":13EEE
               Style           =   1  'Graphical
               TabIndex        =   165
               Top             =   810
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   2
               Left            =   1245
               Picture         =   "frmMiniCut2d.frx":14518
               Style           =   1  'Graphical
               TabIndex        =   164
               Top             =   270
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   1
               Left            =   690
               Picture         =   "frmMiniCut2d.frx":14B2A
               Style           =   1  'Graphical
               TabIndex        =   163
               Top             =   270
               Width           =   495
            End
            Begin VB.OptionButton optManuel 
               BackColor       =   &H00FFC0C0&
               Height          =   465
               Index           =   0
               Left            =   120
               Picture         =   "frmMiniCut2d.frx":15154
               Style           =   1  'Graphical
               TabIndex        =   162
               Top             =   270
               Width           =   495
            End
         End
         Begin VB.Image imgVitesseManuel 
            Height          =   180
            Left            =   240
            Picture         =   "frmMiniCut2d.frx":15766
            ToolTipText     =   "Vitesse"
            Top             =   165
            Width           =   330
         End
         Begin VB.Label lblValeurVitesseManuel 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "X.X mm/s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2460
            TabIndex        =   157
            Top             =   210
            Width           =   810
         End
      End
      Begin VB.Frame frameInformationFil 
         Caption         =   "Information"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   180
         TabIndex        =   111
         Top             =   1950
         Width           =   3645
         Begin MSComctlLib.ProgressBar progrbarChauffeManu 
            Height          =   255
            Left            =   90
            TabIndex        =   130
            Top             =   1155
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lblAvertissementFil2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Avertissement Découpe"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   915
            TabIndex        =   131
            Top             =   1080
            Width           =   1755
         End
         Begin VB.Label lblProcedure 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Procédure demandée"
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
            Left            =   915
            TabIndex        =   113
            Top             =   330
            Width           =   1815
         End
         Begin VB.Label lblAvertissementFil 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Avertissement Découpe"
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   915
            TabIndex        =   112
            Top             =   765
            Width           =   1755
         End
      End
      Begin VB.Frame frameAnnulationFil 
         Caption         =   "Quitter"
         ForeColor       =   &H00C00000&
         Height          =   1065
         Left            =   180
         TabIndex        =   109
         Top             =   7500
         Width           =   3645
         Begin VB.CommandButton cmdAnnulerFil 
            BackColor       =   &H00FFC0C0&
            Height          =   585
            Left            =   150
            Picture         =   "frmMiniCut2d.frx":15D00
            Style           =   1  'Graphical
            TabIndex        =   110
            TabStop         =   0   'False
            ToolTipText     =   "Quitter"
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Label lblPiloterLeFil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     Piloter le fil     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1260
         TabIndex        =   108
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdSimulation 
      BackColor       =   &H00C0E0FF&
      Height          =   990
      Left            =   2370
      Picture         =   "frmMiniCut2d.frx":167A2
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      ToolTipText     =   "Voir le déplacement du fil"
      Top             =   1725
      Width           =   1320
   End
   Begin VB.PictureBox pctMasquage 
      Height          =   2460
      Left            =   3840
      ScaleHeight     =   2400
      ScaleWidth      =   1425
      TabIndex        =   121
      Top             =   825
      Width           =   1485
   End
   Begin VB.CommandButton cmdSauver 
      Height          =   405
      Index           =   1
      Left            =   1785
      Picture         =   "frmMiniCut2d.frx":1721C
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Enregistrer"
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton cmdNouveauProjet 
      Height          =   405
      Left            =   45
      Picture         =   "frmMiniCut2d.frx":177A6
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Nouveau projet"
      Top             =   30
      Width           =   495
   End
   Begin VB.PictureBox pctDecoupe 
      AutoRedraw      =   -1  'True
      Height          =   2955
      Left            =   8505
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   165
      Width           =   6645
      Begin VB.Timer TimerScreenSaver 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   1455
         Top             =   2295
      End
      Begin VB.Timer TimerClignotement 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2865
         Top             =   1980
      End
      Begin MSComDlg.CommonDialog DialogueFichiers 
         Left            =   2820
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer TimerFilm 
         Enabled         =   0   'False
         Left            =   2280
         Top             =   1890
      End
   End
   Begin VB.PictureBox pctTransf 
      DragIcon        =   "frmMiniCut2d.frx":17D30
      Height          =   4875
      Left            =   4095
      MouseIcon       =   "frmMiniCut2d.frx":185FA
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   740
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4830
      Width           =   11160
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9990
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":18EC4
               Key             =   "Mesurer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1979E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1A078
               Key             =   "Transferer"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1A952
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1B22C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1BB06
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1C3E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1CCBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1D594
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1DE6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1E748
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1F022
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9270
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   22
         ImageHeight     =   19
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1F8FC
               Key             =   "Bibliotheque"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":1FF6A
               Key             =   "Dossier"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":205D8
               Key             =   "Decoupe"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMiniCut2d.frx":20C46
               Key             =   "DossierFerme"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblAvertissementTransf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Avertissement Transférées"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   9180
         TabIndex        =   52
         Top             =   3435
         Width           =   1905
      End
      Begin VB.Label lblMesure 
         AutoSize        =   -1  'True
         Caption         =   "Mesure avec l'outil mètre"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   9300
         TabIndex        =   10
         Top             =   3105
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdSauver 
      Height          =   405
      Index           =   0
      Left            =   1095
      Picture         =   "frmMiniCut2d.frx":212B4
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Enregistrer sous..."
      Top             =   30
      Width           =   660
   End
   Begin VB.PictureBox pctBandeauSaisie 
      Height          =   375
      Left            =   4095
      ScaleHeight     =   315
      ScaleWidth      =   11100
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9705
      Width           =   11160
      Begin VB.TextBox txtMesures 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   3
         Text            =   "360.69"
         Top             =   30
         Width           =   615
      End
      Begin VB.TextBox txtMesures 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   2355
         TabIndex        =   4
         Text            =   "360.69"
         Top             =   30
         Width           =   615
      End
      Begin VB.Label lblDimensionSelection 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sélection :"
         Height          =   195
         Left            =   10275
         TabIndex        =   14
         Top             =   60
         Width           =   750
      End
      Begin VB.Label lblMetre 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "mm"
         Height          =   195
         Left            =   165
         TabIndex        =   9
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblMesures 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y (mm) :"
         Height          =   195
         Index           =   1
         Left            =   1740
         TabIndex        =   8
         Top             =   60
         Width           =   570
      End
      Begin VB.Label lblMesures 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Angle (°) :"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   13
         Top             =   60
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOuvrirFichierSequ 
      Height          =   405
      Left            =   570
      Picture         =   "frmMiniCut2d.frx":218BE
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Ouvrir un projet"
      Top             =   30
      Width           =   495
   End
   Begin VB.PictureBox pctSequ 
      DragIcon        =   "frmMiniCut2d.frx":21E58
      Height          =   4815
      Left            =   4095
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   11160
      Begin VB.PictureBox pctAjoutPoint 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   35
         TabIndex        =   172
         ToolTipText     =   "Ajouter un point"
         Top             =   60
         Width           =   555
      End
      Begin VB.Label lblAvertissementSequ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Avertissement Séquences"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   3330
         TabIndex        =   54
         Top             =   60
         Width           =   1875
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Création"
      TabPicture(0)   =   "frmMiniCut2d.frx":22722
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDimensionsBloc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblBlocMaxi(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblMm(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAlignerAjuster"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line3(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line4(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line5(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line6(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblBoiteOutils"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblMm(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "shapeBloc"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblBlocMaxi(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line3(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line4(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line5(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Line6(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line7"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line8"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line9"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Line10"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Line12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Line13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line14"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line15"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Tree"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "optOutils(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdAligner(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdAligner(1)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdAligner(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdAligner(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdAligner(4)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdAligner(5)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdInsererPoint"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmdDupliquer"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdUndo(0)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtBloc(0)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "optOutils(5)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "chkCentrer(0)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "chkCentrer(1)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chkCentrer(2)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdPoubelle"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdMiroir"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "cmdInverser"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "optOutils(2)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "optOutils(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "optOutils(0)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cmdUndo(1)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtBloc(1)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cmdImporterProfil"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "cmdRafraichir"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "cmdEffacerFichier"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "optOutils(6)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "cmdImporterImage"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "Découpe"
      TabPicture(1)   =   "frmMiniCut2d.frx":2273E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTrajet"
      Tab(1).Control(1)=   "lblDecoupe"
      Tab(1).Control(2)=   "lblFil"
      Tab(1).Control(3)=   "lblTitreChauffe"
      Tab(1).Control(4)=   "cmdDecouper"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frameSortieBloc"
      Tab(1).Control(6)=   "frameEntreeBloc"
      Tab(1).Control(7)=   "frameMateriaux"
      Tab(1).Control(8)=   "frameProcedures"
      Tab(1).Control(9)=   "frameManuel"
      Tab(1).Control(10)=   "frameSimulation"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdImporterImage 
         Height          =   420
         Left            =   3180
         Picture         =   "frmMiniCut2d.frx":2275A
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Vectoriser une image"
         Top             =   810
         Width           =   510
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   6
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":22D0C
         Style           =   1  'Graphical
         TabIndex        =   135
         TabStop         =   0   'False
         ToolTipText     =   "Premier point d'un trajet"
         Top             =   4410
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdEffacerFichier 
         Height          =   420
         Left            =   2280
         Picture         =   "frmMiniCut2d.frx":235D6
         Style           =   1  'Graphical
         TabIndex        =   134
         ToolTipText     =   "Supprimer (corbeille)"
         Top             =   375
         Width           =   420
      End
      Begin VB.CommandButton cmdRafraichir 
         Height          =   420
         Left            =   2715
         Picture         =   "frmMiniCut2d.frx":23BB8
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Actualiser"
         Top             =   375
         Width           =   420
      End
      Begin VB.Frame frameSimulation 
         Caption         =   "Simulation"
         Height          =   1770
         Left            =   -72765
         TabIndex        =   129
         Top             =   825
         Width           =   1590
      End
      Begin VB.CommandButton cmdImporterProfil 
         Height          =   420
         Left            =   3150
         Picture         =   "frmMiniCut2d.frx":2419A
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Importer un fichier dans la bibliothèque"
         Top             =   375
         Width           =   540
      End
      Begin VB.TextBox txtBloc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   38
         Text            =   "200"
         Top             =   8250
         Width           =   600
      End
      Begin VB.Frame frameManuel 
         Caption         =   "Piloter"
         ForeColor       =   &H00FF0000&
         Height          =   1020
         Left            =   -74820
         TabIndex        =   78
         Top             =   7155
         Width           =   3660
         Begin VB.CommandButton cmdDeplacementsManuels 
            BackColor       =   &H00FFC0C0&
            Height          =   540
            Left            =   165
            Picture         =   "frmMiniCut2d.frx":24804
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Piloter le fil en direct"
            Top             =   255
            Width           =   3285
         End
      End
      Begin VB.Frame frameProcedures 
         Caption         =   "Positions"
         ForeColor       =   &H00C00000&
         Height          =   1110
         Left            =   -74820
         TabIndex        =   66
         Top             =   8310
         Width           =   3660
         Begin VB.CommandButton cmdPlierLePortique 
            BackColor       =   &H00FFC0C0&
            Height          =   570
            Left            =   180
            Picture         =   "frmMiniCut2d.frx":254BE
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Aller en position de rangement"
            Top             =   330
            Width           =   1545
         End
         Begin VB.CommandButton cmdRetourOrigine 
            BackColor       =   &H00FFC0C0&
            Height          =   585
            Left            =   1920
            Picture         =   "frmMiniCut2d.frx":25DA0
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            ToolTipText     =   "Ramener le fil à l'origine"
            Top             =   315
            Width           =   1545
         End
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00F0E2F0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   150
         Picture         =   "frmMiniCut2d.frx":26682
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Refaire"
         Top             =   5415
         Width           =   615
      End
      Begin VB.Frame frameMateriaux 
         ForeColor       =   &H000000C0&
         Height          =   2220
         Left            =   -74820
         TabIndex        =   58
         Top             =   3090
         Width           =   3675
         Begin VB.HScrollBar hscVitesseBDD 
            Height          =   255
            LargeChange     =   5
            Left            =   750
            Max             =   60
            Min             =   1
            TabIndex        =   153
            Top             =   555
            Value           =   1
            Width           =   1455
         End
         Begin VB.PictureBox imgChauffe 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   225
            Picture         =   "frmMiniCut2d.frx":26BBC
            ScaleHeight     =   195
            ScaleWidth      =   375
            TabIndex        =   125
            ToolTipText     =   "Chauffe"
            Top             =   270
            Width           =   375
         End
         Begin VB.HScrollBar hscChauffeDecoupe 
            Height          =   255
            LargeChange     =   10
            Left            =   750
            Max             =   100
            Min             =   1
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   270
            Value           =   1
            Width           =   1860
         End
         Begin VB.CommandButton cmdGestionMatiere 
            BackColor       =   &H00C0FFFF&
            Height          =   705
            Index           =   2
            Left            =   2515
            Picture         =   "frmMiniCut2d.frx":27156
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Supprimer matériau"
            Top             =   1260
            Width           =   945
         End
         Begin VB.CommandButton cmdGestionMatiere 
            BackColor       =   &H00C0FFFF&
            Height          =   705
            Index           =   1
            Left            =   1370
            Picture         =   "frmMiniCut2d.frx":27A70
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Remplacer la valeur de la chauffe"
            Top             =   1260
            Width           =   945
         End
         Begin VB.CommandButton cmdGestionMatiere 
            BackColor       =   &H00C0FFFF&
            Height          =   705
            Index           =   0
            Left            =   200
            Picture         =   "frmMiniCut2d.frx":2869A
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Nouveau matériau"
            Top             =   1260
            Width           =   945
         End
         Begin VB.ComboBox comboMatieres 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "@Arial Unicode MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmMiniCut2d.frx":28FB4
            Left            =   90
            List            =   "frmMiniCut2d.frx":28FB6
            Style           =   2  'Dropdown List
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   765
            Width           =   3450
         End
         Begin VB.Label lblVitesseDecoupe 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "X.X mm/s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2450
            TabIndex        =   154
            Top             =   585
            Width           =   810
         End
         Begin VB.Image imgVitesse 
            Height          =   180
            Left            =   225
            Picture         =   "frmMiniCut2d.frx":28FB8
            ToolTipText     =   "Vitesse"
            Top             =   540
            Width           =   330
         End
         Begin VB.Label lblChauffeDecoupe 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "XX %"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2865
            TabIndex        =   124
            Top             =   300
            Width           =   420
         End
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   0
         Left            =   3300
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":29552
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Déplacer"
         Top             =   4395
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   1
         Left            =   2670
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":29E1C
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Tourner"
         Top             =   4395
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   2
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":2A6E6
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Etirer"
         Top             =   4395
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdInverser 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Left            =   3300
         Picture         =   "frmMiniCut2d.frx":2AFB0
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Inverser le sens"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CommandButton cmdMiroir 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Left            =   2670
         Picture         =   "frmMiniCut2d.frx":2B87A
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Miroir"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CommandButton cmdPoubelle 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Left            =   780
         Picture         =   "frmMiniCut2d.frx":2C144
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Supprimer"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CheckBox chkCentrer 
         BackColor       =   &H00F0E2F0&
         Height          =   540
         Index           =   2
         Left            =   2940
         Picture         =   "frmMiniCut2d.frx":2CA0E
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Centrer verticalement"
         Top             =   6300
         Width           =   615
      End
      Begin VB.CheckBox chkCentrer 
         BackColor       =   &H00F0E2F0&
         Height          =   540
         Index           =   1
         Left            =   2325
         Picture         =   "frmMiniCut2d.frx":2D200
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Centrer horizontalement"
         Top             =   6300
         Width           =   615
      End
      Begin VB.CheckBox chkCentrer 
         BackColor       =   &H00F0E2F0&
         Height          =   540
         Index           =   0
         Left            =   1710
         Picture         =   "frmMiniCut2d.frx":2D9F2
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Mettre à l'échelle du bloc"
         Top             =   6300
         Width           =   615
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   5
         Left            =   780
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":2E1E4
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Couper un trajet"
         Top             =   4410
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtBloc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Index           =   0
         Left            =   195
         TabIndex        =   37
         Text            =   "200"
         Top             =   8805
         Width           =   600
      End
      Begin VB.Frame frameEntreeBloc 
         Caption         =   "Entrée"
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   -74835
         TabIndex        =   33
         Top             =   825
         Width           =   1935
         Begin VB.OptionButton optEntrerBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   0
            Left            =   75
            Picture         =   "frmMiniCut2d.frx":2EAAE
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Value           =   -1  'True
            Width           =   540
         End
         Begin VB.OptionButton optEntrerBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   1
            Left            =   660
            Picture         =   "frmMiniCut2d.frx":2F198
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Width           =   540
         End
         Begin VB.OptionButton optEntrerBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   2
            Left            =   1260
            Picture         =   "frmMiniCut2d.frx":2F882
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame frameSortieBloc 
         Caption         =   "Sortie"
         ForeColor       =   &H00008000&
         Height          =   855
         Left            =   -74835
         TabIndex        =   29
         Top             =   1740
         Width           =   1935
         Begin VB.OptionButton optSortirBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   0
            Left            =   75
            Picture         =   "frmMiniCut2d.frx":2FF6C
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   240
            Width           =   540
         End
         Begin VB.OptionButton optSortirBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   1
            Left            =   660
            Picture         =   "frmMiniCut2d.frx":30656
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   540
         End
         Begin VB.OptionButton optSortirBloc 
            BackColor       =   &H00C0E0FF&
            Height          =   495
            Index           =   2
            Left            =   1260
            Picture         =   "frmMiniCut2d.frx":30D40
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdDecouper 
         BackColor       =   &H0080FF80&
         Height          =   705
         Left            =   -74745
         Picture         =   "frmMiniCut2d.frx":3142A
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Découper le projet"
         Top             =   5880
         Width           =   3480
      End
      Begin VB.CommandButton cmdUndo 
         BackColor       =   &H00F0E2F0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   150
         Picture         =   "frmMiniCut2d.frx":31FF4
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Défaire"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CommandButton cmdDupliquer 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Left            =   2040
         Picture         =   "frmMiniCut2d.frx":3252E
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Dupliquer"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CommandButton cmdInsererPoint 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Left            =   1410
         Picture         =   "frmMiniCut2d.frx":32DF8
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Insérer un point"
         Top             =   5100
         Width           =   615
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   5
         Left            =   3150
         Picture         =   "frmMiniCut2d.frx":336C2
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Aligner à droite"
         Top             =   7035
         Width           =   435
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   4
         Left            =   2415
         Picture         =   "frmMiniCut2d.frx":33F8C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Aligner au milieu"
         Top             =   7035
         Width           =   405
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   3
         Left            =   1665
         Picture         =   "frmMiniCut2d.frx":34856
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Aligner à gauche"
         Top             =   7020
         Width           =   435
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   435
         Index           =   2
         Left            =   360
         Picture         =   "frmMiniCut2d.frx":35120
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Aligner en haut"
         Top             =   6240
         Width           =   615
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   405
         Index           =   1
         Left            =   360
         Picture         =   "frmMiniCut2d.frx":359EA
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Aligner au milieu"
         Top             =   6720
         Width           =   615
      End
      Begin VB.CommandButton cmdAligner 
         BackColor       =   &H00F0E2F0&
         Height          =   435
         Index           =   0
         Left            =   360
         Picture         =   "frmMiniCut2d.frx":362B4
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Aligner en bas"
         Top             =   7200
         Width           =   615
      End
      Begin VB.OptionButton optOutils 
         BackColor       =   &H00F0E2F0&
         Height          =   615
         Index           =   4
         Left            =   150
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMiniCut2d.frx":36B7E
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Mesurer"
         Top             =   4410
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin MSComctlLib.TreeView Tree 
         Height          =   3420
         Left            =   45
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   360
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   6033
         _Version        =   393217
         Indentation     =   794
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
      End
      Begin VB.Label lblTitreChauffe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Chauffe et vitesse "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -73845
         TabIndex        =   40
         Top             =   2820
         Width           =   1710
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00404040&
         X1              =   1510
         X2              =   1455
         Y1              =   8985
         Y2              =   8810
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00404040&
         X1              =   1455
         X2              =   1395
         Y1              =   8820
         Y2              =   8985
      End
      Begin VB.Line Line13 
         BorderColor     =   &H00404040&
         X1              =   1395
         X2              =   1455
         Y1              =   9225
         Y2              =   9375
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00404040&
         X1              =   1470
         X2              =   1530
         Y1              =   9360
         Y2              =   9210
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00404040&
         X1              =   1455
         X2              =   1455
         Y1              =   8820
         Y2              =   9375
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00404040&
         X1              =   1650
         X2              =   1215
         Y1              =   9380
         Y2              =   9380
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00404040&
         X1              =   1650
         X2              =   1200
         Y1              =   8810
         Y2              =   8810
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         X1              =   2775
         X2              =   2940
         Y1              =   8720
         Y2              =   8655
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   2920
         X2              =   2775
         Y1              =   8655
         Y2              =   8580
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   1710
         X2              =   1845
         Y1              =   8670
         Y2              =   8715
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   1710
         X2              =   1845
         Y1              =   8655
         Y2              =   8580
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   1725
         X2              =   2920
         Y1              =   8655
         Y2              =   8655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         Index           =   2
         X1              =   2930
         X2              =   2930
         Y1              =   8760
         Y2              =   8220
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         Index           =   1
         X1              =   1710
         X2              =   1710
         Y1              =   8760
         Y2              =   8220
      End
      Begin VB.Label lblBlocMaxi 
         AutoSize        =   -1  'True
         Caption         =   "< XXX mm"
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
         Left            =   3030
         TabIndex        =   120
         Top             =   8295
         Width           =   840
      End
      Begin VB.Shape shapeBloc 
         BorderColor     =   &H00808000&
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   600
         Left            =   1695
         Top             =   8790
         Width           =   1245
      End
      Begin VB.Label lblMm 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   2430
         TabIndex        =   67
         Top             =   8295
         Width           =   330
      End
      Begin VB.Label lblFil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Fil  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -73275
         TabIndex        =   65
         Top             =   6855
         Width           =   480
      End
      Begin VB.Label lblDecoupe 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Découpe "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -73470
         TabIndex        =   49
         Top             =   5505
         Width           =   945
      End
      Begin VB.Label lblBoiteOutils 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Boite à outils "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1335
         TabIndex        =   45
         Top             =   4020
         Width           =   1230
      End
      Begin VB.Line Line6 
         Index           =   0
         X1              =   2865
         X2              =   3105
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   2145
         X2              =   2370
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   1200
         X2              =   1095
         Y1              =   7410
         Y2              =   7410
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   1200
         X2              =   1095
         Y1              =   6495
         Y2              =   6495
      End
      Begin VB.Line Line2 
         X1              =   1200
         X2              =   1095
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1200
         X2              =   1200
         Y1              =   6495
         Y2              =   7425
      End
      Begin VB.Label lblAlignerAjuster 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Aligner et ajuster "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1155
         TabIndex        =   44
         Top             =   5910
         Width           =   1590
      End
      Begin VB.Label lblMm 
         AutoSize        =   -1  'True
         Caption         =   "mm"
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
         Left            =   825
         TabIndex        =   43
         Top             =   8850
         Width           =   330
      End
      Begin VB.Label lblBlocMaxi 
         AutoSize        =   -1  'True
         Caption         =   "< XXX mm"
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
         Left            =   165
         TabIndex        =   42
         Top             =   9150
         Width           =   840
      End
      Begin VB.Label lblDimensionsBloc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dimensions du bloc "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1065
         TabIndex        =   41
         Top             =   7860
         Width           =   1800
      End
      Begin VB.Label lblTrajet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Trajet "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   -73350
         TabIndex        =   39
         Top             =   510
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdSettings 
      Height          =   405
      Left            =   3480
      Picture         =   "frmMiniCut2d.frx":37448
      Style           =   1  'Graphical
      TabIndex        =   102
      TabStop         =   0   'False
      ToolTipText     =   "A propos..."
      Top             =   30
      Width           =   540
   End
   Begin VB.CommandButton cmdLangue 
      Height          =   405
      Left            =   2970
      Picture         =   "frmMiniCut2d.frx":37A2A
      Style           =   1  'Graphical
      TabIndex        =   133
      ToolTipText     =   "Langue"
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "frmMiniCut2d"
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

'********* INITIALISATION de la FENETRE ***********
Private Sub Form_Load()
   Dim i As Long, j As Long
   Dim ctrl As Control
   Dim CleDeLaSection() As String
   Dim lngReturnCode As Long
   Dim singleTemp As Single
   Dim stringTemp As String
   
   '************** EN TOUT PREMIER, le .INI ***************
   'Contrôler la présence du fichier .ini, on le crée s'il est absent avec les paramètres par défaut
   If Dir(App.Path & "\MiniCut2d_Software.ini", vbNormal) = "" Then
      'Il n'y a pas de fichier .ini : on en crée un avec les paramètres par défaut
      'L'écriture d'une clé engendre automatiquement le fichier s'il n'existe pas
      'Les matières :
      'La section [BDDMatières] comprend le nombre de matières et le nom de celle qui doit être affichée dans la combobox
      EcritFichierIni "BDDMatieres", "MatiereUtilisee", "Reglage par defaut"  'c'est la dernière matière utilisée et chargée à l'ouverture
      'Chaque matière fait ensuite l'objet d'une section avec la valeur de chauffe en clé
      EcritFichierIni "Matiere_" & "Reglage par defaut", "ChauffeDecoupe", 65
      EcritFichierIni "Matiere_" & "Reglage par defaut", "VitesseDecoupe", "4.0"
      'On définit égelement le dossier de sauvegarde des projets par défaut
      EcritFichierIni "Fichiers", "DernierRepertoire", App.Path & "\Bibliotheque\Mes Decoupes\"
      strLangue = "francais"
      EcritFichierIni "Parametres", "Langue", strLangue  'au premier lancement du soft, la langue est le français
      ModeSoft = "Normal"
      EcritFichierIni "Logiciel", "Mode", ModeSoft 'au premier lancement du soft, pas de mode expert
   Else  'le fichier .ini existe, on vérifie qu'il est conforme
      'Pour internationnalisation, on supprime les accents de "Matière_" & "Réglage par défaut"
      If LitFichierIni("Matiere_" & "Réglage par défaut", "ChauffeDecoupe", "") <> "" Then  'si ça renvoie une clé valide, la section existe
         'On efface l'ancienne section
         lngReturnCode = WritePrivateProfileString("Matiere_" & "Réglage par défaut", 0&, 0&, App.Path & "\MiniCut2d_Software.ini")
         'Et on réécrit la nouvelle
         EcritFichierIni "Matiere_" & "Reglage par defaut", "ChauffeDecoupe", 65
         EcritFichierIni "Matiere_" & "Reglage par defaut", "VitesseDecoupe", "4.0"
      End If
      'on peut encore avoir un accent dans "MatiereUtilisee", on le supprime
      If LitFichierIni("BDDMatieres" & "MatiereUtilisee", "") = "Réglage par défaut" Then  'si ça renvoie une clé valide, la section existe
         EcritFichierIni "BDDMatieres", "MatiereUtilisee", "Reglage par defaut"
      End If
   End If

   'prise en compte du mode du logiciel
   ModeSoft = LitFichierIni("Logiciel", "Mode")
   If ModeSoft = "" Then
      ModeSoft = "Normal"
      EcritFichierIni "Logiciel", "Mode", ModeSoft
   End If
   
   '******** Préparation des contrôles pour affichage et comportement *****
   'attention, cette procédure lance le Form_Load de toutes les feuilles affectées
   Call InitialiserControles     'dans modPlacementControles.bas
   
   '****** uniformiser les icones ********
   With frmMiniCut2d
      frmSplashScreen.Icon = .Icon
      frmAboutAndSettings.Icon = .Icon
      frmDecoupeInactive.Icon = .Icon
      frmLangue.Icon = .Icon
      frmParametres.Icon = .Icon
      frmImpConv.Icon = .Icon
   End With
  
   'apparition des éléments liés au mode
   Call AffichageEnFonctionDuModeSoft(ModeSoft)

   'prise en compte de la langue
   strLangue = LitFichierIni("Parametres", "Langue")
   If strLangue = "" Then
      strLangue = "francais"
      EcritFichierIni "Parametres", "Langue", strLangue  'pour vieille version du .ini
   End If
   Call GestionLangue(strLangue)
   'on met le bouton radio sur la bonne valeur
   Select Case strLangue
   Case "francais"
      frmLangue.optLangue(0).Value = True
   Case "USA"
      frmLangue.optLangue(1).Value = True
   Case "english"
      frmLangue.optLangue(2).Value = True
   Case "deutsch"
      frmLangue.optLangue(3).Value = True
   Case "espanol"
      frmLangue.optLangue(4).Value = True
   End Select
   '******** Affichage du Splash Screen *********
   ' Création du .ini s'il est absent, choix de la machine
   frmSplashScreen.Show vbModal, Me
     
   TimerScreenSaver.Enabled = True 'on active le timer qui empêche la mise en veille
   
   flagTableEcriteDansIPL = False  'permet le branchement à chaud : sera contrôlé avant toute demande d'accès à un endroit du soft où on peut faire bouger/chauffer
   flagPositionPliage = True 'indique que la machine est en position de rangement => contrôle à la fermeture du soft
   flagPliageApresMsgBox = False
   
   'Rendre impossible la mémorisation dans undo tant que le Form Load n'est pas terminé (appelé dans TraceTransf)
   flagMemoUndoDansTraceTransf = False
   'Afficher le sens dans le TraceSequ
   flagAfficherSens = True
   
   'Initialisation de la variable pour effacement de fichiers
   CheminFichier = ""
   
   'Suppression des accents pour fonctionnement international : on teste si "Bibliothèque" existe et on le remplace par "Bibliotheque"
   If VerifierExistenceRepertoire(App.Path & "\Bibliothèque") = True Then
      Name App.Path & "\Bibliothèque" As App.Path & "\Bibliotheque"
   End If
   
   'Vérifier l'existence du dossier "Bibliotheque" et le créer si besoin
   If VerifierExistenceRepertoire(App.Path & "\Bibliotheque") = False Then
      MsgBox Message(Corps, 2), vbInformation, Message(Titre, 2) 'répértoire Bibliothèque absent, il sera créé
      MkDir (App.Path & "\Bibliotheque")
      MkDir (App.Path & "\Bibliotheque\Mes Decoupes")
   End If
   'Si "Bibliothèque" existe, il faut tester "Mes Decoupes"
   If VerifierExistenceRepertoire(App.Path & "\Bibliotheque\Mes Decoupes") = False Then
      MkDir (App.Path & "\Bibliotheque\Mes Decoupes")
   End If
   'ON vérifie que le Dernier Répertoire mémorisé existe (si le chemin contient des accents, il n'existe plus puisqu'on a viré les accents
   ' et ça plantera à la lecture)
   If VerifierExistenceRepertoire(LitFichierIni("Fichiers", "DernierRepertoire")) = False Then
      EcritFichierIni "Fichiers", "DernierRepertoire", App.Path & "\Bibliotheque\Mes Decoupes\"
   End If
   
   'On efface l'ancienne section "Table" qui deviendra "Machine" (qui aura plus de sens pour le newbie et en anglais)
   lngReturnCode = WritePrivateProfileString("Table", 0&, 0&, App.Path & "\MiniCut2d_Software.ini")
   
   'prise en compte du type de machine
   TypeMachine = LitFichierIni("Machine", "Type")
   If TypeMachine = "" Then
      TypeMachine = "MiniCut2d_v1.2"
      EcritFichierIni "Machine", "Type", TypeMachine  'pour vieille version du .ini
   End If
   Call ParametresMachine
   
   Call GestionLangue(strLangue)  'pour le About, et tout ce qui tient compte de TypeMachine
   'Création des tableaux du dessin de la table de l'onglet découpe et calcul des extremums
   'on a besoin des courses pour les dessins, elles sont définies ci-dessus
   Call TableauTraceMachine  'remplissage des tableaux de définition du dessin de la table (module modDessinTable)
      
   '************* CREATION DE LA TABLE DES ACCELERATIONS *************
      'Ensuite, il faut créer les tables nécessaires aux accélérations pour l'interpolateur
   ReDim TableLIN(1 To 5, 0 To 255)  'les indices 1 à 5 correspondent aux fréquences 10000 à 50000 Hz
   For j = 1 To 5
      For i = 0 To 255
         TableLIN(j, i) = Round(6000000 / ((j * 10000 / 255) * (i) + (255 - i) * 6000000 / (65535 * 255)), 0)
      Next i
   Next j

   '**************************** PARAMETRES TABLE ET BLOC ******************************
   'Initialisation du bloc par défaut : on va utiliser 80% des courses utiles arrondi à la dizaine
   Call InitialisationCoursesEtBloc
   
   '******** Préparation des contrôles pour affichage et comportement *****
   Call InitialiserControles     'dans modPlacementControles.bas
   
   '******** Initialisation des scrollbar de chauffe *********
   hscChauffeDecoupe.Max = 100
   hscChauffeStop.Max = hscChauffeDecoupe.Max
   hscChauffeFilManuel.Max = hscChauffeDecoupe.Max
   hscChauffePendantDecoupe.Max = hscChauffeDecoupe.Max
   
   '******** Initialisation du scrollbar de vitesse *********
   hscVitesseBDD.Min = 1   'rappel : la vitesse est cette valeur divisée par 10
   hscVitesseBDD.Max = 60  'donc 60/10=6mm/s maxi
   hscVitesseManuel.Min = hscVitesseBDD.Min
   hscVitesseManuel.Max = hscVitesseBDD.Max

   '****** Gestion des TabStop pour navigation avec la touche TAB *****
   On Error Resume Next
   For Each ctrl In Controls
      ctrl.TabStop = False
   Next
   On Error GoTo 0
   txtMesures(0).TabStop = True
   txtMesures(1).TabStop = True
   txtBloc(0).TabStop = True
   txtBloc(1).TabStop = True
   
   flagSimulationLancee = False  'au début, la simulation n'est pas lancée
   flagTracePourSimulation = False  'lors du premier tracé, la découpe est sur le plan provisoire
   flagPremierPoint = False
   flagFenetreAgrandie = False
   flagLeProjetAUnNom = False  'le projet n'a pas de nom
   'initialisation de l'outil undo/redo
   IndexUndo = 0
   ReDim UndoRedo(1 To 1)
   
   'Initialisation de la mise à l'échelle du bloc
   CoeffBloc = 1

   'on initialise les variables des outils
   Call InitialiserOutils
   
   'On active l'outil Déplacer et les curseurs associés
   Call optOutils_Click(Deplacer)
   
   'Matières :
   'contrôle que la matière par défaut n'a pas une vitesse modifiée et on corrige si besoin
   stringTemp = LitFichierIni("Matiere_" & "Reglage par defaut", "VitesseDecoupe", "")
   If stringTemp <> "4.0" Then
      EcritFichierIni "Matiere_" & "Reglage par defaut", "VitesseDecoupe", "4.0"
   End If
   
   Call ListerMatieresDuIni
   
   MatiereUtilisee.Nom = LitFichierIni("BDDMatieres", "MatiereUtilisee")
       
   'Réglage des scrollbar sur la valeur de la matière utilisée ; si elle est trop élevée, on la bride au maxi
   singleTemp = MyVal(LitFichierIni("Matiere_" & MatiereUtilisee.Nom, "ChauffeDecoupe"))
   If singleTemp <= hscChauffeDecoupe.Max Then
      hscChauffeDecoupe.Value = singleTemp
   Else
      MsgBox Message(Corps, 4), vbCritical, Message(Titre, 4)  'chauffe trop élevée, bridage
      hscChauffeDecoupe.Value = hscChauffeDecoupe.Max
   End If
   ChauffeCourante = hscChauffeDecoupe.Value
   lblChauffeDecoupe.Caption = Format(ChauffeCourante, "##0") & " %"
   hscChauffeStop.Value = hscChauffeDecoupe.Value
   lblChauffeStop.Caption = Format(ChauffeCourante, "##0") & " %"
   hscChauffeFilManuel.Value = hscChauffeDecoupe.Value
   lblChauffeManuel.Caption = Format(ChauffeCourante, "##0") & " %"
   hscChauffePendantDecoupe.Value = hscChauffeDecoupe.Value
   lblChauffePendantDecoupe.Caption = Format(ChauffeCourante, "##0") & " %"
   
   TempoChauffeFil = CalculTempoChauffe(ChauffeCourante)  'délai de mise en température du fil

   'chauffe cadre déplacements manuels
   hscChauffeFilManuel.Value = ChauffeCourante
   lblChauffeManuel.Caption = Format(ChauffeCourante, "##0") & " %"
   
   'réglage du scrollbar de vitesse
   singleTemp = MyVal(LitFichierIni("Matiere_" & MatiereUtilisee.Nom, "VitesseDecoupe"))
   If singleTemp * 10 <= hscVitesseBDD.Max Then
      hscVitesseBDD.Value = singleTemp * 10
      hscVitesseManuel.Value = hscVitesseBDD.Value
   Else
      MsgBox Message(Corps, 4), vbCritical, Message(Titre, 4)  'chauffe trop élevée (bidouille du .ini), bridage
      hscVitesseBDD.Value = hscVitesseBDD.Max
      hscVitesseManuel.Value = hscVitesseBDD.Value
   End If
   
   CoeffNett = 5000  'précision du nettoyage des profils à l'import
      
   'Définition des curseurs du drag and drop
   pctSequ.DragIcon = ImageList2.ListImages(3).Picture
   pctTransf.DragIcon = ImageList2.ListImages(9).Picture
   pctAjoutPoint.DragIcon = ImageList2.ListImages(3).Picture
   
   frmMiniCut2d.KeyPreview = True     'pour interception des touches Suppr, Tab, Copier, Couper, Coller
   Call DesactiverDeplacer 'affichage des saisies clavier désactivées
   '****************** ON TRAITE TOUS LES GRAPHIQUES EN PIXELS (pas en twips)*******************
   '**** L'affichage dans la PictureBox du haut se fait à l'aide d'un coefficient, sans ********
   '**** modifier l'échelle de la fenêtre, mais celui dans la PictureBox du bas se fait ********
   '**** en coordonnées réelles en utilisant des échelles personnalisées centrées sur la *******
   '**** position basse du fil ; le bloc est décalé vers la droite de la valeur MargeFil *******
   NbTransf = 0
   NbSequSel = 0
   NumSequSel = 0
   NbTransfSel = 0
   NumTransfSel = 0
   pctSequ.Cls  'effacement de l'éventuel dessin précédent
   pctTransf.Cls  'effacement de l'éventuel dessin précédent
   'on inverse l'axe des Y pour avoir l'origine en bas à gauche
   pctSequ.ScaleTop = pctSequ.ScaleHeight
   pctSequ.ScaleHeight = -pctSequ.ScaleHeight
   'idem pour en bas
   pctTransf.ScaleTop = pctTransf.ScaleHeight
   pctTransf.ScaleHeight = -pctTransf.ScaleHeight
   
   MargInit = 8  'pour affichage total des bords du dessin
   
   'Création du treeview des dossiers et fichiers
   ChemRepRacine = App.Path & "\Bibliotheque"
   Call CreerTreeView
   
   Call EchelleTransf(0, CourseX, 0, CourseY)    'Définition du repère de représentation de la fenêtre du bas
   flagMemoUndoDansTraceTransf = True
   Call TraceTransf       'pour tracé du bloc
   
   'Tester la présence et la validité (numéro de version) de la dll de communication
   On Error GoTo DLL_Introuvable  'gestion de l'erreur si la dll n'est pas trouvée et fait planter le soft
   'Si la dll est installée, il faut encore tester la version
   'la fonction est "IPL5X_Dll_Version", on l'utilise directement, sans variable intermédiaire
   If IPL5X_Dll_Version() < 16 Then
      MsgBox Message(Corps, 5), vbCritical, Message(Titre, 5)  'version d'IP5XComm à changer !
      End
   End If
   On Error GoTo 0   'annule le gestionnaire d'erreur de la dll

   'Contrôle de la présence d'une interface compatible
   If IPL5X_IsConnected() = 1 Then 'l'interpolateur est connecté
      'Chargement de la table dans l'interface
      Call EcrireTable
      If ErrIPL <> 1 Then 'il s'est produit une erreur
         GoTo Erreur
      Else
         flagTableEcriteDansIPL = True
      End If
   End If
   
   Unload frmImpConv  'pour pouvoir la recharger et la positionner lors du show
   
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   Exit Sub
DLL_Introuvable:
   MsgBox Message(Corps, 7), vbCritical, Message(Titre, 7)  'le fichier ipl5xcomm.dll n'a pas été trouvé
   Exit Sub
End Sub

Public Sub InitialisationCoursesEtBloc()
   LongBloc = Int(0.8 * (CourseX - MargeFil) / 10) * 10
   HautBloc = Int(0.8 * (CourseY - MargeFil) / 10) * 10
   txtBloc(0).Text = Format(HautBloc, "##0")
   txtBloc(1).Text = Format(LongBloc, "##0")
   lblBlocMaxi(0).Caption = "< " & Format(CourseY - MargeFil, "##0") & " mm"
   lblBlocMaxi(1).Caption = "< " & Format(CourseX - 2 * MargeFil, "##0") & " mm"
End Sub

Private Sub cmdSettings_Click()
   Call GestionLangue(strLangue)
   frmAboutAndSettings.Show vbModal
End Sub


'************************************************************************
'****** CREATION DU TREEVIEW DE VISUALISATION DES FICHIERS PROFILS ******
'************************************************************************
Public Sub CreerTreeView()
   Tree.ImageList = ImageList1
   Tree.Nodes.Clear
   'Création de la racine, avec le chemin du dossier en clé
   Set NoeudX = Tree.Nodes.Add(, , ChemRepRacine, "Bibliotheque", "Bibliotheque")
   Set fso = New FileSystemObject
   Set dossier = fso.GetFolder(ChemRepRacine)
   Scan dossier  'remplissage du treeview
   Tree.Nodes(ChemRepRacine).Expanded = True
   Set dossier = Nothing  'destruction de l'objet
End Sub

'******************************************************************************
'Procédure récursive de parcours des sous-dossiers pour remplissage du treeview
'******************************************************************************
Public Sub Scan(ByVal dossier As Folder)
   Dim N As Node
   
   'On utilise le File System Object  (vérifier que l'installateur fait le nécessaire!)
   For Each sousdossier In dossier.SubFolders   'exporation récursive de tous les sous-dossiers
      Set NoeudX = Tree.Nodes.Add(CStr(dossier), tvwChild, CStr(sousdossier), ExtraitNomFichier(sousdossier), "DossierFerme")
 '     If ExtraitNomFichier(sousdossier) = "Alphabets" Then  'Développement du dossier "Alphabets"
 '        NoeudX.Expanded = True
 '        NoeudX.Image = "Dossier"
 '     End If
      Scan sousdossier     'fonction du FSO
   Next
   For Each fichier In dossier.Files   'liste des fichiers
      Select Case ExtraitExtensionFichier(fichier)
      Case "DAT", "dat", "DXF", "dxf", "plt", "PLT", "eps", "EPS", "mnc", "MNC", "cpx", "CPX", "fc", "FC", "txt", "TXT" 'on affiche seulement les fichiers connus
         Set NoeudX = Tree.Nodes.Add(CStr(dossier), tvwChild, CStr(fichier), ExtraitNomFichier(fichier), "Decoupe")
      End Select
   Next
End Sub

'*** gestion du redimensionnement de la feuille ***
Private Sub Form_Resize()
   If frmMiniCut2d.WindowState <> vbMinimized Then
      Call TailleFenetre
   End If
End Sub

Private Sub TailleFenetre()
   Dim A As Single  'variable intermédiaire pour test des nombres négatifs
   
   'pour que tous les boutons soient dans le même repère, ils sont posés sur la Form, pas dans les PictureBox.
   With frmMiniCut2d
      chkZoomDecoupe.Top = 3
      If .ScaleWidth > 312 Then
         pctSequ.Width = .ScaleWidth - 274
         pctTransf.Width = pctSequ.Width
         pctBandeauSaisie.Width = pctSequ.Width
         cmdAfficherSensSequ.Left = .ScaleWidth - cmdAfficherSensSequ.Width - 3
         cmdInverserSensSequ.Left = cmdAfficherSensSequ.Left
         cmdMiroirFichierSource.Left = cmdAfficherSensSequ.Left
         chkVoirPoints.Left = cmdAfficherSensSequ.Left
         chkCouleurProfils.Left = cmdAfficherSensSequ.Left
         cmdAgrandirRetrecir.Left = cmdAfficherSensSequ.Left
         chkZoomProjet.Left = cmdAfficherSensSequ.Left
         pctZoomInfo.Left = chkZoomProjet.Left + 7
         chkZoomDecoupe.Left = cmdAfficherSensSequ.Left
      End If
      If .ScaleHeight < 305 Then
         pctBandeauSaisie.Top = 305 - 27
         If flagFenetreAgrandie = False Then
            pctSequ.Height = Int((305 - pctBandeauSaisie.Height - 3) / 2)
            pctTransf.Top = pctSequ.Top + pctSequ.Height
            pctTransf.Height = 305 - pctSequ.Height - pctBandeauSaisie.Height - 3
         Else
            pctTransf.Top = pctSequ.Top
            pctTransf.Height = 305 - 27 - 1
         End If
      ElseIf .ScaleHeight >= 305 And .ScaleHeight < 674 Then
         pctBandeauSaisie.Top = .ScaleHeight - 27
         If flagFenetreAgrandie = False Then
            pctSequ.Height = Int((.ScaleHeight - pctBandeauSaisie.Height - 3) / 2)
            pctTransf.Top = pctSequ.Top + pctSequ.Height
            pctTransf.Height = .ScaleHeight - pctSequ.Height - pctBandeauSaisie.Height - 3
         Else
            pctTransf.Top = pctSequ.Top
            pctTransf.Height = .ScaleHeight - 27 - 1
         End If
      ElseIf .ScaleHeight >= 674 Then
         pctBandeauSaisie.Top = .ScaleHeight - 27
         If flagFenetreAgrandie = False Then
            pctSequ.Height = 321
            pctTransf.Top = pctSequ.Top + pctSequ.Height
            pctTransf.Height = .ScaleHeight - pctSequ.Height - pctBandeauSaisie.Height - 3
         Else
            pctTransf.Top = pctSequ.Top
            pctTransf.Height = .ScaleHeight - 27 - 1
         End If
      End If
      cmdMiroirFichierSource.Top = pctSequ.Top + pctSequ.Height - 36
      cmdInverserSensSequ.Top = cmdMiroirFichierSource.Top - 33
      cmdAfficherSensSequ.Top = cmdInverserSensSequ.Top - 33
      If flagFenetreAgrandie = False Then
         chkVoirPoints.Top = cmdMiroirFichierSource.Top + 39
      Else
         chkVoirPoints.Top = pctTransf.Top + 2
      End If
      chkCouleurProfils.Top = chkVoirPoints.Top + 33
      cmdAgrandirRetrecir.Top = chkCouleurProfils.Top + 33
      chkZoomProjet.Top = cmdAgrandirRetrecir.Top + 33
      pctZoomInfo.Top = chkZoomProjet.Top + chkZoomProjet.Height + 5
      A = frmMiniCut2d.ScaleHeight - 2
      If A > 0 Then .pctDecoupe.Height = frmMiniCut2d.ScaleHeight - 2
      pctDecoupe.Width = pctSequ.Width
   End With
   pctAjoutPoint.Width = 30
   pctAjoutPoint.Height = 30
   pctAjoutPoint.Left = -1
   pctAjoutPoint.Top = 29
   pctAjoutPoint.Line (8, 9)-(19, 20), vbBlack
   pctAjoutPoint.Line (8, 20)-(19, 9), vbBlack
   'mise à jour des graphiques
   If SSTab1.Tab = 0 Then
      pctSequ.ScaleTop = -pctSequ.ScaleHeight
      Call CalculAffichageInitial
      Call TraceSequ
     ' pctTransf.ScaleTop = pctTransf.ScaleHeight
      If chkZoomProjet.Value = vbChecked Then
         Call EchelleTransf(0, LongBloc, 0, HautBloc)
      ElseIf chkZoomProjet.Value = vbUnchecked Then
         ' Call ZoomAutoToutVoir
      End If
      flagMemoUndoDansTraceTransf = False
      Call TraceTransf
   ElseIf SSTab1.Tab = 1 Then
      Call CalculRepAffDecoupe   'mise à l'échelle de la box
      Call TraceTableDecoupe     'tracé des rectangles et des lignes constitutifs du dessin
      Call TraceBlocEtOrigine    'tracé du bloc, de la zone utile, du cercle de l'origine
      Call TraceDecoupe
   End If
End Sub

Private Sub hscChauffePendantDecoupe_Change()
   If pctDecoupe.Visible = True Then
      hscChauffeStop.Value = hscChauffePendantDecoupe.Value
      hscChauffeDecoupe.Value = hscChauffePendantDecoupe.Value
      hscChauffeFilManuel.Value = hscChauffePendantDecoupe.Value
      ChauffeCourante = hscChauffePendantDecoupe.Value
      lblChauffePendantDecoupe.Caption = Format(ChauffeCourante, "##0") & " %"
      flagModifChauffePendantDecoupe = True 'on passe le flag à True et on gère la chauffe dans la fonction Decoupe()
   End If
End Sub

Private Sub hscChauffeStop_Change()
   If pctDecoupe.Visible = True Then
      hscChauffeDecoupe.Value = hscChauffeStop.Value
      hscChauffeFilManuel.Value = hscChauffeStop.Value
      hscChauffePendantDecoupe.Value = hscChauffeStop.Value
      ChauffeCourante = hscChauffeStop.Value
      lblChauffeStop.Caption = Format(ChauffeCourante, "##0") & " %"
   End If
End Sub

Private Sub hscChauffeStop_Scroll()
   lblChauffeStop.Caption = Format(hscChauffeStop.Value, "##0") & " %"
End Sub

Private Sub hscChauffePendantDecoupe_Scroll()
   lblChauffePendantDecoupe.Caption = Format(hscChauffePendantDecoupe.Value, "##0") & " %"
End Sub

Private Sub hscVitesseManuel_Change()
   hscVitesseBDD.Value = hscVitesseManuel.Value
   VitesseDecoupe = hscVitesseManuel.Value / 10
   lblValeurVitesseManuel.Caption = Format(VitesseDecoupe, "0.0") & " mm/s"
End Sub

Private Sub hscVitesseManuel_Scroll()
   lblValeurVitesseManuel.Caption = Format(hscVitesseManuel.Value / 10, "0.0") & " mm/s"
End Sub


Private Sub optChauffePendantDecoupe_Click(Index As Integer)
   Select Case Index
   Case 0
      optChauffePendantDecoupe(1).ZOrder
      hscChauffePendantDecoupe.Enabled = True
   Case 1
      optChauffePendantDecoupe(0).ZOrder
      hscChauffePendantDecoupe.Enabled = False
   End Select
End Sub

'*************** GESTION DES DECALAGES MODE EXPERT *************************
Private Sub optDecalageExpert_Click(Index As Integer)
   Select Case Index
   Case 0
      lblDecalageExpert.Caption = "0mm"
   Case 1
      lblDecalageExpert.Caption = "-0.5mm (S=1mm)"
   Case 2
      lblDecalageExpert.Caption = "-0.55mm (S=1.1mm)"
   Case 3
      lblDecalageExpert.Caption = "-0.6mm (S=1.2mm)"
   Case 4
      lblDecalageExpert.Caption = "-0.65mm (S=1.3mm)"
   Case 5
      lblDecalageExpert.Caption = "-0.7mm (S=1.4mm)"
   Case 6
      lblDecalageExpert.Caption = "+0.5mm (S=1mm)"
   Case 7
      lblDecalageExpert.Caption = "+0.55mm (S=1.1mm)"
   Case 8
      lblDecalageExpert.Caption = "+0.6mm (S=1.2mm)"
   Case 9
      lblDecalageExpert.Caption = "+0.65mm (S=1.3mm)"
   Case 10
      lblDecalageExpert.Caption = "+0.7mm (S=1.4mm)"
   End Select
   If pctValidationDecoupe.Visible = True And ModeSoft = "Expert" Then 'permet de faire des modifs sans effet quand invisible
      If optDecalageExpert(1).Value = True Then
         SequDecalee = DecalerFil(-0.5)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(2).Value = True Then
         SequDecalee = DecalerFil(-0.55)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(3).Value = True Then
         SequDecalee = DecalerFil(-0.6)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(4).Value = True Then
         SequDecalee = DecalerFil(-0.65)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(5).Value = True Then
         SequDecalee = DecalerFil(-0.7)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(6).Value = True Then
         SequDecalee = DecalerFil(0.5)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(7).Value = True Then
         SequDecalee = DecalerFil(0.55)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(8).Value = True Then
         SequDecalee = DecalerFil(0.6)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(9).Value = True Then
         SequDecalee = DecalerFil(0.65)  'la fonction renvoie une séquence
      ElseIf optDecalageExpert(10).Value = True Then
         SequDecalee = DecalerFil(0.7)  'la fonction renvoie une séquence
      Else
         SequDecalee = SequDecoupe  'pas de décalage
      End If
      Call CalculDepassementCourses(SequDecalee) '(modDepassements) : On calcule le dépassement du bloc ou des courses pour éventuellement modifier le profil
      If flagDepassementDecoupe = True Then
         MsgBox Message(Corps, 52), vbInformation, Message(Titre, 52) 'dépassement des courses, projet tronqué
         flagDepassementDecoupe = False
      End If
   
      Call EntreeSortieSiDecalage
      Call CreerSequMouvement  'on associe les différents trajets, ça marche aussi si optDecalage(2)=true (pas de décalage)
      Call CalculDureeEtSalve
      Call TraceDecoupe
   End If
End Sub

Private Sub optHomeX_Click()
   Dim ReponseFonction As Integer

   optHomeY.Enabled = False
   If optGoManuel(0).Value = True Then
      optGoManuel(1).Value = True 'on stope le mouvement
      optGoManuel(0).ZOrder
   End If
   optGoManuel(0).Enabled = False 'on empêche tout mouvement
   'test du dégagement des inters
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then
      optHomeY.Enabled = True
      Exit Sub   'la machine n'est pas prête
   End If
   
   lblProcedure.Caption = Label(34) '  "Retour automatique horizontal"
   lblProcedure.Visible = True
   lblAvertissementFil2.Caption = Label(18)  '  "LE FIL SE DEPLACE"
   lblAvertissementFil2.BackColor = vbRed
   lblAvertissementFil2.Visible = True
   
   flagAppuiSTOP = False
   flagAppuiStopSansMsgBox = False
   flagPositionPliage = False
   ReponseFonction = 1
   ReponseFonction = RetourOrigine("X", VitesseDecoupe) 'les erreurs sont traités dans DemandeHomeY, qui modifie PasParcourusXX
   If ReponseFonction = 1 Then   'on est à l'origine, on va à la position de repos du fil
      Call MouvementUniquePas(NbrPasToOriXG, 0, NbrPasToOriXD, 0, VitesseDecoupe, NBRt)
      If flagAppuiSTOP = True Then
         flagAppuiSTOP = False
         GoTo Arret_Stop
      End If
      If flagAppuiStopSansMsgBox = True Then
         flagAppuiStopSansMsgBox = False
         GoTo Arret_Stop_Sans_MsgBox
      End If
      If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
      If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
      If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
      lblAvertissementFil2.Caption = Label(38)  ' "Origine horizontale atteinte"
   End If
   optHomeY.Enabled = True
   optAnnulerHome.Value = True
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   optHomeY.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Stop:
   If optChauffe(0).Value = True Then  'si la chauffe est active, on l'arrête
      optChauffe(1).Value = True
      optChauffe(0).ZOrder
   End If
   MsgBox Message(Corps, 25), vbCritical, Message(Titre, 25)  'opération annulée, arrêt d'urgence
   optHomeY.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Stop_Sans_MsgBox:
   If optChauffe(0).Value = True Then  'si la chauffe est active, on la remet
      Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on réactive la chauffe qui a été coupée par le stop
   End If
   optHomeY.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   optHomeY.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   optHomeY.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
End Sub

Private Sub optHomeY_Click()
   Dim ReponseFonction As Integer
   
   optHomeX.Enabled = False
   If optGoManuel(0).Value = True Then
      optGoManuel(1).Value = True 'on stope le mouvement
      optGoManuel(0).ZOrder
   End If
   optGoManuel(0).Enabled = False 'on empêche tout mouvement
   'test du dégagement des inters
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then
      optHomeX.Enabled = True
      Exit Sub   'la machine n'est pas prête
   End If
   
   lblProcedure.Caption = Label(33) '  "Retour automatique vertical"
   lblProcedure.Visible = True
   lblAvertissementFil2.Caption = Label(18)  '  "LE FIL SE DEPLACE"
   lblAvertissementFil2.BackColor = vbRed
   lblAvertissementFil2.Visible = True
   
   flagAppuiSTOP = False
   flagAppuiStopSansMsgBox = False
   flagPositionPliage = False
   ReponseFonction = 1
   ReponseFonction = RetourOrigine("Y", VitesseDecoupe) 'les erreurs sont traités dans DemandeHomeY, qui modifie PasParcourusXX
   If ReponseFonction = 1 Then   'on est à l'origine, on va à la position de repos du fil
      Call MouvementUniquePas(0, -NbrPasToOriYG, 0, -NbrPasToOriYD, VitesseDecoupe, NBRt)
      If flagAppuiSTOP = True Then
         flagAppuiSTOP = False
         GoTo Arret_Stop
      End If
      If flagAppuiStopSansMsgBox = True Then
         flagAppuiStopSansMsgBox = False
         GoTo Arret_Stop_Sans_MsgBox
      End If
      If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
      If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
      If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
      lblAvertissementFil2.Caption = Label(35)  ' "Origine verticale atteinte"
   End If
   optHomeX.Enabled = True
   optAnnulerHome.Value = True
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   optHomeX.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Stop:
   If optChauffe(0).Value = True Then  'si la chauffe est active, on l'arrête
      optChauffe(1).Value = True
      optChauffe(0).ZOrder
   End If
   MsgBox Message(Corps, 25), vbCritical, Message(Titre, 25)  'opération annulée, arrêt d'urgence
   optHomeX.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Stop_Sans_MsgBox:
   If optChauffe(0).Value = True Then  'si la chauffe est active, on la remet
      Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on réactive la chauffe qui a été coupée par le stop
   End If
   optHomeX.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   optHomeX.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   optHomeX.Enabled = True
   optAnnulerHome.Value = True 'à tester !
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True
   Exit Sub
End Sub
Private Sub optAnnulerHome_Click()
   lblProcedure.Visible = False
   flagAppuiStopSansMsgBox = True
   optHomeX.Enabled = True
   optHomeY.Enabled = True
End Sub

Private Sub pctAjoutPoint_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   If Source.Name = "pctAjoutPoint" Then  'pour éviter l'intéraction avec pctTransf et pctSequ
      pctAjoutPoint.AutoRedraw = False
      pctAjoutPoint.Cls
      pctTransf.AutoRedraw = False
      pctTransf.Cls
      VecteurX = x - MemoX
      VecteurY = y - MemoY
      pctAjoutPoint.Line (14 - 4 + VecteurX, 14 - 4 + VecteurY)-(14 + 4 + VecteurX, 14 + 4 + VecteurY), vbBlue
      pctAjoutPoint.Line (14 - 4 + VecteurX, 14 + 4 + VecteurY)-(14 + 4 + VecteurX, 14 - 4 + VecteurY), vbBlue
   End If
End Sub

Private Sub pctAjoutPoint_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Select Case Button
   Case vbLeftButton
      TransfTemp.Etat = 0  'initialisation de l'état de sélection dans Transf
      TransfTemp.NbPoints = 1
      ReDim TransfTemp.Point(1 To 1)
      TransfTemp.Point(1).x = 1
      TransfTemp.Point(1).y = 1
      Call MaxiMiniSequ(TransfTemp)
      MemoX = x
      MemoY = y
      pctAjoutPoint.Drag
   End Select
End Sub

Private Sub pctZoomInfo_Click()
   MsgBox Message(Corps, 55), vbInformation, Message(Titre, 55) 'fonctionnement du zoom
End Sub

Private Sub sliderZoom_Scroll()
   Select Case sliderZoom.Value
   Case -1  'zoom positif : on doit agrandir la figure
      With pctTransf
            .ScaleWidth = 0.9 * .ScaleWidth
            .ScaleLeft = XZoomMolette - Abs(.ScaleWidth / 2)
            .ScaleHeight = 0.9 * .ScaleHeight
            .ScaleTop = YZoomMolette + Abs(.ScaleHeight / 2)
      End With

      sliderZoom.Value = 0
   Case 1   'zoom négatif : on doit rétrécir la figure
      With pctTransf
            .ScaleWidth = 1.1 * .ScaleWidth
            .ScaleLeft = XZoomMolette - Abs(.ScaleWidth / 2)
            .ScaleHeight = 1.1 * .ScaleHeight
            .ScaleTop = YZoomMolette + Abs(.ScaleHeight / 2)
      End With
      sliderZoom.Value = 0
   End Select
   UnPixelToMm = pctTransf.ScaleX(1, vbPixels, vbUser)
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub

Private Sub TimerScreenSaver_Timer()
   'toutes les 30s, on active le screensaver dans la base de registre, ce qui réinitialise le compteur de temps
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, 0, 0
End Sub

Private Sub Tree_Collapse(ByVal Node As MSComctlLib.Node)
   If Node.Key <> ChemRepRacine Then
      Node.Image = "DossierFerme"
   End If
End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
   If Node.Key <> ChemRepRacine Then
      Node.Image = "Dossier"
   End If
End Sub

'****************************************************************
'******* CLIC sur un fichier de contours dans le TREEVIEW *******
'****************************************************************
Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
   Dim NbPoints As Long
   Dim Comment As String
   Dim t() As String
   Dim i As Long, j As Long, k As Long
   Dim Ext As String
   Dim Xgauche As Single, NumPtXGauche As Long
   Dim FF As Integer 'retour de FreeFile pour ouverture de fichier texte
   Dim BufferFichierTexte As String 'pour stockage intégralité fichier texte en mémoire vive
   Dim BufferLignes() As String
   Dim LongueurLigne As Long
   Dim TypeLigneCourante As Boolean, TypeLigneSuivante As Boolean, TypeLignePrecedente As Boolean

   On Error GoTo ErreurOuvertureFichier   'en cas de fichier bidouillé et faisant planter le soft
   
   CheminFichier = Node.Key  'initialisé à "" à l'ouverture du soft
   Ext = ExtraitExtensionFichier(CheminFichier)
   Select Case Ext
   Case "DXF", "dxf", "dat", "DAT", "plt", "PLT", "eps", "EPS"  'fichiers ouverts par cnctool.dll
      '**** appel de CNCTools.dll pour ouverture des autres types de fichier ****
      ReDim profil(1)                           'on vide Profil() avant son remplissage par CNCTools.dll
      Comment = Space$(260) 'initialisation pour reconnaissance pas la dll
      NbPoints = LireFichier(profil(), CheminFichier, Xmin, Xmax, Ymin, Ymax, Comment) 'les variables sont passées
                                                                           'par référence et modifiées par CNCTools.dll
      If NbPoints <= -100 Then  'les code d'erreur de cnctool sont <=-100
         Erase profil
         GoTo ErreurOuvertureFichier
      End If
      Comment = Trim$(Comment)
      Call ExtractionSequ
      Erase profil   'on libère la mémoire
      For i = 1 To NbSequ
         With Sequ(i)
            If .NbPoints > 1 Then 'on doit conserver les points uniques
               If Ext = "dat" Or Ext = "DAT" Then        'Dans le cas d'un .dat, on met la corde sur 100
                  For j = 1 To .NbPoints
                     .Point(j).x = .Point(j).x / .DeltaX * 100
                     .Point(j).y = .Point(j).y / .DeltaX * 100
                  Next j
                  Call MaxiMiniSequ(Sequ(i))
               End If
               If Abs(.Point(.NbPoints).x - .Point(1).x) < 0.001 And Abs(.Point(.NbPoints).y - .Point(1).y) < 0.001 Then
                  'les points d'entrée et sortie sont considérés comme confondus si à moins de 0.001mm en X et Y
                  'dans ce cas, on change le premier point : on prend celui qui est le plus à gauche
                  Xgauche = .Point(1).x
                  NumPtXGauche = 1
                  For k = 1 To .NbPoints
                     If .Point(k).x < Xgauche Then
                        Xgauche = .Point(k).x
                        NumPtXGauche = k
                     End If
                  Next k
                  If NumPtXGauche = .NbPoints Then NumPtXGauche = 1 'on va enlever le dernier point
                  .NbPoints = .NbPoints - 1  'on vire le dernier point
                  ReDim Preserve .Point(1 To .NbPoints)
                  'puis on renumérote
                  ReDim Preserve .Point(1 To .NbPoints + NumPtXGauche - 1)
                  For k = 1 To NumPtXGauche - 1
                     .Point(.NbPoints + k).x = .Point(k).x
                     .Point(.NbPoints + k).y = .Point(k).y
                  Next k
                  For k = 1 To .NbPoints
                     .Point(k).x = .Point(k + NumPtXGauche - 1).x
                     .Point(k).y = .Point(k + NumPtXGauche - 1).y
                  Next k
                  .NbPoints = .NbPoints + 1
                  ReDim Preserve .Point(1 To .NbPoints)
                  .Point(.NbPoints) = .Point(1)
                  '***
               End If
               Call InverserSensSequ(Sequ(i))
               Call MaxiMiniSequ(Sequ(i))
            End If
         End With
      Next i
   Case "mnc", "MNC" 'pour les fichiers projet, on va afficher les séquences individuellement et un assemblage
      NbSequ = 1 + MyVal(LitDecoupeSousFormeIni("ParametresSequences", "NombreSequences", CheminFichier))
      ReDim Sequ(1 To NbSequ)
      ReDim SequTrace(1 To NbSequ)
      For i = 1 To NbSequ - 1
         With Sequ(i)
            .NbPoints = MyVal(LitDecoupeSousFormeIni("Sequence" & LTrim(Str(i)), "NombrePoints", CheminFichier))
            ReDim .Point(1 To .NbPoints)
            ReDim SequTrace(i).Point(1 To .NbPoints)
            For j = 1 To .NbPoints
               t() = Split(LitDecoupeSousFormeIni("Sequence" & LTrim(Str(i)), Str(j), CheminFichier), ":")
               .Point(j).x = MyVal(t(0))
               .Point(j).y = MyVal(t(1))
            Next j
            Call MaxiMiniSequ(Sequ(i))
         End With
      Next i
      Erase t  'on libère la mémoire
      With Sequ(NbSequ) 'définition de la séquence assemblée
         .NbPoints = 0
         For i = 1 To NbSequ - 1
            .NbPoints = .NbPoints + Sequ(i).NbPoints
         Next i
         ReDim Sequ(NbSequ).Point(1 To .NbPoints)
         i = 0
         For j = 1 To NbSequ - 1
            For k = 1 To Sequ(j).NbPoints
               i = i + 1
               .Point(i) = Sequ(j).Point(k)
            Next k
         Next j
         'la séquence assemblée est superposée aux autres séquences, il faut la décaler
         Call MaxiMiniSequ(Sequ(NbSequ))
         For i = 1 To .NbPoints
            .Point(i).x = .Point(i).x + .DeltaX + 0.1 * .DeltaX
         Next i
      End With
   Case "fc", "FC"
      Call LireProfilsFC(CheminFichier)
      'Dans RPFC, l'origine est à droite, on fait un miroir horizontal sans inverser le sens
      For i = 1 To NbSequ
         With Sequ(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Xmin + .Xmax - .Point(j).x
            Next j
         End With
      Next i
   Case "cpx", "CPX"
      Call LireCPX(CheminFichier)
      'Dans Complexes, l'origine est à droite, on fait un miroir horizontal sans inverser le sens
      For i = 1 To NbSequ
         With Sequ(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Xmin + .Xmax - .Point(j).x
            Next j
         End With
      Next i
   Case "txt", "TXT"
      '
      'les fichiers .txt (texte) acceptés sont des fichiers texte du type :
      '
      'Face1
      '2.45:5.432
      '3.44:2.34
      '4.789:-5
      'Face2
      '...
      '
      ' Ouverture du fichier en 'Binary'
      FF = FreeFile
      Open CheminFichier For Binary As #FF
         ' préallocation d'un buffer à la taille du fichier
         BufferFichierTexte = Space$(LOF(FF))
         ' lecture complète du fichier
         Get #FF, , BufferFichierTexte
      Close #FF
      BufferFichierTexte = BufferFichierTexte & vbLf 'au cas où il manquerait
      BufferLignes = Split(BufferFichierTexte, vbLf) 'BufferLignes est un tableau qui contient toutes les lignes
      BufferFichierTexte = "" 'on libère la mémoire
      NbSequ = 0
      i = 0  'LBound(BufferLignes)
      Do
         TypeLigneCourante = DebutDUnNombre(BufferLignes(i)) 'est-ce le début d'un nombre?
         If TypeLigneCourante = True Then 'oui
            'If i > 0 Then TypeLignePrecedente = DebutDUnNombre(BufferLignes(i - 1))
            'If NbSequ = 0 Or (i > 0 And TypeLignePrecedente = False) Then 'soit on est au début du fichier, soit on est après un intervale entre deux séquences
            NbSequ = NbSequ + 1
            ReDim Preserve Sequ(1 To NbSequ)
            ReDim Preserve SequTrace(1 To NbSequ)
            Sequ(NbSequ).NbPoints = 0
            With Sequ(NbSequ)
               Do
                  .NbPoints = .NbPoints + 1
                  ReDim Preserve .Point(1 To .NbPoints)
                  ReDim Preserve SequTrace(NbSequ).Point(1 To .NbPoints)
                  t() = Split(BufferLignes(i), ":")
                  .Point(.NbPoints).x = MyVal(t(0))
                  .Point(.NbPoints).y = MyVal(t(1))
                  If i = UBound(BufferLignes) Then Exit Do 'dernière ligne
                  If DebutDUnNombre(BufferLignes(i + 1)) = False Then
                     'si le début de la ligne suivante n'est pas un nombre,
                     i = i + 1
                     Exit Do
                  End If
                  i = i + 1
               Loop
               Call MaxiMiniSequ(Sequ(NbSequ))
            End With
            'End If
         Else 'la ligne n'est pas une ligne de coordonnées
            If i = UBound(BufferLignes) Then  'si on est arrivé au bout du tableau,
               Exit Do                         'on sort
            End If
            i = i + 1
         End If
      Loop
      Erase t  'on libère la mémoire
      Erase BufferLignes
      If NbSequ = 0 Then 'si le fichier ne contient pas de séquence, on efface juste l'écran de visu
         pctSequ.AutoRedraw = True
         pctSequ.Cls
      Else 'le fichier est représentable et possède des séquences
         'Tous les tracés sont superposés lors d'un export Sketchup, il faut les répartir, on les aligne à droite de la première
         If NbSequ > 1 Then
            i = 1
            Do
               With Sequ(i + 1)
                  For j = 1 To .NbPoints
                     .Point(j).x = .Point(j).x + Sequ(i).Xmax - .Xmin + 0.1 * (Sequ(i).Xmax - Sequ(i).Xmin)
                     .Point(j).y = .Point(j).y - (.Ymin - Sequ(i).Ymin)
                  Next j
               End With
               Call MaxiMiniSequ(Sequ(i + 1))
               i = i + 1
               If i = NbSequ Then Exit Do
            Loop
         End If
         'on renumérote les séquences fermées en partant du point en haut à gauche
         For i = 1 To NbSequ
            With Sequ(i)
               If .NbPoints > 1 Then 'on doit conserver les points uniques
                  If Abs(.Point(.NbPoints).x - .Point(1).x) < 0.001 And Abs(.Point(.NbPoints).y - .Point(1).y) < 0.001 Then
                     'les points d'entrée et sortie sont considérés comme confondus si à moins de 0.001mm en X et Y
                     'dans ce cas, on change le premier point : on prend celui qui est le plus à gauche
                     Xgauche = .Point(1).x
                     NumPtXGauche = 1
                     For k = 1 To .NbPoints
                        If .Point(k).x < Xgauche Then
                           Xgauche = .Point(k).x
                           NumPtXGauche = k
                        End If
                     Next k
                     If NumPtXGauche = .NbPoints Then NumPtXGauche = 1 'on va enlever le dernier point
                     .NbPoints = .NbPoints - 1  'on vire le dernier point
                     ReDim Preserve .Point(1 To .NbPoints)
                     'puis on renumérote
                     ReDim Preserve .Point(1 To .NbPoints + NumPtXGauche - 1)
                     For k = 1 To NumPtXGauche - 1
                        .Point(.NbPoints + k).x = .Point(k).x
                        .Point(.NbPoints + k).y = .Point(k).y
                     Next k
                     For k = 1 To .NbPoints
                        .Point(k).x = .Point(k + NumPtXGauche - 1).x
                        .Point(k).y = .Point(k + NumPtXGauche - 1).y
                     Next k
                     .NbPoints = .NbPoints + 1
                     ReDim Preserve .Point(1 To .NbPoints)
                     .Point(.NbPoints) = .Point(1)
                     '***
                  End If
                  Call InverserSensSequ(Sequ(i))
                  Call MaxiMiniSequ(Sequ(i))
               End If
            End With
         Next i
      End If
   End Select
   
   '****************NETTOYAGE
      For j = 1 To NbSequ
      Call MaxiMiniSequ(Sequ(j))
      With Sequ(j)
         ReDim profil(0 To .NbPoints - 1)
         For i = 1 To .NbPoints
            profil(i - 1).x = .Point(i).x
            profil(i - 1).y = .Point(i).y
         Next i
         If .DeltaX >= .DeltaY Then
            EpsilonNettoyage = .DeltaX / CoeffNett     'CoeffNett défini dans le Form Load
         Else
            EpsilonNettoyage = .DeltaY / CoeffNett
         End If
         NbPoints = Nettoyage(EpsilonNettoyage)     'nettoyage des points trop rapprochés
         .NbPoints = NbPoints
         ReDim .Point(1 To .NbPoints)               'tableau du profil initial, inclus dans le type "Trajet"
         ReDim SequTrace(j).Point(1 To .NbPoints)
         For i = 1 To .NbPoints                      'Transfert des points dans mon type de tableau
            .Point(i).x = profil(i - 1).x
            .Point(i).y = profil(i - 1).y
         Next i
      End With
      Call MaxiMiniSequ(Sequ(j))
   Next j
   'Suppression des points doubles (si point simple dans le dxf, CNCTools envoie deux points avec PU et PD, cf. type)
   For j = 1 To NbSequ
      With Sequ(j)
         If .NbPoints = 2 Then
            If .Point(1).x = .Point(2).x And .Point(1).y = .Point(2).y Then
               ReDim Preserve .Point(1 To 1) 'on vire le doublon
               .NbPoints = 1
            End If
         End If
      End With
   Next j
   Erase profil  'libérer la mémoire
   NumSequSel = 0
   '****************
   
   NumSequSel = 0 'initialisation pour mouse_move
   NbSequSel = 0
   CoeffBloc = 1 'reset car nouveau fichier
   Call CalculAffichageInitial
   Call TraceSequ
   Exit Sub
ErreurOuvertureFichier:
   MsgBox Message(Corps, 10), vbCritical, Message(Titre, 10) 'impossible de lire ce fichier
End Sub

'*********************************************
'********* FERMETURE de l'APPLICATION ********
'*********************************************
Private Sub Form_QueryUnload(Cancel As Integer, unloadmode As Integer)
   Dim ReponseYesNo As Integer, ReponseOKCancel As Integer
   Dim i As Integer, j As Integer
   
   Unload frmImpConv  'on décharge la feuille de vectorisation
   
   If IPL5X_IsConnected() = 1 Then    'l'interpolateur est connecté
      If pctFil.Visible = True Then    'si on est en fenêtre de déplacements, on relève les boutons
         If optGoManuel(0).Value = True Then
            optGoManuel(1).Value = True
         End If
         If optChauffe(0).Value = True Then
            optChauffe(1).Value = True
         End If
      End If
      Do
         Call EnvoiBytes(&H53)   'on envoie le stop pour arrêter le mouvement en cours
         If (ByteIPL(2) And &H2) = &H2 Then Exit Do
      Loop
      Do
         Call EnvoiBytes(&H50, &H0) 'on stope le PWM
         If ByteIPL(2) = &H0 Then Exit Do
      Loop
      If flagPositionPliage = False And flagTableEcriteDansIPL = True Then
         ReponseYesNo = MsgBox(Message(Corps, 8), vbQuestion + vbYesNo, Message(Titre, 8)) 'il semble que la machine n'est pas en position de rangement, y aller?
         If ReponseYesNo = vbYes Then
            flagPliageApresMsgBox = True
            cmdPlierLePortique.Value = True
         End If
      End If
   End If
   If optChauffePendantDecoupe(0).Value = True Then 'on relève le cadenas
      optChauffePendantDecoupe(1).Value = True
   End If
   ReponseOKCancel = MsgBox(Message(Corps, 9), vbExclamation + vbYesNo, Message(Titre, 9)) 'validation pour quitter (sauver, alim...)
   If ReponseOKCancel = vbNo Then
      Cancel = True  'annulation de la fermeture du soft
      Exit Sub
   End If
   'on mémorise la dernière matière utilisée
   MatiereUtilisee.Nom = MatieresAffichees(comboMatieres.ListIndex).Nom
   EcritFichierIni "BDDMatieres", "MatiereUtilisee", MatiereUtilisee.Nom
   'tout vider de la mémoire
   End
End Sub

'***********************************************************************************************
'****** Déplacement du focus sur les textbox après sélection de l'outil avec la touche tab *****
'***********************************************************************************************
Private Sub optOutils_LostFocus(Index As Integer)
'   Select Case Index
'   Case Deplacer
'      If GetTabState Then txtMesures(0).SetFocus
'   End Select
End Sub

'************** Inverser en miroir les séquences sources *****************
Private Sub cmdMiroirFichierSource_Click()
   Dim i As Long, j As Long
   
   If NbSequ > 0 Then
      For i = 1 To NbSequ
         With Sequ(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Xmin + .Xmax - .Point(j).x
            Next j
            Call InverserSensSequ(Sequ(i))
            Call MaxiMiniSequ(Sequ(i))
         End With
      Next i
      NumSequSel = 0 'initialisation pour mouse_move
      NbSequSel = 0
      Call CalculAffichageInitial
      Call TraceSequ
   End If
End Sub
'************* Inverser le sens des séquences sources **********
Private Sub cmdInverserSensSequ_Click()
   Dim i As Long
   
   If NbSequ > 0 Then
      For i = 1 To NbSequ
         Call InverserSensSequ(Sequ(i))
         Call MaxiMiniSequ(Sequ(i))
      Next i
      NumSequSel = 0 'initialisation pour mouse_move
      NbSequSel = 0
      Call CalculAffichageInitial
      Call TraceSequ
   End If
End Sub
'************* Afficher le sens des Sequ *************
Private Sub cmdAfficherSensSequ_Click()
   If flagAfficherSens = True Then
      flagAfficherSens = False
      Call TraceSequ
      Exit Sub
   Else
      flagAfficherSens = True
      Call CalculAffichageInitial
      Call TraceSequ
      Exit Sub
   End If
End Sub


'*************************************************
'****** Ouvrir une feuille de projet vierge ******
'*************************************************
Private Sub cmdNouveauProjet_Click()
   Dim Reponse As Integer
   
   'on doit pouvoir sauver le projet en cours ou annuler la manip OK, NO, CANCEL
   If NbTransf > 0 Then
      Reponse = vbNo
      Reponse = MsgBox(Message(Corps, 11), vbExclamation + vbYesNoCancel, Message(Titre, 11)) 'voulez-vous sauver?
      Select Case Reponse
      Case vbYes
         If flagLeProjetAUnNom = False Then
            Call cmdSauver_Click(0) 'on ouvre la fenêtre de sauvegarde
         Else
            Call cmdSauver_Click(1) 'on écrase le fichier
         End If
      Case vbCancel
         Exit Sub
      End Select
   End If
   'à partir d'ici, le projet courant est écrasé
   chkZoomProjet.Value = vbUnchecked   'on annule le zoom
   NbTransf = 0
   NbTransfSel = 0
   CoeffBloc = 1
   Erase Transf
   flagLeProjetAUnNom = False
   cmdSauver(1).Enabled = False 'désactivation du bouton de sauvegarde simple
   frmMiniCut2d.Caption = "MiniCut2d Software"
   Call EchelleTransf(0, CourseX, 0, CourseY) 'on réinitialise le zoom de la fenêtre du bas
   Call TraceTransf
   Call DesactiverMesures  'il n'y a plus de séquences, on désactive les textbox de saisie manuelle
   SSTab1.Tab = 0 'on repasse en mode création
End Sub

'**********************************************************
'********** OUVRIR un fichier SANS le TreeView ************
'**********************************************************
Private Sub cmdOuvrirFichierSequ_Click()
   Dim NbPoints As Long
   Dim NomFichier As String
   Dim Comment As String
   Dim t() As String
   Dim i As Long, j As Long, k As Long
   Dim Ext As String
   Dim Xgauche As Single, NumPtXGauche As Long
   Dim MatiereProjet As String, ChauffeProjet As Single
   Dim flagMatiereExiste As Boolean
   Dim ReponseOKCancel As Integer
   Dim RepertoireOuvert As String
   'pour ouverture fichier texte
   Dim FF As Integer 'retour de FreeFile pour ouverture de fichier texte
   Dim BufferFichierTexte As String 'pour stockage intégralité fichier texte en mémoire vive
   Dim BufferLignes() As String
   Dim LongueurLigne As Long
   Dim TypeLigneCourante As Boolean, TypeLigneSuivante As Boolean, TypeLignePrecedente As Boolean

   DialogueFichiers.InitDir = LitFichierIni("Fichiers", "DernierRepertoire", App.Path)  'si la clé n'existe pas, on prend App.path
   DialogueFichiers.FileName = ""
   DialogueFichiers.DialogTitle = "Choisissez un fichier profil (.dxf, .dat, .plt, .eps, .txt)"
   DialogueFichiers.CancelError = True
   'on limite le filtre "tous les fichiers" à "tous les fichiers connus" (voir fin de ligne)
   DialogueFichiers.Filter = "MiniCut2d Software (*.mnc)|*.mnc|DAO (*.dxf)|*.dxf|Coordonnées et export SketchUp ou Scratch (*.txt)|*.txt|Profil (*.dat)|*.dat|Plotter (*.plt)|*.plt|Encapsulated PostScript (*.eps)|*.eps|Complexes (*.cpx)|*.cpx|RPFC (*.fc)|*.fc|Tous les fichiers connus|*.mnc;*.dxf;*.txt;*.dat;*.plt;*.cpx;*.fc"
   DialogueFichiers.FilterIndex = 9

   On Error GoTo Annuler                     'si fermeture de la fenêtre sans sélection de fichier
   DialogueFichiers.ShowOpen                              ' afficher la fenêtre d'ouverture
   NomFichier = DialogueFichiers.FileName                 'mémorisation du nom du fichier
   On Error GoTo 0
   Ext = ExtraitExtensionFichier(NomFichier)
   Select Case Ext
   Case "mnc", "MNC"    'il s'agit d'un fichier projet MiniCut2d Software
      'Il faut d'abord écraser l'ancien projet : on fait comme si on ouvrait une feuille blanche
      cmdNouveauProjet.Value = True
      
      'définition du bloc
      LongBloc = MyVal(LitDecoupeSousFormeIni("Bloc", "BlocX", NomFichier))
      HautBloc = MyVal(LitDecoupeSousFormeIni("Bloc", "BlocY", NomFichier))
      'on limite le bloc à la zone utile de la machine
      If LongBloc > CourseX - MargeFil Then LongBloc = CourseX - MargeFil
      If HautBloc > CourseY - MargeFil Then HautBloc = CourseY - MargeFil
      txtBloc(0).Text = Format(HautBloc, "##0")
      txtBloc(1).Text = Format(LongBloc, "##0")
      'lecture des séquences du projet
      NbTransf = MyVal(LitDecoupeSousFormeIni("ParametresSequences", "NombreSequences", NomFichier))
      If NbTransf > 0 Then 'on peut envisager un fichier vide pour attribuer un nom
         ReDim Transf(1 To NbTransf)
         For i = 1 To NbTransf
            With Transf(i)
               .NbPoints = MyVal(LitDecoupeSousFormeIni("Sequence" & LTrim(Str(i)), "NombrePoints", NomFichier))
               ReDim Transf(i).Point(1 To .NbPoints)
               For j = 1 To .NbPoints
                  t() = Split(LitDecoupeSousFormeIni("Sequence" & LTrim(Str(i)), Str(j), NomFichier), ":")
                  .Point(j).x = MyVal(t(0))
                  .Point(j).y = MyVal(t(1))
                  .Point(j).Etat = MyVal(t(2))
               Next j
               Call MaxiMiniSequ(Transf(i))
            End With
         Next i
         Erase t  'on libère la mémoire
      ElseIf NbTransf = 0 Then
         Erase Transf
      End If
      If SSTab1.Tab = 0 Then  'on est dans l'onglet de création
         ' Call ZoomAutoToutVoir
         Call TraceTransf
      ElseIf SSTab1.Tab = 1 Then 'on est dans l'onglet de découpe
         Call InitialisationDessinDecoupe
      End If
      optEntrerBloc(CInt(LitDecoupeSousFormeIni("Decoupe", "TypeEntree", NomFichier))).Value = True
      optSortirBloc(CInt(LitDecoupeSousFormeIni("Decoupe", "TypeSortie", NomFichier))).Value = True
      
      'on mémorise le dernier nom de projet, dans MiniCut2d_Software.ini
      EcritFichierIni "Fichiers", "DernierProjet", NomFichier
      
      'on actualise le titre de la fenêtre
      frmMiniCut2d.Caption = "MiniCut2d Software - " & ExtraitNomFichier(NomFichier)
      flagLeProjetAUnNom = True
      cmdSauver(1).Enabled = True
   Case "dxf", "DXF", "dat", "DAT", "plt", "PLT", "eps", "EPS" 'cnctool va - normalement - réussir à l'ouvrir
      ReDim profil(1)                           'on vide Profil() avant son remplissage par CNCTools.dll
      '**** appel de CNCTools.dll pour ouverture ****
      Comment = Space$(260) 'initialisation pour reconnaissance pas la dll
      NbPoints = LireFichier(profil(), NomFichier, Xmin, Xmax, Ymin, Ymax, Comment) 'les variables sont passées
                                                                           'par référence et modifiées par CNCTools.dll
      If NbPoints <= -100 Then  'les code d'erreur de cnctool sont <=-100
         Erase profil
         GoTo ErreurOuvertureFichier
      End If
      Comment = Trim$(Comment)
      Call ExtractionSequ
      Erase profil  'on libère la mémoire
      For i = 1 To NbSequ
         With Sequ(i)
            If .NbPoints > 1 Then 'on doit conserver les points uniques
               If Ext = "dat" Or Ext = "DAT" Then        'Dans le cas d'un .dat, on met la corde sur 100
                  For j = 1 To .NbPoints
                     .Point(j).x = .Point(j).x / .DeltaX * 100
                     .Point(j).y = .Point(j).y / .DeltaX * 100
                  Next j
                  Call MaxiMiniSequ(Sequ(i))
               End If
               If Abs(.Point(.NbPoints).x - .Point(1).x) < 0.001 And Abs(.Point(.NbPoints).y - .Point(1).y) < 0.001 Then
                  'les points d'entrée et sortie sont considérés comme confondus si à moins de 0.001mm en X et Y
                  'dans ce cas, on change le premier point : on prend celui qui est le plus à gauche
                  Xgauche = .Point(1).x
                  NumPtXGauche = 1
                  For k = 1 To .NbPoints
                     If .Point(k).x < Xgauche Then
                        Xgauche = .Point(k).x
                        NumPtXGauche = k
                     End If
                  Next k
                  If NumPtXGauche = .NbPoints Then NumPtXGauche = 1 'on va enlever le dernier point
                  .NbPoints = .NbPoints - 1  'on vire le dernier point
                  ReDim Preserve .Point(1 To .NbPoints)
                  'puis on renumérote
                  ReDim Preserve .Point(1 To .NbPoints + NumPtXGauche - 1)
                  For k = 1 To NumPtXGauche - 1
                     .Point(.NbPoints + k).x = .Point(k).x
                     .Point(.NbPoints + k).y = .Point(k).y
                  Next k
                  For k = 1 To .NbPoints
                     .Point(k).x = .Point(k + NumPtXGauche - 1).x
                     .Point(k).y = .Point(k + NumPtXGauche - 1).y
                  Next k
                  .NbPoints = .NbPoints + 1
                  ReDim Preserve .Point(1 To .NbPoints)
                  .Point(.NbPoints) = .Point(1)
                  '***
               End If
               Call InverserSensSequ(Sequ(i))
               Call MaxiMiniSequ(Sequ(i))
            End If
         End With
      Next i
   Case "fc", "FC"  'ouverture d'un fichier RP-FC
      Call LireProfilsFC(NomFichier)
      'Dans RPFC, l'origine est à droite, on fait un miroir horizontal sans inverser le sens
      For i = 1 To NbSequ
         With Sequ(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Xmin + .Xmax - .Point(j).x
            Next j
         End With
      Next i
   Case "cpx", "CPX"  'ouverture d'un fichier Complexes
      Call LireCPX(NomFichier)
      'Dans Complexes, l'origine est à droite, on fait un miroir horizontal sans inverser le sens
      For i = 1 To NbSequ
         With Sequ(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Xmin + .Xmax - .Point(j).x
            Next j
         End With
      Next i
   Case "txt", "TXT"
      'les fichiers .txt (texte) acceptés sont des fichiers texte du type :
      '
      'Face1
      '2.45:5.432
      '3.44:2.34
      '4.789:-5
      'Face2
      '...
      '
      ' Ouverture du fichier en 'Binary'
      FF = FreeFile
      Open NomFichier For Binary As #FF
         ' préallocation d'un buffer à la taille du fichier
         BufferFichierTexte = Space$(LOF(FF))
         ' lecture complète du fichier
         Get #FF, , BufferFichierTexte
      Close #FF
      BufferFichierTexte = BufferFichierTexte & vbLf 'au cas où il manquerait
      BufferLignes = Split(BufferFichierTexte, vbLf) 'BufferLignes est un tableau qui contient toutes les lignes
      BufferFichierTexte = "" 'on libère la mémoire
      NbSequ = 0
      i = 0  'LBound(BufferLignes)
      Do
         TypeLigneCourante = DebutDUnNombre(BufferLignes(i)) 'est-ce le début d'un nombre?
         If TypeLigneCourante = True Then 'oui
            'If i > 0 Then TypeLignePrecedente = DebutDUnNombre(BufferLignes(i - 1))
            'If NbSequ = 0 Or (i > 0 And TypeLignePrecedente = False) Then 'soit on est au début du fichier, soit on est après un intervale entre deux séquences
            NbSequ = NbSequ + 1
            ReDim Preserve Sequ(1 To NbSequ)
            ReDim Preserve SequTrace(1 To NbSequ)
            Sequ(NbSequ).NbPoints = 0
            With Sequ(NbSequ)
               Do
                  .NbPoints = .NbPoints + 1
                  ReDim Preserve .Point(1 To .NbPoints)
                  ReDim Preserve SequTrace(NbSequ).Point(1 To .NbPoints)
                  t() = Split(BufferLignes(i), ":")
                  .Point(.NbPoints).x = MyVal(t(0))
                  .Point(.NbPoints).y = MyVal(t(1))
                  If i = UBound(BufferLignes) Then Exit Do 'dernière ligne
                  If DebutDUnNombre(BufferLignes(i + 1)) = False Then
                     'si le début de la ligne suivante n'est pas un nombre,
                     i = i + 1
                     Exit Do
                  End If
                  i = i + 1
               Loop
               Call MaxiMiniSequ(Sequ(NbSequ))
            End With
            'End If
         Else 'la ligne n'est pas une ligne de coordonnées
            If i = UBound(BufferLignes) Then  'si on est arrivé au bout du tableau,
               Exit Do                         'on sort
            End If
            i = i + 1
         End If
      Loop
      Erase t  'on libère la mémoire
      Erase BufferLignes
      If NbSequ = 0 Then 'si le fichier ne contient pas de séquence, on efface juste l'écran de visu
         pctSequ.AutoRedraw = True   'le tracé des séquences se fait sur le plan permanent
         pctSequ.Cls
      Else
         'Tous les tracés sont superposés lors d'un export Sketchup, il faut les répartir, on les aligne à droite de la première
         If NbSequ > 1 Then
            i = 1
            Do
               With Sequ(i + 1)
                  For j = 1 To .NbPoints
                     .Point(j).x = .Point(j).x + Sequ(i).Xmax - .Xmin + 0.1 * (Sequ(i).Xmax - Sequ(i).Xmin)
                     .Point(j).y = .Point(j).y - (.Ymin - Sequ(i).Ymin)
                  Next j
               End With
               Call MaxiMiniSequ(Sequ(i + 1))
               i = i + 1
               If i = NbSequ Then Exit Do
            Loop
         End If
         'on renumérote les séquences fermées en partant du point en haut à gauche
         For i = 1 To NbSequ
            With Sequ(i)
               If .NbPoints > 1 Then 'on doit conserver les points uniques
                  If Abs(.Point(.NbPoints).x - .Point(1).x) < 0.001 And Abs(.Point(.NbPoints).y - .Point(1).y) < 0.001 Then
                     'les points d'entrée et sortie sont considérés comme confondus si à moins de 0.001mm en X et Y
                     'dans ce cas, on change le premier point : on prend celui qui est le plus à gauche
                     Xgauche = .Point(1).x
                     NumPtXGauche = 1
                     For k = 1 To .NbPoints
                        If .Point(k).x < Xgauche Then
                           Xgauche = .Point(k).x
                           NumPtXGauche = k
                        End If
                     Next k
                     If NumPtXGauche = .NbPoints Then NumPtXGauche = 1 'on va enlever le dernier point
                     .NbPoints = .NbPoints - 1  'on vire le dernier point
                     ReDim Preserve .Point(1 To .NbPoints)
                     'puis on renumérote
                     ReDim Preserve .Point(1 To .NbPoints + NumPtXGauche - 1)
                     For k = 1 To NumPtXGauche - 1
                        .Point(.NbPoints + k).x = .Point(k).x
                        .Point(.NbPoints + k).y = .Point(k).y
                     Next k
                     For k = 1 To .NbPoints
                        .Point(k).x = .Point(k + NumPtXGauche - 1).x
                        .Point(k).y = .Point(k + NumPtXGauche - 1).y
                     Next k
                     .NbPoints = .NbPoints + 1
                     ReDim Preserve .Point(1 To .NbPoints)
                     .Point(.NbPoints) = .Point(1)
                     '***
                  End If
                  Call InverserSensSequ(Sequ(i))
                  Call MaxiMiniSequ(Sequ(i))
               End If
            End With
         Next i
      End If
   Case Else
      'on n'est pas sensé arriver ici, mais c'est une précaution
      MsgBox Message(Corps, 12), vbCritical, Message(Titre, 12)  'extension non valide
      Exit Sub
   End Select
   '****************NETTOYAGE
   For j = 1 To NbSequ
      Call MaxiMiniSequ(Sequ(j))
      With Sequ(j)
         ReDim profil(0 To .NbPoints - 1)
         For i = 1 To .NbPoints
            profil(i - 1).x = .Point(i).x
            profil(i - 1).y = .Point(i).y
         Next i
         If .DeltaX >= .DeltaY Then
            EpsilonNettoyage = .DeltaX / CoeffNett     'CoeffNett défini dans le Form Load
         Else
            EpsilonNettoyage = .DeltaY / CoeffNett
         End If
         NbPoints = Nettoyage(EpsilonNettoyage)     'nettoyage des points trop rapprochés
         .NbPoints = NbPoints
         ReDim .Point(1 To .NbPoints)               'tableau du profil initial, inclus dans le type "Trajet"
         ReDim SequTrace(j).Point(1 To .NbPoints)
         For i = 1 To .NbPoints                      'Transfert des points dans mon type de tableau
            .Point(i).x = profil(i - 1).x
            .Point(i).y = profil(i - 1).y
         Next i
      End With
      Call MaxiMiniSequ(Sequ(j))
   Next j
   'Suppression des points doubles (si point simple dans le dxf, CNCTools envoie deux points avec PU et PD, cf. type)
   For j = 1 To NbSequ
      With Sequ(j)
         If .NbPoints = 2 Then
            If .Point(1).x = .Point(2).x And .Point(1).y = .Point(2).y Then
               ReDim Preserve .Point(1 To 1) 'on vire le doublon
               .NbPoints = 1
            End If
         End If
      End With
   Next j
   Erase profil  'libérer la mémoire
   '****************
   NumSequSel = 0 'initialisation pour mouse_move
   NbSequSel = 0
   CoeffBloc = 1 'reset car nouveau fichier
   Call CalculAffichageInitial
   Call TraceSequ
   'on mémorise le dernier dossier utilisé dans MiniCut2d_Software.ini
   RepertoireOuvert = GetPathName(NomFichier)
   EcritFichierIni "Fichiers", "DernierRepertoire", RepertoireOuvert
   EcritFichierIni "Fichiers", "DernierProjet", NomFichier
   Exit Sub
ErreurOuvertureFichier:
   MsgBox Message(Corps, 10), vbCritical, Message(Titre, 10) 'impossible de lire ce fichier
   Exit Sub
Annuler:
End Sub

'********************************************************************************
'********* CALCUL DES COEFFICIENTS de l'AFFICHAGE DU FICHIER INITIAL ************
'********************************************************************************
Public Sub CalculAffichageInitial()
   Dim i As Long, j As Long
   
   If NbSequ > 0 Then
      '*** calcul des maxi et mini du total des séquences pour représentation graphique ***
      XminSequ = Sequ(1).Point(1).x
      XmaxSequ = Sequ(1).Point(1).x
      YminSequ = Sequ(1).Point(1).y
      YmaxSequ = Sequ(1).Point(1).y
      For i = 1 To NbSequ
         For j = 1 To Sequ(i).NbPoints
            If Sequ(i).Point(j).x < XminSequ Then XminSequ = Sequ(i).Point(j).x
            If Sequ(i).Point(j).x > XmaxSequ Then XmaxSequ = Sequ(i).Point(j).x
            If Sequ(i).Point(j).y < YminSequ Then YminSequ = Sequ(i).Point(j).y
            If Sequ(i).Point(j).y > YmaxSequ Then YmaxSequ = Sequ(i).Point(j).y
         Next j
      Next i
      XcentreSequ = (XmaxSequ + XminSequ) / 2
      YcentreSequ = (YmaxSequ + YminSequ) / 2
      If XmaxSequ = XminSequ And YmaxSequ = YminSequ Then    'tous les points au même point!
         CoeffSequ = 1
      ElseIf XmaxSequ = XminSequ Then    'tous les points sur une droite verticale
         CoeffSequ = (pctSequ.Height - 2 * MargInit) / (Abs(YmaxSequ - YminSequ))
      ElseIf YmaxSequ = YminSequ Then    'tous les points sur une droite horizontale
         CoeffSequ = (pctSequ.Width - 2 * MargInit) / (Abs(XmaxSequ - XminSequ))
      Else     'cas classique
         If (pctSequ.Width - 2 * MargInit) / (Abs(XmaxSequ - XminSequ)) < (pctSequ.Height - 2 * MargInit) / (Abs(YmaxSequ - YminSequ)) Then
            CoeffSequ = (pctSequ.Width - 2 * MargInit) / (Abs(XmaxSequ - XminSequ))
         Else
            CoeffSequ = (pctSequ.Height - 2 * MargInit) / (Abs(YmaxSequ - YminSequ))
         End If
      End If
      '*** Création du tableau de représentation ***
      SequTrace = Sequ
      For i = 1 To NbSequ
         With SequTrace(i)
            For j = 1 To .NbPoints
               .Point(j).x = (Sequ(i).Point(j).x - XcentreSequ) * CoeffSequ + pctSequ.Width / 2
               .Point(j).y = (Sequ(i).Point(j).y - YcentreSequ) * CoeffSequ + pctSequ.Height / 2
            Next j
            Call MaxiMiniSequ(SequTrace(i))   'calcul des maxi et mini de toutes les séquences tracées
         End With
      Next i
   End If
End Sub

'**********************************************************
'********* TRACE DES LIGNES DU FICHIER INITIAL ************
'**********************************************************
Public Sub TraceSequ()
   Dim i As Long, j As Long
   Dim XFS1 As Single, YFS1 As Single, XFS2 As Single, YFS2 As Single, XFS As Single, YFS As Single
   
   If NbSequ > 0 Then
      pctSequ.AutoRedraw = True   'le tracé des séquences se fait sur le plan permanent
      pctSequ.Cls
      'séquences initiales
      For i = 1 To NbSequ
         If Sequ(i).NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
            With SequTrace(i)
               pctSequ.Line (.Point(1).x - 4, .Point(1).y - 4)-(.Point(1).x + 4, .Point(1).y + 4), vbBlack
               pctSequ.Line (.Point(1).x - 4, .Point(1).y + 4)-(.Point(1).x + 4, .Point(1).y - 4), vbBlack
            End With
         Else           'gestion des séquences consituées de plusieurs points
            With SequTrace(i)
               For j = 1 To Sequ(i).NbPoints - 1
                  pctSequ.Line (.Point(j).x, .Point(j).y)-(.Point(j + 1).x, .Point(j + 1).y), vbBlack
                  If .Point(j).Mark = True Then
                     pctSequ.Circle (.Point(j).x, .Point(j).y), 4, vbRed
                  End If
               Next j
               XFS1 = .Point(1).x
               YFS1 = .Point(1).y
               XFS2 = .Point(2).x
               YFS2 = .Point(2).y
            End With
            If flagAfficherSens = True Then
               With pctSequ
                  .FillColor = RGB(255, 107, 107)
                  .FillStyle = 0  'rempli
                  .DrawMode = vbMaskPen
                  'tracé de la flèche
                  If YFS2 <> YFS1 Or XFS2 <> XFS1 Then
                     XFS = 10 * (XFS2 - XFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     YFS = 10 * (YFS2 - YFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     .DrawWidth = 2
                     pctSequ.Line (XFS1, YFS1)-Step(XFS, YFS), RGB(230, 61, 61) ' corps de la flèche
                     .DrawWidth = 1
                  End If
                  .DrawMode = vbCopyPen
                  'tracé du point d'entrée
                  .FillStyle = 0 'rempli
                  .FillColor = vbRed
                  pctSequ.Circle (XFS1, YFS1), 2, vbRed
                  .FillStyle = 1 'transparent
               End With
            End If
         End If
      Next i
      pctSequ.AutoRedraw = False   'par défaut on est sur le plan temporaire
   End If
End Sub

'**************************************************************************************************
'***** Définition des échelles d'affichage dans la picturebox de transfert (ZOOM AUTOMATIQUE) *****
'**************************************************************************************************
Private Sub EchelleTransf(ByVal XminFigure As Single, ByVal XmaxFigure As Single, ByVal YminFigure As Single, ByVal YmaxFigure As Single)
   Dim LongueurFigure As Single, HauteurFigure As Single
   Dim MargeAutour   As Single     'sert à voir les bords de la figure
   
      LongueurFigure = Abs(XmaxFigure - XminFigure)
      HauteurFigure = Abs(YmaxFigure - YminFigure)
      'Calcul des échelles d'affichage dans la picturebox de transfert : on représente
      ' la surface utile + XXmm tout autour
      MargeAutour = 10
      With pctTransf
         If .Width / (LongueurFigure + 2 * MargeAutour) > .Height / (HauteurFigure + 2 * MargeAutour) Then
            .ScaleHeight = -(HauteurFigure + 2 * MargeAutour) '  le "-" devant inverse le sens de l'axe
            .ScaleTop = YmaxFigure + MargeAutour
            .ScaleWidth = .Width * Abs(.ScaleHeight) / .Height
            .ScaleLeft = -MargeAutour + XminFigure - (Abs(.ScaleWidth) - (LongueurFigure + 2 * MargeAutour)) / 2
         Else
            .ScaleLeft = XminFigure - MargeAutour
            .ScaleWidth = LongueurFigure + 2 * MargeAutour
            .ScaleHeight = -.Height * .ScaleWidth / .Width
            .ScaleTop = YmaxFigure + MargeAutour + (Abs(.ScaleHeight) - Abs(HauteurFigure + 2 * MargeAutour)) / 2
         End If
      End With
         
      UnPixelToMm = pctTransf.ScaleX(1, vbPixels, vbUser)
      
End Sub

'*******************************************************
'******** Transfert de séquence par double-clic ********
'*******************************************************
Private Sub pctSequ_DblClick()
   Dim j As Long, i As Long, k As Long
   Dim P1 As PointProfil, P2 As PointProfil
   Dim DeltaX As Single, DeltaY As Single
   Dim EpsilonPoint As Single
   
   '**** ATTENTION : il faut tenir compte du sens de rotation des profils ****
   'contrairement aux .dat, les profils tournent maintenant dans le sens horaire, avec l'entrée à gauche
   
   Jonction = "BC" 'pour info, équivalence par rapport au Drag and Drop
   
   If NumSequSel <> 0 Then
      TransfTemp = Sequ(NumSequSel)
      TransfTemp.Etat = 0     'initialisation de l'état
      With TransfTemp
         For i = 1 To .NbPoints
            .Point(i).x = .Point(i).x * CoeffBloc  'mise à l'échelle liée au zoom auto
            .Point(i).y = .Point(i).y * CoeffBloc
         Next i
      End With
      Call MaxiMiniSequ(TransfTemp)
      If NbTransfSel = 0 Then 'pas de séquence sélectionnée, on se met à la fin
         NbTransf = NbTransf + 1
         ReDim Preserve Transf(1 To NbTransf)
         Transf(NbTransf) = TransfTemp
         If NbTransf = 1 Then 'si c'est la première séquence, on la met au début du bloc, moins les marges
            DeltaX = (MargeFil + MargeInterieureX) - Transf(1).Xmin
            DeltaY = MargeInterieureY - Transf(1).Ymin
         ElseIf NbTransf > 1 Then  'si on a déjà d'autres séquences, on met la nouvelle à une distance égale à 10% de sa largeur
            If Transf(NbTransf).NbPoints = 1 Then
               EpsilonPoint = CourseX / 20
            Else
               EpsilonPoint = 0
            End If
            P1.x = Transf(NbTransf - 1).Xmax + 0.1 * (Transf(NbTransf).Xmax - Transf(NbTransf).Xmin) + EpsilonPoint
            P1.y = Transf(NbTransf - 1).Point(Transf(NbTransf - 1).NbPoints).y 'P1 et P2 sont les extrémités du vecteur de translation
            P2.x = Transf(NbTransf).Xmin
            P2.y = Transf(NbTransf).Point(1).y
            DeltaX = P1.x - P2.x
            DeltaY = P1.y - P2.y
         End If
         With Transf(NbTransf)
            For j = 1 To .NbPoints
               .Point(j).x = .Point(j).x + DeltaX
               .Point(j).y = .Point(j).y + DeltaY
            Next j
         End With
         Call MaxiMiniSequ(Transf(NbTransf))
      Else  'il y a des séquences sélectionnées, on insère après elles
         i = 0
         Do
            i = i + 1
            If (Transf(i).Etat And 1) = 1 Then
               NbTransf = NbTransf + 1
               ReDim Preserve Transf(1 To NbTransf)
               For j = NbTransf To i + 1 Step -1
                  Transf(j) = Transf(j - 1)
               Next j
               Transf(i + 1) = TransfTemp
                  
               If Transf(i + 1).NbPoints = 1 Then
                  EpsilonPoint = CourseX / 20
               Else
                  EpsilonPoint = 0
               End If
               P1.x = Transf(i).Xmax + 0.1 * (Transf(i + 1).Xmax - Transf(i + 1).Xmin) + EpsilonPoint
               P1.y = Transf(i).Point(Transf(i).NbPoints).y 'P1 et P2 sont les extrémités du vecteur de translation
               P2.x = Transf(i + 1).Xmin
               P2.y = Transf(i + 1).Point(1).y
               DeltaX = P1.x - P2.x
               DeltaY = P1.y - P2.y
               With Transf(i + 1)   'cette translation sert à passer du repère du fichier source au repère du projet
                  For j = 1 To .NbPoints
                     .Point(j).x = .Point(j).x + DeltaX
                     .Point(j).y = .Point(j).y + DeltaY
                  Next j
                  Call MaxiMiniSequ(Transf(i + 1))
               End With
               If NbTransf > i + 1 Then
                  If Transf(i + 2).NbPoints = 1 Then
                     EpsilonPoint = CourseX / 20
                  Else
                     EpsilonPoint = 0
                  End If         'on écarte seulement suivant X pour faire de la place
                  P1.x = Transf(i + 1).Xmax + 0.1 * (Transf(i + 2).Xmax - Transf(i + 1).Xmin) + EpsilonPoint
              '    P1.Y = Transf(i + 1).Point(Transf(i + 1).NbPoints).Y 'P1 et P2 sont les extrémités du vecteur de translation
                  P2.x = Transf(i + 2).Xmin
                  P2.y = Transf(i + 2).Point(1).y
                  DeltaX = P1.x - P2.x
               '   DeltaY = P1.Y - P2.Y
                  For k = i + 2 To NbTransf
                     With Transf(k)
                        For j = 1 To .NbPoints
                           .Point(j).x = .Point(j).x + DeltaX
                '           .Point(j).Y = .Point(j).Y + DeltaY
                        Next j
                     End With
                     Call MaxiMiniSequ(Transf(k))
                  Next k
               End If
               If i = NbTransf - 1 Then Exit Do
            End If
            If i = NbTransf Then Exit Do
         Loop
                  
      End If
      If NbTransf > 0 Then
         Call AppliquerCentrageEtAjustement
         ' Call ZoomAutoToutVoir 'on vérifie si ça dépasse
         Call TraceTransf
      End If
   End If
End Sub

Private Sub MaxiMiniTotalTransff()
   Dim i As Long, j As Long
   
   If NbTransf > 0 Then
      '*** calcul des maxi et mini du total des séquences Transférées ***
      XminTransf = Transf(1).Point(1).x
      XmaxTransf = Transf(1).Point(1).x
      YminTransf = Transf(1).Point(1).y
      YmaxTransf = Transf(1).Point(1).y
      For i = 1 To NbTransf
         For j = 1 To Transf(i).NbPoints
            If Transf(i).Point(j).x < XminTransf Then XminTransf = Transf(i).Point(j).x
            If Transf(i).Point(j).x > XmaxTransf Then XmaxTransf = Transf(i).Point(j).x
            If Transf(i).Point(j).y < YminTransf Then YminTransf = Transf(i).Point(j).y
            If Transf(i).Point(j).y > YmaxTransf Then YmaxTransf = Transf(i).Point(j).y
         Next j
      Next i
   Else
      XminTransf = 0
      XmaxTransf = CourseX
      YminTransf = 0
      YmaxTransf = CourseY
   End If
End Sub

'*************************************************************
'********* TRACE DES LIGNES DES Séquences Transférées ********
'*************************************************************
'*** Si Memo=true, on ne trace pas les séquences sélectionnées, ce qui permet de les voir bouger avec le DragOver
Public Sub TraceTransf(Optional ByVal Memo As Boolean = False)
   Dim i As Long, j As Long
   Dim XFS1 As Single, YFS1 As Single, XFS2 As Single, YFS2 As Single, XFS As Single, YFS As Single
   Dim XS1 As Single, YS1 As Single, XS2 As Single, YS2 As Single
   Dim Color As Long, Epaisseur As Integer, flagColor As Integer
   Dim RayonPoints As Single, UnPixel As Single
   
   RayonPoints = 3 * pctTransf.ScaleWidth / pctTransf.Width 'toujours 3 pixels
   UnPixel = pctTransf.ScaleWidth / pctTransf.Width 'pour être insensible au zoom
   pctTransf.AutoRedraw = True
   pctTransf.FillStyle = 1  'transparent
   pctTransf.DrawMode = vbCopyPen
   pctTransf.DrawWidth = 1
   pctTransf.Cls
   'rectangle de la surface utile
   pctTransf.DrawStyle = 2
   pctTransf.Line (0, MargePlateau)-(CourseX, CourseY - MargeFil), vbBlack, B
   'rectangle du bloc
   pctTransf.DrawStyle = 0
   pctTransf.Line (MargeFil, 0)-(LongBloc + MargeFil, HautBloc), RGB(110, 110, 219), B
   If flagMemoUndoDansTraceTransf = True Then
      Call MemoriserPourUndo
   End If
   flagMemoUndoDansTraceTransf = True
   If NbTransf > 0 Then
      '**** On représente les séquences du tableau Transf() ****
      flagColor = 0
      For i = 1 To NbTransf
         'gestion des couleurs et de l'épaisseur : d'abord comme si la séquence n'était pas sélectionnée
         If chkCouleurProfils.Value = vbUnchecked Then  'toutes les séquences de la même couleur
            With Transf(i)
               If (.Etat And 1) = 1 Then     'si la séquence est sélectionnée, on écrase les valeurs précédentes
                  Color = vbBlue
                  Epaisseur = 2
               Else
                  Color = vbBlack
                  Epaisseur = 1
               End If
            End With
         ElseIf chkCouleurProfils.Value = vbChecked Then 'alternance de couleurs
            If flagColor = 0 Then
               With Transf(i)
                  If (.Etat And 1) = 1 Then     'si la séquence est sélectionnée, on écrase les valeurs précédentes
                     Color = vbBlue
                     Epaisseur = 2
                  Else
                     Color = vbBlack
                     Epaisseur = 1
                  End If
                  flagColor = 1
               End With
            ElseIf flagColor = 1 Then
               With Transf(i)
                  If (.Etat And 1) = 1 Then     'si la séquence est sélectionnée, on écrase les valeurs précédentes
                     Color = RGB(0, 169, 207)
                     Epaisseur = 2
                  Else
                     Color = vbRed
                     Epaisseur = 1
                  End If
                  flagColor = 0
               End With
            End If
         End If
         With Transf(i)
            If Memo = False Or (Memo = True And (.Etat And 1) = 0) Then
               If Transf(i).NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
                  pctTransf.DrawWidth = Epaisseur
                  pctTransf.Line (.Point(1).x - 4 * UnPixel, .Point(1).y - 4 * UnPixel)-(.Point(1).x + 4 * UnPixel, .Point(1).y + 4 * UnPixel), Color
                  pctTransf.Line (.Point(1).x - 4 * UnPixel, .Point(1).y + 4 * UnPixel)-(.Point(1).x + 4 * UnPixel, .Point(1).y - 4 * UnPixel), Color
                  pctTransf.DrawWidth = 1
               Else           'gestion des séquences consituées de plusieurs points
                  pctTransf.DrawWidth = Epaisseur
                  For j = 1 To Transf(i).NbPoints - 1
                        pctTransf.Line (.Point(j).x, .Point(j).y)-(.Point(j + 1).x, .Point(j + 1).y), Color
                  Next j
                  pctTransf.DrawWidth = 1
               End If
               'affichage du sens sur la première séquence si elle est formée de plusieurs points et que les deux premiers sont distincts
               If i = 1 And .NbPoints > 1 Then
                  pctTransf.FillColor = RGB(255, 107, 107)
                  pctTransf.FillStyle = 0  'rempli
                  pctTransf.DrawMode = vbMaskPen
                  'tracé de la flèche
                  XFS1 = .Point(1).x
                  YFS1 = .Point(1).y
                  XFS2 = .Point(2).x
                  YFS2 = .Point(2).y
                  If YFS2 <> YFS1 Or XFS2 <> XFS1 Then
                     XFS = 10 * UnPixel * (XFS2 - XFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     YFS = 10 * UnPixel * (YFS2 - YFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     pctTransf.DrawWidth = 2
                     pctTransf.Line (XFS1, YFS1)-Step(XFS, YFS), RGB(230, 61, 61) ' corps de la flèche
                     pctTransf.DrawWidth = 1
                  End If
                  pctTransf.FillStyle = 1  'transparent
                  pctTransf.DrawMode = vbCopyPen
                  pctTransf.DrawWidth = 1
               End If
               'Affichage des points si demandé
               If chkVoirPoints.Value = vbChecked Then
                  For j = 1 To Transf(i).NbPoints 'points des profils
                     pctTransf.Circle (.Point(j).x, .Point(j).y), RayonPoints, RGB(144, 0, 82)
                  Next j
               End If
               'tracé des segments de liaison
               If i < NbTransf Then
                  If Memo = False Or (Memo = True And (Transf(i + 1).Etat And 1) = 0) Then
                     If .Point(.NbPoints).x <> Transf(i + 1).Point(1).x Or _
                           .Point(.NbPoints).y <> Transf(i + 1).Point(1).y Then
                           XS1 = .Point(.NbPoints).x
                           YS1 = .Point(.NbPoints).y
                           XS2 = Transf(i + 1).Point(1).x
                           YS2 = Transf(i + 1).Point(1).y
                           pctTransf.Line (XS1, YS1)-(XS2, YS2), RGB(0, 122, 11)
                     End If
                  End If
               End If
            End If
         End With
      Next i
      'tracé du point d'entrée
      If Memo = False Or (Memo = True And (Transf(1).Etat And 1) = 0) Then
         pctTransf.FillStyle = 0 'rempli
         pctTransf.FillColor = vbRed
         pctTransf.DrawWidth = 1
         pctTransf.Circle (Transf(1).Point(1).x, Transf(1).Point(1).y), RayonPoints, vbRed
         pctTransf.FillStyle = 1 'transparent
      End If
       'tracé du point de sortie
      If Memo = False Or (Memo = True And (Transf(NbTransf).Etat And 1) = 0) Then
         pctTransf.FillStyle = 1 'transparent
         With Transf(NbTransf)
            pctTransf.Circle (.Point(.NbPoints).x, .Point(.NbPoints).y), 1.5 * RayonPoints, RGB(0, 144, 61)
         End With
      End If
      ' calcul des dimensions de la sélection pour affichage rectangle et texte
      If NbTransfSel > 0 Then
         For i = 1 To NbTransf   'la première boucle sert à initialiser les variables des maxi/mini
            If (Transf(i).Etat And 1) = 1 Then
               XminSel = Transf(i).Xmin
               XmaxSel = Transf(i).Xmax
               YminSel = Transf(i).Ymin
               YmaxSel = Transf(i).Ymax
               Exit For
            End If
         Next i
         For i = 1 To NbTransf   'la seconde boucle est la boucle de comparaison
            With Transf(i)
               If (.Etat And 1) = 1 Then
                  If .Xmin < XminSel Then
                      XminSel = .Xmin
                  End If
                  If .Xmax > XmaxSel Then
                     XmaxSel = .Xmax
                  End If
                  If .Ymin < YminSel Then
                     YminSel = .Ymin
                  End If
                  If .Ymax > YmaxSel Then
                     YmaxSel = .Ymax
                  End If
               End If
            End With
         Next i
         LargeurSelection = Abs(XmaxSel - XminSel)
         HauteurSelection = Abs(YmaxSel - YminSel)
         lblDimensionSelection.Caption = Label(5) & Format(LargeurSelection, "##0.0") & " mm x " _
                                   & Format(HauteurSelection, "##0.0") & " mm" 'sélection
         lblDimensionSelection.Visible = True
         'si le rectangle de sélection est plat, on l'élargit (mais après avoir indiqué la taille de la sélection)
         If (Abs(XmaxSel - XminSel) < 10) Then
            XminSel = (XminSel + XmaxSel) / 2 - 5
            XmaxSel = (XminSel + XmaxSel) / 2 + 5
         End If
         If (Abs(YmaxSel - YminSel) < 10) Then
            YminSel = (YminSel + YmaxSel) / 2 - 5
            YmaxSel = (YminSel + YmaxSel) / 2 + 5
         End If
         'on trace le rectangle autour
         pctTransf.Line (XminSel - 1, YminSel - 1)-(XmaxSel + 1, YmaxSel + 1), RGB(112, 0, 105), B 'RECTANGLE de la sélection
         'outil étirement : on ajoute les poignées
         If OutilEnCours = Etirer Then
            pctTransf.DrawStyle = vbSolid
            pctTransf.FillColor = RGB(220, 114, 228)
            pctTransf.FillStyle = 0  'rempli
            pctTransf.DrawMode = vbMaskPen
            DemiCarre = 3 * pctTransf.ScaleWidth / pctTransf.Width 'pour avoir un carré de 6 pixels de côté
            pctTransf.Line (XminSel - DemiCarre, YminSel - DemiCarre)-(XminSel + DemiCarre, YminSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line (XminSel - DemiCarre, (YminSel + YmaxSel) / 2 - DemiCarre)-(XminSel + DemiCarre, (YminSel + YmaxSel) / 2 + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line (XminSel - DemiCarre, YmaxSel - DemiCarre)-(XminSel + DemiCarre, YmaxSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line ((XminSel + XmaxSel) / 2 - DemiCarre, YminSel - DemiCarre)-((XminSel + XmaxSel) / 2 + DemiCarre, YminSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line ((XminSel + XmaxSel) / 2 - DemiCarre, YmaxSel - DemiCarre)-((XminSel + XmaxSel) / 2 + DemiCarre, YmaxSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line (XmaxSel - DemiCarre, YmaxSel - DemiCarre)-(XmaxSel + DemiCarre, YmaxSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line (XmaxSel - DemiCarre, (YminSel + YmaxSel) / 2 - DemiCarre)-(XmaxSel + DemiCarre, (YminSel + YmaxSel) / 2 + DemiCarre), RGB(156, 86, 228), B
            pctTransf.Line (XmaxSel - DemiCarre, YminSel - DemiCarre)-(XmaxSel + DemiCarre, YminSel + DemiCarre), RGB(156, 86, 228), B
            pctTransf.FillStyle = 1  'transparent
            pctTransf.DrawMode = vbCopyPen
         End If
         'outil tourner : on ajoute les axes
         If OutilEnCours = Tourner Then
            Xcentre = (XminSel + XmaxSel) / 2
            Ycentre = (YminSel + YmaxSel) / 2
            DemiLargeurSelection = Abs(XmaxSel - XminSel) / 2
            DemiHauteurSelection = Abs(YmaxSel - YminSel) / 2
            'tracé des axes :
            pctTransf.DrawStyle = vbDashDot
            pctTransf.Line (Xcentre - DemiLargeurSelection, Ycentre)-(Xcentre + DemiLargeurSelection, Ycentre), RGB(112, 0, 105)
            pctTransf.Line (Xcentre, Ycentre - DemiHauteurSelection)-(Xcentre, Ycentre + DemiHauteurSelection), RGB(112, 0, 105)
            pctTransf.DrawStyle = vbSolid
         End If
      ElseIf NbTransfSel = 0 Then
         lblDimensionSelection.Visible = False
      End If
   End If
   'Tracé des points des angles si outil mesurer
   If optOutils(Mesurer).Value = True And chkVoirPoints.Value = vbChecked Then
      pctTransf.FillStyle = 1 'transparent
      'points du rectangle de la surface utile
      pctTransf.Circle (0, MargePlateau), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (CourseX, CourseY - MargeFil), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (0, CourseY - MargeFil), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (CourseX, MargePlateau), RayonPoints, RGB(144, 0, 82)
      'rectangle du bloc
      pctTransf.Circle (MargeFil, 0), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (LongBloc + MargeFil, HautBloc), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (LongBloc + MargeFil, 0), RayonPoints, RGB(144, 0, 82)
      pctTransf.Circle (MargeFil, HautBloc), RayonPoints, RGB(144, 0, 82)
   End If
   pctTransf.AutoRedraw = False
End Sub

'******** Déselection de toutes les séquences transférées *************
Private Sub DeselecTransf()
   Dim i As Long
   For i = 1 To NbTransf
      If (Transf(i).Etat And 1) = 1 Then Transf(i).Etat = Transf(i).Etat - 1 'on les désélectionne toutes
   Next i
   NbTransfSel = 0
   Call InitialiserOutils
End Sub

'*******************************************************************
'********* SELECTION à LA SOURIS DANS séquences INITIALES **********
'*******************************************************************

'***** Survol des séquences initiales : surbrillance de la sélection *****
Private Sub pctSequ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Long, j As Long
   Dim EpsilonX As Single, EpsilonY As Single

   If x <> MemoMouseDownSequX Then lblAvertissementSequ.Visible = False   'le MouseMove tourne en permanence, mais on veut seulement effacer le cadre si on a bougé
   
   If Button = 0 Then   'on est pas en drag and drop
      pctSequ.MouseIcon = ImageList2.ListImages(11).Picture  'curseur normal
      pctSequ.MousePointer = 99
   End If
   pctTransf.Cls  'on efface le plan temporaire du bas (au cas où on en vient)
   NbSequSel = 0
   For i = 1 To NbSequ
      With SequTrace(i)
         If .Xmin = .Xmax Then      'segment vertical
            EpsilonX = 3
         Else
            EpsilonX = 0
         End If
         If .Ymin = .Ymax Then      'segment horizontal
            EpsilonY = 3
         Else
            EpsilonY = 0
         End If
         If y + EpsilonY >= .Ymin And y - EpsilonY <= .Ymax Then  'on sépare la contrôle suivant X et Y pour aller plus vite
            If x + EpsilonX >= .Xmin And x - EpsilonX <= .Xmax Then
               NbSequSel = NbSequSel + 1     'nbsequsel est initialisé à zéro dans le form load
               If i <> NumSequSel Then
                  NumSequSel = i
                  pctSequ.DrawWidth = 2
                  pctSequ.Cls
                  If Sequ(i).NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
                     With SequTrace(i)
                        pctSequ.Line (.Point(1).x - 4, .Point(1).y - 4)-(.Point(1).x + 4, .Point(1).y + 4), vbBlue
                        pctSequ.Line (.Point(1).x - 4, .Point(1).y + 4)-(.Point(1).x + 4, .Point(1).y - 4), vbBlue
                     End With
                  Else           'gestion des séquences consituées de plusieurs points
                     For j = 1 To Sequ(i).NbPoints - 1
                        pctSequ.Line (.Point(j).x, .Point(j).y)-(.Point(j + 1).x, .Point(j + 1).y), vbBlue
                     Next j
                  End If
                  pctSequ.DrawWidth = 1
               End If
            End If
         End If
      End With
   Next i
   If NbSequSel = 0 Then
      pctSequ.Cls
      NumSequSel = 0
   End If
End Sub

'********************************************************************
'********* Transfert par Drag and Drop ******************************
'********************************************************************

'*** A l'enfoncement du bouton on mémorise le point de départ *******
Private Sub pctSequ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Long
   
   MemoMouseDownSequX = x
   Select Case Button
   Case vbLeftButton
      If NbSequ = 0 Then
         lblAvertissementSequ.Top = y + 20
         lblAvertissementSequ.Left = x + 3
         lblAvertissementSequ.Caption = Label(6)  '  " Pas de contour visible. Sélectionnez un fichier dans la bibliothèque. "
         lblAvertissementSequ.Visible = True
      ElseIf NbSequSel = 0 Then
         lblAvertissementSequ.Top = y + 36
         lblAvertissementSequ.Left = x + 3
         lblAvertissementSequ.Caption = Label(7) '  " Pas de contour sous le pointeur. Double-cliquez " & vbCrLf & "  sur un contour ou faites le glisser dans le bloc."
         lblAvertissementSequ.Visible = True
      Else
         If NbSequSel = 1 Then
            TransfTemp = Sequ(NumSequSel)
            TransfTemp.Etat = 0  'initialisation de l'état de sélection dans Transf
            With TransfTemp
               For i = 1 To .NbPoints
                  .Point(i).x = .Point(i).x * CoeffBloc  'mise à l'échelle du bloc : 90% du bloc pour la plus grande séquence
                  .Point(i).y = .Point(i).y * CoeffBloc
               Next i
            End With
            Call MaxiMiniSequ(TransfTemp)
            MemoX = x
            MemoY = y
            pctSequ.Drag
         End If
      End If
   End Select
End Sub

'*** si on le repose en haut, il ne se passe rien ***
Private Sub pctSequ_DragDrop(Source As Control, x As Single, y As Single)
   Dim i As Long
   
   If Source.Name = "pctSequ" Then  'pour éviter les intéractions avec pctTransf
      VecteurX = x - MemoX
      VecteurY = y - MemoY
      pctSequ.AutoRedraw = False
      pctSequ.Cls
      pctAjoutPoint.AutoRedraw = False
      pctAjoutPoint.Cls
      If NbSequSel > 0 Then
         With Sequ(NumSequSel)
            For i = 1 To .NbPoints
               .Point(i).x = ((SequTrace(NumSequSel).Point(i).x + VecteurX) - pctSequ.Width / 2) / CoeffSequ + XcentreSequ
               .Point(i).y = ((SequTrace(NumSequSel).Point(i).y + VecteurY) - pctSequ.Height / 2) / CoeffSequ + YcentreSequ
            Next i
         End With
         Call MaxiMiniSequ(Sequ(NumSequSel))
         Call CalculAffichageInitial
         Call TraceSequ
      End If
      NumSequSel = 0   'pour réinitialisation de la procédure Mouse_move
      NbSequSel = 0
   End If
   pctAjoutPoint.AutoRedraw = False
   pctAjoutPoint.Cls
End Sub

'***** On dessine la séquence sur les plans temporaires des contrôles survolés *****
'**** Survol de la PictureBox de départ ****
Private Sub pctSequ_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   Dim i As Long
   
   If Source.Name = "pctSequ" Then  'pour éviter l'intéraction avec pctTransf
      pctSequ.AutoRedraw = False
      pctSequ.Cls
      pctTransf.AutoRedraw = False
      pctTransf.Cls
      With SequTrace(NumSequSel)
         VecteurX = x - MemoX
         VecteurY = y - MemoY
         If Sequ(NumSequSel).NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
            pctSequ.Line (.Point(1).x - 4 + VecteurX, .Point(1).y - 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y + 4 + VecteurY), vbBlue
            pctSequ.Line (.Point(1).x - 4 + VecteurX, .Point(1).y + 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y - 4 + VecteurY), vbBlue
         Else           'gestion des séquences consituées de plusieurs points
            For i = 1 To Sequ(NumSequSel).NbPoints - 1
               pctSequ.Line (.Point(i).x + VecteurX, .Point(i).y + VecteurY)-(.Point(i + 1).x + VecteurX, .Point(i + 1).y + VecteurY), vbBlue
            Next i
         End If
      End With
   End If
      pctAjoutPoint.AutoRedraw = False
      pctAjoutPoint.Cls
End Sub

'**** Survol de la PictureBox de destination ****
Private Sub pctTransf_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   Dim i As Long, j As Long
   Dim MemoFinTempX As Single, MemoFinTempY As Single
   Dim XS1 As Single, YS1 As Single, XS2 As Single, YS2 As Single
   Dim VecteurX1 As Single, VecteurY1 As Single, VecteurX2 As Single, VecteurY2 As Single
   Dim XA As Single, YA As Single
   Dim XB As Single, YB As Single
   Dim XC As Single, YC As Single
   Dim XD As Single, YD As Single
   Dim AC As Single, AD As Single, BC As Single, BD As Single
   Dim CD As Single, AB As Single
   'C et D sont les extrémités gauche et droite de la séquence transférée
   ' donc C est le dernier point et D est le point n°1
   'A et B sont les extrémités gauche et droite de l'ensemble déjà Transfféré
   ' donc A est le dernier point de la dernière séquence et B est le premier point de la première séquence
   '   ____            __    _____
   '  /    \       A__/  \__/     \__B
   ' C      D
   
   If Source.Name = "pctSequ" Or Source.Name = "pctAjoutPoint" Then   'il s'agit de l'ajout d'une séquence
      pctSequ.Cls
      pctAjoutPoint.Cls
      pctTransf.AutoRedraw = False
      pctTransf.Cls
      With TransfTemp
         Call MaxiMiniSequ(TransfTemp)   'calcul des maxi et mini
         VecteurX = x - (.Xmax + .Xmin) / 2
         VecteurY = y - (.Ymax + .Ymin) / 2
         If .NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
            pctTransf.Line (.Point(1).x - 4 + VecteurX, .Point(1).y - 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y + 4 + VecteurY), Color
            pctTransf.Line (.Point(1).x - 4 + VecteurX, .Point(1).y + 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y - 4 + VecteurY), Color
         Else           'gestion des séquences consituées de plusieurs points
            For i = 1 To .NbPoints - 1
               pctTransf.Line (.Point(i).x + VecteurX, .Point(i).y + VecteurY)-(.Point(i + 1).x + VecteurX, .Point(i + 1).y + VecteurY), vbBlue
            Next i
         End If
         If NbTransf > 0 Then  's'il y a déjà des séquences, on représente la jonction
            XA = Transf(NbTransf).Point(Transf(NbTransf).NbPoints).x
            YA = Transf(NbTransf).Point(Transf(NbTransf).NbPoints).y
            XB = Transf(1).Point(1).x
            YB = Transf(1).Point(1).y
            XC = .Point(TransfTemp.NbPoints).x + VecteurX
            YC = .Point(TransfTemp.NbPoints).y + VecteurY
            XD = .Point(1).x + VecteurX
            YD = .Point(1).y + VecteurY
            AC = Sqr((XC - XA) ^ 2 + (YC - YA) ^ 2)
            AD = Sqr((XD - XA) ^ 2 + (YD - YA) ^ 2)
            BC = Sqr((XC - XB) ^ 2 + (YC - YB) ^ 2)
            BD = Sqr((XD - XB) ^ 2 + (YD - YB) ^ 2)
            CD = Sqr((XD - XC) ^ 2 + (YD - YC) ^ 2) 'pour vérifier si profil fermé
            AB = Sqr((XB - XA) ^ 2 + (YB - YA) ^ 2) 'pour vérifier si profil fermé

            If CD > 0.0001 And AB > 0.0001 Then    'les deux sont ouverts
               If AC < AD And AC < BC And AC < BD Then
                  Jonction = "AC"
                  pctTransf.Line (XA, YA)-(XC, YC), vbMagenta
               ElseIf AD < AC And AD < BC And AD < BD Then
                  Jonction = "AD"
                  pctTransf.Line (XA, YA)-(XD, YD), vbMagenta
               ElseIf BC < AC And BC < AD And BC < BD Then
                  Jonction = "BC"
                  pctTransf.Line (XB, YB)-(XC, YC), vbMagenta
               ElseIf BD < AC And BD < AD And BD < BC Then
                  Jonction = "BD"
                  pctTransf.Line (XB, YB)-(XD, YD), vbMagenta
               End If
            ElseIf AB > 0.0001 And CD <= 0.0001 Then  'le profil qui bouge est fermé, le projet en place est ouvert
               If AD < BD Then
                  Jonction = "AD"
                  pctTransf.Line (XA, YA)-(XD, YD), vbMagenta
               Else
                  Jonction = "BC"
                  pctTransf.Line (XB, YB)-(XC, YC), vbMagenta
               End If
            ElseIf CD > 0.0001 And AB <= 0.0001 Then  'le profil qui bouge est ouvert, le projet en place est fermé
               If AD < AC Then
                  Jonction = "AD"
                  pctTransf.Line (XA, YA)-(XD, YD), vbMagenta
               Else
                  Jonction = "BC"
                  pctTransf.Line (XB, YB)-(XC, YC), vbMagenta
               End If
            ElseIf AB <= 0.0001 And CD <= 0.0001 Then 'les deux sont fermés
                  Jonction = "BC"
                  pctTransf.Line (XB, YB)-(XC, YC), vbMagenta
            End If
         End If
      End With
   ElseIf Source.Name = "pctTransf" Then     'il s'agit de la mise en oeuvre d'un outil
      '********** Outil de DEPLACEMENT **********
      'lors du DragOver, on ne déplace pas réellement la séquence, on le fait seulement au DragDrop
      If OutilEnCours = Deplacer Then
         flagMemoUndoDansTraceTransf = False
         Call TraceTransf(True)  'on retrace le fond sans les séquences sélectionnées
         pctTransf.Cls           'on efface le plan provisoire
         
         PositionX = PositionX - VecteurX    'il faut enlever le vecteur de la précédente itération (X,Y) pour ajouter le nouveau
         PositionY = PositionY - VecteurY
         
         VecteurX = x - MemoX
         VecteurY = y - MemoY
         For j = 1 To NbTransf
            If (Transf(j).Etat And 1) = 1 Then
               '***** Séquence FANTOME ***************
               '**** segments de liaison ****
               If NbTransf > 1 And j > 1 Then
                  If (Transf(j - 1).Etat And 1) = 0 Then    'traits avant la séquences
                     VecteurX1 = 0
                     VecteurY1 = 0
                  Else
                     VecteurX1 = VecteurX
                     VecteurY1 = VecteurY
                  End If
                  VecteurX2 = VecteurX
                  VecteurY2 = VecteurY
                  XS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).x + VecteurX1
                  YS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).y + VecteurY1
                  XS2 = Transf(j).Point(1).x + VecteurX2
                  YS2 = Transf(j).Point(1).y + VecteurY2
                  pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
               End If
               If NbTransf > 1 And j < NbTransf Then     'trait après la séquence
                  If (Transf(j + 1).Etat And 1) = 0 Then
                     XS1 = Transf(j).Point(Transf(j).NbPoints).x + VecteurX
                     YS1 = Transf(j).Point(Transf(j).NbPoints).y + VecteurY
                     XS2 = Transf(j + 1).Point(1).x
                     YS2 = Transf(j + 1).Point(1).y
                  pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
                  End If
               End If
               '***** Séquence ***************
               With Transf(j)
                  If .NbPoints > 1 Then
                     For i = 1 To .NbPoints - 1
                        pctTransf.Line (.Point(i).x + VecteurX, .Point(i).y + VecteurY)-(.Point(i + 1).x + VecteurX, .Point(i + 1).y + VecteurY), vbBlue
                     Next i
                  ElseIf .NbPoints = 1 Then
                     pctTransf.Line (.Point(1).x - 4 + VecteurX, .Point(1).y - 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y + 4 + VecteurY), vbBlue
                     pctTransf.Line (.Point(1).x - 4 + VecteurX, .Point(1).y + 4 + VecteurY)-(.Point(1).x + 4 + VecteurX, .Point(1).y - 4 + VecteurY), vbBlue
                  End If
               End With
            End If
         Next j
         PositionX = PositionX + VecteurX
         PositionY = PositionY + VecteurY
         txtMesures(0).Text = Format(PositionX, "##0.0")
         txtMesures(1).Text = Format(PositionY, "##0.0")
      '********************************************
      '************** Outil de ROTATION ***********
      '********************************************
      ElseIf OutilEnCours = Tourner Then
         flagMemoUndoDansTraceTransf = False
         Call TraceTransf(True)  'on retrace le fond sans les séquences sélectionnées
         pctTransf.Cls           'on efface le plan provisoire
         Rotation = Rotation - AngleRelatif
         If x - Xcentre > 0 Then
            If y - Ycentre > 0 Then
               alpha = Atn((y - Ycentre) / (x - Xcentre)) - AngleTotal
               AngleTotal = Atn((y - Ycentre) / (x - Xcentre))
            Else
               alpha = 2 * pi + Atn((y - Ycentre) / (x - Xcentre)) - AngleTotal
               AngleTotal = 2 * pi + Atn((y - Ycentre) / (x - Xcentre))
               
            End If
         ElseIf x - Xcentre < 0 Then
            alpha = pi + Atn((y - Ycentre) / (x - Xcentre)) - AngleTotal
            AngleTotal = pi + Atn((y - Ycentre) / (x - Xcentre))
         Else
            AngleTotal = 0
         End If
         pctTransf.Cls
         For j = 1 To NbTransf
            If (Transf(j).Etat And 1) = 1 Then
               '**** segments de liaison ****
               If NbTransf > 1 And j > 1 Then
                  If (Transf(j - 1).Etat And 1) = 0 Then    'traits avant la séquences
                     XS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).x
                     YS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).y
                  Else
                     With Transf(j - 1)
                        Xtemp = .Point(Transf(j - 1).NbPoints).x
                        Ytemp = .Point(Transf(j - 1).NbPoints).y
                        XS1 = Xcentre + (Xtemp - Xcentre) * Cos(alpha) - (Ytemp - Ycentre) * Sin(alpha)
                        YS1 = Ycentre + (Xtemp - Xcentre) * Sin(alpha) + (Ytemp - Ycentre) * Cos(alpha)
                     End With
                  End If
                  Xtemp = Transf(j).Point(1).x
                  Ytemp = Transf(j).Point(1).y
                  XS2 = Xcentre + (Xtemp - Xcentre) * Cos(alpha) - (Ytemp - Ycentre) * Sin(alpha)
                  YS2 = Ycentre + (Xtemp - Xcentre) * Sin(alpha) + (Ytemp - Ycentre) * Cos(alpha)
                  pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
               End If
               If NbTransf > 1 And j < NbTransf Then     'trait après la séquence
                  If (Transf(j + 1).Etat And 1) = 0 Then
                     With Transf(j)
                        Xtemp = .Point(Transf(j).NbPoints).x
                        Ytemp = .Point(Transf(j).NbPoints).y
                        XS1 = Xcentre + (Xtemp - Xcentre) * Cos(alpha) - (Ytemp - Ycentre) * Sin(alpha)
                        YS1 = Ycentre + (Xtemp - Xcentre) * Sin(alpha) + (Ytemp - Ycentre) * Cos(alpha)
                     End With
                     XS2 = Transf(j + 1).Point(1).x
                     YS2 = Transf(j + 1).Point(1).y
                  pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
                  End If
               End If
               '******** Séquence fantôme ***********
               With Transf(j)
                  For i = 1 To .NbPoints
                     Xtemp = .Point(i).x
                     Ytemp = .Point(i).y
                     .Point(i).x = Xcentre + (Xtemp - Xcentre) * Cos(alpha) - (Ytemp - Ycentre) * Sin(alpha)
                     .Point(i).y = Ycentre + (Xtemp - Xcentre) * Sin(alpha) + (Ytemp - Ycentre) * Cos(alpha)
                  Next i
                  'tracé du fantome :
                  If .NbPoints > 1 Then
                     For i = 1 To .NbPoints - 1
                        pctTransf.Line (.Point(i).x, .Point(i).y)-(.Point(i + 1).x, .Point(i + 1).y), vbBlue
                     Next i
                  ElseIf .NbPoints = 1 Then
                     pctTransf.Line (.Point(1).x - 4, .Point(1).y - 4)-(.Point(1).x + 4, .Point(1).y + 4), vbBlue
                     pctTransf.Line (.Point(1).x - 4, .Point(1).y + 4)-(.Point(1).x + 4, .Point(1).y - 4), vbBlue
                  End If
               End With
            End If
         Next j
         'tracé des lignes de référence :
         pctTransf.DrawStyle = vbDashDot
         pctTransf.Line (Xcentre - DemiLargeurSelection, Ycentre)-(Xcentre + DemiLargeurSelection, Ycentre), vbBlue
         pctTransf.Line (Xcentre, Ycentre - DemiHauteurSelection)-(Xcentre, Ycentre + DemiHauteurSelection), vbBlue
         pctTransf.DrawStyle = vbSolid
         pctTransf.Line (Xcentre, Ycentre)-(x, y), vbBlue
         
         AngleRelatif = AngleTotal - AngleInitial
         Rotation = Rotation + AngleRelatif
         txtMesures(0).Text = Format((Rotation * 180 / pi + 360) Mod 360, "##0.0")

      '****************************
      '******* Outil ETIRER *******
      '****************************
      ElseIf OutilEnCours = Etirer Then
         AgrandissementX = AgrandissementX / kX
         AgrandissementY = AgrandissementY / kY
         Select Case Poignee     'avec les poignées des angles, la touche Shift rend le truc proportionnel
         Case "HG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XmaxSel
               OrigineY = YminSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = -1
            Else
               If (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) > Abs(y - ValeurInitialeY) Then 'touche shift appuyée
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = kX
               ElseIf (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) <= Abs(y - ValeurInitialeY) Then
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
                  kX = kY
               Else
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
               End If
            End If
            
         Case "HM"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = 0
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = 0
               OrigineY = YminSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = 1
               kY = -1
            Else
               kX = 1
               kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
            End If
         Case "HD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XminSel
               OrigineY = YminSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = -1
            Else
               If (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) > Abs(y - ValeurInitialeY) Then
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = kX
               ElseIf (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) <= Abs(y - ValeurInitialeY) Then
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
                  kX = kY
               Else
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
               End If
            End If
         Case "MD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = 0
            Else
               OrigineX = XminSel
               OrigineY = 0
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = 1
            Else
               kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
               kY = 1
            End If
         Case "BD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XminSel
               OrigineY = YmaxSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = -1
            Else
               If (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) > Abs(y - ValeurInitialeY) Then
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = kX
               ElseIf (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) <= Abs(y - ValeurInitialeY) Then
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
                  kX = kY
               Else
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
               End If
            End If
         Case "BM"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = 0
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = 0
               OrigineY = YmaxSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = 1
               kY = -1
            Else
               kX = 1
               kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
            End If
         Case "BG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XmaxSel
               OrigineY = YmaxSel
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = -1
            Else
               If (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) > Abs(y - ValeurInitialeY) Then
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = kX
               ElseIf (GetAsyncKeyState(16) <> 0) And Abs(x - ValeurInitialeX) <= Abs(y - ValeurInitialeY) Then
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
                  kX = kY
               Else
                  kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
                  kY = (y - OrigineY) / (ValeurInitialeY - OrigineY)
               End If
            End If
         Case "MG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = 0
            Else
               OrigineX = XmaxSel
               OrigineY = 0
            End If
            If GetAsyncKeyState(18) <> 0 Then   'si on appuye pas sur Alt
               kX = -1
               kY = 1
            Else
               kX = (x - OrigineX) / (ValeurInitialeX - OrigineX)
               kY = 1
            End If
         End Select
         AgrandissementX = AgrandissementX * kX
         AgrandissementY = AgrandissementY * kY
         txtMesures(0).Text = Format(AgrandissementX * 100, "##0.0")
         txtMesures(1).Text = Format(AgrandissementY * 100, "##0.0")
         lblDimensionSelection.Caption = Label(5) & Format(LargeurSelection * kX, "##0.0") & " mm x " _
                                   & Format(HauteurSelection * kY, "##0.0") & " mm" 'sélection
         pctTransf.Cls
         For j = 1 To NbTransf
            If (Transf(j).Etat And 1) = 1 Then
               ReDim TransfTemp.Point(1 To Transf(j).NbPoints)
               With Transf(j)
                  For i = 1 To .NbPoints
                     TransfTemp.Point(i).x = kX * (.Point(i).x - OrigineX) + OrigineX
                     TransfTemp.Point(i).y = kY * (.Point(i).y - OrigineY) + OrigineY
                  Next i
                  TransfTemp.NbPoints = Transf(j).NbPoints
                  MemoFinTempX = TransfTemp.Point(TransfTemp.NbPoints).x
                  MemoFinTempY = TransfTemp.Point(TransfTemp.NbPoints).y
               End With
               '**** segments de liaison avant et après les séquences sélectionnées ****
               If NbTransf > 1 Then
                  If j > 1 Then
                     If (Transf(j - 1).Etat And 1) = 0 Then    'traits avant la séquences
                        XS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).x
                        YS1 = Transf(j - 1).Point(Transf(j - 1).NbPoints).y
                        XS2 = TransfTemp.Point(1).x
                        YS2 = TransfTemp.Point(1).y
                        pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
                     Else
                        XS1 = MemoFinTempX
                        YS1 = MemoFinTempY
                        XS2 = TransfTemp.Point(1).x
                        YS2 = TransfTemp.Point(1).y
                        pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
                     End If
                  End If
                  If j < NbTransf Then     'trait après la séquence
                     If (Transf(j + 1).Etat And 1) = 0 Then
                        XS1 = TransfTemp.Point(Transf(j).NbPoints).x
                        YS1 = TransfTemp.Point(Transf(j).NbPoints).y
                        XS2 = Transf(j + 1).Point(1).x
                        YS2 = Transf(j + 1).Point(1).y
                        pctTransf.Line (XS1, YS1)-(XS2, YS2), vbBlue
                     End If
                  End If
               End If
               With TransfTemp
                  'tracé du fantome :
                  If .NbPoints > 1 Then
                     For i = 1 To .NbPoints - 1
                        pctTransf.Line (.Point(i).x, .Point(i).y)-(.Point(i + 1).x, .Point(i + 1).y), vbBlue
                     Next i
                  ElseIf .NbPoints = 1 Then
                     pctTransf.Line (.Point(1).x - 4, .Point(1).y - 4)-(.Point(1).x + 4, .Point(1).y + 4), vbBlue
                     pctTransf.Line (.Point(1).x - 4, .Point(1).y + 4)-(.Point(1).x + 4, .Point(1).y - 4), vbBlue
                  End If
               End With
            End If
         Next j
      End If
   End If
End Sub

'***************** Drop de séquences dans la fenêtre du bas *******************
Private Sub pctTransf_DragDrop(Source As Control, x As Single, y As Single)
   Dim i As Long, j As Long
   Dim VecteurReelX As Single, VecteurReelY As Single
   
   If Source.Name = "pctSequ" Or Source.Name = "pctAjoutPoint" Then 'Insertion d'une nouvelle séquence par glisser-déposer
      If NbTransf = 0 Then
         NbTransf = 1
         Call MaxiMiniSequ(TransfTemp)
         With TransfTemp
            VecteurReelX = x
            VecteurReelY = y
            For i = 1 To .NbPoints
               .Point(i).x = .Point(i).x - (.Xmax + .Xmin) / 2 + VecteurReelX
               .Point(i).y = .Point(i).y - (.Ymax + .Ymin) / 2 + VecteurReelY
            Next i
         End With
         ReDim Transf(1 To NbTransf)
         Transf(1) = TransfTemp
         Call MaxiMiniSequ(Transf(1))
      ElseIf NbTransf > 0 Then   's'il y a déjà des séquences, il faut regarder le trait de liaison
         Call MaxiMiniSequ(TransfTemp)
         With TransfTemp
            VecteurReelX = x
            VecteurReelY = y
            For i = 1 To .NbPoints
               .Point(i).x = .Point(i).x - (.Xmax + .Xmin) / 2 + VecteurReelX
               .Point(i).y = .Point(i).y - (.Ymax + .Ymin) / 2 + VecteurReelY
            Next i
         End With
         Call MaxiMiniSequ(TransfTemp)
         NbTransf = NbTransf + 1
         ReDim Preserve Transf(1 To NbTransf)
         If Jonction = "AC" Then    'la séquence Transfférée passe en dernière posisiton et son sens s'inverse
            Transf(NbTransf) = TransfTemp
            Call InverserSensSequ(Transf(NbTransf))
         ElseIf Jonction = "AD" Then    'la séquence Transfférée passe en dernière position sans inverser son sens
            Transf(NbTransf) = TransfTemp
         ElseIf Jonction = "BC" Then    'la séquence Transfférée passe en première position sans inverser son sens
            For j = NbTransf To 2 Step -1
               Transf(j) = Transf(j - 1)
            Next j
            Transf(1) = TransfTemp
         ElseIf Jonction = "BD" Then    'la séquence Transférée passe en première position et son sens s'inverse
            For j = NbTransf To 2 Step -1
               Transf(j) = Transf(j - 1)
            Next j
            Transf(1) = TransfTemp
            Call InverserSensSequ(Transf(1))
         End If
      End If
      If NbTransf > 0 Then
         Call AppliquerCentrageEtAjustement
         ' Call ZoomAutoToutVoir  'ajustement automatique
         Call TraceTransf
      End If
   ElseIf Source.Name = "pctTransf" Then     'relâchement lors de l'utilisation d'un outil
      If OutilEnCours = Deplacer Then '**** déplacement des séquences sélectionnées ****
         
         '**** à ce stade, il n'y a pas eu de déplacement, on a seulement tracé des traits, maintenant on affecte la
         '**** nouvelle position à la sélection
         For j = 1 To NbTransf
            If (Transf(j).Etat And 1) = 1 Then
               With Transf(j)
                  For i = 1 To .NbPoints
                     .Point(i).x = .Point(i).x + VecteurX
                     .Point(i).y = .Point(i).y + VecteurY
                  Next i
               End With
               Call MaxiMiniSequ(Transf(j))
            End If
         Next j
      ElseIf OutilEnCours = Tourner Then '**** rotation des séquences sélectionnées ****
         'le fantôme est tracé avec TraceTransf, donc il suffit de retracer
         For j = 1 To NbTransf
            Call MaxiMiniSequ(Transf(j))
         Next j
      ElseIf OutilEnCours = Etirer Then  '******* OUTIL ETIREMENT ********
         
         Select Case Poignee     'avec les poignées des angles, la touche Shift rend le truc proportionnel
         Case "HG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XmaxSel
               OrigineY = YminSel
            End If
         Case "HM"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = 0
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = 0
               OrigineY = YminSel
            End If
         Case "HD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XminSel
               OrigineY = YminSel
            End If
         Case "MD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = 0
            Else
               OrigineX = XminSel
               OrigineY = 0
            End If
         Case "BD"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XminSel
               OrigineY = YmaxSel
            End If
         Case "BM"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = 0
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = 0
               OrigineY = YmaxSel
            End If
         Case "BG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = (YmaxSel + YminSel) / 2
            Else
               OrigineX = XmaxSel
               OrigineY = YmaxSel
            End If
         Case "MG"
            If GetAsyncKeyState(17) <> 0 Then   'touche Ctrl appuyée : homothétie autour du milieu
               OrigineX = (XmaxSel + XminSel) / 2
               OrigineY = 0
            Else
               OrigineX = XmaxSel
               OrigineY = 0
            End If
         End Select
         For j = 1 To NbTransf
            If (Transf(j).Etat And 1) = 1 Then
               With Transf(j)
                  For i = 1 To .NbPoints
                     .Point(i).x = kX * (.Point(i).x - OrigineX) + OrigineX
                     .Point(i).y = kY * (.Point(i).y - OrigineY) + OrigineY
                  Next i
               End With
               Call MaxiMiniSequ(Transf(j))
            End If
         Next j
      End If
      Call AppliquerCentrageEtAjustement
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub

'***********************************************************
'******* GESTION des EVENEMENTS CLAVIER ********************
'***********************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   '****************************************************************
   '***** CTRL : changement du curseur de l'outil Selectionner *****
   '****************************************************************
   Select Case KeyCode
   Case vbKeyControl 'si on appuye sur Ctrl avec l'outil sélectionner
      Select Case OutilEnCours
      Case Deplacer, Tourner, Etirer
         pctTransf.MouseIcon = ImageList2.ListImages(6).Picture
         pctTransf.MousePointer = 99
      End Select
   End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim i As Long, j As Long
   
   Select Case KeyCode
   '************************************************************************************
   '****************** Relèvement de la touche CTRL : changement du curseur de sélection
   '************************************************************************************
   Case vbKeyControl
      Select Case OutilEnCours
      Case Deplacer
         pctTransf.Cls  'pour faire disparaître le cadre
         pctTransf.MouseIcon = ImageList2.ListImages(2).Picture
      Case Tourner
         pctTransf.Cls  'pour faire disparaître le cadre
         pctTransf.MouseIcon = ImageList2.ListImages(4).Picture
      Case Etirer
         pctTransf.Cls  'pour faire disparaître le cadre
         pctTransf.MouseIcon = ImageList2.ListImages(5).Picture
      End Select
   '**********************************************************
   '****** ESCAPE : DESELECTION DE TOUTES LES SEQUENCES ******
   '**********************************************************
   Case vbKeyEscape
      If OutilEnCours = Mesurer And flagPremierPoint = True Then 'on veut annuler la saisie du deuxième point de la mesure
         flagPremierPoint = False
         Call TraceTransf
      Else
         If NbTransf > 0 Then
            For i = 1 To NbTransf
               If (Transf(i).Etat And 1) = 1 Then
                  Transf(i).Etat = Transf(i).Etat - 1
               End If
            Next i
            NbTransfSel = 0
            Call TraceTransf
         End If
      End If
   '**********************************************
   '****** SUPPR : suppression de séquences ******
   '**********************************************
   Case vbKeyDelete
      If NbTransf > 0 Then
         j = 0 'vérification du nombre de séquences sélectionnées
         For i = 1 To NbTransf
            If (Transf(i).Etat And 1) = 1 Then j = j + 1
         Next i
         'si elles le sont toutes, il faut réinitialiser
         If j = NbTransf Then
            NbTransf = 0
            NbTransfSel = 0
            Erase Transf
            Call EchelleTransf(0, CourseX, 0, CourseY)
            Call TraceTransf
            Call DesactiverMesures  'il n'y a plus de séquences, on désactive les textbox de saisie manuelle
            Exit Sub
         End If
         'sinon, on compacte le tableau
         i = 1
         Do
            If (Transf(i).Etat And 1) = 1 Then
               NbTransf = NbTransf - 1
               For j = i To NbTransf
                  Transf(j) = Transf(j + 1)
               Next j
               ReDim Preserve Transf(1 To NbTransf)
               If i = NbTransf + 1 Then Exit Do
               i = i - 1
            End If
            If i = NbTransf Then Exit Do
            i = i + 1
         Loop
         NbTransfSel = 0   'puisqu'elles ont été supprimées
         ' Call ZoomAutoToutVoir
         Call TraceTransf
      End If
   End Select
End Sub

'********* Effacement des plans temporaires des picturebox quand on en sort (on est sur Form) ***************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   pctSequ.Cls
   pctTransf.Cls
   pctAjoutPoint.AutoRedraw = False
   pctAjoutPoint.Cls
End Sub

'**************************************
'**** modification HAUTEUR du BLOC ****
'*************************µ************
Private Sub txtBloc_GotFocus(Index As Integer)  'ancien txtBloc
   txtBloc(Index).SelStart = 0
   txtBloc(Index).SelLength = Len(txtBloc(Index).Text)
   txtMesures(0).TabStop = False
   txtMesures(1).TabStop = False
   txtBloc(0).TabStop = True
   txtBloc(1).TabStop = True
End Sub
      
'si on quitte le textbox sans valider par Entrée, c'est qu'il n'y a pas eu de modification, on affiche la valeur en mémoire
Private Sub txtBloc_LostFocus(Index As Integer)
   If GetTabState Then  'la perte de focus s'est produite par appui de la touche Tab
      Call txtBloc_KeyUp(Index, 13, 0)  'on simule l'appui sur Entrée
      If Index = 0 Then
         txtBloc(1).SetFocus
      Else
         txtBloc(0).SetFocus
      End If
   Else
      lblBlocMaxi(Index).BackColor = vbButtonFace
      lblBlocMaxi(Index).ForeColor = vbButtonText
      If Index = 0 Then
         txtBloc(0).Text = V2P(Format(HautBloc, "#####0"))
      Else
         txtBloc(1).Text = V2P(Format(LongBloc, "#####0"))
      End If
   End If
End Sub

Private Sub txtBloc_KeyPress(Index As Integer, KeyAscii As Integer)
   'à l 'appui sur une touche, on autorise seulement les nombres positifs
   Call VerifierEntiers(KeyAscii)  'Keyascii est passé par référence
End Sub

'la validation se fait par appui sur Entrée
Private Sub txtBloc_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   
   Select Case Index
   Case 0
      If KeyCode = 13 Then
         If Not MyIsNumeric(txtBloc(0).Text) Then       'si erreur de saisie
            lblBlocMaxi(0).BackColor = vbRed
            lblBlocMaxi(0).ForeColor = vbWhite
            txtBloc(0).Text = V2P(Format(HautBloc, "#####0"))
            txtBloc(0).SelStart = 0
            txtBloc(0).SelLength = Len(txtBloc(0).Text)
         Else
            If MyVal(txtBloc(0).Text) <= 1 Or MyVal(txtBloc(0).Text) > CourseY - MargeFil Then 'on interdit l'écartement nul pour éviter les problèmes de division (dièdre)
               lblBlocMaxi(0).BackColor = vbRed
               lblBlocMaxi(0).ForeColor = vbWhite
               txtBloc(0).Text = V2P(Format(HautBloc, "#####0"))
               txtBloc(0).SelStart = 0
               txtBloc(0).SelLength = Len(txtBloc(0).Text)
               Exit Sub
            End If
            lblBlocMaxi(0).BackColor = vbButtonFace
            lblBlocMaxi(0).ForeColor = vbButtonText
            HautBloc = MyVal(txtBloc(0).Text)
            txtBloc(0).Text = V2P(Format(HautBloc, "#####0"))
            Call AppliquerCentrageEtAjustement
            If chkZoomProjet.Value = vbChecked Then
               Call EchelleTransf(0, LongBloc, 0, HautBloc)
            ElseIf chkZoomProjet.Value = vbUnchecked Then
               ' Call ZoomAutoToutVoir 'on vérifie si ça dépasse
            End If
            Call TraceTransf
         End If
      End If
   Case 1
      If KeyCode = 13 Then
         If Not MyIsNumeric(txtBloc(1).Text) Then       'si erreur de saisie
            lblBlocMaxi(1).BackColor = vbRed
            lblBlocMaxi(1).ForeColor = vbWhite
            txtBloc(1).Text = V2P(Format(LongBloc, "#####0"))
            txtBloc(1).SelStart = 0
            txtBloc(1).SelLength = Len(txtBloc(1).Text)
         Else
            If MyVal(txtBloc(1).Text) <= 1 Or MyVal(txtBloc(1).Text) > CourseX - MargeFil Then 'on interdit l'écartement nul pour éviter les problèmes de division (dièdre)
               lblBlocMaxi(1).BackColor = vbRed
               lblBlocMaxi(1).ForeColor = vbWhite
               txtBloc(1).Text = V2P(Format(LongBloc, "#####0"))
               txtBloc(1).SelStart = 0
               txtBloc(1).SelLength = Len(txtBloc(1).Text)
               Exit Sub
            End If
            lblBlocMaxi(1).BackColor = vbButtonFace
            lblBlocMaxi(1).ForeColor = vbButtonText
            LongBloc = MyVal(txtBloc(1).Text)
            txtBloc(1).Text = V2P(Format(LongBloc, "#####0"))
            Call AppliquerCentrageEtAjustement
            If chkZoomProjet.Value = vbChecked Then
               Call EchelleTransf(0, LongBloc, 0, HautBloc)
            ElseIf chkZoomProjet.Value = vbUnchecked Then
               ' Call ZoomAutoToutVoir 'on vérifie si ça dépasse
            End If
            Call TraceTransf
         End If
      End If
   End Select
End Sub

'*********************************************************************
'********* SELECTION à LA SOURIS DANS séquences transférées **********
'*********************************************************************
'La sélection fonctionne selon le modèle des dessins dans Word 97
'Au MouseDown, si une séquence est sous le curseur et n'est pas sélectionnée
'elle devient l'unique séquence sélectionnée,
'par contre si elle est sélectionnée, rien ne change
Private Sub pctTransf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim flagDansLaZone As Boolean
   Dim p As POINTAPI
   Dim ret As Long
   
   If Button = vbRightButton Or Button = vbMiddleButton Then 'clic droit : activation du zoom (molette ou flèches haut-bas)
      XZoomMolette = x
      YZoomMolette = y
      With pctTransf
         p.x = 0
         p.y = 0
         'Get information about the form's left and top
         ret = ClientToScreen(.hwnd, p)
         p.x = p.x + Abs(.Width / 2) - 2
         p.y = p.y + Abs(.Height / 2) - 2
         'Set the cursor to the middle of the form
         ret = SetCursorPos&(p.x, p.y)
         .ScaleLeft = XZoomMolette - Abs(.ScaleWidth / 2)
         .ScaleTop = YZoomMolette + Abs(.ScaleHeight / 2)
      End With
      flagMemoUndoDansTraceTransf = False
      Call TraceTransf
      sliderZoom.Value = 0
      sliderZoom.SetFocus
   ElseIf Button = vbLeftButton Then   'clic gauche : outils
      If OutilEnCours <> Mesurer And OutilEnCours <> CouperProfil And OutilEnCours <> PointNumero1 Then
      
         TransfSousCur = SousCursTransf(x, y)  'on teste tout de suite si on a cliqué sur une séquence
         MemoMouseDownX = x  'mémorisation pour effacement du cadre d'aide lors du mouse_move
         lblMesure.Visible = False  'initialisation du texte de cotation
         flagRectSel = False  'initialisation du flag qui permet de savoir dans le mouseup si on est en rectangle de sélection
         
         flagDansLaZone = False
         If x >= XminSel And x <= XmaxSel And y >= YminSel And y <= YmaxSel Then
            flagDansLaZone = True
         End If
      
         'Avertissement si pas de contours dans le bloc
         If NbTransf = 0 Then  'si pas de séquences, avertissement
            lblAvertissementTransf.Top = y + 34
            lblAvertissementTransf.Left = x + 3
            lblAvertissementTransf.Caption = Label(8) '  " Pas de contour transféré dans le bloc. Sélectionnez un fichier de la " & vbCrLf & _
                                                      ' "  bibliothèque puis double-cliquez sur un contour ou faites-le glisser ici."
            lblAvertissementTransf.Visible = True
            Exit Sub
         End If
      
         If NbTransf > 0 Then 'contrôle qu'il y a bien des séquences en bas et qu'on a cliqué gauche
            If OutilEnCours = Etirer And pctTransf.MousePointer <> 99 Then  'lorsque le curseur n'est pas personnalisé, c'est qu'on est sur une poignée (curseur VB)
               kX = 1   'pour éviter la division par zéro
               kY = 1
               Select Case Poignee
                  Case "BG", "MG", "HG", "BM", "HM", "BD", "MD", "HD"
                  ValeurInitialeX = x
                  ValeurInitialeY = y
                  pctTransf.Drag
               End Select
            ElseIf OutilEnCours <> Mesurer And OutilEnCours <> CouperProfil Then 'dans le cas des outils mesurer et découper, on ne gère pas la sélection
               'si la touche ctrl est appuyée OU qu'on est à la première sélection (pas de zone), on bascule juste l'état
               If Shift = vbCtrlMask Or (Shift = 0 And NbTransfSel = 0) Then
                  If TransfSousCur > 0 Then  'si on était sur une séquence dans le mousemove
                     If (Transf(TransfSousCur).Etat And 1) = 0 Then 'si la séquence n'est pas sélectionnée
                        Transf(TransfSousCur).Etat = Transf(TransfSousCur).Etat + 1  'et on sélectionne celle qui est sous le curseur
                        NbTransfSel = NbTransfSel + 1
                        Call InitialiserOutils
                        '***le déplacement est activé pour la séquence qui vient d'être sélectionnée
                        If OutilEnCours = Deplacer Then
                           MemoX = x      'mémorisation pour calcul du vecteur
                           MemoY = y
                           VecteurX = 0     'réinitialisation du vecteur du déplacement précédent (utilisé dans le DragOver)
                           VecteurY = 0
                           pctTransf.Drag
                        End If
                        '***
                     Else  'sinon on la désélectionne
                        Transf(TransfSousCur).Etat = Transf(TransfSousCur).Etat - 1
                        NbTransfSel = NbTransfSel - 1
                        Call InitialiserOutils
                     End If
                     Call TraceTransf
                     Exit Sub
                  ElseIf TransfSousCur = 0 Then    'si rien sous le curseur, on commence une zone de sélection
                     flagRectSel = True
                     X1SelTransf = x   'début du déplacement pour tracé du rectangle de sélection si la souris est bougée avec le bouton maintenu
                     Y1SelTransf = y
                  End If
               End If
               
               'si pas de touche (shift, ctrl...) et existence d'une zone
               If Shift = 0 And NbTransfSel > 0 Then
                  If flagDansLaZone = False Then  'si on n'est pas dans la zone de sélection, on est en rectangle de sélection
                     If TransfSousCur = 0 Then
                        flagRectSel = True
                        Call DeselecTransf      'on les désélectionne toutes
                        X1SelTransf = x   'début du déplacement pour tracé du rectangle de sélection si la souris est bougée avec le bouton maintenu
                        Y1SelTransf = y
                     Else
                        Call DeselecTransf      'on les désélectionne toutes
                        Transf(TransfSousCur).Etat = Transf(TransfSousCur).Etat + 1  'et on sélectionne celle qui est sous le curseur
                        NbTransfSel = NbTransfSel + 1
                        Call InitialiserOutils
                        '***le déplacement est activé pour la séquence qui vient d'être sélectionnée
                        If OutilEnCours = Deplacer Then
                           MemoX = x      'mémorisation pour calcul du vecteur
                           MemoY = y
                           VecteurX = 0     'réinitialisation du vecteur du déplacement précédent (utilisé dans le DragOver)
                           VecteurY = 0
                           pctTransf.Drag
                        End If
                        '***
                     End If
                     Call TraceTransf
                     Exit Sub
                  Else
                     If OutilEnCours = Deplacer Then  'Outil DEPLACER
                        MemoX = x      'mémorisation pour calcul du vecteur
                        MemoY = y
                        VecteurX = 0     'réinitialisation du vecteur du déplacement précédent (utilisé dans le DragOver)
                        VecteurY = 0
                        pctTransf.Drag
                     ElseIf OutilEnCours = Tourner Then 'outil TOURNER : il faut définir le centre de la sélection
                        AngleRelatif = 0
                        Xcentre = (XminSel + XmaxSel) / 2
                        Ycentre = (YminSel + YmaxSel) / 2
                        'initialisation de l'angle pour rotation à la souris autour du centre
                        If x - Xcentre > 0 Then
                           If y - Ycentre > 0 Then
                              AngleTotal = Atn((y - Ycentre) / (x - Xcentre))
                           Else
                              AngleTotal = 2 * pi + Atn((y - Ycentre) / (x - Xcentre))
                           End If
                        ElseIf x - Xcentre < 0 Then
                           AngleTotal = pi + Atn((y - Ycentre) / (x - Xcentre))
                        Else
                           AngleTotal = 0
                        End If
                        AngleInitial = AngleTotal
                        pctTransf.Drag
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub
'************* DEPLACEMENT DE LA SOURIS SUR SEQUENCES TRANSFEREES *************
Private Sub pctTransf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Long, j As Long
   Dim RayonCercles As Single
   Dim NombrePixelsAttraction As Single
   Dim ToleranceAttraction As Single
   
   
   NombrePixelsAttraction = 4
   ToleranceAttraction = NombrePixelsAttraction * UnPixelToMm
   RayonCercles = NombrePixelsAttraction * pctTransf.ScaleWidth / pctTransf.Width
   
   pctTransf.MousePointer = 99 'utilisation exclusive des curseurs personnalisés
   If x <> MemoMouseDownX Then lblAvertissementTransf.Visible = False   'le MouseMove tourne en permanence, mais on veut seulement effacer le cadre si on a bougé
   '*********** outil MESURER **********
   If OutilEnCours = Mesurer Then   ' tracé des points survolé avec le mètre
      flagPointSousCur = False
      pctTransf.FillColor = RGB(220, 114, 228)
      pctTransf.FillStyle = 0  'rempli
      pctTransf.DrawMode = vbMaskPen
      pctTransf.DrawWidth = 2
      If NbTransf > 0 Then
         'test sur les points des séquences
         For i = 1 To NbTransf
            With Transf(i)
               For j = 1 To .NbPoints
                  If Abs(x - .Point(j).x) < ToleranceAttraction Then
                     If Abs(y - .Point(j).y) < ToleranceAttraction Then
                        pctTransf.Cls
                        pctTransf.Circle (.Point(j).x, .Point(j).y), RayonCercles, RGB(156, 86, 228)
                        XPointMesure = .Point(j).x
                        YPointMesure = .Point(j).y
                        flagPointSousCur = True
                     End If
                  End If
               Next j
            End With
         Next i
      End If
      'test sur les angles du bloc
      If flagPointSousCur = False Then
         If Abs(x - MargeFil) < ToleranceAttraction Then
            If Abs(y - 0) < ToleranceAttraction Then
               XPointMesure = MargeFil
               YPointMesure = 0
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            ElseIf Abs(y - HautBloc) < ToleranceAttraction Then
               XPointMesure = MargeFil
               YPointMesure = HautBloc
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            End If
         ElseIf Abs(x - (MargeFil + LongBloc)) < ToleranceAttraction Then
            If Abs(y - 0) < ToleranceAttraction Then
               XPointMesure = MargeFil + LongBloc
               YPointMesure = 0
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            ElseIf Abs(y - HautBloc) < ToleranceAttraction Then
               XPointMesure = MargeFil + LongBloc
               YPointMesure = HautBloc
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            End If
         End If
      End If
      'test sur les angles de la zone de travail
      If flagPointSousCur = False Then
         If Abs(x - 0) < ToleranceAttraction Then
            If Abs(y - MargePlateau) < ToleranceAttraction Then
               XPointMesure = 0
               YPointMesure = MargePlateau
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            ElseIf Abs(y - (CourseY - MargeFil)) < ToleranceAttraction Then
               XPointMesure = 0
               YPointMesure = CourseY - MargeFil
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            End If
         ElseIf Abs(x - CourseX) < ToleranceAttraction Then
            If Abs(y - MargePlateau) < ToleranceAttraction Then
               XPointMesure = CourseX
               YPointMesure = MargePlateau
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            ElseIf Abs(y - (CourseY - MargeFil)) < ToleranceAttraction Then
               XPointMesure = CourseX
               YPointMesure = CourseY - MargeFil
               pctTransf.Cls
               pctTransf.Circle (XPointMesure, YPointMesure), RayonCercles, RGB(156, 86, 228)
               flagPointSousCur = True
            End If
         End If
      End If
      If flagPointSousCur = False Then
         pctTransf.Cls
         If flagPremierPoint = True Then
            lblMesure.Visible = False
         End If
      Else
         If flagPremierPoint = True Then
            lblMesure.Visible = True
         End If
      End If
      If flagPremierPoint = True Then
         pctTransf.Circle (XPremierPoint, YPremierPoint), RayonCercles * 1.3, RGB(156, 86, 228)
         pctTransf.DrawStyle = vbDash
         pctTransf.DrawWidth = 1
         pctTransf.Line (XPremierPoint, YPremierPoint)-(x, y), vbBlack
         pctTransf.DrawStyle = vbSolid
         Xmesure = Abs(XPointMesure - XPremierPoint)
         Ymesure = Abs(YPointMesure - YPremierPoint)
         ValeurMesure = Sqr(Xmesure ^ 2 + Ymesure ^ 2)
         lblMesure.Caption = Format(ValeurMesure, "##0.0")
         lblMesure.Left = (XPointMesure + XPremierPoint) / 2 - lblMesure.Width / 2
         lblMesure.Top = (YPointMesure + YPremierPoint) / 2 + lblMesure.Height
'            lblMesure.Visible = True
         lblMetre.Caption = Label(9) & Format(Xmesure, "##0.0") & Label(10) & Format(Ymesure, "##0.0") & " mm"  'suivant
      End If
      pctTransf.FillStyle = 1  'transparent
      pctTransf.DrawMode = vbCopyPen
      pctTransf.DrawWidth = 1
      Exit Sub  'on sort de la procédure
   End If
   '*********** outil COUPER PROFIL **********
   If OutilEnCours = CouperProfil Or OutilEnCours = PointNumero1 Then
      If NbTransf > 0 Then    ' tracé des points survolé avec les ciseaux ou le premier point
         flagPointSousCur = False
         NumTransfDecoupe = 0
         NumPointDecoupe = 0
         pctTransf.FillColor = RGB(220, 114, 228)
         pctTransf.FillStyle = 0  'rempli
         pctTransf.DrawMode = vbMaskPen
         pctTransf.DrawWidth = 2
         'test sur les points des séquences
         For i = 1 To NbTransf
            With Transf(i)
               For j = 1 To .NbPoints
                  If Abs(x - .Point(j).x) < ToleranceAttraction Then
                     If Abs(y - .Point(j).y) < ToleranceAttraction Then
                        pctTransf.Cls
                        pctTransf.Circle (.Point(j).x, .Point(j).y), RayonCercles, RGB(156, 86, 228)
                        flagPointSousCur = True
                        NumTransfDecoupe = i
                        NumPointDecoupe = j
                     End If
                  End If
               Next j
            End With
         Next i
         If flagPointSousCur = False Then
            pctTransf.Cls
         End If
      End If
      pctTransf.FillStyle = 1  'transparent
      pctTransf.DrawMode = vbCopyPen
      pctTransf.DrawWidth = 1
      Exit Sub
   End If
   'Les outils Mesurer ou CouperProfil ne sont pas sélectionnés
   'les autres outils passent par le drag and drop, donc si on est ici avec le bouton appuyé,
   'c'est qu'on est en train de tracer un rectangle de sélection
   Select Case Button
   Case vbLeftButton    'on est en train de sélectionner
      If flagRectSel = True Then
         pctTransf.Cls
         pctTransf.DrawStyle = vbDash 'trait grands pointillés
         pctTransf.Line (X1SelTransf, Y1SelTransf)-(x, y), vbBlack, B  'tracé du rectangle de sélection
      End If
   Case 0   'si aucun bouton, on gère les curseurs, poignées et éléments de références
      '*********** Gestion des poignées de l'outil ETIRER ********** TRES IMPORTANT, c'est ici qu'est testé la poignée sur laquelle on tire au MouseDown
      If OutilEnCours = Etirer And NbTransfSel > 0 Then
         'si des séquences sont sélectionnées, le rectangle d'étirement et les poignées sont affichées dans TraceTransf
         'changement du curseur en fonction de la poignée
         DemiCarreSelec = 3 * UnPixelToMm
         Poignee = "00"
         If x >= XminSel - DemiCarreSelec And x <= XminSel + DemiCarreSelec Then 'côté gauche du  rectangle de sélection
            If y >= YminSel - DemiCarreSelec And y <= YminSel + DemiCarreSelec Then 'en bas à gauche
               pctTransf.MousePointer = ccSizeNESW
               Poignee = "BG" 'en bas à gauche
            ElseIf y >= (YminSel + YmaxSel) / 2 - DemiCarreSelec And y <= (YminSel + YmaxSel) / 2 + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeEW 'au milieu à gauche
               Poignee = "MG" 'au milieu à gauche
            ElseIf y >= YmaxSel - DemiCarreSelec And y <= YmaxSel + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeNWSE 'en haut à gauche
               Poignee = "HG" 'en haut à gauche
            End If
         ElseIf x >= (XminSel + XmaxSel) / 2 - DemiCarreSelec And x <= (XminSel + XmaxSel) / 2 + DemiCarreSelec Then   'axe vertical du rectangle
            If y >= YminSel - DemiCarreSelec And y <= YminSel + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeNS 'en bas au milieu
               Poignee = "BM" 'en bas au milieu
            ElseIf y >= YmaxSel - DemiCarreSelec And y <= YmaxSel + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeNS 'en haut au milieu
               Poignee = "HM" 'en haut au milieu
            End If
         ElseIf x >= XmaxSel - DemiCarreSelec And x <= XmaxSel + DemiCarreSelec Then 'côté droit du rectangle de sélection
            If y >= YminSel - DemiCarreSelec And y <= YminSel + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeNWSE 'en bas à droite
               Poignee = "BD" 'en bas à droite
            ElseIf y >= (YminSel + YmaxSel) / 2 - DemiCarreSelec And y <= (YminSel + YmaxSel) / 2 + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeEW 'au milieu à droite
               Poignee = "MD" 'au milieu à droite
            ElseIf y >= YmaxSel - DemiCarreSelec And y <= YmaxSel + DemiCarreSelec Then
               pctTransf.MousePointer = ccSizeNESW 'en haut à droite
               Poignee = "HD" 'en haut à droite
            End If
         End If
      End If
      If OutilEnCours = Tourner And NbTransfSel > 0 Then
         pctTransf.Cls
         If x >= XminSel And x <= XmaxSel And y >= YminSel And y <= YmaxSel Then
            pctTransf.Line (Xcentre, Ycentre)-(x, y), vbBlue
         End If
      End If
   End Select
End Sub

'******* Relâchement du bouton dans les séquences transférées ********
Private Sub pctTransf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Long
   Dim L1X As Single, L2X As Single, L1Y As Single, L2Y As Single 'limites de la sélection
   
   '********* fin de rectangle de sélection ***********
   Select Case Button
   Case vbLeftButton
      If flagRectSel = True Then
         pctTransf.Cls
         pctTransf.DrawStyle = vbSolid  'traits continus
         'on calcule les limites de la sélection
         L1X = (x + X1SelTransf) / 2 - Abs((x - X1SelTransf) / 2)
         L2X = (x + X1SelTransf) / 2 + Abs((x - X1SelTransf) / 2)
         L1Y = (y + Y1SelTransf) / 2 - Abs((y - Y1SelTransf) / 2)
         L2Y = (y + Y1SelTransf) / 2 + Abs((y - Y1SelTransf) / 2)
         For i = 1 To NbTransf
            With Transf(i)     'on va différencier les points de passage et les séquences
               If Transf(i).NbPoints = 1 Then   'point de passage, rectangle de 10 x 10 pixels
                  If L1X - 5 < .Point(1).x And L2X + 5 > .Point(1).x And L1Y - 5 < .Point(1).y And L2Y + 5 > .Point(1).y Then
                     If (Transf(i).Etat And 1) = 0 Then
                        Transf(i).Etat = Transf(i).Etat + 1
                        NbTransfSel = NbTransfSel + 1
                        Call InitialiserOutils
                     Else
                        Transf(i).Etat = Transf(i).Etat - 1
                        NbTransfSel = NbTransfSel - 1
                        Call InitialiserOutils
                     End If
                  End If
               Else           'séquence
                  If (L1X < .Xmin And L2X > .Xmax And L1Y < .Ymin And L2Y > .Ymax) Or (L1X > .Xmin And L2X < .Xmax And L1Y > .Ymin And L2Y < .Ymax) Then
                     If (Transf(i).Etat And 1) = 0 Then
                        Transf(i).Etat = Transf(i).Etat + 1
                        NbTransfSel = NbTransfSel + 1
                        Call InitialiserOutils
                     Else
                        Transf(i).Etat = Transf(i).Etat - 1
                        NbTransfSel = NbTransfSel - 1
                        Call InitialiserOutils
                     End If
                  End If
               End If
            End With
         Next i
         Call TraceTransf
      End If
   End Select
End Sub
Private Sub InitialiserOutils()
   flagInitSelection = True
   AgrandissementX = 1
   AgrandissementY = 1
   PositionX = 0
   PositionY = 0
   Rotation = 0
   Select Case OutilEnCours
   Case Etirer, Tourner, Deplacer
      If NbTransfSel > 0 Then
         Call ActiverMesures
      ElseIf NbTransfSel = 0 Then
         Call DesactiverMesures
      End If
   Case CouperProfil, Mesurer
      Call ActiverMesures
   End Select
End Sub
'**************************************************************
'****** Gestion de l'outil de mesure sur événement click ******
'**************************************************************
Private Sub pctTransf_Click()
   Dim Xmesure As Single, Ymesure As Single
   Dim i As Long, j As Long
   Dim Xdebut As Single, Ydebut As Single, Xfin As Single, Yfin As Single

   If OutilEnCours = Mesurer And flagPointSousCur = True Then
      If flagPremierPoint = False Then    'on est sur le premier point de la cote
         flagPremierPoint = True
         XPremierPoint = XPointMesure  'définit dans le mousemove
         YPremierPoint = YPointMesure
      Else  'on est sur le deuxième point de la cote
         Xmesure = Abs(XPointMesure - XPremierPoint)
         Ymesure = Abs(YPointMesure - YPremierPoint)
         ValeurMesure = Sqr((Xmesure) ^ 2 + (Ymesure) ^ 2)
         lblMesure.Caption = Format(ValeurMesure, "##0.0")
         lblMesure.Left = (XPointMesure + XPremierPoint) / 2 - lblMesure.Width / 2
         lblMesure.Top = (YPointMesure + YPremierPoint) / 2 + lblMesure.Height
         lblMesure.Visible = True
         XDernierPoint = XPointMesure
         YDernierPoint = YPointMesure
         flagPremierPoint = False
         flagMemoUndoDansTraceTransf = False
         Call TraceTransf
         'tracé de la cote et de ses extrémités
         pctTransf.AutoRedraw = True
         pctTransf.DrawStyle = vbDash
         pctTransf.DrawWidth = 1
         pctTransf.Line (XPremierPoint, YPremierPoint)-(XPointMesure, YPointMesure), vbRed
         pctTransf.Line (XPremierPoint - 5, YPremierPoint - 5)-(XPremierPoint + 5, YPremierPoint + 5), vbRed
         pctTransf.Line (XPremierPoint - 5, YPremierPoint + 5)-(XPremierPoint + 5, YPremierPoint - 5), vbRed
         pctTransf.Line (XPointMesure - 5, YPointMesure - 5)-(XPointMesure + 5, YPointMesure + 5), vbRed
         pctTransf.Line (XPointMesure - 5, YPointMesure + 5)-(XPointMesure + 5, YPointMesure - 5), vbRed
         pctTransf.DrawStyle = vbSolid
         pctTransf.AutoRedraw = False
         lblMetre.Caption = Label(9) & Format(Xmesure, "##0.0") & Label(10) & Format(Ymesure, "##0.0") & " mm"
      End If
   End If
   If optOutils(CouperProfil).Value = True And NbTransf > 0 And flagPointSousCur = True Then
      If NumTransfDecoupe > 0 Then 'on a cliqué sur un point d'une séquence
         If NumPointDecoupe > 1 And NumPointDecoupe < Transf(NumTransfDecoupe).NbPoints Then 'pas sur le premier ou le dernier point
            ReDim Preserve Transf(1 To NbTransf + 1)  'on agrandit le tableau d'une séquence
            If NumTransfDecoupe < NbTransf Then 'si on est sur la dernière séquence, pas besoin de faire de transfert
               For i = NbTransf To NumTransfDecoupe + 1 Step -1
                  Transf(i + 1) = Transf(i)
               Next i
            End If
            NbTransf = NbTransf + 1
            With Transf(NumTransfDecoupe + 1)
               .NbPoints = Transf(NumTransfDecoupe).NbPoints - NumPointDecoupe + 1
               ReDim .Point(1 To .NbPoints)
               j = NumPointDecoupe
               For i = 1 To .NbPoints
                  .Point(i) = Transf(NumTransfDecoupe).Point(j)
                  j = j + 1
               Next i
            End With
            Call MaxiMiniSequ(Transf(NumTransfDecoupe + 1))
            ReDim Preserve Transf(NumTransfDecoupe).Point(1 To NumPointDecoupe)  'on efface la fin de la séquence
            Transf(NumTransfDecoupe).NbPoints = NumPointDecoupe
            Call MaxiMiniSequ(Transf(NumTransfDecoupe))
            Call DeselecTransf 'on déselectionne tout pour pouvoir facilement déplacer ce qui a été coupé
            If chkCouleurProfils.Value = vbUnchecked Then
               chkCouleurProfils.Value = vbChecked
            Else
               Call TraceTransf
            End If
         ElseIf NumPointDecoupe = 1 Then
            MsgBox Message(Corps, 13), vbInformation, Message(Titre, 13)  'coupe impossible sur point 1
         ElseIf NumPointDecoupe = Transf(NumTransfDecoupe).NbPoints Then
            MsgBox Message(Corps, 14), vbInformation, Message(Titre, 14)  'coupe impossible sur dernier point
         End If
      End If
   End If
   If optOutils(PointNumero1).Value = True And NbTransf > 0 And flagPointSousCur = True Then 'renumérotation
      If NumTransfDecoupe > 0 Then 'on a cliqué sur un point d'une séquence
         If NumPointDecoupe > 1 Then 'pas sur le premier point
            With Transf(NumTransfDecoupe)
               'Si le profil est fermé, on vire le dernier point
               Xdebut = .Point(1).x
               Ydebut = .Point(1).y
               Xfin = .Point(.NbPoints).x
               Yfin = .Point(.NbPoints).y
               If Abs(Xdebut - Xfin) < 0.001 And Abs(Ydebut - Yfin) < 0.001 Then
                  .NbPoints = .NbPoints - 1
                  ReDim Preserve .Point(1 To .NbPoints)
               End If
               'puis on renumérote
               ReDim Preserve .Point(1 To .NbPoints + NumPointDecoupe - 1)
               For i = 1 To NumPointDecoupe - 1
                  .Point(.NbPoints + i).x = .Point(i).x
                  .Point(.NbPoints + i).y = .Point(i).y
               Next i
               For i = 1 To .NbPoints
                  .Point(i).x = .Point(i + NumPointDecoupe - 1).x
                  .Point(i).y = .Point(i + NumPointDecoupe - 1).y
               Next i
               .NbPoints = .NbPoints + 1
               ReDim Preserve .Point(1 To .NbPoints)
               .Point(.NbPoints) = .Point(1)
            End With
            Call TraceTransf
         End If
      End If
   End If
End Sub

'******* Mise à jour de l'affichage en cas de changement d'outil (sens des séquences notamment) ********
'******* Mise à jour du champ de saisie manuelle en cas de changement d'outils *************************
Private Sub optOutils_Click(Index As Integer)
   OutilEnCours = Index 'on mémorise l'outil en cours d'utilisation
   Select Case Index
   Case Mesurer      'définition du curseur personnalisé
      If chkVoirPoints.Value = vbUnchecked Then chkVoirPoints.Value = vbChecked 'on visualise les points
      pctTransf.MouseIcon = ImageList2.ListImages(1).Picture
   Case CouperProfil
      If chkVoirPoints.Value = vbUnchecked Then chkVoirPoints.Value = vbChecked 'on visualise les points
      pctTransf.MouseIcon = ImageList2.ListImages(10).Picture
   Case Tourner
      If chkVoirPoints.Value = vbChecked Then chkVoirPoints.Value = vbUnchecked 'on n'a pas besoin des points
      pctTransf.MouseIcon = ImageList2.ListImages(4).Picture
   Case Deplacer
      If chkVoirPoints.Value = vbChecked Then chkVoirPoints.Value = vbUnchecked 'on n'a pas besoin des points
      pctTransf.MouseIcon = ImageList2.ListImages(2).Picture
   Case Etirer
      If chkVoirPoints.Value = vbChecked Then chkVoirPoints.Value = vbUnchecked 'on n'a pas besoin des points
      pctTransf.MouseIcon = ImageList2.ListImages(5).Picture
   Case PointNumero1
      If chkVoirPoints.Value = vbUnchecked Then chkVoirPoints.Value = vbChecked 'on visualise les points
      pctTransf.MouseIcon = ImageList2.ListImages(12).Picture
   End Select
   pctTransf.MousePointer = 99   'application du curseur personnalisé
   lblMesure.Visible = False  'initialisation du texte de cotation
   If Index = CouperProfil Then
      chkCouleurProfils.Value = vbChecked  'déclenche le tracetransf
   Else
      flagMemoUndoDansTraceTransf = False
      Call TraceTransf
   End If
   Select Case OutilEnCours
   Case Etirer, Tourner, Deplacer
      If NbTransfSel = 0 Then
         Call DesactiverMesures     'pour affichage des bons textbox
      ElseIf NbTransfSel > 0 Then
         Call ActiverMesures
      End If
   Case CouperProfil, Mesurer, PointNumero1
      Call ActiverMesures
   End Select
End Sub

Private Sub txtMesures_GotFocus(Index As Integer)
   txtMesures(Index).SelStart = 0
   txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
   txtBloc(0).TabStop = False
   txtBloc(1).TabStop = False
   txtMesures(0).TabStop = True
   txtMesures(1).TabStop = True
End Sub

Private Sub txtMesures_LostFocus(Index As Integer)
   Select Case OutilEnCours
   Case Deplacer
      If GetTabState Then  'la perte de focus s'est produite par appui de la touche Tab
         Call ValiderDeplacer(Index)
         If Index = 1 Then
            txtMesures(0).SetFocus
         ElseIf Index = 0 Then
            txtMesures(1).SetFocus
         End If
      Else
         txtMesures(0).Text = Format(PositionX, "##0.0")
         txtMesures(1).Text = Format(PositionY, "##0.0")
      End If
   Case Tourner
      If GetTabState Then  'la perte de focus s'est produite par appui de la touche Tab
         Call ValiderTourner(Index)
         txtMesures(0).SetFocus
      Else
         txtMesures(0).Text = Format(Rotation, "##0.0")
      End If
   Case Etirer
      If GetTabState Then  'la perte de focus s'est produite par appui de la touche Tab
         Call ValiderEtirer(Index)
         If Index = 1 Then
            txtMesures(0).SetFocus
         ElseIf Index = 0 Then
            txtMesures(1).SetFocus
         End If
      Else
         txtMesures(0).Text = Format(AgrandissementX * 100, "##0.0")
         txtMesures(1).Text = Format(AgrandissementY * 100, "##0.0")
      End If
   End Select
End Sub

'***********************************
'**** DEPLACEMENT de la sélection ****
'***********************************
Private Sub txtMesures_KeyPress(Index As Integer, KeyAscii As Integer)
   'vérification de la saisie à l'appui sur les touches
   If optOutils(Deplacer).Value = True Or optOutils(Tourner).Value = True Or optOutils(Etirer).Value = True Then
      Call VerifierDecimauxRelatifs(KeyAscii)  'Keyascii est passé par référence
   End If
End Sub
'********* Procédures de validation des saisies clavier *************
Private Sub ValiderDeplacer(Index As Integer)
   Dim i As Long, j As Long
   
   If Not MyIsNumeric(MyVal(txtMesures(Index).Text)) Then
      MsgBox Message(Corps, 15), vbCritical + vbOKOnly, Message(Titre, 15)    'erreur de saisie, décimal positif ou négatif
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
   Else
      For j = 1 To NbTransf
         If (Transf(j).Etat And 1) = 1 Then
            With Transf(j)
               For i = 1 To .NbPoints
                  If Index = 0 Then    'déplacement suivant X
                     .Point(i).x = .Point(i).x + MyVal(txtMesures(Index).Text) - PositionX
                     .Point(i).y = .Point(i).y
                  ElseIf Index = 1 Then   'déplacement suivant Y
                     .Point(i).x = .Point(i).x
                     .Point(i).y = .Point(i).y + MyVal(txtMesures(Index).Text) - PositionY
                  End If
               Next i
            End With
            Call MaxiMiniSequ(Transf(j))
         End If
      Next j
      PositionX = MyVal(txtMesures(0).Text)
      PositionY = MyVal(txtMesures(1).Text)
      txtMesures(0).Text = Format(PositionX, "##0.0")
      txtMesures(1).Text = Format(PositionY, "##0.0")
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub

Private Sub ValiderTourner(Index As Integer)
   Dim i As Long, j As Long
   
   Dim AngleTemp As Single
   If Not MyIsNumeric(txtMesures(Index).Text) Then
      MsgBox Message(Corps, 15), vbCritical + vbOKOnly, Message(Titre, 15)    'erreur de saisie, décimal positif ou négatif
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
   Else
      AngleTemp = MyVal(txtMesures(Index).Text) * pi / 180
      Xcentre = (XminSel + XmaxSel) / 2
      Ycentre = (YminSel + YmaxSel) / 2
      For j = 1 To NbTransf
         If (Transf(j).Etat And 1) = 1 Then
            With Transf(j)
               For i = 1 To .NbPoints
                  Xtemp = .Point(i).x
                  Ytemp = .Point(i).y
                  .Point(i).x = Xcentre + (Xtemp - Xcentre) * Cos(AngleTemp - Rotation) - (Ytemp - Ycentre) * Sin(AngleTemp - Rotation)
                  .Point(i).y = Ycentre + (Xtemp - Xcentre) * Sin(AngleTemp - Rotation) + (Ytemp - Ycentre) * Cos(AngleTemp - Rotation)
               Next i
            End With
            Call MaxiMiniSequ(Transf(j))
         End If
      Next j
      Rotation = MyVal(txtMesures(Index).Text) * pi / 180
      txtMesures(Index).Text = Format((Rotation * 180 / pi + 360) Mod 360, "##0.0")
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub
Private Sub ValiderEtirer(Index As Integer)
   Dim i As Long, j As Long
   
   If Not MyIsNumeric(txtMesures(Index).Text) Then
      MsgBox Message(Corps, 15), vbCritical + vbOKOnly, Message(Titre, 15)    'erreur de saisie, décimal positif ou négatif
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
   Else
      OrigineX = (XmaxSel + XminSel) / 2
      OrigineY = (YmaxSel + YminSel) / 2
      kX = MyVal(txtMesures(0).Text) / 100
      kY = MyVal(txtMesures(1).Text) / 100
      For j = 1 To NbTransf
         If (Transf(j).Etat And 1) = 1 Then
            With Transf(j)
               For i = 1 To .NbPoints
                  .Point(i).x = kX / AgrandissementX * (.Point(i).x - OrigineX) + OrigineX
                  .Point(i).y = kY / AgrandissementY * (.Point(i).y - OrigineY) + OrigineY
               Next i
            End With
            Call MaxiMiniSequ(Transf(j))
         End If
      Next j
      AgrandissementX = MyVal(txtMesures(0).Text) / 100
      AgrandissementY = MyVal(txtMesures(1).Text) / 100
      txtMesures(0).Text = Format(AgrandissementX * 100, "##0.0")
      txtMesures(1).Text = Format(AgrandissementY * 100, "##0.0")
      txtMesures(Index).SelStart = 0
      txtMesures(Index).SelLength = Len(txtMesures(Index).Text)
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub

Private Sub txtMesures_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   
   '************************************
   '**** Saisie manuelle des outils ****
   '************************************
   If KeyCode = 13 Then 'validation par entrée
      If OutilEnCours = Deplacer Then
         Call ValiderDeplacer(Index)
      ElseIf OutilEnCours = Tourner Then
         Call ValiderTourner(Index)
      ElseIf OutilEnCours = Etirer Then
         Call ValiderEtirer(Index)
      End If
   End If
End Sub

'*****************************************************************************
'***** Insérer un point de déviation de trajectoire entre deux séquences *****
'*****************************************************************************
Private Sub cmdInsererPoint_Click()
   Dim i As Long, j As Long
   
   If NbTransfSel >= 2 Then
      i = 0
      Do
         i = i + 1
         If (Transf(i).Etat And 1) = 1 And (Transf(i + 1).Etat And 1) = 1 Then
            ReDim Preserve Transf(1 To NbTransf + 1)
            NbTransf = NbTransf + 1
            For j = NbTransf To i + 2 Step -1
               Transf(j) = Transf(j - 1)
            Next j
            With Transf(i + 1)
               .NbPoints = 1
               .Point(1).x = (Transf(i).Point(Transf(i).NbPoints).x + Transf(i + 2).Point(1).x) / 2
               .Point(1).y = (Transf(i).Point(Transf(i).NbPoints).y + Transf(i + 2).Point(1).y) / 2
               .Etat = 99  'l'état 99 est l'état d'un point qui vient d'être créé : utile pour la gestion des sélections ci-dessous
            End With
            i = i + 1
         End If
         If i + 1 = NbTransf Then Exit Do  'on est arrivé à la fin
      Loop
      'on sélectionne uniquement les points créés
      For i = 1 To NbTransf
         If Transf(i).Etat = 99 Then
            Transf(i).Etat = 1
         Else
            If (Transf(i).Etat And 1) = 1 Then  'si la séquence est sélectionnée, on la désélectionne
               Transf(i).Etat = Transf(i).Etat - 1
            End If
         End If
         Call MaxiMiniSequ(Transf(i))
      Next i
   End If
   If NbTransf > 0 Then
      Call InitialiserOutils  'activation des textbox de saisie manuelle
      'on calcule l'affichage avant toute chose pour permettre le "change"
      Call TraceTransf
   End If
End Sub

'*********************************************
'****** GESTION DES OUTILS D'ALIGNEMENT ******
'*********************************************
Private Sub cmdAligner_Click(Index As Integer)
   Dim i As Long, j As Long
   
   If NbTransf > 0 Then
      Select Case Index
      'déplacement suivant Y pour aligner sur le plus bas
      Case 0
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).y = .Point(j).y - (.Ymin - YminSel)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      'déplacement suivant Y pour aligner au milieu
      Case 1
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).y = .Point(j).y - ((.Ymax + .Ymin) / 2 - (YmaxSel + YminSel) / 2)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      'déplacement suivant Y pour aligner sur le plus haut
      Case 2
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).y = .Point(j).y + (YmaxSel - .Ymax)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      'déplacement suivant X pour aligner sur la gauche
      Case 3
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).x = .Point(j).x - (.Xmin - XminSel)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      'déplacement suivant X pour aligner au milieu
      Case 4
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).x = .Point(j).x - ((.Xmax + .Xmin) / 2 - (XmaxSel + XminSel) / 2)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      'déplacement suivant X pour aligner à droite
      Case 5
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call MaxiMiniSequ(Transf(i))
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).x = .Point(j).x + (XmaxSel - .Xmax)
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))
         End If
      Next i
      End Select
      'on calcule l'affichage avant toute chose pour permettre le "change"
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub

'****************************************
'***** Gestion de l'outil Dupliquer *****
'****************************************
Private Sub cmdDupliquer_Click()
   Dim i As Long, j As Long
   Dim MemoNbTransf As Long
   
   If NbTransfSel > 0 Then
      MemoNbTransf = NbTransf
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            ReDim Preserve Transf(1 To NbTransf + 1)
            NbTransf = NbTransf + 1
            Transf(NbTransf) = Transf(i)
            Transf(i).Etat = Transf(i).Etat - 1
         End If
      Next i
      If MemoNbTransf <> NbTransf Then
         For i = MemoNbTransf + 1 To NbTransf
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).x = .Point(j).x + 30
                  .Point(j).y = .Point(j).y + 30
               Next j
            End With
         Call MaxiMiniSequ(Transf(i))  'important pour pouvoir sélectionner la nouvelle séquence
         Next i
      End If
      ' Call ZoomAutoToutVoir
      Call TraceTransf
      optOutils(Deplacer).Value = True
   End If
End Sub

'***********************************
'**** Outil d'INVERSION de sens ****
'***********************************
Private Sub cmdInverser_Click()
   Dim i As Long
   
   If NbTransfSel > 0 Then
      Call InverserOrdreTransf   'l'ordre des séquences qui se suivent doit être inversé
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            Call InverserSensSequ(Transf(i)) 'l'ordre des points des séquences sélectionnées doit être inversé
         End If
      Next i
   End If
   Call TraceTransf
End Sub

'******************
'**** Poubelle ****
'******************
Private Sub cmdPoubelle_Click()
   Call Form_KeyUp(vbKeyDelete, 0)
End Sub

'***************************
'**** Miroir horizontal ****
'***************************
Private Sub cmdMiroir_Click()
   Dim i As Long, j As Long
   Dim XminiMiroir As Single, XmaxiMiroir As Single
   
   If NbTransfSel > 0 Then
      For i = 1 To NbTransf
         If (Transf(i).Etat And 1) = 1 Then
            XminiMiroir = Transf(i).Xmin
            XmaxiMiroir = Transf(i).Xmax
            Exit For
         End If
      Next i
      For i = 1 To NbTransf
         With Transf(i)
            If (.Etat And 1) = 1 Then
               If .Xmin < XminiMiroir Then XminiMiroir = .Xmin
               If .Xmax > XmaxiMiroir Then XmaxiMiroir = .Xmax
            End If
         End With
      Next i
      For i = 1 To NbTransf
         With Transf(i)
            If (.Etat And 1) = 1 Then
               For j = 1 To .NbPoints
                  .Point(j).x = XminiMiroir + XmaxiMiroir - .Point(j).x
               Next j
            End If
            Call MaxiMiniSequ(Transf(i))
         End With
      Next i
      Call TraceTransf
   End If
End Sub

'****** Zoom sur le bloc du projet *******
Private Sub chkZoomProjet_Click()
   If chkZoomProjet.Value = vbChecked Then
      Call EchelleTransf(0, LongBloc, 0, HautBloc)
   ElseIf chkZoomProjet.Value = vbUnchecked Then
      Call ZoomAutoToutVoir
   End If
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub
'****** curseur pour survol *****
Private Sub chkZoomProjet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   pctTransf.MousePointer = 0
End Sub
'******** Agrandir / Rétrécir la fenêtre de travail *******
Private Sub cmdAgrandirRetrecir_Click()
   lblMesure.Visible = False  'initialisation du texte de cotation
   If flagFenetreAgrandie = False Then
      cmdAgrandirRetrecir.Picture = frmImages.imgRetrecir.Picture
      flagFenetreAgrandie = True
      cmdMiroirFichierSource.Visible = False
      cmdInverserSensSequ.Visible = False
      cmdAfficherSensSequ.Visible = False
      Call TailleFenetre
   Else
      cmdAgrandirRetrecir.Picture = frmImages.imgAgrandir.Picture
      flagFenetreAgrandie = False
      cmdMiroirFichierSource.Visible = True
      cmdInverserSensSequ.Visible = True
      cmdAfficherSensSequ.Visible = True
      Call TailleFenetre
   End If
   Call ZoomAutoToutVoir
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub

'******** Survol du bouton d'agrandissement de la fenêtre de travail ***********
Private Sub cmdAgrandirRetrecir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   pctTransf.MousePointer = 0
End Sub

'********** Affichage des points ********
Private Sub chkVoirPoints_Click()
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub

Private Sub chkVoirPoints_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   pctTransf.MousePointer = 0
End Sub
'********** Alternance de couleurs ********
Private Sub chkCouleurProfils_Click()
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub

Private Sub chkCouleurProfils_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   pctTransf.MousePointer = 0
End Sub

'*****************************************
'**** GESTION des BOUTONS UNDO / REDO ****
'*****************************************
Private Sub cmdUndo_Click(Index As Integer)
   Dim i As Long
   
   'gestion de avant-arrière et de l'endroit où on se trouve dans le tableau UndoRedo()
   Select Case Index
   Case 0   'si on va en arrière : il faut qu'il y ait un élément dans le tableau
      If IndexUndo >= 2 Then
         IndexUndo = IndexUndo - 1
         cmdUndo(1).Enabled = True
      End If
   Case 1   'si on va en avant, ça veut dire qu'on est allé en arrière, on peut avancer jusqu'à la fin du tableau
      If IndexUndo < UBound(UndoRedo) Then
         IndexUndo = IndexUndo + 1
         cmdUndo(0).Enabled = True
      End If
   End Select
   
   If IndexUndo = 1 Then cmdUndo(0).Enabled = False
   If IndexUndo = UBound(UndoRedo) Then cmdUndo(1).Enabled = False
   
   '****
   With UndoRedo(IndexUndo)
      NbTransf = 0  'le fait de mettre cette variable à zéro permet d'annuler l'effet du clic sur les chkCentrer
      If .CheckAjusterEchelle = True Then chkCentrer(0).Value = vbChecked
      If .CheckCentrerX = True Then chkCentrer(1).Value = vbChecked
      If .CheckCentrerY = True Then chkCentrer(2).Value = vbChecked
      If .NbTransfUndo = 0 Then
         Erase Transf
      ElseIf .NbTransfUndo > 0 Then
         For i = 0 To 2
            chkCentrer(i).Value = vbUnchecked
         Next i
         NbTransf = .NbTransfUndo
         NbTransfSel = .NbTransfSelUndo
         ReDim Transf(1 To NbTransf)
         For i = 1 To NbTransf
            Transf(i) = .TransfUndo(i)
         Next i
      Else
         MsgBox "NbTransf <0 dans cmdUndo_click => contact@minicut2d.com", vbCritical, "Erreur"
      End If
      HautBloc = .HautBlocUndo
      LongBloc = .LongBlocUndo
      txtBloc(0).Text = Format(HautBloc, "##0")
      txtBloc(1).Text = Format(LongBloc, "##0")
      CoeffBloc = .CoeffBlocUndo
   End With
   flagMemoUndoDansTraceTransf = False
   Call TraceTransf
End Sub

'***************************************
'**** Mémorisation pour UNDO / REDO ****
'***************************************
Private Sub MemoriserPourUndo()
   Dim i As Long
   Dim NombreUndoMaxi As Integer
   
   NombreUndoMaxi = 20  'on autorise 20 retours en arrière
   'IndexUndo est initialisé à 0 dans le Form Load et UndoRedo à (1 to 1)
   If UBound(UndoRedo) < NombreUndoMaxi Then
      IndexUndo = IndexUndo + 1
      ReDim Preserve UndoRedo(1 To IndexUndo) 'le redim rajoute une ligne
   ElseIf UBound(UndoRedo) = NombreUndoMaxi Then
      For i = 1 To NombreUndoMaxi - 1  'à partir de NombreUndoMaxi, on remonte tout d'une ligne et on stocke dans la dernière
         UndoRedo(i) = UndoRedo(i + 1)
      Next i
      IndexUndo = NombreUndoMaxi
   End If
   If IndexUndo > 1 Then
      cmdUndo(0).Enabled = True  'à partir du moment où il y a une mémorisation, on peut revenir en arrière
      cmdUndo(1).Enabled = False 'mais la suite est effacée : on ne peut pas refaire
   End If
   With UndoRedo(IndexUndo)
      .NbTransfUndo = NbTransf
      .NbTransfSelUndo = NbTransfSel
      If .NbTransfUndo > 0 Then
         ReDim .TransfUndo(1 To .NbTransfUndo)
         For i = 1 To NbTransf
            .TransfUndo(i) = Transf(i)
         Next i
      ElseIf .NbTransfUndo = 0 Then
         Erase .TransfUndo
      End If
      .HautBlocUndo = HautBloc
      .LongBlocUndo = LongBloc
      .CoeffBlocUndo = CoeffBloc
      If chkCentrer(0).Value = vbChecked Then
         .CheckAjusterEchelle = True
      ElseIf chkCentrer(0).Value = vbUnchecked Then
         .CheckAjusterEchelle = False
      End If
      If chkCentrer(1).Value = vbChecked Then
         .CheckCentrerX = True
      ElseIf chkCentrer(1).Value = vbUnchecked Then
         .CheckCentrerX = False
      End If
      If chkCentrer(2).Value = vbChecked Then
         .CheckCentrerY = True
      ElseIf chkCentrer(2).Value = vbUnchecked Then
         .CheckCentrerY = False
      End If
   End With
End Sub

'******** Boutons d'AJUSTEMENT par rapport au bloc (ajuster, centrer X, centrer Y) *********
Private Sub chkCentrer_Click(Index As Integer)
   Dim i As Integer
   
   For i = 0 To 2
      If chkCentrer(i).Value = vbChecked Then
         chkCentrer(i).BackColor = &HFF
      Else
         chkCentrer(i).BackColor = &H8000000F
      End If
   Next i
   If NbTransf > 0 Then
      Call AppliquerCentrageEtAjustement
      ' Call ZoomAutoToutVoir
      Call TraceTransf
   End If
End Sub

'************* Ajustement des séquences par rapport au bloc **************
Private Sub AppliquerCentrageEtAjustement()
   Dim i As Long, j As Long
   Dim CoeffAjustement As Single
   Dim L As Single, H As Single
   
   'Index=0 : Ajustement au bloc : la découpe est mise à l'échelle du bloc
   'Index=1 : la découpe est juste centrée suivant X
   'Index=2 : la découpe est juste centrée suivant Y
   If chkCentrer(0).Value = vbChecked Then
      CoeffAjustement = 1
      Call MaxiMiniTotalTransff
      DeltaXMaxTransf = XmaxTransf - XminTransf
      DeltaYMaxTransf = YmaxTransf - YminTransf
      L = LongBloc - 2 * MargeInterieureX
      H = HautBloc - 2 * MargeInterieureY
      If DeltaXMaxTransf <> 0 And DeltaYMaxTransf <> 0 Then
         If DeltaXMaxTransf > L And DeltaYMaxTransf <= H Then 'les séquences dépassent en X
            CoeffAjustement = L / DeltaXMaxTransf
         ElseIf DeltaYMaxTransf > H And DeltaXMaxTransf <= L Then 'les séquences dépassent en Y
            CoeffAjustement = H / DeltaYMaxTransf
         ElseIf DeltaXMaxTransf > L And DeltaYMaxTransf > H Then  'trop grand sur X et Y
            If L / DeltaXMaxTransf <= H / DeltaYMaxTransf Then
               CoeffAjustement = L / DeltaXMaxTransf
            Else
               CoeffAjustement = H / DeltaYMaxTransf
            End If
         ElseIf DeltaXMaxTransf <= L And DeltaYMaxTransf <= H Then 'trop petit sur X et Y
            If L / DeltaXMaxTransf <= H / DeltaYMaxTransf Then
               CoeffAjustement = L / DeltaXMaxTransf
            Else
               CoeffAjustement = H / DeltaYMaxTransf
            End If
         Else
            MsgBox "pas de coeff défini => contact@minicut2d.com"
         End If
         'on fait une homothétie de centre le point le plus en bas à gauche
         For i = 1 To NbTransf
            With Transf(i)
               For j = 1 To .NbPoints
                  .Point(j).x = CoeffAjustement * (.Point(j).x - XminTransf) + XminTransf
                  .Point(j).y = CoeffAjustement * (.Point(j).y - YminTransf) + YminTransf
               Next j
            End With
            Call MaxiMiniSequ(Transf(i))   'mise à jour pour éviter bugs affichage, distance, etc.
         Next i
         CoeffBloc = CoeffBloc * CoeffAjustement  'pour que les transferts suivants se fassent à la même échelle
      Else
         MsgBox Message(Corps, 16), vbInformation, Message(Titre, 16) 'calcul impossible, une dimension =0
         Exit Sub
      End If
   Else
      CoeffBloc = 1
   End If
   If chkCentrer(1).Value = vbChecked Then
   'translation suivant X
      Call MaxiMiniTotalTransff  'pour définir le centre
      XcentreTransff = (XminTransf + XmaxTransf) / 2
      For i = 1 To NbTransf
         With Transf(i)
            For j = 1 To .NbPoints
            .Point(j).x = .Point(j).x + ((MargeFil + LongBloc / 2) - XcentreTransff)
            Next j
         End With
         Call MaxiMiniSequ(Transf(i))   'mise à jour pour éviter bugs affichage, distance, etc.
      Next i
   End If
   If chkCentrer(2).Value = vbChecked Then
   'translation suivant Y
      Call MaxiMiniTotalTransff  'pour définir le centre
      YcentreTransff = (YminTransf + YmaxTransf) / 2
      For i = 1 To NbTransf
         With Transf(i)
            For j = 1 To .NbPoints
            .Point(j).y = .Point(j).y + ((HautBloc / 2) - YcentreTransff)
            Next j
         End With
         Call MaxiMiniSequ(Transf(i))   'mise à jour pour éviter bugs affichage, distance, etc.
      Next i
   End If
End Sub

'************ AJUSTEMENT AUTOMATIQUE du ZOOM EN CAS DE DEPASSEMENT DE LA ZONE UTILE LORS DES TRANSFERTS **************
Public Sub ZoomAutoToutVoir()
   Dim Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single
   
   If chkZoomProjet.Value = vbUnchecked Then    'l'ajustement automatique ne fonctionne pas en mode zoom bloc
      Call MaxiMiniTotalTransff
      
      If XminTransf >= 0 Then Xmin = 0 Else Xmin = XminTransf
      If XmaxTransf < CourseX Then Xmax = CourseX Else Xmax = XmaxTransf
      If YminTransf >= 0 Then Ymin = 0 Else Ymin = YminTransf
      If YmaxTransf < CourseY Then Ymax = CourseY Else Ymax = YmaxTransf
      
      Call EchelleTransf(Xmin, Xmax, Ymin, Ymax)
   End If
End Sub

'********* Gestion des onglets **********
Private Sub SSTab1_Click(PreviousTab As Integer)
   Select Case SSTab1.Tab
   Case 0   'onglet de création, manipulation des profils
      cmdMiroirFichierSource.Visible = True
      cmdInverserSensSequ.Visible = True
      cmdAfficherSensSequ.Visible = True
      chkVoirPoints.Visible = True
      chkCouleurProfils.Visible = True
      cmdAgrandirRetrecir.Visible = True
      chkZoomProjet.Visible = True
      chkZoomDecoupe.Visible = False
      pctZoomInfo.Visible = True
      
      cmdSimulation.Visible = False
      pctSequ.Visible = True
      pctTransf.Visible = True
      pctBandeauSaisie.Visible = True
      pctDecoupe.Visible = False
      flagMemoUndoDansTraceTransf = False
      optOutils(0).Value = True
      Call TailleFenetre
      
'      Call TraceTransf
   Case 1   'onglet de gestion de la découpe
      cmdMiroirFichierSource.Visible = False
      cmdInverserSensSequ.Visible = False
      cmdAfficherSensSequ.Visible = False
      chkVoirPoints.Visible = False
      chkCouleurProfils.Visible = False
      cmdAgrandirRetrecir.Visible = False
      chkZoomProjet.Visible = False
      chkZoomDecoupe.Visible = True
      pctZoomInfo.Visible = False
      
      cmdSimulation.Visible = True
      'on masque les box de travail sur les profils
      pctSequ.Visible = False
      pctTransf.Visible = False
      pctBandeauSaisie.Visible = False
      lblDimensionSelection.Visible = False
      
      Call InitialisationDessinDecoupe  'valable pour le passage dans l'onglet ou pour l'ouverture d'un nouveau projet
   End Select
End Sub
'*********** Dessin de la table et de la découpe : soit au passage dans l'onglet, soit à l'ouverture d'un projet en étant déjà dans l'onglet *********
Public Sub InitialisationDessinDecoupe()
   'on trace la table dans la box de découpe avant de la rendre visible (si elle ne l'est pas)
   Call CalculRepAffDecoupe   'mise à l'échelle de la box
   Call TraceTableDecoupe     'tracé des rectangles et des lignes constitutifs du dessin
   pctDecoupe.Visible = True  'si changement d'onglet
   Call TraceBlocEtOrigine    'tracé du bloc, de la zone utile, du cercle de l'origine
   'découpe à proprement parler
   If NbTransf > 0 Then 'il faut qu'il y ait des séquences
      Call AssemblageSequencesPourDecoupe    'on assemble toutes les séquences en une seule appelée SequDecoupe
      Call CalculDepassementCourses(SequDecoupe) '(modDepassements) : On calcule le dépassement du bloc ou des courses pour éventuellement modifier le profil
       If flagDepassementDecoupe = True Then
         MsgBox Message(Corps, 17), vbInformation, Message(Titre, 17) 'dépassement => projet tronqué
         flagDepassementDecoupe = False
      End If
      Call CreationTrajetsEntreeSortie 'en fonction de la position des boutons
      Call TraceDecoupe   'contrôle du nombre de points et tracé de la découpe
   End If
End Sub

Public Sub CalculRepAffDecoupe()
   If chkZoomDecoupe.Value = vbUnchecked Then
      'Calcul des échelles d'affichage de la table réelle dans le repère de la picturebox de découpe à partir des coordonnées réelles
      If pctDecoupe.Width / Abs(MaxiDecoupeX - MiniDecoupeX) > pctDecoupe.Height / Abs(MaxiDecoupeY - MiniDecoupeY) Then
         pctDecoupe.ScaleHeight = -Abs(MaxiDecoupeY - MiniDecoupeY) + 1   ' le 1 sert à la marge en haut ; le "-" devant inverse le sens de l'axe
         pctDecoupe.ScaleTop = MaxiDecoupeY + 1
         pctDecoupe.ScaleWidth = pctDecoupe.Width * Abs(pctDecoupe.ScaleHeight) / pctDecoupe.Height
         pctDecoupe.ScaleLeft = MiniDecoupeX - (Abs(pctDecoupe.ScaleWidth) - Abs(MaxiDecoupeX - MiniDecoupeX)) / 2
      Else
         pctDecoupe.ScaleLeft = MiniDecoupeX + 1   'le +1 sert à faire disparaître le bord du rectangle
         pctDecoupe.ScaleWidth = Abs(MaxiDecoupeX - MiniDecoupeX) + 1   ' le 1 sert à la marge à droite
         pctDecoupe.ScaleHeight = -pctDecoupe.Height * pctDecoupe.ScaleWidth / pctDecoupe.Width
         pctDecoupe.ScaleTop = MaxiDecoupeY + (Abs(pctDecoupe.ScaleHeight) - Abs(MaxiDecoupeY - MiniDecoupeY)) / 2
      End If
   ElseIf chkZoomDecoupe.Value = vbChecked Then
      'Calcul des échelles d'affichage de la table réelle dans le repère de la picturebox de découpe à partir des coordonnées réelles
      If pctDecoupe.Width / Abs(LongBloc + MargeFil) > pctDecoupe.Height / Abs(HautBloc) Then
         pctDecoupe.ScaleHeight = -Abs(HautBloc) + 1   ' le 1 sert à la marge en haut ; le "-" devant inverse le sens de l'axe
         pctDecoupe.ScaleTop = HautBloc + 1
         pctDecoupe.ScaleWidth = pctDecoupe.Width * Abs(pctDecoupe.ScaleHeight) / pctDecoupe.Height
         pctDecoupe.ScaleLeft = 0 - (Abs(pctDecoupe.ScaleWidth) - Abs(LongBloc)) / 2
      Else
         pctDecoupe.ScaleLeft = 1   'le +1 sert à faire disparaître le bord du rectangle
         pctDecoupe.ScaleWidth = Abs(LongBloc + MargeFil) + 1 ' le 1 sert à la marge à droite
         pctDecoupe.ScaleHeight = -pctDecoupe.Height * pctDecoupe.ScaleWidth / pctDecoupe.Width
         pctDecoupe.ScaleTop = HautBloc + (Abs(pctDecoupe.ScaleHeight) - Abs(HautBloc)) / 2
      End If
   End If
End Sub

Public Sub TraceTableDecoupe()
   'tracé de la table dans l'onglet découpe
   Dim i As Integer
   
   pctDecoupe.AutoRedraw = True
   pctDecoupe.FillStyle = 1  'transparent
   pctDecoupe.DrawMode = vbCopyPen
   pctDecoupe.DrawWidth = 1
   pctDecoupe.Cls
   
   '**** Tracé des rectangles ****
   For i = 1 To NbRect
      If RECT(i).Rempli = True Then
         pctDecoupe.FillStyle = 0 'plein
         pctDecoupe.FillColor = RECT(i).CoulFond
      Else
         pctDecoupe.FillStyle = 1 'transparent
      End If
      With RECT(i)
         pctDecoupe.Line (.X1, .Y1)-(.X2, .Y2), .CoulTour, B
      End With
   Next i
   '**** Tracé des lignes ****
'   For i = 1 To NbLignes
'      With Ligne(i)
'         pctDecoupe.Line (.x1, .y1)-(.x2, .y2), .Couleur
'      End With
'   Next i
End Sub

Public Sub TraceBlocEtOrigine()
   
   '****************************** WARNING - ATTENTION ***********************************
   '*** l'origine du repère des coordonnées se trouve au niveau du plateau de la table ***
   '*** MAIS la position de repos du fil est en haut et se trouve donc en (0, CourseY) ***
   '**************************************************************************************
   
      'position de repos du fil
      pctDecoupe.FillStyle = 0 'plein
      pctDecoupe.FillColor = vbBlack
      pctDecoupe.Circle (0, CourseY), 4 * pctDecoupe.ScaleWidth / pctDecoupe.Width, vbBlack '(on veut un cercle de 4 pixels alors qu'on est en mm)
      
      'bloc de polystyrène
      With RectBloc
         .X1 = MargeFil
         .Y1 = 0
         .X2 = MargeFil + LongBloc
         .Y2 = HautBloc
         .CoulFond = 16045215
         .CoulTour = 8157820     'même tour que le plateau pour ligne du bas confondue
         .Rempli = True
         .TypeTrait = 0
      End With
      If RectBloc.Rempli = True Then
         pctDecoupe.FillStyle = 0 'plein
         pctDecoupe.FillColor = RectBloc.CoulFond
      Else
         pctDecoupe.FillStyle = 1 'transparent
      End If
      With RectBloc
         pctDecoupe.Line (.X1, .Y1)-(.X2, .Y2), .CoulTour, B
      End With
      'zone utile
      pctDecoupe.FillStyle = 1 'transparent
      pctDecoupe.DrawStyle = vbDot
      pctDecoupe.Line (MargeFil, MargePlateau)-(CourseX, CourseY - MargeFil), vbBlack, B
      pctDecoupe.DrawStyle = vbSolid
         
      'axes de l'origine
'      pctDecoupe.DrawStyle = vbDashDot
'      pctDecoupe.Line (0, 0)-(0, CourseY)
'      pctDecoupe.Line (0, CourseY)-(CourseX, CourseY)
'      pctDecoupe.DrawStyle = vbSolid
End Sub

Public Sub AssemblageSequencesPourDecoupe()
'conversion des toutes les séquences en une seule séquence pour représentation dans l'onglet découpe
   Dim i As Long, j As Long
   
   SequDecoupe.NbPoints = 0
   For i = 1 To NbTransf
      With Transf(i)
         For j = 1 To .NbPoints
            SequDecoupe.NbPoints = SequDecoupe.NbPoints + 1
            ReDim Preserve SequDecoupe.Point(1 To SequDecoupe.NbPoints)
            SequDecoupe.Point(SequDecoupe.NbPoints) = .Point(j)
         Next j
      End With
   Next i
   'on effectue un nettoyage pour supprimer les points superposés ; l'algo est le même que celui
   'utilisé lors de l'importation des profils ; il utilise le tableau Profil comme tableau de transfert (pas très heureux,
   'mais lié aux indices qui ne commencent pas au même numéro, il faudrait revoir la fonction Nettoyage pour passer le tableau par référence)
   Call MaxiMiniSequ(SequDecoupe)
   With SequDecoupe
      ReDim profil(0 To .NbPoints - 1)
      For i = 1 To .NbPoints
         profil(i - 1).x = .Point(i).x
         profil(i - 1).y = .Point(i).y
      Next i
      If .DeltaX >= .DeltaY Then
         EpsilonNettoyage = .DeltaX / CoeffNett     'CoeffNett défini dans le Form Load
      Else
         EpsilonNettoyage = .DeltaY / CoeffNett
      End If
      NbPoints = Nettoyage(EpsilonNettoyage)     'nettoyage des points trop rapprochés
      .NbPoints = NbPoints
      ReDim .Point(1 To .NbPoints)               'tableau du profil initial, inclus dans le type "Trajet"
      For i = 1 To .NbPoints                      'Transffert des points dans mon type de tableau
         .Point(i).x = profil(i - 1).x
         .Point(i).y = profil(i - 1).y
      Next i
   End With
   Call MaxiMiniSequ(SequDecoupe)
   Erase profil 'libération de mémoire

End Sub

Public Sub CreationTrajetsEntreeSortie()
   
   'Approche jusqu'au bord du bloc :
   If optEntrerBloc(0).Value = True Then ' entrée par la gauche
      With SequEntree
         .NbPoints = 3
         ReDim SequEntree.Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = 0
         .Point(2).y = SequDecoupe.Point(1).y
         .Point(3).x = SequDecoupe.Point(1).x
         .Point(3).y = SequDecoupe.Point(1).y
      End With
   ElseIf optEntrerBloc(1).Value = True Then    'entrée par le haut
      With SequEntree
         .NbPoints = 3
         ReDim SequEntree.Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = SequDecoupe.Point(1).x
         .Point(2).y = CourseY
         .Point(3).x = SequDecoupe.Point(1).x
         .Point(3).y = SequDecoupe.Point(1).y
      End With
   ElseIf optEntrerBloc(2).Value = True Then    'entrée par la droite
      With SequEntree
         .NbPoints = 4
         ReDim SequEntree.Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = MargeFil + LongBloc + MargeFil
         If .Point(2).x > CourseX Then .Point(2).x = CourseX  'gestion du dépassement pour grand bloc
         .Point(2).y = .Point(1).y
         .Point(3).x = .Point(2).x
         .Point(3).y = SequDecoupe.Point(1).y
         .Point(4).x = SequDecoupe.Point(1).x
         .Point(4).y = SequDecoupe.Point(1).y
      End With
  End If

   'définition de la sortie : on traite d'abord la partie hors matière, pour disposer du point de sortie de la matière
   If optSortirBloc(0).Value = True Then  'sortie par la gauche
      With SequSortie
         .NbPoints = 3
         ReDim SequSortie.Point(1 To .NbPoints)
         .Point(1).x = SequDecoupe.Point(SequDecoupe.NbPoints).x
         .Point(1).y = SequDecoupe.Point(SequDecoupe.NbPoints).y
         .Point(2).x = 0
         .Point(2).y = .Point(1).y
         .Point(3).x = 0
         .Point(3).y = CourseY
      End With
   ElseIf optSortirBloc(1).Value = True Then    'sortie par le haut
      With SequSortie
         .NbPoints = 3
         ReDim SequSortie.Point(1 To .NbPoints)
         .Point(1).x = SequDecoupe.Point(SequDecoupe.NbPoints).x
         .Point(1).y = SequDecoupe.Point(SequDecoupe.NbPoints).y
         .Point(2).x = .Point(1).x
         .Point(2).y = CourseY
         .Point(3).x = 0
         .Point(3).y = CourseY
      End With
   ElseIf optSortirBloc(2).Value = True Then    'sortie par la droite
      With SequSortie
         .NbPoints = 4
         ReDim SequSortie.Point(1 To .NbPoints)
         .Point(1).x = SequDecoupe.Point(SequDecoupe.NbPoints).x
         .Point(1).y = SequDecoupe.Point(SequDecoupe.NbPoints).y
         .Point(2).x = MargeFil + LongBloc + MargeFil
         If .Point(2).x > CourseX Then .Point(2).x = CourseX  'gestion du dépassement pour grand bloc
         .Point(2).y = .Point(1).y
         .Point(3).x = .Point(2).x
         .Point(3).y = CourseY
         .Point(4).x = 0
         .Point(4).y = CourseY
      End With
   End If
End Sub

Public Sub TraceDecoupe()
   Dim i As Integer
   Dim Pixel As Single

   Pixel = pctDecoupe.ScaleWidth / pctDecoupe.Width
   If NbTransf > 0 Then    'il faut que des séqences soient transférées
      If flagTracePourSimulation = False Then
         pctDecoupe.AutoRedraw = False    'on n'est pas en simulation, on trace sur le plan provisoire
         pctDecoupe.Cls
      Else
         pctDecoupe.AutoRedraw = True    'on est en simulation, on trace sur le plan permanent
      End If
      pctDecoupe.FillStyle = 1 'transparent
      'tracé de l'entrée ; on décale de 1 pixel vers la droite et le haut pour éviter la superposition des traits
      With SequEntree
         pctDecoupe.DrawWidth = 1
         pctDecoupe.Line (.Point(1).x + Pixel, .Point(1).y + Pixel)-(.Point(2).x + Pixel, .Point(2).y + Pixel), vbRed
         For i = 3 To .NbPoints
            pctDecoupe.Line (.Point(i - 1).x + Pixel, .Point(i - 1).y + Pixel)-(.Point(i).x + Pixel, .Point(i).y + Pixel), vbRed
         Next i
      End With
      'tracé de la séquence
      pctDecoupe.DrawStyle = vbSolid
      pctDecoupe.DrawWidth = 1
      With SequDecoupe
         For i = 2 To .NbPoints
            pctDecoupe.Line (.Point(i - 1).x, .Point(i - 1).y)-(.Point(i).x, .Point(i).y), vbBlue
         Next i
      End With
      'tracé de la sortie ; on décale de 1 pixel vers la gauche et le bas pour éviter la superposition des traits
      pctDecoupe.DrawStyle = vbSolid
      With SequSortie
         For i = 2 To .NbPoints - 1
            pctDecoupe.Line (.Point(i - 1).x - Pixel, .Point(i - 1).y - Pixel)-(.Point(i).x - Pixel, .Point(i).y - Pixel), RGB(0, 144, 61)
         Next i
         pctDecoupe.DrawWidth = 1
         pctDecoupe.Line (.Point(.NbPoints - 1).x - Pixel, .Point(.NbPoints - 1).y - Pixel)-(.Point(.NbPoints).x - Pixel, .Point(.NbPoints).y - Pixel), RGB(0, 144, 61)
      End With
      'tracé du point d'entrée
      pctDecoupe.FillStyle = 0 'rempli
      pctDecoupe.FillColor = vbRed
      With SequEntree
         pctDecoupe.Circle (.Point(.NbPoints).x, .Point(.NbPoints).y), 3 * Pixel, vbRed
      End With
      pctDecoupe.FillStyle = 1 'transparent
     'tracé du point de sortie
      With SequSortie
         pctDecoupe.Circle (.Point(1).x, .Point(1).y), 4 * Pixel, RGB(0, 144, 61)
      End With
      pctDecoupe.DrawWidth = 1
      'Tracé du décalage
      If pctValidationDecoupe.Visible = True And ((ModeSoft = "Normal" And optDecalage(2).Value = False) Or _
                                                   (ModeSoft = "Expert" And optDecalageExpert(0).Value = False)) Then 'il faut afficher le décalage
         pctDecoupe.DrawWidth = 2
         With SequDecalee
            For i = 2 To .NbPoints
               pctDecoupe.Line (.Point(i - 1).x, .Point(i - 1).y)-(.Point(i).x, .Point(i).y), RGB(240, 136, 0)
            Next i
         End With
         pctDecoupe.DrawWidth = 1
      End If
      pctDecoupe.AutoRedraw = True
   End If
End Sub

Private Sub chkZoomDecoupe_Click()
   Call InitialisationDessinDecoupe
End Sub

'*****************************************************************************************************
'******************** Définition des trajectoires d'entrée et de sortie du fil ***********************
'*****************************************************************************************************
Private Sub optEntrerBloc_Click(Index As Integer)
   If SSTab1.Tab = 1 And NbTransf > 0 Then
      Call CreationTrajetsEntreeSortie 'en fonction de la position des boutons
      Call TraceDecoupe
   End If
End Sub

Private Sub optSortirBloc_Click(Index As Integer)
   If SSTab1.Tab = 1 And NbTransf > 0 Then
      Call CreationTrajetsEntreeSortie 'en fonction de la position des boutons
      Call TraceDecoupe
   End If
End Sub

'*****************************************************
'*************** Gestion des matériaux ***************
'*****************************************************
Private Sub cmdGestionMatiere_Click(Index As Integer)
'sauvegarde de la matière affichée, dans le .ini
   Dim AreYouSure As Integer
   Dim NomTemporaire As String 'nom de la matière pour comparaison
   Dim i As Integer
   
   Select Case Index
   Case 0  'sauvegarde d'une nouvelle matière
      NomTemporaire = InputBox("Nom de la matière :", "Base de données Matières")
      NomTemporaire = Trim(NomTemporaire)  'on enlève les espaces avant et après pour éviter les bugs du .ini
      'il faut avant tout contrôler que le nom n'existe pas déjà
      'si on utilise un nom déjà utilisé, il faut valider l'écrasement des données
      If NomTemporaire <> "" Then
         For i = 0 To UBound(MatieresDeLaBase)
            If NomTemporaire = MatieresDeLaBase(i).Nom Then
               If ModeSoft = "Expert" Or (ModeSoft = "Normal" And MatieresDeLaBase(i).Vitesse = "4.0") Then
                  AreYouSure = MsgBox(Message(Corps, 18), vbYesNo, Message(Titre, 18))  'nom existe, écraser valeurs?
               Else
                  AreYouSure = MsgBox(Message(Corps, 57), vbYesNo, Message(Titre, 57))  'nom existe en mémoire (Expert mode), écraser valeurs?
               End If
               If AreYouSure = vbNo Then  'si la réponse est non
                  Exit Sub
               Else
                  Call SauverMatiere(MatieresDeLaBase(i).Nom, ChauffeCourante, VitesseDecoupe)
                  comboMatieres.Text = comboMatieres.List(i) 'on affiche la matière sauvée
                  Unload Me  'on sort de la sub et on ferme la fenêtre
               End If
            End If
         Next i
         'si on n'est pas sorti de la sub, il faut sauver la matière
         Call SauverMatiere(NomTemporaire, ChauffeCourante, VitesseDecoupe) 'seulement dans le .ini
         Call ListerMatieresDuIni  'lit le .ini et remplit la combobox
         comboMatieres.Text = comboMatieres.List(comboMatieres.ListCount - 1) 'on affiche la dernière matière
         '********** A OPTIMISER : PAS BON SI REMPLACEMENT ***********
         'les opérations sur comboMatieres vont automatiquement lancer comboMatieres_click() qui va actualiser les valeurs des curseurs
      End If
   Case 1   'remplacement de la chauffe dans la matière courante
      If MatiereUtilisee.Nom = "Reglage par defaut" And VitesseDecoupe <> 4 Then  'si on est sur la matière par défaut, il ne faut pas changer la vitesse
         MsgBox Message(Corps, 59), vbOKOnly, Message(Titre, 59) 'ligne ne peut pas être modifiée
         Exit Sub
      End If
      Call SauverMatiere(MatiereUtilisee.Nom, ChauffeCourante, VitesseDecoupe) 'en .ini
      Call ListerMatieresDuIni  'lit le .ini et remplit la combobox
   Case 2   'suppression de la matière courante
      If MatiereUtilisee.Nom = "Reglage par defaut" Then
         MsgBox Message(Corps, 19), vbOKOnly, Message(Titre, 19) 'ligne ne peut pas être effacée
         Exit Sub
      Else
         AreYouSure = MsgBox(Message(Corps, 20), vbYesNo, Message(Titre, 20)) 'valider l'effacement ?
         If AreYouSure = vbNo Then  'si la réponse est non
            Exit Sub
         Else
            EffacerSectionIni "Matiere_" & MatiereUtilisee.Nom 'efface toutes les clés
            EffacerLigneDansIni "[Matiere_" & MatiereUtilisee.Nom & "]" 'efface la section
            MatiereUtilisee.Nom = MatieresDeLaBase(0).Nom 'on repasse sur la matière par défaut
            Call ListerMatieresDuIni  'lit le .ini et remplit la combobox
         End If
      End If
   End Select
End Sub

'choix de matière dans la liste déroulante
Private Sub comboMatieres_click()
   If MatieresAffichees(comboMatieres.ListIndex).Chauffe >= hscChauffeDecoupe.Min And MatieresAffichees(comboMatieres.ListIndex).Chauffe <= hscChauffeDecoupe.Max Then
      hscChauffeDecoupe.Value = MatieresAffichees(comboMatieres.ListIndex).Chauffe  'il y a correspondance entre les index du Combobox et du tableau MatieresAffichees
      ChauffeCourante = hscChauffeDecoupe.Value 'par sécurité
   Else
      MsgBox Message(Corps, 21), vbCritical, Message(Titre, 21) 'valeur de chauffe hors limites
   End If
   If MyVal(MatieresAffichees(comboMatieres.ListIndex).Vitesse) * 10 >= hscVitesseBDD.Min And MyVal(MatieresAffichees(comboMatieres.ListIndex).Vitesse) * 10 <= hscVitesseBDD.Max Then
      hscVitesseBDD.Value = MyVal(MatieresAffichees(comboMatieres.ListIndex).Vitesse) * 10 'il y a correspondance entre les index du Combobox et du tableau MatieresAffichees
      hscVitesseManuel.Value = hscVitesseBDD.Value
      VitesseDecoupe = hscVitesseBDD.Value / 10 'par sécurité
   Else
      MsgBox Message(Corps, 58), vbCritical, Message(Titre, 58) 'valeur de vitesse hors limites
   End If
   MatiereUtilisee.Nom = MatieresAffichees(comboMatieres.ListIndex).Nom
   EcritFichierIni "BDDMatieres", "MatiereUtilisee", MatiereUtilisee.Nom
End Sub

'gestion des scrollbar de réglage
Private Sub hscChauffeDecoupe_Change()
   If pctDecoupe.Visible = True Then
      hscChauffeStop.Value = hscChauffeDecoupe.Value
      hscChauffeFilManuel.Value = hscChauffeDecoupe.Value
      hscChauffePendantDecoupe.Value = hscChauffeDecoupe.Value
      ChauffeCourante = hscChauffeDecoupe.Value
      lblChauffeDecoupe.Caption = Format(ChauffeCourante, "##0") & " %"
   End If
End Sub

Private Sub hscChauffeDecoupe_Scroll()
   lblChauffeDecoupe.Caption = Format(hscChauffeDecoupe.Value, "##0") & " %"
End Sub

'gestion des scrollbar de réglage
Private Sub hscVitesseBDD_Change()
   hscVitesseManuel.Value = hscVitesseBDD.Value
   VitesseDecoupe = hscVitesseBDD.Value / 10
   lblVitesseDecoupe.Caption = Format(VitesseDecoupe, "0.0") & " mm/s"
End Sub

Private Sub hscVitesseBDD_Scroll()
   lblVitesseDecoupe.Caption = Format(hscVitesseBDD.Value / 10, "0.0") & " mm/s"
End Sub

Private Sub cmdSTOP_Click()
   flagAppuiSTOP = True
End Sub

Private Sub cmdRetourOrigine_Click()
   Dim ReponseFonction As Integer
   
   If IPL5X_IsConnected() <> 1 Then 'l'interpolateur n'est pas connecté
      flagTableEcriteDansIPL = False
      frmDecoupeInactive.Show vbModal
      Exit Sub
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
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then Exit Sub   'la machine n'est pas prête
   lblProcedure.Caption = Label(11) '  "Retour automatique à la position de repos"
   'on utilise la picturebox des déplacements manuels en masquant certains contrôles
   pctValidationDecoupe.Visible = False
   pctRepriseDecoupe.Visible = False
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   frameFilManuel.Visible = False
   frameChauffeFil.Visible = False
   pctFil.Visible = True
   If MsgBox(Message(Corps, 23), vbYesNo, Message(Titre, 23)) = vbYes Then 'validation retour repos
      flagAppuiSTOP = False
      lblAvertissementFil.Caption = Label(12)  '  "Recherche des interrupteurs"
      lblAvertissementFil.Visible = True
      flagPositionPliage = False
      ReponseFonction = 1
      ReponseFonction = RetourOrigine("XetY", VMaxAvecAcc) 'les erreurs sont traités dans RetourOrigine, qui modifie PasParcourusXX
      If ReponseFonction = 1 Then   'on est à l'origine, on va à la position de repos du fil
         lblAvertissementFil.Caption = Label(13) ' "Décalage vers la position de repos"
         Call MouvementUniquePas(NbrPasToOriXG, -NbrPasToOriYG, NbrPasToOriXD, -NbrPasToOriYD, VitesseDecoupe, NBRt)
         If flagAppuiSTOP = True Then
            flagAppuiSTOP = False
            GoTo Arret_Stop
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         lblAvertissementFil.Caption = Label(14)  ' "Position de repos"
         MsgBox Message(Corps, 24), vbInformation, Message(Titre, 24) 'fil en position de repos
         pctFil.Visible = False
         Exit Sub
      Else
         pctFil.Visible = False
         Exit Sub
      End If
   Else
      pctFil.Visible = False
      Exit Sub
   End If
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   Exit Sub
Arret_Stop:
   MsgBox Message(Corps, 25), vbCritical, Message(Titre, 25)  'opération annulée, arrêt d'urgence
   pctFil.Visible = False
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   pctFil.Visible = False
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   pctFil.Visible = False
   Exit Sub
End Sub

'***** Procédure de pliage : idem retour à l'origine, mais avec un déplacement plus grand en X ********
Private Sub cmdPlierLePortique_Click()
   Dim ReponseFonction As Integer
   Dim ReponseBox As Integer
   
   If IPL5X_IsConnected() <> 1 Then 'l'interpolateur n'est pas connecté
      flagTableEcriteDansIPL = False
      frmDecoupeInactive.Show vbModal
      Exit Sub
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
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then Exit Sub   'la machine n'est pas prête
   lblProcedure.Caption = Label(15) ' "Préparation pour rangement"
   pctValidationDecoupe.Visible = False
   pctRepriseDecoupe.Visible = False
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   frameFilManuel.Visible = False
   frameChauffeFil.Visible = False
   pctFil.Visible = True
   
   ReponseBox = vbNo
   If flagPliageApresMsgBox = False Then  'le pliage n'intervient pas après la box de demande mais par appui utilisateur sur le bouton
      ReponseBox = MsgBox(Message(Corps, 28), vbYesNo, Message(Titre, 28)) 'validation rangement
   End If
   If ReponseBox = vbYes Or flagPliageApresMsgBox = True Then
      flagPliageApresMsgBox = False
      flagAppuiSTOP = False
      lblAvertissementFil.Caption = Label(12) ' "Recherche des interrupteurs"
      lblAvertissementFil.Visible = True
      ReponseFonction = 1
      ReponseFonction = RetourOrigine("XetY", VMaxAvecAcc)  'les erreurs sont traités dans RetourOrigine, qui modifie PasParcourusXX
      If ReponseFonction = 1 Then   'on est à l'origine, on va à la position de repos du fil
         lblAvertissementFil.Caption = Label(16) ' "Déplacement vers la position de pliage"
         'on dissocie les mouvements en X et Y pour éviter les bruits bizarres
         Call MouvementUniquePas(0, -NbrPasToOriYG, 0, -NbrPasToOriYD, VMaxAvecAcc, NBRt)
         If flagAppuiSTOP = True Then
            flagAppuiSTOP = False
            GoTo Arret_Stop
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         Call MouvementUniquePas(150 * PasParTourXG / MmParTourXG, 0, 150 * PasParTourXD / MmParTourXD, 0, VMaxAvecAcc, NBRt)
         If flagAppuiSTOP = True Then
            flagAppuiSTOP = False
            GoTo Arret_Stop
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         lblAvertissementFil.Caption = Label(17) ' "Position de repos"
         MsgBox Message(Corps, 29), vbInformation, Message(Titre, 29)  'position pliage atteinte
         flagPositionPliage = True
         pctFil.Visible = False
         Exit Sub
      Else
         pctFil.Visible = False
         Exit Sub
      End If
   Else
      pctFil.Visible = False
      Exit Sub
   End If
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   Exit Sub
Arret_Stop:
   MsgBox Message(Corps, 25), vbCritical, Message(Titre, 25)  'opération annulée, arrêt d'urgence
   pctFil.Visible = False
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   pctFil.Visible = False
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   pctFil.Visible = False
   Exit Sub
End Sub

'---------------------------------------
'Annulation dans la fenêtre Manuel / Fil
'---------------------------------------
Private Sub cmdAnnulerFil_Click()
   If frameChauffeFil.Visible = True Then  'si le cadre de chauffe du fil est visible, on est en déplacements manuels
      If optGoManuel(0).Value = True Then 'si on bouge en manuel, on stoppe tout
         optGoManuel(1).Value = True
         If optChauffe(0).Value = True Then optChauffe(1).Value = True  'arrêt de la chauffe
      Else 'sinon on contrôle juste la chauffe
         If optChauffe(0).Value = True Then optChauffe(1).Value = True  'arrêt de la chauffe
      End If
      optManuel(6).Value = True 'on force le bouton vers le bas
      frameFilManuel.Visible = False
      frameChauffeFil.Visible = False
      pctFil.Visible = False
   Else 'si le cadre de chauffe est invisible, on est en procédure de retour à l'origine.
      flagAppuiSTOP = True
   End If
End Sub

Private Function VerifierDegagementInters() As Integer
   'On vérifie que les inters sont bien actifs sur la table courante
   Call EnvoiBytes(&H4F, &H0) 'do not override
   If (ByteIPL(2) And &H2) = &H0 Then   'la table courante n'a pas les switch actifs
      MsgBox Message(Corps, 30), vbOKOnly, Message(Titre, 30) 'interrupteurs inactifs, procédure impossible
      VerifierDegagementInters = -1
      Exit Function
   End If
   'on vérifie qu'on est bien dégagé
   If EtatInterFDC() = True Then 'true=ouvert
      MsgBox Message(Corps, 31), vbOKOnly, Message(Titre, 31) 'un inter fdc ouvert, à dégager à la main
      VerifierDegagementInters = -2
      Exit Function
   End If
   If EtatInterOrigine() = True Then 'true=ouvert
      MsgBox Message(Corps, 32), vbOKOnly, Message(Titre, 32) 'interrupteur origine ouvert, impossible faire procéedure
      VerifierDegagementInters = -3
      Exit Function
   End If
   VerifierDegagementInters = 1 'OK
End Function

Private Function RetourOrigine(Axe As String, VitesseMaxDemandee As Single) As Integer
'************* RETOUR A L'ORIGINE avec Mesure des décalages entre les inters d'origine et la position courante du fil ********************
   Dim PasXG As Single, PasYG As Single, PasXD As Single, PasYD As Single
   Dim Petit_MouvementPasXG As Long, Petit_MouvementPasYG As Long, Petit_MouvementPasXD As Long, Petit_MouvementPasYD As Long
   Dim Loop_count As Long
   Dim Compteur_Pas As Long
   
   NBRt = 0
   NBRc = 0
   RetourOrigine = 0 'initialisation du retour de la fonction ; passera à 1 si tout s'est bien déroulé ou à d'autres valeurs suivant les erreurs éventuelles
    
   'Distances utilisées par la procédure de mise à l'origine : on doit mettre plus que la longueur de la machine
   PasXG = -Round((CourseX + 200) * PasParTourXG / MmParTourXG, 0)
   PasYG = Round((CourseY + 200) * PasParTourYG / MmParTourYD, 0)
   PasXD = -Round((CourseX + 200) * PasParTourXD / MmParTourXD, 0)
   PasYD = Round((CourseY + 200) * PasParTourYD / MmParTourYD, 0)
   
   PasParcourusXG = 0
   PasParcourusYG = 0
   PasParcourusXD = 0
   PasParcourusYD = 0
   Select Case TypeMachine
   Case "MaxiCut2d"
      Petit_MouvementPasXG = Round((0.2) * PasParTourXG / MmParTourXG, 0)      'on sort des inters par segments de 0.2mm
      Petit_MouvementPasYG = Round((0.2) * PasParTourYG / MmParTourYG, 0)      'on sort des inters par segments de 0.2mm
      Petit_MouvementPasXD = Round((0.2) * PasParTourXD / MmParTourXD, 0)      'on sort des inters par segments de 0.2mm
      Petit_MouvementPasYD = Round((0.2) * PasParTourYD / MmParTourYD, 0)      'on sort des inters par segments de 0.2mm
   Case Else
      Petit_MouvementPasXG = Round((0.1) * PasParTourXG / MmParTourXG, 0)      'on sort des inters par segments de 0.1mm
      Petit_MouvementPasYG = Round((0.1) * PasParTourYG / MmParTourYG, 0)      'on sort des inters par segments de 0.1mm
      Petit_MouvementPasXD = Round((0.1) * PasParTourXD / MmParTourXD, 0)      'on sort des inters par segments de 0.1mm
      Petit_MouvementPasYD = Round((0.1) * PasParTourYD / MmParTourYD, 0)      'on sort des inters par segments de 0.1mm
   End Select
   Select Case Axe  'on annule le mouvement sur un des axes
   Case "Y"  'home sur Y, on annule les valeurs sur X
      PasXG = 0
      PasXD = 0
   Case "X"  'home sur X, on annule les valeurs sur Y
      PasYG = 0
      PasYD = 0
   Case Else
      'si pas seulement X ou Y, on laisse les valeurs : remontée en diagonale (vitesse à contrôler)
   End Select
   Loop_count = 0
   ' *** début de l'enchaînement des déplacements ***
   Do
      Do
         Call MouvementUniquePas(PasXG, PasYG, PasXD, PasYD, VitesseMaxDemandee, NBRt)   'pour le permier, tous les axes bougent,
         If flagAppuiSTOP = True Then                                              'mais ensuite PasXX de l'axe qui a
            flagAppuiSTOP = False                                                'touché le premier un inter passe à 0.
            GoTo Arret_Stop
         End If
         If flagAppuiStopSansMsgBox = True Then
            flagAppuiStopSansMsgBox = False
            GoTo Arret_Stop_Sans_MsgBox
         End If
         If (ByteIPL(12) And &H10) = &H10 Then     'stop par inter origine
            If optHomeX.Value = True Or optHomeY.Value = True Then 'si on est en procédure manuelle de retour auto
               optAnnulerHome.Enabled = False 'on interdit le stop sans msgbox pendant la sortie des inters.
            End If
            Exit Do  'un inter origine ouvert, c'est ce qu'on veut
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'un FDC ouvert, anormal
      Loop  'On boucle au cas où le mouvement ne suffirait pas
      'un inter est ouvert, on calcule le nombre de pas parcourus
      'les infos de la salve information sont encore disponibles dans ByteIPL() ; NBRt donne le nombre de pulses qui étaient prévues
      NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
      If NBRt > 0 Then
         PasParcourusXG = PasParcourusXG + Abs(PasXG) * (NBRt - NBRc - 1) / NBRt
         PasParcourusYG = PasParcourusYG + Abs(PasYG) * (NBRt - NBRc - 1) / NBRt
         PasParcourusXD = PasParcourusXD + Abs(PasXD) * (NBRt - NBRc - 1) / NBRt
         PasParcourusYD = PasParcourusYD + Abs(PasYD) * (NBRt - NBRc - 1) / NBRt
      Else
         MsgBox "Erreur, NBRt=0, division impossible => contact@minicut2d.com"
         Exit Function
      End If
      
      'ensuite on enclenche la procédure de remise à l'origine inters
      
      Call EnvoiBytes(&H4F, &H1) 'Override_ON : on a ouvert un inter, il faut en sortir ; le premier inter qui en sort est à l'origine inter
      
      'On ressort par petites avances de 40pas (0.1mm), si on atteint 800pas (2mm) et que les inters ne sont pas fermés, on passe à l'axe suivant
      If PasYG <> 0 And EtatInterOrigine() = True Then      'true -> ouvert (Attention, en VB, True= -1)
         Compteur_Pas = 0
         Do
            Call MouvementUniquePas(0, -Petit_MouvementPasYG, 0, 0, VitesseDecoupe, NBRt)
            If flagAppuiSTOP = True Then
               flagAppuiSTOP = False
               GoTo Arret_Stop
            End If
            If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
            Compteur_Pas = Compteur_Pas + Petit_MouvementPasYG
            PasParcourusYG = PasParcourusYG - Petit_MouvementPasYG
            If (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H0 Then 'mouvement terminé et boucle des inters origine fermée, c'est ce qu'on voulait
               PasYG = 0  'YG est à l'origine, mais pas forcément YD, XG et XD
               Call MouvementUniquePas(0, -Petit_MouvementPasYG, 0, 0, VitesseDecoupe, NBRt) 'on rajoute un mouvement anti-bagotage des inters
               If flagAppuiSTOP = True Then
                  flagAppuiSTOP = False
                  GoTo Arret_Stop
               End If
               If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
               PasParcourusYG = PasParcourusYG - Petit_MouvementPasYG
               Exit Do
            ElseIf (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H2 Then
               'mouvement terminé et inter origine toujours ouvert, il faut boucler
            ElseIf (ByteIPL(8) And &H4) = &H4 Then
               GoTo Arret_FDC 'on contrôle les fdc par acquit de conscience (override_on)
            End If
            If Compteur_Pas > 2 * PasParTourYG / MmParTourYG Then 'Au bout de (2mm), on sort : ce n'est pas cet axe qui a ouvert l'inter
               Exit Do
            End If
         Loop
      End If
      
      'Même chose pour YD :
      If PasYD <> 0 And EtatInterOrigine() = True Then      'true -> ouvert
         Compteur_Pas = 0
         Do
            Call MouvementUniquePas(0, 0, 0, -Petit_MouvementPasYD, VitesseDecoupe, NBRt)
            If flagAppuiSTOP = True Then
               flagAppuiSTOP = False
               GoTo Arret_Stop
            End If
            If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
            Compteur_Pas = Compteur_Pas + Petit_MouvementPasYD
            PasParcourusYD = PasParcourusYD - Petit_MouvementPasYD
            If (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H0 Then 'mouvement terminé et boucle des inters origine fermée, c'est ce qu'on voulait
               PasYD = 0  'YD est à l'origine, mais pas forcément YG et X
               Call MouvementUniquePas(0, 0, 0, -Petit_MouvementPasYD, VitesseDecoupe, NBRt) 'on rajoute un mouvement anti-bagotage des inters
               If flagAppuiSTOP = True Then
                  flagAppuiSTOP = False
                  GoTo Arret_Stop
               End If
               If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
               PasParcourusYD = PasParcourusYD - Petit_MouvementPasYD
               Exit Do
            ElseIf (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H2 Then
               'mouvement terminé et inter origine toujours ouvert, il faut boucler
            ElseIf (ByteIPL(8) And &H4) = &H4 Then
               GoTo Arret_FDC 'on contrôle les fdc par acquit de conscience (override_on)
            End If
            If Compteur_Pas > 2 * PasParTourYD / MmParTourYD Then 'Au bout de (2mm), on sort : ce n'est pas cet axe qui a ouvert l'inter
               Exit Do
            End If
         Loop
      End If
      'Même chose pour XG :
      If PasXG <> 0 And EtatInterOrigine() = True Then      'True -> ouvert
         Compteur_Pas = 0
         Do
            Call MouvementUniquePas(Petit_MouvementPasXG, 0, 0, 0, VitesseDecoupe, NBRt)
            If flagAppuiSTOP = True Then
               flagAppuiSTOP = False
               GoTo Arret_Stop
            End If
            If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
            Compteur_Pas = Compteur_Pas + Petit_MouvementPasXG
            PasParcourusXG = PasParcourusXG - Petit_MouvementPasXG
            If (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H0 Then 'mouvement terminé et boucle des inters origine fermée, c'est ce qu'on voulait
               PasXG = 0  'XG est à l'origine, mais pas forcément YG, XD et YD
               Call MouvementUniquePas(Petit_MouvementPasXG, 0, 0, 0, VitesseDecoupe, NBRt) 'on rajoute un mouvement anti-bagotage des inters
               If flagAppuiSTOP = True Then
                  flagAppuiSTOP = False
                  GoTo Arret_Stop
               End If
               If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
               PasParcourusXG = PasParcourusXG - Petit_MouvementPasXG
               Exit Do
            ElseIf (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H2 Then
               'mouvement terminé et inter origine toujours ouvert, il faut boucler
            ElseIf (ByteIPL(8) And &H4) = &H4 Then
               GoTo Arret_FDC 'on contrôle les fdc par acquit de conscience (override_on)
            End If
            If Compteur_Pas > 2 * PasParTourXG / MmParTourXG Then 'Au bout de (2mm), on sort : ce n'est pas cet axe qui a ouvert l'inter
               Exit Do
            End If
         Loop
      End If
      'POUR LA MAXICUT2D, même chose pour XD :
      If TypeMachine = "MaxiCut2d" Then
         If PasXD <> 0 And EtatInterOrigine() = True Then      'True -> ouvert
            Compteur_Pas = 0
            Do
               Call MouvementUniquePas(0, 0, Petit_MouvementPasXD, 0, VitesseDecoupe, NBRt)
               If flagAppuiSTOP = True Then
                  flagAppuiSTOP = False
                  GoTo Arret_Stop
               End If
               If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
               Compteur_Pas = Compteur_Pas + Petit_MouvementPasXD
               PasParcourusXD = PasParcourusXD - Petit_MouvementPasXD
               If (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H0 Then 'mouvement terminé et boucle des inters origine fermée, c'est ce qu'on voulait
                  PasXD = 0  'XD est à l'origine, mais pas forcément XG, YG et YD
                  Call MouvementUniquePas(0, 0, Petit_MouvementPasXD, 0, VitesseDecoupe, NBRt) 'on rajoute un mouvement anti-bagotage des inters
                  If flagAppuiSTOP = True Then
                     flagAppuiSTOP = False
                     GoTo Arret_Stop
                  End If
                  If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
                  PasParcourusXD = PasParcourusXD - Petit_MouvementPasXD
                  Exit Do
               ElseIf (ByteIPL(12) And &H80) = &H80 And (ByteIPL(8) And &H2) = &H2 Then
                  'mouvement terminé et inter origine toujours ouvert, il faut boucler
               ElseIf (ByteIPL(8) And &H4) = &H4 Then
                  GoTo Arret_FDC 'on contrôle les fdc par acquit de conscience (override_on)
               End If
               If Compteur_Pas > 2 * PasParTourXD / MmParTourXD Then 'Au bout de (2mm), on sort : ce n'est pas cet axe qui a ouvert l'inter
                  Exit Do
               End If
            Loop
         End If
      Else 'SINON pour "MiniCut2d" => SEULEMENT 3 axes
         PasXD = 0  'pour seulement trois axes
      End If
      Call EnvoiBytes(&H4F, &H0) 'do not override : on a mis un axe à l'origine, on va reculer à nouveau ou sortir, il faut donc activer les FDC
      If (PasYG = 0 And PasYD = 0 And PasXG = 0 And PasXD = 0) Then 'si tous les axes sont à l'origine inters
         Exit Do                             'on sort de la procédure de mise à l'origine
      End If
      Loop_count = Loop_count + 1
      Select Case TypeMachine  'on passe dans la boucle une fois 2mm par axe => 3x pour 3axes, 4x pour 4 axes...
      Case "MaxiCut2d"  '4 axes
         If Loop_count = 4 Then
            MsgBox Message(Corps, 33), vbCritical, Message(Titre, 33) 'il faut plus de 2mm pour sortir des inters...on re-essaye
         End If
         If Loop_count > 4 Then
            MsgBox Message(Corps, 34), vbCritical, Message(Titre, 34) 'deuxième essai avorté, contrôlez la machine
         End If
      Case Else  '3 axes
         If Loop_count = 3 Then
            MsgBox Message(Corps, 33), vbCritical, Message(Titre, 33) 'il faut plus de 2mm pour sortir des inters...on re-essaye
         End If
         If Loop_count > 3 Then
            MsgBox Message(Corps, 34), vbCritical, Message(Titre, 34) 'deuxième essai avorté, contrôlez la machine
         End If
      End Select
   Loop
   Call EnvoiBytes(&H4F, &H0) 'on réactive l'override
   optAnnulerHome.Enabled = True 'on réactive le stop manu des procédures auto
   RetourOrigine = 1 'tout s'est bien passé ; la procédure appelante peut lire les PasParcourusXX
   Exit Function
Arret_Stop:
   MsgBox Message(Corps, 25), vbCritical, Message(Titre, 25)  'opération annulée, arrêt d'urgence
   Call EnvoiBytes(&H4F, &H0) 'on réactive l'override
   optAnnulerHome.Enabled = True 'on réactive le stop manu des procédures auto
   RetourOrigine = 2
   Exit Function
Arret_Stop_Sans_MsgBox:
   'uniquement dans le mouvement avant de toucher l'inter
   'pas de message, on se sert du flag pour traiter dans la procédure appelante
   Call EnvoiBytes(&H4F, &H0) 'on réactive l'override
   RetourOrigine = 3
   Exit Function
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   optAnnulerHome.Enabled = True 'on réactive le stop manu des procédures auto
   RetourOrigine = 4
   Exit Function
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   optAnnulerHome.Enabled = True 'on réactive le stop manu des procédures auto
   RetourOrigine = 5
   Exit Function
End Function

Private Sub optManuel_Click(Index As Integer)
   If optGoManuel(0).Value = True Then
      flagChangementDirectionPendantMouvement = True
   End If
End Sub

Private Sub optGoManuel_Click(Index As Integer)
   Dim ManuXG As Single, ManuYG As Single, ManuXD As Single, ManuYD As Single
   Dim SalveManu As SalveData
   
   '*********************** WARNING ***********************
   '**** sur la forme, le conteneur de optGoManuel(x) ne doit pas être le même que celui de optManuel(x), sinon ils intérragissent ****
   '*******************************************************
   
   Select Case Index
   Case 0 'GO
      flagAppuiFeuRouge = False
      flagPositionPliage = False
      optGoManuel(1).ZOrder
      lblProcedure.Caption = Label(30)
      lblProcedure.Visible = True
      lblAvertissementFil2.Caption = Label(18)  '  "LE FIL SE DEPLACE"
      lblAvertissementFil2.BackColor = vbRed
      lblAvertissementFil2.Visible = True
      Do 'On reste dans cette boucle tant qu'on ne stoppe pas le mouvement
         flagChangementDirectionPendantMouvement = False
         'InfiniX et InfiniY sont calculés à l'entrée dans la fenêtre de déplacements manuels
         ' pour éviter d'avoir à les recalculer à chaque appui sur le bouton go.
         If optManuel(0).Value = True Then 'en haut à gauche = HG
            ManuXG = -InfiniX
            ManuYG = InfiniY
            ManuXD = -InfiniX
            ManuYD = InfiniY
         ElseIf optManuel(1).Value = True Then 'H
            ManuXG = 0
            ManuYG = InfiniY
            ManuXD = 0
            ManuYD = InfiniY
         ElseIf optManuel(2).Value = True Then 'HD
            ManuXG = InfiniX
            ManuYG = InfiniY
            ManuXD = InfiniX
            ManuYD = InfiniY
         ElseIf optManuel(3).Value = True Then 'G
            ManuXG = -InfiniX
            ManuYG = 0
            ManuXD = -InfiniX
            ManuYD = 0
         ElseIf optManuel(4).Value = True Then 'D
            ManuXG = InfiniX
            ManuYG = 0
            ManuXD = InfiniX
            ManuYD = 0
         ElseIf optManuel(5).Value = True Then 'BG
            ManuXG = -InfiniX
            ManuYG = -InfiniY
            ManuXD = -InfiniX
            ManuYD = -InfiniY
         ElseIf optManuel(6).Value = True Then 'B
            ManuXG = 0
            ManuYG = -InfiniY
            ManuXD = 0
            ManuYD = -InfiniY
         ElseIf optManuel(7).Value = True Then 'BD
            ManuXG = InfiniX
            ManuYG = -InfiniY
            ManuXD = InfiniX
            ManuYD = -InfiniY
         End If
         'On gère le mouvement en local pour pouvoir stopper, changer de sens, et gérer les erreurs
         'Reset buffer + Data source: USB
         Call EnvoiBytes(&H42, &H0)
         Call EnvoiBytes(&H42, &H1)
         SalveManu = CalculSalveDataPas(VitesseDecoupe, VitesseDecoupe, VitesseDecoupe, ManuXG, ManuYG, ManuXD, ManuYD)
         'Transfert de la salve de mouvement dans le tableau de communication, on récupère le nombre de pulses de la salve
         NBRt = 0
         NBRt = TransfertSalve2ByteIPL(SalveManu)
         'stockage de la demande de mouvement dans le buffer
         ErrIPL = IPL5X_Send(ByteIPL(), 0)
         'c'est un mouvement unique, il faut envoyer l'instruction
         'de fin de découpe dans le buffer (pour pouvoir lire l'état de l'interface je crois)
         Call EnvoiBytes(&H44, &H80)
         'Go buffer : on lance l'exécution des mouvements stockés dans le buffer ce qui met les moteurs on si l'auto on/off a été paramétré
         Call EnvoiBytes(&H42, &H80)
         Do
            DoEvents  ' on regarde si l'utilisateur n'a pas cliqué sur STOP ou changé de sens
            If flagAppuiFeuRouge = True Then 'il y a eu appui sur le feu rouge
               Call EnvoiBytes(&H53, &H0) 'stop rapide
               If optChauffe(0).Value = True Then  'si la chauffe est active, on la remet
                  Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on réactive la chauffe qui a été coupée par le stop
               End If
               lblAvertissementFil2.Visible = False
               optGoManuel(0).ZOrder
               Exit Sub 'on sort de la procédure
            End If
            If flagChangementDirectionPendantMouvement = True Then
               Call EnvoiBytes(&H53, &H0)  'on envoie le stop pour arrêter le mouvement en cours
               If optChauffe(0).Value = True Then  'si la chauffe est active, on la remet
                  Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on réactive la chauffe qui a été coupée par le stop
               End If
            End If
            Call EnvoiBytes(&H49) 'On envoi "Information" en boucle jusqu'à ce que le mouvement soit terminé
            If (ByteIPL(12) And &H2) = &H2 Then   'step activity stopped = on est à l'arrêt
               Exit Do   'on sort de la boucle pour analyser la cause du stop
            End If
         Loop
   
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton STOP sur la machine
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         'si on arrive ici c'est qu'on a changé de direction, on boucle
      Loop
   Case 1  'STOP (feu rouge)
      lblProcedure.Visible = False
      flagAppuiFeuRouge = True
   End Select
   Exit Sub
Arret_Stop:
   If optChauffe(0).Value = True Then optChauffe(1).Value = True 'on relève la chauffe
   If optGoManuel(0).Value = True Then optGoManuel(1).Value = True 'on relève le feu vert
   lblAvertissementFil2.Visible = False
   lblAvertissementFil.Visible = False
   lblProcedure.Visible = False
   lblAvertissementFil.Caption = ""
   lblAvertissementFil.BackColor = &H800000
   optChauffe(0).ZOrder
   optGoManuel(0).ZOrder
   MsgBox Message(Corps, 35), vbCritical, Message(Titre, 35)  'arrêt bouton STOP
   Exit Sub
Arret_Origine:
   If optChauffe(0).Value = True Then optChauffe(1).Value = True 'on relève la chauffe
   If optGoManuel(0).Value = True Then optGoManuel(1).Value = True 'on relève le feu vert
   lblAvertissementFil2.Visible = False
   lblAvertissementFil.Visible = False
   lblAvertissementFil.Caption = ""
   lblAvertissementFil.BackColor = &H800000
   optChauffe(0).ZOrder
   optGoManuel(0).ZOrder
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   Exit Sub
Arret_FDC:
   If optChauffe(0).Value = True Then optChauffe(1).Value = True 'on relève la chauffe
   If optGoManuel(0).Value = True Then optGoManuel(1).Value = True 'on relève le feu vert
   lblAvertissementFil2.Visible = False
   lblAvertissementFil.Visible = False
   lblAvertissementFil.Caption = ""
   lblAvertissementFil.BackColor = &H800000
   optChauffe(0).ZOrder
   optGoManuel(0).ZOrder
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   Exit Sub
End Sub
Private Sub hscChauffeFilManuel_Change()
   If pctDecoupe.Visible = True Then
      hscChauffeStop.Value = hscChauffeFilManuel.Value
      hscChauffeDecoupe.Value = hscChauffeFilManuel.Value
      hscChauffePendantDecoupe.Value = hscChauffeFilManuel.Value
      ChauffeCourante = hscChauffeFilManuel.Value
      lblChauffeManuel.Caption = Format(ChauffeCourante, "##0") & " %"
      progrbarChauffeManu.Max = CalculTempoChauffe(ChauffeCourante)
      If optChauffe(0).Value = True Then
         Do
            Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on ajuste le PWM
            If ByteIPL(2) = &H1 Then Exit Do
         Loop
      End If
   End If
End Sub

Private Sub hscChauffeFilManuel_Scroll()
   lblChauffeManuel.Caption = Format(hscChauffeFilManuel.Value, "##0") & " %"
End Sub

Private Sub TimerTempoChauffeManu_Timer()
   With progrbarChauffeManu
      If .Value < .Max Then
         Call EnvoiBytes(&H49) 'On envoi "Information"
         If (ByteIPL(7) And &H10) = 0 Then    'le PWM est stoppé
            optChauffe(1).Value = True  'se chargera de tout arrêter
         End If
         .Value = .Value + 1
      Else
         lblAvertissementFil.Caption = Label(19) ' "LE FIL CHAUFFE"
         .Visible = False
         optGoManuel(0).Enabled = True 'on libère les mouvements
         TimerTempoChauffeManu.Enabled = False
      End If
   End With
End Sub
Private Sub optChauffe_Click(Index As Integer)
   Select Case Index
   Case 0 'chauffe ON
      If optGoManuel(0).Value = True Then optGoManuel(1).Value = True 'on stope le mouvement
      If optHomeX.Value = True Or optHomeY.Value = True Then optAnnulerHome.Value = True
      optGoManuel(0).Enabled = False 'on empêche tout mouvement
      TempoChauffeFil = CalculTempoChauffe(ChauffeCourante)
      With progrbarChauffeManu
         .Max = TempoChauffeFil
         .Value = 0
         .Visible = True
      End With
      lblAvertissementFil.Caption = Label(20) '  "MISE EN TEMPERATURE DU FIL"
      lblAvertissementFil.BackColor = vbRed
      lblAvertissementFil.Visible = True
      Do
         Call EnvoiBytes(&H50, &H1, ChauffeCourante * ChauffeMaxi / 100 * 2.55) 'on active le PWM
         If ByteIPL(2) = &H1 Then Exit Do
      Loop
      TimerTempoChauffeManu.Enabled = True
      optChauffe(1).ZOrder
   Case 1 'chauffe OFF
      Do
         Call EnvoiBytes(&H50, &H0) 'on stope le PWM
         If ByteIPL(2) = &H0 Then Exit Do
      Loop
      If TimerTempoChauffeManu.Enabled = True Then TimerTempoChauffeManu.Enabled = False  'on stoppe la progress bar
      If optGoManuel(0).Enabled = False Then optGoManuel(0).Enabled = True 'on libère les mouvements
      progrbarChauffeManu.Visible = False
      lblAvertissementFil.Visible = False
      lblAvertissementFil.Caption = ""
      lblAvertissementFil.BackColor = &H800000
      optChauffe(0).ZOrder
   End Select
End Sub

Private Sub CreerSequMouvement()
   Dim i As Long
   
   'on crée une séquence unique avec l'entrée, la découpe et la sortie
   With SequMouvement
      If (ModeSoft = "Normal" And optDecalage(2).Value = True) Or (ModeSoft = "Expert" And optDecalageExpert(0).Value = True) Then 'pas de décalage
         'ajout de l'entrée
         .NbPoints = SequEntree.NbPoints
         ReDim .Point(1 To .NbPoints)
         For i = 1 To .NbPoints
            .Point(i) = SequEntree.Point(i)
         Next i
         'ajout de la découpe (on ne prend pas le premier point qui est déjà là)
         ReDim Preserve .Point(1 To .NbPoints + SequDecoupe.NbPoints - 1)
         For i = 2 To SequDecoupe.NbPoints
            .Point(.NbPoints + i - 1) = SequDecoupe.Point(i)
         Next i
         .NbPoints = .NbPoints + SequDecoupe.NbPoints - 1
         'ajout de la sortie
         ReDim Preserve .Point(1 To .NbPoints + SequSortie.NbPoints - 1)
         For i = 2 To SequSortie.NbPoints
            .Point(.NbPoints + i - 1) = SequSortie.Point(i)
         Next i
         .NbPoints = .NbPoints + SequSortie.NbPoints - 1
      Else
         'ajout de l'entrée
         .NbPoints = SequEntreeDecalee.NbPoints
         ReDim .Point(1 To .NbPoints)
         For i = 1 To .NbPoints
            .Point(i) = SequEntreeDecalee.Point(i)
         Next i
         'ajout de la découpe (on ne prend pas le premier point qui est déjà là)
         ReDim Preserve .Point(1 To .NbPoints + SequDecalee.NbPoints - 1)
         For i = 2 To SequDecalee.NbPoints
            .Point(.NbPoints + i - 1) = SequDecalee.Point(i)
         Next i
         .NbPoints = .NbPoints + SequDecalee.NbPoints - 1
         'ajout de la sortie
         ReDim Preserve .Point(1 To .NbPoints + SequSortieDecalee.NbPoints - 1)
         For i = 2 To SequSortieDecalee.NbPoints
            .Point(.NbPoints + i - 1) = SequSortieDecalee.Point(i)
         Next i
         .NbPoints = .NbPoints + SequSortieDecalee.NbPoints - 1
      End If
      'on attribue la vitesse de coupe à tous les points
      For i = 1 To .NbPoints
         .Point(i).Vitesse = VitesseDecoupe
         .Point(i).Acceleration = False 'pas d'accélérations
      Next i
   End With
End Sub

Private Sub CalculDureeEtSalve()
   Dim Retour As Single
   Dim Jours As Single, Heures As Single, Minutes As Single, Secondes As Single 'on met tout en single pour éviter les pb dans les calculs

   TempoChauffeFil = CalculTempoChauffe(ChauffeCourante)
   Retour = CalculSalvesDecoupe(SequMouvement, TempoChauffeFil, ChauffeCourante)
   If Retour <> -1 Then 'tout s'est bien passé, la fonction renvoie la durée
      Retour = Retour + TempoChauffeFil
      Jours = Int(Retour / 84600)
      Retour = Retour - 84600 * Jours
      Heures = Int(Retour / 3600)
      Retour = Retour - 3600 * Heures
      Minutes = Int(Retour / 60)
      Secondes = Int(Retour - 60 * Minutes)
      TexteDuree = Label(32)
      If Jours > 0 Then
         TexteDuree = TexteDuree & Str(Jours) & " j. " & _
           Str(Heures) & " h. " & Str(Minutes) & " min. " & Str(Secondes) & " s."
      Else
         If Heures > 0 Then
            TexteDuree = TexteDuree & Str(Heures) & " h. " & _
               Str(Minutes) & " min. " & Str(Secondes) & "s."
         Else
            If Minutes > 0 Then
               TexteDuree = TexteDuree & Str(Minutes) & " min. " & _
                  Str(Secondes) & " s."
            Else
               TexteDuree = TexteDuree & Str(Secondes) & " s."
            End If
         End If
      End If
      If ModeSoft = "Expert" Then
         TexteDuree = TexteDuree & vbCrLf & Label(21) & Str(TempoChauffeFil) & Label(22) & vbCrLf & Label(23) & Str(ChauffeCourante) & " %, " _
         & Label(31) & " " & V2P(Format(VitesseDecoupe, "0.0")) & " mm/s)."
      Else
         TexteDuree = TexteDuree & vbCrLf & Label(21) & Str(TempoChauffeFil) & Label(22) & vbCrLf & Label(23) & Str(ChauffeCourante) & " %)."
      End If
      lblDureeDecoupe.Caption = TexteDuree
   Else
      MsgBox Message(Corps, 36), vbCritical, Message(Titre, 36) 'annulation car pb calcul temps découpe
      flagAnnulationDemandeDecoupe = True
   End If
End Sub

Private Sub cmdAnnulerDecoupe_Click()
   pctValidationDecoupe.Visible = False
   optDecalage(2).Value = True 'on annule le décalage
   optDecalageExpert(0).Value = True
   Erase SequMouvement.Point
   Erase SequDecalee.Point
   Erase SalveDecoupe
   Call TraceDecoupe
End Sub

Private Sub cmdFaireOrigineAvantDecoupe_Click()
   Dim ReponseFonction As Integer
   Dim bActive As Boolean

   If IPL5X_IsConnected() <> 1 Then 'l'interpolateur n'est pas connecté
      flagTableEcriteDansIPL = False
      frmDecoupeInactive.Show vbModal
      Exit Sub
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
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then
      Exit Sub   'la machine n'est pas prête
   End If
   Call GestionBoxDebutDecoupe(1)
   flagPositionPliage = False
   Call LancerDecoupe(False) 'on lance la découpe mais ce n'est pas une reprise
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
End Sub

Private Sub GestionBoxDebutDecoupe(ByVal Situation As Integer)
   Select Case Situation
   Case 0   'initialisation de la box
      cmdAnnulerDecoupe.Visible = True
      cmdSTOP.Visible = False
      frameAction.Visible = True
      frameDecalage.Visible = True
      flagAnnulationDemandeDecoupe = False
      lblAvertissementDecoupe.Caption = ""
      lblAvertissementDecoupe.Visible = False
      lblDureeDecoupe.Caption = ""
      progrbarChauffe.Value = 0
      progrbarChauffe.Visible = False
   Case 1   'début de découpe normal à partir de l'écran principal
      lblAvertissementDecoupe.Visible = False
      cmdAnnulerDecoupe.Visible = False
      cmdSTOP.Visible = True
      frameAction.Visible = False
      frameDecalage.Visible = False
      flagAnnulationDemandeDecoupe = False
   End Select
End Sub

'********** GESTION DE LA DECOUPE *********
Private Sub cmdDecouper_Click()
   Dim Reponse As Integer
   Dim Temp1 As Single
   Dim Temp2 As Single
   
   optDecalage(2).Value = True 'on annule le décalage pour éviter les plantages
   optDecalageExpert(0).Value = True
   If NbTransf = 0 Then
      MsgBox Message(Corps, 37), vbInformation, Message(Titre, 37)  'pas de tracé chargé, impossible
      Exit Sub
   End If
   Temp1 = LitFichierIni("Matiere_" & MatiereUtilisee.Nom, "ChauffeDecoupe")
   Temp2 = MyVal(LitFichierIni("Matiere_" & MatiereUtilisee.Nom, "VitesseDecoupe"))
   If Temp1 <> ChauffeCourante Or Temp2 <> VitesseDecoupe Then 'si la chauffe ou la vitesse a été modifiée sans actualiser la base
      Reponse = MsgBox(Message(Corps, 38), vbExclamation + vbYesNoCancel, Message(Titre, 38)) 'chauffe différente de la matière, remplacer ?
      Select Case Reponse
      Case vbYes
         Call cmdGestionMatiere_Click(1)  'on appuye sur le bouton d'actualisation de la chauffe
      Case vbCancel
         Exit Sub
      End Select
   End If
   
   hscChauffeStop.Value = hscChauffeDecoupe.Value
   Call GestionBoxDebutDecoupe(0)
   Call CreerSequMouvement  'on associe les différents trajets
   Call CalculDureeEtSalve
   'on s'assure que le verrouillage de la variation de chauffe est relevé avant de la désactiver
   If optChauffePendantDecoupe(0).Value = True Then optChauffePendantDecoupe(1).Value = True
   hscChauffePendantDecoupe.Enabled = False
   optChauffePendantDecoupe(0).Enabled = False
   pctValidationDecoupe.Visible = True 'on demande à l'utilisateur de choisir les paramètres de sa découpe et on lui laisse une possibilité d'annuler
   'les salves de la découpe sont calculées par CalculSalvesDecoupes à partir de SequMouvement (renvoie le temps total), valeurchauffe et tempochauffe
   Call TraceDecoupe  'à mettre en dernier car la visibilité de pctValidationdecoupe est un critère (il faut avoir calculé le décalage)
End Sub

Private Sub LancerDecoupe(ByVal Reprise As Boolean)
   Dim i As Long
   Dim ReponseFonction As Integer
 
   'ORIGINE puis DECOUPE
   flagAppuiSTOP = False
   If Reprise = False Then   '
      lblAvertissementDecoupe.Caption = Label(24) ' "Procédure de remise à l'origine"
      lblAvertissementDecoupe.Visible = True
      ReponseFonction = RetourOrigine("XetY", VMaxAvecAcc)  'les erreurs sont traités dans RetourOrigine, qui modifie PasParcourusXX
      If ReponseFonction = 1 Then   'on est à l'origine, on va à la position de repos du fil
         Call MouvementUniquePas(NbrPasToOriXG, -NbrPasToOriYG, NbrPasToOriXD, -NbrPasToOriYD, VitesseDecoupe, NBRt)
         If flagAppuiSTOP = True Then
            flagAppuiSTOP = False
            GoTo Arret_Stop
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
      Else
         'la procédure a échoué : on se remet en situation de demander la découpe
         Call GestionBoxDebutDecoupe(0)
         lblDureeDecoupe.Caption = TexteDuree
         Exit Sub
      End If
      'si on arrive là, on est à la position de repos du fil
   End If
   SegmentCourant = 0
   Do
      flagModifChauffePendantDecoupe = False
      ReponseFonction = Decoupe()
      '****** Quand on sort de la fonction, on est à l'arrêt, il faut analyser le numéro de segment et la cause de l'arrêt ******
      'si on passe ici c'est qu'il y a eu un stop qui n'était pas dû à la fin du buffer
      If flagSTOPAvantGoBuffer = True Then 'il y a eu un clic sur Stop avant le Go Buffer
         flagSTOPAvantGoBuffer = False
         Call EnvoiBytes(&H42, &H0) 'on reset le buffer
         GoTo Arret_Stop_Chauffe
      End If
      If SegmentCourant = 1 Then 'il a eu un clic/appui sur Stop pendant la chauffe du fil
         If flagAppuiSTOP = True Then
            flagAppuiSTOP = False
            Call EnvoiBytes(&H42, &H0) 'on reset le buffer
            GoTo Arret_Stop_Chauffe
         End If
         If (ByteIPL(12) And &H40) = &H40 Then GoTo Arret_Stop_Chauffe  'arrêt par stop bouton
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         Exit Sub  'normalement on n'arrive jamais ici!
      End If
      If SegmentCourant >= 2 Then
         'stop à partir du segment 2 : on est en cours de mouvement ************
         If (ByteIPL(12) And &H40) = &H40 Or (ByteIPL(12) And &H4) = &H4 Then  'stop par BP ou ordre
            If (ByteIPL(12) And &H4) = &H4 Then flagAppuiSTOP = False   'réinitialisation du flag stop si besoin
            ReponseFonction = CalculPositionApresStop()
            'avant de tester le retour du calcul de position, on s'assure que le PWM est coupé :
            Do
               Call EnvoiBytes(&H50, &H0) 'on stope le PWM
               If ByteIPL(2) = &H0 Then Exit Do
            Loop
            'on teste le retour du calcul de position
            If ReponseFonction <> 1 Then
               MsgBox "Erreur dans le calcul de la position du Stop => contact@minicut2d.com"
               Exit Sub
            End If
            pctRepriseDecoupe.Visible = True
            Exit Sub
         End If
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
      End If
      If (ByteIPL(12) And &H80) = &H80 Then  'arrêt par fin du buffer
         'Si on arrive ici, tout s'est bien passé, la chauffe a été coupée automatiquement
         MsgBox Message(Corps, 39), vbInformation, Message(Titre, 39) 'découpe terminée
         Call TraceDecoupe
         Call GestionBoxDebutDecoupe(0)
         Exit Sub
      End If
   Loop
   'Si on arrive ici, tout s'est bien passé, on coupe la chauffe
   Call EnvoiBytes(&H50, &H0)
   If ByteIPL(2) <> 0 Then
      MsgBox Message(Corps, 40), vbCritical, Message(Titre, 40)  'chauffe pas coupée, pb de comm
   Else
      MsgBox Message(Corps, 39), vbInformation, Message(Titre, 39) 'découpe terminée
      Call TraceDecoupe
      Call GestionBoxDebutDecoupe(0)
   End If
   Exit Sub
Arret_Stop_Chauffe:
   'on s'assure que le PWM est coupé :
   Do
      Call EnvoiBytes(&H50, &H0) 'on stope le PWM
      If ByteIPL(2) = &H0 Then Exit Do
   Loop
   MsgBox Message(Corps, 41), vbCritical, Message(Titre, 41) 'STOP pendant la mise en température du fil
   Call GestionBoxDebutDecoupe(0)
   lblDureeDecoupe.Caption = TexteDuree
   Exit Sub
Arret_Stop:
   MsgBox Message(Corps, 42), vbCritical, Message(Titre, 42) 'Arrêt d'urgence par STOP
   Call GestionBoxDebutDecoupe(0)
   lblDureeDecoupe.Caption = TexteDuree
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
End Sub

Private Sub cmdAnnulerReprise_Click()
   'on reset la chauffe
   Do
      Call EnvoiBytes(&H50, &H0) 'on stope le PWM
      If ByteIPL(2) = &H0 Then Exit Do
   Loop
   'on reset le buffer
   Call EnvoiBytes(&H42, &H0)
   MsgBox Message(Corps, 43), vbInformation, Message(Titre, 43)  'Arrêt dans zone utile, pas à position repos
   lblAvertissementDecoupe.Caption = ""
   lblAvertissementDecoupe.Visible = False
   progrbarChauffe.Visible = False
   pctValidationDecoupe.Visible = False
   pctRepriseDecoupe.Visible = False
   
   pctValidationDecoupe.Visible = False
   optDecalage(2).Value = True 'on annule le décalage
   optDecalageExpert(0).Value = True
   Erase SequMouvement.Point
   Erase SequDecalee.Point
   Erase SalveDecoupe
   Call TraceDecoupe

End Sub

Private Sub cmdRepriseDecoupe_Click()
   Dim i As Long
   Dim ReponseFonction As Single
   Dim bActive As Boolean
   
   'on modifie les salves et on repart comme pour un nouveau mouvement
   With SequMouvement
      .Point(SegmentCourant - 1).x = PasArretX / PasParTourXG * MmParTourXG
      .Point(SegmentCourant - 1).y = PasArretY / PasParTourYG * MmParTourYG
      If SegmentCourant > 2 Then
         For i = SegmentCourant - 1 To .NbPoints
            .Point(i - (SegmentCourant - 2)) = .Point(i)
         Next i
         .NbPoints = .NbPoints - (SegmentCourant - 2)
         ReDim Preserve .Point(1 To .NbPoints)
         .Point(1).Acceleration = False
      End If
   End With
   Call CalculDureeEtSalve
   pctRepriseDecoupe.Visible = False
   pctValidationDecoupe.Visible = True
   Call LancerDecoupe(True)   'on relance la découpe en précisant qu'il s'agit d'une reprise (pas de remise à l'origine)
End Sub

Private Function CalculPositionApresStop() As Integer
   Dim SegmentArret As Long
   Dim A As Single, B As Single, C As Single
   
   NBRt = 0
   NBRc = 0
   CalculPositionApresStop = 0  'initialisation du retour de la fonction
   SegmentArret = ConcatOctets(ByteIPL(4), ByteIPL(5), ByteIPL(6))  'si on est ici, SegmentArret est >1
   lblNumSegmentStop.Caption = Label(25) & Str(SegmentArret - 1)  'arrêt sur le segment xx
   NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
   With SalveDecoupe(SegmentArret - 1)
      NBRt = ConcatOctets(.NBRL, .NBRM, .NBRH, .NBRU)
   End With
   If NBRt > 0 Then
      With SequMouvementPas
         A = (NBRt - NBRc - 1) / NBRt
         B = (.PointPas(SegmentCourant).XPas - .PointPas(SegmentCourant - 1).XPas)
         C = (.PointPas(SegmentCourant).YPas - .PointPas(SegmentCourant - 1).YPas)
         PasArretX = .PointPas(SegmentCourant - 1).XPas + CLng(A * B)
         PasArretY = .PointPas(SegmentCourant - 1).YPas + CLng(A * C)
      End With
      CalculPositionApresStop = 1
   Else
      MsgBox "Erreur, NBRt=0, division impossible => contact@minicut2d.com"
      Exit Function
   End If
End Function

Private Sub cmdStopRetourApresReprise_Click()
   flagAppuiSTOP = True
   optDecalage(2).Value = True 'on annule le décalage
   optDecalageExpert(0).Value = True
   Do
      Call EnvoiBytes(&H50, &H0) 'on stope le PWM
      If ByteIPL(2) = &H0 Then Exit Do
   Loop
End Sub

Private Sub cmdLancerRetourApresStop_Click()
   Dim Vitesse As Single
   Dim PasRetourX As Long, PasRetourY As Long
   Dim i As Long
   Dim Salve() As SalveData
   Dim Segment As Long
      
   Erase SequMouvement.Point
   Erase SequDecalee.Point
   Erase SalveDecoupe
   lblArretParBoutonStop.Caption = ""
   frameAnnulationReprise.Visible = False
   cmdLancerRetourApresStop.Visible = False
   cmdStopRetourApresReprise.Visible = True
   PasRetourX = -PasArretX
   PasRetourY = Round(CourseY * PasParTourYG / MmParTourYG, 0) - PasArretY

   TempoChauffeFil = CalculTempoChauffe(ChauffeCourante)
   Vitesse = VitesseDecoupe
   If optTrajetRetour(0).Value = True Then   'en diagonale
      ReDim Salve(0 To 2)
      Salve(0) = SalveAttenteChauffe(TempoChauffeFil, ChauffeCourante)
      Salve(1) = CalculSalveDataPas(Vitesse, Vitesse, Vitesse, PasRetourX, PasRetourY, PasRetourX, PasRetourY)
      Salve(2).CMD = &H80  'fin du mouvement
   ElseIf optTrajetRetour(1).Value = True Then  'à gauche
      ReDim Salve(0 To 3)
      Salve(0) = SalveAttenteChauffe(TempoChauffeFil, ChauffeCourante)
      Salve(1) = CalculSalveDataPas(Vitesse, Vitesse, Vitesse, PasRetourX, 0, PasRetourX, 0)
      Salve(2) = CalculSalveDataPas(Vitesse, Vitesse, Vitesse, 0, PasRetourY, 0, PasRetourY)
      Salve(3).CMD = &H80  'fin du mouvement
   ElseIf optTrajetRetour(2).Value = True Then  'en haut
      ReDim Salve(0 To 3)
      Salve(0) = SalveAttenteChauffe(TempoChauffeFil, ChauffeCourante)
      Salve(1) = CalculSalveDataPas(Vitesse, Vitesse, Vitesse, 0, PasRetourY, 0, PasRetourY)
      Salve(2) = CalculSalveDataPas(Vitesse, Vitesse, Vitesse, PasRetourX, 0, PasRetourX, 0)
      Salve(3).CMD = &H80  'fin du mouvement
   End If
   'seulement deux ou trois salves, ça tient dans le buffer
   Call EnvoiBytes(&H42, &H0) 'Reset buffer
   Call EnvoiBytes(&H42, &H1) 'Data source: USB
   For i = 0 To UBound(Salve)
      Call TransfertData2Bytes(Salve(i)) 'on transfère les bytes de la salve Data dans les bytes du tableau de communication
      ErrIPL = IPL5X_Send(ByteIPL(), 0)   'on injecte dans le buffer
   Next i
   Call EnvoiBytes(&H42, &H80)   'go buffer
   lblNumSegmentStop.Caption = Label(28) ' "Mise en température du fil"
   progrbarRetour.Value = 0
   progrbarRetour.Visible = True
   Do
      DoEvents  ' on regarde si l'utilisateur n'a pas cliqué sur STOP
      If flagAppuiSTOP = True Then
         Call EnvoiBytes(&H53)   'on envoie le stop pour arrêter le mouvement en cours
      End If
      Call EnvoiBytes(&H49)  'On envoi "Information" en boucle jusqu'à ce que le mouvement soit terminé (7 segments dans le buffer)
      If (ByteIPL(12) And &H2) = &H2 Then   'step activity stopped = on est à l'arrêt
         If (ByteIPL(12) And &H40) = &H40 Or (ByteIPL(12) And &H4) = &H4 Then
            If (ByteIPL(12) And &H4) = &H4 Then flagAppuiSTOP = False
            GoTo Arret_Stop  'arrêt par bouton ou ordre STOP, on annule
         End If
         If (ByteIPL(12) And &H10) = &H10 Then GoTo Arret_Origine 'arrêt par ouverture des switch d'origine
         If (ByteIPL(12) And &H20) = &H20 Then GoTo Arret_FDC  'arrêt par ouverture des FDC
         If (ByteIPL(12) And &H80) = &H80 Then Exit Do
      End If
      'si le stop n'a pas été demandé, on actualise la chauffe
      Segment = ConcatOctets(ByteIPL(4), ByteIPL(5), ByteIPL(6))
      If Segment = 1 Then
         lblNumSegmentStop.Caption = Label(28) ' "Mise en température du fil"  le premier segment est celui de la tempo de chauffe
         NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
         NBRt = ConcatOctets(Salve(0).NBRL, Salve(0).NBRM, Salve(0).NBRH, Salve(0).NBRU)
         If NBRt > 0 Then
            progrbarRetour.Value = progrbarRetour.Max * (NBRt - NBRc - 1) / NBRt
         Else
            MsgBox "NBRt de la chauffe du fil égal à zéro, division impossible => contact@minicut2d.com", vbCritical, "Opération annulée"
            Exit Sub
         End If
      ElseIf Segment > 1 Then
         progrbarRetour.Visible = False
         lblNumSegmentStop.Caption = Label(26) ' "Retour position de repos"
      End If
   Loop
   'arrivé ici, le mouvement est terminée correctement
   lblNumSegmentStop.Caption = Label(27) ' "Position de repos"
   MsgBox Message(Corps, 44), vbInformation, Message(Titre, 44)  'le fil est en position de repos
   cmdStopRetourApresReprise.Visible = False
   cmdLancerRetourApresStop.Visible = True
   progrbarRetour.Visible = False
   pctRepriseDecoupe.Visible = False
   frameAnnulationReprise.Visible = True
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
Arret_Stop:
   MsgBox Message(Corps, 42), vbCritical, Message(Titre, 42) 'Arrêt d'urgence par STOP
   cmdStopRetourApresReprise.Visible = False
   cmdLancerRetourApresStop.Visible = True
   progrbarRetour.Visible = False
   pctRepriseDecoupe.Visible = False
   frameAnnulationReprise.Visible = True
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
Arret_Origine:
   MsgBox Message(Corps, 26), vbCritical, Message(Titre, 26) 'ouverture boucle origines, dégager à la main
   cmdStopRetourApresReprise.Visible = False
   cmdLancerRetourApresStop.Visible = True
   progrbarRetour.Visible = False
   pctRepriseDecoupe.Visible = False
   frameAnnulationReprise.Visible = True
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
Arret_FDC:
   MsgBox Message(Corps, 27), vbCritical, Message(Titre, 27)  'ouverture boucle FDC, dégager à la main
   cmdStopRetourApresReprise.Visible = False
   cmdLancerRetourApresStop.Visible = True
   progrbarRetour.Visible = False
   pctRepriseDecoupe.Visible = False
   frameAnnulationReprise.Visible = True
   pctValidationDecoupe.Visible = False
   Call TraceDecoupe
   Exit Sub
End Sub

Private Function Decoupe() As Integer
   Dim NumSalve As Long
   Dim j As Integer
   Dim NumSalveDataEnd As Long
   Dim DureeDecoupe As Single
   Dim SalvePourTempo As SalveData
   Dim RetourFonction As Integer
   Dim HauteurRectangle As Single
   Dim A As Single, B As Single, C As Single
   Dim ModifChauffeDemandee As Byte

   NBRt = 0
   NBRc = 0
   Decoupe = 0 'initialisation du retour de la fonction
   SegmentCourant = 0

   'On vérifie que les inters sont bien actifs sur la table courante
   Call EnvoiBytes(&H4F, &H0) 'do not override
   If (ByteIPL(2) And &H2) = &H0 Then   'la table courante n'a pas les switch actifs
      MsgBox Message(Corps, 30), vbOKOnly, Message(Titre, 30) 'interrupteurs inactifs, procédure impossible
      Exit Function
   End If
   'on vérifie qu'on est bien dégagé
   If EtatInterFDC() = True Or EtatInterOrigine = True Then 'true=ouvert
      MsgBox Message(Corps, 45), vbOKOnly, Message(Titre, 45)  'un inter ouvert, découpe impossible
      Exit Function
   End If

   'Le tableau des salves a déjà été rempli dans frmValidationDecoupe
   'Reset buffer
   Call EnvoiBytes(&H42, &H0)
   'Data source: USB
   Call EnvoiBytes(&H42, &H1)
   
   'Tableau des salves :
   ' 0 : tempo de chauffe du fil
   ' 1 : 1er segment du mouvement
   ' 2 : 2ème segment du mouvement
   ' 3 : ...
   ' ...
   ' ubound(SalveDecoupe)-1 : dernier segment du mouvement
   ' ubound(SalveDecoupe) : DATA END  (&H44 , &H80) , soit la fin de la découpe (stop des pas, de la chauffe...)
   ' donc pour 5 segments, on a un tableau qui va de 0 à 6
   ' => SalveDecoupe (0 to 6)
   
   NumSalveDataEnd = UBound(SalveDecoupe)
   NumSalve = 0
   'Une première boucle sert à remplir le buffer (ou à tout envoyer s'il y a peu de vecteurs!)
   Do
      Call TransfertData2Bytes(SalveDecoupe(NumSalve)) 'on transfère les bytes de la salve Data dans les bytes du tableau de communication
      'stockage de la demande de mouvement dans le buffer
      ErrIPL = IPL5X_Send(ByteIPL(), 0)
      If ByteIPL(2) = &H1 Then  'si l'interpolateur répond &H1, la valeur a bien été stockée dans le buffer
         If NumSalve = NumSalveDataEnd Then  'si on a trop peu de salves pour remplir le buffer,
            Exit Do                          ' on sort de la boucle de remplissage
         End If
         NumSalve = NumSalve + 1
      Else           'l'interpolateur ne répond pas &H1 : le buffer est plein, il faut envoyer le mouvement
         Exit Do
      End If
   Loop
   DoEvents  'Le buffer est plein, juste avant le mouvement, on regarde si on n'a pas cliqué sur STOP
             'le buffer n'ayant pas été lancé, il n'y a pas de chauffe et pas de mouvement
   If flagAppuiSTOP = True Then
      flagAppuiSTOP = False
      flagSTOPAvantGoBuffer = True
      Exit Function
   End If
   'Go buffer : on lance l'exécution des mouvements stockés dans le buffer ce qui met les moteurs on si l'auto on/off a été paramétré
   Call EnvoiBytes(&H42, &H80)
   lblAvertissementDecoupe.Caption = Label(28)  ' "Mise en température du fil"
   lblAvertissementDecoupe.Visible = True       'on affiche le label
   progrbarChauffe.Visible = True               'on affiche la progress bar de chauffe
   'Maintenant on envoie la suite en boucle jusqu'à ce que le buffer ait accepté un à un tous les ordres (il acceptera les data au fur et à mesure
   ' de leur exécution) =>polling
   If NumSalve <> NumSalveDataEnd Then 'si on n'a pas tout chargé, il faut faire du polling sur le buffer
      While NumSalve <= NumSalveDataEnd
         DoEvents  ' on regarde si l'utilisateur n'a pas cliqué sur STOP ou modifié la chauffe
         If flagAppuiSTOP = True Then  'STOP demandé
            Call EnvoiBytes(&H53)   'on envoie le stop pour arrêter le mouvement en cours
            Call EnvoiBytes(&H49)   'on demande la salve information pour avoir les infos au même endroit qu'avec le retour de la salve
            Exit Function
         End If
         
         If flagModifChauffePendantDecoupe = True Then  'Modif de la chauffe demandée
            Do
               ModifChauffeDemandee = Round(hscChauffePendantDecoupe.Value * ChauffeMaxi / 100 * 2.55, 0)
               Call EnvoiBytes(&H50, &H1, ModifChauffeDemandee) 'on ajuste le PWM à la chauffe demandée
               If ByteIPL(4) = ModifChauffeDemandee Then Exit Do
            Loop
            flagModifChauffePendantDecoupe = False
         End If
         
         Call TransfertData2Bytes(SalveDecoupe(NumSalve)) 'on transfère les bytes de la salve Data dans les bytes du tableau de communication
         ErrIPL = IPL5X_Send(ByteIPL(), 0)      'c'est un envoi Data : on reçoit Information en retour, même si le buffer refuse la salve
         If ByteIPL(2) = &H1 Then   'data stored in the buffer
            NumSalve = NumSalve + 1
         End If
         If (ByteIPL(12) And &H2) = &H2 Then   'la tentative d'envoi de la salve renvoie step activity stopped = on est à l'arrêt
            Exit Function
         End If
                 
         SegmentCourant = ConcatOctets(ByteIPL(4), ByteIPL(5), ByteIPL(6))
         If SegmentCourant = 1 Then
            lblAvertissementDecoupe.Caption = Label(28) ' "Mise en température du fil"  'le premier segment est celui de la tempo de chauffe
            lblAvertissementDecoupe.Visible = True
            NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
            NBRt = ConcatOctets(SalveDecoupe(0).NBRL, SalveDecoupe(0).NBRM, SalveDecoupe(0).NBRH, SalveDecoupe(0).NBRU)
            If NBRt > 0 Then
               progrbarChauffe.Value = progrbarChauffe.Max * (NBRt - NBRc - 1) / NBRt
            Else
               MsgBox "NBRt de la chauffe du fil égal à zéro, division impossible => contact@minicut2d.com", vbCritical, "Opération annulée"
               Exit Function
            End If
         ElseIf SegmentCourant > 1 Then
            
            If optChauffePendantDecoupe(0).Enabled = False Then  ' au premier passage dans la boucle,
               optChauffePendantDecoupe(0).Enabled = True         ' on libère la possibilité de changer la chauffe pendant la découpe
            End If
            
            progrbarChauffe.Visible = False
            progrbarChauffe.Value = 0
            lblAvertissementDecoupe.Caption = Label(29) & Str(SegmentCourant - 1) & "/" & Str(NumSalveDataEnd - 1) 'découpe segment n°
            NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
            With SalveDecoupe(SegmentCourant - 1)
               NBRt = ConcatOctets(.NBRL, .NBRM, .NBRH, .NBRU)
            End With
            If NBRt > 0 Then
               With SequMouvementPas
                  A = (NBRt - NBRc - 1) / NBRt
                  B = (.PointPas(SegmentCourant).XPas - .PointPas(SegmentCourant - 1).XPas)
                  C = (.PointPas(SegmentCourant).YPas - .PointPas(SegmentCourant - 1).YPas)
                  PasArretX = .PointPas(SegmentCourant - 1).XPas + CLng(A * B)
                  PasArretY = .PointPas(SegmentCourant - 1).YPas + CLng(A * C)
               End With
               pctDecoupe.DrawWidth = 2
               pctDecoupe.AutoRedraw = False
               pctDecoupe.Line (SequMouvement.Point(SegmentCourant - 1).x, SequMouvement.Point(SegmentCourant - 1).y)-(PasArretX / PasParTourXG * MmParTourXG, PasArretY / PasParTourYG * MmParTourYG), vbBlack
            Else
               MsgBox "Erreur, NBRt=0, division impossible => contact@minicut2d.com"
               Exit Function
            End If
         End If
      Wend
   End If
   'Arrivé à ce stade, toutes les Data ont été envoyées, il faut continuer de suivre l'activité du buffer qui se vide
   ' et actualiser la position du fil.
   Do
      DoEvents  ' on regarde si l'utilisateur n'a pas cliqué sur STOP
      If flagAppuiSTOP = True Then
         Call EnvoiBytes(&H53)   'on envoie le stop pour arrêter le mouvement en cours
      End If
      
      If flagModifChauffePendantDecoupe = True Then  'Modif de la chauffe demandée
         Do
            ModifChauffeDemandee = Round(hscChauffePendantDecoupe.Value * ChauffeMaxi / 100 * 2.55, 0)
            Call EnvoiBytes(&H50, &H1, ModifChauffeDemandee) 'on ajuste le PWM à la chauffe demandée
            If ByteIPL(4) = ModifChauffeDemandee Then Exit Do
         Loop
         flagModifChauffePendantDecoupe = False
      End If
      
      Call EnvoiBytes(&H49)  'On envoi "Information" en boucle jusqu'à ce que le mouvement soit terminé (7 segments dans le buffer)
      If (ByteIPL(12) And &H2) = &H2 Then   'step activity stopped = on est à l'arrêt
         SegmentCourant = ConcatOctets(ByteIPL(4), ByteIPL(5), ByteIPL(6))
         Exit Function
      End If
      
      'si le stop n'a pas été demandé, on actualise la position du fil
      SegmentCourant = ConcatOctets(ByteIPL(4), ByteIPL(5), ByteIPL(6))
      If SegmentCourant = 1 Then
         lblAvertissementDecoupe.Caption = "Mise en température du fil"  'le premier segment est celui de la tempo de chauffe
         lblAvertissementDecoupe.Visible = True
         NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
         NBRt = ConcatOctets(SalveDecoupe(0).NBRL, SalveDecoupe(0).NBRM, SalveDecoupe(0).NBRH, SalveDecoupe(0).NBRU)
         If NBRt > 0 Then
            progrbarChauffe.Value = progrbarChauffe.Max * (NBRt - NBRc - 1) / NBRt
         Else
            MsgBox "NBRt de la chauffe du fil égal à zéro, division impossible => contact@minicut2d.com", vbCritical, "Opération annulée"
            Exit Function
         End If
      ElseIf SegmentCourant > 1 Then
                  
         If optChauffePendantDecoupe(0).Enabled = False Then  ' au premier passage dans la boucle,
            optChauffePendantDecoupe(0).Enabled = True         ' on libère la possibilité de changer la chauffe pendant la découpe
         End If

         progrbarChauffe.Visible = False
         lblAvertissementDecoupe.Caption = Label(29) & Str(SegmentCourant - 1) & "/" & Str(NumSalveDataEnd - 1)
         NBRc = ConcatOctets(ByteIPL(13), ByteIPL(14), ByteIPL(15), ByteIPL(16))
         With SalveDecoupe(SegmentCourant - 1)
            NBRt = ConcatOctets(.NBRL, .NBRM, .NBRH, .NBRU)
         End With
         If NBRt > 0 Then
            With SequMouvementPas
               A = (NBRt - NBRc - 1) / NBRt
               B = (.PointPas(SegmentCourant).XPas - .PointPas(SegmentCourant - 1).XPas)
               C = (.PointPas(SegmentCourant).YPas - .PointPas(SegmentCourant - 1).YPas)
               PasArretX = .PointPas(SegmentCourant - 1).XPas + CLng(A * B)
               PasArretY = .PointPas(SegmentCourant - 1).YPas + CLng(A * C)
            End With
            pctDecoupe.DrawWidth = 2
            pctDecoupe.AutoRedraw = False
            pctDecoupe.Line (SequMouvement.Point(SegmentCourant - 1).x, SequMouvement.Point(SegmentCourant - 1).y)-(PasArretX / PasParTourXG * MmParTourXG, PasArretY / PasParTourYG * MmParTourYG), vbBlack
         Else
            MsgBox "Erreur, NBRt=0, division impossible => contact@minicut2d.com"
            Exit Function
         End If
      End If
   Loop
   'arrivé ici, la découpe est terminée correctement
   Exit Function
Y_a_un_pb:
      'On regarde en premier les inters, puis le bouton Stop
      If (ByteIPL(8) And &H2) = &H2 Then  'un switch Origine s'est ouvert pendant la découpe
         MsgBox Message(Corps, 46), vbCritical, Message(Titre, 46) 'arrêt anormal par ouverture inter origine
      End If
      If (ByteIPL(8) And &H4) = &H4 Then  'un switch FDC s'est ouvert pendant la découpe
         MsgBox Message(Corps, 47), vbCritical, Message(Titre, 47) 'arrêt anormal par ouverture inter fin de course
      End If
      If (ByteIPL(12) And &H40) = &H40 Then  'le BP d'arrêt d'urgence a été appuyé durant la découpe
         MsgBox Message(Corps, 48), vbCritical, Message(Titre, 48)  'arrêt par BP, découpe annulée
      End If
      
         'ICI IL FAUT Prévoir une reprise
         'ICI IL FAUT REACTIVER LES BOUTONS
      Exit Function
Erreur:
   MsgBox DecodeErrIPL(ErrIPL), vbInformation, Message(Titre, 40) 'problème de comm avec l'interpolateur
   Exit Function
End Function

'*************** Lancement de la fenêtre des déplacements manuels **********
Private Sub cmdDeplacementsManuels_Click()
   Dim ReponseFonction As Integer
   
   If IPL5X_IsConnected() <> 1 Then 'l'interpolateur n'est pas connecté
      flagTableEcriteDansIPL = False
      frmDecoupeInactive.Show vbModal
      Exit Sub
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
   ReponseFonction = 1
   ReponseFonction = VerifierDegagementInters
   If ReponseFonction <> 1 Then Exit Sub   'la machine n'est pas prête
   InfiniX = Round(6000 * PasParTourXG / MmParTourXG, 0) 'déplacement de 6m
   InfiniY = Round(6000 * PasParTourYG / MmParTourYG, 0)
   frameFilManuel.Visible = True
   frameChauffeFil.Visible = True
   lblProcedure.Visible = False
   pctValidationDecoupe.Visible = False
   pctRepriseDecoupe.Visible = False
   lblAvertissementFil.Visible = False
   lblAvertissementFil2.Visible = False
   pctFil.Visible = True
   Exit Sub
Erreur:
   MsgBox DecodeErrIPL(ErrIPL) & vbCrLf & Message(Corps, 6), vbInformation, Message(Titre, 6) 'l'intialisation de l'interface usb pose pb
   Exit Sub
End Sub

'******** SIMULATION VISUELLE DE LA DECOUPE *********
Private Sub cmdSimulation_Click()
   Dim i As Long
   
   If NbTransf > 0 Then
      If flagSimulationLancee = False Then  'on lance la simulation
         NbClignotements = 4  'initialisation, ira jusqu'à 6 -> clignote 1 fois
         TimerClignotement.Enabled = True
         flagSimulationLancee = True
         cmdSimulation.Picture = frmImages.imgStopSimulation.Picture
         flagTracePourSimulation = True
         pctMasquage.Visible = True
         Call DesactiverOngletDecoupe
         Call TraceDecoupe
         With SequEntree
            ReDim PointsSimulation(1 To .NbPoints)
            For i = 1 To .NbPoints
               PointsSimulation(i).x = .Point(i).x
               PointsSimulation(i).y = .Point(i).y
            Next i
         End With
         With SequDecoupe
            For i = 2 To .NbPoints
               ReDim Preserve PointsSimulation(1 To UBound(PointsSimulation) + 1)
               PointsSimulation(UBound(PointsSimulation)).x = .Point(i).x
               PointsSimulation(UBound(PointsSimulation)).y = .Point(i).y
            Next i
         End With
         With SequSortie
            For i = 2 To .NbPoints
               ReDim Preserve PointsSimulation(1 To UBound(PointsSimulation) + 1)
               PointsSimulation(UBound(PointsSimulation)).x = .Point(i).x
               PointsSimulation(UBound(PointsSimulation)).y = .Point(i).y
            Next i
         End With
         Call FaireLeFilm(PointsSimulation, 2) 'pas
         ProgressionDuFilm = 1 'point de départ du film
         TimerFilm.Interval = 20
         pctDecoupe.AutoRedraw = False
         pctDecoupe.FillColor = RGB(240, 136, 0)
         pctDecoupe.FillStyle = vbSolid
         TimerFilm.Enabled = True
      Else                                'on stoppe la simulation
         lblAvertissementDecoupe.Visible = False
         lblAvertissementDecoupe.Caption = ""
         TimerFilm.Enabled = False
         flagSimulationLancee = False
         cmdSimulation.Picture = frmImages.imgSimuler.Picture
         Call TraceTableDecoupe
         Call TraceBlocEtOrigine
         flagTracePourSimulation = False
         Call TraceDecoupe
         Call ActiverOngletDecoupe
         pctMasquage.Visible = False
      End If
   End If
End Sub

Private Sub DesactiverOngletDecoupe()
   'désactivation des boutons et scrollbars
   Dim i As Integer
   
   hscChauffeDecoupe.Enabled = False
   comboMatieres.Enabled = False
   For i = 0 To 2
      cmdGestionMatiere(i).Enabled = False
      optEntrerBloc(i).Enabled = False
      optSortirBloc(i).Enabled = False
   Next i
   cmdRetourOrigine.Enabled = False
   cmdDecouper.Enabled = False
   SSTab1.Enabled = False
End Sub

Private Sub ActiverOngletDecoupe()
   'désactivation des boutons et scrollbars
   Dim i As Integer
   
   SSTab1.Enabled = True
   hscChauffeDecoupe.Enabled = True
   comboMatieres.Enabled = True
   For i = 0 To 2
      cmdGestionMatiere(i).Enabled = True
      optEntrerBloc(i).Enabled = True
      optSortirBloc(i).Enabled = True
   Next i
   cmdRetourOrigine.Enabled = True
   cmdDecouper.Enabled = True
End Sub

Private Sub TimerFilm_Timer()
   
   pctDecoupe.Cls
   pctDecoupe.Circle (PointsSimulation(ProgressionDuFilm).x, PointsSimulation(ProgressionDuFilm).y), 1.5, RGB(240, 136, 0)
   If ProgressionDuFilm = UBound(PointsSimulation) Then
      Call TraceTableDecoupe
      Call TraceBlocEtOrigine
      flagTracePourSimulation = False
      Call TraceDecoupe
      Call ActiverOngletDecoupe
      TimerFilm.Enabled = False
      flagSimulationLancee = False
      cmdSimulation.Picture = frmImages.imgSimuler.Picture
      pctMasquage.Visible = False
   End If
   ProgressionDuFilm = ProgressionDuFilm + 1
End Sub

'********************************
'***** SAUVEGARDE DU PROJET *****
'********************************
Private Sub cmdSauver_Click(Index As Integer)
   Dim ExportProfil() As ProfilDATDXF
   Dim NomFichier As String
   Dim RepertoireSauvegarde As String
   Dim Reponse As Integer
   Dim Comment As String
   Dim CodeErreur As Long
   Dim i As Long, j As Long, k As Long
   
' ON NE SAUVE PLUS LA MATIERE
'   If MatiereUtilisee.Chauffe <> ChauffeCourante Then  'si la chauffe a été modifiée sans actualiser la base
'      Reponse = vbNo
'      Reponse = MsgBox(Message(Corps, 49), vbExclamation + vbYesNoCancel, Message(Titre, 49))  'chauffe différente de matière, actualiser?
'      Select Case Reponse
'      Case vbYes
'         Call cmdGestionMatiere_Click(1)  'on appuye sur le bouton d'actualisation de la chauffe
'      Case vbCancel
'         Exit Sub
'      End Select
'   End If
   Select Case Index
   Case 0   'Enregistrer sous...
      RepertoireSauvegarde = LitFichierIni("Fichiers", "DernierRepertoire", App.Path)  'si la clé n'existe pas, on prend App.path
   
      DialogueFichiers.DialogTitle = "Enregistrer/Exporter"
      DialogueFichiers.CancelError = True
      DialogueFichiers.Filter = "Projet MiniCut2d Software (*.mnc)|*.mnc|DAO (*.dxf)|*.dxf"
      DialogueFichiers.FLAGS = cdlOFNOverwritePrompt 'message pour écrasement de fichier
      DialogueFichiers.FileName = ""
      DialogueFichiers.InitDir = RepertoireSauvegarde
      DialogueFichiers.FilterIndex = 1
      On Error GoTo Annuler                     'si fermeture de la fenêtre sans sélection de fichier
      DialogueFichiers.ShowSave                           ' afficher la fenêtre de sauvegarde
      On Error GoTo 0   'on annule la gestion d'erreur pour voir les plantages
'*** MNC
      If DialogueFichiers.FilterIndex = 1 Then  'sauvegarde en .mnc
         NomFichier = DialogueFichiers.FileName  'nom entré par l'utilisateur (chemin complet)
         Call SauverMNC(NomFichier) 'on recrée le treeview pour faire apparaître le fichier et les éventuels nouveaux dossiers
         EcritFichierIni "Fichiers", "DernierProjet", NomFichier 'sauvegarde du nom de fichier dans MiniCut2d_Software.ini
      ElseIf DialogueFichiers.FilterIndex = 2 Then 'sauvegarde en .dxf
'*** DXF
         NomFichier = DialogueFichiers.FileName  'nom entré par l'utilisateur (chemin complet)
         If NbTransf = 0 Then 'on ne peut pas exporter un trajet vide
            MsgBox Message(Corps, 53), vbCritical, Message(Titre, 53)
            Exit Sub
         End If
         ReDim ExportProfil(1 To 1)
         k = 0
         For i = 1 To NbTransf
            With Transf(i)
               For j = 1 To .NbPoints
                  k = k + 1
                  ReDim Preserve ExportProfil(1 To k)
                  ExportProfil(k).x = .Point(j).x
                  ExportProfil(k).y = .Point(j).y
               Next j
            End With
         Next i
         Comment = ""
         CodeErreur = EcrireFichier(ExportProfil, NomFichier, Comment) 'appel de la fonction de CNCTools.dll
         If CodeErreur <> 1 Then 'en cas d'erreur d'export
            MsgBox Message(Corps, 54), vbCritical, Message(Titre, 54)
            Erase ExportProfil
            Exit Sub
         End If
         Erase ExportProfil
'***
      End If
      'mémorisation du dernier dossier utilisé dans MiniCut2d_Software.ini
      RepertoireSauvegarde = GetPathName(NomFichier)
      EcritFichierIni "Fichiers", "DernierRepertoire", RepertoireSauvegarde

      'il faut ajouter le fichier créé dans le treeview, mais seulement s'il ne s'agit pas d'un écrasement
      Call CreerTreeView
   
      'Si on est dans l'onglet découpe, il faut régénérer la vue
      If SSTab1.Tab = 1 Then
         Call TraceDecoupe
      End If

      frmMiniCut2d.Caption = "MiniCut2d Software - " & ExtraitNomFichier(NomFichier)
      flagLeProjetAUnNom = True
      cmdSauver(1).Enabled = True
   Case 1      'Enregistrer : ne fonctionne que si le projet a déjà un fichier de sauvegarde
      If flagLeProjetAUnNom = True Then
         NomFichier = LitFichierIni("Fichiers", "DernierProjet")
         Call SauverMNC(NomFichier)
      End If
   End Select
   
   Exit Sub
Annuler:
'si pas de fichier enregistré
End Sub

'*********************************************
'****** Sauvegarde du projet de découpe ******
'*********************************************
Private Sub SauverMNC(ByVal NomFichier As String)
   Dim i As Long, j As Long
   Dim Point As String
   
   'si on écrase un fichier, on l'efface d'abord
   If VerifierExistenceFichier(NomFichier) = True Then Kill (NomFichier)
   
   'On crée une première section avec les paramètres des séquences : nombre de séquences et nombre de points de chaque séquence
   EcritDecoupeSousFormeIni "ParametresSequences", "NombreSequences", LTrim(Str(NbTransf)), NomFichier
   'ensuite, on crée une section "SequenceXXX" pour chaque séquence dans laquelle les points sont représentés par leur indice
   If NbTransf > 0 Then
      For i = 1 To NbTransf
         With Transf(i)
            EcritDecoupeSousFormeIni "Sequence" & LTrim(Str(i)), "NombrePoints", LTrim(Str(.NbPoints)), NomFichier
            For j = 1 To .NbPoints
               Point = V2P(Format(.Point(j).x, "#####0.000000")) & ":" & V2P(Format(.Point(j).y, "#####0.000000")) & ":" & LTrim(Str(.Point(j).Etat))
               'les points seront de la forme SxxxPyyy = coordX : coordY : Etat , où xxx est le numéro de séquence et yyy le numéro de point
               EcritDecoupeSousFormeIni "Sequence" & LTrim(Str(i)), LTrim(Str(j)), Point, NomFichier
            Next j
         End With
      Next i
   End If
   'le bloc
   EcritDecoupeSousFormeIni "Bloc", "BlocX", V2P(Format(LongBloc, "#####0.000000")), NomFichier
   EcritDecoupeSousFormeIni "Bloc", "BlocY", V2P(Format(HautBloc, "#####0.000000")), NomFichier
   'les paramètres de l'onglet découpe
   If optEntrerBloc(0).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeEntree", "0", NomFichier
   ElseIf optEntrerBloc(1).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeEntree", "1", NomFichier
   ElseIf optEntrerBloc(2).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeEntree", "2", NomFichier
   End If
   If optSortirBloc(0).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeSortie", "0", NomFichier
   ElseIf optSortirBloc(1).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeSortie", "1", NomFichier
   ElseIf optSortirBloc(2).Value = True Then
      EcritDecoupeSousFormeIni "Decoupe", "TypeSortie", "2", NomFichier
   End If
End Sub

'********** CLIGNOTEMENT DU MESSAGE DE DEPASSEMENT ************
Private Sub TimerClignotement_Timer()
   If lblAvertissementDecoupe.Visible = True Then
      lblAvertissementDecoupe.Visible = False
   Else
      lblAvertissementDecoupe.Visible = True
   End If
   NbClignotements = NbClignotements + 1
   If NbClignotements = 6 Then TimerClignotement.Enabled = False
End Sub

'*************** GESTION DES DECALAGES MODE NORMAL *************************
Private Sub optDecalage_Click(Index As Integer)
   If pctValidationDecoupe.Visible = True And ModeSoft = "Normal" Then 'permet de faire des modifs sans effet quand invisible
      If optDecalage(1).Value = True Then
         SequDecalee = DecalerFil(-0.5)  'la fonction renvoie une séquence
      ElseIf optDecalage(3).Value = True Then
         SequDecalee = DecalerFil(0.5)  'la fonction renvoie une séquence
      Else
         SequDecalee = SequDecoupe 'pas de décalage
      End If
      Call CalculDepassementCourses(SequDecalee) '(modDepassements) : On calcule le dépassement du bloc ou des courses pour éventuellement modifier le profil
      If flagDepassementDecoupe = True Then
         MsgBox Message(Corps, 52), vbInformation, Message(Titre, 52) 'dépassement des courses, projet tronqué
         flagDepassementDecoupe = False
      End If

      Call EntreeSortieSiDecalage
      Call CreerSequMouvement  'on associe les différents trajets, ça marche aussi si optDecalage(2)=true (pas de décalage)
      Call CalculDureeEtSalve
      Call TraceDecoupe
   End If
End Sub
Private Sub EntreeSortieSiDecalage()
   With SequEntreeDecalee
      If optEntrerBloc(0).Value = True Then ' entrée par la gauche
         .NbPoints = 3
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = 0
         .Point(2).y = SequDecalee.Point(1).y
         .Point(3).x = SequDecalee.Point(1).x
         .Point(3).y = SequDecalee.Point(1).y
      ElseIf optEntrerBloc(1).Value = True Then    'entrée par le haut
         .NbPoints = 3
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = SequDecalee.Point(1).x
         .Point(2).y = CourseY
         .Point(3).x = SequDecalee.Point(1).x
         .Point(3).y = SequDecalee.Point(1).y
      ElseIf optEntrerBloc(2).Value = True Then    'entrée par la droite
         .NbPoints = 4
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = 0
         .Point(1).y = CourseY
         .Point(2).x = MargeFil + LongBloc + MargeFil
         If .Point(2).x > CourseX Then .Point(2).x = CourseX  'gestion du dépassement pour grand bloc
         .Point(2).y = .Point(1).y
         .Point(3).x = .Point(2).x
         .Point(3).y = SequDecalee.Point(1).y
         .Point(4).x = SequDecalee.Point(1).x
         .Point(4).y = SequDecalee.Point(1).y
      End If
   End With
   With SequSortieDecalee
      If optSortirBloc(0).Value = True Then  'sortie par la gauche
         .NbPoints = 3
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = SequDecalee.Point(SequDecalee.NbPoints).x
         .Point(1).y = SequDecalee.Point(SequDecalee.NbPoints).y
         .Point(2).x = 0
         .Point(2).y = .Point(1).y
         .Point(3).x = 0
         .Point(3).y = CourseY
      ElseIf optSortirBloc(1).Value = True Then    'sortie par le haut
         .NbPoints = 3
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = SequDecalee.Point(SequDecalee.NbPoints).x
         .Point(1).y = SequDecalee.Point(SequDecalee.NbPoints).y
         .Point(2).x = .Point(1).x
         .Point(2).y = CourseY
         .Point(3).x = 0
         .Point(3).y = CourseY
      ElseIf optSortirBloc(2).Value = True Then    'sortie par la droite
         .NbPoints = 4
         ReDim .Point(1 To .NbPoints)
         .Point(1).x = SequDecalee.Point(SequDecalee.NbPoints).x
         .Point(1).y = SequDecalee.Point(SequDecalee.NbPoints).y
         .Point(2).x = MargeFil + LongBloc + MargeFil
         If .Point(2).x > CourseX Then .Point(2).x = CourseX  'gestion du dépassement pour grand bloc
         .Point(2).y = .Point(1).y
         .Point(3).x = .Point(2).x
         .Point(3).y = CourseY
         .Point(4).x = 0
         .Point(4).y = CourseY
      End If
   End With
End Sub

'*********************************************************
'***** Importation d'un fichier dans la bibliothèque *****
'*********************************************************
Private Sub cmdImporterProfil_Click()
   Dim ChaineFichiers As String
   Dim NomsFichiers() As String
   Dim spath As String
   Dim strFolderName As String
   Dim iEndPath As Integer
   Dim i As Integer

   With DialogueFichiers
      .FileName = ""
      .DialogTitle = "Choisissez les fichiers à copier dans la bibliothèque et cliquez sur Ouvrir"
      .FLAGS = cdlOFNAllowMultiselect + cdlOFNExplorer 'permet la sélection multiple
      .CancelError = True
      .Filter = "MiniCut2d Software (*.mnc)|*.mnc|DAO (*.dxf)|*.dxf|Coordonnées et export SketchUp ou Scratch (*.txt)|*.txt|Profil (*.dat)|*.dat|Plotter (*.plt)|*.plt|Encapsulated PostScript (*.eps)|*.eps|Complexes (*.cpx)|*.cpx|RPFC (*.fc)|*.fc|Tous les fichiers connus|*.mnc;*.dxf;*.txt;*.dat;*.plt;*.cpx;*.fc|Tous les fichiers |*.*"
      .FilterIndex = 9
      .InitDir = GetSpecialFolder(CSIDL_PERSONAL) 'Dossier "Mes Documents" ou "Documents"
      On Error GoTo Annuler                     'si fermeture de la fenêtre sans sélection de fichier
      .ShowOpen                              ' afficher la fenêtre d'ouverture
      ChaineFichiers = .FileName         'mémorisation des noms de fichiers, séparés par des espaces
      On Error GoTo 0
   End With
   
   NomsFichiers = Split(ChaineFichiers, vbNullChar)  'Les noms de fichiers se retrouvent dans NomsFichiers(0), NomFichiers(1), etc. séparés par un caractère Null
   
  'the path used in the Browse function
  'must be correctly formatted depending
  'on whether the path is a drive, a
  'folder, or "".
  spath = FixPath(App.Path & "\Bibliotheque")
  'call the function, returning the path
  'selected (or "" if cancelled)
   strFolderName = BrowseForFolderByPath(spath)  'affiche le browser de dossiers et récupère le chemin du rep sélectionné
   If strFolderName = "" Then
      MsgBox Message(Corps, 50), vbInformation, Message(Titre, 50) 'copie annulée
      Exit Sub
   End If
   If UBound(NomsFichiers) = 0 Then    'dans ce cas, un seul fichier a été sélectionné
      CopyFile NomsFichiers(0), strFolderName & "\" & ExtraitNomFichier(NomsFichiers(0)), False
   Else
      For i = 1 To UBound(NomsFichiers)  'NomFichiers(0) contient le chemin, puis viennent les fichiers sans chemin
         CopyFile NomsFichiers(i), strFolderName & "\" & NomsFichiers(i), False
      Next i
   End If
   
   'il faut recréer le Treeview :
   Call CreerTreeView
   MsgBox Message(Corps, 51), vbInformation, Message(Titre, 51)  'validation de la copie
   Exit Sub
Annuler:
End Sub

Private Sub cmdRafraichir_Click()
   'rafraichissement du treeview
   Call CreerTreeView
End Sub

'****** Mettre un fichier / dossier dans la corbeille *******
Private Sub cmdEffacerFichier_Click()
   Dim Reponse As Integer
   Dim Ext As String
   
   If CheminFichier = "" Then
      MsgBox "Vous devez d'abord cliquer sur le fichier à supprimer", vbInformation, "Pas de sélection"
   Else
      If DansCorbeille(CheminFichier, Me.hwnd) Then  'DansCorbeille est une fonction
          MsgBox "Déplacement dans la corbeille effectué", vbInformation, "Suppression"
          cmdRafraichir.Value = True
      Else
          MsgBox "Le déplacement dans la corbeille n'a pas pu être effectué", vbCritical, "Echec de la suppression"
      End If
   End If
End Sub

'********************************************************************************************
'****** Fonctions permettant d'ouvrir l'explorateur de dossier à l'emplacement souhaité *****
'********************************************************************************************
Private Function BrowseForFolderByPath(sSelPath As String) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim lpSelPath As Long
   Dim spath As String * MAX_PATH
   
   With BI
      .hOwner = Me.hwnd
      .pidlRoot = 0
      .lpszTitle = "Sélectionnez le dossier de destination"
      .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
    
      lpSelPath = LocalAlloc(LPTR, Len(sSelPath) + 1)
      CopyMemory ByVal lpSelPath, ByVal sSelPath, Len(sSelPath) + 1
      .lParam = lpSelPath
    
   End With
    
   pidl = SHBrowseForFolder(BI)
   
   If pidl Then
     
      If SHGetPathFromIDList(pidl, spath) Then
         BrowseForFolderByPath = Left$(spath, InStr(spath, vbNullChar) - 1)
      Else
         BrowseForFolderByPath = ""
      End If
      
      Call CoTaskMemFree(pidl)
   
   Else
      BrowseForFolderByPath = ""
   End If
   
  Call LocalFree(lpSelPath)

End Function

Private Function IsValidDrive(spath As String) As Boolean

   Dim buff As String
   Dim nBuffsize As Long
   
  'Call the API with a buffer size of 0.
  'The call fails, and the required size
  'is returned as the result.
   nBuffsize = GetLogicalDriveStrings(0&, buff)

  'pad a buffer to hold the results
   buff = Space$(nBuffsize)
   nBuffsize = Len(buff)
   
  'and call again
   If GetLogicalDriveStrings(nBuffsize, buff) Then
   
     'if the drive letter passed is in
     'the returned logical drive string,
     'return True.
      IsValidDrive = InStr(1, buff, spath, vbTextCompare) > 0
   
   End If

End Function


Private Function FixPath(spath As String) As String

  'The Browse callback requires the path string
  'in a specific format - trailing slash if a
  'drive only, or minus a trailing slash if a
  'file system path. This routine assures the
  'string is formatted correctly.
  '
  'In addition, because the calls to LocalAlloc
  'requires a valid path for the call to succeed,
  'the path defaults to C:\ if the passed string
  'is empty.
  
  'Test 1: check for empty string. Since
  'we're setting it we can assure it is
  'formatted correctly, so can bail.
   If Len(spath) = 0 Then
      FixPath = "C:\"
      Exit Function
   End If
   
  'Test 2: is path a valid drive?
  'If this far we did not set the path,
  'so need further tests. Here we ensure
  'the path is properly terminated with
  'a trailing slash as needed.
  '
  'Drives alone require the trailing slash;
  'file system paths must have it removed.
   If IsValidDrive(spath) Then
      
     'IsValidDrive only determines if the
     'path provided is contained in
     'GetLogicalDriveStrings. Since
     'IsValidDrive() will return True
     'if either C: or C:\ is passed, we
     'need to ensure the string is formatted
     'with the trailing slash.
      FixPath = QualifyPath(spath)
   Else
     'The string passed was not a drive, so
     'assume it's a path and ensure it does
     'not have a trailing space.
      FixPath = UnqualifyPath(spath)
   End If
   
End Function

Private Function QualifyPath(spath As String) As String
 
   If Len(spath) > 0 Then
      If Right$(spath, 1) <> "\" Then
         QualifyPath = spath & "\"
      Else
         QualifyPath = spath
      End If
   Else
      QualifyPath = ""
   End If
   
End Function

Private Function UnqualifyPath(spath As String) As String
  'Qualifying a path involves assuring that its format
  'is valid, including a trailing slash, ready for a
  'filename. Since SHBrowseForFolder will not pre-select
  'the path if it contains the trailing slash, it must be
  'removed, hence 'unqualifying' the path.
   If Len(spath) > 0 Then
      If Right$(spath, 1) = "\" Then
         UnqualifyPath = Left$(spath, Len(spath) - 1)
         Exit Function
      End If
   End If
   UnqualifyPath = spath
End Function

'****** ouverture de la fenêtre de vectorisation *****
Private Sub cmdImporterImage_Click()
   Call GestionLangue(strLangue) 'on modifie la langue, ce qui va charger automatiquement la feuille
   frmImpConv.Show vbModal
End Sub

'****** ouverture de la fenêtre de choix de la langue ******
Private Sub cmdLangue_Click()
   frmLangue.Show vbModal
End Sub



