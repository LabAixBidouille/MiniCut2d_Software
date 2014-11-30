VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImpConv 
   Caption         =   "Vectoriser une image"
   ClientHeight    =   9345
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14325
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   623
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSauverVecto 
      Height          =   435
      Left            =   135
      Picture         =   "frmImportImage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7305
      Width           =   690
   End
   Begin VB.PictureBox pctVisuVecto 
      Height          =   6120
      Left            =   5010
      ScaleHeight     =   6060
      ScaleWidth      =   8580
      TabIndex        =   17
      Top             =   2955
      Width           =   8640
   End
   Begin VB.CommandButton cmdQuitterImpConv 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7305
      Width           =   1035
   End
   Begin VB.PictureBox pctCouleurOriginaleRecadree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2130
      Left            =   7800
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   13
      Top             =   135
      Width           =   3525
   End
   Begin VB.Frame frmLisserTransferer 
      Caption         =   "Vectoriser"
      Height          =   2520
      Left            =   90
      TabIndex        =   11
      Top             =   4665
      Width           =   1965
      Begin VB.Frame frameInterieurExterieur 
         Height          =   750
         Left            =   255
         TabIndex        =   24
         Top             =   220
         Width           =   1470
         Begin VB.OptionButton optInterieur 
            Height          =   450
            Index           =   0
            Left            =   135
            Picture         =   "frmImportImage.frx":060A
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Seulement les contours extérieurs"
            Top             =   180
            Value           =   -1  'True
            Width           =   585
         End
         Begin VB.OptionButton optInterieur 
            Height          =   450
            Index           =   1
            Left            =   750
            Picture         =   "frmImportImage.frx":0BEC
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Contours extérieurs et intérieurs"
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.HScrollBar hscLissageVecto 
         Height          =   240
         LargeChange     =   5
         Left            =   180
         Max             =   20
         Min             =   3
         TabIndex        =   21
         Top             =   1470
         Value           =   10
         Width           =   1590
      End
      Begin VB.CheckBox chkVoirPointsVecto 
         Height          =   480
         Left            =   1335
         Picture         =   "frmImportImage.frx":11CE
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Afficher les points"
         Top             =   1905
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.CommandButton cmdContoursLissageTransfert 
         BackColor       =   &H00C0E0FF&
         Height          =   480
         Left            =   105
         Picture         =   "frmImportImage.frx":1848
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Vectoriser"
         Top             =   1905
         Width           =   1080
      End
      Begin VB.Image imgVecto 
         Height          =   270
         Index           =   1
         Left            =   1425
         Picture         =   "frmImportImage.frx":2252
         Top             =   1095
         Width           =   360
      End
      Begin VB.Image imgVecto 
         Height          =   285
         Index           =   0
         Left            =   120
         Picture         =   "frmImportImage.frx":2894
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label lblLissageVecto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
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
         Left            =   915
         TabIndex        =   22
         Top             =   1170
         Width           =   105
      End
   End
   Begin VB.Frame frmApercu 
      Caption         =   "Noir et Blanc"
      Height          =   1395
      Left            =   90
      TabIndex        =   8
      Top             =   2835
      Width           =   1965
      Begin VB.OptionButton optApercu 
         Height          =   510
         Index           =   1
         Left            =   1305
         Picture         =   "frmImportImage.frx":2EF2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Recommencer"
         Top             =   735
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton optApercu 
         BackColor       =   &H00C0E0FF&
         Height          =   510
         Index           =   0
         Left            =   105
         Picture         =   "frmImportImage.frx":3534
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Convertir en noir et blanc"
         Top             =   735
         Width           =   1095
      End
      Begin VB.HScrollBar hscSensibilite 
         Height          =   270
         LargeChange     =   10
         Left            =   105
         Max             =   200
         Min             =   1
         TabIndex        =   9
         Top             =   315
         Value           =   100
         Width           =   1080
      End
      Begin VB.Label lblSensibilite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "seuil"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1380
         TabIndex        =   10
         ToolTipText     =   "Seuil noir/blanc"
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.Frame frmRecadrer 
      Caption         =   "Recadrer"
      Height          =   930
      Left            =   390
      TabIndex        =   6
      Top             =   1440
      Width           =   1665
      Begin VB.CommandButton cmdRecadrerLaSelection 
         Height          =   510
         Left            =   855
         Picture         =   "frmImportImage.frx":3F3E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Recadrer sur la sélection"
         Top             =   255
         Width           =   630
      End
      Begin VB.Image imgSelectionerUneZone 
         Height          =   390
         Left            =   186
         Picture         =   "frmImportImage.frx":4538
         ToolTipText     =   "Sélectionner une zone puis cliquez"
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame frmImporterImage 
      Caption         =   "Image"
      Height          =   915
      Left            =   90
      TabIndex        =   3
      Top             =   105
      Width           =   1965
      Begin VB.CommandButton cmdImporterImage 
         Height          =   510
         Index           =   1
         Left            =   1050
         Picture         =   "frmImportImage.frx":4DFA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Coller"
         Top             =   270
         Width           =   630
      End
      Begin VB.CommandButton cmdImporterImage 
         Height          =   510
         Index           =   0
         Left            =   270
         Picture         =   "frmImportImage.frx":53C0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Importer"
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.PictureBox pctCouleurOriginale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   11760
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   2
      Top             =   345
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox pctNoirBlancRedim 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   10800
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   1
      Top             =   1515
      Width           =   2430
   End
   Begin VB.PictureBox pctCouleurRedim 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5700
      Left            =   2160
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   363
      TabIndex        =   0
      Top             =   300
      Width           =   5475
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H000000FF&
         Height          =   15
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin MSComDlg.CommonDialog COMDLG 
      Left            =   2595
      Top             =   6270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pbConvertir 
      Height          =   270
      Left            =   2835
      TabIndex        =   20
      Top             =   150
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label lblCouleur 
      AutoSize        =   -1  'True
      Caption         =   "Couleur"
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
      Left            =   2235
      TabIndex        =   23
      Top             =   60
      Width           =   660
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   71
      X2              =   64
      Y1              =   296
      Y2              =   308
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   59
      X2              =   65
      Y1              =   296
      Y2              =   308
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   65
      X2              =   65
      Y1              =   288
      Y2              =   308
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   82
      X2              =   77
      Y1              =   80
      Y2              =   90
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   71
      X2              =   76
      Y1              =   80
      Y2              =   89
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   82
      X2              =   77
      Y1              =   175
      Y2              =   185
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   71
      X2              =   76
      Y1              =   174
      Y2              =   184
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   9
      X2              =   16
      Y1              =   169
      Y2              =   181
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   22
      X2              =   16
      Y1              =   169
      Y2              =   181
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   77
      X2              =   77
      Y1              =   165
      Y2              =   184
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   16
      X2              =   16
      Y1              =   74
      Y2              =   181
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   77
      X2              =   77
      Y1              =   74
      Y2              =   90
   End
   Begin VB.Menu menuImage 
      Caption         =   "Image"
      Visible         =   0   'False
      Begin VB.Menu menuImporter 
         Caption         =   "&Importer image"
      End
      Begin VB.Menu menuColler 
         Caption         =   "&Coller image"
      End
   End
End
Attribute VB_Name = "frmImpConv"
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
   Set frmImporterImage.Font = PoliceNormal
   Set frmApercu.Font = PoliceNormal
   Set frmRecadrer.Font = PoliceNormal
   Set frmRecadrer.Font = PoliceNormal
   Set frmLisserTransferer.Font = PoliceNormal
   Set cmdQuitterImpConv.Font = PoliceNormal
   Set lblSensibilite.Font = PoliceGras
   Set lblLissageVecto.Font = PoliceGras

   Me.KeyPreview = True  'pour intercepter Ctrl+v
   
   
   With frmImpConv
      .Top = frmMiniCut2d.Top + frmMiniCut2d.Height * 0.2
      .Left = frmMiniCut2d.Left + 600
      .Height = frmMiniCut2d.Height - frmMiniCut2d.Height * 0.2 - 110
      .Width = frmMiniCut2d.Width - 700
   End With
   
   With pctCouleurOriginale
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .Appearance = 0
      .AutoSize = True
      .Visible = False
      .ZOrder
   End With
    
   With pctCouleurRedim
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .Appearance = 0
      .ZOrder
   End With
   
   With pctNoirBlancRedim
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .Appearance = 0
      .Top = pctCouleurRedim.Top
      .Left = pctCouleurRedim.Left
      .Visible = False
   End With
   
   With pctCouleurOriginaleRecadree
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .Appearance = 0
      .Visible = False
   End With
   
   With pctVisuVecto
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .Appearance = 0
      .Top = pctCouleurRedim.Top
      .Left = pctCouleurRedim.Left
      .Visible = False
      .ZOrder
      .ScaleTop = .ScaleHeight   'inversion de l'axe Y
      .ScaleHeight = -.ScaleHeight
   End With
   
   With pbConvertir
      .Visible = False
      .Left = pctCouleurRedim.Left + 5
      .Top = pctCouleurRedim.Top + 5
      .Value = 0
      .ZOrder
   End With
   lblCouleur.Caption = ""
   pbConvertir.ZOrder
   
   pctCouleurRedim.Visible = False
   
   LargeurMaxiPct = frmImpConv.ScaleWidth - pctCouleurRedim.Left - 5   'mémorisation des dimensions relatives de pctCouleurRedim
   HauteurMaxiPct = frmImpConv.ScaleHeight - 26

   
   COMDLG.Filter = "Images(*.bmp; *.jpg)|*.bmp; *.jpg"
   
   lblSensibilite.Caption = hscSensibilite.Value
   lblLissageVecto.Caption = hscLissageVecto.Value
   ToleranceNoir = hscSensibilite.Value
   
   blnImageTracee = True
   blnRecadree = False
   blnSelected = False
   cmdRecadrerLaSelection.Enabled = False
   cmdContoursLissageTransfert.Enabled = False
   chkVoirPointsVecto.Enabled = False
   hscLissageVecto.Enabled = False
   optInterieur(0).Enabled = False
   optInterieur(1).Enabled = False
   optApercu(0).Enabled = False
   optApercu(1).Enabled = False
   cmdSauverVecto.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If frmImpConv.Visible = True Then
      If NbSequVecto <> 0 Then
         Call MiseEchelleMachine
         Sequ = SequVecto
         NbSequ = NbSequVecto
         NumSequSel = 0 'initialisation pour mouse_move
         NbSequSel = 0
         CoeffBloc = 1 'reset car nouveau fichier
         Call frmMiniCut2d.CalculAffichageInitial
         Call frmMiniCut2d.TraceSequ
      End If
      Erase BitsPict
      Erase MatricePict
      Erase MatriceBinaire
      Erase CopieMatriceBinaire
      Erase Contours
      Erase NumeroMarqueurs
      Erase profil
      Erase SequVecto
      Erase SequVectoTrace
      NbSequVecto = 0
      pctCouleurOriginale.Cls
      pctCouleurRedim.Cls
      pctNoirBlancRedim.Cls
      pctCouleurOriginaleRecadree.Cls
      pctVisuVecto.Cls
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim CtrlDown As Integer
   'Définition des constantes
   Const vbCtrlMask As Integer = 2
   
   CtrlDown = (Shift And vbCtrlMask) > 0
   If KeyCode = vbKeyV Then   ' interception de Ctrl+v
      If CtrlDown Then
         cmdImporterImage(1).Value = True 'on clique sur le bouton "coller"
      End If
   End If
End Sub

Private Sub Form_Resize()
   LargeurMaxiPct = frmImpConv.ScaleWidth - pctCouleurRedim.Left - 5   'mémorisation des dimensions relatives de pctCouleurRedim
   HauteurMaxiPct = frmImpConv.ScaleHeight - 26
   If pctVisuVecto.Visible = True Then
      Call RepresenterVecto
   End If
End Sub

Private Sub cmdImporterImage_Click(Index As Integer)
   Dim ImagePressePapier As Boolean
   
   'importation de l'image originale par ouverture fichier ou collage presse-papier
   flagVectoRedimCourses = False
   pctCouleurOriginale.Cls  'premier truc à faire pour vider la mémoire
   pctCouleurRedim.Cls
   pctNoirBlancRedim.Cls
   pctCouleurOriginaleRecadree.Cls
   pctVisuVecto.Cls
   Erase SequVecto
   Erase BitsPict
   Erase MatricePict
   Erase MatriceBinaire
   Erase CopieMatriceBinaire
   Erase Contours
   
   NbSequVecto = 0
   Shape1.Move 1, 1, 1, 1  'fait disparaître le cadre de sélection
      
   On Error GoTo Erreur
   
   Select Case Index
   Case 0  'ouvrir un fichier
      blnImageTracee = False
      COMDLG.ShowOpen
      If COMDLG.FileName = "" Then
         Exit Sub
      End If
      pctCouleurOriginale.Picture = LoadPicture(COMDLG.FileName)  'pctCouleurOriginale contient l'image initiale
   Case 1  'coller le contenu du presse-papier
      ImagePressePapier = False
      If Clipboard.GetFormat(vbCFBitmap) Or Clipboard.GetFormat(vbCFDIB) Then
         ImagePressePapier = True
      Else
         MsgBox "Le presse-papier ne contient pas d'image compatible."
         Exit Sub
      End If
      
      blnImageTracee = False
      pctCouleurOriginale.Picture = Clipboard.GetData
   End Select
   pctCouleurRedim.Visible = True
   pctNoirBlancRedim.Visible = False
   lblCouleur.Visible = True
   
   Call TransferOriginaleDansRedim  'redimensionnement pour affichage à l'écran en taille maxi
   
   'initialisation des contrôles
   cmdRecadrerLaSelection.Enabled = True
   If optApercu(1).Value = False Then optApercu(1).Value = True
   If optApercu(0).Enabled = False Then optApercu(0).Enabled = True
   blnSelected = False
   
   Exit Sub
Erreur:
   MsgBox "Impossible d'afficher l'image"
End Sub

Private Sub TransferOriginaleDansRedim()
   'redimensionnement pour affichage à l'écran en taille maxi
   pctCouleurRedim.AutoRedraw = True
   pctCouleurRedim.Cls
   If pctCouleurOriginale.Height > HauteurMaxiPct Or pctCouleurOriginale.Width > LargeurMaxiPct Then
      'si l'image est trop grande, on la redimensionne pour l'affichage
      If pctCouleurOriginale.Width / pctCouleurOriginale.Height >= LargeurMaxiPct / HauteurMaxiPct Then
         pctCouleurRedim.Width = LargeurMaxiPct
         pctCouleurRedim.Height = pctCouleurRedim.Width * pctCouleurOriginale.Height / pctCouleurOriginale.Width
      Else
         pctCouleurRedim.Height = HauteurMaxiPct
         pctCouleurRedim.Width = pctCouleurRedim.Height * pctCouleurOriginale.Width / pctCouleurOriginale.Height
      End If
      pctCouleurRedim.PaintPicture pctCouleurOriginale.Image, 2, 2, pctCouleurRedim.Width - 6, pctCouleurRedim.Height - 6 'pctCouleurRedim affiche l'image redimensionnée
   Else  'sinon affichage à la taille d'origine
      pctCouleurRedim.Height = pctCouleurOriginale.ScaleHeight
      pctCouleurRedim.Width = pctCouleurOriginale.ScaleWidth
      pctCouleurRedim.AutoRedraw = True
      pctCouleurRedim.AutoSize = True
      pctCouleurRedim.Picture = pctCouleurOriginale.Picture
   End If
   pbConvertir.Width = pctCouleurRedim.Width - 10
   pctNoirBlancRedim.Width = pctCouleurRedim.Width
   pctNoirBlancRedim.Height = pctCouleurRedim.Height
   pctCouleurRedim.ZOrder
   pctCouleurRedim.Visible = True
   lblCouleur.Caption = "Image couleur ou niveaux de gris"
   DoEvents 'pour déclencher éventuellement le mousedown/up sur la picturebox en-dessous
   blnImageTracee = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   'effacer la croix sur le plan temporaire quand on n'est plus sur le pctbox de l'image
   pctCouleurRedim.AutoRedraw = False
   pctCouleurRedim.Cls
   pctCouleurRedim.AutoRedraw = True
End Sub

Private Sub hscLissageVecto_Change()
   lblLissageVecto.Caption = hscLissageVecto.Value
End Sub

Private Sub hscLissageVecto_Scroll()
   lblLissageVecto.Caption = hscLissageVecto.Value
End Sub

Private Sub pctCouleurRedim_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'tracé du rectangle de sélection ou menu clic-droit
   
   Select Case Button
   
   Case vbLeftButton
      If blnImageTracee = True Then
         With RectangleSelection
            blnSelected = False
            OrigX = CInt(x)
            OrigY = CInt(y)
            
            .Left = OrigX
            .Top = OrigY
            .Right = OrigX
            .Bottom = OrigY
            Shape1.Move .Left, .Top, (.Right - .Left), (.Bottom - .Top)
         End With
      End If
      
   Case vbRightButton
      PopupMenu menuImage
      
   End Select
   
End Sub

Private Sub pctCouleurRedim_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   'rectangle et croix de sélection sur pctCouleurRedim
   'les valeurs du rectangle de sélection sont mémorisées dans une entité "RectangleSelection"
   
   If Button = vbKeyLButton Then
      If blnImageTracee = True Then
         pctCouleurRedim.AutoRedraw = False
         pctCouleurRedim.Cls
         With RectangleSelection
            If x < OrigX Then
               .Left = x
            ElseIf x > OrigX Then
               .Right = x
            Else
               .Left = x
               .Right = x
            End If
            
            If y < OrigY Then
               .Top = y
            ElseIf y > OrigY Then
               .Bottom = y
            Else
               .Top = y
               .Bottom = y
            End If
            Shape1.Move .Left, .Top, .Right - .Left, .Bottom - .Top
         End With
         pctCouleurRedim.AutoRedraw = True
      End If
   Else
      With pctCouleurRedim
         .AutoRedraw = False
         .Cls
         pctCouleurRedim.Line (0, y)-(.Width, y), vbRed
         pctCouleurRedim.Line (x, 0)-(x, .Height), vbRed
         .AutoRedraw = True
      End With
   End If
End Sub

Private Sub pctCouleurRedim_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'fin sélection
   
   With RectangleSelection
      If Abs(.Right - .Left) > 2 And Abs(.Top - .Bottom) > 2 Then
         blnSelected = True
         cmdRecadrerLaSelection.Enabled = True
      End If
   End With
End Sub

Private Sub cmdRecadrerLaSelection_Click()
   'recadrer en fonction du rectangle de sélection
   
   If blnSelected Then
      CoeffRedim = pctCouleurOriginale.Width / pctCouleurRedim.Width  'coeff de proportionnalité entre les deux picturebox (l'originale et celle affichée)
      
      'on va redimensionner les pctbox de visu pour qu'elles aient le même rapport H/L que la sélection
      With RectangleSelection
         If Abs(.Right - .Left) / Abs(.Bottom - .Top) > LargeurMaxiPct / HauteurMaxiPct Then
            pctCouleurRedim.Width = LargeurMaxiPct
            pctCouleurRedim.Height = LargeurMaxiPct * Abs(.Bottom - .Top) / Abs(.Right - .Left)
            pctCouleurRedim.Cls
         Else
            pctCouleurRedim.Height = HauteurMaxiPct
            pctCouleurRedim.Width = HauteurMaxiPct * Abs(.Right - .Left) / Abs(.Bottom - .Top)
            pctCouleurRedim.Cls
         End If
         pctNoirBlancRedim.Height = pctCouleurRedim.Height
         pctNoirBlancRedim.Width = pctCouleurRedim.Height
         pctCouleurRedim.Cls
         pctNoirBlancRedim.Cls
         
         pctCouleurOriginaleRecadree.Cls
         pctCouleurOriginaleRecadree.Width = Abs(.Right - .Left) * CoeffRedim
         pctCouleurOriginaleRecadree.Height = Abs(.Bottom - .Top) * CoeffRedim
         pctCouleurRedim.AutoRedraw = True
         
         pctCouleurRedim.PaintPicture pctCouleurOriginale.Image, 0, 0, pctCouleurRedim.Width, pctCouleurRedim.Height, _
                                 .Left * CoeffRedim, .Top * CoeffRedim, Abs(.Right - .Left) * CoeffRedim, Abs(.Bottom - .Top) * CoeffRedim
         
         pctCouleurOriginaleRecadree.PaintPicture pctCouleurOriginale.Image, 0, 0, pctCouleurOriginaleRecadree.Width, pctCouleurOriginaleRecadree.Height, _
                                 .Left * CoeffRedim, .Top * CoeffRedim, Abs(.Right - .Left) * CoeffRedim, Abs(.Bottom - .Top) * CoeffRedim
         Shape1.Move 1, 1, 1, 1
         .Left = 0
         .Right = 0
         .Top = 0
         .Bottom = 0
      End With
      blnRecadree = True
      cmdRecadrerLaSelection.Enabled = False 'on ne peut recadrer qu'une fois
   End If
   blnSelected = False
End Sub

Private Sub optApercu_Click(Index As Integer)
   flagVectoRedimCourses = False
   Select Case Index
   Case 0   'passage en noir et blanc
      If blnSelected = True Then  'si un cadre est tracé, on applique la sélection
         cmdRecadrerLaSelection.Value = True
      End If
      Call ImageVersMatrice(pctCouleurRedim)        'on travaille pour l'instant uniquement sur l'image réduite pour ne pas avoir
      pctNoirBlancRedim.Width = LargeurPictDestination  ' à calculer le noir et blanc sur l'image totale
      pctNoirBlancRedim.Height = HauteurPictDestination
      Call MatriceVersNoirBlanc(True, pctNoirBlancRedim) 'true = pas de matrice binaire, uniquement de la visu
      lblCouleur.Caption = "Image noir et blanc"
      pctNoirBlancRedim.Visible = True
      pctCouleurRedim.Visible = False
      
      optApercu(0).Enabled = False
      optApercu(1).Enabled = True
      cmdContoursLissageTransfert.Enabled = True
      cmdSauverVecto.Enabled = False
      chkVoirPointsVecto.Enabled = True
      hscLissageVecto.Enabled = True
      optInterieur(0).Enabled = True
      optInterieur(1).Enabled = True

   Case 1   'annulation
      pctVisuVecto.Visible = False
      pctCouleurRedim.Visible = True
      pctNoirBlancRedim.Visible = False
      
      cmdContoursLissageTransfert.Enabled = False
      cmdSauverVecto.Enabled = False
      chkVoirPointsVecto.Enabled = False
      hscLissageVecto.Enabled = False
      optInterieur(0).Enabled = False
      optInterieur(1).Enabled = False
      optApercu(0).Enabled = True
      optApercu(1).Enabled = False
   End Select
   cmdRecadrerLaSelection.Enabled = False  'on ne peut plus recadrer
End Sub

Private Sub hscSensibilite_Change()
   If optApercu(1).Value = False Then optApercu(1).Value = True
   lblSensibilite.Caption = hscSensibilite.Value
   ToleranceNoir = hscSensibilite.Value
End Sub
Private Sub ImageVersMatrice(Picture As PictureBox) '
   'Copie de l'image vers une matrice
   Dim i As Long, j As Long
   Dim Z As Long
   Dim k As Long, D As Long
   Dim InfoPict As BITMAP
   Dim NL As Integer  'numéro de ligne
   
   GetObject Picture.Image, Len(InfoPict), InfoPict
   Size = InfoPict.bmWidth * InfoPict.bmBitsPixel * InfoPict.bmHeight / 8
   
   ReDim BitsPict(Size) As Byte
   ReDim MatricePict(InfoPict.bmHeight, InfoPict.bmWidth) ' As Pixel
   
   pbConvertir.Min = 0
   pbConvertir.Max = 2 * InfoPict.bmHeight * (InfoPict.bmWidth + 1) / 1000
   pbConvertir.Visible = True
   pbConvertir.ZOrder
   
   GetBitmapBits Picture.Image, Size, BitsPict(1)
   k = InfoPict.bmWidth * 4
   For i = 1 To InfoPict.bmHeight
      D = ((i - 1) * k) + 1
      For j = 1 To InfoPict.bmWidth
         Z = D + ((j - 1) * 4)
         MatricePict(i, j).Blue = BitsPict(Z)
         MatricePict(i, j).Green = BitsPict(Z + 1)
         MatricePict(i, j).Red = BitsPict(Z + 2)
      Next j
      pbConvertir.Value = Int(i * j / 1000)
   Next i
   pbConvertir.Visible = False
   Erase BitsPict 'on libère la mémoire
   HauteurPictDestination = InfoPict.bmHeight + 2
   LargeurPictDestination = InfoPict.bmWidth + 2
End Sub
Private Sub MatriceVersNoirBlanc(PourVisu As Boolean, Optional Picture As PictureBox)
   'si PourVisu=true on fait uniquement l'image
   'si PouVisu=false, on fait uniquement la matrice binaire pour détection des contours
   Dim i As Long, j As Long
   Dim UBy As Integer
   Dim Z As Long
   Dim Mx As Long, My As Long
   
 '  On Error GoTo Erreur
   
   'Copie de la matrice vers picture contours noir
   If PourVisu = True Then  'on va représenter l'image noir et blanc
      ReDim BitsPict(UBound(MatricePict(), 1) * UBound(MatricePict(), 2) * 4)
   Else
      'Copie vers une matrice binaire pour détection des contours
      '******** ATTENTION : on va ajouter une
      '******** bordure "blanche" (=0) tout autour de l'image
      '******** dans la matrice binaire => on commence à 0 et on termine à ubound+1
      'ATTENTION : inversion des X et Y...
      ReDim MatriceBinaire(0 To UBound(MatricePict(), 2) + 1, 0 To UBound(MatricePict(), 1) + 1)
      'le tableau est créé "vide", donc il n'y a rien besoin de faire, il faut juste mettre les
      'valeurs à l'intérieur du rectangle
      '00000000000000000000000000
      '0vvvvvvvvvvvvvvvvvvvvvvvv0
      '0vvvvvvvvvvvvvvvvvvvvvvvv0
      '0vvvvvvvvvvvvvvvvvvvvvvvv0
      '00000000000000000000000000
   End If
   Z = 0
   
   UBy = UBound(MatricePict(), 1) + 1  'pour remettre les Y de bas en haut dans la matrice de 0 et de 1
   
   'on paramètre la progressbar
   pbConvertir.Min = 0
   pbConvertir.Max = 2 * UBound(MatricePict(), 1) * (UBound(MatricePict(), 2) + 1) / 1000
   pbConvertir.Value = pbConvertir.Max / 2
   pbConvertir.Visible = True
   For i = 1 To UBound(MatricePict(), 1)  'variations suivant Y, de haut en bas
      For j = 1 To UBound(MatricePict(), 2)  'variations suivant X, de gauche à droite
         Z = Z + 1
         DoEvents
         If 0.299 * MatricePict(i, j).Red + 0.587 * MatricePict(i, j).Green + 0.114 * MatricePict(i, j).Blue < ToleranceNoir Then
            If PourVisu = True Then  'on va représenter l'image noir et blanc
               BitsPict(Z + 2) = 0
               BitsPict(Z + 1) = 0
               BitsPict(Z) = 0
            Else
               MatriceBinaire(j, UBy - i) = 1 '1 = pixel noir
            End If
         Else
            If PourVisu = True Then  'on va représenter l'image noir et blanc
               BitsPict(Z + 2) = 255
               BitsPict(Z + 1) = 255
               BitsPict(Z) = 255
            Else
               MatriceBinaire(j, UBy - i) = 0 '1 = pixel blanc
            End If
         End If
         Z = Z + 3
         DoEvents
      Next j
      pbConvertir.Value = pbConvertir.Max / 2 + Int(i * j / 1000)
      DoEvents
   Next i
   pbConvertir.Visible = False
   
   If PourVisu = True Then  'on va représenter l'image noir et blanc
      Picture.Cls
      SetBitmapBits Picture.Image, UBound(BitsPict), BitsPict(1)
      Picture.Refresh
      Erase BitsPict  'on libère la mémoire
   Else
      CopieMatriceBinaire = MatriceBinaire  'on duplique pour pouvoir tester ensuite plusieurs lissages
   End If
   
   Exit Sub
Erreur:
   MsgBox "Arrêt de la conversion"
End Sub

Public Function TrouverContours() As Long
   Dim sx As Long
   Dim sy As Long
   Dim decY As Long
   Dim trouve As Boolean
   Dim Premier As Boolean
   Dim TP As Boolean
   Dim Fin As Boolean
   Dim x As Long
   Dim y As Long
   Dim X1 As Long
   Dim Y1 As Long
   Dim fx As Long
   Dim fy As Long
   Dim direction As Integer
        '0 de gauche à droite
        '1 de bas en haut
        '2 de droite à gauche
        '3 de haut en bas
   Dim DebutX As Long
   Dim DebutY As Long
   Dim i As Long, j As Long
   
   TP = True 'on tourne dans le sens trigo (on est toujours à l'extérieur)
   
   'Attention : on cherche le contour "externe" des pixels, donc si un pixel a pour
   'coordonnées (2,2) il a 4 points qui forment son contour, dans un repère légèrement
   'décalé : (2,2) (3,2) (3,3) (2,3)
   
   'recherche des contours
   NombreContours = 0
   ReDim Contours(1 To 1)
   DebutX = 0
   DebutY = 0
   Do
      'On cherche le premier pixel noir (=1):
      Premier = True
      trouve = False
      For sy = DebutY To UBound(MatriceBinaire(), 2)
         For sx = DebutX To UBound(MatriceBinaire(), 1)
            If MatriceBinaire(sx, sy) = 1 Then
               trouve = True
               Exit For
            End If
         Next sx
         If trouve = True Then Exit For
         If Premier = True Then
            Premier = False
            DebutX = 0
         End If
      Next sy
      If trouve = False Then Exit Do
            
      'on a trouvé un pixel, donc un contour, il a au moins deux points correspondant au côté
      'inférieur du pixel (cf. explication ci-dessus)
      NombreContours = NombreContours + 1
      ReDim Preserve Contours(1 To NombreContours)
      '****
      With Contours(NombreContours)
         .NombrePoints = 1
         ReDim .Point(1 To 1)
         'On détecte le contour :
         'le premier point a la même coordonnée que le pixel (coin bas gauche du pixel)
         x = sx
         y = sy
         X1 = x
         Y1 = y
         .Point(1).x = x
         .Point(1).y = y
         'on tourne dans le sens trigo, le deuxième points est en bas à droite du pixel
         .NombrePoints = 2
         ReDim Preserve .Point(1 To .NombrePoints)
         x = sx + 1
         y = sy
         .Point(2).x = x
         .Point(2).y = y
         .Point(2).D = 0 'on vient de la gauche
         direction = 0
         'Remarque : on évite de nombreux tests en considérant que les bords de l'image sont blancs
         Do
            Select Case direction
               Case 0 'de gauche à droite
                  'On teste le pixel en haut à droite :
                  fx = x
                  fy = y
                  If MatriceBinaire(fx, fy) = 1 Then
                     'le pixel en haut à droite est noir
                     'on doit tester deux autres possibilités
                     'on teste le pixel en bas à droite :
                     fx = x
                     fy = y - 1
                     If MatriceBinaire(fx, fy) = 1 Then
                        'le pixel en haut à droite est noir, le pixel en bas à droite est noir
                        ' => on descend
                        direction = 3
                        y = y - 1
                     Else
                        'le pixel en haut à droite est noir, le pixel en bas à droite est blanc
                        ' => on va à droite
                        direction = 0
                        x = x + 1
                     End If
                  Else
                     'le pixel en haut à droite est blanc
                     If TP Then
                        ' => on monte :
                        direction = 1
                        y = y + 1
                     Else
                        'on teste le pixel en bas à droite :
                        fx = x
                        fy = y - 1
                        If MatriceBinaire(fx, fy) = 1 Then
                           'le pixel en bas à droite est noir
                           ' => on descend
                           direction = 3
                           y = y - 1
                        Else
                           'le pixel en bas à droite est blanc
                           ' => on monte :
                           direction = 1
                           y = y + 1
                        End If
                     End If
                  End If

               Case 1 'de bas en haut
                  'On teste le pixel en haut à gauche :
                  fx = x - 1
                  fy = y
                  If MatriceBinaire(fx, fy) = 1 Then
                     'le pixel en haut à gauche est noir
                     'on doit tester deux autres possibilités
                     'on teste le pixel en haut à droite :
                     fx = x
                     fy = y
                     If MatriceBinaire(fx, fy) = 1 Then
                        'le pixel en haut à gauche est noir, le pixel en haut à droite est noir
                        ' => on va à droite
                        direction = 0
                        x = x + 1
                     Else
                        'le pixel en haut à gauche est noir, le pixel en haut à droite est blanc
                        ' => on monte :
                        direction = 1
                        y = y + 1
                     End If
                  Else
                     'le pixel en haut à gauche est blanc
                     If TP Then
                        ' => on va à gauche :
                        direction = 2
                        x = x - 1
                     Else
                        'on teste le pixel en haut à droite :
                        fx = x
                        fy = y
                        If MatriceBinaire(fx, fy) = 1 Then
                           'le pixel en haut à droite est noir
                           ' => on va à droite
                           direction = 0
                           x = x + 1
                        Else
                           'le pixel en haut à droite est blanc
                           ' => on va à gauche :
                           direction = 2
                           x = x - 1
                        End If
                     End If
                  End If

               Case 2 'de droite à gauche
                  'On teste le pixel en bas à gauche :
                  fx = x - 1
                  fy = y - 1
                  If MatriceBinaire(fx, fy) = 1 Then
                     'le pixel en bas à gauche est noir
                     'on doit tester deux autres possibilités
                     'on teste le pixel en haut à gauche :
                     fx = x - 1
                     fy = y
                     If MatriceBinaire(fx, fy) = 1 Then
                        'le pixel en haut à gauche est noir
                        ' => on monte :
                        direction = 1
                        y = y + 1
                     Else
                        'le pixel en haut à gauche est blanc
                        ' => on va à gauche :
                        direction = 2
                        x = x - 1
                     End If
                  Else
                     'le pixel en bas à gauche est blanc
                     If TP Then
                        '=> on descend
                        direction = 3
                        y = y - 1
                     Else
                        'on teste le pixel en haut à gauche :
                        fx = x - 1
                        fy = y
                        If MatriceBinaire(fx, fy) = 1 Then
                           'le pixel en haut à gauche est noir
                           ' => on monte :
                           direction = 1
                           y = y + 1
                        Else
                           'le pixel en haut à gauche est blanc
                           '=> on descend
                           direction = 3
                           y = y - 1
                        End If
                     End If
                  End If

               Case 3 'de haut en bas
                  'On teste le pixel en bas à droite :
                  fx = x
                  fy = y - 1
                  If MatriceBinaire(fx, fy) = 1 Then
                     'le pixel en bas à droite est noir
                     'on doit tester deux autres possibilités
                     'on teste le pixel en bas à gauche :
                     fx = x - 1
                     fy = y - 1
                     If MatriceBinaire(fx, fy) = 1 Then
                        'le pixel en bas à droite est noir, le pixel en bas à gauche est noir
                        ' => on va à gauche :
                        direction = 2
                        x = x - 1
                     Else
                        'le pixel en bas à droite est noir, le pixel en bas à gauche est blanc
                        '=> on descend
                        direction = 3
                        y = y - 1
                     End If
                  Else
                     'le pixel en bas à droite est blanc
                     If TP Then
                        ' => on va à droite
                        direction = 0
                        x = x + 1
                     Else
                        'on teste le pixel en bas à gauche :
                        fx = x - 1
                        fy = y - 1
                        If MatriceBinaire(fx, fy) = 1 Then
                           'le pixel en bas à gauche est noir
                           ' => on va à gauche :
                           direction = 2
                           x = x - 1
                        Else
                           'le pixel en bas à gauche est blanc
                           ' => on va à droite
                           direction = 0
                           x = x + 1
                        End If
                     End If
                  End If
            End Select
            .NombrePoints = .NombrePoints + 1
            ReDim Preserve .Point(1 To .NombrePoints)
            .Point(.NombrePoints).x = x
            .Point(.NombrePoints).y = y
            .Point(.NombrePoints).D = direction
            
            If x = X1 And y = Y1 Then 'si on est revenu sur le premier point, on a terminé
               'on ne met pas de direction verticale au premier point (sinon pb sur les couples car premier=dernier)
               'MsgBox "Nombre de points: " & .NombrePoints
               Exit Do
            End If
         Loop
      
      End With
      
      'Pour pouvoir trouver les autres contours, il faut faire disparaître celui qu'on vient de trouver
      Call CalculerCouplesContour(Contours(NombreContours)) 'on trouve les couples de points horizontaux
      Call BlanchirContour(Contours(NombreContours)) 'pour le faire disparaître pour le prochain passage (voir dans la procédure pour les détails)
      
      DebutX = sx
      DebutY = sy
   Loop
   TrouverContours = NombreContours
   'calcul des mini-maxi
   For i = 1 To NombreContours
      With Contours(i)
         .Xmin = .Point(1).x
         .Xmax = .Point(1).x
         .Ymin = .Point(1).y
         .Ymax = .Point(1).y
         For j = 2 To .NombrePoints
            If .Point(j).x < .Xmin Then .Xmin = .Point(j).x
            If .Point(j).x > .Xmax Then .Xmax = .Point(j).x
            If .Point(j).y < .Ymin Then .Ymin = .Point(j).y
            If .Point(j).y > .Ymax Then .Ymax = .Point(j).y
         Next j
      End With
   Next i
End Function
Private Sub CalculerCouplesContour(C As Contour)
   'Il faut peindre l'intérieur du contour, c'est à dire inverser la valeur de tous les pixels contenus dans le contour
   'On procéde pour chaque seguement de droite contenu dans le contour
   'Il faut d'abord faire la liste des vecteurs vers le haut (B,D,F,H) et des vecteurs vers le bas (A,C,E,G)
   'Ensuite on forme des couples de vecteurs (A,B) (C,D) (E,F) (G,H)
   'Puis on inverse les seguements de droite entre chaque couple de vecteur
   '------------------------------------------------------------------------
   '-----------A***************************B------------C*****D-------------
   '-------E******************************************************F---------
   '---------------G***********************H--------------------------------
   '------------------------------------------------------------------------

   Dim ListeH() As PointOriente
   Dim ListeB() As PointOriente
   Dim NbPtH As Long
   Dim NbPtB As Long
   Dim NbCouples As Long
   Dim i As Long, j As Long
   Dim flagP3 As Boolean
   Dim pctNoirBlancRedim As PointOriente
   'La direction du point p(i) indique le sens du vecteur p(i-1)p(i)
   'le point p(i) désigne aussi le vecteur p(i-1)p(i)

   'Liste des vecteurs :
   ReDim ListeH(1 To 1)
   ReDim ListeB(1 To 1)
   NbPtH = 0
   NbPtB = 0
   For i = 1 To C.NombrePoints
      Select Case C.Point(i).D
         Case 1 'vers le haut
            NbPtH = NbPtH + 1
            ReDim Preserve ListeH(1 To NbPtH)
            ListeH(NbPtH) = C.Point(i)
         Case 3 'vers le bas
            NbPtB = NbPtB + 1
            ReDim Preserve ListeB(1 To NbPtB)
            ListeB(NbPtB) = C.Point(i)
       End Select
   Next i

   If NbPtB = NbPtH Then
       'Formation des couples :
       NbCouples = 0
       ReDim C.Couples(1 To 1)
       For i = 1 To NbPtH
           'Il faut chercher le point p3 correspondant au couple p,p3
           flagP3 = False
           For j = 1 To NbPtB
               'On teste si p et p2 sont sur la même ligne  (p=c.Point(i) et p2=c.Point(j))
               If ListeH(i).y = ListeB(j).y + 1 Then '+1 à cause de la flèche du vecteur?
                   'On teste si p et p2 sont dans le bon sens
                   If ListeB(j).x < ListeH(i).x Then
                       If flagP3 = False Then
                           pctNoirBlancRedim = ListeB(j)
                           flagP3 = True
                       Else
                           'If p.X - p2.X < p.X - p3.X Then
                           If ListeB(j).x > pctNoirBlancRedim.x Then
                               pctNoirBlancRedim = ListeB(j)
                               flagP3 = True
                           End If
                       End If
                   End If
               End If
           Next
           If flagP3 = False Then
               MsgBox "Erreur - impossible de former un couple"
               End
           Else
               NbCouples = NbCouples + 1
               ReDim Preserve C.Couples(1 To NbCouples)
               C.Couples(NbCouples).Point1 = ListeH(i)
               C.Couples(NbCouples).Point2 = pctNoirBlancRedim
           End If
       Next i
   Else
       MsgBox "Erreur - le nombre de vecteurs vers le haut est différent du nombre de vecteurs vers le bas"
   End If
End Sub
Private Sub BlanchirContour(C As Contour) 'tout mettre à 0 pour le faire disparaître
   Dim i As Long, x As Long, j As Long
   
   If optInterieur(0).Value = True Then 'on colorie tout l'intérieur
      For i = 1 To UBound(C.Couples)
         With C.Couples(i)
            For x = .Point2.x To .Point1.x - 1
               MatriceBinaire(x, .Point2.y) = 0
            Next x
         End With
      Next i
   ElseIf optInterieur(1).Value = True Then 'on inverse l'intérieur
      For i = 1 To UBound(C.Couples)
         With C.Couples(i)
            For x = .Point2.x To .Point1.x - 1
               If MatriceBinaire(x, .Point2.y) = 0 Then
                  MatriceBinaire(x, .Point2.y) = 1
               ElseIf MatriceBinaire(x, .Point2.y) = 1 Then
                  MatriceBinaire(x, .Point2.y) = 0
               End If
            Next x
         End With
      Next i
   End If
End Sub

'_________________________________________________________________________________________________
'Private Sub LisserContours_Contour(ByRef C As Contour, ByVal DemiPasMaxLissage As Integer, _
'            ByVal PourcentageDemiPasMinLissageCoins As Double, ByVal ErreurMaxLissage As Double)
'            'j'ai viré deux-trois paramètre concernant les filtres et la suppression des
'            'points alignés
'
'   Dim c2 As Contour
'
'   Dim p As PointOriente
'   Dim pprec As PointOriente
'   Dim psuiv As PointOriente
'
'   Dim pr As PointOriente
'
'   Dim tmpListe As Contour
'
'   Dim DemiPasMinLissage As Integer
'   DemiPasMinLissage = DemiPasMaxLissage * PourcentageDemiPasMinLissageCoins / 100
'
'   Dim nbFiltres As Integer
'   nbFiltres = DemiPasMaxLissage - DemiPasMinLissage + 1
'   Dim FiltreQ() As Double
'   ReDim FiltreQ(nbFiltres)
'   Dim FiltreM() As Double
'   ReDim FiltreM(DemiPasMaxLissage + 1, nbFiltres)
'   Dim FiltreDemiPas() As Integer
'   ReDim FiltreDemiPas(nbFiltres)
'
'   Dim t As Double
'
'   Dim SommeX As Double
'   Dim SommeY As Double
'
'   Dim Dist As Double
'
'   Dim i As Integer
'   Dim j As Integer
'   Dim k As Integer
'   Dim iprec As Integer
'   Dim isuiv As Integer
'   Dim jmax As Integer
'
'   'CALCUL DES COEFFICIENTS DES FILTRES DE LISSAGE A PAS VARIABLE ===============================================================
'
'    'Les DemiPas :
'    For i = 0 To nbFiltres - 1
'       'Le filtre i a un demiPas égal à DemiPasMaxLissage, le filtre nbFiltres a un demiPas égal à DemiPasMinLissage
'       FiltreDemiPas(i) = DemiPasMaxLissage - i
'    Next i
'   'Les Coefficients
'   'Filtre gaussien    (NB : j'ai viré les autres par rapport au code initial d'electroremy)
'      For i = 0 To nbFiltres - 1
'          For j = 0 To DemiPasMaxLissage 'Amélioration du filtre - For j = 0 To FiltreDemiPas(i)
'              t = j * 3 / FiltreDemiPas(i)
'              FiltreM(j, i) = Exp(-t * t / 2)
'          Next j
'          FiltreDemiPas(i) = DemiPasMaxLissage 'Amélioration du filtre
'      Next i
'   'Les Quotiens pour la normalisation finale
'   For i = 0 To nbFiltres - 1
'      FiltreQ(i) = FiltreM(0, i)
'      For j = 1 To FiltreDemiPas(i)
'         FiltreQ(i) = FiltreQ(i) + 2 * FiltreM(j, i)
'      Next j
'   Next i
'   'On va pré-diviser, cela va éviter de diviser à chaque fois dans la boucle SommeX et SommeY :
'   For i = 0 To nbFiltres - 1
'      For j = 0 To FiltreDemiPas(i)
'         FiltreM(j, i) = FiltreM(j, i) / FiltreQ(i)
'      Next
'   Next i
'   'LISSAGE A PAS VARIABLE ==================================================================================================
'
'   'ReDim tmpListe.Point(1 To 1)
'   For i = 1 To C.NombrePoints
'      'SommeX(j) et SommeY(j) contiennent la somme pour la moyenne glissante de demiPas=j
'      'Il y a moyen d'optimiser ce code pour ne pas tout recalculer pour chaque point i :
'      p = C.Point(i)
'      'Calcul du pr(i) correspondant à p(i)
'      For k = 0 To nbFiltres - 1 'For demiPas = DemiPasMaxLissage To DemiPasMinLissage Step -1
'         'Contribution du point p(i)
'         SommeX = p.x * FiltreM(0, k)
'         SommeY = p.y * FiltreM(0, k)
'         jmax = FiltreDemiPas(k)
'         If jmax > C.NombrePoints Then jmax = C.NombrePoints
'         For j = 1 To jmax
'            iprec = i - j
'            isuiv = i + j
'            If iprec < 1 Then iprec = 1 'c.NombrePoints + iprec
'            If isuiv > C.NombrePoints Then isuiv = C.NombrePoints 'isuiv - c.NombrePoints
'            pprec = C.Point(iprec)
'            psuiv = C.Point(isuiv)
'            'Contribution des points p(i-j) et p(i+j)
'            SommeX = SommeX + FiltreM(j, k) * (pprec.x + psuiv.x)
'            SommeY = SommeY + FiltreM(j, k) * (pprec.y + psuiv.y)
'         Next
'         pr.x = SommeX
'         pr.y = SommeY
'         'Calcul de la distance entre p(i) et pr(i)
'         Dist = Sqr((pr.x - p.x) * (pr.x - p.x) + (pr.y - p.y) * (pr.y - p.y))
'         'Plus la valeur de ErreurMaxLissage est faible, plus on va diminuer le nombre de pas
'         If Dist <= ErreurMaxLissage Then Exit For
'      Next k
'      tmpListe.NombrePoints = tmpListe.NombrePoints + 1
'      ReDim Preserve tmpListe.Point(1 To tmpListe.NombrePoints)
'      tmpListe.Point(tmpListe.NombrePoints) = pr
'   Next i
'   C = tmpListe
'End Sub

Private Sub cmdContoursLissageTransfert_Click()
   Dim i As Long, j As Long, k As Long, m As Long
   Dim RechercheContours As Long
   Dim DemiPasMaxLiss As Integer, PourcentDemiPasMinLiss As Double, ErrMaxLiss As Double
   Dim ContourTemp() As Contour
   Dim NbrTroncons As Long
   Dim Xgauche As Single, NumPtXGauche As Long

  ' On Error GoTo ErrorVecto  'gestion des erreurs pour fichier trop gros
   
   flagVectoRedimCourses = False
   '******* Recherche des contours
   If blnRecadree = False Then   'pas de recadrage, on travaille sur toute l'image
      Call ImageVersMatrice(pctCouleurOriginale)
   Else        'image recadrée, on va travailler seulement sur une partie
      Call ImageVersMatrice(pctCouleurOriginaleRecadree)
   End If
   
   Call MatriceVersNoirBlanc(False) 'uniquement du calcul, pas de visu
   
'   MatriceBinaire = CopieMatriceBinaire  'si on est déjà passé dans cette procédure, la matrice est blanche
   RechercheContours = TrouverContours()
   
   If RechercheContours = 0 Then
      MsgBox "Aucun contour trouvé"
      cmdContoursLissageTransfert.Enabled = False
      chkVoirPointsVecto.Enabled = False
      hscLissageVecto.Enabled = False
      optInterieur(0).Enabled = False
      optInterieur(1).Enabled = False
      Exit Sub
   End If
   '********* lissage et affichage
   If NombreContours > 0 Then
      'nettoyage
      NbSequVecto = NombreContours
   
      If NbSequVecto <> 0 Then 'si le fichier n'est pas représentable et ne plante pas
         frmMiniCut2d.pctSequ.Cls
         ReDim SequVecto(1 To NbSequVecto)
         SequVecto(NbSequVecto).NbPoints = 0
         ReDim SequVectoTrace(1 To NbSequVecto)
         For i = 1 To NbSequVecto
            With SequVecto(i)
               .NbPoints = Contours(i).NombrePoints
               ReDim Preserve .Point(1 To .NbPoints)
               ReDim Preserve SequVectoTrace(i).Point(1 To .NbPoints)
               For j = 1 To .NbPoints
                  .Point(j).x = Contours(i).Point(j).x
                  .Point(j).y = Contours(i).Point(j).y
                  .Point(j).Mark = Contours(i).Point(j).Mark
               Next j
               Call MaxiMiniSequ(SequVecto(i))
            End With
         Next i
      End If
      '****************NETTOYAGE
      For j = 1 To NbSequVecto
         Call MaxiMiniSequ(SequVecto(j))
         With SequVecto(j)
            ReDim profil(0 To .NbPoints - 1)
            For i = 1 To .NbPoints
               profil(i - 1).x = .Point(i).x
               profil(i - 1).y = .Point(i).y
            Next i
            EpsilonNettoyage = 1.9  'les coordonnées sont en pixels
            NbPoints = NettoyageMoindreCarres(EpsilonNettoyage)     'nettoyage des points, angles déterminés par intersection des droites moyennes
            .NbPoints = NbPoints
            ReDim .Point(1 To .NbPoints)               'tableau du profil initial, inclus dans le type "Trajet"
            ReDim SequVectoTrace(j).Point(1 To .NbPoints)
            For i = 1 To .NbPoints                      'Transfert des points dans mon type de tableau
               .Point(i).x = profil(i - 1).x
               .Point(i).y = profil(i - 1).y
            Next i
         End With
         Call MaxiMiniSequ(SequVecto(j))
      Next j
      'Suppression des éventuels points doubles pour éviter toute surprise au décalage
      For j = 1 To NbSequVecto
         With SequVecto(j)
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
      ' on renumérote à partir du point le plus à gauche et on change de sens
      '****************
      For i = 1 To NbSequVecto
         With SequVecto(i)
            If .NbPoints > 1 Then 'on doit conserver les points uniques
           '    If Abs(.Point(.NbPoints).X - .Point(1).X) < 0.001 And Abs(.Point(.NbPoints).Y - .Point(1).Y) < 0.001 Then
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
            '   End If
               Call InverserSensSequ(SequVecto(i))
               Call MaxiMiniSequ(SequVecto(i))
            End If
         End With
      Next i
      
      '*************************************
      ' affichage
      '*************************************
      pctNoirBlancRedim.Visible = False
      Call RepresenterVecto
      lblCouleur.Caption = "Contours vectorisés"
      cmdSauverVecto.Enabled = True
   End If
   Exit Sub
ErrorVecto:
   MsgBox "Vectorisation error : the picture must be too complexe ; try to use Inkscape with better_dxf_output"
   
End Sub

'***************************************************************
'**** LISSAGE des PROFILS (MOINDRES CARRES ET TRUCS DIVERS *****
'***************************************************************
Private Function NettoyageMoindreCarres(ByVal Dist As Single) As Long     'renvoie le nombre de points

   Dim i As Long, j As Long, k As Long, L As Long
   Dim lined As Long
   Dim Nbr_Out As Long
   Dim A As Single, B As Single
   Dim d2 As Single, d1 As Single
   Dim epsilon As Single
   Dim Xmoy As Single, Ymoy As Single, cov As Single, varX As Single, varY As Single
   Dim Xdernier As Single, Ydernier As Single
   Dim CoeffMC() As MoindreCarres 'Coefficiens A et B des droites de moindre carrés
   Dim Nbr_Mini_Points As Integer
   Dim CoeffDir As Single
   Dim Somme_Dist As Single
   
   ReDim CoeffMC(0 To 0)
   
   If UBound(profil) < 3 Then
      NettoyageMoindreCarres = UBound(profil) + 1
      Exit Function
   End If
   epsilon = Dist * Dist  'On élève au carré pour aller plus vite pour les comparaisons
   i = 0
   Nbr_Out = 0
   Do
      j = 1
      Do
         lined = 0
         j = j + 1
         Somme_Dist = 0
         For k = 1 To j - 1
            'Est-ce que tous les points intermédiaires sont alignés ?
            A = profil(i + j).x - profil(i).x
            B = profil(i + j).y - profil(i).y
            d2 = A * A + B * B
            If d2 <= epsilon Then
               ' Les points i and i+j sont trop près
               d1 = (profil(i + k).x - profil(i).x) * (profil(i + k).x - profil(i).x) + (profil(i + k).y - profil(i).y) * (profil(i + k).y - profil(i).y)
               If d1 <= epsilon Then
                  Somme_Dist = Somme_Dist + d1
                  If Somme_Dist > hscLissageVecto.Value / 10 * epsilon Then
                     Exit For
                  End If
                  lined = lined + 1         ' Le point i+k est aussi trop près
               Else
                  Exit For                  ' Le point i+k n'est pas aligné, pas la peine de continuer !
               End If
            Else
               d1 = ((profil(i + k).x - profil(i).x) * A + (profil(i + k).y - profil(i).y) * B) / d2
               If d1 < 0# Then
                  d1 = 0#                   ' La référence est le point i
               Else
                  If d1 > 1# Then d1 = 1#   ' La référence est le point i+j
               End If
               A = profil(i).x + d1 * A - profil(i + k).x
               B = profil(i).y + d1 * B - profil(i + k).y
               If A * A + B * B <= epsilon Then
                  Somme_Dist = Somme_Dist + A * A + B * B
                  If Somme_Dist > hscLissageVecto.Value / 10 * epsilon Then
                     Exit For
                  End If
                  lined = lined + 1         ' point aligné
               Else
                  Exit For                  ' Le point i+k n'est pas aligné, pas besoin de continuer
               End If
            End If
         Next k
      Loop While lined = j - 1 And i + j + 1 <= UBound(profil) 'Boucle jusqu'à ce qu'un point ne soit pas aligné
      If lined <> j - 1 Or i + j - 1 > UBound(profil) - 2 Then
         ReDim Preserve CoeffMC(0 To Nbr_Out)
         CoeffMC(Nbr_Out).x = profil(i + j - 1).x  'mémorisation systématique du point
         CoeffMC(Nbr_Out).y = profil(i + j - 1).y
         's'il y a très peu de points, on ne fait pas la droite des moindres carrés, on prend la droite qui relie les points i et i+j-1
'         Nbr_Mini_Points = 10
'         If j - 1 < Nbr_Mini_Points Then  'i + j - 1 - i = j - 1
'            'calculer la droite  passant par profil (i) & profil(i + j - 1)
'            If profil(i).x = profil(i + j - 1).x Then  'droite verticale
'               CoeffMC(Nbr_Out).Sup45deg = True
'               CoeffMC(Nbr_Out).C = 0
'               CoeffMC(Nbr_Out).D = profil(i).x
'            Else
'               CoeffMC(Nbr_Out).Sup45deg = False  'ce n'est pas nécessairement vrai, mais c'est pour que l'algo utilise y=A.x+B
'               CoeffMC(Nbr_Out).A = (profil(i + j - 1).y - profil(i).y) / (profil(i + j - 1).x - profil(i).x)
'               CoeffMC(Nbr_Out).B = profil(i).y - CoeffMC(Nbr_Out).A * profil(i).x
'            End If
'         Else
            'on calcule la droite de régression entre les points i et i+j-1
            'calcul des moyennes
            Xmoy = 0
            Ymoy = 0
            For L = i To i + j - 1
               Xmoy = Xmoy + profil(L).x
               Ymoy = Ymoy + profil(L).y
            Next L
            Xmoy = Xmoy / j
            Ymoy = Ymoy / j
            'calcul de la covariance et des variances:
            cov = 0
            varX = 0
            varY = 0
            For L = i To i + j - 1
               cov = cov + (profil(L).x - Xmoy) * (profil(L).y - Ymoy)
               varX = varX + (profil(L).x - Xmoy) * (profil(L).x - Xmoy)
               varY = varY + (profil(L).y - Ymoy) * (profil(L).y - Ymoy)
            Next L
            cov = cov / j
            varX = varX / j
            varY = varY / j
            ' les coefficients de la droite des moindres carrés précédant le changement de direction sont donc :
            If varX <> 0 And varY <> 0 Then 'si on n'est ni vertical, ni horizontal
               If Abs(cov / varX) <= 1 Then 'pente inférieure ou égale à 45°, on prend les moindres carrés verticaux
                  CoeffMC(Nbr_Out).Sup45deg = False
                  CoeffMC(Nbr_Out).A = cov / varX
                  CoeffMC(Nbr_Out).B = Ymoy - CoeffMC(Nbr_Out).A * Xmoy
               Else
                  'on prend les moindre carrés horizontaux et on calcule les coefficients de x=C.y+D
                  CoeffMC(Nbr_Out).Sup45deg = True
                  CoeffMC(Nbr_Out).C = cov / varY
                  CoeffMC(Nbr_Out).D = Xmoy - CoeffMC(Nbr_Out).C * Ymoy
               End If
            Else  'on est vertical ou horizontal
               If varX = 0 And varY <> 0 Then 'on est vertical
                  CoeffMC(Nbr_Out).Sup45deg = True
                  CoeffMC(Nbr_Out).C = cov / varY
                  CoeffMC(Nbr_Out).D = Xmoy - CoeffMC(Nbr_Out).C * Ymoy
               ElseIf varX <> 0 And varY = 0 Then 'on est horizontal
                  CoeffMC(Nbr_Out).Sup45deg = False
                  CoeffMC(Nbr_Out).A = cov / varX
                  CoeffMC(Nbr_Out).B = Ymoy - CoeffMC(Nbr_Out).A * Xmoy
               ElseIf varX = 0 And varY = 0 Then 'on fait des ondulations équivalentes de part et d'autre d'une position centrale
                  MsgBox "Erreur : varX et varY sont nulles"
               End If
            End If
'         End If
         Nbr_Out = Nbr_Out + 1
      End If
      i = i + j - 1
   Loop While (i <= UBound(profil) - 2)
   'on recopie la première dans la dernière
   ReDim Preserve CoeffMC(0 To Nbr_Out)
   CoeffMC(Nbr_Out) = CoeffMC(0)
   'le premier et le dernier segment sont les seuls qui sont susceptibles d'être parallèles, il faut traiter le cas
   Xdernier = profil(UBound(profil)).x
   Ydernier = profil(UBound(profil)).y
   ReDim Preserve profil(0 To 0) 'on garde le premier point, on le remplace seulement s'il y a une intersection (droites non parallèles)
   For i = 1 To Nbr_Out
      ReDim Preserve profil(0 To i)
      If CoeffMC(i).Sup45deg = False And CoeffMC(i - 1).Sup45deg = False Then
         If CoeffMC(i - 1).A = CoeffMC(i).A Then 'même pente, division par zéro
            profil(i).x = CoeffMC(i).x
            profil(i).y = CoeffMC(i).y
         Else
            profil(i).x = (CoeffMC(i).B - CoeffMC(i - 1).B) / (CoeffMC(i - 1).A - CoeffMC(i).A)
            profil(i).y = CoeffMC(i).A * profil(i).x + CoeffMC(i).B
         End If
      ElseIf CoeffMC(i).Sup45deg = True And CoeffMC(i - 1).Sup45deg = True Then
         If CoeffMC(i - 1).C = CoeffMC(i).C Then 'même pente, division par zéro
            profil(i).x = CoeffMC(i).x
            profil(i).y = CoeffMC(i).y
         Else
            profil(i).y = (CoeffMC(i).D - CoeffMC(i - 1).D) / (CoeffMC(i - 1).C - CoeffMC(i).C)
            profil(i).x = CoeffMC(i).C * profil(i).y + CoeffMC(i).D
         End If
      ElseIf CoeffMC(i).Sup45deg = False And CoeffMC(i - 1).Sup45deg = True Then   'i en A,B et i-1 en C,D
         If CoeffMC(i).A = 0 Then
            profil(i).y = CoeffMC(i).B
            profil(i).x = CoeffMC(i - 1).C * profil(i).y + CoeffMC(i - 1).D
         Else
            If CoeffMC(i - 1).C = 0 Then
               profil(i).x = CoeffMC(i - 1).D
               profil(i).y = CoeffMC(i).A * profil(i).x + CoeffMC(i).B
            ElseIf 1 - CoeffMC(i).A * CoeffMC(i - 1).C <> 0 Then
               profil(i).y = (CoeffMC(i).A * CoeffMC(i - 1).D + CoeffMC(i).B) / (1 - CoeffMC(i).A * CoeffMC(i - 1).C)
               profil(i).x = CoeffMC(i - 1).C * profil(i).y + CoeffMC(i - 1).D
            Else
               profil(i).x = CoeffMC(i).x
               profil(i).y = CoeffMC(i).y
            End If
         End If
      ElseIf CoeffMC(i).Sup45deg = True And CoeffMC(i - 1).Sup45deg = False Then 'i-1 en A,B et i en C,D
         If CoeffMC(i - 1).A = 0 Then
            profil(i).y = CoeffMC(i - 1).B
            profil(i).x = CoeffMC(i).C * profil(i).y + CoeffMC(i).D
         Else
            If CoeffMC(i).C = 0 Then
               profil(i).x = CoeffMC(i).D
               profil(i).y = CoeffMC(i - 1).A * profil(i).x + CoeffMC(i - 1).B
            ElseIf 1 - CoeffMC(i - 1).A * CoeffMC(i).C <> 0 Then
               profil(i).y = (CoeffMC(i - 1).A * CoeffMC(i).D + CoeffMC(i - 1).B) / (1 - CoeffMC(i - 1).A * CoeffMC(i).C)
               profil(i).x = CoeffMC(i).C * profil(i).y + CoeffMC(i).D
            Else
               profil(i).x = CoeffMC(i).x
               profil(i).y = CoeffMC(i).y
            End If
         End If
      End If
      If (profil(i).y - CoeffMC(i - 1).y) * (profil(i).y - CoeffMC(i - 1).y) + (profil(i).x - CoeffMC(i - 1).x) * (profil(i).x - CoeffMC(i - 1).x) > 2 * epsilon Then
         profil(i).x = CoeffMC(i - 1).x
         profil(i).y = CoeffMC(i - 1).y
      End If
   
   Next i
   ReDim Preserve profil(0 To Nbr_Out)
   profil(0) = profil(Nbr_Out)
   NettoyageMoindreCarres = Nbr_Out + 1
End Function

Private Sub cmdQuitterImpConv_Click()
   Unload Me
End Sub

Private Sub menuColler_Click()  'menu contextuel "coller"
   cmdImporterImage(1).Value = True
End Sub

Private Sub menuImporter_Click()   'menu contextuel "importer"
   cmdImporterImage(0).Value = True
End Sub

'Private Function MarquerNettoyageEtAngles(ByVal Dist As Single, ByRef profil() As PointOriente) As Long 'renvoie le nombre de points
''marquer les tronçons à lisser
'   Dim i As Long, j As Long, k As Long
'   Dim lined As Long
'   Dim Nbr_Mark As Long
'   Dim A As Single, B As Single
'   Dim d2 As Single, d1 As Single
'   Dim epsilon As Single
'   Dim ProfilTemp() As PointOriente
'   Dim Angle1 As Single, Angle2 As Single
'
'   'on utilise un tableau de travail pour récupérer les indices
'   ReDim ProfilTemp(0 To UBound(profil) - 1)
'   For i = 0 To UBound(profil) - 1
'      ProfilTemp(i) = profil(i + 1)
'   Next i
'
'   'initialisation des marqueurs
'   For i = 0 To UBound(ProfilTemp)
'      ProfilTemp(i).Mark = False
'   Next i
'
'   'moins de 3 points
'   If UBound(ProfilTemp) < 3 Then
'      For i = 0 To UBound(ProfilTemp)
'         ProfilTemp(i).Mark = True
'      Next i
'      MarquerNettoyageEtAngles = UBound(ProfilTemp) + 1
'      Exit Function
'   End If
'
'   'cas courant
'   epsilon = Dist * Dist  'On élève au carré pour aller plus vite pour les comparaisons
'   i = 0
'   Nbr_Mark = 1
'   ProfilTemp(0).Mark = True
'   Do
'      j = 1
'      Do
'         lined = 0
'         j = j + 1
'         For k = 1 To j - 1
'            'Est-ce que tous les points intermédiaires sont alignés ?
'            A = ProfilTemp(i + j).x - ProfilTemp(i).x
'            B = ProfilTemp(i + j).y - ProfilTemp(i).y
'            d2 = A * A + B * B
'            If d2 <= epsilon Then
'               ' Les points i and i+j sont trop près
'               d1 = (ProfilTemp(i + k).x - ProfilTemp(i).x) * (ProfilTemp(i + k).x - ProfilTemp(i).x) + _
'                  (ProfilTemp(i + k).y - ProfilTemp(i).y) * (ProfilTemp(i + k).y - ProfilTemp(i).y)
'               If d1 <= epsilon Then
'                  lined = lined + 1         ' Le point i+k est aussi trop près
'               Else
'                  Exit For                  ' Le point i+k n'est pas aligné, pas la peine de continuer !
'               End If
'            Else
'               d1 = ((ProfilTemp(i + k).x - ProfilTemp(i).x) * A + (ProfilTemp(i + k).y - ProfilTemp(i).y) * B) / d2
'               If d1 < 0# Then
'                  d1 = 0#                   ' La référence est le point i
'               Else
'                  If d1 > 1# Then d1 = 1#   ' La référence est le point i+j
'               End If
'               A = ProfilTemp(i).x + d1 * A - ProfilTemp(i + k).x
'               B = ProfilTemp(i).y + d1 * B - ProfilTemp(i + k).y
'               If A * A + B * B <= epsilon Then
'                  lined = lined + 1         ' point aligné
'               Else
'                  Exit For                  ' Le point i+k n'est pas aligné, pas besoin de continuer
'               End If
'            End If
'         Next k
'      Loop While lined = j - 1 And i + j + 1 <= UBound(ProfilTemp) 'Boucle jusqu'à ce qu'un point ne soit pas aligné
'      i = i + j - 1
'      If lined <> j - 1 Then
'        ProfilTemp(i).Mark = True
'        Nbr_Mark = Nbr_Mark + 1
'      End If
'   Loop While (i <= UBound(ProfilTemp) - 2)
'   ProfilTemp(UBound(ProfilTemp)).Mark = True
'   Nbr_Mark = Nbr_Mark + 1
'   MarquerNettoyageEtAngles = Nbr_Mark
'
'   'on récupère les marqueurs depuis le tableau de travail et on stocke les numéros des marqueurs pour calcul de l'angle
'   ReDim NumeroMarqueurs(1 To Nbr_Mark)
'   j = 1
'   For i = 0 To UBound(ProfilTemp)
'      profil(i + 1).Mark = ProfilTemp(i).Mark
'      If profil(i + 1).Mark = True Then
'         NumeroMarqueurs(j) = i + 1
'         j = j + 1
'      End If
'   Next i
'
'   'sélection auto par angle supérieur à 30°
'   'on mesure les angles par rapport à X et on fait la soustraction
'   For i = 1 To Nbr_Mark - 2
'      Angle1 = Angle_Segment(profil(NumeroMarqueurs(i)).x, profil(NumeroMarqueurs(i)).y, profil(NumeroMarqueurs(i + 1)).x, profil(NumeroMarqueurs(i + 1)).y) * 180 / pi
'      Angle2 = Angle_Segment(profil(NumeroMarqueurs(i + 1)).x, profil(NumeroMarqueurs(i + 1)).y, profil(NumeroMarqueurs(i + 2)).x, profil(NumeroMarqueurs(i + 2)).y) * 180 / pi
'      If Not (Abs(Angle2 - Angle1) >= 30 And Abs(Angle2 - Angle1) <= 330) Then
'         profil(NumeroMarqueurs(i + 1)).Mark = False
'         Nbr_Mark = Nbr_Mark - 1
'      End If
'   Next i
'
'   'on récupère uniquement les marqueurs True
'   ReDim NumeroMarqueurs(1 To Nbr_Mark)
'   j = 1
'   For i = 1 To UBound(profil)
'      If profil(i).Mark = True Then
'         NumeroMarqueurs(j) = i
'         j = j + 1
'      End If
'   Next i
'
'   MarquerNettoyageEtAngles = Nbr_Mark
'
'End Function

Private Sub RepresenterVecto()
   Dim XminSequVecto As Single, XmaxSequVecto As Single
   Dim YminSequVecto As Single, YmaxSequVecto As Single
   Dim XcentreSequVecto As Single, YcentreSequVecto As Single
   Dim CoeffSequVecto As Single
   
   'ajustement de la fenêtre de visualisation de la vectorisation
   If pctCouleurRedim.Height < HauteurMaxiPct And pctCouleurRedim.Width < LargeurMaxiPct Then
      'on agrandit
      If pctCouleurOriginale.Width / pctCouleurOriginale.Height >= LargeurMaxiPct / HauteurMaxiPct Then
         pctVisuVecto.Width = LargeurMaxiPct
         pctVisuVecto.Height = pctVisuVecto.Width * pctCouleurOriginale.Height / pctCouleurOriginale.Width
      Else
         pctVisuVecto.Height = HauteurMaxiPct
         pctVisuVecto.Width = pctVisuVecto.Height * pctCouleurOriginale.Width / pctCouleurOriginale.Height
      End If
   Else
      pctVisuVecto.Height = pctCouleurRedim.Height
      pctVisuVecto.Width = pctCouleurRedim.Width
   End If
   pctVisuVecto.ScaleTop = -pctVisuVecto.ScaleHeight      'recalage de l'axe Y
   pctVisuVecto.Visible = True

   '*****************
   'calcul préalables
   '*****************
   Dim i As Long, j As Long
   Dim XFS1 As Single, YFS1 As Single, XFS2 As Single, YFS2 As Single, XFS As Single, YFS As Single

   If NbSequVecto > 0 Then
      '*** calcul des maxi et mini du total des séquences pour représentation graphique ***
      XminSequVecto = SequVecto(1).Point(1).x
      XmaxSequVecto = SequVecto(1).Point(1).x
      YminSequVecto = SequVecto(1).Point(1).y
      YmaxSequVecto = SequVecto(1).Point(1).y
      For i = 1 To NbSequVecto
         For j = 1 To SequVecto(i).NbPoints
            If SequVecto(i).Point(j).x < XminSequVecto Then XminSequVecto = SequVecto(i).Point(j).x
            If SequVecto(i).Point(j).x > XmaxSequVecto Then XmaxSequVecto = SequVecto(i).Point(j).x
            If SequVecto(i).Point(j).y < YminSequVecto Then YminSequVecto = SequVecto(i).Point(j).y
            If SequVecto(i).Point(j).y > YmaxSequVecto Then YmaxSequVecto = SequVecto(i).Point(j).y
         Next j
      Next i
      XcentreSequVecto = (XmaxSequVecto + XminSequVecto) / 2
      YcentreSequVecto = (YmaxSequVecto + YminSequVecto) / 2
      If XmaxSequVecto = XminSequVecto And YmaxSequVecto = YminSequVecto Then    'tous les points au même point!
         CoeffSequVecto = 1
      ElseIf XmaxSequVecto = XminSequVecto Then    'tous les points sur une droite verticale
         CoeffSequVecto = (pctVisuVecto.Height - 2 * MargInit) / (Abs(YmaxSequVecto - YminSequVecto))
      ElseIf YmaxSequVecto = YminSequVecto Then    'tous les points sur une droite horizontale
         CoeffSequVecto = (pctVisuVecto.Width - 2 * MargInit) / (Abs(XmaxSequVecto - XminSequVecto))
      Else     'cas classique
         If (pctVisuVecto.Width - 2 * MargInit) / (Abs(XmaxSequVecto - XminSequVecto)) < (pctVisuVecto.Height - 2 * MargInit) / (Abs(YmaxSequVecto - YminSequVecto)) Then
            CoeffSequVecto = (pctVisuVecto.Width - 2 * MargInit) / (Abs(XmaxSequVecto - XminSequVecto))
         Else
            CoeffSequVecto = (pctVisuVecto.Height - 2 * MargInit) / (Abs(YmaxSequVecto - YminSequVecto))
         End If
      End If
     
      '*** Création du tableau de représentation ***
      SequVectoTrace = SequVecto
      For i = 1 To NbSequVecto
         With SequVectoTrace(i)
            For j = 1 To .NbPoints
               .Point(j).x = (SequVecto(i).Point(j).x - XcentreSequVecto) * CoeffSequVecto + pctVisuVecto.Width / 2
               .Point(j).y = (SequVecto(i).Point(j).y - YcentreSequVecto) * CoeffSequVecto + pctVisuVecto.Height / 2
            Next j
            Call MaxiMiniSequ(SequVectoTrace(i))   'calcul des maxi et mini de toutes les séquences tracées
         End With
      Next i
   End If
   '*****************
   '  tracé
   '*****************
   
   If NbSequVecto > 0 Then
      pctVisuVecto.AutoRedraw = True   'le tracé des séquences se fait sur le plan permanent
      pctVisuVecto.Cls
      'séquences initiales
      For i = 1 To NbSequVecto
         If SequVecto(i).NbPoints = 1 Then  'gestion des séquences constituées d'un point unique pour passage du fil
            With SequVectoTrace(i)
               pctVisuVecto.Line (.Point(1).x - 4, .Point(1).y - 4)-(.Point(1).x + 4, .Point(1).y + 4), vbBlack
               pctVisuVecto.Line (.Point(1).x - 4, .Point(1).y + 4)-(.Point(1).x + 4, .Point(1).y - 4), vbBlack
            End With
         Else           'gestion des séquences consituées de plusieurs points
            With SequVectoTrace(i)
               For j = 1 To SequVecto(i).NbPoints - 1
                  pctVisuVecto.Line (.Point(j).x, .Point(j).y)-(.Point(j + 1).x, .Point(j + 1).y), vbBlack 'segments
               Next j
               If chkVoirPointsVecto.Value = vbChecked Then
                  For j = 1 To SequVecto(i).NbPoints - 1
                     pctVisuVecto.Circle (.Point(j).x, .Point(j).y), 3, vbRed  'cercles et premier point
                  Next j
                  pctVisuVecto.Circle (.Point(.NbPoints).x, .Point(.NbPoints).y), 3, vbRed
                  XFS1 = .Point(1).x
                  YFS1 = .Point(1).y
                  XFS2 = .Point(2).x
                  YFS2 = .Point(2).y
                  pctVisuVecto.FillColor = RGB(255, 107, 107)
                  pctVisuVecto.FillStyle = 0  'rempli
                  pctVisuVecto.DrawMode = vbMaskPen
                  'tracé de la flèche
                  If YFS2 <> YFS1 Or XFS2 <> XFS1 Then
                     XFS = 10 * (XFS2 - XFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     YFS = 10 * (YFS2 - YFS1) / Sqr((YFS2 - YFS1) ^ 2 + (XFS2 - XFS1) ^ 2)
                     pctVisuVecto.DrawWidth = 2
                     pctVisuVecto.Line (XFS1, YFS1)-Step(XFS, YFS), RGB(230, 61, 61) ' corps de la flèche
                     pctVisuVecto.DrawWidth = 1
                  End If
                  pctVisuVecto.DrawMode = vbCopyPen
                  'tracé du point d'entrée
                  pctVisuVecto.FillStyle = 0 'rempli
                  pctVisuVecto.FillColor = vbRed
                  pctVisuVecto.Circle (XFS1, YFS1), 2, vbRed
                  pctVisuVecto.FillStyle = 1 'transparent
               End If
            End With
         End If
      Next i
      pctVisuVecto.AutoRedraw = False   'par défaut on est sur le plan temporaire
   End If
End Sub

Private Sub chkVoirPointsVecto_Click()
   If NbSequVecto > 0 Then
      Call RepresenterVecto
   End If
End Sub

Private Sub cmdSauverVecto_Click()
   Dim ExportProfil() As ProfilDATDXF
   Dim NomFichier As String
   Dim RepertoireSauvegarde As String
   Dim Reponse As Integer
   Dim Comment As String
   Dim CodeErreur As Long
   Dim i As Long, j As Long, k As Long
   Dim N As Node
   
   RepertoireSauvegarde = LitFichierIni("Fichiers", "DernierRepertoire", App.Path)  'si la clé n'existe pas, on prend App.path

   COMDLG.DialogTitle = "Enregistrer"
   COMDLG.CancelError = True
'*** DXF
   COMDLG.Filter = "Contours vectorisés (*.dxf)|*.dxf"
   COMDLG.FLAGS = cdlOFNOverwritePrompt 'message pour écrasement de fichier
   COMDLG.FileName = ""
   COMDLG.InitDir = RepertoireSauvegarde
   COMDLG.FilterIndex = 1
   On Error GoTo Annuler                     'si fermeture de la fenêtre sans sélection de fichier
   COMDLG.ShowSave                           ' afficher la fenêtre de sauvegarde
   On Error GoTo 0   'on annule la gestion d'erreur pour voir les plantages
   NomFichier = COMDLG.FileName  'nom entré par l'utilisateur (chemin complet)
   If NbSequVecto = 0 Then 'on ne peut pas exporter un trajet vide
      Exit Sub
   End If
   'mise à l'échelle de la machine pour que ce ne soit pas trop grand quand on l'importe
   Call MiseEchelleMachine
   
   ReDim ExportProfil(1 To 1)
   k = 0
   For i = 1 To NbSequVecto
      With SequVecto(i)
         For j = 1 To .NbPoints
            k = k + 1
            ReDim Preserve ExportProfil(1 To k)
            ExportProfil(k).x = .Point(j).x
            ExportProfil(k).y = .Point(j).y
            ExportProfil(k).NumSequ = i
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
      
   'mémorisation du dernier dossier utilisé dans MiniCut2d_Software.ini
   RepertoireSauvegarde = GetPathName(NomFichier)
   EcritFichierIni "Fichiers", "DernierRepertoire", RepertoireSauvegarde

   'il faut ajouter le fichier créé dans le treeview, mais seulement s'il ne s'agit pas d'un écrasement
   Call frmMiniCut2d.CreerTreeView
   For Each N In frmMiniCut2d.Tree.Nodes
      If N.Key = NomFichier Then
         N.Selected = True
         N.Expanded = True
      End If
   Next
   Exit Sub
Annuler:
'si pas de fichier enregistré
End Sub

Private Sub MiseEchelleMachine()
   Dim XminTotal As Single, YminTotal As Single, XmaxTotal As Single, YmaxTotal As Single
   Dim CoeffEchelle As Single
   Dim i As Long, j As Long
   
   If flagVectoRedimCourses = False Then
      XminTotal = SequVecto(1).Xmin
      XmaxTotal = SequVecto(1).Xmax
      YminTotal = SequVecto(1).Ymin
      YmaxTotal = SequVecto(1).Ymax
      For i = 1 To NbSequVecto
         If SequVecto(i).Xmin < XminTotal Then XminTotal = SequVecto(i).Xmin
         If SequVecto(i).Xmax > XmaxTotal Then XmaxTotal = SequVecto(i).Xmax
         If SequVecto(i).Ymin < YminTotal Then YminTotal = SequVecto(i).Ymin
         If SequVecto(i).Ymax > YmaxTotal Then YmaxTotal = SequVecto(i).Ymax
      Next i
      If Abs(YmaxTotal - YminTotal) <> 0 Then
         If Abs(XmaxTotal - XminTotal) / Abs(YmaxTotal - YminTotal) > CourseX / CourseY Then  'la valeur en X est "maître"
            CoeffEchelle = (CourseX - 14) / Abs(XmaxTotal - XminTotal) 'on laisse une marge de 7mm à droite et à gauche
         Else
            CoeffEchelle = (CourseY - 14) / Abs(YmaxTotal - YminTotal) 'on laisse une marge de 7mm en haut et en bas
         End If
      Else 'l'image est une ligne horizontale
         CoeffEchelle = (CourseX - 14) / Abs(XmaxTotal - XminTotal) 'on laisse une marge de 7mm autour
      End If
      For i = 1 To NbSequVecto
         With SequVecto(i)
            For j = 1 To .NbPoints
               .Point(j).x = .Point(j).x * CoeffEchelle
               .Point(j).y = .Point(j).y * CoeffEchelle
            Next j
         End With
      Next i
      flagVectoRedimCourses = True
   End If
End Sub
