Attribute VB_Name = "modTypesDeclarations"
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

'Types pour conversion d'image
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type Pixel
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

'type pour la s�lection par rectangle dans l'image
Public Type ShapeRectangle
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

'type pour droite des moindres carr�s
Public Type MoindreCarres
   A As Single
   B As Single
   C As Single
   D As Single
   x As Single
   y As Single
   Sup45deg As Boolean
End Type
   
'les points pour la d�tection des contours
Public Type PointOriente
   x As Single
   y As Single
   D As Integer
   Mark As Boolean
End Type

Public Type Couple
   Point1 As PointOriente
   Point2 As PointOriente
End Type

Public Type Contour
   Point() As PointOriente
   NombrePoints As Long
   Couples() As Couple
   Xmin As Single
   Xmax As Single
   Ymin As Single
   Ymax As Single
End Type

' Type pour recherche des dossiers sp�ciaux (Bureau, Mes Documents...)
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

'Type pour gestion du browser de la biblioth�que lors de l'importation d'un fichier
Type BROWSEINFO                                                          '*
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String                                                       '*
  lpszTitle As String
  ulFlags As Long
  lpfn As Long                                                                   '*
  lParam As Long
  iImage As Long
End Type                                                                         '*

'Type pour r�cup�ration d'un profil avec CNCTools.dll
Type ProfilDATDXF
    x As Single
    y As Single
    Cde As String * 2           'Commande (PU ou PD)
    NumSequ As String * 2   'N� de s�quence de la d�coupe
End Type

'point de profil
Public Type PointProfil
   x As Single
   y As Single
   Etat As Integer
   Vitesse As Single
   Acceleration As Boolean 'true=acc�l�ration sur le segment suivant
   Mark As Boolean
End Type

'Point simple (calcul de la simulation)
Public Type PointSimple
   x As Single
   y As Single
End Type

Public Type POINTAPI    'utilis� pour le dessin de la table avec l'API Polygon
   x As Long
   y As Long
End Type

'Segment (calcul des longueurs pour la simulation)
Public Type Segment
   Point1 As PointSimple
   Point2 As PointSimple
   Longueur As Single
End Type

'Tableau d'une s�quence
Public Type Sequ
   NbPoints As Long
   Point() As PointProfil
   Xmin As Single
   Xmax As Single
   Ymin As Single
   Ymax As Single
   DeltaX As Single
   DeltaY As Single
   Etat As Integer
End Type

Public Type PointPas
   XPas As Long
   YPas As Long
   Vitesse As Single
   Acceleration As Boolean 'true=acc�l�ration sur le segment suivant
End Type

Public Type SequPas
   NbPoints As Long
   PointPas() As PointPas
End Type

'Entit�s g�om�triques pour repr�sentation
Type Rectangle    'd�fini � partir d'une diagonale
   X1 As Single
   Y1 As Single
   X2 As Single
   Y2 As Single
   Rempli As Boolean    'si oui, on regarde la couleur de fond
   CoulTour As Long
   CoulFond As Long
   TypeTrait As Integer
End Type

Type Ligne
   X1 As Single
   Y1 As Single
   X2 As Single
   Y2 As Single
   Couleur As Long
   TypeTrait As Integer
End Type

'Salve Data pour le fil chaud, le PWM est g�r� � part
Public Type SalveData
   CMD As Byte
   NBRL As Byte
   NBRM As Byte
   NBRH As Byte
   NBRU As Byte
   S1L As Byte
   S1M As Byte
   S1H As Byte
   S1U As Byte
   S2L As Byte
   S2M As Byte
   S2H As Byte
   S2U As Byte
   S3L As Byte
   S3M As Byte
   S3H As Byte
   S3U As Byte
   S4L As Byte
   S4M As Byte
   S4H As Byte
   S4U As Byte
   S5L As Byte
   S5M As Byte
   S5H As Byte
   S5U As Byte
   F_ACC As Byte
   F_DEC As Byte
   DECL As Byte
   DECM As Byte
   DECH As Byte
   DECU As Byte
   TempsTotal As Single 'dur�e du d�placement en s
   CodeErreur As Integer
End Type

Public Type Matiere
   Nom As String
   Chauffe As Single
   Vitesse As String
   IndexInitial As Integer 'num�ro dans le tableau initial
End Type

Public Type UndoRedo
   TransfUndo() As Sequ
   NbTransfUndo As Long
   NbTransfSelUndo As Long
   HautBlocUndo As Single
   LongBlocUndo As Single
   CoeffBlocUndo As Single
   CheckAjusterEchelle As Boolean
   CheckCentrerX As Boolean
   CheckCentrerY As Boolean
End Type

'**************************************
'* D�claration des variables globales *
'**************************************

'polices de caract�res
Public PoliceNormal As New StdFont
Public PoliceGras As New StdFont

Public profil() As ProfilDATDXF  'tableau temporaire pour r�ception des points issu de cnctools.dll
Public Xmin As Single, Ymin As Single, Xmax As Single, Ymax As Single 'pour communication avec cnctools.dll
Public NomFichier As String   'pour Transfmission � CNCTools.dll
Public NbPoints As Long

Public Const pi = 3.14159265358979
Public CoeffNett As Single
Public EpsilonNettoyage As Single

Public NbSequ As Long
Public NbSequSel As Long
Public NbTransf As Long
Public NbTransfSel As Long
Public NumeroSequ As Long
Public NumeroTransf As Long
Public MargInit As Integer
Public NumSequSel As Long
Public NumTransfSel As Long

Public DeltaXMax As Single
Public DeltaYMax As Single
Public DeltaXMaxTransf As Single
Public DeltaYMaxTransf As Single

Public Numero As String    'liste alphab�tique des s�quences
Public Sequ() As Sequ 'pour choix dans dxf
Public SequTrace() As Sequ 'pour trac� et s�lection par clic

Public Transf() As Sequ
Public XminTotalTransf As Integer, XmaxTotalTransf As Integer
Public YminTotalTransf As Integer, YmaxTotalTransf As Integer

Public TransfTemp As Sequ   'pour op�rations de mise � l'�chelle

'UNDO/REDO
Public UndoRedo() As UndoRedo   'pour op�ration de Undo
Public IndexUndo As Single
Public flagMemoUndoDansTraceTransf As Boolean

'Onglet D�coupe
Public SequDecoupe As Sequ  's�quence unique d'affichage du trajet du fil
Public SequDecalee As Sequ   's�quence unique d�cal�e
Public SequEntree As Sequ  'trajet d'entr�e jusqu'� la d�coupe
Public SequEntreeDecalee As Sequ
Public SequSortie As Sequ  'trajet de sortie depuis la d�coupe
Public SequSortieDecalee As Sequ
Public SequMouvement As Sequ  's�quence � envoyer � l'interpolateur
Public SequMouvementPas As SequPas 'la m�me, en pas
Public PointsSimulation() As PointSimple 'liste des points de la simulation
Public ProgressionDuFilm As Long 'nombre d'images qui ont d�fil�
Public flagTracePourSimulation As Boolean 'pour savoir s'il faut tracer la d�coupe sur le plan permanent
Public flagSimulationLancee As Boolean 'pour gestion du bouton de simulation

'dimensions du bloc
Public HautBloc As Single  'dimension utile
Public LongBloc As Single

'Marge entre le fil et le bloc
Public MargeFil As Single
'Marge entre le fil et la table
Public MargePlateau As Single

'Marge auto entre le bord du bloc et le trac�
Public MargeInterieureX As Single
Public MargeInterieureY As Single

' repr�sentations graphiques :
Public XminSequ As Single
Public XmaxSequ As Single
Public YminSequ As Single
Public YmaxSequ As Single
Public XcentreSequ As Single
Public YcentreSequ As Single
Public CoeffSequ As Single
Public CoeffBloc As Single 'adaptation � la taille du bloc
Public XminTransf As Single
Public XmaxTransf As Single
Public YminTransf As Single
Public YmaxTransf As Single
Public XcentreTransff As Single
Public YcentreTransff As Single
Public CoeffTransf As Single
Public LongBox As Single
Public HautBox As Single
Public LongBoxTransf As Single
Public HautBoxTransf As Single
Public X0Box As Single
Public Y0Box As Single
Public flagLeProjetAUnNom As Boolean  'pour gestion sauvegarde

Public flagFenetreAgrandie As Boolean '�tat de la fen�tre du bas

Public NoeudX As Node            'objet permettant de cr�er les noeuds du treeview
Public fso As FileSystemObject, dossier As Folder, sousdossier As Folder, fichier As File

Public Jonction As String 'identificateur du type de jonction entre les s�quences d�j� transf�r�es et la s�quence drop�e

'index dex outils
Public Const Deplacer As Integer = 0
Public Const Tourner As Integer = 1
Public Const Etirer As Integer = 2
Public Const Inverser As Integer = 3
Public Const Mesurer As Integer = 4
Public Const CouperProfil As Integer = 5
Public Const PointNumero1 As Integer = 6
Public OutilEnCours As Integer

'gestion selection et outils
Public X1SelTransf As Single
Public Y1SelTransf As Single
Public X1Transf As Single
Public Y1Transf As Single
Public MemoX1Transf As Single
Public MemoY1Transf As Single
Public MemoNumTransfMouseDown As Integer
Public MemoNumTransfMouseUp As Integer
Public MemoMouseDownX As Single
Public MemoMouseDownY As Single
Public flagZeroSelectionAuMouseDown As Boolean
Public MemoMouseDownSequX As Single
Public flagInitSelection As Boolean    'pour initialisation des textbox manuelles
Public TransfSousCur As Integer  'num�ro de la s�quence sous le curseur
Public MemoTransfSousCur As Integer

'outil d�placement
Public MemoX As Single, MemoY As Single
Public VecteurX As Single, VecteurY As Single
Public PositionX As Single, PositionY As Single  'position absolue pour m�morisation de la valeur

'pour l'outil rotation :
Public Xcentre As Single
Public Ycentre As Single
Public XminSel As Single
Public XmaxSel As Single
Public YminSel As Single
Public YmaxSel As Single
Public Xtemp As Single, Ytemp As Single
Public angle As Single
Public DemiLargeurSelection As Single
Public DemiHauteurSelection As Single
Public LargeurSelection As Single
Public HauteurSelection As Single

Public flagRectSel As Boolean
Public AngleTotal As Single  'pour outil rotation
Public AngleInitial As Single
Public alpha As Single
Public Rotation As Single  'rotation absolue pour m�morisation
Public AngleRelatif As Single    'variable interm�diaire

Public flagPremierPoint As Boolean  'pour savoir si on est au d�but ou a la fin d'une mesure
Public XPremierPoint As Single
Public YPremierPoint As Single
Public XDernierPoint As Single
Public YDernierPoint As Single
Public XPointMesure As Single
Public YPointMesure As Single
Public ValeurMesure As Single
Public flagPointSousCur As Boolean
Public Xmesure As Single
Public Ymesure As Single
Public Poignee As String


Public NumTransfDecoupe As Long 'outil de d�coupe de profil
Public NumPointDecoupe As Long

'outil �tirer
Public ValeurFinaleX As Single
Public ValeurFinaleY As Single
Public ValeurInitialeX As Single
Public ValeurInitialeY As Single
Public OrigineX As Single
Public OrigineY As Single
Public EtirerXetY As Single
Public kX As Single 'coefficient d'�tirement
Public kY As Single
Public AgrandissementX As Single, AgrandissementY As Single 'valeur absolue de l'�tirement pour m�morisation
Public DemiCarre As Single   'demi-c�t� des poign�es de l'outil �tirer
Public DemiCarreSelec As Single  'demi-cot� du carr� affich� pour s�lection de la poign�e

'********************************************************************
'** D�claration des variables pour la repr�sentation de la d�coupe **
'********************************************************************
'repr�sentation graphique
Public RECT() As Rectangle
Public Ligne() As Ligne
Public NbRect As Integer
Public NbLignes As Integer
Public CourseX As Single
Public CourseY As Single
Public MaxiDecoupeX As Single, MiniDecoupeX As Single
Public MaxiDecoupeY As Single, MiniDecoupeY As Single

Public RectBloc As Rectangle  'rectangle du bloc

Public flagDepassementDecoupe  'si la d�coupe est plus grande que la surface utile
Public NbClignotements As Single

Public MatieresDeLaBase() As Matiere  'tableau des mati�res
Public MatieresAffichees() As Matiere  'tableau des mati�res dans le combobox
Public MatiereUtilisee As Matiere
Public VitesseDecoupe As Single
Public VitesseRapide As Single
Public ChauffeCourante As Single
Public ChauffeMaxi As Single
Public ChauffeCouranteSur255 As Integer

'variables interpolateur

Public Const EnCours As Byte = 1, Arretee As Byte = 0
Public Const EnAttente As Byte = 1, PasDeStop As Byte = 0
Public Const USB As Byte = 1, Flash As Byte = 0
Public Const Marche As Byte = 1, Arret As Byte = 0
Public Const Oui As Byte = 1, Non As Byte = 0
Public Const Manuel As Byte = 1, Auto As Byte = 0
Public Const Ouvert As Byte = 1, Ferme As Byte = 0
Public Const PWMOff As Byte = 1, PWMOn As Byte = 0
Public Const Plus5V As Byte = 1, GND As Byte = 0
Public Const Relache As Byte = 1, Presse As Byte = 0
Public Const Absent As Byte = 1, Present As Byte = 0
Public Const K_Pressee As Byte = 1, K_Relachee As Byte = 0
Public Const Active As Byte = 0, Stoppee As Byte = 1
Public Const InterFDC As Byte = 1, InterOrigine As Byte = 2, ProgStopBouton As Byte = 4, EndOfDatas As Byte = 8


' Gestion des erreurs retourn�es par l'interpolateur :
Public ErrIPL As Long 'code d'erreur renvoy� par la fonction IPL5X_Send
Public CodeErrDepl As Long  'voir proc�dure de d�placement dans le module modDeplacements
                           'et fonction de d�codage

' Tableau des octets de conversation avec l'interpolateur
Public ByteIPL(0 To 64) As Byte   'tableau des bytes � envoyer � IPL5X et de la r�ponse
                                       'mettre des z�ro dans les cases non utilis�es
                                       'le premier octet doit �tre � 0 ByteIPL(0)=0
                                       'l'instruction commence � ByteIPL(1)

Public TableLIN() As Single   'Tableaux servant au calcul des acc�l�rations d'IPL5X (cf. feuille Excell)
                              'ce tableau est cr�� � l'ouverture de l'appli
'****** Caract�ristiques de la table en cours ******
Public NomTable As String
Public Frequence As Long
Public PenteAcceleration As Byte
Public PWMmaxi As Byte
Public DelaiMoteursOff As Byte
Public VMaxSansAcc As Single
Public VMaxAvecAcc As Single
Public MmParTourXG As Single  'r�duction m�canique
Public MmParTourYG As Single
Public MmParTourXD As Single
Public MmParTourYD As Single
Public PasParTourMoteurXG As Single  'caract�ristique moteur
Public PasParTourMoteurYG As Single
Public PasParTourMoteurXD As Single
Public PasParTourMoteurYD As Single
Public MicroPasXG As Single   'r�glage interface micropas
Public MicroPasYG As Single
Public MicroPasXD As Single
Public MicroPasYD As Single
Public PasParTourXG As Long   'r�sultat final : nombre de pas � envoyer pour faire un tour
Public PasParTourYG As Long
Public PasParTourXD As Long
Public PasParTourYD As Long
Public NbrPasToOriXG As Long   'd�calage de l'origine par rapport aux inters
Public NbrPasToOriYG As Long
Public NbrPasToOriXD As Long
Public NbrPasToOriYD As Long
Public MmToOriXG As Single
Public MmToOriYG As Single
Public MmToOriXD As Single
Public MmToOriYD As Single
Public TypeAxeInterpolateur(1 To 5) As Byte
Public DefinitionSortie(1 To 10) As Byte

'pr�sence / absence d'une table compatible = d�coupe active/inactive
Public flagTableEcriteDansIPL As Boolean

'tableau des salves de la d�coupe
Public SalveDecoupe() As SalveData

Public PasArretX As Long
Public PasArretY As Long

Public NBRt As Long  'nombre total de pulses pr�vues pour le segment (appelation Dev_Guide)
Public NBRc As Long  'nombre de pulses perdues sur le segment

Public PasParcourusXG As Long
Public PasParcourusYG As Long
Public PasParcourusXD As Long
Public PasParcourusYD As Long

Public flagRepriseDecoupe As Boolean
Public flagAnnulationDecoupeEnCours As Boolean

Public TempoChauffeFil As Integer
Public flagAppuiSTOP As Boolean
Public flagSTOPAvantGoBuffer As Boolean
Public SegmentCourant As Long
Public flagModifChauffePendantDecoupe As Boolean

'******** Choix utilisateur avant une d�coupe ********
Public flagAnnulationDemandeDecoupe As Boolean

Public TexteDuree As String

Public ChemRepRacine As String  'racine du treeview
Public flagAfficherSens As Boolean
Public flagPositionPliage As Boolean

'******* Traduction ********
Public strLangue As String    'langue de traduction
Public Message() As String
Public Const Corps = 1  'pour msgbox
Public Const Titre = 2
Public Label() As String

'******* Pour ne pas avoir � valider deux fois le pliage ********
Public flagPliageApresMsgBox As Boolean

'******* Pour pouvoir effacer un fichier *********
Public CheminFichier As String

'******* Zoom � la molette apr�s clic droit *******
Public XZoomMolette As Long
Public YZoomMolette As Long

'******* Mouvements manuels *******
Public InfiniX As Long
Public InfiniY As Long
Public flagChangementDirectionPendantMouvement As Boolean
Public flagAppuiFeuRouge As Boolean
Public flagAppuiStopSansMsgBox As Boolean
'******* Type de machine ********
Public TypeMachine As String
Public flagChoixFait As Boolean

'******* Importation d'image *******
Public LargeurMaxiPct As Long
Public HauteurMaxiPct As Long
Public BitsPict() As Byte
Public MatricePict() As Pixel
Public Size As Long
Public ToleranceNoir As Byte
Public MatriceBinaire() As Byte
Public CopieMatriceBinaire() As Byte
Public Contours() As Contour
Public NombreContours As Integer
Public RectangleSelection As ShapeRectangle
Public blnSelected As Boolean
Public OrigX As Integer
Public OrigY As Integer
Public blnImageTracee As Boolean
Public CoeffRedim As Single
Public blnRecadree As Boolean
Public HauteurPictDestination As Long
Public LargeurPictDestination As Long

'********* s�quences vectoris�es *********
Public NbSequVecto As Long
Public NbSequVectoSel As Long
Public SequVecto() As Sequ
Public SequVectoTrace() As Sequ
Public flagVectoRedimCourses As Boolean

'********* Pixel Pour Detection Point ********
Public UnPixelToMm As Single

'marqueurs de lissage
Public NumeroMarqueurs() As Long

'********* Mode Expert **********
Public ModeSoft As String
Public ModeSoftTemp As String 'pour check � l'ouverture de la fen�tre de settings
Public flagFenetreChargee As Boolean
