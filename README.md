Le code de ce repository est celui du logiciel MiniCut2d Software, �crit en VB6.

Il est publi� avec les objectifs suivants :
- r��crire le logiciel pour qu'il puisse fonctionner sous Windows, Linux et Mac,
- conserver la simplicit� et la stabilit� du logiciel actuel,
- am�liorer l'aspect visuel du logiciel (et peut-�tre aussi l'ergonomie).

Les sources sont fournies sous licence CECILL (compatible GNU GPL).

Ci-dessous la description des diff�rents modules du logiciel
------------------------------------------------------------

Informations g�n�rales :
- Un fichier projet MiniCut2d (.mnc) contient les informations suivantes :
	* le nombre de s�quences qui constituent le projet,
	* le nombre de point de chaque s�quence,
	* les coordonn�es X:Y de chaque plan dans le rep�re de la machine,
	* les dimensions du bloc de mati�re,
	* le type d'entr�e: 0=par la gauche; 1=par le haut; 2=par la droite,
	* le type de sortie (idem).
	
- Penser tout de suite au multi-langues qui peut �ventuellement �tre g�r� dans des fichiers texte s�par�s (actuellement les traductions sont contenues dans des modules compil�s avec le .exe), mais pr�voir toujours une valeur pas d�faut dans le code en cas d'absence du fichier de langue.
	
Description des diff�rents modules pour MiniCut2d Software (en essayant de pr�voir l'�volution):

1 - Communication avec l'interface de type IPL5X reconnue comme p�riph�rique HID (actuellement ce travail est r�alis� par IPL5XCom.dll sur la base d'un tableau de 65 octets dont la premi�re case est toujours vide, code en C fourni � la demande)
	* Identification du p�riph�rique
	* Envoi d'octets
	* R�ception d'octets
	
2 - Visionneuse de fichiers :
	* Cr�ation d'une biblioth�que � l'installation du logiciel (=> ToDo : settings, choix de l'emplacement?).
	* TreeView de repr�sentation permettant de parcourir la biblioth�que.
	* Fen�tre de visualisation du contenu du fichier.
	* Fichiers trait�s : DXF, DAT, PLT (trait�s actuellement par cnctools.dll dont le code en C peut fourni � la demande), TXT (format : cf. site MiniCut2d), MNC (projet MiniCut2d, voir code), ToDo : SVG.
	* Les profils ferm�s sont r�orient�s/renum�rot�s pour mettre le premier point vers l'origine suivant X (voir code) et la rotation en sens horaire.
	* Les profils issus de DXF sont nettoy�s (suppression des points align�s ou trop proches), voir code et fonction Nettoyage.
	* Op�rations �l�mentairs sur le contenu affich� : affichage du premier point, miroir, changement de sens.

3 - Cr�ation de projet :
	* Fen�tre r�pr�sentant la zone utile de la machine et le bloc.
	* Saisie des dimensions du bloc.
	* Int�gration des s�quences par drag and drop ou double-clic depuis la visionneuse.
	* Affichage des s�quences (en noir) et des jonctions (en vert) entre les s�quences.
	* Affichage dynamique de la jonction en fonction de l'emplacement de la nouvelle s�quence import�e par drag and drop (en rose tant que la s�quence n'est pas pos�e).
	* Zoom molette, d�placement de la vue, zoom bloc.
	* Affichage des points.
	* Alternance des couleurs des s�quences pour les distinguer les unes des autres.
	* Outils de s�lection/d�placement/rotation/�tirement (� la main ou par saisie de valeur), affichage dynamique de la dimension de la s�lection.
	* Outils de modification du premier point, de mesure, de coupe des s�quences, de suppression de la s�lection.
	* Outils d'inversion de sens d'une ou plusieurs s�quences.
	* Outil d'insertion de point.
	* Outil miroir.
	* Outil dupliquer (ToDo : remplacer l'outil duliquer par Copier+Coller).
	* Ajout de point par Drag and Drop.
	* Outils de mise � l'�chelle du bloc et de centrage horizontal et vertical (ToDo : case � cocher pour verrouiller ou non ces boutons).
	* Outils d'alignement (ToDo : r�duire la place prise par ces outils : boutons d�roulants?).
	* ToDo : outils de modification des points d'une s�quence : d�placement d'un point, ajout, suppression.

4 - Gestion de fichiers :
	* Sauvegarde en .mnc du projet, le .mnc a la structure d'un fichier .ini (voir code pour les sections et les cl�s).
	* Export en .dxf du projet (actuellement r�alis� par CNCTools.dll).
	* Sauvegarde des r�pertoires utilis�s dans "MiniCut2d Software.ini".
	* Boutons (et menus?) "Nouveau Projet", "Ouvrir", "Enregister", "Enregistrer sous...".

5 - Vectorisation :
	* Vectorisation d'images .jpg, .bmp ou .png int�gr�e au logiciel.
	* Recadrage.
	* Curseur de r�glage du contraste.
	* Curseur de r�glage du nettoyage des points.
	* Sauvegarde en .dxf dans la biblioth�que (mise � jour du TreeView).
	* Transfert dans la visionneuse.

6 - Repr�sentation de la d�coupe :
	* Repr�sentation graphique (simplifi�e) de la machine en fonction de ses caract�ristiques r�elles (voir code).
	* Repr�sentation du projet : bloc, s�quences assembl�es.
	* Choix du type d'entr�e, du type de sortie et repr�sentation graphique (attention, si elles sont de m�me type, d�caler la repr�sentation graphique pour qu'on voie les deux).
	* Zoom bloc (Todo : Zoom molette et d�placement de la vue).
	* Simulation du d�placement du fil.

7 - Choix du mode Normal - Expert

8 - Base de donn�es mati�res :
	* Caract�ristiques d'une mati�re : nom, chauffe, vitesse.
	* M�morisation dans "MiniCut2d Software.ini".
	* R�glage chauffe.
	* R�glage vitesse (mode Expert).
	* Affichage par liste d�roulante.
	* Cr�ation d'une nouvelle mati�re.
	* Suppression mati�re.
	* Mise � jour mati�re.

9 - Boutons de d�placements automatiques (les interrupteurs forment deux boucles, une boucle pour les origines et une boucle pour les fin de course) :
	* Retour � l'origine (voir code).
	* Retour en position de repos.

10 - Pilotage manuel du fil :
	* Marche - arr�t et r�glage dynamique de la chauffe (pr�voir tempo pour la mise en temp�rature du fil).
	* Information sur l'�tat en cours.
	* Mouvements suivant 8 directions + marche - arr�t.
	* Retour automatique suivant X, suivant Y.
	* ToDo : d�placement d'une valeur saisie au clavier.

11 - Lancement de la d�coupe :
	* Choix du d�calage (valeur fixe en mode Normal, plusieurs valeurs en mode Expert) : ext�rieur - sans - int�rieur.
	* Repr�sentation graphique du d�calage.
	* Information sur l'�tat en cours.
	* Lancement de la d�coupe.
	* Modification dynamique de la chauffe.
	* Arr�t de la d�coupe.
	
12 - Reprise de la d�coupe apr�s stop :
	* Choix entre annulation totale, retour origine, ou reprise de la d�coupe.
	* Modification possible de la chauffe avant reprise.
	* Choix du trajet de retour : horizontal, vertical, diagonal.
	
13 - Choix de la langue.

14 - A propos.


www.minicut2d.com - Renaud ILTIS
