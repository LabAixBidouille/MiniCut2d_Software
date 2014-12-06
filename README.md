###Objet :
**Le code de ce repository est celui du logiciel MiniCut2d Software, écrit en VB6 par Renaud ILTIS pour la machine de découpe par fil chaud "MiniCut2d".**

Il est publié avec les objectifs suivants :
- réécrire le logiciel pour qu'il puisse fonctionner sous Windows, Linux et Mac,
- conserver la simplicité et la stabilité du logiciel actuel,
- améliorer l'aspect visuel du logiciel (et peut-être aussi l'ergonomie).

Les sources sont fournies sous licence CECILL (compatible GNU GPL).

###Informations générales :
- Un fichier projet MiniCut2d (.mnc) contient les informations suivantes :
	* le nombre de séquences qui constituent le projet,
	* le nombre de point de chaque séquence,
	* les coordonnes X:Y de chaque plan dans le repère de la machine,
	* les dimensions du bloc de matière,
	* le type d'entrée: 0=par la gauche; 1=par le haut; 2=par la droite,
	* le type de sortie (idem).
	
- Penser tout de suite au multi-langues qui peut éventuellement être géré dans des fichiers texte séparés (actuellement les traductions sont contenues dans des modules compilés avec le .exe), mais prévoir toujours une valeur pas défaut dans le code en cas d'absence du fichier de langue.

- L'interface USB de la MiniCut2d est la version de base de l'interpolateur IPL5X disponible sur le site 5xproject.
	
###Description des différents modules du logiciel
1. Communication avec l'interface de type IPL5X reconnue comme périphérique HID (actuellement ce travail est réalisé par IPL5XCom.dll sur la base d'un tableau de 65 octets dont la première case est toujours vide, code en C fourni à la demande)
	* Identification du périphérique
	* Envoi d'octets
	* Réception d'octets
	
2. Visionneuse de fichiers :
	* Création d'une bibliothèque à l'installation du logiciel (=> ToDo : settings, choix de l'emplacement?).
	* TreeView de représentation permettant de parcourir la bibliothèque.
	* Fenêtre de visualisation du contenu du fichier.
	* Fichiers traités : DXF, DAT, PLT (traités actuellement par cnctools.dll dont le code en C peut fourni à la demande), TXT (format : cf. site MiniCut2d), MNC (projet MiniCut2d, voir code), ToDo : SVG.
	* Les profils fermés sont réorientés/renumérotés pour mettre le premier point vers l'origine suivant X (voir code) et la rotation en sens horaire.
	* Les profils issus de DXF sont nettoyés (suppression des points alignés ou trop proches), voir code et fonction Nettoyage.
	* Opérations élémentairs sur le contenu affiché : affichage du premier point, miroir, changement de sens.

3. Création de projet :
	* Fenêtre réprésentant la zone utile de la machine et le bloc.
	* Saisie des dimensions du bloc.
	* Intégration des séquences par drag and drop ou double-clic depuis la visionneuse.
	* Affichage des séquences (en noir) et des jonctions (en vert) entre les séquences.
	* Affichage dynamique de la jonction en fonction de l'emplacement de la nouvelle séquence importée par drag and drop (en rose tant que la séquence n'est pas posée).
	* Zoom molette, déplacement de la vue, zoom bloc.
	* Affichage des points.
	* Alternance des couleurs des séquences pour les distinguer les unes des autres.
	* Outils de sélection/déplacement/rotation/étirement (à la main ou par saisie de valeur), affichage dynamique de la dimension de la sélection.
	* Outils de modification du premier point, de mesure, de coupe des séquences, de suppression de la sélection.
	* Outils d'inversion de sens d'une ou plusieurs séquences.
	* Outil d'insertion de point.
	* Outil miroir.
	* Outil dupliquer (ToDo : remplacer l'outil duliquer par Copier+Coller).
	* Ajout de point par Drag and Drop.
	* Outils de mise à l'échelle du bloc et de centrage horizontal et vertical (ToDo : case à cocher pour verrouiller ou non ces boutons).
	* Outils d'alignement (ToDo : réduire la place prise par ces outils : boutons déroulants?).
	* ToDo : outils de modification des points d'une séquence : déplacement d'un point, ajout, suppression.

4. Gestion de fichiers :
	* Sauvegarde en .mnc du projet, le .mnc a la structure d'un fichier .ini (voir code pour les sections et les clés).
	* Export en .dxf du projet (actuellement réalisé par CNCTools.dll).
	* Sauvegarde des répertoires utilisés dans "MiniCut2d Software.ini".
	* Boutons (et menus?) "Nouveau Projet", "Ouvrir", "Enregister", "Enregistrer sous...".

5. Vectorisation :
	* Vectorisation d'images .jpg, .bmp ou .png intégrée au logiciel.
	* Recadrage.
	* Curseur de réglage du contraste.
	* Curseur de réglage du nettoyage des points.
	* Sauvegarde en .dxf dans la bibliothèque (mise à jour du TreeView).
	* Transfert dans la visionneuse.

6. Représentation de la découpe :
	* Représentation graphique (simplifiée) de la machine en fonction de ses caractéristiques réelles (voir code).
	* Représentation du projet : bloc, séquences assemblées.
	* Choix du type d'entrée, du type de sortie et représentation graphique (attention, si elles sont de même type, décaler la représentation graphique pour qu'on voie les deux).
	* Zoom bloc (Todo : Zoom molette et déplacement de la vue).
	* Simulation du déplacement du fil.

7. Choix du mode Normal - Expert

8. Base de données matières :
	* Caractéristiques d'une matière : nom, chauffe, vitesse.
	* Mémorisation dans "MiniCut2d Software.ini".
	* Réglage chauffe.
	* Réglage vitesse (mode Expert).
	* Affichage par liste déroulante.
	* Création d'une nouvelle matière.
	* Suppression matière.
	* Mise à jour matière.

9. Boutons de déplacements automatiques (les interrupteurs forment deux boucles, une boucle pour les origines et une boucle pour les fin de course) :
	* Retour à l'origine (voir code).
	* Retour en position de repos.

10. Pilotage manuel du fil :
	* Marche - arrêt et réglage dynamique de la chauffe (prévoir tempo pour la mise en température du fil).
	* Information sur l'état en cours.
	* Mouvements suivant 8 directions + marche - arrêt.
	* Retour automatique suivant X, suivant Y.
	* ToDo : déplacement d'une valeur saisie au clavier.

11. Lancement de la découpe :
	* Choix du décalage (valeur fixe en mode Normal, plusieurs valeurs en mode Expert) : extérieur - sans - intérieur.
	* Représentation graphique du décalage.
	* Information sur l'état en cours.
	* Lancement de la découpe.
	* Modification dynamique de la chauffe.
	* Arrêt de la découpe.
	
12. Reprise de la découpe après stop :
	* Choix entre annulation totale, retour origine, ou reprise de la découpe.
	* Modification possible de la chauffe avant reprise.
	* Choix du trajet de retour : horizontal, vertical, diagonal.
	
13. Choix de la langue.

14. A propos.

###Demande d'informations complémentaires :

Renaud ILTIS - contact@minicut2d.com
