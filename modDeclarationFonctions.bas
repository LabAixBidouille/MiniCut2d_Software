Attribute VB_Name = "modDeclarationFonctions"
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

'Fonctions pour conversion d'image
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function GetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Function SetBitmapBits Lib "gdi32" ( _
    ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
 


'****************************************************************************
'D�claration utilisation CNCTools.dll (fonction de lecture de fichier profil)
'les param�tres sont pass�s byref et modifi�s par la dll
'****************************************************************************

Public Declare Function LireFichier Lib "CNCTools.dll" (tableau() As ProfilDATDXF, NomFic As String, _
                                                Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single, _
                                                Comment As String) As Long

Public Declare Function EcrireFichier Lib "CNCTools.dll" (tableau() As ProfilDATDXF, NomFic As String, Comment As String) As Long

Public Declare Function EcrireFichier3D Lib "CNCTools.dll" (Emplanture() As ProfilDATDXF, Saumon() As ProfilDATDXF, _
                                                      NomFic As String, Ecartement As Single) As Long
                                                      
Public Declare Function LissageSequ Lib "CNCTools.dll" (tableau() As ProfilDATDXF) As Long
                                                      
Public Declare Function Version Lib "CNCTools.dll" () As Long
'fonction de lecture de la version de la dll
'la fonction renvoie le num�ro de version
                                                                                                          
'Copie de pixels pour cr�ation de curseur personnalis�
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'Dessin de polygones pour le dessin de la table
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
                                    
'Copie d'une image en enlevant une couleur
Public Declare Function TransfparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransfparent As Long) As Boolean
                                    
'Tempo pour acquisition
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Pour tester l'�tat d'une touche
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_TAB = &H9 'pour la touche tab
Public Const VK_DOWN = &H28 'pour la touche curseur bas
Public Const VK_RIGHT = &H27 'pour la touche curseur droite
Public Const VK_UP = &H26 'pour la touche curseur haut
Public Const VK_LEFT = &H25 'pour la touche curseur gauche

'Pour les dimensions de la fen�tre
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Pour tester la perte de focus
Public Declare Function GetActiveWindow Lib "user32" () As Long

'Lecture/�criture des fichiers .ini

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" _
     (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
     
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" _
(ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Les fonctions pour le browser de r�pertoires lors de l'importation de fichier
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias _
            "SHGetPathFromIDListA" (ByVal pidl As Long, _
            ByVal pszPath As String) As Long                                     '*
            
Public Declare Function SHBrowseForFolder Lib "Shell32.dll" Alias _
            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
            As Long                                                              '*
            
Public Const BIF_RETURNONLYFSDIRS = &H1

' R�cup�ration des dossiers sp�ciaux (voir module CheminRepertoires
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

'API de copie des fichiers
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" _
(ByVal lpExistingFileName As String, ByVal lpNewFileName As String, _
ByVal bFailIfExists As Long) As Long
'bFailIfExists doit etre � false lors de l'appel pour permettre 'l'overwriting

'Pour g�rer l'�conomiseur d'�cran
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
               (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Long _
               , ByVal fuWinIni As Long) As Long
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPI_GETSCREENSAVEACTIVE As Long = &H10
Public Const SPI_GETSCREENSAVERRUNNING As Long = &H72
Public Const SPI_GETSCREENSAVETIMEOUT = 14

'activation de l'�conomiseur au niveau de la base de registre :
'SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 1, 0, 0
'd�sactivation de l'�conomiseur au niveau de la base de registre :
'SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0, 0, 0
'pour savoir si le screensaver est en cours d'ex�cution :
'dim bRunning as boolean
'SystemParametersInfo SPI_GETSCREENSAVERRUNNING, 0, bRunning, False
' et on teste bRunning
'pour savoir si le screensaver est activ� au niveau de la base de registre
'Dim bActive As Boolean
'SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, bActive, False
'    If bActive Then
'        Me.Caption = "Screen saver is active"
'    Else
'        Me.Caption = "Screen saver is not active"
'    End If

'Pour d�finir la position du pointeur lors du zoom
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 

