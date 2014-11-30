Attribute VB_Name = "modRepertoiresSpeciaux"
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

'Récupération des chemin des dossiers spéciaux : bureau, Mes Documents, etc.

' Enumération des dossiers spéciaux
Public Enum SpecialFoldersConstants
    CSIDL_ADMINTOOLS = &H30
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A
    CSIDL_BITBUCKET = &HA
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_FAVORITES = &H6
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_MASK = &HFF00
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_MYPICTURES = &H27
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PERSONAL = &H5
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAMS = &H2
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25
    CSIDL_SYSTEMX86 = &H29
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24
End Enum

Public Function GetSpecialFolder(SpecialFolder As SpecialFoldersConstants) As String

    ' Les variables
    Dim RC As Long
    Dim IDL As ITEMIDLIST
    Dim spath As String

    ' Récupère le dossier spécial
    RC = SHGetSpecialFolderLocation(100, SpecialFolder, IDL)
    If RC = 0 Then
        ' Crée un tampon
        spath = Space$(MAX_PATH)
        ' Récupère le path à partir de l'IDList
        SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal spath
        ' Supprime les chr$(0) inutiles
        spath = Left$(spath, InStr(spath, Chr$(0)) - 1)
        If Right$(spath, 1) <> "\" Then spath = spath & "\"
    Else
        spath = ""
    End If
    GetSpecialFolder = spath

End Function



