Attribute VB_Name = "modGestionINI"
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

'Lecture des cl�s d'un fichier de type .ini
Public Function LitFichierTypeIni(ByVal Section As String, ByVal Cle As String, fichier As String, _
      Optional ByVal ValeurParDefaut As String = "") As String
   Dim strReturn As String
   strReturn = String(255, 0)
   GetPrivateProfileString Section, Cle, ValeurParDefaut, strReturn, Len(strReturn), fichier
   LitFichierTypeIni = Left(strReturn, InStr(strReturn, Chr(0)) - 1)
End Function

'Ecriture des sections et cl�s du fichier MiniCut2d_Software.ini
Public Function EcritFichierIni(ByVal Section As String, ByVal Cle As String, _
                                     ByVal Valeur As String) As Long
   EcritFichierIni = WritePrivateProfileString(Section, Cle, Valeur, App.Path & "\MiniCut2d_Software.ini")
End Function

'Lecture des cl�s du fichier MiniCut2d_Software.ini
Public Function LitFichierIni(ByVal Section As String, ByVal Cle As String, _
      Optional ByVal ValeurParDefaut As String = "") As String
   Dim strReturn As String
   strReturn = String(255, 0)
   GetPrivateProfileString Section, Cle, ValeurParDefaut, strReturn, Len(strReturn), App.Path & "\MiniCut2d_Software.ini"
   LitFichierIni = Left(strReturn, InStr(strReturn, Chr(0)) - 1)
End Function

'Effacement de toutes les cl�s d'une section du fichier MiniCut2d_Software.ini
Public Function EffacerSectionIni(ByVal Section As String) As Boolean
   EffacerSectionIni = WritePrivateProfileSection(Section, "", App.Path & "\MiniCut2d_Software.ini")
End Function

'Etablir la liste de toutes les sections du MiniCut2d_Software.ini
Public Function ListeSectionIni(Section() As String)
    Dim strReturn As String
    strReturn = String(8192, 0)
    GetPrivateProfileSectionNames strReturn, Len(strReturn), App.Path & "\MiniCut2d_Software.ini"
    Section = Split(Left(strReturn, InStr(1, strReturn, vbNullChar & vbNullChar) - 1), vbNullChar)
End Function

'Etablir la liste de toutes les cl�s d'une section du .ini
Public Function ListeSectionKey(ByVal Path As String, ByVal Section As String, Key() As String)
    Dim strReturn As String
    strReturn = String(8192, 0)
    GetPrivateProfileSection Section, strReturn, 8192, Path
    Key = Split(Left(strReturn, InStr(1, strReturn, vbNullChar & vbNullChar) - 1), vbNullChar)
'Utilisation :
'Private Sub Command1_Click()
'    Dim Key() As String
'    ListeSectionKey "C:\test.ini", "SectionName1", Key '-- le param�tre Key est pass� byRef
'    For Index = LBound(Key) To UBound(Key)
'        Debug.Print Key(Index)
'    Next
'End Sub
End Function
 
'**** Effacement d'une ligne du fichier .ini, en deux fonctions ****
' Cette fonction lit le contenu du fichier szFileName et retourne
' ce contenu. En cas d'erreur, elle retourne une cha�ne vide et
' renseigne le code d'erreur et la description de l'erreur
'
Public Function ReadFileToBuffer(ByVal szFileName As String, _
                                ByRef errCode As Integer, _
                                ByRef errString As String) As String
    Dim f As Integer
    Dim buffer As String

    ' trappe les erreurs
    On Error GoTo ReadFileToBuffer_ERR

    ' Ouverture du fichier en 'Binary'
    f = FreeFile
    Open szFileName For Binary As #f
        ' pr�allocation d'un buffer � la taille du fichier
        buffer = Space$(LOF(f))
        ' lecture compl�te du fichier
        Get #f, , buffer
    Close #f
    ReadFileToBuffer = buffer
ReadFileToBuffer_END:
    Exit Function
    
ReadFileToBuffer_ERR:
    ' Gestion d'erreur
    ReadFileToBuffer = ""
    errCode = Err.Number
    errString = Err.Description
    Resume ReadFileToBuffer_END
End Function
' La fonction suivante supprime la ligne comprennant une cha�ne de caract�res donn�e (Patten)
' Si mode=0, supprime seulement la premi�re occurrence de Pattern
' Si mode=1, supprime toutes les occurrences de pattern
Public Sub EffacerLigneDansIni(ByVal Pattern As String)
    Dim f As Integer, errCode As Integer, errString As String
    Dim buffer As String
    Dim t() As String
    Dim i As Long
    Dim nbOcc As Long
    Dim mode As Integer
    mode = 1  'j'ai choisi de ne pas utiliser ce param�tre
    buffer = ReadFileToBuffer(App.Path & "\MiniCut2d_Software.ini", errCode, errString)
    t() = Split(buffer, vbCrLf)
    f = FreeFile
    Open App.Path & "\MiniCut2d_Software.ini" For Output As #f
        For i = 0 To UBound(t()) - 1
            If InStr(t(i), Pattern) > 0 Then
                ' si on trouve le pattern
                nbOcc = nbOcc + 1
                If (nbOcc = 1) Or (mode = 1) Then
                    ' on n'�crit pas la ligne
                Else
                    Print #f, t(i)
                End If
            Else
                Print #f, t(i)
            End If
        Next i
    Close #f
End Sub
' Exemple d'appel
'Private Sub Command3_Click()
    ' Suppression de la premi�re ligne contenant "Bernard"
 '   Call RemoveLineFromFileByPattern("c:\publicdata\test.txt", "Bernard", 0)
    ' Suppression de toutes les lignes contenant "Henri"
  '  Call RemoveLineFromFileByPattern("c:\publicdata\test.txt", "Henri", 1)
'End Sub

'Sauvegarde d'une mati�re dans le fichier .ini
Public Sub SauverMatiere(ByVal Nom As String, ByVal Chauffe As Single, ByVal Vitesse As Single)
   EcritFichierIni "Matiere_" & Nom, "ChauffeDecoupe", Chauffe
   EcritFichierIni "Matiere_" & Nom, "VitesseDecoupe", V2P(Format(Vitesse, "0.0"))  'V2P pour toujours utiliser le point comme s�parateur
End Sub

Public Sub ListerMatieresDuIni()
'Chercher toutes les tables dans le fichier MiniCut2d_Software.ini
   Dim buffer As String
   Dim errCode As Integer, errString As String

   Dim t() As String
   Dim i As Integer, j As Integer
   Dim NomTemp As String
      
   buffer = ReadFileToBuffer(App.Path & "\MiniCut2d_Software.ini", errCode, errString) 'on r�cup�re le texte du .ini
   t() = Split(buffer, vbCrLf)   'on le coupe en lignes qu'on met dans un tableau
   ReDim MatieresDeLaBase(0 To 0) 'l'indice 0 est r�serv� aux r�glages par d�faut
   j = 1
   For i = 0 To UBound(t()) - 1  'on parcoure le tableau du .ini et on recopie le nom des tables
      If Left$(t(i), 9) = "[Matiere_" Then
         NomTemp = Mid$(t(i), 10, Len(t(i)) - 10)
         If NomTemp = "Reglage par defaut" Then
            MatieresDeLaBase(0).Nom = NomTemp
            MatieresDeLaBase(0).Chauffe = LitFichierIni((Mid$(t(i), 2, Len(t(i)) - 2)), "ChauffeDecoupe")
            'ajout de la vitesse pour le mode Expert, il faut rajouter la vitesse dans les anciens .ini
            MatieresDeLaBase(0).Vitesse = V2P(LitFichierIni("Matiere_" & MatieresDeLaBase(0).Nom, "VitesseDecoupe"))  'on change les virgules en point
            If MatieresDeLaBase(0).Vitesse = "" Then 'c'est qu'on est sur un ancien .ini
               MatieresDeLaBase(0).Vitesse = "4.0" 'c'est une variable string
               EcritFichierIni "Matiere_" & MatieresDeLaBase(0).Nom, "VitesseDecoupe", MatieresDeLaBase(0).Vitesse
            End If
         Else
            ReDim Preserve MatieresDeLaBase(0 To j)
            MatieresDeLaBase(j).Nom = NomTemp
            MatieresDeLaBase(j).Chauffe = LitFichierIni((Mid$(t(i), 2, Len(t(i)) - 2)), "ChauffeDecoupe")
            'ajout de la vitesse pour le mode Expert, il faut rajouter la vitesse dans les anciens .ini
            MatieresDeLaBase(j).Vitesse = V2P(LitFichierIni("Matiere_" & MatieresDeLaBase(j).Nom, "VitesseDecoupe"))  'on change les virgules en point
            If MatieresDeLaBase(j).Vitesse = "" Then 'c'est qu'on est sur un ancien .ini
               MatieresDeLaBase(j).Vitesse = "4.0" 'c'est une variable string
               EcritFichierIni "Matiere_" & MatieresDeLaBase(j).Nom, "VitesseDecoupe", MatieresDeLaBase(j).Vitesse
            End If
            j = j + 1
         End If
      End If
   Next i
   j = 0
   ReDim MatieresAffichees(0)
   For i = 0 To UBound(MatieresDeLaBase)
      If ModeSoft = "Expert" Or (ModeSoft = "Normal" And MatieresDeLaBase(i).Vitesse = "4.0") Then
         ReDim Preserve MatieresAffichees(0 To j)
         MatieresAffichees(j) = MatieresDeLaBase(i)
         j = j + 1
      End If
   Next i
   'Remplissage du ComboBox des mati�res :
   'Chercher toutes les mati�res dans le fichier .ini et les mettre dans MatieresDeLaBase()
   ' et celles � afficher dans le combo dans MatieresAffichees()
   'Remplir le combobox avec toutes les mati�res du .ini
   'Comme la mati�re par d�faut a �t� cr��e en premier dans le .ini, elle sera toujours en haut
   '(sauf modification de l'ordre dans le .ini)

   'affichage dans le combobox
   With frmMiniCut2d
      .comboMatieres.Clear  'effacement
      For i = 0 To UBound(MatieresAffichees)
         If ModeSoft = "Expert" Then 'en mode Expert, on affiche la vitesse
            .comboMatieres.AddItem MatieresAffichees(i).Nom & " - " & MatieresAffichees(i).Chauffe & "%" _
                                                          & " - " & MatieresAffichees(i).Vitesse & "mm/s"
         Else
            .comboMatieres.AddItem MatieresAffichees(i).Nom & " - " & MatieresAffichees(i).Chauffe & "%"
         End If
      Next i
      For i = 0 To UBound(MatieresAffichees)
         If MatieresAffichees(i).Nom = MatiereUtilisee.Nom Then
            .comboMatieres.Text = .comboMatieres.List(i)
            Exit For
         End If
      Next i
      If .comboMatieres.Text = "" Then 'on �tait sur une valeur Expert et on est pass� en Normal
         .comboMatieres.Text = .comboMatieres.List(0)  'va d�clencher la mise � jour des valeurs et du .ini
      End If
   End With
End Sub

'********************************************************************************************
'******* Fonctions pour sauvegarde du projet en fichier *.mnc en fichier de type .ini *******
'********************************************************************************************

'Ecriture des sections et cl�s du fichier de d�coupe
Public Function EcritDecoupeSousFormeIni(ByVal Section As String, ByVal Cle As String, _
                                     ByVal Valeur As String, ByVal fichier As String) As Long
   EcritDecoupeSousFormeIni = WritePrivateProfileString(Section, Cle, Valeur, fichier)
End Function
'utilisation : EcritDansFichierIni "MaSection", "MaSousSection", MaValeur, CheminFichierIni

'Lecture des cl�s du fichier de d�coupe
Public Function LitDecoupeSousFormeIni(ByVal Section As String, ByVal Cle As String, ByVal fichier As String, _
      Optional ByVal ValeurParDefaut As String = "") As String
   Dim strReturn As String
   strReturn = String(255, 0)
   GetPrivateProfileString Section, Cle, ValeurParDefaut, strReturn, Len(strReturn), fichier
   LitDecoupeSousFormeIni = Left(strReturn, InStr(strReturn, Chr(0)) - 1)
End Function
'utilisation : Dim CheminFichierIni As String
               'CheminFichierIni = App.Path & "\MonFichier.ini"
               'TxtDateCreation.Text = LitDecoupeSousFormeIni("APROPOS", "DateCreation", CheminFichierIni)
               'LblAuteur1.Caption = LitDecoupeSousFormeIni("APROPOS", "Auteur1", CheminFichierIni)

'Effacement de toutes les cl�s d'une section du fichier de d�coupe
Public Function EffacerSectionDecoupeSousFormeIni(ByVal Section As String, ByVal fichier As String) As Boolean
   EffacerSectionDecoupeSousFormeIni = WritePrivateProfileSection(Section, "", fichier)
End Function

