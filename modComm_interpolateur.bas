Attribute VB_Name = "modComm_interpolateur"
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

'********** fonctions de communication avec IPL5X (IPL5XCom.dll)*************
Public Declare Function IPL5X_Dll_Version Lib "IPL5XCom.dll" () As Long
'fonction de lecture de la version de la dll de communication
'la fonction renvoie le num�ro de version
Public Declare Function IPL5X_Open Lib "IPL5XCom.dll" () As Long
'fonction d'ouverture de la communication avec IPL5X
'renvoie -201 si pas de comm possible , 1 si IPL5X trouv� et comm possible.
Public Declare Sub IPL5X_Close Lib "IPL5XCom.dll" ()
'proc�dure de fermeture de la communication avec IPL5X
'il faut penser � fermer la communication avec IPL5X avant la fermeture du soft
Public Declare Function IPL5X_IsConnected Lib "IPL5XCom.dll" () As Long
'la fonction renvoie -200 si IPL5X n'est pas trouv�e et 1 si IPL5X est trouv�e et la communication ouverte
Public Declare Function IPL5X_Send Lib "IPL5XCom.dll" (Data() As Byte, Fast As Long) As Long
'fonction d'envoi des ByteIPL � IPL5X
'les ByteIPL doivent �tre envoy�es dans un tableau de 65 cases num�rot�es de 0 � 64
' dont les donn�es doivent �tre de type Byte en Visual Basic 6
'la premi�re case doit �tre un 0, le premier octet doit �tre dans la deuxi�me case
'la fonction renvoie :
'1 Send/receive OK, data stream is valid
'-200      IPL5X not found
'-201      Can�t open communication channel
'-202      An error occurred while sending data to IPL5X
'-203      An error occurred while receiving data to IPL5X
'-204      Send/receive data mismatch
'-101      Can�t lock the table *
'-102      Table is too small *

'******** VIDAGE DU TABLEAU DE COMMUNICATION ********
Public Sub RAZByteComm()
   Dim i As Integer
   'On met 0 dans toutes les cases du tableau d'ByteIPL puis on rempli les cases qui nous int�ressent
   For i = 0 To 64
      ByteIPL(i) = &H0
   Next i
End Sub

'******** ENVOI D'INSTRUCTIONS COURTES ********
Public Sub EnvoiBytes(ByVal ByteNum0 As Byte, Optional ByVal ByteNum1 As Byte = &H0, Optional ByVal ByteNum2 As Byte = &H0)
         Call RAZByteComm
         ByteIPL(1) = ByteNum0
         ByteIPL(2) = ByteNum1
         ByteIPL(3) = ByteNum2
         ErrIPL = IPL5X_Send(ByteIPL(), 0)
         'ErrIPL est une variable globale qui sera d�cod�e par la proc�dure appelante
End Sub

'********* TEST DU RETOUR DE LA FONCTION IPL5X_Send : Gestion de l'erreur *******************
Public Function DecodeErrIPL(ByVal Erreur As Long) As String
   Select Case Erreur
   Case -200
      DecodeErrIPL = "Erreur -200 : l'interpolateur n'a pas �t� trouv�."
   Case -201
      DecodeErrIPL = "Erreur -201 : impossible d'ouvrir le canal."
   Case -202
      DecodeErrIPL = "Erreur -202 : une erreur s'est produite durant l'envoi des donn�es � l'interpolateur."
   Case -203
      DecodeErrIPL = "Erreur -203 : une erreur s'est produite lors de la r�ception de la r�ponse de l'interpolateur."
   Case -204
      DecodeErrIPL = "Erreur -204 : probl�me de coh�rence entre les informations envoy�es � l'interpolateur et la r�ponse re�ue."
   Case -500  'code perso pour pb de DLL
      DecodeErrIPL = "Le fichier IPL5XCom.dll n'a pas �t� trouv�," & vbCrLf & _
                     "veuillez l'installer sur votre ordinateur, � la bonne place."
   Case -101
      DecodeErrIPL = "Erreur -101 : can't lock the table."
   Case -102
      DecodeErrIPL = "Erreur -102 : table is too small."
   Case Else
      DecodeErrIPL = "Erreur non r�pertori�e."
   End Select
End Function

'********* Ecriture d'une table dans IPL5X ***********
Public Sub EcrireTable()
   Dim i As Integer

   Dim CMD As Byte 'commande
   Dim NBR As Byte 'num�ro de la table (0 dans notre cas)
   'low part
   Dim NomTableTemp As String
   Dim N(1 To 8) As Byte, FREQ As Byte, FLAGS As Byte, IO1 As Byte, LANG As Byte
   'high part
   Dim NbrPas1MmXG As Long, NbrPas1MmYG As Long, NbrPas1MmXD As Long, NbrPas1MmYD As Long
   Dim NbrPas1sSsAccXG As Long, NbrPas1sSsAccYG As Long, NbrPas1sSsAccXD As Long, NbrPas1sSsAccYD As Long
   Dim NbrPas1sAvAccXG As Long, NbrPas1sAvAccYG As Long, NbrPas1sAvAccXD As Long, NbrPas1sAvAccYD As Long
   Dim SL(1 To 5) As Byte, SM(1 To 5) As Byte, VML(1 To 5) As Byte, VMM(1 To 5), VMAL(1 To 5) As Byte, VMAM(1 To 5) As Byte
   'upper part
   Dim ORIL(1 To 5) As Byte, ORIM(1 To 5) As Byte
   
   'D�finition des caract�ristiques des tables � partir des variables d�finies dans la proc�dure appelante
   NomTableTemp = NomTable
    If Len(NomTableTemp) < 8 Then  'on ajoute des espaces pour compl�ter � 8 caract�res
      For i = 1 To (8 - Len(NomTableTemp))
         NomTableTemp = NomTableTemp & " "
      Next i
   End If
   For i = 1 To 8
      N(i) = AscB(Mid(NomTableTemp, i, 1))
   Next i
   FREQ = Frequence / 10000 ' 1=10kHz, 2=20kHz, 3=30kHz, 4=40kHz, 5=50kHz
   NbrPas1MmXG = PasParTourXG / MmParTourXG
   NbrPas1MmYG = PasParTourYG / MmParTourYG
   NbrPas1MmXD = PasParTourXD / MmParTourXD
   NbrPas1MmYD = PasParTourYD / MmParTourYD
   'sans acc�l�ration en 1s
   NbrPas1sSsAccXG = NbrPas1MmXG * VMaxSansAcc
   NbrPas1sSsAccYG = NbrPas1MmYG * VMaxSansAcc
   NbrPas1sSsAccXD = NbrPas1MmXD * VMaxSansAcc
   NbrPas1sSsAccYD = NbrPas1MmYD * VMaxSansAcc
   'avec acc�l�ration en 1s
   NbrPas1sAvAccXG = NbrPas1MmXG * VMaxAvecAcc
   NbrPas1sAvAccYG = NbrPas1MmYG * VMaxAvecAcc
   NbrPas1sAvAccXD = NbrPas1MmXD * VMaxAvecAcc
   NbrPas1sAvAccYD = NbrPas1MmYD * VMaxAvecAcc
   
   'Donn�es low part :
   FLAGS = BaseToBase("00000101", 2, 10) '0(na) 0(na) 0(na) 0(na) 0(na) 1(signal motor on invers�) 0(orientation table normale) 1(FDC activ�s)
   IO1 = BaseToBase("00000001", 2, 10) 'input
   LANG = 0 'french
   SL(1) = NbrPas1MmXG And &HFF&
   SM(1) = (NbrPas1MmXG And &HFF00&) / &H100&
   SL(2) = NbrPas1MmYG And &HFF&
   SM(2) = (NbrPas1MmYG And &HFF00&) / &H100&
   SL(3) = NbrPas1MmXD And &HFF&
   SM(3) = (NbrPas1MmXD And &HFF00&) / &H100&
   SL(4) = NbrPas1MmYD And &HFF&
   SM(4) = (NbrPas1MmYD And &HFF00&) / &H100&
   SL(5) = 0  'N.A.
   SM(5) = 0  'N.A.
   VML(1) = NbrPas1sSsAccXG And &HFF&
   VMM(1) = (NbrPas1sSsAccXG And &HFF00&) / &H100&
   VML(2) = NbrPas1sSsAccYG And &HFF&
   VMM(2) = (NbrPas1sSsAccYG And &HFF00&) / &H100&
   VML(3) = NbrPas1sSsAccXD And &HFF&
   VMM(3) = (NbrPas1sSsAccXD And &HFF00&) / &H100&
   VML(4) = NbrPas1sSsAccYD And &HFF&
   VMM(4) = (NbrPas1sSsAccYD And &HFF00&) / &H100&
   VML(5) = 0 'N.A.
   VMM(5) = 0 'N.A.
   VMAL(1) = NbrPas1sAvAccXG And &HFF&
   VMAM(1) = (NbrPas1sAvAccXG And &HFF00&) / &H100&
   VMAL(2) = NbrPas1sAvAccYG And &HFF&
   VMAM(2) = (NbrPas1sAvAccYG And &HFF00&) / &H100&
   VMAL(3) = NbrPas1sAvAccXD And &HFF&
   VMAM(3) = (NbrPas1sAvAccXD And &HFF00&) / &H100&
   VMAL(4) = NbrPas1sAvAccYD And &HFF&
   VMAM(4) = (NbrPas1sAvAccYD And &HFF00&) / &H100&
   VMAL(5) = 0 'N.A.
   VMAM(5) = 0 'N.A.
   ORIL(1) = NbrPasToOriXG And &HFF&
   ORIM(1) = (NbrPasToOriXG And &HFF00&) / &H100&
   ORIL(2) = NbrPasToOriYG And &HFF&
   ORIM(2) = (NbrPasToOriYG And &HFF00&) / &H100&
   ORIL(3) = NbrPasToOriXD And &HFF&
   ORIM(3) = (NbrPasToOriXD And &HFF00&) / &H100&
   ORIL(4) = NbrPasToOriYD And &HFF&
   ORIM(4) = (NbrPasToOriYD And &HFF00&) / &H100&
   ORIL(5) = 0  'N.A.
   ORIM(5) = 0  'N.A.
'Ecriture :
   'low part:
   Call RAZByteComm  'on met z�ro partout
   ByteIPL(1) = &H54 'commande d'acc�s aux tables
   ByteIPL(2) = &H4 'CMD commande demand�e
   ByteIPL(3) = &H0   'NBR num�ro de la table
   For i = 1 To 8
      ByteIPL(3 + i) = N(i)
   Next i
   For i = 1 To 5
      ByteIPL(11 + i) = TypeAxeInterpolateur(i)
   Next i
   For i = 1 To 10
      ByteIPL(16 + i) = DefinitionSortie(i)
   Next i
   ByteIPL(27) = FREQ
   ByteIPL(28) = PenteAcceleration 'de 0=pente douce � 15=pente raide
   ByteIPL(29) = PWMmaxi ' de 0 � 255 (convertir en %)
   ByteIPL(30) = DelaiMoteursOff ' de 0 � 127 secondes
   ByteIPL(31) = FLAGS
   ByteIPL(32) = IO1
   ByteIPL(33) = LANG
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL <> 1 Then  'erreur
      Exit Sub
   End If
   
   'Ecriture high part :
   Call RAZByteComm  'on met z�ro partout
   ByteIPL(1) = &H54 'commande d'acc�s aux tables
   ByteIPL(2) = &H6 'CMD commande demand�e
   ByteIPL(3) = 0   'NBR num�ro de la table
   ByteIPL(4) = SL(1)
   ByteIPL(5) = SM(1)
   ByteIPL(6) = SL(2)
   ByteIPL(7) = SM(2)
   ByteIPL(8) = SL(3)
   ByteIPL(9) = SM(3)
   ByteIPL(10) = SL(4)
   ByteIPL(11) = SM(4)
   ByteIPL(12) = SL(5)
   ByteIPL(13) = SM(5)
   ByteIPL(14) = VML(1)
   ByteIPL(15) = VMM(1)
   ByteIPL(16) = VML(2)
   ByteIPL(17) = VMM(2)
   ByteIPL(18) = VML(3)
   ByteIPL(19) = VMM(3)
   ByteIPL(20) = VML(4)
   ByteIPL(21) = VMM(4)
   ByteIPL(22) = VML(5)
   ByteIPL(23) = VMM(5)
   ByteIPL(24) = VMAL(1)
   ByteIPL(25) = VMAM(1)
   ByteIPL(26) = VMAL(2)
   ByteIPL(27) = VMAM(2)
   ByteIPL(28) = VMAL(3)
   ByteIPL(29) = VMAM(3)
   ByteIPL(30) = VMAL(4)
   ByteIPL(31) = VMAM(4)
   ByteIPL(32) = VMAL(5)
   ByteIPL(33) = VMAM(5)
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL <> 1 Then  'erreur
      Exit Sub
   End If
   
   'Ecriture upper part :
   Call RAZByteComm  'on met z�ro partout
   ByteIPL(1) = &H54 'commande d'acc�s aux tables
   ByteIPL(2) = &H8 'CMD commande demand�e
   ByteIPL(3) = 0   'NBR num�ro de la table
   ByteIPL(4) = ORIL(1)
   ByteIPL(5) = ORIM(1)
   ByteIPL(6) = ORIL(2)
   ByteIPL(7) = ORIM(2)
   ByteIPL(8) = ORIL(3)
   ByteIPL(9) = ORIM(3)
   ByteIPL(10) = ORIL(4)
   ByteIPL(11) = ORIM(4)
   ByteIPL(12) = ORIL(5)
   ByteIPL(13) = ORIM(5)
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL <> 1 Then  'erreur
      Exit Sub
   End If
   
   ' La table a �t� �crite dans l'EEPROM, il faut la recharger en RAM
   Call RAZByteComm  'on met z�ro partout
   ByteIPL(1) = &H54 'commande d'acc�s aux tables
   ByteIPL(2) = &H1 'sp�cifier le num�ro
   ByteIPL(3) = &H0 'table 0
   ErrIPL = IPL5X_Send(ByteIPL(), 0) 'envoi du tableau des ByteIPL
   If ErrIPL <> 1 Then  'erreur
      Exit Sub
   End If

End Sub

'******* Transfert des donn�es de la salve Data dans le tableau de communication avec l'interpolateur *****
'**** Rappel : la salve data est la salve de d�placements ******
Public Sub TransfertData2Bytes(SalveATransferer As SalveData)
   Call RAZByteComm  'on met z�ro partout
   With SalveATransferer
      ByteIPL(1) = &H44
      ByteIPL(2) = .CMD
      ByteIPL(3) = .NBRL
      ByteIPL(4) = .NBRM
      ByteIPL(5) = .NBRH
      ByteIPL(6) = .NBRU
      ByteIPL(7) = .S1L
      ByteIPL(8) = .S1M
      ByteIPL(9) = .S1H
      ByteIPL(10) = .S1U
      ByteIPL(11) = .S2L
      ByteIPL(12) = .S2M
      ByteIPL(13) = .S2H
      ByteIPL(14) = .S2U
      ByteIPL(15) = .S3L
      ByteIPL(16) = .S3M
      ByteIPL(17) = .S3H
      ByteIPL(18) = .S3U
      ByteIPL(19) = .S4L
      ByteIPL(20) = .S4M
      ByteIPL(21) = .S4H
      ByteIPL(22) = .S4U
      ByteIPL(23) = .S5L
      ByteIPL(24) = .S5M
      ByteIPL(25) = .S5H
      ByteIPL(26) = .S5U
      ByteIPL(27) = .F_ACC
      ByteIPL(28) = .F_DEC
      ByteIPL(29) = .DECL
      ByteIPL(30) = .DECM
      ByteIPL(31) = .DECH
      ByteIPL(32) = .DECU
   End With
End Sub

