Attribute VB_Name = "modLangueEspanol"
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

Public Sub LangueEspanol(LangueAUtiliser As String)
   '**** traduction du SplashScreen
   With frmSplashScreen
      .Caption = "MiniCut2d Software - Bienvenido !"
      .lblCliquez = "Haga clic en el tipo de máquina que utiliza. "
      .cmdPasDeMachine.Caption = "Si usted no tiene una máquina, haga clic aquí"
   End With
   '**** traduction du A propos ****
   With frmAboutAndSettings
      .Caption = "Parámetros"
      .lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
      .lblTitle.Caption = TypeMachine & " Software"
      .lblParametres.Caption = "Recorrido total X : " & Format(CourseX, "#####") & "mm - Recorrido total Y : " & Format(CourseY, "#####") & "mm." & _
               vbCrLf & "Desfase inter/origen X : " & Format(MmToOriXG, "#0.0##") & "mm, ósea " & Format(NbrPasToOriXG, "#######") & " micro pasos." & _
               vbCrLf & "Desfase inter/origen YG : " & Format(MmToOriYG, "#0.0##") & "mm, ósea " & Format(NbrPasToOriYG, "#######") & " micro pasos." & _
               vbCrLf & "Desfase inter/origen YD : " & Format(MmToOriYD, "#0.0##") & "mm, ósea " & Format(NbrPasToOriYD, "#######") & " micro pasos."
      .lblTraduction.Caption = "Traducción:" & vbCrLf & vbCrLf & _
                              "Ingles : aiR-C²/Hugh Potter" & vbCrLf & _
                              "Aleman : Charles Wittmer" & vbCrLf & _
                              "Espanol : Enrique Iglesias"
      .frmParametres.Caption = "Parámetros"
      .frmAPropos.Caption = "Acerca de"
      .frmModeExpert.Caption = "El funcionamiento normal Normal / Experto"
      .chkActiverLeChangementDeMode.Caption = "Activar"
      .optNormalExpert(0).Caption = "Modo Normal"
      .optNormalExpert(1).Caption = "Modo Experto"
   End With
   '**** traduction de la form "Découpe Inactive" *****
   With frmDecoupeInactive
      .lblDecoupeInactive.Caption = "El acceso a esta parte del software es imposible" & vbCrLf & "ya que no se detecta el interface."
   End With
   '**** traduction de la form "Paramètres machine"
   With frmParametres
      .Caption = "Parámetros de la máquina"
   End With
   '**** traduction de la form principale ****
   With frmMiniCut2d
      .cmdLangue.Picture = frmImages.imgDrapeauEspagnol.Picture   'le drapeau
      .SSTab1.TabCaption(0) = "Creación"
      .SSTab1.TabCaption(1) = "Recorte"
      .lblBoiteOutils.Caption = " Caja de herramientas "
      .lblAlignerAjuster.Caption = " Alinear y ajustar "
      .lblDimensionsBloc.Caption = " Dimensiones del bloque "
      .lblTitreChauffe.Caption = " Calor "
      .lblTrajet.Caption = " Trayecto "
      .lblDecoupe.Caption = " Recorte "
      .lblFil.Caption = " Hilo "
      .lblCadreDecoupe.Caption = " Recorte "
      .lblStopReprise.Caption = " Parar / Reanudar "
      .lblPiloterLeFil.Caption = " Conducir el hilo "
      .frameEntreeBloc.Caption = "Entrada"
      .frameSortieBloc.Caption = "Salida"
      .frameSimulation.Caption = "Simulación"
      .frameManuel.Caption = "Pilotar"
      .frameProcedures.Caption = "Posiciones"
      .frameDecalage.Caption = "Desplazar el hilo"
      .frameInformation.Caption = "Información"
      .frameAction.Caption = "Lanzar el recorte"
      .frameAnnulationStop.Caption = "Anulación - Parar/Reanudar"
      .frameChauffeEnCoursDecoupe.Caption = "Calentamiento"
      .frameInformationStop.Caption = "Información"
      .frameModifierLaChauffe.Caption = "Modificar el calentamiento"
      .frameRetourOrigine.Caption = "Volver a origen"
      .frameTrajetRetour.Caption = "Trayecto"
      .frameReprendre.Caption = "Acabar el recorte"
      .frameAnnulationReprise.Caption = "Cancelar"
      .frameChauffeFil.Caption = "Calentar"
      .frameInformationFil.Caption = "Información"
      .frameFilManuel.Caption = "Desplazar"
      .frameAnnulationFil.Caption = "Cancelar / Salir"
      .cmdNouveauProjet.ToolTipText = "Nuevo proyecto"
      .cmdOuvrirFichierSequ.ToolTipText = "Abrir un proyecto"
      .cmdSauver(0).ToolTipText = "Guardar como..."
      .cmdSauver(1).ToolTipText = "Guardar"
      .cmdLangue.ToolTipText = "Cambiar el idioma"
      .cmdSettings.ToolTipText = "Parámetros"
      .cmdRafraichir.ToolTipText = "Actualizar"
      .cmdEffacerFichier.ToolTipText = "Borrar (basura)"
      .cmdImporterProfil.ToolTipText = "Importar un fichero en la biblioteca"
      .cmdSimulation.ToolTipText = "Ver el desplazamiento del hilo"
      .optOutils(4).ToolTipText = "Medir"
      .optOutils(5).ToolTipText = "Cortar un trayecto"
      .optOutils(2).ToolTipText = "Estirar"
      .optOutils(1).ToolTipText = "Rotar"
      .optOutils(0).ToolTipText = "Desplazar"
      .optOutils(6).ToolTipText = "Cambiar el punto de entrada"
      .cmdUndo(0).ToolTipText = "Deshacer"
      .cmdUndo(1).ToolTipText = "Rehacer"
      .cmdPoubelle.ToolTipText = "Suprimir"
      .cmdInsererPoint.ToolTipText = "Insertar un punto"
      .cmdDupliquer.ToolTipText = "Duplicar"
      .cmdMiroir.ToolTipText = "Espejo"
      .cmdInverser.ToolTipText = "Invertir sentido"
      .cmdAligner(0).ToolTipText = "Alinear abajo"
      .cmdAligner(1).ToolTipText = "Alinear en medio"
      .cmdAligner(2).ToolTipText = "Alinear arriba"
      .cmdAligner(3).ToolTipText = "Alinear a la izquierda"
      .cmdAligner(4).ToolTipText = "Alinear en el centro"
      .cmdAligner(5).ToolTipText = "Alinear a la derecha"
      .chkCentrer(0).ToolTipText = "Escalar con el bloque"
      .chkCentrer(1).ToolTipText = "Centrar horizontalmente"
      .chkCentrer(2).ToolTipText = "Centrar verticalmente"
      .chkVoirPoints.ToolTipText = "Mostrar los puntos"
      .chkCouleurProfils.ToolTipText = "Alternar colores de perfiles"
      .cmdAgrandirRetrecir.ToolTipText = "Tamaño de la ventana"
      .chkZoomProjet.ToolTipText = "Zoom bloque"
      .pctZoomInfo.ToolTipText = "Zoom: click derecho + rueda del ratón o las flechas arriba y abajo"
      .cmdGestionMatiere(0).ToolTipText = "Nueva materia"
      .cmdGestionMatiere(1).ToolTipText = "Remplazar el valor de calentamiento"
      .cmdGestionMatiere(2).ToolTipText = "Suprimir la materia"
      .cmdDecouper.ToolTipText = "Recortar el proyecto"
      .cmdDeplacementsManuels.ToolTipText = "Pilotar el hilo en directo"
      .cmdPlierLePortique.ToolTipText = "Ir a posición de reposo"
      .cmdRetourOrigine.ToolTipText = "Traer hilo a origen"
      .optDecalage(1).ToolTipText = "-0.5mm"
      .optDecalage(2).ToolTipText = "0mm"
      .optDecalage(3).ToolTipText = "0.5mm"
      .cmdFaireOrigineAvantDecoupe.ToolTipText = "Lanzar el recorte"
      .cmdAnnulerDecoupe.ToolTipText = "Cancelar el recorte"
      .cmdSTOP.ToolTipText = "Parada emergencia!”"
      .optTrajetRetour(0).ToolTipText = "En diagonal"
      .optTrajetRetour(1).ToolTipText = "Por la izquierda"
      .optTrajetRetour(2).ToolTipText = "Por arriba"
      .cmdLancerRetourApresStop.ToolTipText = "Lanzar el retorno"
      .cmdStopRetourApresReprise.ToolTipText = "Parada emergencia!”"
      .cmdRepriseDecoupe.ToolTipText = "Acabar el recorte"
      .cmdAnnulerReprise.ToolTipText = "Cancelar todo"
      .optChauffe(0).ToolTipText = "Calentar"
      .optChauffe(1).ToolTipText = "Parar el calentamiento"
      .optGoManuel(0).ToolTipText = "Lanzar el movimiento"
      .optGoManuel(1).ToolTipText = "Parar el movimiento"
      .cmdAnnulerFil.ToolTipText = "Salir"
      .optHomeY.ToolTipText = "Origen vertical"
      .optHomeX.ToolTipText = "Origen horizontal"
      .optAnnulerHome.ToolTipText = "Salir"
   End With
   '**** les MsgBox ****
   ReDim Message(1 To 2, 1 To 1)
   'MessageBox n°1
   Message(Corps, 1) = "Dos puntos consecutivos están confundidos." & vbCrLf & "Imposible definir un desfase."
   Message(Titre, 1) = "Calculo imposible."
   'MessageBox n°2
   ReDim Preserve Message(Corps To Titre, 1 To 2)
   Message(Corps, 2) = "El directorio \Biblioteca no se encuentra, se creara pero quedara vacio. Tendrá que rellenarlo!"
   Message(Titre, 2) = "Creación de la biblioteca"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 3)
   Message(Corps, 3) = "El fichero ""MiniCut2d Software.ini"" contiene llaves que empiezan por ""NbrPasToOri...""." & vbCrLf & _
                        "Estas llaves ya no son validas y serán substituidas por las " & vbCrLf & _
                        "nuevas llaves de tipo ""MmToOri..."" para que MiniCut2d Software pueda funcionar correctamente."
   Message(Titre, 3) = "Antigua versión del .ini"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 4)
   Message(Corps, 4) = "El valor de calentamiento memorizado con una materia, sobrepasa el valor máximo posible de la maquina." _
                        & vbCrLf & "El calentamiento será sujeto al valor máximo."
   Message(Titre, 4) = "Calentamiento demasiado alto"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 5)
   Message(Corps, 5) = "El fichero IPL5XComm.dll presente en su ordenador es antiguo." & vbCrLf & _
                        "MiniCut2d Software no puede funcionar." & vbCrLf & _
                        "Tiene que descargar la ultima versión y actualizar."
   Message(Titre, 5) = "Comunicación con la maquina imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 6)
   Message(Corps, 6) = "El arranque del interface USB da problemas" & vbCrLf & _
                        "El acceso a las funciones de recorte puede no hacerse"
   Message(Titre, 6) = "Problema de arranque del interpolador."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 7)
   Message(Corps, 7) = "El fichero IPL5XCom.dll no se encuentra," & vbCrLf & _
                        "Debe instalarlo en su ordenador, en el lugar correcto." & vbCrLf & _
                        "El recorte se desactivara."
   Message(Titre, 7) = "Comunicación con la maquina imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 8)
   Message(Corps, 8) = "Parece que la MiniCut2d no esta en posición de guardado." & vbCrLf & _
                        "Quiere lanzar el procedimiento de guardado del platillo y del hilo, antes de salir del programa?"
   Message(Titre, 8) = "Solicitud de cierre de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 9)
   Message(Corps, 9) = "Corte la alimentación de la MiniCut2d antes de salir del programa." & vbCrLf & _
   "Verifique igualmente que su proyecto está guardado." & vbCrLf & vbCrLf & _
   "Confirma el cierre de la aplicación ?"
   Message(Titre, 9) = "Solicitud de cierre de MiniCut2d Software"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 10)
   Message(Corps, 10) = "Imposible presentar este fichero."
   Message(Titre, 10) = "Error de lectura del fichero"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 11)
   Message(Corps, 11) = "Desea salvaguardar el proyecto en curso?"
   Message(Titre, 11) = "Nuevo proyecto"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 12)
   Message(Corps, 12) = "Extensión del fichero no valida. MiniCut2d Software solo abre los .mnc, .dxf, .dat, .plt, .eps (y .txt con las coordenadas de los puntos y coma separados)."
   Message(Titre, 12) = "Apertura imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 13)
   Message(Corps, 13) = "Ha cortado sobre el primer punto de un perfil, es imposible cortar en este sitio."
   Message(Titre, 13) = "Operación imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 14)
   Message(Corps, 14) = "Ha cortado sobre el último punto de un perfil, es imposible cortar en este sitio."
   Message(Titre, 14) = "Operación imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 15)
   Message(Corps, 15) = "Debe introducir un numero decimal positivo o negativo."
   Message(Titre, 15) = "Error de tecleo"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 16)
   Message(Corps, 16) = "La operación solicitada es imposible, una dimensión es igual a 0."
   Message(Titre, 16) = "Calculo imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 17)
   Message(Corps, 17) = "La zona útil de la maquina ha sido sobrepasada, el proyecto es cortado automáticamente."
   Message(Titre, 17) = "Sobrepaso de los recorridos"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 18)
   Message(Corps, 18) = "El nombre de este material ya existe, desea sobrescribir los valores?"
   Message(Titre, 18) = "Modificación de un material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 19)
   Message(Corps, 19) = "Esta línea no puede ser borrada."
   Message(Titre, 19) = "Operación imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 20)
   Message(Corps, 20) = "Validar el borrado de este material?"
   Message(Titre, 20) = "Borrado de un material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 21)
   Message(Corps, 21) = "Hay un problema: el valor de calentamiento para el material seleccionado, no está en los límites previstos."
   Message(Titre, 21) = "Valor incorrecto"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 22)
   Message(Corps, 22) = "Inicialización de la " & TypeMachine
   Message(Titre, 22) = "Interface USB detectada"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 23)
   Message(Corps, 23) = "Ejecutar el procedimiento de retorno a posición de descanso?"
   Message(Titre, 23) = "Validación de seguridad"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 24)
   Message(Corps, 24) = "El hilo esta en posición de descanso."
   Message(Titre, 24) = "Retorno realizado"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 25)
   Message(Corps, 25) = "La operación a sido anulada."
   Message(Titre, 25) = "Parada de emergencia"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 26)
   Message(Corps, 26) = "El bucle de los interruptores de origen esta abierto." & vbCrLf & _
                        "No es normal." & vbCrLf & "Se deben liberar los interruptores haciendo girar los motores manualmente" & _
                        vbCrLf & "(Cortar alimentación si se resisten)."
   Message(Titre, 26) = "Salida de recorridos"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 27)
   Message(Corps, 27) = "El bucle de los interruptores de final de recorrido está abierto." & vbCrLf & _
                        "No es normal." & vbCrLf & "Se deben liberar los interruptores haciendo girar los motores manualmente" & _
                        vbCrLf & "(Cortar alimentación si se resisten)."
   Message(Titre, 27) = "Salida de recorridos"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 28)
   Message(Corps, 28) = "Desplazamiento en posición de guardado?"
   Message(Titre, 28) = "Validación de seguridad"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 29)
   Message(Corps, 29) = "Posición de doblado alcanzada."
   Message(Titre, 29) = "Movimiento realizado"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 30)
   Message(Corps, 30) = "Los interruptores de la mesa actual no están activados."
   Message(Titre, 30) = "Imposible realizar el procedimiento."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 31)
   Message(Corps, 31) = "Un interruptor de final de recorrido esta abierto, el procedimiento no puede iniciarse." & vbCrLf & _
                        "Libere todos los interruptores haciendo girar los motores con la mano."
   Message(Titre, 31) = "Imposible realizar el procedimiento."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 32)
   Message(Corps, 32) = "Un interruptor de origen esta abierto, el procedimiento no puede iniciarse." & vbCrLf & _
                        "Libere todos los interruptores haciendo girar los motores con la mano."
   Message(Titre, 32) = "Imposible realizar el procedimiento."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 33)
   Message(Corps, 33) = "Hay un  problema, es necesario que haya más de 2mm para escapar de los interruptores. Probemos de nuevo..."
   Message(Titre, 33) = "Búsqueda del origen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 34)
   Message(Corps, 34) = "Los interruptores están aun abiertos. la búsqueda del origen no es posible, debe controlar la maquina."
   Message(Titre, 34) = "Búsqueda del origen"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 35)
   Message(Corps, 35) = "Parada calentamiento y movimiento por apretado del botón STOP de la MiniCut2d."
   Message(Titre, 35) = "Parada de emergencia"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 36)
   Message(Corps, 36) = "Ha habido un  problema en el cálculo del tiempo del recorte."
   Message(Titre, 36) = "Anulación"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 37)
   Message(Corps, 37) = "No hay trazado cargado." & vbCrLf & "Imposible acceder a los parámetros de recorte."
   Message(Titre, 37) = "Operación imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 38)
   Message(Corps, 38) = "Si los valores de la memoria por los valores se muestra?"
   Message(Titre, 38) = "Material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 39)
   Message(Corps, 39) = "Recorte acabado, hilo en posición de descanso, calentamiento apagado."
   Message(Titre, 39) = "MiniCut2d disponible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 40)
   Message(Corps, 40) = "Cuidado el calentamiento no está apagado!"
   Message(Titre, 40) = "Problema de comunicación con la maquina"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 41)
   Message(Corps, 41) = "El botón STOP ha sido presionado antes de el principio de movimiento." & vbCrLf & "El calentamiento a sido parado."
   Message(Titre, 41) = "Parada durante el calentamiento"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 42)
   Message(Corps, 42) = "El botón STOP ha sido presionado, operación anulada."
   Message(Titre, 42) = "Parada de emergencia"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 43)
   Message(Corps, 43) = "Atención, el hilo no está en posición de descanso."
   Message(Titre, 43) = "Parada en la zona útil"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 44)
   Message(Corps, 44) = "El hilo está en posición de descanso."
   Message(Titre, 44) = "Retorno realizado"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 45)
   Message(Corps, 45) = "Un interruptor está abierto, el recorte no puede iniciarse." & vbCrLf & _
                        "Libere todos los interruptores haciendo girar las varillas rosca, con la mano."
   Message(Titre, 45) = "Imposible hacer el recorte."
   '
   ReDim Preserve Message(Corps To Titre, 1 To 46)
   Message(Corps, 46) = "Un interruptor de puesta a origen está abierto. No es normal, libere la pieza y tome de nuevo el punto de origen."
   Message(Titre, 46) = "Parada en medio de recorte!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 47)
   Message(Corps, 47) = "Un interruptor de final de recorrido está abierto. No es normal, libere la pieza y tome de nuevo el punto de origen."
   Message(Titre, 47) = "Parada en medio de recorte!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 48)
   Message(Corps, 48) = "El botón de Parada de emergencia a sido apretado, anulación del recorte."
   Message(Titre, 48) = "Parada en medio de recorte!"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 49)
'  Message(Corps, 49) = "Tenga en cuenta que la calefacción y el calentamiento del material seleccionado actualmente tienen diferentes valores." & vbCrLf & _
'                       "Antes de salvaguardar, debo reemplazar la almacenada para el caso por el calentador de calefacción actual?"

'  Message(Titre, 49) = "Salvaguardado del material usado"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 50)
   Message(Corps, 50) = "Copia anulada"
   Message(Titre, 50) = "Copia de ficheros"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 51)
   Message(Corps, 51) = "Los ficheros han sido copiados en el lugar solicitado."
   Message(Titre, 51) = "Copia de ficheros"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 52)
   Message(Corps, 52) = "La zona útil de la maquina esta sobrepasada, el proyecto será cortado automáticamente."
   Message(Titre, 52) = "Sobrepaso de los recorridos"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 53)
   Message(Corps, 53) = "No hay contorno transferido al bloque."
   Message(Titre, 53) = "Operación imposible"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 54)
   Message(Corps, 54) = "Operación imposible."
   Message(Titre, 54) = "Anulación"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 55)
   Message(Corps, 55) = "Haga clic derecho para ampliar la ubicación" & vbCrLf & _
                        "a continuación, la rueda del ratón" & vbCrLf & _
                        "o hacia arriba y abajo las teclas de flecha."
   Message(Titre, 55) = "Zoom"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 56)
   Message(Corps, 56) = "El valor de velocidad memorizado con una materia, sobrepasa el valor máximo posible de la maquina." _
                        & vbCrLf & "El velocidad será sujeto al valor máximo."
   Message(Titre, 56) = "Velocidad demasiado alto"
   '   '
   ReDim Preserve Message(Corps To Titre, 1 To 57)
   Message(Corps, 57) = "El nombre de este material ya está en la memoria (modo experto)." & vbCrLf & _
                        "Desea sobrescribir los valores?"
   Message(Titre, 57) = "Modificación de un material"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 58)
   Message(Corps, 58) = "Hay un problema: el valor de velocidad para el material seleccionado, no está en los límites previstos."
   Message(Titre, 58) = "Valor incorrecto"
   '
   ReDim Preserve Message(Corps To Titre, 1 To 59)
   Message(Corps, 59) = "La velocidad de este material no se puede cambiar. Crear un nuevo material."
   Message(Titre, 59) = "Modificación imposible."
   '

   '**** Les Labels des outils, informations, avertissements ****
   ReDim Label(1 To 1)
   Label(1) = "Haga clic en un punto para cortarlo en dos trayectos."
   '
   ReDim Preserve Label(1 To 2)
   Label(2) = "Haga clic en el nuevo punto de entrada"
   '
   ReDim Preserve Label(1 To 3)
   Label(3) = "Angulo (°) :"
   '
   ReDim Preserve Label(1 To 4)
   Label(4) = "Ctrl=centrar, Shift=proporción, Alt=espejo"
   '
   ReDim Preserve Label(1 To 5)
   Label(5) = "Selección : "
   '
   ReDim Preserve Label(1 To 6)
   Label(6) = " No hay contorno visible. Seleccione un fichero en la biblioteca. "
   '
   ReDim Preserve Label(1 To 7)
   Label(7) = " No hay contornos bajo el puntero. Doble-clic " & vbCrLf & "  en un contorno o bien desliza-lo en el bloque."
   '
   ReDim Preserve Label(1 To 8)
   Label(8) = " No hay contorno transferido al bloque. Seleccione un fichero de la " & vbCrLf & "  biblioteca y haga doble clic en su contorno o deslice-lo aquí."
   '
   ReDim Preserve Label(1 To 9)
   Label(9) = "según X : "  ' il s'agit d'une dimensions suivant X, le mot à traduire est "Suivant"
   '
   ReDim Preserve Label(1 To 10)
   Label(10) = " mm - según Y : "  ' il s'agit d'une dimensions suivant Y
   '
   ReDim Preserve Label(1 To 11)
   Label(11) = "Retorno automático a la posición de descanso"
   '
   ReDim Preserve Label(1 To 12)
   Label(12) = "Búsqueda de los interruptores"
   '
   ReDim Preserve Label(1 To 13)
   Label(13) = "Desfase hacia la posición de descanso"
   '
   ReDim Preserve Label(1 To 14)
   Label(14) = "Posición de descanso"
   '
   ReDim Preserve Label(1 To 15)
   Label(15) = "Preparación para el guardado"
   '
   ReDim Preserve Label(1 To 16)
   Label(16) = "Desplazamiento hacia la posición de pliegue"
   '
   ReDim Preserve Label(1 To 17)
   Label(17) = "Posición de descanso"
   '
   ReDim Preserve Label(1 To 18)
   Label(18) = "EL HILO SE DESPLAZA"
   '
   ReDim Preserve Label(1 To 19)
   Label(19) = "EL HILO SE CALIENTA"
   '
   ReDim Preserve Label(1 To 20)
   Label(20) = "HILO TOMANDO TEMPERATURA"
   '
   ReDim Preserve Label(1 To 21)
   Label(21) = "del cual "
   '
   ReDim Preserve Label(1 To 22)
   Label(22) = " s. de puesta en temperatura."
   '
   ReDim Preserve Label(1 To 23)
   Label(23) = "(Calentamiento"
   '
   ReDim Preserve Label(1 To 24)
   Label(24) = "Procedimiento de retorno a origen"
   '
   ReDim Preserve Label(1 To 25)
   Label(25) = "Parada sobre segmento n° "
   '
   ReDim Preserve Label(1 To 26)
   Label(26) = "Retorno a posición de descanso"
   '
   ReDim Preserve Label(1 To 27)
   Label(27) = "Posición de descanso"
   '
   ReDim Preserve Label(1 To 28)
   Label(28) = "Hilo tomando temperatura"
   '
   ReDim Preserve Label(1 To 29)
   Label(29) = "Recorte segmento n°"
   '
   ReDim Preserve Label(1 To 30)
   Label(30) = "Pilotaje del hilo"
   '
   ReDim Preserve Label(1 To 31)
   Label(31) = "velocidad"
   '
   ReDim Preserve Label(1 To 32)
   Label(32) = "Duración :"
   '

End Sub
