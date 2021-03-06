' Script de t�l�chargement des comics Cyanide and Hapiness
' Ce script doit �tre ex�cut� avec la commande cscript et non wscript (par d�faut)

'region Constantes
'***********************************************
'|                 Constantes                  |
'***********************************************
'region gestion des erreurs
Dim ERREUR_TECHNIQUE: ERREUR_TECHNIQUE = 1
Dim ERREUR_IMG_INEXISTANTE: ERREUR_IMG_INEXISTANTE = 100
'endregion gestion des erreurs

'region gestion des retours OK
Dim RETOUR_SANS_EXECUTION: RETOUR_SANS_EXECUTION = 10
Dim RETOUR_OK: RETOUR_OK = 0
'endregion gestion des retours OK

'region gestion des arguments
Dim ARG_SILENTMODE: ARG_SILENTMODE = "s"
Dim ARG_HELP: ARG_HELP = "h"
'endregion gestion des arguments

'region stockage
' Emplacement des fichiers de log
Dim LOGROOT: LOGROOT = "D:\temp\"
' Emplacement des images t�l�charg�es
Dim IMGROOT: IMGROOT = "D:\temp\CandH\"
'endregion stockage

'region URLs
' URL principale
Dim URL_MAIN: URL_MAIN = "http://explosm.net/comics/"

' URL du premier comic
Dim URL_OLDEST: URL_OLDEST = "http://explosm.net/comics/oldest"

' URL du dernier comic
Dim URL_LATEST: URL_LATEST = "http://explosm.net/comics/latest"

' URL de base de t�l�chargement des images
Dim URL_DLROOT: URL_DLROOT = "http://files.explosm.net/comics/"
'endregion URL

'region Identification des �l�ments
' ID de l'�l�ment contenant l'image dans le cas d'une page normale
Dim IMG_ID: IMG_STD_ID = "main-comic"
'endregion Identification des �l�ments

'region Gestion des expression r�guli�res
' Pattern d'extraction du chemin de l'image actuelle, depuis la page compl�te
' Exemple de chaine : <input id="permalink" type="text" value="http://explosm.net/comics/4376/" onclick="this.select()">
Dim PATTERN_URL_CUR: PATTERN_URL_CUR = "id=""permalink"" .* value=""(.*)"" .*"
' Expression r�guli�re d'extraction du lien de l'image actuelle
Set regExtractURLCur = New RegExp
regExtractURLCur.Pattern = PATTERN_URL_CUR

' Pattern d'extraction du chemin du lien de l'image pr�c�dente, depuis la page compl�te
' Exemple de chaine : <li><a href="/comics/4375/" class="previous-comic " title="Previous comic"><img src="/img/nav-button_previous@2x.png"/></a></li>
Dim PATTERN_URL_PREV: PATTERN_URL_PREV = "href=""/(.*)"" .* title=""Previous comic"""
' Expression r�guli�re d'extraction du lien de l'image pr�c�dente
Set regExtractURLPrev = New RegExp
regExtractURLPrev.Pattern = PATTERN_URL_PREV

' Pattern d'extraction du chemin du lien de l'image suivante, depuis la page compl�te
' Exemple de chaine : <li><a href="/comics/4377/" class="next-comic " title="Next comic"><img src="/img/nav-button_next@2x.png"/></a></li>
Dim PATTERN_URL_NEXT: PATTERN_URL_NEXT = "href=""/(.*)"" .* title=""Next comic"""
' Expression r�guli�re d'extraction du lien de l'image suivante
Set regExtractURLNext = New RegExp
regExtractURLNext.Pattern = PATTERN_URL_NEXT

' Pattern d'extraction du num�ro de l'image, depuis un lien de page d'image, extrait avec regExtractURLCur/Prev/Next
' Exemple de chaine : cf. href dans les extractions de chemin
Dim PATTERN_NUM_IMG: PATTERN_NUM_IMG = "comics/(.*)/"
' Expression r�guli�re d'extraction du num�ro de l'image
Set regExtractNumImg = New RegExp
regExtractNumImg.Pattern = PATTERN_NUM_IMG

' Expression r�guli�re d'extraction de l'URL relative d'une image sur une page
Dim PATTERN_URL_REL_IMG: PATTERN_URL_REL_IMG = "id=""main-comic"" src=""\/\/files\.explosm\.net\/comics\/(.*)"""
Set regExtractURLRelImg = new RegExp
regExtractURLRelImg.Pattern = PATTERN_URL_REL_IMG

' Expression r�guli�re d'extraction du nom de l'image � partir de son URL compl�te
Dim PATTERN_NAME_IMG: PATTERN_NAME_IMG = ".*\/([^?]*)(\?.*)?"
Set regExtractNameImg = new RegExp
regExtractNameImg.Pattern = PATTERN_NAME_IMG

' Expression r�guli�re d'extraction de la date pour le comic actuel
Dim PATTERN_DATE_BRUTE: PATTERN_DATE_BRUTE = "<h3 .*><a .*>(.*)<\/a><\/h3>"
Set regExtractDateBrute = New RegExp
regExtractDateBrute.Pattern = PATTERN_DATE_BRUTE

' Expression r�guli�re d'extraction des �l�ments d'une date brute
Dim PATTERN_ELEM_DATE_BRUTE: PATTERN_ELEM_DATE_BRUTE = "[0-9]{1,4}\.[0-9]{1,2}\.[0-9]{1,2}"
Set regExtractElemDateBrute = New RegExp
regExtractElemDateBrute.Pattern = PATTERN_ELEM_DATE_BRUTE
'endregion Gestion des expressions r�guli�res

'region Objets divers
' objet permettant la gestion d'appel HTTP
Dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP")
' stream qui va servir pour l'�criture de fichier
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
'endregion Objets divers
'endregion Constantes

'region Variables
'***********************************************
'|                 Variables                   |
'***********************************************

'region variables d'ex�cution
Dim silentMode
'endregion variables d'ex�cution

'region Intervalle de recherche
' Intervalle de recherche inf�rieur (num�ro et date)
Dim rMin, dMin
' Intervalle de recherche sup�rieur (num�ro et date)
Dim rMax, dMax
'endregion Intervalle de recherche

'region URLs
' URL des images
Dim URLImg, URLImgPrev, URLImgNext
' URL des pages vers les images
Dim URLPageImg, URLPageImgPrev, URLPageImgNext
' URL relative des pages vers les images
Dim URLRelImg, URLRelImgPrev, URLRelImgNext
'endregion URLs

'region Identification des �l�ments
' Num�ro des images
Dim numImgOldest, dateImgOldest, numImgLatest, dateImgLatest, numImgPrev, numImgNext
'endregion Identification des �l�ments
'endregion Variables

'region M�tier
'***********************************************
'|                 M�tier                      |
'***********************************************
'region Main

Call CalcIntervalleMaxPerma

Call ExtractArguments

'Call GetImgsIntervalle(0, 40)

'endregion Main

'region Appel
'*********************************************************
' Purpose: R�cup�re la page indiqu� par le param�tre p_URLPageImg
' Inputs: p_URLPageImg : l'URL de la page pour laquelle r�cup�rer le contenu
'         avecLog : si 1, log de l'appel dans le fichier d�clar� ; sinon, l'appel est silencieux
' Returns: Le contenu de la page sous forme d'objet, dans l'objet d�clar� dans la fonction
'*********************************************************
Function Appel(p_URLPageImg, avecLog)
  xHttp.Open "GET", p_URLPageImg, False
  xHttp.Send
  If(avecLog = 1) Then
    ' �criture dans fichier
    With bStrm
      .type = 1 '//binary
      .open
      .write xHttp.responseBody
      .savetofile LOGROOT & "CandH_dl.html", 2 '//overwrite
      .close
    End With
  End If
End Function
'endregion Appel

'region Extraction des liens de page, num�ros et noms d'images
'*********************************************************
' Purpose: Extraction du lien vers la page de l'image, en fonction de la regEx pass�e en param�tre
' Inputs: regExLien : expression r�guli�re d'un lien en particulier
' Returns: Si un lien est extrait, le lien lui-m�me ; sinon (pas de lien trouv�), erreur
'*********************************************************
Function ExtraitLien(regExLien)
  Set objMatches = regExLien.Execute(xHttp.responseText)
  Dim nbMatches: nbMatches = objMatches.Count
  If(nbMatches > 0) Then
    Dim result: result = objMatches(0)
    ExtraitLien = objMatches(0).SubMatches(0)
  Else
    ExtraitLien = ERREUR_IMG_INEXISTANTE
  End If
End Function

'*********************************************************
' Purpose: Extraction du num�ro de l'image en fonction de son URL relative
' Inputs: regExImg : expression r�guli�re d'extraction d'un num�ro d'image dans une URL relative
'         p_URLRelativeImg : URL relative (sans la partie URLROOT) de l'image
' Returns: le num�ro de l'image extrait
'*********************************************************
Function ExtraitNumImg(regExImg, p_URLRelativeImg)
  ' TODO : gestion d'erreur si le num�ro ne parvient pas � �tre extrait
  Set objMatches = regExImg.Execute(p_URLRelativeImg)
  Dim nbMatches: nbMatches = objMatches.Count
  Dim result: result = objMatches(0)
  ExtraitNumImg = objMatches(0).SubMatches(0)
End Function

'*********************************************************
' Purpose: Extraction de la date du comic
' Inputs: regExImg : expression r�guli�re d'extraction d'une date dans une URL relative
' Returns: La date de la page extraite
'*********************************************************
Function ExtraitDateImg(regExImg)
  ' TODO : gestion d'erreur si la date ne parvient pas � �tre extraite
  Set objMatches = regExImg.Execute(xHttp.responseText)
  Dim nbMatches: nbMatches = objMatches.Count
  Dim result: result = objMatches(0)
  ExtraitDateImg = objMatches(0).SubMatches(0)
End Function

'*********************************************************
' Purpose: Extraction du nom de l'image en fonction de son URL relative
' Inputs: p_URLImg : l'URL relative (sans la partie URLROOT) de l'image
' Returns: le nom de l'image extrait
'*********************************************************
Function ExtractImgName(p_URLImg)
  Set objMatches = regExtractNameImg.Execute(p_URLImg)
  Dim nbMatches: nbMatches = objMatches.Count
  Dim result: result = objMatches(0)
  ExtractImgName = objMatches(0).SubMatches(0)
End Function
'endregion Extraction des liens, num�ros et noms d'images

'region Gestion des intervalles minimal et maximal
'*********************************************************
' Purpose: Calcul de l'intervalle maximal en fonction du lien permanent (permalink)
' Returns: l'intervalle maximal ; dans la m�me temps, la variable globale est aussi
'          valoris� avec celui-ci
'*********************************************************
Sub CalcIntervalleMaxPerma()

  ' Appel initial vers le dernier comic en date
  Call Appel(URL_LATEST, 1)

  ' Extraction du lien � partir du permalien de la page
  URLPageImg = ExtraitLien(regExtractURLCur)

  ' Extraction du num�ro de l'image sur la page qui d�terminera l'intervalle maximal
  numImgLatest = ExtraitNumImg(regExtractNumImg, URLPageImg)

  ' Extraction de la date de la page actuelle
  dateImgLatest = ExtraitDateImg(regExtractDateBrute)

  ' Appel vers le premier comic
  Call Appel(URL_OLDEST, 1)
'
  ' Extraction du lien � partir du permalien de la page
  URLPageImg = ExtraitLien(regExtractURLCur)

  ' Extraction du num�ro de l'image sur la page qui d�terminera l'intervalle maximal
  numImgOldest = ExtraitNumImg(regExtractNumImg, URLPageImg)

  ' Extraction de la date brute de la page actuelle
  dateImgOldest = ExtraitDateImg(regExtractDateBrute)

  ' Valorisation de l'intervalle maximal
  rMax = numImgLatest

  ' Valorisation de la date maximale
  dMax = DateBruteToDate(dateImgLatest)

  ' Valorisation de l'intervalle minimal
  rMin = numImgOldest

  ' Valorisation de la date minimale
  dMin = DateBruteToDate(dateImgOldest)
End Sub

'*********************************************************
' Purpose: R�cup�ration des images dans l'intervalle renseign�
' Inputs: dateMin : date minimum
'         dateMax : date maximum
'*********************************************************
Function GetImgsIntervalleDate(dateMin, dateMax)
  ' Il n'est pas possible de t�l�charger une image directement depuis sa date ;
  ' on t�l�charge donc la premi�re image en se basant sur les archives et on consid�re
  ' que l'on change de jour en changeant le num�ro de l'image

  Dim numMin, numMax

  numMin = GetImgNumFromDate(dateMin)
  numMax = GetImgNumFromDate(dateMax)
  Call GetImgsIntervalle(numMin, numMax)
End Function

'*********************************************************
' Purpose: R�cup�ration des images dans l'intervalle renseign�, pour chaque
'          num�ro d'image disponible sur le site
' Inputs: numImgMin : l'intervalle bas (num�ro d'image minimum)
'         numImgMax : l'intervalle haut (num�ro d'image maximum)
'*********************************************************
Function GetImgsIntervalle(numImgMin, numImgMax)

  ' Si l'intervalle n'est pas correct, erreur
  If (numImgMin > numImgMax) Then
    WScript.Echo "Erreur : l'intervalle minimal doit �tre inf�rieur ou �gal � l'intervalle maximal"
    WScript.Quit ERREUR_TECHNIQUE
    ' Sinon, on boucle dans l'intervalle
  Else
    For counter = numImgMin To numImgMax
      DlImg(counter)
    Next
  End If
End Function
'endregion Gestion des intervalles minimal et maximal

'region Gestion des t�l�chargements

'*********************************************************
' Purpose: R�cup�ration du num�ro d'une image en fonction de sa date
' Inputs: dateImg : la date de l'image
' Return: le num�ro de l'image
'*********************************************************
Function GetImgNumFromDate(dateImg)
  Dim URLRootArchive: URLRootArchive = "http://explosm.net/comics/archive/"
  Dim annee: annee = DatePart("yyyy", dateImg)
  Dim mois: mois = DatePart("m", dateImg)
  If (Len(mois) = 1) Then
    mois = "0" & mois
  End If
  Dim jour: jour = DatePart("d", dateImg)
  If (Len(jour) = 1) Then
    jour = "0" & jour
  End If

  ' Appel de l'URL pour les archives du mois demand�
  Call Appel(URLRootArchive & annee & "/" & mois, 1)

  ' Extraction du num�ro
  Dim patternExtractNum: patternExtractNum = "<a href=""(.*)\/([0-9]*)\/"">" & annee & "." & mois & "." & jour
  Set regExtractNum = New RegExp
  regExtractNum.Pattern = patternExtractNum

  Set objMatches = regExtractNum.Execute(xHttp.ResponseText)
  Dim nbMatches: nbMatches = objMatches.Count
  If(nbMatches > 0) Then
    Dim result: result = objMatches(0)
    GetImgNumFromDate = objMatches(0).SubMatches(1)
  Else
    WScript.Echo "Erreur : il n'y a pas de comic en date du " & dateImg & "."
    WScript.Quit ERREUR_TECHNIQUE
  End If
End Function

'*********************************************************
' Purpose: R�cup�ration d'une image en particulier, en fonction de son num�ro d'image
' Inputs: p_numImg : le num�ro de l'image (c'est � dire son num�ro de page)
'*********************************************************
Function DlImg(p_numImg)

  ' Objet permettant la connexion et la r�cup�ration d'informations via HTTP
  Dim objXMLHTTPImg: Set objXMLHTTPImg = CreateObject("Microsoft.XMLHTTP")
  Dim objStreamImg

  ' Cr�ation de l'URL de la page en fonction du num�ro de page pass� en param�tre
  Dim URLPageCurrentImg: URLPageCurrentImg = URL_MAIN & p_numImg

  ' appel de l'URL cr��e
  Call Appel(URLPageCurrentImg, 1)

  ' extraction de l'URL de l'image
  Dim URLRelCurrentImg: URLRelCurrentImg = ExtraitLien(regExtractURLRelImg)

  If (URLRelCurrentImg <> ERREUR_IMG_INEXISTANTE) Then
    Dim URLCurrentImg: URLCurrentImg = URL_DLROOT & URLRelCurrentImg
    ' extraction du nom de l'image
    Dim imgName: imgName = ExtractImgName(URLCurrentImg)

    Dim imgDateBrute, imgDate

    Set matchDateBrute = regExtractDateBrute.Execute(xHttp.responseText)
    imgDateBrute = matchDateBrute(0).SubMatches(0)
    imgDate = DateBruteToDate(imgDateBrute)
    Dim imgDateToWriteMonth: imgDateToWriteMonth = DatePart("m", imgDate)
    If(Len(imgDateToWriteMonth) = 1) Then
      imgDateToWriteMonth = "0" & imgDateToWriteMonth
    End If
    Dim imgDateToWriteDay: imgDateToWriteDay = DatePart("d", imgDate)
    If(Len(imgDateToWriteDay) = 1) Then
      imgDateToWriteDay = "0" & imgDateToWriteDay
    End If
    Dim imgDateToWrite: imgDateToWrite = DatePart("yyyy", imgDate) & imgDateToWriteMonth & imgDateToWriteDay

    ' Initialisation des emplacements source et cible pour l'image � t�l�charger
    Dim URLSourceCurrentImg: URLSourceCurrentImg = URLCurrentImg
    Dim URLTargetCurrentImg: URLTargetCurrentImg = IMGROOT & p_numImg & " - " & imgDateToWrite & " - " & imgName

    ' T�l�chargement de l'image
    objXMLHTTPImg.Open "GET", URLSourceCurrentImg, False
    objXMLHTTPImg.Send

    If objXMLHTTPImg.statusText = "OK" Then
      Set objStreamImg = CreateObject("ADODB.Stream")
      objStreamImg.Type = 1 '// binary
      objStreamImg.Open
      objStreamImg.Write objXMLHTTPImg.ResponseBody
      objStreamImg.SavetoFile URLTargetCurrentImg, 2 '//adSaveCreateOverwrite
      objStreamImg.Close
      GetImage = "OK"
    Else
      GetImage = objXMLHTTPImg.statusText
    End If
  Else
    WScript.Echo "Image " & p_numImg & " inexistante"
  End If
End Function
'endregion Gestion des t�l�chargements

'*********************************************************
' Purpose: Extraction des arguments de l'appel du programme
'*********************************************************
Function ExtractArguments()
  If(WScript.Arguments.Count = 0) Then
    WScript.Echo "Veuillez saisir au moins 1 argument."
    Call Help()
    WScript.Quit
  Else
    ' erreur si des arguments non nomm�s sont pass�s
    If WScript.Arguments.Unnamed.Count > 0 Then
      Call AffichageParametre("Erreur : veuillez n'utiliser que des arguments nomm�s (/arg)", 1)
        WScript.Quit ERREUR_TECHNIQUE
    End If

    ' V�rification du fait que l'on demande de l'aide
    If(WScript.Arguments.Named.Exists(ARG_HELP)) Then
      Call Help()
      WScript.Quit RETOUR_SANS_EXECUTION
    End If

    ' V�rification du fait que l'on soit en mode silencieux
    If(WScript.Arguments.Named.Exists(ARG_SILENTMODE)) Then
      silentMode = True
    End If

    ' V�rification du fait que l'on ne cherche pas � t�l�charger � la fois un intervalle
    ' et tout depuis la derni�re fois
    If (WScript.Arguments.Named.Exists("dl") And WScript.Arguments.Named.Exists("dll")) Then
      WScript.Echo "Erreur : impossible de t�l�charger � la fois depuis un intervalle et depuis la derni�re date. " & _
                   "Veuillez ne saisir qu'un seul type de t�l�chargement."
    End If

    ' extraction des arguments en tenant compte du mode silencieux et de l'aide
    For Each arg in WScript.Arguments.Named
      Select case arg
        ' TODO : gestion des arguments possibles
        ' TODO : /dl:[date]
        ' TODO : /dl:[[dateMin];[dateMax]]
        ' TODO : /dl:[[numImg]]
        ' TODO : /dl:[[numImgMin];[numImgMax]]
        ' TODO : /dll
        case "dl"
          Call ProcessDL(WScript.Arguments.Named.Item(arg))
        case "dll"
          Call ProcessDLL()
        case "s", "h"
        case Else
          Call AffichageParametre("Erreur : argument incorrect : " + arg _
          + "." & vbCrLf & "Taper /h pour de l'aide.", 1)
          WScript.Quit ERREUR_TECHNIQUE
        End Select
    Next
  End If
End Function

'*********************************************************
' Purpose: Fonction d'aide : appel�e avec le param�tre "-h" tout seul.
'*********************************************************
Function Help()
  WScript.Echo "CandH_Downloader v1.0 - Benoit Masson-Bedeau" & vbCrLf & _
  "Aide" & vbCrLf & _
  "Param�tres d'appel :" & vbCrLf & _
  "/h : aide. Le param�tre s'utilise seul." & vbCrLf & _
  "/s : mode silencieux. Est sp�cifi� en m�me temps que les autres param�tres." & vbCrLf & _
  "/dl:[date] : t�l�chargement d�une image en particulier � partir de sa date. Format : aaaa-MM-jj" & vbCrLf & _
  "/dl:[[dateMin];[dateMax]] : t�l�chargement des images dans l�intervalle [dateMin ; dateMax]. Format : aaaa-MM-jj" & vbCrLf & _
  "/dl:[numImg] : t�l�chargement d�une image en particulier � partir de son num�ro de page." & vbCrLf & _
  "/dl:[[numImgMin];[numImgMax]] : t�l�chargement des images dans l�intervalle [numImgMin ; numImgMax]" & vbCrLf & _
  "/dll : t�l�chargement de la derni�re image en date." & vbCrLf
End Function


'*********************************************************
' Purpose: Traite le passage d'argument /dl
' Inputs: value : valeur de l'argument
'*********************************************************
Function ProcessDL(value)
  Dim modeDate: modeDate = 1
  Dim modeIntervalle: modeIntervalle = 2
  Dim mode

  Dim paramDate, paramDateIntervalleMin, paramDateIntervalleMax

  ' TODO : ProcessDL sur un num�ro
  ' TODO : ProcessDL sur un intervalle de num�ro

  ' Pattern de date
  ' Exemple de chaine : 2017-01-01
  Dim PATTERN_DATE_ARG: PATTERN_DATE_ARG = "^[0-9]{4}-[0-9]{2}-[0-9]{2}$"
  ' Expression r�guli�re d'extraction d'une date
  Set regExtractDateArg = New RegExp
  regExtractDateArg.Pattern = PATTERN_DATE_ARG

  Set objMatchesDate = regExtractDateArg.Execute(value)
  Dim nbMatchesDate: nbMatchesDate = objMatchesDate.Count
  If(nbMatchesDate > 0) Then
    mode = modeDate
    If (IsDate(value)) Then
      paramDate = CDate(value)
    Else
      WScript.Echo "Erreur : la date est invalide."
      WScript.Quit ERREUR_TECHNIQUE
    End If
  End If

  ' Pattern d'intervalle de date
  ' Exemple de chaine : [2017-01-01;2017-05-01]
  Dim PATTERN_INTERVALLE: PATTERN_INTERVALLE = "^\[([0-9]{4}-[0-9]{2}-[0-9]{2});([0-9]{4}-[0-9]{2}-[0-9]{2})]$"
  ' Expression r�guli�re d'extraction d'une date
  Set regExtractIntervalle = New RegExp
  regExtractIntervalle.Pattern = PATTERN_INTERVALLE

  Set objMatchesIntervalle = regExtractIntervalle.Execute(value)
  Dim nbMatchesIntervalle: nbMatchesIntervalle = objMatchesIntervalle.Count
  If(nbMatchesIntervalle > 0) Then
    mode = modeIntervalle
    If (IsDate(objMatchesIntervalle(0).SubMatches(0))) Then
      paramDateIntervalleMin = CDate(objMatchesIntervalle(0).SubMatches(0))
    Else
      WScript.Echo "Erreur : la date de la borne inf�rieure est invalide."
      WScript.Quit ERREUR_TECHNIQUE
    End If
    If (IsDate(objMatchesIntervalle(0).SubMatches(1))) Then
      paramDateIntervalleMax = CDate(objMatchesIntervalle(0).SubMatches(1))
    Else
      WScript.Echo "Erreur : la date de la borne sup�rieure est invalide."
      WScript.Quit ERREUR_TECHNIQUE
    End If
  End If

  ' Contr�le des dates
  If (mode = modeDate) Then
    If(DateDiff("d", paramDate, dMin) > 0) Then
      WScript.Echo "La date est inf�rieure � la date du premier comic : elle a �t� remplac�e par celle-ci."
      paramDate = dMin
    End If
    If(DateDiff("d", paramDate, dMax) < 0) Then
      WScript.Echo "Erreur : la borne sup�rieure est sup�rieure � la date du dernier comic : elle a �t� remplac�e par celle-ci."
      paramDateIntervalleMax = dMax
    End If
  ElseIf (mode = modeIntervalle) Then
    ' Pour l'intervalle, la borne inf�rieure doit en plus �tre inf�rieure ou �gale � la borne sup�rieure
    If (DateDiff("d", paramDateIntervalleMin, paramDateIntervalleMax) < 0) Then
      WScript.Echo "Erreur : la borne inf�rieure doit �tre sup�rieure ou �gale � la borne sup�rieure."
      WScript.Quit ERREUR_TECHNIQUE
    End If

    If(DateDiff("d", paramDateIntervalleMin, dMin) > 0) Then
      WScript.Echo "La borne inf�rieure est inf�rieure � la date du premier comic : elle a �t� remplac�e par celle-ci."
      paramDateIntervalleMin = dMin
    End If
    If(DateDiff("d", paramDateIntervalleMax, dMax) < 0) Then
      WScript.Echo "La borne sup�rieure est sup�rieure � la date du dernier comic : elle a �t� remplac�e par celle-ci."
      paramDateIntervalleMax = dMax
    End If

    Call GetImgsIntervalleDate(paramDateIntervalleMin, paramDateIntervalleMax)
  End If
End Function

'*********************************************************
' Purpose: Traite le passage d'argument /dll
'*********************************************************
Function ProcessDLL()
  ' TODO : ProcessDLL : r�cup�ration du num�ro de la derni�re image t�l�charg� puis t�l�chargement
  ' � partir de n + 1
  WScript.Echo "ProcessDLL value : " & value
End Function

'*********************************************************
' Purpose: Converti une date brute (2005.12.31) en une vraie date
' Inputs: value : valeur de la date brute
'*********************************************************
Function DateBruteToDate(value)
  Dim PATTERN_EXTRACT: PATTERN_EXTRACT = "^([0-9]{4})\.([0-9]{2})\.([0-9]{2})$"
  ' Expression r�guli�re d'extraction d'une date
  Set regExtract = New RegExp
  regExtract.Pattern = PATTERN_EXTRACT

  Set objMatchesExtract = regExtract.Execute(value)
  Dim nbMatches: nbMatches = objMatchesExtract.Count
  If(nbMatches > 0) Then
    Dim annee: annee = objMatchesExtract(0).SubMatches(0)
    Dim mois: mois = objMatchesExtract(0).SubMatches(1)
    Dim jour: jour = objMatchesExtract(0).SubMatches(2)
    DateBruteToDate = CDate(jour & "/" & mois & "/" & annee)
  End If
End Function
'endregion M�tier

'region Util
'*********************************************************
' Purpose: affiche le texte � afficher accompagn� d'un timestamp si pr�cis�, si
'          le mode silencieux n'est pas activ�
' Inputs: texte : texte � afficher
'         avecTS : si True, affichage du timestamp ; pas d'affichage sinon
'*********************************************************
Function AffichageParametre(texte, avecTS)
  Select case avecTS
    case 0
      If(NOT silentMode) Then WScript.Echo texte End If
    case 1
      If(NOT silentMode) Then WScript.Echo Timestamp + " - " + texte End If
    case Else
      If(NOT silentMode) Then WScript.Echo Timestamp + " - " + "Erreur : choix incorrect : " + avecTS End If
        WScript.Quit ERREUR_TECHNIQUE
    End Select
End Function

'*********************************************************
' Purpose: Cr�ation d'un timestamp sous la forme yyyy-MM-dd HH:mm:ss
'*********************************************************
Function Timestamp()
  dim dateNow, currentYear, currentMonth, currentDay, currentHour
  dim currentMinute, currentSecond, currentNano, dateFormated

  dateNow         = now
  currentYear     = Year(dateNow)
  currentMonth    = Right("0" & Month(dateNow), 2)
  currentDay      = Right("0" & Day(dateNow), 2)
  currentHour     = Right("0" & Hour(dateNow), 2)
  currentMinute   = Right("0" & Minute(dateNow), 2)
  currentSecond   = Right("0" & Second(dateNow), 2)
  dateFormated    = currentYear & "-" & currentMonth & "-" & currentDay & " " & currentHour & ":" & _
  currentMinute & ":" & currentSecond

  Timestamp = dateFormated
End Function
'endregion Util
