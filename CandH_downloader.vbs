' Script de téléchargement des comics Cyanide and Hapiness
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
' Emplacement des images téléchargées
Dim IMGROOT: IMGROOT = "D:\temp\CandH\"
'endregion stockage

'region URLs
' URL principale
Dim URL_MAIN: URL_MAIN = "http://explosm.net/comics/"

' URL du dernier comic
Dim URL_LATEST: URL_LATEST = "http://explosm.net/comics/latest"

' URL de base de téléchargement des images
Dim URL_DLROOT: URL_DLROOT = "http://files.explosm.net/comics/"
'endregion URL

'region Identification des éléments
' ID de l'élément contenant l'image dans le cas d'une page normale
Dim IMG_ID: IMG_STD_ID = "main-comic"
'endregion Identification des éléments

'region Gestion des expression régulières
' Pattern d'extraction du chemin de l'image actuelle, depuis la page complète
' Exemple de chaine : <input id="permalink" type="text" value="http://explosm.net/comics/4376/" onclick="this.select()">
Dim PATTERN_URL_CUR: PATTERN_URL_CUR = "id=""permalink"" .* value=""(.*)"" .*"
' Expression régulière d'extraction du lien de l'image actuelle
Set regExtractURLCur = New RegExp
regExtractURLCur.Pattern = PATTERN_URL_CUR

' Pattern d'extraction du chemin du lien de l'image précédente, depuis la page complète
' Exemple de chaine : <li><a href="/comics/4375/" class="previous-comic " title="Previous comic"><img src="/img/nav-button_previous@2x.png"/></a></li>
Dim PATTERN_URL_PREV: PATTERN_URL_PREV = "href=""/(.*)"" .* title=""Previous comic"""
' Expression régulière d'extraction du lien de l'image précédente
Set regExtractURLPrev = New RegExp
regExtractURLPrev.Pattern = PATTERN_URL_PREV

' Pattern d'extraction du chemin du lien de l'image suivante, depuis la page complète
' Exemple de chaine : <li><a href="/comics/4377/" class="next-comic " title="Next comic"><img src="/img/nav-button_next@2x.png"/></a></li>
Dim PATTERN_URL_NEXT: PATTERN_URL_NEXT = "href=""/(.*)"" .* title=""Next comic"""
' Expression régulière d'extraction du lien de l'image suivante
Set regExtractURLNext = New RegExp
regExtractURLNext.Pattern = PATTERN_URL_NEXT

' Pattern d'extraction du numéro de l'image, depuis un lien de page d'image, extrait avec regExtractURLCur/Prev/Next
' Exemple de chaine : cf. href dans les extractions de chemin
Dim PATTERN_NUM_IMG: PATTERN_NUM_IMG = "comics/(.*)/"
' Expression régulière d'extraction du numéro de l'image
Set regExtractNumImg = New RegExp
regExtractNumImg.Pattern = PATTERN_NUM_IMG

' Expression régulière d'extraction de l'URL relative d'une image sur une page
Dim PATTERN_URL_REL_IMG: PATTERN_URL_REL_IMG = "id=""main-comic"" src=""\/\/files\.explosm\.net\/comics\/(.*)"""
Set regExtractURLRelImg = new RegExp
regExtractURLRelImg.Pattern = PATTERN_URL_REL_IMG

' Expression régulière d'extraction du nom de l'image à partir de son URL complète
Dim PATTERN_NAME_IMG: PATTERN_NAME_IMG = ".*\/(.*)"
Set regExtractNameImg = new RegExp
regExtractNameImg.Pattern = PATTERN_NAME_IMG
'endregion Gestion des expressions régulières

'region Objets divers
' objet permettant la gestion d'appel HTTP
Dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
' stream qui va servir pour l'écriture de fichier
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
'endregion Objets divers
'endregion Constantes

'region Variables
'***********************************************
'|                 Variables                   |
'***********************************************

'region variables d'exécution
Dim silentMode
'endregion variables d'exécution

'region Intervalle de recherche
' Intervalle de recherche inférieur
Dim rMin
' Intervalle de recherche supérieur
Dim rMax
'endregion Intervalle de recherche

'region URLs
' URL des images
Dim URLImg, URLImgPrev, URLImgNext
' URL des pages vers les images
Dim URLPageImg, URLPageImgPrev, URLPageImgNext
' URL relative des pages vers les images
Dim URLRelImg, URLRelImgPrev, URLRelImgNext
'endregion URLs

'region Identification des éléments
' Numéro des images
Dim numImg, numImgPrev, numImgNext
'endregion Identification des éléments
'endregion Variables

'region Métier
'***********************************************
'|                 Métier                      |
'***********************************************
'region Main

Call CalcIntervalleMaxPerma

Call ExtractArguments

'Call GetImgsIntervalle(0, 40)

'endregion Main

'region Appel
'*********************************************************
' Purpose: Récupère la page indiqué par le paramètre p_URLPageImg
' Inputs: p_URLPageImg : l'URL de la page pour laquelle récupérer le contenu
'         avecLog : si 1, log de l'appel dans le fichier déclaré ; sinon, l'appel est silencieux
' Returns: Le contenu de la page sous forme d'objet, dans l'objet déclaré dans la fonction
'*********************************************************
Function Appel(p_URLPageImg, avecLog)
  xHttp.Open "GET", p_URLPageImg, False
  xHttp.Send
  If(avecLog = 1) Then
    ' écriture dans fichier
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

'region Extraction des liens de page, numéros et noms d'images
'*********************************************************
' Purpose: Extraction du lien vers la page de l'image, en fonction de la regEx passée en paramètre
' Inputs: regExLien : expression régulière d'un lien en particulier
' Returns: Si un lien est extrait, le lien lui-même ; sinon (pas de lien trouvé), erreur
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
' Purpose: Extraction du numéro de l'image en fonction de son URL relative
' Inputs: regExImg : expression régulière d'extraction d'un numéro d'image dans une URL relative
'         p_URLRelativeImg : URL relative (sans la partie URLROOT) de l'image
' Returns: le numéro de l'image extrait
'*********************************************************
Function ExtraitNumImg(regExImg, p_URLRelativeImg)
  ' TODO : gestion d'erreur si le numéro ne parvient pas à être extrait
  Set objMatches = regExImg.Execute(p_URLRelativeImg)
  Dim nbMatches: nbMatches = objMatches.Count
  Dim result: result = objMatches(0)
  ExtraitNumImg = objMatches(0).SubMatches(0)
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
'endregion Extraction des liens, numéros et noms d'images

'region Gestion des intervalles minimal et maximal
'*********************************************************
' Purpose: Calcul de l'intervalle maximal en fonction du lien permanent (permalink)
' Returns: l'intervalle maximal ; dans la même temps, la variable globale est aussi
'          valorisé avec celui-ci
'*********************************************************
Function CalcIntervalleMaxPerma()

  ' Appel initial vers le dernier comic en date
  Call Appel(URL_LATEST, 1)

  ' Extraction du lien à partir du permalien de la page
  URLPageImg = ExtraitLien(regExtractURLCur)

  ' Extraction du numéro de l'image sur la page qui déterminera l'intervalle maximal
  numImg = ExtraitNumImg(regExtractNumImg, URLPageImg)

  ' Valorisation de l'intervalle maximal et retour de fonction
  rMax = numImg

  ' TODO : valorisation de la date max

  CalcIntervalleMaxPerma = rMax
End Function

'*********************************************************
' Purpose: Récupération des images dans l'intervalle renseigné, pour chaque
'          numéro d'image disponible sur le site
' Inputs: numImgMin : l'intervalle bas (numéro d'image minimum)
'         numImgMax : l'intervalle haut (numéro d'image maximum)
'*********************************************************
Function GetImgsIntervalle(numImgMin, numImgMax)

  ' Si l'intervalle n'est pas correct, erreur
  If (numImgMin > numImgMax) Then
    WScript.Echo "Erreur : l'intervalle minimal doit être inférieur ou égal à l'intervalle maximal"
    WScript.Quit ERREUR_TECHNIQUE
    ' Sinon, on boucle dans l'intervalle
  Else
    For counter = numImgMin To numImgMax
      DlImg(counter)
    Next
  End If
End Function
'endregion Gestion des intervalles minimal et maximal

'region Gestion des téléchargements
'*********************************************************
' Purpose: Récupération d'une image en particulier, en fonction de son numéro d'image
' Inputs: p_numImg : le numéro de l'image (c'est à dire son numéro de page)
'*********************************************************
Function DlImg(p_numImg)

  ' Objet permettant la connexion et la récupération d'informations via HTTP
  Dim objXMLHTTPImg: Set objXMLHTTPImg = CreateObject("Microsoft.XMLHTTP")
  Dim objStreamImg

  ' Création de l'URL de la page en fonction du numéro de page passé en paramètre
  Dim URLPageCurrentImg: URLPageCurrentImg = URL_MAIN & p_numImg

  ' appel de l'URL créée
  Call Appel(URLPageCurrentImg, 1)

  ' extraction de l'URL de l'image
  Dim URLRelCurrentImg: URLRelCurrentImg = ExtraitLien(regExtractURLRelImg)

  If (URLRelCurrentImg <> ERREUR_IMG_INEXISTANTE) Then
    Dim URLCurrentImg: URLCurrentImg = URL_DLROOT & URLRelCurrentImg
    ' extraction du nom de l'image
    Dim imgName: imgName = ExtractImgName(URLCurrentImg)

    ' Initialisation des emplacements source et cible pour l'image à télécharger
    Dim URLSourceCurrentImg: URLSourceCurrentImg = URLCurrentImg
    Dim URLTargetCurrentImg: URLTargetCurrentImg = IMGROOT & p_numImg & " - " & imgName

    ' Téléchargement de l'image
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
'endregion Gestion des téléchargements

'*********************************************************
' Purpose: Extraction des arguments de l'appel du programme
'*********************************************************
Function ExtractArguments()
  If(WScript.Arguments.Count = 0) Then
    WScript.Echo "Veuillez saisir au moins 1 argument."
    Call Help()
    WScript.Quit
  Else
    ' erreur si des arguments non nommés sont passés
    If WScript.Arguments.Unnamed.Count > 0 Then
      Call AffichageParametre("Erreur : veuillez n'utiliser que des arguments nommés (/arg)", 1)
        WScript.Quit ERREUR_TECHNIQUE
    End If

    ' Vérification du fait que l'on demande de l'aide
    If(WScript.Arguments.Named.Exists(ARG_HELP)) Then
      Call Help()
      WScript.Quit RETOUR_SANS_EXECUTION
    End If

    ' Vérification du fait que l'on soit en mode silencieux
    If(WScript.Arguments.Named.Exists(ARG_SILENTMODE)) Then
      silentMode = True
    End If

    ' Vérification du fait que l'on ne cherche pas à télécharger à la fois un intervalle
    ' et tout depuis la dernière fois
    If (WScript.Arguments.Named.Exists("dl") And WScript.Arguments.Named.Exists("dll")) Then
      WScript.Echo "Erreur : impossible de télécharger à la fois depuis un intervalle et depuis la dernière date. " & _
                   "Veuillez ne saisir qu'un seul type de téléchargement."
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
          Call ProcessDLL(WScript.Arguments.Named.Item(arg))
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
' Purpose: Fonction d'aide : appelée avec le paramètre "-h" tout seul.
'*********************************************************
Function Help()
  WScript.Echo "CandH_Downloader v1.0 - Benoit Masson-Bedeau" & vbCrLf & _
  "Aide" & vbCrLf & _
  "Paramètres d'appel :" & vbCrLf & _
  "/h : aide. Le paramètre s'utilise seul." & vbCrLf & _
  "/s : mode silencieux. Est spécifié en même temps que les autres paramètres." & vbCrLf & _
  "/dl:[date] : téléchargement d’une image en particulier à partir de sa date. Format : aaaa-MM-jj" & vbCrLf & _
  "/dl:[[dateMin];[dateMax]] : téléchargement des images dans l’intervalle [dateMin ; dateMax]. Format : aaaa-MM-jj" & vbCrLf & _
  "/dl:[numImg] : téléchargement d’une image en particulier à partir de son numéro de page." & vbCrLf & _
  "/dl:[[numImgMin];[numImgMax]] : téléchargement des images dans l’intervalle [numImgMin ; numImgMax]" & vbCrLf & _
  "/dll : téléchargement de la dernière image en date." & vbCrLf
End Function


'*********************************************************
' Purpose: Traite le passage d'argument /dl
' Inputs: value : valeur de l'argument
'*********************************************************
Function ProcessDL(value)
  WScript.Echo "ProcessDL value : " & value
End Function

'*********************************************************
' Purpose: Traite le passage d'argument /dll
' Inputs: value : valeur de l'argument
'*********************************************************
Function ProcessDLL(value)
  WScript.Echo "ProcessDLL value : " & value
End Function
'endregion Métier

'region Util
'*********************************************************
' Purpose: affiche le texte à afficher accompagné d'un timestamp si précisé, si
'          le mode silencieux n'est pas activé
' Inputs: texte : texte à afficher
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
' Purpose: Création d'un timestamp sous la forme yyyy-MM-dd HH:mm:ss
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
