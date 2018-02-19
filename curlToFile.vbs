'*********************************************************************************
'|                          DÉCLARATION DES CONSTANTES                           |
'*********************************************************************************
' les différents chemins et URL
Dim URLPRESTA: URLPRESTA = "http://presta-si.cm-cic.fr/FOU1_PRESTA/devbooster.aspx?_pid=Menu&_fid=DoUtilisateVisu&_state=2050&_saguid=oS0g2QCCEkG4vNaA33GTxQ%3d%3dMTAuNDYuNC4xMzI6ODAwNg%3d%3d&_PTL=D&_rendertype=WebXForm&_rendertypeversion=2"
Dim URLROOT: URLROOT = "http://presta-si.cm-cic.fr/FOU1_PRESTA/"
Dim LOGROOT: LOGROOT = "U:\tmp\"
' textes concernant le statut et les actions possibles
Dim BADGER_TXT: BADGER_TXT = "badger"
Dim DEBADGER_TXT: DEBADGER_TXT = "débadger"
Dim TOGGLE_TXT: TOGGLE_TXT = "toggle"
Dim STATUT_BADGE_TXT: STATUT_BADGE_TXT = "badgé"
Dim STATUT_DEBADGE_TXT: STATUT_DEBADGE_TXT = "débadgé"
' gestion des expressions régulières
Dim PATTERN_SESSION_EN_COURS: PATTERN_SESSION_EN_COURS = "class=""act"" href=""(.*?)"""
Dim PATTERN_NOUVELLE_JOURNEE: PATTERN_NOUVELLE_JOURNEE = "form id=""DoValidate"" action=""(.*?)""[\s\S]*Identifier l'utilisateur"
' gestion des arguments
Dim ARG_SILENTMODE: ARG_SILENTMODE = "-s"
Dim ARG_HELP: ARG_HELP = "-h"
Dim ARG_BADGER: ARG_BADGER = "-b"
Dim ARG_DEBADGER: ARG_DEBADGER = "-d"
' gestion des erreurs
Dim ERREUR_TECHNIQUE: ERREUR_TECHNIQUE = 1
' gestion des retours OK
Dim RETOUR_SANS_EXECUTION: RETOUR_SANS_EXECUTION = 10
Dim RETOUR_OK: RETOUR_OK = 0

'*********************************************************************************
'|                          DÉCLARATION DES VARIABLES                            |
'*********************************************************************************
' variables d'exécution
Dim silentMode
' liens brut extrait depuis l'appel à PRESTA, pour badger et débadger
Dim linkExtrait, linkBadger, linkDebadger
' booléen indiquant le statut actuel, mis à jour après chaque vérification
' ainsi que son action possible associée
Dim statut, actionPossible, actionDesiree
' objet permettant la gestion d'appel HTTP
Dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
' stream qui va servir pour l'écriture de fichier
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
' les expressions régulières pour déterminer si c'est une session en cours ou une
' nouvelle journée
Set regNouvelleJournee = New RegExp
Set regSessionEnCours = New RegExp
regNouvelleJournee.Pattern = PATTERN_NOUVELLE_JOURNEE
regSessionEnCours.Pattern = PATTERN_SESSION_EN_COURS

'*********************************************************************************
'|                                    MÉTIER                                     |
'*********************************************************************************
' lecture et traitement des arguments passés au script
Call ExtractArguments()

' appel initial
Call AppelInitial(URLPRESTA, 1)

' extraction du bon lien
Call ExtractLienCourant

'détermination de l'URL à appeler dans le cas où l'on veut badger ou débadger
Call ExtractURLsBadgerDebadger

'Call PlanifieAction(DEBADGER_TXT, "24/04/2015", "14:35")
'WScript.Quit 99
' TODO : gestion en fonction des paramètres extraits. Fonction ?

If(Len(actionDesiree) > 0) Then
  Call ExecuteActionDemandee(actionDesiree, 0)
Else
  WScript.Quit RETOUR_SANS_EXECUTION
End If

'Call Toggle(0)

'*********************************************************************************
'|                          DÉCLARATION DES FONCTIONS                            |
'*********************************************************************************
Function Timestamp()
' création d'un timestamp sous la forme yyyy-MM-dd HH:mm:ss
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

Function ExtractArguments()
' Fonction d'extraction des arguments
  Dim veutBadger
  Dim veutDebadger
  Dim veutToggle
  Dim conflitAction

  If(WScript.Arguments.Count = 0) Then
    WScript.Echo Timestamp + " - " + "Erreur : appel du script sans argument."
    WScript.Quit ERREUR_TECHNIQUE
  Else
    ' on fait une première boucle pour voir si l'on est en mode silencieux
    For Each arg in WScript.Arguments
      If(arg = ARG_SILENTMODE) Then
        silentMode = True
      End If
    Next
    ' boucle pour savoir si l'on demande l'aide
    For Each arg in WScript.Arguments
      If(arg = ARG_HELP) Then
        Call Help()
        WScript.Quit RETOUR_SANS_EXECUTION
      End If
    Next
    ' extraction des arguments en tenant compte du mode silencieux
    For Each arg in WScript.Arguments
      Select case arg
        case "-a"
          Call AffichageParametre("Vous voulez connaitre votre statut actuel.", 1)
          Call DetermineStatut()
        case "-b"
          veutBadger = True
          Call AffichageParametre("Vous voulez badger.", 1)
        case "-d"
          veutDebadger = True
          Call AffichageParametre("Vous voulez débadger.", 1)
        case "-t"
          veutToggle = True
          Call AffichageParametre("Vous voulez faire un toggle de votre statut.", 1)
        case "-s", "-h"
        case Else
          Call AffichageParametre("Erreur : argument incorrect : " + arg _
                                  + "." & vbCrLf & "Taper -h pour de l'aide.", 1)
          WScript.Quit ERREUR_TECHNIQUE
      End Select
    Next
  End If
  ' gestion du résultat
  If (((veutBadger And veutDebadger) = True) Or ((veutBadger And veutToggle) = True) Or ((veutDebadger And veutToggle) = True)) Then
    conflitAction = True
    ElseIf (veutBadger) Then
      actionDesiree = BADGER_TXT
    ElseIf (veutDebadger) Then
      actionDesiree = DEBADGER_TXT
    ElseIf (veutToggle) Then
      actionDesiree = TOGGLE_TXT
  End If

  If (conflitAction) Then
    Call AffichageParametre("Erreur : une seule action sur le statut est possible à la fois.", 1)
    WScript.Quit ERREUR_TECHNIQUE
  End If
End Function

Function Help()
' Fonction d'aide : appelée avec le paramètre "-h" tout seul.
WScript.Echo "CurlToFile v1.0 - Benoit Masson-Bedeau" & vbCrLf & _
"Aide" & vbCrLf & _
"Paramètres d'appel :" & vbCrLf & _
"-h : aide. Le paramètre s'utilise seul." & vbCrLf & _
"-a : connaitre le statut actuel." & vbCrLf & _
"-b : badger. Une information est retournée si le statut est déjà à badgé." & vbCrLf & _
"-d : débadger. Une information est retournée si le statut est déjà à débadgé." & vbCrLf & _
"-t : toggle. Inverse le statut." & vbCrLf & _
"-s : mode silencieux. Est spécifié en même temps que les autres paramètres." & vbCrLf
End Function

Function AppelInitial(urlAppelInitial, avecLog)
' Fonction d'appel initial de PRESTA ; c'est à partir de cette réponse que l'on va
' savoir si l'on est dans un cas ou l'on a déjà badgé ou s'il faut badger
xHttp.Open "GET", urlAppelInitial, False
xHttp.Send
If(avecLog = 1) Then
' écriture dans fichier
    With bStrm
        .type = 1 '//binary
        .open
        .write xHttp.responseBody
        .savetofile LOGROOT & "presta.html", 2 '//overwrite
        .close
    End With
End If
End Function

Function ExtraitLien(regEx)
' Extraction du lien pour badger ou débadger
    Set objMatches = regEx.Execute(xHttp.responseText)
    Dim nbMatches: nbMatches = objMatches.Count
    Dim result: result = objMatches(0)
    ExtraitLien = objMatches(0).SubMatches(0)
End Function

Function ExtractLienCourant()
' Fonction d'extraction du lien en fonction d'une nouvelle journée
' ou d'une session en cours

  ' Nouvelle journée
  If regNouvelleJournee.Test(xHttp.responseText) Then
      linkExtrait = ExtraitLien(regNouvelleJournee)
  ' Session en cours
  ElseIf regSessionEnCours.Test(xHttp.responseText) Then
      linkExtrait = ExtraitLien(regSessionEnCours)
  Else
      linkExtrait = ""
  End If
End Function

Function DetermineStatut()
  If(Len(statut) > 0) Then
    DetermineStatut = statut
  Else
    ' TODO
    Call AppelInitial(URLPRESTA, 1)
    Call ExtractLienCourant
    DetermineStatut = DetermineBadgerDebadger(linkExtrait)
    Call AffichageParametre("Votre statut est actuellement : " + statut, 1)
  End If
End Function

Function DetermineBadgerDebadger(lien)
  ' Fonction qui va déterminer si l'on est dans le cas où l'on doit badger ou débadger
  ' valeurs dans l'URL récupérée
  Dim valeurBadger: valeurBadger = "1"
  Dim valeurDebadger: valeurDebadger = "2"

  Set regLienMatin = New RegExp
  ' on envoie ce lien uniquement pour le matin ; néanmoins, attention : ce lien
  ' est aussi présent dans la page en journée
  regLienMatin.Pattern = "FID=DoValidate"
  regLienMatin.IgnoreCase = False

  Set regLienJournee = New RegExp
  regLienJournee.Pattern = "actionsg_current=(.)"
  regLienJournee.IgnoreCase = False

  ' gestion nouvelle journée : on est forcément dans le cas où il faut badger
  If(regLienMatin.Test(lien)) Then
    DetermineBadgerDebadger = BADGER_TXT
    statut = STATUT_DEBADGE_TXT
  Else
    ' sinon on traite comme une journée classique
    Dim courant: courant = regLienJournee.Execute(lien)(0).SubMatches(0)
    ' gestion du résultat
    Select case courant
      case valeurBadger
        DetermineBadgerDebadger = BADGER_TXT
        statut = STATUT_DEBADGE_TXT
      case valeurDebadger
        DetermineBadgerDebadger = DEBADGER_TXT
        statut = STATUT_BADGE_TXT
      case Else
        DetermineBadgerDebadger = "erreur : code d'entrée non reconnu : " & lien
    End Select
  End If
End Function

Function ExtractURLsBadgerDebadger()
' extraction des liens pour badger ou débadger
  actionPossible = DetermineBadgerDebadger(linkExtrait)
  If(actionPossible = BADGER_TXT) Then
    linkBadger = URLROOT & linkExtrait
  ElseIf(actionPossible = DEBADGER_TXT) Then
    linkDebadger = URLROOT & linkExtrait
  End If
End Function

Function Toggle(avecLog, avecConfirmation)
' Fonction qui fait un toggle sur l'état badgé / débadgé
' TODO : rendre ce bloc if paramétrable ? Et du coup forcer
' le if suivant à true dans le cas ou ce if n'est pas appelé
  Dim demandeSwitch
  If(avecConfirmation = 1) Then
    If((actionPossible = BADGER_TXT) Or (actionPossible = DEBADGER_TXT)) Then
      demandeSwitch = MsgBox("Le statut est actuellement à : " & statut & ". Voulez vous " & actionPossible & " ?", vbYesNo)
    Else
      Call AffichageParametre("Erreur de traitement.", 1)
      WScript.Quit ERREUR_TECHNIQUE
    End If
  End If
  If(demandeSwitch = vbYes Or avecConfirmation = 0) Then
    Dim linkToSubmit
    If(Len(linkDebadger) > 0) Then
      linkToSubmit = linkDebadger
      ElseIf (Len(linkBadger) > 0) Then
        linkToSubmit = linkBadger
    End If
    xHttp.Open "GET", linkToSubmit, False
    xHttp.Send
    ' mise à jour du statut et de l'action possible
    Dim lienRetour: lienRetour = ExtraitLien(regSessionEnCours)
    Call DetermineBadgerDebadger(lienRetour)
    ' log si demandé
    If(avecLog = 1) Then
      With bStrm
        ' TODO : gérer si open ou non
        .write xHttp.responseBody
        .savetofile LOGROOT & "prestaResponseToggle.html", 2 '//overwrite
      End With
    End If
    Call AffichageParametre("Votre statut est désormais : " + statut, 1)
    WScript.Quit RETOUR_OK
  Else
    Call AffichageParametre("Aucune opération effectuée. Votre statut est toujours : " + statut, 1)
    WScript.Quit RETOUR_SANS_EXECUTION
  End If
End Function

'Function Badger(avecLog)
Function ExecuteActionDemandee(actionDemandee, avecLog)
' Fonction qui permet d'essayer d'exécuter la fonction demandée.

  ' Gestion de la demande du Toggle
  If(actionDemandee = TOGGLE_TXT) Then
    Call Toggle(avecLog, 0)
  ' Gestion badger / débadger
  ElseIf(actionPossible = actionDemandee) Then
    Dim linkToSubmit
    If(Len(linkDebadger) > 0) Then
      linkToSubmit = linkDebadger
    ElseIf (Len(linkBadger) > 0) Then
      linkToSubmit = linkBadger
    End If

    ' Soumission de l'URL pour badger / débadger
    xHttp.Open "GET", linkToSubmit, False
    xHttp.Send
    ' mise à jour du statut et de l'action possible
    Dim lienRetour: lienRetour = ExtraitLien(regSessionEnCours)
    Call DetermineBadgerDebadger(lienRetour)
    ' log si demandé
    If(avecLog = 1) Then
      With bStrm
        ' TODO : gérer si open ou non
        .write xHttp.responseBody
        .savetofile LOGROOT & "prestaResponse" & actionDemandee & ".html", 2 '//overwrite
      End With
    End If
    Call AffichageParametre("Votre statut est désormais : " + statut, 1)
    WScript.Quit RETOUR_OK
  Else
    Call AffichageParametre("Erreur : le statut est déjà " + statut + ".", 1)
    WScript.Quit ERREUR_TECHNIQUE
  End If
End Function

Function PlanifieAction(actionAPlanifier, datePlanif, heureMinutesPlanif)
' Fonction de création d'une tâche planifiée avec une action associée
' TODO
  Dim wshshell: Set wshshell = createobject("wscript.shell")
  Dim fullPath: fullPath = WScript.ScriptFullName
  ' on prend un nom de tâche un peu générique
  Dim nom: nom = "checkBMA"

  Dim commande: commande = "schtasks /create /tn """ & nom &_
  """ /tr " & fullPath & " /sc once /sd " & datePlanif & " /st " & heureMinutesPlanif
  
  WScript.Echo commande

  'wshshell.run "cscript //nologo ospp.vbs /sethst:sn.fr",,true
End Function


' TODO : -> OK ! checker gestion badger le matin, lorsque l'on doit "s'identifier"
' TODO : -> OK ! checker gestion débadger le soir, lorsque la session est en cours
' TODO : log "court" de l'opération de badge ; pop-up paramétrable ?
'        -> date, heure
' TODO : demande si l'utilisateur veut débadger aujourd'hui / badger demain
'        -> création tâche planifiée
' TODO : -> OK ! finir l'appel avec arguments (cscript curlToFile "Badger" "silent" par exemple)
' TODO : -> OK ! remplacer tous les MsgBox par des WScript.Echo
' TODO : -> OK ! créer une fonction d'affichage de message d'info ou d'erreur qui gère le silentMode et le timestamp
' TODO : -> OK ! ajouter un paramètre pour connaitre le statut actuel
' TODO : ajout paramètre pour lancer la création d'une tâche planifiée
