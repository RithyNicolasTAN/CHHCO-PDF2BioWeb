#include <File.au3>
#include <Array.au3>
#include <Date.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPIFiles.au3>
#include <Misc.au3>

If _Singleton("pdf2bioweb", 1) = 0 Then Exit
If _Singleton("attribution", 1) = 0 Then Exit
If _Singleton("apicrypt", 1) = 0 Then Exit
If _Singleton("integration", 1) = 0 Then Exit
If _Singleton("maj_scanbac", 1) = 0 Then Exit

$aIgnoreList=FileReadToArray(@ScriptDir&"/Ignore_list.txt")
;~ _ArrayDisplay($aIgnoreList)

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Interface Bactériologie/CERBA BioWin --> BioWeb et HM", 624, 135, 192, 124)
$Label1 = GUICtrlCreateLabel("", 8, 8, 601, 45)
$Progress1 = GUICtrlCreateProgress(8, 64, 601, 25)
$Button1 = GUICtrlCreateButton("Quitter", 216, 96, 185, 33)
GUICtrlSetState(-1, $GUI_DISABLE)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Local $aMyDate2, $aMyTime2
_DateTimeSplit( _DateAdd("d", -31, _NowCalcDate()), $aMyDate2, $aMyTime2)
$a2 = stringRight($aMyDate2[1], 2)
if $aMyDate2[2] < 10 then
	$m2 = "0" & $aMyDate2[2]
Else
	$m2 = $aMyDate2[2]
 EndIf


#Region ### CHARGEMENTS DES FICHIERS ET DES TABLES ###

;COPIE ET CONVERSION DE LA BASE DE DONNEE SCANNER DE BIOWIN
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 1/5" & @CRLF & "Copie des fichiers depuis BioWin...")
GUICtrlSetData($Progress1, 0)
FileCopy("Y:\Biowin\Scanner.fic", @ScriptDir & "\Scanner.fic", 1)
GUICtrlSetData($Progress1, 2.5)
FileCopy("Y:\Biowin\Scanner.ndx", @ScriptDir & "\Scanner.ndx", 1)
GUICtrlSetData($Progress1, 5)
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 2/5" & @CRLF & "Conversion des fichiers...")
Runwait(@ScriptDir & "\hyperfile2xml.exe Scanner.fic Scanner.xml")
GUICtrlSetData($Progress1, 10)
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 3/5" & @CRLF & "Lecture du fichier SCANNER")
Local $hFileOpen = FileOpen(@ScriptDir & "\Scanner.xml", $FO_READ)
$lign = _FileCountLines(@ScriptDir & "\scanner.xml")
FileReadLine($hFileOpen)
FileReadLine($hFileOpen)
Local $aSCAN[(($lign - 3) / 11)][8]
for $i = 0 to UBound($aSCAN, 1) - 1
	GUICtrlSetData($Progress1, 10 + ($i / (UBound($aSCAN, 1) - 1)) * 2.5)
	FileReadLine($hFileOpen)
	$aSCAN[$i][0] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8) & StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8)
	for $j = 2 to 8
		$aSCAN[$i][$j - 1] = StringStripWS(StringSplit(StringSplit(FileReadLine($hFileOpen), ">")[2], "<")[1], 8)
	Next
	FileReadLine($hFileOpen)
Next
FileClose($hFileOpen)
;~ _ArrayDisplay($aSCAN)

local $aList1[UBound($aSCAN, 1)]
for $i = 1 to UBound($aSCAN, 1) - 1
	GUICtrlSetData($Progress1, 12.5 + ($i / (UBound($aSCAN, 1) - 1)) * 2.5)
	if Number(stringleft($aSCAN[$i][0], 5)) >= Number("2" & $a2 & $m2) Then
	   if _ArraySearch($aIgnoreList, $aSCAN[$i][0])=-1 then $aList1[$i] = $aSCAN[$i][0]



	EndIf
 Next
;~ _ArrayDisplay($alist1)
$aList1 = _ArrayUnique($aList1)
;~ _ArrayDisplay($alist1)

;VERIFICATION POUR CHAQUE LIGNE QUE LA VERSION DU FICHIER SUR BIOWEB EST > A 1.5 ET CREATION DE LA LISTE DE TRAVAIL $aLIST2
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 4/5" & @CRLF & "Vérification des fichiers sur BioWeb...")
GUICtrlSetData($Progress1, 15)

local $aList2[1]
for $i = 2 to UBound($aList1) - 1
	GUICtrlSetData($Progress1, 15 + ($i / (UBound($aList1) - 1)) * 2.5)
	GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 4/5" & @CRLF & "Vérification des fichiers sur BioWeb (" & $aList1[$i] & ")")
	$hFileOpen = FileOpen("Y:\CONSER\pdf\V_1" & $aList1[$i] & ".pdf", $FO_READ + $FO_UTF8)
	if @error = -1 Then
		$herror = FileOpen(@ScriptDir & "\Erreurs.txt", $FO_APPEND)
		FileWriteLine($herror, "Fichier (PDF) sur BioWeb non trouvé : " & "Y:\CONSER\pdf\V_1" & $aList1[$i] & ".pdf")
		FileClose($herror)
	Else
		if stringleft(FileReadLine($hFileOpen), 8) = "%PDF-1.4" then
			ReDim $aList2[UBound($aList2) + 1]
			$aList2[0] = UBound($aList2) - 1
			$aList2[UBound($aList2) - 1] = $aList1[$i]
		EndIf
	EndIf
	FileClose($hFileOpen)
Next

;VERIFICATION POUR CHAQUE DOSSIER CERBA QUE L'EXEMPLAIRE MEDECIN "-1.PDF" EXISTE SINON LE COPIER
$aCERBA = _FileListToArrayRec("Y:\SAUVPDF\2017", "*.PDF", 1, 1, 1, 2)
$aCERBA17 = _FileListToArrayRec("Y:\SAUVPDF\2018", "*.PDF", 1, 1, 1, 2)
$aCERBA[0]=$aCERBA[0]+$aCERBA17[0]
_ArrayDelete ($aCERBA17, 0 )
_ArrayAdd($aCERBA, $aCERBA17)

Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
;~ _ArrayDisplay($aCERBA)
$hOK = FileOpen(@ScriptDir & "\Ok.txt", $FO_APPEND)
for $i = 1 to UBound($aCERBA) - 1
	if stringright($aCERBA[$i], 6) <> "-1.PDF" then
		$aPathSplit = _PathSplit($aCERBA[$i], $sDrive, $sDir, $sFileName, $sExtension)
		if StringLeft(stringright($aCERBA[$i], 8), 1) = "_" Then ;FORMAT AAAMMJJDDDD_P-X.PDF
			if _ArraySearch($aCERBA, stringsplit($aCERBA[$i], "_")[1] & "-1.PDF") = -1 Then
				$ret = FileCopy($aCERBA[$i], stringsplit($aCERBA[$i], "_")[1] & "-1.PDF")
				if $ret = 1 Then FileWriteLine($hOK, _Now() & " : PDF CERBA COPIE : " & $aCERBA[$i] & " -> " & stringsplit($aCERBA[$i], "_")[1] & "-1.PDF")
			Endif
		ElseIf StringLeft(stringright($aCERBA[$i], 8), 1) <> "_" And StringLeft(stringright($aCERBA[$i], 6), 1) = "-" Then ;FORMAT AAAMMJJDDDD-X.PFF
			if _ArraySearch($aCERBA, $sDrive & $sDir & stringleft(stringright($aCERBA[$i], 17), 11) & "-1.PDF") = -1 Then
				$ret = FileCopy($aCERBA[$i], $sDrive & $sDir & stringleft(stringright($aCERBA[$i], 17), 11) & "-1.PDF")
				if $ret = 1 Then FileWriteLine($hOK, _Now() & " : PDF CERBA COPIE : " & $aCERBA[$i] & " -> " & $sDrive & $sDir & stringleft(stringright($aCERBA[$i], 17), 11) & "-1.PDF")
			Endif
		Else
		EndIf
	EndIf
Next
FileClose($hOK)

;VERIFICATION POUR CHAQUE DOSSIER CERBA QUE LA VERSION SUR BIOWEB EST > A 1.5 ET CREATION DE LA LISTE DE TRAVAIL $aCERBA2
;~ $aCERBA = _FileListToArrayRec("Y:\SAUVPDF\2016", "*.PDF", 1, 1, 1, 2)
;~ $aCERBA17 = _FileListToArrayRec("Y:\SAUVPDF\2017", "*.PDF", 1, 1, 1, 2)
;~ $aCERBA[0]=$aCERBA[0]+$aCERBA17[0]
;~ _ArrayDelete ($aCERBA17, 0 )
;~ _ArrayAdd($aCERBA, $aCERBA17)

$aCERBA = _FileListToArrayRec("Y:\SAUVPDF\2018", "*.PDF", 1, 1, 1, 2)

local $aCERBA2[1]
for $i = 1 to UBound($aCERBA) - 1

	GUICtrlSetData($Progress1, 17.5 + ($i / (UBound($aCERBA) - 1)) * 2.5)
	GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 4/5" & @CRLF & "Vérification des fichiers sur BioWeb (" & $aCERBA[$i] & ")")

	if stringright($aCERBA[$i], 6) = "-1.PDF" then

		$hFileOpen = FileOpen("Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA[$i], 17), 11) & ".pdf", $FO_READ + $FO_UTF8)
		if @error = -1 Then
			$herror = FileOpen(@ScriptDir & "\Erreurs.txt", $FO_APPEND)
			FileWriteLine($herror, "Fichier CERBA sur BioWeb non trouvé : " & "Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA[$i], 17), 11) & ".pdf")
			FileClose($herror)
		Else
			if stringleft(FileReadLine($hFileOpen), 8) = "%PDF-1.4" then
				ReDim $aCERBA2[UBound($aCERBA2) + 1]
				$aCERBA2[0] = UBound($aCERBA2) - 1
				$aCERBA2[UBound($aCERBA2) - 1] = $aCERBA[$i]
			EndIf
		EndIf
		FileClose($hFileOpen)
	EndIf
Next

;SUPPRESSION DES FICHIERS DOUBLONS HPR DANS LE REPERTOIRE D:\HPR2HM\COPIE
GUICtrlSetData($Label1, "Chargement..." & @CRLF & "Etape 5/5" & @CRLF & "Suppression des doublons HPR...")
GUICtrlSetData($Progress1, 20)
$aHPR = _FileListToArray("Y:\HPR2HM\COPIE", "*.OK", 1, 0)
$aHPR2 = _FileListToArray("Y:\HPR2HM\COPIE", "*.pdf", 1, 0)
for $i = 1 to $aHPR2[0]
	GUICtrlSetData($Progress1, 20 + ($i / $aHPR2[0]))
	Local $iIndex2 = _ArrayFindAll($aHPR, StringRight(StringLeft($aHPR2[$i], 13), 11), 0, 0, 0, 1, 0, 0)
	if @error <> 6 AND UBound($iIndex2) > 1 Then
		local $ares[UBound($iIndex2)][2]
		for $j = 0 to UBound($iIndex2) - 1
			$ares[$j][0] = stringright(stringsplit($aHPR[$iIndex2[$j]], "-")[1], 8) & stringsplit($aHPR[$iIndex2[$j]], "-")[2]
			$ares[$j][1] = $aHPR[$iIndex2[$j]]
		Next
		_ArraySort($ares)
		for $j = 0 to UBound($ares) - 2
			FileDelete("Y:\HPR2HM\COPIE\" & StringSplit($ares[$j][1], ".")[1] & ".HPR")
			FileDelete("Y:\HPR2HM\COPIE\" & StringSplit($ares[$j][1], ".")[1] & ".OK")
		Next
	EndIf
Next
$aHPR = _FileListToArray("Y:\HPR2HM\COPIE", "*.ok", 1, 0)
GUICtrlSetData($Progress1, 25)
#EndRegion ### CHARGEMENTS DES FICHIERS ET DES TABLES ###


$hOK = FileOpen(@ScriptDir & "\Ok.txt", $FO_APPEND)
#Region ### IMPORTATION DES FICHIERS SCANNES SUR BIOWEB ### LISTE DE TRAVAil $aLIST2 ###
For $i = 1 to $aList2[0]
	GUICtrlSetData($Label1, "Traitement PDF (BioWeb)..." & @CRLF & "Fichier " & $i & "/" & $aList2[0] & @CRLF & "Dossier n°" & $aList2[$i])
	GUICtrlSetData($Progress1, 25 + ($i / ($aList2[0])) * 12.5)
	FileDelete(@ScriptDir & "\createpdf\a.pdf")
	FileDelete(@ScriptDir & "\createpdf\b.pdf")
	Local $iIndex2 = _ArrayFindAll($aSCAN, $aList2[$i], 0, 0, 0, 1, 0, 0)
	Local $hFileOpenppdf = FileOpen(@ScriptDir & "\createpdf\pdf.ntr", $FO_OVERWRITE)
	For $j = 0 To UBound($iIndex2) - 1
		FileWriteLine($hFileOpenppdf, "Y:\SCANNER\" & $aSCAN[$iIndex2[$j]][3])
	Next
	FileClose($hFileOpenppdf)
	Runwait(@ScriptDir & "\createpdf\createpdf.exe")
	RunWait("pdftk.exe " & "Y:\CONSER\pdf\T_1" & $aList2[$i] & ".pdf " & @ScriptDir & "\createpdf\a.pdf cat output " & @ScriptDir & "\createpdf\b.pdf")
	FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\CONSER\pdf\T_1" & $aList2[$i] & ".pdf", 1)



	RunWait("pdftk.exe " & "Y:\CONSER\pdf\V_1" & $aList2[$i] & ".pdf " & @ScriptDir & "\createpdf\a.pdf cat output " & @ScriptDir & "\createpdf\b.pdf")
	FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\CONSER\pdf\V_1" & $aList2[$i] & ".pdf", 1)
	FileWriteLine($hOK, _Now() & " : PDF Importé sur BioWeb : " & "Y:\CONSER\pdf\T_1" & $aList2[$i] & ".pdf")
Next
#EndRegion ### IMPORTATION DES FICHIERS SCANNES SUR BIOWEB ### LISTE DE TRAVAil $aLIST2 ###


#Region ### IMPORTATION DES FICHIERS SCANNES SUR HOPITAL MANAGER ###
for $i = 1 to UBound($aHPR, 1) - 1
	GUICtrlSetData($Label1, "Traitement PDF (HM) ..." & @CRLF & "Fichier " & ($i - 1) & "/" & UBound($aHPR, 1) - 1 & @CRLF & "Dossier n°" & stringleft(stringsplit($aHPR[$i], "-")[3], 11))
	GUICtrlSetData($Progress1, 37.5 + (($i - 1) / (UBound($aHPR, 1) - 2)) * 12.5)
	FileDelete(@ScriptDir & "\createpdf\a.pdf")
	FileDelete(@ScriptDir & "\createpdf\b.pdf")
	if FileExists("Y:\HPR2HM\COPIE\" & stringleft($aHPR[$i], 31) & ".HPR") AND FileExists("Y:\HPR2HM\COPIE\" & stringleft($aHPR[$i], 31) & ".OK") AND FileExists("Y:\HPR2HM\COPIE\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf") Then ; Si les fichiers ok, hprim et pdf existent
		Local $hFileOpen2 = FileOpen("Y:\HPR2HM\COPIE\" & stringleft($aHPR[$i], 31) & ".HPR", $FO_READ); $hFileOpen2 = Fichier source
		Local $hFileOpen3 = FileOpen("Y:\HPRRES\" & stringleft($aHPR[$i], 31) & ".HPR", $FO_OVERWRITE) ; $hFileOpen3 = Fichier destination
		$txtobr = ""
		$txtpre = ""
		$codepre = ""
		$bact = 0
		$cerba = 0

		While 1 = 1
			$txt = FileReadLine($hFileOpen2) ; Lecture de la ligne du fichier source
			if @error = -1 then ExitLoop ; Si on est à la fin, on quitte la bouche
			if StringSplit($txt, "|")[1] = "OBX" Then ; Si le début est OBX
				if $txtobr <> "" Then ; Si le buffer $txtobr est plein, on l'écrit et on le supprime
					FileWriteLine($hFileOpen3, $txtobr) ; On écrit l'OBR
					$txtobr = ""
				EndIf

				if StringSplit(StringSplit($txt, "|")[4], "~")[1] = "CR_SGL" Then ; Si CR_SGL, on recopie la ligne dans le fichier source
					FileWriteLine($hFileOpen3, $txt)
				EndIf


				#Region ### TRAITEMENT FICHIERS SCANNES ###
				if StringSplit(StringSplit($txt, "|")[4], "~")[1] = "*BACT01" Then ; Si *BACT01, on crée le nouveau pdf
					$t = stringleft(stringsplit($aHPR[$i], "-")[3], 11) ; on stocke le nom du dossier
					Local $iIndex2 = _ArrayFindAll($aSCAN, $t, 0, 0, 0, 1, 0, 0) ; On recherche le nom du dossier dans la liste des images scannées
					if @error <> 6 Then ; si quelque chose a été trouvé
						Local $hFileOpenppdf = FileOpen(@ScriptDir & "\createpdf\pdf.ntr", $FO_OVERWRITE) ; on ouvre le fichier .ntr
						For $j = 0 To UBound($iIndex2) - 1
							FileWriteLine($hFileOpenppdf, "Y:\SCANNER\" & $aSCAN[$iIndex2[$j]][3]) ; on écrit le répertoire et le nom de chaque image
						Next
						FileClose($hFileOpenppdf)
						Runwait(@ScriptDir & "\createpdf\createpdf.exe") ; on crée le pdf --> a.pdf
						RunWait("pdftk.exe " & "Y:\HPR2HM\COPIE\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf " & @ScriptDir & "\createpdf\a.pdf cat output " & @ScriptDir & "\createpdf\b.pdf") ; On fusion le pdf créé par biowin et le pdf des images
;~ 						Filemove("Y:\HPR2HM\COPIE\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", @ScriptDir & "\backup\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", 1) ; on sauvergarde le pdf d'origine de biowin
						FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\HPRRES\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", 1) ; on copie le fichier fusionné
						$bact = 1 ; c'était un dossier de bactériologie
					EndIf
				EndIf
				#EndRegion ### TRAITEMENT FICHIERS SCANNES ###

				#Region ### TRAITEMENT PDF CERBA ###
				if StringSplit(StringSplit($txt, "|")[4], "~")[1] = "*CERBA03" Then ; Si *CERBA03, on crée le nouveau pdf
					$t = stringleft(stringsplit($aHPR[$i], "-")[3], 11) ; on stocke le nom du dossier
;~ 							   ConsoleWrite("OK"&$t&@CRLF)
					Local $iIndex2 = _ArrayFindAll($aCERBA, $t & "-1.PDF", 0, 0, 0, 1, 0, 0) ; On recherche le nom du dossier dans la liste des images scannées
					if @error <> 6 Then ; si quelque chose a été trouvé

						RunWait("pdftk.exe " & "Y:\HPR2HM\COPIE\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf " & $aCERBA[$iIndex2[0]] & " cat output " & @ScriptDir & "\createpdf\b.pdf") ; On fusion le pdf créé par biowin et le pdf des images
;~ 						Filemove("Y:\HPR2HM\COPIE\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", @ScriptDir & "\backup\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", 1) ; on sauvergarde le pdf d'origine de biowin
						FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\HPRRES\T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf", 1) ; on copie le fichier fusionné
						$cerba = 1 ; c'était un dossier cerba
					EndIf
				EndIf
				#EndRegion ### TRAITEMENT PDF CERBA ###

				$codepre = "OBX"

			Elseif StringSplit($txt, "|")[1] = "C" Then ; Si le début est C ==> On suprime la ligne
				$codepre = "C"

			Elseif StringSplit($txt, "|")[1] = "OBR" Then ; Si le début est OBR, on modifie l'entete d'affichage
				$txtpre = $txt
				Local $atemp = StringSplit($txt, "|")
				Local $atemp2 = StringSplit($atemp[5], "^")
				$txt2 = "CR_SGL~COMPTE RENDU PDF^"
				$atemp[5] = StringLeft($txt2, StringLen($txt2) - 1)
				$txt = ""
				for $j = 1 to $atemp[0]
					$txt = $txt & $atemp[$j] & "|"
				Next
				$txtobr = StringLeft($txt, StringLen($txt) - 1)
				$codepre = "OBR"

			Elseif StringSplit($txt, "|")[1] = "A" Then ; Si le début est A

				if $codepre = "OBR" Then ; Si le code précédent est OBR
					$txt = $txtpre & StringRight($txt, StringLen($txt) - 2)
					$txtpre = $txt
					Local $atemp = StringSplit($txt, "|")
					Local $atemp2 = StringSplit($atemp[5], "^")
					$txt2 = "CR_SGL~COMPTE RENDU PDF^"
					$atemp[5] = StringLeft($txt2, StringLen($txt2) - 1)

					$txt = ""
					for $j = 1 to $atemp[0]
						$txt = $txt & $atemp[$j] & "|"
					Next
					$txtobr = StringLeft($txt, StringLen($txt) - 1)
				Else ; Sinon on supprime la ligne
				EndIf


			Else ; Si le code est ni OBX, ni C, ni A, ni OBR, on recopie la ligne telle quelle
				FileWriteLine($hFileOpen3, $txt)
				$codepre = ""
			EndIf

		WEnd

		FileClose($hFileOpen2)
		FileClose($hFileOpen3)

		if $bact = 1 or $cerba = 1 Then ; Si c'est un dossier de bactério ou cerba, on le déplace les fichiers HPR et OK dans le dossier sur surveillé par HM

			FileMove("Y:\HPR2HM\COPIE\" & stringleft($aHPR[$i], 31) & ".OK", @ScriptDir & "\backup\" & stringleft($aHPR[$i], 31) & ".OK", 1)
			FileMove("Y:\HPR2HM\COPIE\" & stringleft($aHPR[$i], 31) & ".HPR", @ScriptDir & "\backup\" & stringleft($aHPR[$i], 31) & ".HPR", 1)
			FileCopy(@ScriptDir & "\backup\" & stringleft($aHPR[$i], 31) & ".OK", "Y:\HPRRES\" & stringleft($aHPR[$i], 31) & ".OK", 1)
			if $bact = 1 Then FileWriteLine($hOK, _Now() & " : PDF Importé sur HM : T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf")
			if $cerba = 1 Then FileWriteLine($hOK, _Now() & " : CERBA Importé sur HM : T_" & stringleft(stringsplit($aHPR[$i], "-")[3], 11) & ".pdf")

		Else ; Sinon on détruit le ficher HPRIM créé
			FileDelete("Y:\HPRRES\" & stringleft($aHPR[$i], 31) & ".HPR")
		EndIf
	EndIf

Next
#EndRegion ### IMPORTATION DES FICHIERS SCANNES SUR HOPITAL MANAGER ###


#Region ### IMPORTATION DES PDF CERBA SCANNES SUR BIOWEB ### LISTE DE TRAVAil $aCERBA2 ###
For $i = 1 to $aCERBA2[0]
	GUICtrlSetData($Label1, "Traitement CERBA (BioWeb)..." & @CRLF & "Fichier " & ($i - 1) & "/" & $aCERBA2[0] & @CRLF & "Dossier n°" & stringleft(stringright($aCERBA2[$i], 17), 11))
	GUICtrlSetData($Progress1, 50 + (($i - 1) / (UBound($aHPR, 1) - 2)) * 25)

	FileDelete(@ScriptDir & "\createpdf\b.pdf")

	RunWait("pdftk.exe " & "Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf " & $aCERBA2[$i] & " cat output " & @ScriptDir & "\createpdf\b.pdf")
;~ 		FileCopy("Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i],17),11) & ".pdf", @ScriptDir & "\backup\T_1" & stringleft(stringright($aCERBA2[$i],17),11) & ".pdf", 1)
	FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 1)

	RunWait("pdftk.exe " & "Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf " & $aCERBA2[$i] & " cat output " & @ScriptDir & "\createpdf\b.pdf")
;~ 		FileCopy("Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i],17),11) & ".pdf", @ScriptDir & "\backup\V_1" & stringleft(stringright($aCERBA2[$i],17),11) & ".pdf", 1)
	FileCopy(@ScriptDir & "\createpdf\b.pdf", "Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 1)

	FileWriteLine($hOK, _Now() & " : CERBA Importé sur BioWeb : " & "Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf")

	;CHANGEMENT VERSION PDF BIOWEB DE 1.4 EN 1.5 (SI PDF CERBA EST EN 1.4)
	if _HexRead("Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, 1) = 0x34 Then
		_HexWrite("Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, Binary("0x35"))
		FileWriteLine($hOK, _Now() & " : CERBA Importé sur BioWeb : " & "Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf (modification version PDF 1.4 en 1.5 [" & _HexRead("Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, 1) & " à l'octet 0x7])")
	EndIf

	if _HexRead("Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, 1) = 0x34 Then
		_HexWrite("Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, Binary("0x35"))
		FileWriteLine($hOK, _Now() & " : CERBA Importé sur BioWeb : " & "Y:\CONSER\pdf\T_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf (modification version PDF 1.4 en 1.5 [" & _HexRead("Y:\CONSER\pdf\V_1" & stringleft(stringright($aCERBA2[$i], 17), 11) & ".pdf", 0x7, 1) & " à l'octet 0x7])")
	EndIf

Next
#EndRegion ### IMPORTATION DES PDF CERBA SCANNES SUR BIOWEB ### LISTE DE TRAVAil $aCERBA2 ###

FileClose($hOK)

GUICtrlSetData($Label1, "Traitement terminé, vous pouvez quitter le programme...")
GUICtrlSetData($Progress1, 100)
GUICtrlSetState($Button1, $GUI_ENABLE)

;~ While 1
;~ 	$nMsg = GUIGetMsg()
;~ 	Switch $nMsg
;~ 		Case $Button1
;~ 			Exit
;~ 		Case $GUI_EVENT_CLOSE
;~ 			Exit

;~ 	EndSwitch
;~ WEnd
Exit

Func _HexRead($FilePath, $Offset, $Length)
	Local $Buffer, $ptr, $fLen, $hFile, $Result, $Read, $err, $Pos

	;## Parameter Checks
	If Not FileExists($FilePath) Then Return SetError(1, @error, 0)
	$fLen = FileGetSize($FilePath)
	If $Offset > $fLen Then Return SetError(2, @error, 0)
	If $fLen < $Offset + $Length Then Return SetError(3, @error, 0)

	;## Define the dll structure to store the data.
	$Buffer = DllStructCreate("byte[" & $Length & "]")
	$ptr = DllStructGetPtr($Buffer)

	;## Open File
	$hFile = _WinAPI_CreateFile($FilePath, 2, 2, 0)
	If $hFile = 0 Then Return SetError(5, @error, 0)

	;## Move file pointer to offset location
	$Pos = $Offset
	$Result = _WinAPI_SetFilePointer($hFile, $Pos)
	$err = @error
	If $Result = 0xFFFFFFFF Then
		_WinAPI_CloseHandle($hFile)
		Return SetError(6, $err, 0)
	EndIf

	;## Read from file
	$Read = 0
	$Result = _WinAPI_ReadFile($hFile, $ptr, $Length, $Read)
	$err = @error
	If Not $Result Then
		_WinAPI_CloseHandle($hFile)
		Return SetError(7, $err, 0)
	EndIf

	;## Close File
	_WinAPI_CloseHandle($hFile)
	If Not $Result Then Return SetError(8, @error, 0)

	;## Return Data
	$Result = DllStructGetData($Buffer, 1)

	Return $Result
EndFunc   ;==>_HexRead

Func _HexWrite($FilePath, $Offset, $BinaryValue)
	Local $Buffer, $ptr, $bLen, $fLen, $hFile, $Result, $Written

	;## Parameter Checks
	If Not FileExists($FilePath) Then Return SetError(1, @error, 0)
	$fLen = FileGetSize($FilePath)
	If $Offset > $fLen Then Return SetError(2, @error, 0)
	If Not IsBinary($BinaryValue) Then Return SetError(3, @error, 0)
	$bLen = BinaryLen($BinaryValue)
	If $bLen > $Offset + $fLen Then Return SetError(4, @error, 0)

	;## Place the supplied binary value into a dll structure.
	$Buffer = DllStructCreate("byte[" & $bLen & "]")

	DllStructSetData($Buffer, 1, $BinaryValue)
	If @error Then Return SetError(5, @error, 0)

	$ptr = DllStructGetPtr($Buffer)

	;## Open File
	$hFile = _WinAPI_CreateFile($FilePath, 2, 4, 0)
	If $hFile = 0 Then Return SetError(6, @error, 0)

	;## Move file pointer to offset location
	$Result = _WinAPI_SetFilePointer($hFile, $Offset)
	$err = @error
	If $Result = 0xFFFFFFFF Then
		_WinAPI_CloseHandle($hFile)
		Return SetError(7, $err, 0)
	EndIf

	;## Write new Value
	$Result = _WinAPI_WriteFile($hFile, $ptr, $bLen, $Written)
	$err = @error
	If Not $Result Then
		_WinAPI_CloseHandle($hFile)
		Return SetError(8, $err, 0)
	EndIf

	;## Close File
	_WinAPI_CloseHandle($hFile)
	If Not $Result Then Return SetError(9, @error, 0)
EndFunc   ;==>_HexWrite
