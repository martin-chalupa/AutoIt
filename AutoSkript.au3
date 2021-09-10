#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <GuiScrollBars.au3>
#include <Date.au3>

; aktivni TP = Edit33
; vlozeni cisla dilu = Edit1
; cislo dilu = Edit26
; cislo op = Edit21
; cislo naslednika = Edit28
; cislo v priz materialu Edit34
; radek pricina zmen Edit2
; varianta (pro check F11) = Edit15
; Platnost do = Edit29

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

#Region ### START Koda GUI section ### Form=C:\Users\chalupa\Desktop\autoit\koda_1.7.3.0\Forms\AutoSkriptGUI.kxf
Global $g_bPaused = False
$Form1 = GUICreate("Zmena technologickych postupu", 345, 260, -1, -1)
$ButSpustit = GUICtrlCreateButton("Spustit", 16, 150, 150, 41)
GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
$TxtNazev = GUICtrlCreateInput("", 16, 113, 310, 21)
GUICtrlSetFont(-1, 8, 400, 0, "MS Sans Serif")
$Label1 = GUICtrlCreateLabel("Zadejte cestu k excel souboru s požadovanými informacemi (bez nebo včetně uvozovek):", 16, 77, 282, 28)
$Label2 = GUICtrlCreateLabel("Data pro správnou funkci skriptu musí být na PRVNÍM listě v zadaném sešitu.", 16, 17, 330, 50)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("Logovací soubor naleznete zde: " & @ScriptDir & "\LogSvazků.txt", 16, 205, 330, 50)
$Label3 = GUICtrlCreateLabel("Program je možné ukončit stiskem klávesy ESC.", 16, 233, 330, 50)
$ButExit = GUICtrlCreateButton("Zavřít", 176, 150, 150, 41)
GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
GUISetState(@SW_SHOW)
ControlFocus("Zmena technologickych postupu","", "Edit1")
#EndRegion ### END Koda GUI section ###

While 1
    HotKeySet("{ESC}", "Terminate")
	HotKeySet("{ENTER}", "StartSkriptu")
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		 Case $ButSpustit
			If GUICtrlRead($TxtNazev) = "" Then
			   HotKeySet("{ENTER}")
			   MsgBox($MB_ICONERROR,"Chyba","Zadejte cestu k souboru!")
			Else
			   ZmenaTP()
			EndIf
		 Case $ButExit
			Exit
	EndSwitch
WEnd

Func Terminate()
   Exit
EndFunc

Func StartSkriptu()
   ControlClick("Zmena technologickych postupu","Spustit","Button1")
EndFunc

Func ZmenaTP()
HotKeySet("{ENTER}")
;Nacteni potrebnych daz excelu ve formatu cislo svazku; pocet 9.3218.000; pocet 9.4820.000
Local $oExcel = _Excel_Open()
If StringInStr(GUICtrlRead($TxtNazev),'"') <> 0 Then
   Local $oSesit = _Excel_BookOpen($oExcel, StringMid((GUICtrlRead($TxtNazev)),2,StringLen((GUICtrlRead($TxtNazev)))-2))
Else
   Local $oSesit = _Excel_BookOpen($oExcel, (GUICtrlRead($TxtNazev)))
EndIf
If @error Then
   MsgBox($MB_ICONERROR,"Chyba","Zadaná cesta neexistuje!")
   ControlSetText("Zmena technologickych postupu","", "Edit1", "")
   ControlFocus("Zmena technologickych postupu","", "Edit1")
   Return
EndIf
Local $iPocetRadku = $oSesit.Sheets(1).Range("A6000").End(-4162).Row
Local $aSvazky = _Excel_RangeRead($oSesit, Default, "A2:A" & $iPocetRadku)
Global $a3218 = _Excel_RangeRead($oSesit, Default, "B2:B" & $iPocetRadku)
Global $a4820 = _Excel_RangeRead($oSesit, Default, "C2:C" & $iPocetRadku)
AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("WinTitleMatchMode", 2)
Local $hWnd = WinActivate("[CLASS:OWL_Window]","")
;Zajisteni spravne masky v Xpertu
If StringInStr(WinGetTitle("[ACTIVE]"),"Udržování technolog. postupů") = 0 Then
   MsgBox($MB_ICONERROR,"Chyba","Nejste v masce Udržování technolog. postupů!")
   Exit
EndIf

If ControlGetFocus("Udržování technolog. postupů","") <> "Edit1" Then MouseClick("primary", 515, 85, 1, 0)

;Poslání čísla svazku
For $i = 0 to $iPocetRadku - 2
   Global $indexpole = $i
   Send($aSvazky[$i])
   Send("{ENTER}")
   Local $sBarva = 0
   Local $error = WinWaitActive("Výběr záhlaví technol. postupu","", 5)
   If $error = 0 Then
	  $sBarva = PixelGetColor(877,83, $hWnd)
	  If $sBarva = "16711680" Then
;~ 		 Logsvazku($aSvazky[$i], $sBarva)
		 ContinueLoop
	  Else
		 ErrorHandler($error)
	  EndIf
   EndIf
   Global $iPocetAktTP = 0
   Send("^{PGUP}")
   Sleep(100)
   Do
	  Send("{RIGHT}")
	  Sleep(100)
   Until ControlGetFocus("Výběr záhlaví technol. postupu") = "Edit33"
   ;Zjištění počtu aktuvních TP
   Do
	  Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
	  Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  If ControlGetText("Výběr záhlaví technol. postupu","","Edit33") = 1 Then $iPocetAktTP = $iPocetAktTP + 1
	  Send("{DOWN}")
	  Sleep(100)
	  $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
   Until $iPoziceBaru1 = $iPoziceBaru2
   Do
	  Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  Send("{PGUP}")
	  Sleep(100)
	  $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
   Until $iPoziceBaru1 = $iPoziceBaru2
   ;Výběr prvního aktivního TP (do žádného dalšíhon program nevstupuje). V případě více aktivních TP, je toto zapsáno do logu
   While ControlGetText("Výběr záhlaví technol. postupu","","Edit33") <> 1
	  Send("{DOWN}")
	  Sleep(100)
   Wend
   Send("^{RIGHT}")
   ProhledaniTP($sBarva, $aSvazky[$i])
   Send("{F1}")
   $error = WinWaitActive("Udržování technolog. postupů", "", 10)
   ErrorHandler($error)
Next
MsgBox($MB_ICONINFORMATION,"SUCCESS!!!","Všechny zadané svazky byly změněny!")
EndFunc

Func ProhledaniTP($sBarva, $sSvazek)
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")
   Local $error = WinWaitActive("Přehled pracovních operací","", 10)
   ErrorHandler($error)
   ;Zjištění režimu zobrazení (všetně materiálu nebo ne) a přepnutí na zobrazení s materiálem
   If ControlGetText("Přehled pracovních operací","","Edit15") = "5" Then
	  Send("{F11}")
	  $error = WinWaitActive("Přehled pracovních operací","", 10)
	  ErrorHandler($error)
	  Sleep(1000)
   EndIf
   ;Přemístění do prvního řádku
   Do
	  Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
	  Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  Send("{PGUP}")
	  Sleep(100)
	  $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
   Until $iPoziceBaru1 = $iPoziceBaru2
   Global $iPocet4820 = 0
   Global $iPocet3218 = 0
   ;Prohledávání pro zadaný materiál. Proběhne tolikrát, kolikrát byl počet materiálu zadán v datovém excelu. V souboru je nutné mít min 2 řádky
	  While $a4820[$indexpole] > $iPocet4820 Or $a3218[$indexpole] > $iPocet3218
		 Do
			Send("{RIGHT}")
			Sleep(100)
		 Until ControlGetFocus("Přehled pracovních operací") = "Edit26"
		 Local $sMaterial = ControlGetText("Přehled pracovních operací","","Edit26")
   ;Vyhledávání pozice materiálu s pomocí pozice posuvníku
		 While $sMaterial <> "9.4820.000" And $sMaterial <> "9.3218.000"
			Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
			Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
			Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
			Send("{DOWN}")
			Sleep(100)
			$sMaterial = ControlGetText("Přehled pracovních operací","","Edit26")
			$tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
			Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
			If $iPoziceBaru1 = $iPoziceBaru2 Then
			   MsgBox ($MB_ICONERROR,"Chyba","Nepodařilo se vyhledat všechny zadané ITRs!")
			   Exit
			EndIf
		 WEnd
		 If $sMaterial = "9.4820.000" Then
			$iPocet4820 = $iPocet4820 + 1
			If $iPocet4820 = 1 Then
			   Local $sNaslednik4820 = ZmenaMat($sMaterial)
			Else
			   Local $sNaslednik4820 = $sNaslednik4820 & ", " & ZmenaMat($sMaterial)
			EndIf
		 Else
			$iPocet3218 = $iPocet3218 + 1
			If $siPocet3218 = 1 Then
			   Local $sNaslednik3218 = ZmenaMat($sMaterial) & ", "
			Else
			   Local $sNaslednik3218 = $sNaslednik3218 & ZmenaMat($sMaterial) & ", "
			EndIf
		 EndIf
	  WEnd
   Send("{F1}")
   $error = WinWaitActive("Evidovat příčinu změny","", 10)
   ErrorHandler($error)
   If ControlGetFocus("Evidovat příčinu změny","") <> "Edit5" Then MouseClick("primary", 30, 195, 2, 0)
   Send("Smazání operací, přidání podskupiny ITRs")
   Sleep(100)
   Send("{ENTER}")
   $error = WinWaitActive("Výběr záhlaví technol. postupu","", 10)
   ErrorHandler($error)
   LogSvazku($sSvazek, $sNaslednik4820, $sNaslednik3128, $sBarva)
EndFunc

Func ZmenaMat($sNovyMat)
   Sleep(100)
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")
   Send("^{RIGHT}")
   Local $error = WinWaitActive("Prac. operace","", 10)
   ErrorHandler($error)
   Send("{F3}")
   $error = WinWaitActive("Přiřazené materiály","", 10)
   ErrorHandler($error)
   MouseClick("secondary", 30, 55, 1, 0)
   For $i = 1 to 2
	  Send("{DOWN}")
	  Sleep(100)
   Next
   Send("{ENTER}")
   $error = WinWaitActive("Potvrzení smazání", "", 10)
   ErrorHandler($error)
   MouseClick("primary", 75, 75, 1, 0) ;MouseClick("primary", 203, 75, 1, 0)=NE
   $error = WinWaitActive("Přiřazené materiály","", 10)
   ErrorHandler($error)
   MouseClick("primary", 780, -14, 1, 0)
   $error = WinWaitActive("Prac. operace","", 10)
   ErrorHandler($error)
   Send("{ENTER}")
   $error = WinWaitActive("Přehled pracovních operací","", 10)
   ErrorHandler($error)
   Send("{UP}")
   Sleep(100)
   Do
	  Send("{RIGHT}")
	  Sleep(100)
   Until ControlGetFocus("Přehled pracovních operací") = "Edit28"
   Local $sNaslednik = ControlGetText("Přehled pracovních operací","","Edit28")
   If StringLen($sNaslednik) > 4 Or StringLen($sNaslednik) < 1 Then
	  MsgBox($MB_ICONERROR,"Chyba","Hodnota následník nenalezena!")
	  Exit
   EndIf
   Local $aPos = ControlGetPos("Přehled pracovních operací","","Edit21")
   MouseClick("secondary", 75, $aPos[1] + 10, 1, 0)
   For $i = 1 to 3
	  Send("{DOWN}")
	  Sleep(100)
   Next
   Send("{ENTER}")
   $error = WinWaitActive("Potvrzení smazání","", 10)
   ErrorHandler($error)
   MouseClick("primary", 75, 75, 1, 0) ;MouseClick("primary", 203, 75, 1, 0)=NE
   While ControlGetText("Přehled pracovních operací","","Edit21") <> $sNaslednik
	  Send("{DOWN}")
	  Sleep(100)
   WEnd
   Send("^{RIGHT}")
   Send("{F3}")
   $error = WinWaitActive("Přiřazené materiály","", 10)
   ErrorHandler($error)
   MouseClick("primary", 715, 388, 1, 0)
   $error = WinWaitActive("Nové založení pozice kusovníku","", 10)
   ErrorHandler($error)
   MouseClick("primary", 203, 83, 1, 0)
   Send($sNovyMat & ".032")
   Send("{ENTER}")
   $error = WinWaitActive("Zpracovat pozice kusovníku","", 10)
   ErrorHandler($error)
   Send("1")
   Send("{ENTER}")
   $error = WinWaitActive("Přiřazené materiály","", 10)
   ErrorHandler($error)
   MouseClick("primary", 75, 385, 1, 0)
   $error = WinWaitActive("Prac. operace","", 10)
   ErrorHandler($error)
   Send("{ENTER}")
   $error = WinWaitActive("Přehled pracovních operací","", 10)
   ErrorHandler($error)
   Do
	  Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
	  Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  Send("{PGUP}")
	  Sleep(100)
	  $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
	  Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
   Until $iPoziceBaru1 = $iPoziceBaru2
   Return $sNaslednik
EndFunc

Func ErrorHandler($error)
   If $error = 0 Then
	  MsgBox($MB_ICONERROR,"Chyba","Xpert neodpovídá, skript bude ukončen!")
	  Exit
   EndIf
EndFunc

Func Logsvazku($sSvazek, $sNaslednik4820, $sNaslednik3218, $sChyba)
   Local $hndl = FileOpen(@ScriptDir & "\LogSvazků.txt",1)
   If $hndl = -1 Then
	  MsgBox($MB_ICONERROR,"Chyba","Nebyl nalezen soubor logu!")
	  Exit
   EndIf
   If $indexpole = 0 Then
	  FileWrite($hndl, @CRLF & _Now() & @CRLF)
   EndIf
   If $sChyba > 0 Then
	  FileWrite($hndl, $sSvazek & "; svazek neexistuje nebo je blokován uživatelem" & @CRLF)
   Else
	  If $iPocetAktTP > 1 Then
		 FileWrite($hndl, $sSvazek & ";" & $iPocet3218 & ";" & $iPocet4820 & "; více akt. postupů" & @CRLF)
	  Else
		 FileWrite($hndl, $sSvazek & ";" & $iPocet3218 & ";" & $iPocet4820 & @CRLF)
	  EndIf
   EndIf
   FileClose($hndl)
EndFunc
