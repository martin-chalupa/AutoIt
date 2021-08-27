
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <GuiScrollBars.au3>
#include <Date.au3>
#include <ScreenCapture.au3>
#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <GDIPlus.au3>
#include <WinAPIHObj.au3>


Test()



#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

;~ #Region ### START Koda GUI section ### Form=C:\Users\chalupa\Desktop\autoit\koda_1.7.3.0\Forms\AutoSkriptGUI.kxf
;~ Global $g_bPaused = False
;~ $Form1 = GUICreate("Graficke znazorneni nastrihaku", 345, 260, -1, -1)
;~ $ButSpustit = GUICtrlCreateButton("Spustit", 16, 150, 150, 41)
;~ GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
;~ $TxtNazev = GUICtrlCreateInput("", 16, 113, 310, 21)
;~ GUICtrlSetFont(-1, 8, 400, 0, "MS Sans Serif")
;~ $Label1 = GUICtrlCreateLabel("Zadejte cestu k excel souboru s po�adovan�mi informacemi (bez nebo v�etn� uvozovek):", 16, 77, 282, 28)
;~ $Label2 = GUICtrlCreateLabel("Data pro spr�vnou funkci skriptu mus� b�t na PRVN�M list� v zadan�m se�itu.", 16, 17, 330, 50)
;~ GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
;~ $Label3 = GUICtrlCreateLabel("Logovac� soubor naleznete zde: " & @ScriptDir & "\LogN�st�ih�k�.txt", 16, 205, 330, 50)
;~ $Label3 = GUICtrlCreateLabel("Program je mo�n� ukon�it stiskem kl�vesy ESC.", 16, 233, 330, 50)
;~ $ButExit = GUICtrlCreateButton("Zav��t", 176, 150, 150, 41)
;~ GUICtrlSetFont(-1, 14, 400, 0, "MS Sans Serif")
;~ GUISetState(@SW_SHOW)
;~ ControlFocus("Zmena technologickych postupu","", "Edit1")
;~ #EndRegion ### END Koda GUI section ###

;~ While 1
;~     HotKeySet("{ESC}", "Terminate")
;~ 	HotKeySet("{ENTER}", "StartSkriptu")
;~ 	$nMsg = GUIGetMsg()
;~ 	Switch $nMsg
;~ 		Case $GUI_EVENT_CLOSE
;~ 			Exit
;~ 		 Case $ButSpustit
;~ 			If GUICtrlRead($TxtNazev) = "" Then
;~ 			   HotKeySet("{ENTER}")
;~ 			   MsgBox($MB_ICONERROR,"Chyba","Zadejte cestu k souboru!")
;~ 			Else
;~ 			   ZmenaTP()
;~ 			EndIf
;~ 		 Case $ButExit
;~ 			Exit
;~ 	EndSwitch
;~  WEnd

Func Terminate()
   Exit
EndFunc

Func StartSkriptu()
   ControlClick("Graficke znazorneni nastrihaku","Spustit","Button1")
EndFunc

Func Nastrihaky()
   HotKeySet("{ENTER}")
   Local $oExcel = _Excel_Open()
   If StringInStr(GUICtrlRead($TxtNazev),'"') <> 0 Then
	  Local $oSesit = _Excel_BookOpen($oExcel, StringMid((GUICtrlRead($TxtNazev)),2,StringLen((GUICtrlRead($TxtNazev)))-2))
   Else
	  Local $oSesit = _Excel_BookOpen($oExcel, (GUICtrlRead($TxtNazev)))
   EndIf
   If @error Then
	  MsgBox($MB_ICONERROR,"Chyba","Zadan� cesta neexistuje!")
	  ControlSetText("Zmena technologickych postupu","", "Edit1", "")
	  ControlFocus("Zmena technologickych postupu","", "Edit1")
	  Return
   EndIf
   Local $iPocetRadku = $oSesit.Sheets(1).Range("A6000").End(-4162).Row
   Local $aSvazky = _Excel_RangeRead($oSesit, Default, "A2:A" & $iPocetRadku)
   Local $aOperace = _Excel_RangeRead($oSesit, Default, "B2:B" & $iPocetRadku)
   Local $aNazev = _Excel_RangeRead($oSesit, Default, "C2:C" & $iPocetRadku)
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")

   If StringInStr(WinGetTitle("[ACTIVE]"),"Udr�ov�n� technolog. postup�") = 0 Then
	  MsgBox($MB_ICONERROR,"Chyba","Nejste v masce Udr�ov�n� technolog. postup�!")
	  Exit
   EndIf

   If ControlGetFocus("Udr�ov�n� technolog. postup�","") <> "Edit1" Then MouseClick("primary", 515, 85, 1, 0)
   For $i = 0 to $PocetRadku - 2
	  Local $indexSvazku = $i
	  Send($aSvazky[$i])
	  Send("{ENTER}")
	  Local $error = WinWaitActive("V�b�r z�hlav� technol. postupu","", 10)
	  ErrorHandler($error)
	  Send("^{PGUP}")
	  Sleep(100)
	  Do
		 Send("{RIGHT}")
		 Sleep(100)
	  Until ControlGetFocus("V�b�r z�hlav� technol. postupu") = "Edit33"
	  While ControlGetText("V�b�r z�hlav� technol. postupu","","Edit33") <> 1
		 Send("{DOWN}")
		 Sleep(100)
	  Wend
	  Send("^{RIGHT}")
      $error = WinWaitActive("P�ehled pracovn�ch operac�","", 10)
	  ErrorHandler($error)

	  Do
		 Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 Send("{PGUP}")
		 Sleep(100)
	  Until DllStructGetData($tSCROLLBARINFO, "xyThumbTop") = "17"
	  Local $sCOP = ControlGetText("P�ehled pracovn�ch operac�","","Edit21")
	  While $aOperace[$i] <> $sCOP
		 Local $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 Local $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 Local $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
		 Send("{DOWN}")
		 Sleep(100)
		 $sCOP = ControlGetText("P�ehled pracovn�ch operac�","","Edit21")
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 Local $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
		 If $iPoziceBaru1 = $iPoziceBaru2 Then
			MsgBox ($MB_ICONERROR,"Chyba","Nepoda�ilo se vyhledat zadanou operaci!")
			Exit
		 EndIf
	  WEnd
	  Send("^{RIGHT}")
	  $error = WinWaitActive("Prac. operace","", 10)
	  ErrorHandler($error)
	  Send("{F7}")
	  $error = WinWaitActive("N�st�ihov� pl�n vodi��","", 10)
	  ErrorHandler($error)
	  Do
		 $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 Send("{PGUP}")
		 Sleep(100)
	  Until DllStructGetData($tSCROLLBARINFO, "xyThumbTop") = "17"
	  Local $iPocetVodicu = 0
	  Do
		 $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
		 $iPocetVodicu = $iPocetVodicu + 1
		 Send("{DOWN}")
		 Sleep(100)
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  Until $iPoziceBaru1 = $iPoziceBaru2
	  _GDIPlus_Startup()
	  Local $iPocetScreenu = Round($iPocetVodicu/2) + 1
	  Local $hBitmap[$iPocetScreenu]
	  Local $hImage[$iPocetScreenu]
	  $hBitmap[0] = _ScreenCapture_CaptureWnd ("", $hWnd, 44, 75, 1484, 315+$iPocetVodicu*21)
	  $hImage[0] = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap[0])
	  Do
		 Send("{F11}")
		 Sleep(1000)
	  Until StringLen(ControlGetText("N�st�ihov� pl�n vodi��","","Edit21")) < 8
	  For $i = 1 To $iPocetScreenu - 1
		 Local $hWnd = WinActivate("N�st�ihov� pl�n vodi��","")
		 $hBitmap[$i] = _ScreenCapture_CaptureWnd ("", $hWnd, 210, 300, 1650, 670)
		 $hImage[$i] = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap[$i])
		 Send("{PGDN}")
		 Sleep(200)
	  Next
	  Local $hGUI = GUICreate("",1440,($iPocetScreenu - 1) * 370 + (240 + $iPocetVodicu * 21))
	  Local $hGraphicGUI = _GDIPlus_GraphicsCreateFromHWND($hGUI)
	  Local $hBMPBuff = _GDIPlus_BitmapCreateFromGraphics(1440,($iPocetScreenu - 1) * 370 + (240 + $iPocetVodicu * 21), $hGraphicGUI)
	  Local $hGraphic = _GDIPlus_ImageGetGraphicsContext($hBMPBuff)
	  If $indexsvazku = 0 Then
		 DirCreate(@ScriptDir & "\Nastrihaky")
	  EndIf
	  _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[0], 0, 0)
	  _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[1], 0, 240 + $iPocetVodicu * 21)

	  For $i = 2 To $iPocetScreenu - 1
		 _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[$i], 0, (240 + $iPocetVodicu * 21) + (($i - 1) * 370))
	  Next
	  _GDIPlus_ImageSaveToFile($hBMPBuff, @ScriptDir & "\Nastrihaky\" & $aNazev[$indexSvazku])
	  _GDIPlus_GraphicsDispose($hGraphic)
	  For $i = 0 to $iPocetScreenu - 1
		 _GDIPlus_ImageDispose($hImage[$i])
		 _WinAPI_DeleteObject($hBitmap[$i])
		 _GDIPlus_Shutdown()
	  Next
	  LogNastrihaku($aSvazek[$indexSvazku],$indexSvazku)
	  Do
		 Send("{F1}")
		 Sleep(100)
	  Until StringLen(ControlGetText("N�st�ihov� pl�n vodi��","","Edit21")) > 8
	  Send("{F1}")
	  Sleep(100)
	  $error = WinWaitActive("Prac. operace","", 10)
	  ErrorHandler($error)
	  Send("{F1}")
	  Sleep(100)
	  $error = WinWaitActive("P�ehled pracovn�ch operac�","", 10)
	  ErrorHandler($error)
	  Send("{F1}")
	  Sleep(100)
	  $error = WinWaitActive("V�b�r z�hlav� technol. postupu","", 10)
	  ErrorHandler($error)
	  Send("{F1}")
	  Sleep(100)
	  $error = WinWaitActive("Udr�ov�n� technolog. postup�","", 10)
	  ErrorHandler($error)
   Next
EndFunc

Func ErrorHandler($error)
   If $error = 0 Then
	  MsgBox($MB_ICONERROR,"Chyba","Xpert neodpov�d�, skript bude ukon�en!")
	  Exit
   EndIf
EndFunc

Func LogNastrihaku($sSvazek, $indexpole)
   Local $hndl = FileOpen(@ScriptDir & "\LogN�st�ih�k�.txt",1)
   If $hndl = -1 Then
	  MsgBox($MB_ICONERROR,"Chyba","Nebyl nalezen soubor logu!")
	  Exit
   EndIf
   If $indexpole = 0 Then
	  FileWrite($hndl, @CRLF & _Now() & @CRLF)
   EndIf
   FileWrite($hndl, $sSvazek & @CRLF)
   FileClose($hndl)
EndFunc

Func Test()
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")
	  Send("{F7}")
	  $error = WinWaitActive("N�st�ihov� pl�n vodi��","", 10)
	  ErrorHandler($error)
	  Do
		 $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 Send("{PGUP}")
		 Sleep(100)
	  Until DllStructGetData($tSCROLLBARINFO, "xyThumbTop") = "17"
	  Local $iPocetVodicu = 0
	  Do
		 $hwndCtrl = ControlGetHandle($hWnd,"","[CLASS:ScrollBar; INSTANCE:1]")
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 $iPoziceBaru1 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
		 $iPocetVodicu = $iPocetVodicu + 1
		 Send("{DOWN}")
		 Sleep(100)
		 $tSCROLLBARINFO = _GUIScrollBars_GetScrollBarInfoEx($hwndCtrl, $OBJID_CLIENT)
		 $iPoziceBaru2 = DllStructGetData($tSCROLLBARINFO, "xyThumbTop")
	  Until $iPoziceBaru1 = $iPoziceBaru2
	  _GDIPlus_Startup()
	  Local $iPocetScreenu = Round($iPocetVodicu/2) + 1
	  Local $hBitmap[$iPocetScreenu]
	  Local $hImage[$iPocetScreenu]
	  $hBitmap[0] = _ScreenCapture_CaptureWnd ("", $hWnd, 44, 75, 1484, 315+$iPocetVodicu*21)
	  $hImage[0] = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap[0])
	  Do
		 Send("{F11}")
		 Sleep(1000)
	  Until StringLen(ControlGetText("N�st�ihov� pl�n vodi��","","Edit21")) < 8
	  For $i = 1 To $iPocetScreenu - 1
		 Local $hWnd = WinActivate("N�st�ihov� pl�n vodi��","")
		 $hBitmap[$i] = _ScreenCapture_CaptureWnd ("", $hWnd, 210, 300, 1650, 670)
		 $hImage[$i] = _GDIPlus_BitmapCreateFromHBITMAP($hBitmap[$i])
		 Send("{PGDN}")
		 Sleep(200)
	  Next
	  Local $hGUI = GUICreate("",1440,($iPocetScreenu - 1) * 370 + (240 + $iPocetVodicu * 21))
	  Local $hGraphicGUI = _GDIPlus_GraphicsCreateFromHWND($hGUI)
	  Local $hBMPBuff = _GDIPlus_BitmapCreateFromGraphics(1440,($iPocetScreenu - 1) * 370 + (240 + $iPocetVodicu * 21), $hGraphicGUI)
	  Local $hGraphic = _GDIPlus_ImageGetGraphicsContext($hBMPBuff)
	  _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[0], 0, 0)
	  _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[1], 0, 240 + $iPocetVodicu * 21)

	  For $i = 2 To $iPocetScreenu - 1
		 _GDIPlus_GraphicsDrawImage($hGraphic, $hImage[$i], 0, (240 + $iPocetVodicu * 21) + (($i - 1) * 370))
	  Next

	  _GDIPlus_ImageSaveToFile($hBMPBuff, "C:\Users\chalupa\Desktop\autoit\Nastrihaky\128.03.000.10_0010.jpg")
	  _GDIPlus_GraphicsDispose($hGraphic)
	  For $i = 0 to $iPocetScreenu - 1
		 _GDIPlus_ImageDispose($hImage[$i])
		 _WinAPI_DeleteObject($hBitmap[$i])
		 _GDIPlus_Shutdown()
	  Next
EndFunc