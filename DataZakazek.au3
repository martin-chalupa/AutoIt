#include <MsgBoxConstants.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <GuiScrollBars.au3>
#include <Date.au3>
; edit45
DatumZakazek()

Func DatumZakazek()
   HotKeySet("{ESC}", "Terminate")
   Local $oExcel = _Excel_Open()
   Local $oSesit = _Excel_BookOpen($oExcel,"C:\Users\chalupa\Desktop\autoit\Data zak�zek\Zakazky.xlsx")
   Local $iPocetRadku = $oSesit.Sheets(1).Range("A6000").End(-4162).Row
   Local $aZakazky = _Excel_RangeRead($oSesit, Default, "C2:C" & $iPocetRadku)
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")


   If ControlGetFocus("Zobrazit d�lensk� zak�zky","") <> "Edit1" Then MouseClick("primary", 680, 83, 1, 0)

   For $i = 0 to $iPocetRadku - 2
	  If ControlGetFocus("Zobrazit d�lensk� zak�zky","") <> "Edit1" Then MouseClick("primary", 680, 83, 1, 0)
	  Send($aZakazky[$i])
	  Send("{ENTER}")
	  Local $error = WinWaitActive("P�ehled pracovn�ch operac�","", 10)
	  ErrorHandler($error)
	  Send("^{RIGHT}")
	  Local $error = WinWaitActive("Pracovn� operace (detail)","", 10)
	  ErrorHandler($error)
	  MouseClick("primary", 263, 338, 1, 0)
	  Sleep(500)
	  MouseClick("primary", 1752, 558, 1, 0)
	  Local $sAdresa = "F" & $i + 2
	  Local $sDatum = ControlGetText("Pracovn� operace (detail)","","Edit45")
	  _Excel_RangeWrite($oSesit,Default,$sDatum,$sAdresa,True)
	  $hWnd = WinActivate("[CLASS:OWL_Window]","")
	  Send("{F1}")
	  Local $error = WinWaitActive("P�ehled pracovn�ch operac�","", 10)
	  ErrorHandler($error)
	  Send("{F1}")
	  Local $error = WinWaitActive("Zobrazit d�lensk� zak�zky","", 10)
	  ErrorHandler($error)
   Next
EndFunc

Func Terminate()
   Exit
EndFunc

Func ErrorHandler($error)
   If $error = 0 Then
	  MsgBox($MB_ICONERROR,"Chyba","Xpert neodpov�d�, skript bude ukon�en!")
	  Exit
   EndIf
EndFunc

Func Test()
   Local $oExcel = _Excel_Open()
   Local $oSesit = _Excel_BookOpen($oExcel,"C:\Users\chalupa\Desktop\autoit\Data zak�zek\Zakazky.xlsx")
   Local $oExcel = _Excel_Open()
   AutoItSetOption("MouseCoordMode", 2)
   AutoItSetOption("WinTitleMatchMode", 2)
   Local $hWnd = WinActivate("[CLASS:OWL_Window]","")
   For $i = 2 to 5
   	  Local $sAdresa = "F" & $i
	  Local $sDatum = ControlGetText("Pracovn� operace (detail)","","Edit45")
	  _Excel_RangeWrite($oSesit,Default,$sDatum,$sAdresa,True)
   Next
EndFunc
