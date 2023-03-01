#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <AutoItConstants.au3>

HotKeySet("{F5}", "StartStop") ;spustí/zastaví skript
HotKeySet("{ESC}", "VypnutiSkriptu") ;vypne skript

$On = False

While 1
    While $On = True
 	  Send("{F4}")
	  Sleep(100)
	  Send("+{TAB}")
	  Sleep(50)
	  Send("+{TAB}")
	  Sleep(50)
	  Send("+{TAB}")
	  Sleep(50)
	  Send("{UP}")
	  Sleep(50)
	  Send("{ENTER}")
	  Sleep(50)
	  Send("{ENTER}")
	  Sleep(50)
    WEnd
    Sleep(100)
WEnd

Func StartStop()
    If $On = False Then
        $On = True
    Else
        $On = False
    EndIf
 EndFunc

 Func VypnutiSkriptu()
    Exit
EndFunc
