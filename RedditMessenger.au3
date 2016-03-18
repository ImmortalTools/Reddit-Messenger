#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=RedditMessenger.ico
#AutoIt3Wrapper_Outfile=RedditMessenger-x86.exe
#AutoIt3Wrapper_Outfile_x64=RedditMessenger.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Description=Desktop Reddit notifications and PM
#AutoIt3Wrapper_Res_Fileversion=0.2.1.0
#AutoIt3Wrapper_Res_LegalCopyright=ImmortalTools
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=Made By|OlavOlsm, ImmortalTools
#AutoIt3Wrapper_Res_File_Add=Reply.png, rt_rcdata, Reply.png
#AutoIt3Wrapper_Res_File_Add=Ignore.png, rt_rcdata, Ignore.png
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

	AutoIt Version:...3.3.12.0
	Author:...........ImmortalTools

	Version:...0.2.1
	Changes:...Word wrap added to reply box and horizontall scrollbar removed
	...........Added support for more message types
	...........Improved view of message notifications
	...........Play Windows notify sound (Windows Notify.wav)
	...........Add check for update and notify when an update is available
	...........Fix that reply is sometimes "v" instead of the message
	...........Fix login prompt when user is not logged in
	...........Fix replying to second comments replies to the first
	...........Fix When different notification types received only one display
	...........Fix crash when getting private message reply
	...........Fix notification window not scaling correctly on title size
	...........Minor bug fixes
	ToDo:..... Added Compose new PM feature

#ce ----------------------------------------------------------------------------

#include <ie.au3>
#include <Array.au3>
#include <String.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GUIConstants.au3>
#include <Math.au3>
#include <Inet.au3>
#include "include/GUICtrlPic.au3"
#include "include/resources.au3"
#include "include/StringSize.au3"

AutoItSetOption ("TrayMenuMode", 1)	; remove default tray menu items
AutoItSetOption ("TrayOnEventMode", 1)	; Enable OnEvent functions notifications for the tray
AutoItSetOption ("SendKeyDownDelay", 10)
AutoItSetOption ("SendKeyDelay", 400)
TraySetToolTip ("Redit Messenger")
OnAutoItExitRegister ( "terminate" )

TrayCreateItem ("Compose new PM")	; tray item option for composing a new message
TrayItemSetOnEvent (-1, "Compose")	; compose a message when compose tray item chosen
TrayCreateItem ("Exit")	; create tray item option exit
TrayItemSetOnEvent (-1, "Terminate")	; exit app when tray item exit is clicked

;WinKill("messages: unread - Internet Explorer", "")
$oIE = _IECreate("http://www.reddit.com/message/unread/", 0, 0)
$hwnd = _IEPropertyGet($oIE, "hwnd")
Global $Version = "0.2.1", $PostTitle, $PostVia="?"

CheckForUpdate()

While 1
	; Get the inner text of the page (source code text)
	$body = _IEPropertyGet($oIE, "innertext")

	; Used for testing
	;$body=FileRead("PM.txt")
	;ClipPut($body)

	; Inform user about error and exit if an error has occurred
	; TODO: Error handling
	If @error Then
		MsgBox(48, "Reddit Messenger", "An error occured, Reddit Messenger will close." & @CRLF & "Error number: " & @Error & ", Extended error number: " & @Extended)
		Exit

	; Prompt the user to sign in to reddit if not signed in
	ElseIf StringInStr($body, "log in") Then
		; Get the current url so we know when the user has signed in
		Local $url = _IEPropertyGet($oIE, "locationurl")

		; Show browser and set the focus to the login field
		WinSetState($hwnd, "", @SW_SHOW)
		WinActivate($hwnd, "")
		Local $loginID = _IEGetObjById($oIE, "user_login")
		_IEAction($loginID, "focus")
		MsgBox(64, "Reddit Messenger", "Please log in to reddit", 5)

		; Wait for the location url to change, then user will be logged in
		WinActivate($hwnd, "")
		While (_IEPropertyGet($oIE, "locationurl") == $url)
			Sleep(1000)
		WEnd

		; Hide browser, wait for page to finish loading and go back to start of while
		WinSetState($hwnd, "", @SW_HIDE)
		_IELoadWait($oIE)
		ContinueLoop
	EndIf

	; Get the notification(s) if notification(s) has been received
	; Notifications has the word "context" or "permalink" in the inner text
	If StringInStr($body, "context") Or StringInStr($body, "permalink") Then

		; Figure out what kind of notification it is and get the message details
		Global $aMessages = GetMessages()

		; Code for testing, display the posts and number of posts
		;_ArrayDisplay($aMessages)
		;MsgBox(0, "test", (UBound($Post)))

		; Display each notification one by one
		If IsArray($aMessages) And UBound($aMessages) > 0 Then

			; Play windows notify sound
			SoundPlay(@WindowsDir & "\media\Windows Notify.wav")

			For $index = (UBound($aMessages)-1) To 0 Step -1
				; Display error and exit if there was an error
				; TODO: Error handling
				If @error=1 Then Exit MsgBox(48, "Reddit Messenger", "An error occured, Reddit Messenger will close." & @CRLF & "Error number: " & @Error & ", Extended error number: " & @Extended)
				; Code for testing (put the post in clipboard)
				;ClipPut($Post[$index])

				; Get post meta data
				$PostTitle = $aMessages[$index][0]
				$PostMsg = $aMessages[$index][1]
				$postUser = $aMessages[$index][2]
				$PostVia = $aMessages[$index][3]
				$PostTime = $aMessages[$index][4]

				DisplayMessage()
			Next
		EndIf
	EndIf

	Sleep(60000)
	_IENavigate($oIE, "http://www.reddit.com/message/unread/", 1)
WEnd

Func GetMessages()
	Local $asMessages[0]
	Local $aMessages[0][5]
	Local $PostTitles[0]

	; Get the messages from body text
	_ArrayConcatenate($asMessages, _StringBetween($body, "px; }", @CRLF & "reply" & @CRLF))
	_ArrayConcatenate($asMessages, _StringBetween($body, @CRLF & "reply" & @CRLF, @CRLF & "reply" & @CRLF))

	;Find message details and add to arrayæ
	For $i = 0 To UBound($asMessages) - 1

		;Find title
		$sTitle = _StringBetween($asMessages[$i], "", "from ")
		If Not IsArray($sTitle) Then $sTitle = _StringBetween($asMessages[$i], "", "via ")
		$sTitle = RemoveBlanks($sTitle[0])
		$sTitle = StringTrimRight($sTitle, 3)
		$sTitle = StringReplace($sTitle, "post reply", "")
		$sTitle = StringReplace($sTitle, "comment reply", "")
		$sTitle = StringReplace($sTitle, "username mention", "")

		;Find from
		$sFrom = _StringBetween($asMessages[$i], "from ", " via")
		If Not IsArray($sFrom) Then $sFrom = _StringBetween($asMessages[$i], "from ", " sent")
		If Not IsArray($sFrom) Then $sFrom = _StringBetween($asMessages[$i], "]", " via")
		$sFrom = $sFrom[0]
		$sFrom = StringReplace($sFrom, "[M]", "")

		;Find via
		$sVia = _StringBetween($asMessages[$i], "via /r/", " sent")
		If Not IsArray($sVia) Then $sVia = "PM"
		If IsArray($sVia) Then $sVia = "/r/" & $sVia[0]
		$sVia = StringReplace($sVia, "[M]", "")

		;Find time
		$sAgo = _StringBetween($asMessages[$i], "sent ", " ago")
		If Not IsArray($sAgo) Then $sAgo = "just now"
		If IsArray($sAgo) Then $sAgo = $sAgo[0]

		;Find message
		$sStart = $sAgo
		If $sAgo <> "just now" Then $sStart = "sent " & $sAgo & " ago"
		$sMessage = _StringBetween($asMessages[$i], $sStart, "permalink")
		If Not IsArray($sMessage) Then $sMessage = _StringBetween($asMessages[$i], $sStart, "context")
		$sMessage = StringReplace($sMessage[0], "show parent", "")
		$sMessage = StringStripWS($sMessage, $STR_STRIPLEADING + $STR_STRIPTRAILING)

		;Add message details to array
		_ArrayAdd($aMessages, $sTitle)
		$aMessages[$i][1] = $sMessage
		$aMessages[$i][2] = $sFrom
		$aMessages[$i][3] = $sVia
		$aMessages[$i][4] = $sAgo
	Next

	#CS Get the message post position in innertext and sort the posts
	If IsArray($Post) Then
		Local $Posts[UBound($Post)][3]
		For $index = 0 TO (UBound($Post)-1) Step 1
			$Posts[$index][0] = StringInStr($body, $Post[$index])
			$Posts[$index][1] = $PostTitles[$index]
			$Posts[$index][2] = $Post[$index]
		Next
	EndIf
	_ArraySort($Posts)
	#CE

	Return $aMessages
EndFunc

; Create GUI and display the notification
Func DisplayMessage()

	; Set the type of user, decides color on the user label
	$UserType = 2
	If $UserType = 1 Then $UserColor = $COLOR_GREEN
	If $UserType = 2 Then $UserColor = $COLOR_BLUE
	If $UserType = 3 Then $UserColor = $COLOR_RED

	If $PostTime <> "just now" Then $PostTime = $PostTime & " ago"

	; Calculate the size of GUI and elements according to desktop size and message size
	; Use StringSize to get the size of each part of the message
	$DesktopClientSize = _GetWorkingArea()
	$PUStringSize =_StringSize($PostUser, 9, 600, 0, "Tahoma")
	$PUStringWidth = $PUStringSize[2]
	$PUStringHeight = $PUStringSize[3]
	$PVStringSize = _StringSize($PostVia, 9, 600, 0, "Tahoma")
	$PVStringWidth = $PVStringSize[2]
	$PVStringHeight = $PVStringSize[3]
	$PTStringSize = _StringSize($PostTitle, 9, 600, 0, "Tahoma", 490)
	$PTStringWidth = $PTStringSize[2]
	$PTStringHeight = $PTStringSize[3]
	$PMStringSize = _StringSize($PostMsg, 10, 500, 0, "Tahoma", 490)
	$PMStringHeight = $PMStringSize[3]
	$PMStringWidth = $PMStringSize[2]
	$TitleStringSize = _StringSize("From " & $PostUser & " via "&$PostVia&" sent "&$PostTime, 9, 600, 0, "Tahoma", 490)
	$TitleStringWidth = $TitleStringSize[2]
	$TitleStringHeight = $TitleStringSize[3]
	$guiWidth = _Max(_Max($PMStringWidth, $TitleStringWidth), $PTStringWidth) + 10
	$guiWidth=_Max($guiWidth, 310) ;minimum width of gui should be 310
	$guiWidth = _Min($guiWidth, 500)	;maximum width of gui should be 500
	$guiHeight = _Min(64 + $TitleStringHeight + $PTStringHeight + $PMStringHeight, $DesktopClientSize[3])	; maximum height of gui should be desktop height
	$guiLeft = $DesktopClientSize[2] - $guiWidth - 2
	$guiTop = $DesktopClientSize[3] - $TitleStringHeight - $PTStringHeight - $PMStringHeight - 66

	; Create the GUI with previously decided size and position, set default font and background color
	GUICreate("Reddit Messenger", $guiWidth, $guiHeight, $guiLeft, $guiTop, $WS_POPUPWINDOW+$WS_EX_TOOLWINDOW, $WS_EX_TOPMOST+$WS_EX_TOOLWINDOW)
	GUISetFont(9, 600, 0, "Tahoma")
	GUISetBkColor($COLOR_WHITE)

	; Create GUI hyperlinks
	$UserURL = _GuiCtrlCreateHyperlink($PostUser, 42, 5, $PUStringWidth, $PUStringHeight)
	$SubURL = _GuiCtrlCreateHyperlink($PostVia, 65 + $PUStringWidth, 5, $PVStringWidth, $PVStringHeight)
	$ReplyURL = _GuiCtrlCreateHyperlink("", $guiWidth/2-150, $PMStringHeight+$PTStringHeight+48, 145, 30)
	$IgnoreURL = _GuiCtrlCreateHyperlink("", $guiWidth/2+5, $PMStringHeight+$PTStringHeight+48, 145, 30)

	; Create GUI labels
	GUICtrlCreateLabel("From ", 5, 5, 35, 15)
	GUICtrlCreateLabel("via", 45+$PUStringWidth, 5, 20)
	GUICtrlCreateLabel("sent "&$PostTime, 68 + $PUStringWidth + $PVStringWidth, 5)
	GUICtrlCreateLabel($PostTitle, 5, 27, $guiWidth-10, $PTStringHeight)
	GUISetFont(10, 500, 0, "Tahoma")
	GUICtrlCreateLabel($PostMsg, 5, 14 + $TitleStringHeight + $PTStringHeight, $guiWidth-10, $PMStringHeight)

	; Set images to GUI controls
	$Image = _GUICtrlPic_Create(@ScriptDir & "\Reply.png", $guiWidth/2-150, $PMStringHeight+$TitleStringHeight+$PTStringHeight+28, 145, 30, BitOR($SS_CENTERIMAGE, $SS_NOTIFY), Default)
	_ResourceSetImageToCtrl($Image, "Reply.png")
	$Image = _GUICtrlPic_Create(@ScriptDir & "\Ignore.png", $guiWidth/2+5, $PMStringHeight+$TitleStringHeight+$PTStringHeight+28, 145, 30, BitOR($SS_CENTERIMAGE, $SS_NOTIFY), Default)
	_ResourceSetImageToCtrl($Image, "Ignore.png")
	; TODO: Add unknown user icon
	;$Image = _GUICtrlPic_Create(@ScriptDir & "\user.png", 125, 2, 23, 25, BitOR($SS_CENTERIMAGE, $SS_NOTIFY), Default)

	GUISetState(@SW_SHOWNOACTIVATE)

	; Handle GUI events
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				_IEQuit($oIE)
				If WinExists($hwnd) Then WinKill($hwnd)
				GUIDelete()
				ExitLoop

			Case $UserURL
				ShellExecute("http://www.reddit.com/user/"&$PostUser)

			Case $SubURL
				ShellExecute("http://www.reddit.com"&$PostVia)

			Case $ReplyURL
				Reply()
				GUIDelete()
				ExitLoop

			Case $IgnoreURL
				GUIDelete()
				ExitLoop

			#CS Open inbox and focus on reply directly
			Case $QuickReply
				GUIDelete()
				_IELinkClickByText($oIE, "reply", $index)
				WinSetState($hwnd, "", @SW_SHOW)
				WinActivate($hwnd)
				ExitLoop
			#CE
		EndSwitch
	WEnd
EndFunc

#cs Func CreateMessageGUI()
	GUICreate ("Reddit Messenger", 500, 300, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetFont(12)
	$inputReply = GUICtrlCreateEdit ("", 0, 0, 500, 260, $ES_WANTRETURN + $WS_VSCROLL + $ES_AUTOVSCROLL)
	GUISetFont(16)
	$buttonReply = GUICtrlCreateButton ("Send", 0, 260, 250, 40)
	$buttonCancel = GUICtrlCreateButton ("Cancel", 250, 260, 250, 40)

	$buttons[0] = $buttonReply
	$buttons[1] = $buttonCancel

	GUISetState()
	return $buttons
EndFunc
#ce

; Show popup for replying to message
Func Reply ()
	GUICreate ("Reddit Messenger", 500, 300, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetFont(12)
	$inputReply = GUICtrlCreateEdit ("", 0, 0, 500, 260, $ES_WANTRETURN + $WS_VSCROLL + $ES_AUTOVSCROLL)
	GUISetFont(16)
	$buttonReply = GUICtrlCreateButton ("Send", 0, 260, 250, 40)
	$buttonCancel = GUICtrlCreateButton ("Cancel", 250, 260, 250, 40)

	GUISetState()

	While 1
		; Handle GUI events
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				Exit

			; Reply to the message when send button is clicked
			Case $buttonReply
				$outputReply = GUICtrlRead($inputReply)
				GUIDelete()
				_IELinkClickByText($oIE, "reply", $Index)
				ControlFocus ($hwnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]")
				Local $temp = ClipGet()

				; Add reply
				ClipPut($outputReply)
				ControlSend($hwnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "^v")

				; Send reply
				ClipPut($temp)
				ControlSend($hwnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{TAB}")
				ControlSend($hwnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{enter}")

				ExitLoop

			Case $buttonCancel
				GUIDelete()
				ExitLoop
		EndSwitch
	WEnd
EndFunc

; Compose a new message
Func Compose ()
	GUICreate ("Reddit Messenger", 500, 350, -1, -1, $WS_POPUP, $WS_EX_TOPMOST)
	GUISetFont(12)
	$inputUsername = GUICtrlCreateInput ("To (username)", 0, 0, 500, 25)
	GUICtrlSetTip(-1, "Username (recipient)")
	$inputTitle = GUICtrlCreateInput ("Subject", 0, 25, 500, 25)
	GUICtrlSetTip(-1, "Title")
	$inputReply = GUICtrlCreateEdit ("Message", 0, 50, 500, 260, $ES_WANTRETURN + $WS_VSCROLL + $ES_AUTOVSCROLL)
	GUICtrlSetTip(-1, "Message")
	GUISetFont(16)
	$buttonReply = GUICtrlCreateButton ("Send", 0, 310, 250, 40)
	$buttonCancel = GUICtrlCreateButton ("Cancel", 250, 310, 250, 40)

	GUISetState()

	While 1
		; Handle GUI events
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				Exit

			; Send the message when send button is clicked
			Case $buttonReply
				$outputUsername = GUICtrlRead($inputUsername)
				$outputTitle = GUICtrlRead($inputTitle)
				$outputReply = GUICtrlRead($inputReply)
				If $outputUsername = "" Or $outputUsername = "To (username)" Then
					MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Reddit Messenger", "Please enter the recipient (username)", 5)
					ContinueLoop
				ElseIf $outputTitle = "" or $outputTitle = "Subject" Then
					MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Reddit Messenger", "Please enter subject (title)", 5)
					ContinueLoop
				ElseIf $outputReply = "" or $outputReply = "Message" Then
					MsgBox($MB_ICONWARNING + $MB_TOPMOST, "Reddit Messenger", "Please enter your message (content)", 5)
					ContinueLoop
				EndIf
				GUIDelete()
				_IENavigate($oIE, "https://www.reddit.com/message/compose/", 1)

				; Get form and form elements
				$oForm = _IEGetObjById($oIE, "compose-message")
				$oFormTo = _IEFormElementGetObjByName($oForm, "to")
				$oformSubject = _IEFormElementGetObjByName($oForm, "subject")
				$oFormMessage = _IEFormElementGetObjByName($oForm, "text")
				$oformSend = _IEFormElementGetObjByName($oForm, "send")

				; Add the message data
				_IEFormElementSetValue($oFormTo, $outputUsername)
				_IEFormElementSetValue($oformSubject, $outputTitle)
				_IEFormElementSetValue($oFormMessage, $outputReply)
				_IEAction($oformSend, "click")

				; TODO: Exception handling (if message not sent)

				ExitLoop

			Case $buttonCancel
				GUIDelete()
				ExitLoop
		EndSwitch
	WEnd
EndFunc

Func CheckForUpdate()
	;read current version info from server.
	Local $UpdateCheck = 0
	Local $CheckForUpdate = _INetGetSource("http://immortaltools.com/CheckForUpdate/RedditMessenger_CheckForUpdate.txt")
	;Split the string and seperate version and update url.
	$CheckForUpdate = StringSplit(StringStripWS(StringStripCR($CheckForUpdate), 4), " " & @LF)
	If IsArray($CheckForUpdate) Then
		$UpdateCheck = $CheckForUpdate[1]
		If UBound($CheckForUpdate)=3 Then
			Local $UpdateURL = $CheckForUpdate[2]
		Else
			$UpdateURL = ""
		EndIf
	Else
		$Update = MsgBox(262144 + 4 + 16, "Reddit Messenger", "Could not check for update! " & @CRLF & "Would you like to check manually?")
		If $Update = 6 Then
			ShellExecute("https://www.immortaltools.com/redditmessenger/")
		EndIf
		Return
	EndIf

	;Check if there is a newer version and ask user to update if there is.
	If $UpdateCheck > $Version Then
		$Update = MsgBox(262144 + 4 + 64, "Reddit Messenger " & $Version, "An update is available! Would you like to update now?", 20)
		If $Update = 6 Then
			If $UpdateURL <> "E" And $UpdateURL <> "" Then
				ShellExecute($UpdateURL)
			Else
				ShellExecute("https://www.immortaltools.com/redditmessenger/")
			EndIf
		EndIf
	EndIf
EndFunc   ;==>CheckForUpdate

Func _GuiCtrlCreateHyperlink($s_Text, $i_Left, $i_Top, $i_Width = -1, $i_Height = -1, $i_Color = 0x0000ff, $s_ToolTip = '', $i_Style = -1, $i_ExStyle = -1)
	Local $i_CtrlID
	$i_CtrlID = GUICtrlCreateLabel($s_Text, $i_Left, $i_Top, $i_Width, $i_Height, $i_Style, $i_ExStyle)
	If $i_CtrlID <> 0 Then
		GUICtrlSetFont($i_CtrlID, -1, -1, 0)	; Set attribute 4 for underline
		GUICtrlSetColor($i_CtrlID, $i_Color)
		GUICtrlSetCursor($i_CtrlID, 0)
	EndIf
	Return $i_CtrlID
EndFunc   ;==>_GuiCtrlCreateHyperlink

Func cutBefore($string, $StringCutBefore)
    Return StringRight($String, StringLen($string) - StringInStr($String, $StringCutBefore)+1)
EndFunc

Func cutAfter($string, $StringCutAfter)
    Return StringLeft($String, StringInStr($String, $StringCutAfter) + StringLen($StringCutAfter)-1)
EndFunc

Func cutFrom($string, $StringCutFrom)
    Return StringRight($String, StringLen($string) - StringInStr($String, $StringCutFrom)-StringLen($StringCutFrom))
EndFunc

Func RemoveBlanks($index)
	$index=StringReplace($index, @CRLF, "")
	While StringLeft($index, 1)=" "
		$index=StringTrimLeft($index, 1)
	WEnd
	While StringRight($index, 1)=" "
		$index=StringTrimRight($index, 1)
	WEnd
	Return $index
EndFunc

Func RemoveLines($string, $opt)
	If $opt=1 Then
		$string=StringReplace($string, @CR, "")
		$string=StringReplace($string, @LF, "")
	ElseIf $opt=2 Then
		While StringIsAlNum(StringLeft($string, 1))=0
			Sleep(1)
			$string=StringTrimLeft($string, 1)
		WEnd
		While StringIsAlNum(StringRight($string, 1))=0
			Sleep(1)
			$string=StringTrimRight($string, 1)
		WEnd
	ElseIf $opt=3 Then
		While StringRight($string, 1)=@CR Or StringRight($string, 1)=@LF
			Sleep(1)
			$string=StringTrimRight($string, 1)
		WEnd
		While StringLeft($string, 1)=@CR Or StringLeft($string, 1)=@LF
			Sleep(1)
			$string=StringTrimLeft($string, 1)
		WEnd
	EndIf
	Return $string
EndFunc

Func _GetWorkingArea()
	#cs
		BOOL WINAPI SystemParametersInfo(UINT uiAction, UINT uiParam, PVOID pvParam, UINT fWinIni);
		uiAction SPI_GETWORKAREA = 48
	#ce
	Local $dRECT = DllStructCreate("long; long; long; long")
	Local $spiRet = DllCall("User32.dll", "int", "SystemParametersInfo", _
			"uint", 48, "uint", 0, "ptr", DllStructGetPtr($dRECT), "uint", 0)
	If @error Then Return 0
	If $spiRet[0] = 0 Then Return 0
	Local $aRet[4] = [DllStructGetData($dRECT, 1), DllStructGetData($dRECT, 2), DllStructGetData($dRECT, 3), DllStructGetData($dRECT, 4)]
	If $aRet[0] > 0 And $aRet[2] = @DesktopWidth Then $aRet[2] = $aRet[2] - $aRet[0] ;minus taskbar width if taskbar is on left side and $aRet[2] is desktopwidth
	If $aRet[1] > 0 And $aRet[3] = @DesktopHeight Then $aRet[3] = $aRet[3] - $aRet[1] ;minus taskbar height if taskbar is on top and $aRet[3] is desktopheight
	Return $aRet
EndFunc   ;==>_GetWorkingArea

Func MyErrFunc()
    ;$oError =$$oIEErrorHandler.description
EndFunc   ;==>MyErrFunc

Func Terminate ()
	_IEQuit($oIE)
	WinClose($hwnd)
	Exit
EndFunc