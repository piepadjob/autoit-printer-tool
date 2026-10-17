#cs ----------------------------------------------------------------------------
 Script  : BatchPrint.au3
 Mô tả   : In hàng loạt file Word/Excel nằm trong một thư mục
           (thư mục nằm chung đường dẫn với tool - mặc định là thư mục chứa file .exe)
           Batch print Word/Excel files in a folder (folder in the same path as the tool).
 Tính năng / Features:
   + In toàn bộ file Word/Excel trong thư mục được chọn
   + Chọn khoảng trang cần in (mặc định: trang 1)
   + Delay (ms) giữa mỗi lệnh gửi tới máy in
   + Tùy chọn in toàn bộ các sheet của Excel
   + Giao diện song ngữ Anh/Việt (EN/VI) - Bilingual UI (EN/VI)
   + Nhật ký in, nút Dừng, bỏ qua file lock tạm (~$...)
 Biên dịch: SciTE for AutoIt -> Tools -> Compile (hoặc Aut2exe.exe) -> BatchPrint.exe
#ce ----------------------------------------------------------------------------

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>

; ======================== HẰNG SỐ ========================
Global Const $gsWordExt = ".doc|.docx|.docm|.dot|.dotx|.rtf"
Global Const $gsExcelExt = ".xls|.xlsx|.xlsm|.xlsb|.xlt|.xltx"

; ======================== BIẾN TOÀN CỤC ========================
Global $g_hGui
Global $g_idDir, $g_idPages, $g_idDelay, $g_idChkWord, $g_idChkExcel, $g_idChkAllSheets
Global $g_idBtnBrowse, $g_idBtnStart, $g_idBtnStop, $g_idLog, $g_idStatus
Global $g_idGrp1, $g_idGrp2, $g_idGrpLog, $g_idLblPages, $g_idLblPagesHint, $g_idLblDelay
Global $g_idLblLang, $g_idCmbLang
Global $g_bStop = False
Global $g_bPrinting = False
Global $g_bComError = False
Global $g_iDelay = 1000
Global $g_bAllSheets = True
Global $g_oWord = 0
Global $g_oExcel = 0
Global $g_sLang = "vi"
Global $g_oStr = 0

; Bắt lỗi COM toàn cục (không để lỗi COM làm treo tool)
Global $g_oComErr = ObjEvent("AutoIt.Error", "_ComErrorFunc")
OnAutoItExitRegister("_OnExit")

_InitStrings()
_Main()

; ======================== VÒNG LẶP CHÍNH ========================
Func _Main()
	_BuildGUI()
	_Log(_T("logReady"))
	While 1
		Local $iMsg = GUIGetMsg()
		Switch $iMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $g_idBtnBrowse
				_BrowseFolder()
			Case $g_idBtnStart
				_DoPrintBatch()
			Case $g_idBtnStop
				$g_bStop = True
				_Log(_T("logStopReq"))
			Case $g_idCmbLang
				_SwitchLanguage()
		EndSwitch
	WEnd
EndFunc

; ======================== GIAO DIỆN ========================
Func _BuildGUI()
	$g_hGui = GUICreate(_T("title"), 580, 600, -1, -1)

	; --- Ngon ngu / Language ---
	$g_idLblLang = GUICtrlCreateLabel(_T("lblLang"), 10, 14, 70, 20)
	$g_idCmbLang = GUICtrlCreateCombo("", 84, 10, 150, 200, $CBS_DROPDOWNLIST)
	GUICtrlSetData($g_idCmbLang, "Tiếng Việt|English", "Tiếng Việt")

	; --- 1. Thu muc in ---
	$g_idGrp1 = GUICtrlCreateGroup(_T("grp1"), 10, 42, 560, 64)
	$g_idDir = GUICtrlCreateInput(@ScriptDir, 25, 66, 465, 24)
	$g_idBtnBrowse = GUICtrlCreateButton(_T("btnBrowse"), 498, 66, 58, 24)

	; --- 2. Tuy chon in ---
	$g_idGrp2 = GUICtrlCreateGroup(_T("grp2"), 10, 116, 560, 172)
	$g_idLblPages = GUICtrlCreateLabel(_T("lblPages"), 25, 144, 105, 18)
	$g_idPages = GUICtrlCreateInput("1", 132, 141, 90, 24)
	$g_idLblPagesHint = GUICtrlCreateLabel(_T("lblPagesHint"), 230, 144, 330, 18)

	$g_idLblDelay = GUICtrlCreateLabel(_T("lblDelay"), 25, 177, 280, 18)
	$g_idDelay = GUICtrlCreateInput("1000", 310, 174, 90, 24)

	$g_idChkWord = GUICtrlCreateCheckbox(_T("chkWord"), 25, 212, 275, 20)
	GUICtrlSetState($g_idChkWord, $GUI_CHECKED)
	$g_idChkExcel = GUICtrlCreateCheckbox(_T("chkExcel"), 310, 212, 245, 20)
	GUICtrlSetState($g_idChkExcel, $GUI_CHECKED)
	$g_idChkAllSheets = GUICtrlCreateCheckbox(_T("chkAllSheets"), 25, 242, 400, 20)
	GUICtrlSetState($g_idChkAllSheets, $GUI_CHECKED)

	; --- Dieu khien ---
	$g_idBtnStart = GUICtrlCreateButton(_T("btnStart"), 10, 298, 160, 36)
	$g_idBtnStop = GUICtrlCreateButton(_T("btnStop"), 180, 298, 100, 36)
	GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
	$g_idStatus = GUICtrlCreateLabel(_T("lblReady"), 292, 304, 276, 24)

	; --- Nhat ky ---
	$g_idGrpLog = GUICtrlCreateGroup(_T("grpLog"), 10, 344, 560, 235)
	$g_idLog = GUICtrlCreateEdit("", 22, 364, 536, 205, BitOR($ES_READONLY, $ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetFont($g_idLog, 9, 400, 0, "Consolas")

	GUISetState(@SW_SHOW, $g_hGui)
EndFunc

; ======================== ĐỔI NGÔN NGỮ ========================
Func _SwitchLanguage()
	If GUICtrlRead($g_idCmbLang) = "English" Then
		$g_sLang = "en"
	Else
		$g_sLang = "vi"
	EndIf
	_ApplyLanguage()
EndFunc

Func _ApplyLanguage()
	WinSetTitle($g_hGui, "", _T("title"))
	GUICtrlSetData($g_idLblLang, _T("lblLang"))
	GUICtrlSetData($g_idGrp1, _T("grp1"))
	GUICtrlSetData($g_idBtnBrowse, _T("btnBrowse"))
	GUICtrlSetData($g_idGrp2, _T("grp2"))
	GUICtrlSetData($g_idLblPages, _T("lblPages"))
	GUICtrlSetData($g_idLblPagesHint, _T("lblPagesHint"))
	GUICtrlSetData($g_idLblDelay, _T("lblDelay"))
	GUICtrlSetData($g_idChkWord, _T("chkWord"))
	GUICtrlSetData($g_idChkExcel, _T("chkExcel"))
	GUICtrlSetData($g_idChkAllSheets, _T("chkAllSheets"))
	GUICtrlSetData($g_idBtnStart, _T("btnStart"))
	GUICtrlSetData($g_idBtnStop, _T("btnStop"))
	GUICtrlSetData($g_idGrpLog, _T("grpLog"))
	; Chỉ đổi nhãn trạng thái khi đang rảnh (không đè tiến trình in)
	If Not $g_bPrinting Then GUICtrlSetData($g_idStatus, _T("lblReady"))
EndFunc

; ======================== DUYỆT THƯ MỤC ========================
Func _BrowseFolder()
	Local $sDir = FileSelectFolder(_T("dlgSelectFolder"), GUICtrlRead($g_idDir), 0)
	If @error Or $sDir = "" Then Return
	GUICtrlSetData($g_idDir, $sDir)
EndFunc

; ======================== QUY TRÌNH IN ========================
Func _DoPrintBatch()
	; ---- Kiểm tra nhập liệu ----
	Local $sDir = StringStripWS(GUICtrlRead($g_idDir), 3)
	$sDir = StringRegExpReplace($sDir, "[\\/]+$", "")
	If $sDir = "" Or Not FileExists($sDir) Then
		MsgBox(48, _T("errTitle"), _T("msgDirNotExist") & @CRLF & $sDir)
		Return
	EndIf

	Local $bWord = (GUICtrlRead($g_idChkWord) = $GUI_CHECKED)
	Local $bExcel = (GUICtrlRead($g_idChkExcel) = $GUI_CHECKED)
	If Not $bWord And Not $bExcel Then
		MsgBox(48, _T("errTitle"), _T("msgNoType"))
		Return
	EndIf

	Local $sPages = StringStripWS(GUICtrlRead($g_idPages), 3)
	If $sPages <> "" Then
		If Not StringRegExp($sPages, "^(\d+(-\d+)?)(,\d+(-\d+)?)*$", 0) Then
			MsgBox(48, _T("errTitle"), _T("msgBadPages") & @CRLF & _T("msgBadPagesHint"))
			Return
		EndIf
	EndIf

	Local $sDelay = GUICtrlRead($g_idDelay)
	If Not StringRegExp($sDelay, "^\d+$", 0) Then $sDelay = "1000"
	$g_iDelay = Int($sDelay)
	$g_bAllSheets = (GUICtrlRead($g_idChkAllSheets) = $GUI_CHECKED)

	; ---- Liệt kê file Word/Excel (bỏ qua file lock tạm ~$ và thư mục con) ----
	Local $aTarget[1] = [""]
	Local $iCount = 0
	Local $sName = "", $sPath = "", $sExt = "", $bMatch = False
	Local $hSearch = FileFindFirstFile($sDir & "\*.*")
	If $hSearch = -1 Then
		MsgBox(48, _T("errTitle"), _T("msgCantReadDir") & @CRLF & $sDir)
		Return
	EndIf
	While 1
		$sName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If StringLeft($sName, 2) = "~$" Then ContinueLoop ; file tam cua Word/Excel
		$sPath = $sDir & "\" & $sName
		If StringInStr(FileGetAttrib($sPath), "D") Then ContinueLoop ; bo qua thu muc
		$sExt = _GetExt($sName)
		$bMatch = False
		If $bWord And _ExtInList($sExt, $gsWordExt) Then $bMatch = True
		If $bExcel And _ExtInList($sExt, $gsExcelExt) Then $bMatch = True
		If Not $bMatch Then ContinueLoop
		$iCount += 1
		ReDim $aTarget[$iCount]
		$aTarget[$iCount - 1] = $sName
	WEnd
	FileClose($hSearch)

	If $iCount = 0 Then
		MsgBox(64, _T("infoTitle"), _T("msgNoFiles") & @CRLF & $sDir)
		Return
	EndIf
	_SortNatural($aTarget)

	; ---- Chuẩn bị chạy ----
	$g_bPrinting = True
	GUICtrlSetState($g_idBtnStart, $GUI_DISABLE)
	GUICtrlSetState($g_idBtnStop, $GUI_ENABLE)
	GUICtrlSetState($g_idDir, $GUI_DISABLE)
	GUICtrlSetState($g_idPages, $GUI_DISABLE)
	GUICtrlSetState($g_idDelay, $GUI_DISABLE)
	GUICtrlSetState($g_idCmbLang, $GUI_DISABLE)
	$g_bStop = False

	Local $sPagesShow = $sPages
	If $sPagesShow = "" Then $sPagesShow = _T("allPages")
	_Log(_Tf("logStart", "{n}", $iCount))
	_Log(_T("logFolder") & " " & $sDir)
	_Log(_Tf("logPagesDelay", "{p}", $sPagesShow, "{d}", $g_iDelay))

	; ---- Vòng in ----
	Local $iOK = 0, $iFail = 0
	Local $i = 0, $sRet = ""
	For $i = 0 To $iCount - 1
		If $g_bStop Then ExitLoop
		$sName = $aTarget[$i]
		$sExt = _GetExt($sName)
		GUICtrlSetData($g_idStatus, _Tf("logPrinting", "{i}", $i + 1, "{n}", $iCount, "{name}", $sName))

		$sRet = ""
		If $bWord And _ExtInList($sExt, $gsWordExt) Then
			_Log("[" & ($i + 1) & "/" & $iCount & "] " & _T("logWord") & " " & $sName)
			$sRet = _PrintWordFile($sDir & "\" & $sName, $sPages)
		ElseIf $bExcel And _ExtInList($sExt, $gsExcelExt) Then
			_Log("[" & ($i + 1) & "/" & $iCount & "] " & _T("logExcel") & " " & $sName)
			$sRet = _PrintExcelFile($sDir & "\" & $sName, $sPages)
		EndIf

		If $sRet = "" Then
			$iOK += 1
			_Log(_T("logOK"))
		Else
			$iFail += 1
			_Log(_T("logFail") & " " & $sRet)
		EndIf
		_PumpEvents() ; bắt sự kiện nút Dừng / đóng cửa sổ giữa 2 file
	Next

	_QuitApps()

	Local $sMsg = _Tf("logDone", "{ok}", $iOK, "{fail}", $iFail)
	If $g_bStop Then $sMsg &= _T("logStoppedEarly")
	_Log("========== " & StringUpper($sMsg) & " ==========")
	GUICtrlSetData($g_idStatus, $sMsg)

	$g_bPrinting = False
	GUICtrlSetState($g_idBtnStart, $GUI_ENABLE)
	GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
	GUICtrlSetState($g_idDir, $GUI_ENABLE)
	GUICtrlSetState($g_idPages, $GUI_ENABLE)
	GUICtrlSetState($g_idDelay, $GUI_ENABLE)
	GUICtrlSetState($g_idCmbLang, $GUI_ENABLE)
	MsgBox(64, _T("doneTitle"), $sMsg)
EndFunc

; ======================== IN FILE WORD ========================
; Trả về "" nếu thành công, hoặc thông báo lỗi
Func _PrintWordFile($sFile, $sPages)
	$g_bComError = False
	Local $oApp = _GetWordApp()
	If Not IsObj($oApp) Then Return _T("errNoWord")

	Local $oDoc = 0
	$oDoc = $oApp.Documents.Open($sFile, False, True, False) ; ConfirmConversions=False, ReadOnly=True, AddToRecent=False
	If $g_bComError Or Not IsObj($oDoc) Then
		Return _T("errOpenFile")
	EndIf

	Local $sRet = ""
	If $sPages = "" Then
		$oDoc.PrintOut() ; in toàn bộ
	Else
		; Range = 4 (wdPrintRangeOfPages): in theo danh sách trang, VD "1", "2-5", "1,3-5"
		$oDoc.PrintOut(False, False, 4, "", "", "", 0, 1, $sPages)
	EndIf
	If $g_bComError Then $sRet = _T("errPrintWord")

	Sleep($g_iDelay) ; delay sau moi lenh gui toi may in
	$oDoc.Close(False)
	Return $sRet
EndFunc

; ======================== IN FILE EXCEL ========================
; Trả về "" nếu thành công, hoặc thông báo lỗi
Func _PrintExcelFile($sFile, $sPages)
	$g_bComError = False
	Local $oApp = _GetExcelApp()
	If Not IsObj($oApp) Then Return _T("errNoExcel")

	Local $oWb = 0
	$oWb = $oApp.Workbooks.Open($sFile, 0, True) ; UpdateLinks=0, ReadOnly=True
	If $g_bComError Or Not IsObj($oWb) Then
		Return _T("errOpenFile")
	EndIf

	Local $sRet = ""
	Local $aParts = 0, $j = 0, $iFrom = 1, $iTo = 1, $aRT = 0
	; Chọn toàn bộ sheet để khoảng trang tính theo TOÀN BỘ file (giống Word)
	If $g_bAllSheets And $oWb.Sheets.Count > 1 Then
		$oWb.Sheets.Select()
		If $g_bComError Then $sRet = _T("errSelectSheets")
	EndIf

	If $sPages = "" Then
		$oWb.PrintOut() ; in toàn bộ
		If $g_bComError And $sRet = "" Then $sRet = _T("errPrintExcel")
		Sleep($g_iDelay)
	Else
		; Excel không có tham số "Pages" -> tách từng khoảng rồi gửi nhiều lệnh in
		$aParts = StringSplit($sPages, ",", 2)
		For $j = 0 To UBound($aParts) - 1
			If $g_bStop Then ExitLoop
			$iFrom = 1
			$iTo = 1
			If StringInStr($aParts[$j], "-") Then
				$aRT = StringSplit($aParts[$j], "-", 2)
				If UBound($aRT) >= 2 Then
					$iFrom = Int($aRT[0])
					$iTo = Int($aRT[1])
				EndIf
			Else
				$iFrom = Int($aParts[$j])
				$iTo = $iFrom
			EndIf
			If $iFrom < 1 Then $iFrom = 1
			If $iTo < $iFrom Then $iTo = $iFrom

			$g_bComError = False
			$oWb.PrintOut($iFrom, $iTo, 1) ; From, To, Copies
			If $g_bComError And $sRet = "" Then $sRet = _Tf("errPrintExcelPart", "{p}", $aParts[$j])
			Sleep($g_iDelay) ; delay sau moi lenh gui toi may in
			_PumpEvents()
		Next
	EndIf

	$oWb.Close(False)
	Return $sRet
EndFunc

; ======================== ỨNG DỤNG COM (tái sử dụng) ========================
Func _GetWordApp()
	If IsObj($g_oWord) Then Return $g_oWord
	$g_bComError = False
	$g_oWord = ObjCreate("Word.Application")
	If @error Or Not IsObj($g_oWord) Then
		$g_oWord = 0
		Return 0
	EndIf
	$g_oWord.Visible = False
	$g_oWord.DisplayAlerts = 0
	$g_oWord.Options.PrintBackground = False ; in đồng bộ để delay có ý nghĩa
	Return $g_oWord
EndFunc

Func _GetExcelApp()
	If IsObj($g_oExcel) Then Return $g_oExcel
	$g_bComError = False
	$g_oExcel = ObjCreate("Excel.Application")
	If @error Or Not IsObj($g_oExcel) Then
		$g_oExcel = 0
		Return 0
	EndIf
	$g_oExcel.Visible = False
	$g_oExcel.DisplayAlerts = False
	Return $g_oExcel
EndFunc

Func _QuitApps()
	If IsObj($g_oWord) Then
		$g_oWord.Quit(0) ; wdDoNotSaveChanges
		$g_oWord = 0
	EndIf
	If IsObj($g_oExcel) Then
		$g_oExcel.Quit(False) ; SaveChanges=False
		$g_oExcel = 0
	EndIf
EndFunc

Func _OnExit()
	_QuitApps()
EndFunc

; ======================== TIỆN ÍCH ========================
; Bắt event trong lúc đang in (nút Dừng / đóng cửa sổ)
Func _PumpEvents()
	Local $iMsg = GUIGetMsg()
	Switch $iMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $g_idBtnStop
			$g_bStop = True
			_Log(_T("logStopReq"))
	EndSwitch
EndFunc

Func _ComErrorFunc($oError)
	$g_bComError = True
	Local $sDesc = $oError.description
	If $sDesc = "" Then $sDesc = _T("noDesc")
	_Log(_Tf("logComErr", "{n}", $oError.number, "{d}", $sDesc))
EndFunc

Func _Log($sMsg)
	If Not IsHWnd($g_idLog) Then Return
	GUICtrlSetData($g_idLog, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "] " & $sMsg & @CRLF, 1)
EndFunc

Func _GetExt($sName)
	Local $iPos = StringInStr($sName, ".", 0, -1)
	If $iPos = 0 Then Return ""
	Return StringLower(StringMid($sName, $iPos))
EndFunc

Func _ExtInList($sExt, $sList)
	Local $a = StringSplit($sList, "|", 2)
	Local $i = 0
	For $i = 0 To UBound($a) - 1
		If $sExt = $a[$i] Then Return True
	Next
	Return False
EndFunc

; ======================== SẮP XẾP THỨ TỰ TỰ NHIÊN (NATURAL SORT) ========================
; Tách chuỗi thành token số / không phải số: "file10a" -> ["file", "10", "a"]
Func _NatSplit($s)
	Local $aR = StringRegExp($s, "\d+|\D+", 3)
	If @error Then
		Local $aE[1] = [$s]
		Return $aE
	EndIf
	Return $aR
EndFunc

; So sánh 2 chuỗi theo thứ tự tự nhiên (natural order):
;   "file2.docx" < "file10.docx"  (2 < 10, KHÔNG phải 1,10,11,2)
; Trả về: <0 nếu a < b, 0 nếu bằng, >0 nếu a > b
Func _NatCompare($sA, $sB)
	Local $aTokA = _NatSplit($sA)
	Local $aTokB = _NatSplit($sB)
	Local $iMax = UBound($aTokA)
	If UBound($aTokB) > $iMax Then $iMax = UBound($aTokB)
	Local $i = 0, $sTa = "", $sTb = "", $sNa = "", $sNb = "", $iCmp = 0
	For $i = 0 To $iMax - 1
		If $i >= UBound($aTokA) Then Return -1 ; A hết token -> A ngắn hơn -> xếp trước
		If $i >= UBound($aTokB) Then Return 1
		$sTa = $aTokA[$i]
		$sTb = $aTokB[$i]
		If StringRegExp($sTa, "^\d+$") And StringRegExp($sTb, "^\d+$") Then
			; Cả 2 token là số: so sánh THEO SỐ (bỏ số 0 đầu để tránh tràn)
			$sNa = StringRegExpReplace($sTa, "^0+", "")
			$sNb = StringRegExpReplace($sTb, "^0+", "")
			If $sNa = "" Then $sNa = "0"
			If $sNb = "" Then $sNb = "0"
			If StringLen($sNa) <> StringLen($sNb) Then ; số nhiều chữ số lớn hơn
				If StringLen($sNa) < StringLen($sNb) Then Return -1
				Return 1
			EndIf
			$iCmp = StringCompare($sNa, $sNb, 0) ; cùng số chữ số -> so chuỗi = so số
			If $iCmp < 0 Then Return -1
			If $iCmp > 0 Then Return 1
		Else
			; Token không phải số (hoặc 1 số 1 chữ cái): so sánh không phân biệt hoa/thường
			$iCmp = StringCompare($sTa, $sTb, 0)
			If $iCmp < 0 Then Return -1
			If $iCmp > 0 Then Return 1
		EndIf
	Next
	Return 0
EndFunc

; Sắp xếp mảng chuỗi theo thứ tự tự nhiên (merge sort: ổn định, O(n log n))
Func _SortNatural(ByRef $aArr)
	Local $n = UBound($aArr)
	If $n < 2 Then Return
	_MergeSortNat($aArr, 0, $n - 1)
EndFunc

Func _MergeSortNat(ByRef $aArr, $iLo, $iHi)
	If $iLo >= $iHi Then Return
	Local $iMid = Int(($iLo + $iHi) / 2)
	_MergeSortNat($aArr, $iLo, $iMid)
	_MergeSortNat($aArr, $iMid + 1, $iHi)

	; Trộn 2 nửa đã sắp sẵn
	Local $iN = $iHi - $iLo + 1
	Local $aTmp[$iN]
	Local $i = $iLo, $j = $iMid + 1, $k = 0
	While $i <= $iMid And $j <= $iHi
		If _NatCompare($aArr[$i], $aArr[$j]) <= 0 Then
			$aTmp[$k] = $aArr[$i]
			$i += 1
		Else
			$aTmp[$k] = $aArr[$j]
			$j += 1
		EndIf
		$k += 1
	WEnd
	While $i <= $iMid
		$aTmp[$k] = $aArr[$i]
		$i += 1
		$k += 1
	WEnd
	While $j <= $iHi
		$aTmp[$k] = $aArr[$j]
		$j += 1
		$k += 1
	WEnd
	For $k = 0 To $iN - 1
		$aArr[$iLo + $k] = $aTmp[$k]
	Next
EndFunc

; ======================== ĐA NGÔN NGỮ / i18n ========================
; Lấy chuỗi theo ngôn ngữ hiện tại
Func _T($sKey)
	If IsObj($g_oStr) Then
		Local $sK = $g_sLang & "|" & $sKey
		If $g_oStr.Exists($sK) Then Return $g_oStr.Item($sK)
	EndIf
	Return $sKey
EndFunc

; Lấy chuỗi và thay thế placeholder: _Tf("key", "{a}", $v1, "{b}", $v2, ...)
Func _Tf($sKey, $s1 = "", $v1 = "", $s2 = "", $v2 = "", $s3 = "", $v3 = "")
	Local $s = _T($sKey)
	If $s1 <> "" Then $s = StringReplace($s, $s1, $v1)
	If $s2 <> "" Then $s = StringReplace($s, $s2, $v2)
	If $s3 <> "" Then $s = StringReplace($s, $s3, $v3)
	Return $s
EndFunc

Func _AddStr($sLang, $sKey, $sVal)
	If IsObj($g_oStr) Then $g_oStr.Add($sLang & "|" & $sKey, $sVal)
EndFunc

Func _InitStrings()
	$g_oStr = ObjCreate("Scripting.Dictionary")
	If Not IsObj($g_oStr) Then Return

	; ---------------- TIẾNG VIỆT ----------------
	_AddStr("vi", "title", "BatchPrint - In hàng loạt file Word/Excel")
	_AddStr("vi", "lblLang", "Ngôn ngữ:")
	_AddStr("vi", "grp1", "1. Thư mục in (thư mục nằm chung đường dẫn với tool)")
	_AddStr("vi", "btnBrowse", "Duyệt...")
	_AddStr("vi", "grp2", "2. Tùy chọn in")
	_AddStr("vi", "lblPages", "Trang cần in:")
	_AddStr("vi", "lblPagesHint", 'VD: "1"  hoặc  "2-5"  hoặc  "1,3,5-8"  (để trống = in toàn bộ)')
	_AddStr("vi", "lblDelay", "Delay giữa mỗi lệnh gửi tới máy in (ms):")
	_AddStr("vi", "chkWord", "In file Word (.doc, .docx, .docm, .dot, .rtf)")
	_AddStr("vi", "chkExcel", "In file Excel (.xls, .xlsx, .xlsm, .xlsb)")
	_AddStr("vi", "chkAllSheets", "In toàn bộ các sheet của Excel (không chỉ sheet đang xem)")
	_AddStr("vi", "btnStart", "Bắt đầu in")
	_AddStr("vi", "btnStop", "Dừng")
	_AddStr("vi", "lblReady", "Sẵn sàng.")
	_AddStr("vi", "grpLog", "Nhật ký in")
	_AddStr("vi", "logReady", "Sẵn sàng. Đặt file Word/Excel cùng thư mục với tool, hoặc bấm 'Duyệt...' để chọn thư mục khác.")
	_AddStr("vi", "logStopReq", ">> Yêu cầu dừng - sẽ dừng sau lệnh hiện tại...")
	_AddStr("vi", "logStart", "========== BẮT ĐẦU IN {n} FILE ==========")
	_AddStr("vi", "logFolder", "Thư mục:")
	_AddStr("vi", "logPagesDelay", "Trang: {p}  |  Delay: {d} ms/lệnh")
	_AddStr("vi", "allPages", "toàn bộ")
	_AddStr("vi", "logPrinting", "Đang in {i}/{n}: {name}")
	_AddStr("vi", "logWord", "Word :")
	_AddStr("vi", "logExcel", "Excel:")
	_AddStr("vi", "logOK", "      -> OK")
	_AddStr("vi", "logFail", "      -> LỖI:")
	_AddStr("vi", "logDone", "Hoàn thành! Thành công: {ok} | Lỗi: {fail}")
	_AddStr("vi", "logStoppedEarly", " | Dừng sớm")
	_AddStr("vi", "logComErr", "      ! Lỗi COM #{n}: {d}")
	_AddStr("vi", "noDesc", "(không có mô tả)")
	_AddStr("vi", "errTitle", "Lỗi")
	_AddStr("vi", "infoTitle", "Thông báo")
	_AddStr("vi", "doneTitle", "In xong")
	_AddStr("vi", "msgDirNotExist", "Thư mục không tồn tại!")
	_AddStr("vi", "msgNoType", "Hãy chọn ít nhất 1 loại file (Word hoặc Excel)!")
	_AddStr("vi", "msgBadPages", "Định dạng không hợp lệ!")
	_AddStr("vi", "msgBadPagesHint", 'Dùng: "1"  hoặc  "2-5"  hoặc  "1,3,5-8"  (để trống = in toàn bộ)')
	_AddStr("vi", "msgCantReadDir", "Không đọc được thư mục!")
	_AddStr("vi", "msgNoFiles", "Không có file Word/Excel nào trong thư mục.")
	_AddStr("vi", "errNoWord", "Không khởi tạo được Word (đã cài Microsoft Word chưa?)")
	_AddStr("vi", "errNoExcel", "Không khởi tạo được Excel (đã cài Microsoft Excel chưa?)")
	_AddStr("vi", "errOpenFile", "Không mở được file (file hỏng, bị mật khẩu hoặc đang bị khóa?)")
	_AddStr("vi", "errPrintWord", "Lỗi khi gửi lệnh in tới Word")
	_AddStr("vi", "errSelectSheets", "Không chọn được toàn bộ sheet")
	_AddStr("vi", "errPrintExcel", "Lỗi khi gửi lệnh in tới Excel")
	_AddStr("vi", "errPrintExcelPart", "Lỗi khi gửi lệnh in tới Excel (phần đoạn {p})")
	_AddStr("vi", "dlgSelectFolder", "Chọn thư mục chứa file cần in")

	; ---------------- ENGLISH ----------------
	_AddStr("en", "title", "BatchPrint - Batch print Word/Excel files")
	_AddStr("en", "lblLang", "Language:")
	_AddStr("en", "grp1", "1. Print folder (folder in the same path as the tool)")
	_AddStr("en", "btnBrowse", "Browse...")
	_AddStr("en", "grp2", "2. Print options")
	_AddStr("en", "lblPages", "Pages to print:")
	_AddStr("en", "lblPagesHint", 'e.g. "1"  or  "2-5"  or  "1,3,5-8"  (empty = all pages)')
	_AddStr("en", "lblDelay", "Delay between each print command (ms):")
	_AddStr("en", "chkWord", "Print Word files (.doc, .docx, .docm, .dot, .rtf)")
	_AddStr("en", "chkExcel", "Print Excel files (.xls, .xlsx, .xlsm, .xlsb)")
	_AddStr("en", "chkAllSheets", "Print all Excel sheets (not just the active one)")
	_AddStr("en", "btnStart", "Start printing")
	_AddStr("en", "btnStop", "Stop")
	_AddStr("en", "lblReady", "Ready.")
	_AddStr("en", "grpLog", "Print log")
	_AddStr("en", "logReady", "Ready. Put Word/Excel files in the same folder as the tool, or click 'Browse...' to choose another folder.")
	_AddStr("en", "logStopReq", ">> Stop requested - will stop after the current command...")
	_AddStr("en", "logStart", "========== START PRINTING {n} FILES ==========")
	_AddStr("en", "logFolder", "Folder:")
	_AddStr("en", "logPagesDelay", "Pages: {p}  |  Delay: {d} ms/command")
	_AddStr("en", "allPages", "all")
	_AddStr("en", "logPrinting", "Printing {i}/{n}: {name}")
	_AddStr("en", "logWord", "Word :")
	_AddStr("en", "logExcel", "Excel:")
	_AddStr("en", "logOK", "      -> OK")
	_AddStr("en", "logFail", "      -> ERROR:")
	_AddStr("en", "logDone", "Done! Success: {ok} | Failed: {fail}")
	_AddStr("en", "logStoppedEarly", " | Stopped early")
	_AddStr("en", "logComErr", "      ! COM error #{n}: {d}")
	_AddStr("en", "noDesc", "(no description)")
	_AddStr("en", "errTitle", "Error")
	_AddStr("en", "infoTitle", "Notice")
	_AddStr("en", "doneTitle", "Print complete")
	_AddStr("en", "msgDirNotExist", "Folder does not exist!")
	_AddStr("en", "msgNoType", "Please select at least one file type (Word or Excel)!")
	_AddStr("en", "msgBadPages", "Invalid format!")
	_AddStr("en", "msgBadPagesHint", 'Use: "1"  or  "2-5"  or  "1,3,5-8"  (empty = all pages)')
	_AddStr("en", "msgCantReadDir", "Cannot read the folder!")
	_AddStr("en", "msgNoFiles", "No Word/Excel files found in the folder.")
	_AddStr("en", "errNoWord", "Cannot start Word (is Microsoft Word installed?)")
	_AddStr("en", "errNoExcel", "Cannot start Excel (is Microsoft Excel installed?)")
	_AddStr("en", "errOpenFile", "Cannot open file (corrupted, password-protected, or locked?)")
	_AddStr("en", "errPrintWord", "Error sending print command to Word")
	_AddStr("en", "errSelectSheets", "Cannot select all sheets")
	_AddStr("en", "errPrintExcel", "Error sending print command to Excel")
	_AddStr("en", "errPrintExcelPart", "Error sending print command to Excel (range {p})")
	_AddStr("en", "dlgSelectFolder", "Select the folder containing files to print")
EndFunc
