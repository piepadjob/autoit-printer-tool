#cs ----------------------------------------------------------------------------
 Script  : BatchPrint.au3
 Mô tả   : In hàng loạt file Word/Excel nằm trong một thư mục
           (thư mục nằm chung đường dẫn với tool - mặc định là thư mục chứa file .exe)
 Tính năng:
   + In toàn bộ file Word/Excel trong thư mục được chọn
   + Chọn khoảng trang cần in (mặc định: trang 1)
   + Delay (ms) giữa mỗi lệnh gửi tới máy in
   + Tùy chọn in toàn bộ các sheet của Excel
   + Nhật ký in, nút Dừng, bỏ qua file lock tạm (~$...)
 Biên dịch: SciTE for AutoIt -> Tools -> Compile (hoặc Aut2exe.exe) -> BatchPrint.exe
#ce ----------------------------------------------------------------------------

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <ButtonConstants.au3>
#include <Array.au3>

; ======================== HẰNG SỐ ========================
Global Const $gsWordExt = ".doc|.docx|.docm|.dot|.dotx|.rtf"
Global Const $gsExcelExt = ".xls|.xlsx|.xlsm|.xlsb|.xlt|.xltx"

; ======================== BIẾN TOÀN CỤC ========================
Global $g_hGui
Global $g_idDir, $g_idPages, $g_idDelay, $g_idChkWord, $g_idChkExcel, $g_idChkAllSheets
Global $g_idBtnBrowse, $g_idBtnStart, $g_idBtnStop, $g_idLog, $g_idStatus
Global $g_bStop = False
Global $g_bComError = False
Global $g_iDelay = 1000
Global $g_bAllSheets = True
Global $g_oWord = 0
Global $g_oExcel = 0

; Bắt lỗi COM toàn cục (không để lỗi COM làm treo tool)
Global $g_oComErr = ObjEvent("AutoIt.Error", "_ComErrorFunc")
OnAutoItExitRegister("_OnExit")

_Main()

; ======================== VÒNG LẶP CHÍNH ========================
Func _Main()
	_BuildGUI()
	_Log("San sang. Dat file Word/Excel cung thu muc voi tool, hoac bam 'Duyet...' de chon thu muc khac.")
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
				_Log(">> Yeu cau dung - se dung sau lenh hien tai...")
		EndSwitch
	WEnd
EndFunc

; ======================== GIAO DIỆN ========================
Func _BuildGUI()
	$g_hGui = GUICreate("BatchPrint - In hang loat file Word/Excel", 580, 560, -1, -1)

	; --- 1. Thu muc in ---
	GUICtrlCreateGroup("1. Thu muc in (thu muc nam chung duong dan voi tool)", 10, 10, 560, 64)
	$g_idDir = GUICtrlCreateInput(@ScriptDir, 25, 34, 465, 24)
	$g_idBtnBrowse = GUICtrlCreateButton("Duyet...", 498, 34, 58, 24)

	; --- 2. Tuy chon in ---
	GUICtrlCreateGroup("2. Tuy chon in", 10, 84, 560, 172)
	GUICtrlCreateLabel("Trang can in:", 25, 112, 105, 18)
	$g_idPages = GUICtrlCreateInput("1", 132, 109, 90, 24)
	GUICtrlCreateLabel('VD: "1"  hoac  "2-5"  hoac  "1,3,5-8"  (de trong = in toan bo)', 230, 112, 330, 18)

	GUICtrlCreateLabel("Delay giua moi lenh gui toi may in (ms):", 25, 145, 280, 18)
	$g_idDelay = GUICtrlCreateInput("1000", 310, 142, 90, 24)

	$g_idChkWord = GUICtrlCreateCheckbox("In file Word (.doc, .docx, .docm, .dot, .rtf)", 25, 180, 275, 20)
	GUICtrlSetState($g_idChkWord, $GUI_CHECKED)
	$g_idChkExcel = GUICtrlCreateCheckbox("In file Excel (.xls, .xlsx, .xlsm, .xlsb)", 310, 180, 245, 20)
	GUICtrlSetState($g_idChkExcel, $GUI_CHECKED)
	$g_idChkAllSheets = GUICtrlCreateCheckbox("In toan bo cac sheet cua Excel (khong chi sheet dang xem)", 25, 210, 400, 20)
	GUICtrlSetState($g_idChkAllSheets, $GUI_CHECKED)

	; --- Dieu khien ---
	$g_idBtnStart = GUICtrlCreateButton("Bat dau in", 10, 266, 160, 36)
	$g_idBtnStop = GUICtrlCreateButton("Dung", 180, 266, 100, 36)
	GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
	$g_idStatus = GUICtrlCreateLabel("San sang.", 292, 272, 276, 24)

	; --- Nhat ky ---
	GUICtrlCreateGroup("Nhat ky in", 10, 312, 560, 235)
	$g_idLog = GUICtrlCreateEdit("", 22, 332, 536, 205, BitOR($ES_READONLY, $ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL))
	GUICtrlSetFont($g_idLog, 9, 400, 0, "Consolas")

	GUISetState(@SW_SHOW, $g_hGui)
EndFunc

; ======================== DUYỆT THƯ MỤC ========================
Func _BrowseFolder()
	Local $sDir = FileSelectFolder("Chon thu muc chua file can in", GUICtrlRead($g_idDir), 0)
	If @error Or $sDir = "" Then Return
	GUICtrlSetData($g_idDir, $sDir)
EndFunc

; ======================== QUY TRÌNH IN ========================
Func _DoPrintBatch()
	; ---- Kiểm tra nhập liệu ----
	Local $sDir = StringStripWS(GUICtrlRead($g_idDir), 3)
	$sDir = StringRegExpReplace($sDir, "[\\/]+$", "")
	If $sDir = "" Or Not FileExists($sDir) Then
		MsgBox(48, "Loi", "Thu muc khong ton tai!" & @CRLF & $sDir)
		Return
	EndIf

	Local $bWord = (GUICtrlRead($g_idChkWord) = $GUI_CHECKED)
	Local $bExcel = (GUICtrlRead($g_idChkExcel) = $GUI_CHECKED)
	If Not $bWord And Not $bExcel Then
		MsgBox(48, "Loi", "Hay chon it nhat 1 loai file (Word hoac Excel)!")
		Return
	EndIf

	Local $sPages = StringStripWS(GUICtrlRead($g_idPages), 3)
	If $sPages <> "" Then
		If Not StringRegExp($sPages, "^(\d+(-\d+)?)(,\d+(-\d+)?)*$", 0) Then
			MsgBox(48, "Loi", 'Dinh dang khong hop le!' & @CRLF & _
				'Dung: "1"  hoac  "2-5"  hoac  "1,3,5-8"  (de trong = in toan bo)')
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
		MsgBox(48, "Loi", "Khong doc duoc thu muc!" & @CRLF & $sDir)
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
		MsgBox(64, "Thong bao", "Khong co file Word/Excel nao trong thu muc." & @CRLF & $sDir)
		Return
	EndIf
	_ArraySort($aTarget)

	; ---- Chuẩn bị chạy ----
	GUICtrlSetState($g_idBtnStart, $GUI_DISABLE)
	GUICtrlSetState($g_idBtnStop, $GUI_ENABLE)
	GUICtrlSetState($g_idDir, $GUI_DISABLE)
	GUICtrlSetState($g_idPages, $GUI_DISABLE)
	GUICtrlSetState($g_idDelay, $GUI_DISABLE)
	$g_bStop = False

	Local $sPagesShow = $sPages
	If $sPagesShow = "" Then $sPagesShow = "toan bo"
	_Log("========== BAT DAU IN " & $iCount & " FILE ==========")
	_Log("Thu muc: " & $sDir)
	_Log("Trang: " & $sPagesShow & "  |  Delay: " & $g_iDelay & " ms/file")

	; ---- Vòng in ----
	Local $iOK = 0, $iFail = 0
	Local $i = 0, $sRet = ""
	For $i = 0 To $iCount - 1
		If $g_bStop Then ExitLoop
		$sName = $aTarget[$i]
		$sExt = _GetExt($sName)
		GUICtrlSetData($g_idStatus, "Dang in " & ($i + 1) & "/" & $iCount & ": " & $sName)

		$sRet = ""
		If $bWord And _ExtInList($sExt, $gsWordExt) Then
			_Log("[" & ($i + 1) & "/" & $iCount & "] Word : " & $sName)
			$sRet = _PrintWordFile($sDir & "\" & $sName, $sPages)
		ElseIf $bExcel And _ExtInList($sExt, $gsExcelExt) Then
			_Log("[" & ($i + 1) & "/" & $iCount & "] Excel: " & $sName)
			$sRet = _PrintExcelFile($sDir & "\" & $sName, $sPages)
		EndIf

		If $sRet = "" Then
			$iOK += 1
			_Log("      -> OK")
		Else
			$iFail += 1
			_Log("      -> LOI: " & $sRet)
		EndIf
		_PumpEvents() ; bắt sự kiện nút Dừng / đóng cửa sổ giữa 2 file
	Next

	_QuitApps()

	Local $sMsg = "Hoan thanh! Thanh cong: " & $iOK & " | Loi: " & $iFail
	If $g_bStop Then $sMsg &= " | Dung som"
	_Log("========== " & StringUpper($sMsg) & " ==========")
	GUICtrlSetData($g_idStatus, $sMsg)

	GUICtrlSetState($g_idBtnStart, $GUI_ENABLE)
	GUICtrlSetState($g_idBtnStop, $GUI_DISABLE)
	GUICtrlSetState($g_idDir, $GUI_ENABLE)
	GUICtrlSetState($g_idPages, $GUI_ENABLE)
	GUICtrlSetState($g_idDelay, $GUI_ENABLE)
	MsgBox(64, "In xong", $sMsg)
EndFunc

; ======================== IN FILE WORD ========================
; Trả về "" nếu thành công, hoặc thông báo lỗi
Func _PrintWordFile($sFile, $sPages)
	$g_bComError = False
	Local $oApp = _GetWordApp()
	If Not IsObj($oApp) Then Return "Khong khoi tao duoc Word (da cai Microsoft Word chua?)"

	Local $oDoc = 0
	$oDoc = $oApp.Documents.Open($sFile, False, True, False) ; ConfirmConversions=False, ReadOnly=True, AddToRecent=False
	If $g_bComError Or Not IsObj($oDoc) Then
		Return "Khong mo duoc file (file hong, bi mat khau hoac dang bi khoa?)"
	EndIf

	Local $sRet = ""
	If $sPages = "" Then
		$oDoc.PrintOut() ; in toàn bộ
	Else
		; Range = 4 (wdPrintRangeOfPages): in theo danh sách trang,VD "1", "2-5", "1,3-5"
		$oDoc.PrintOut(False, False, 4, "", "", "", 0, 1, $sPages)
	EndIf
	If $g_bComError Then $sRet = "Loi khi gui lenh in toi Word"

	Sleep($g_iDelay) ; delay sau moi lenh gui toi may in
	$oDoc.Close(False)
	Return $sRet
EndFunc

; ======================== IN FILE EXCEL ========================
; Trả về "" nếu thành công, hoặc thông báo lỗi
Func _PrintExcelFile($sFile, $sPages)
	$g_bComError = False
	Local $oApp = _GetExcelApp()
	If Not IsObj($oApp) Then Return "Khong khoi tao duoc Excel (da cai Microsoft Excel chua?)"

	Local $oWb = 0
	$oWb = $oApp.Workbooks.Open($sFile, 0, True) ; UpdateLinks=0, ReadOnly=True
	If $g_bComError Or Not IsObj($oWb) Then
		Return "Khong mo duoc file (file hong, bi mat khau hoac dang bi khoa?)"
	EndIf

	Local $sRet = ""
	Local $aParts = 0, $j = 0, $iFrom = 1, $iTo = 1, $aRT = 0
	; Chọn toàn bộ sheet để khoảng trang tính theo TOÀN BỘ file (giống Word)
	If $g_bAllSheets And $oWb.Sheets.Count > 1 Then
		$oWb.Sheets.Select()
		If $g_bComError Then $sRet = "Khong chon duoc toan bo sheet"
	EndIf

	If $sPages = "" Then
		$oWb.PrintOut() ; in toàn bộ
		If $g_bComError And $sRet = "" Then $sRet = "Loi khi gui lenh in toi Excel"
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
			If $g_bComError And $sRet = "" Then $sRet = "Loi khi gui lenh in toi Excel (phan doan " & $aParts[$j] & ")"
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
			_Log(">> Yeu cau dung - se dung sau lenh hien tai...")
	EndSwitch
EndFunc

Func _ComErrorFunc($oError)
	$g_bComError = True
	Local $sDesc = $oError.description
	If $sDesc = "" Then $sDesc = "(khong co mo ta)"
	_Log("      ! Loi COM #" & $oError.number & ": " & $sDesc)
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
