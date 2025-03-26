#AutoIt3Wrapper_usex64=n
#AutoIt3Wrapper_Icon=Google-Noto-Emoji-Animals-Nature-22221-cat.ico
#RequireAdmin
#RequireExplicit
#include "_HttpRequest.au3"
#include "Date.au3"
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <ColorConstants.au3>
#include <Array.au3>
#include <WinAPISys.au3>
#include <StaticConstants.au3>
#include <Constants.au3>
#include <IE.au3>
#include <ListViewConstants.au3>
#include <FontConstants.au3>

_IEErrorNotify(False)


# Set HotKey Exit
HotKeySet("{F3}", "_Exit")


# Căn chỉnh lề file in bằng Regedit IE
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "margin_bottom", "REG_SZ", 0)
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "margin_left", "REG_SZ", 0)
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "margin_right", "REG_SZ", 0)
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "margin_top", "REG_SZ", 0)
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "header", "REG_DWORD", 0)
RegWrite("HKCU\Software\Microsoft\Internet Explorer\PageSetup", "footer", "REG_DWORD", 0)





# Khởi tạo các Folder/File Log và IMG
if NOT FileExists(@ScriptDir &'\Log\') Then DirCreate(@ScriptDir &'\Log\')
if NOT FileExists(@ScriptDir &'\img\') Then DirCreate(@ScriptDir &'\img\')
if NOT FileExists(@ScriptDir &'\img\DoanhThu.ico') 	Then	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/DoanhThu.ico'	,@ScriptDir&'\img\DoanhThu.ico')
if NOT FileExists(@ScriptDir &'\img\Count.ico') 	Then	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/Count.ico'		,@ScriptDir&'\img\Count.ico')
if NOT FileExists(@ScriptDir &'\img\Success.ico') 	Then	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/Success.ico'	,@ScriptDir&'\img\Success.ico')
if NOT FileExists(@ScriptDir &'\img\Error.ico') 	Then 	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/Error.ico'		,@ScriptDir&'\img\Error.ico')
if NOT FileExists(@ScriptDir &'\img\Pending.ico')	Then	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/pending.ico'	,@ScriptDir&'\img\Pending.ico')
if NOT FileExists(@ScriptDir &'\log\comeco.png') 	Then 	 InetGet('https://raw.githubusercontent.com/chinhanh09/PRINT-MOMO-BUSINESS/main/icon/comeco.png'	,@ScriptDir&'\log\comeco.png')


# Khai báo biến

Global $idContextmenu = 0, $idSubmenu1 = 999 ,$idSubmenu2 = 999,	$index, $SubItem
Global $OldBalance 	=0 ;Khởi tạo số dư đầu tiên =0
Local $sUser_Name	=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','user','')
Local $sPWD			=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','pwd','')
Local $OldCount		=0
Local $printname	= _GetPrinter_Name()

# Khởi tạo GUI
Local $Form1 		= GUICreate("Tự in bill MOMO ",  540, 600,@DesktopWidth-550,0)
GUISetBkColor(0xFFFFFF)
Local $ButtonStop = GUICtrlCreateButton("Thoát", 180, 120, 90, 25)
Local $ButtonStart = GUICtrlCreateButton("Đăng nhập", 50, 120, 90, 25)
Local $ButtonDangChay = GUICtrlCreateButton("Đang chạy", 50, 120, 90, 25)

Local $printstate= GUICtrlCreateCheckbox('In bill khi có giao dịch mới',80,70,160)
GUICtrlSetFont(-1,9,0,2,'Segoe UI')

Local $speech= GUICtrlCreateCheckbox('Đọc thông báo',80,90,100)
GUICtrlSetTip(-1, "Bỏ chọn sẽ hiện thông báo bằng text trên máy tính")
GUICtrlSetFont(-1,9,0,2,'Segoe UI')
GUICtrlSetState($printstate, $GUI_CHECKED)
Local $date=GUICtrlCreateLabel('',5,278,530,20,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,9,0,2,'Segoe UI')
Local $lbl1				= GUICtrlCreateLabel('Tên đăng nhập',10,15,100)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')
Local $GUI_UserName		= GUICtrlCreateInput($sUser_Name,110,10,150)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')
Local $lbl2				= GUICtrlCreateLabel('Mật khẩu',10,40,100)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $GUI_PWD				= GUICtrlCreateInput($sPWD,110,35,150,Default,$ES_PASSWORD)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $Store=GUICtrlCreateLabel('Vui lòng đăng nhập tài khoản MOMO Business để sử dụng'&@LF&@LF&@LF&'- Chuột phải vào giao dịch trong danh sách để in lại hoặc xem chi tiết',280,10,255,95)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')
GUICtrlSetState($ButtonDangChay,$GUI_HIDE)

Local $group= GUICtrlCreateGroup("Thống kê", 5, 147, 530, 150)
GUICtrlCreateIcon(@ScriptDir&'\img\DoanhThu.ico',-1,15,160,30,30)
GUICtrlCreateLabel('DOANH THU NGÀY',50,165,150)
GUICtrlSetFont(-1,10,800,0,'Segoe UI')

Local $doanhthu= GUICtrlCreateLabel('',20,185,150,25,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')
GUICtrlSetColor(-1, 0xFF0000)
GUICtrlCreateIcon(@ScriptDir&'\img\DoanhThu.ico',-1,280,160,30,30)
GUICtrlCreateLabel('DOANH THU THÁNG',320,165,150)
GUICtrlSetFont(-1,10,800,0,'Segoe UI')

Local $DOANHTHUMOMO= GUICtrlCreateLabel('',300,185,150,25,BitOR($SS_RIGHT, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')
GUICtrlSetColor(-1, $color_blue)



GUICtrlCreateIcon(@ScriptDir &'\img\Count.ico',-1,5,215,30,30)
GUICtrlCreateLabel('SỐ LƯỢNG GIAO DỊCH',-5+45,215,70,75)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $count= GUICtrlCreateLabel('',5,245,100,25,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')


GUICtrlCreateIcon(@ScriptDir &'\img\Success.ico',-1,135,215,25,25)
GUICtrlCreateLabel('THÀNH CÔNG',-5+40+135-5,220,150)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $success= GUICtrlCreateLabel('',135-10,245,150,25,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')


GUICtrlCreateIcon(@ScriptDir &'\img\Error.ico',-1,-5+80+205,215,25,25)
GUICtrlCreateLabel('THẤT BẠI',-10-5+40+80+205,220,150)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $error= GUICtrlCreateLabel('',-5+80+205-20,245,150,25,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')

GUICtrlCreateIcon(@ScriptDir &'\img\Pending.ico',-1,-5+215+200,215,25,25)
GUICtrlCreateLabel('CHỜ XỬ LÝ',-10-5+40+215+200,220,80)
GUICtrlSetFont(-1,9,800,0,'Segoe UI')

Local $pending= GUICtrlCreateLabel('',20-5+215+200-20,245,100,25,BitOR($SS_CENTER, $SS_CENTERIMAGE))
GUICtrlSetFont(-1,20,800,0,'Segoe UI')


Local $note=GUICtrlCreateLabel('F3 để thoát Tool',5,580,530,20,BitOR($SS_CENTER, $SS_CENTERIMAGE))
Local $note1=GUICtrlCreateLabel('',5,580,50,20,BitOR($SS_CENTER, $SS_CENTERIMAGE))

GUICtrlCreateGroup("Danh sách giao dịch", 5, 310, 530,320)

Global	 $Dsgd = GUICtrlCreateListView("STT   |Thời gian	      |Mã Giao dịch	     |Số tiền	     |Kênh thanh toán	 |Trạng thái", 5,325,530,250,BitOR($LVS_SHOWSELALWAYS, $LVS_NOSORTHEADER, $LVS_REPORT, $LVS_SINGLESEL, $LVS_AUTOARRANGE),$LVS_EX_FULLROWSELECT) ;,$LVS_SORTDESCENDING)

GUISetState(@SW_SHOW,$Form1)


    GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND_BUTTON")
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
    Func _WM_COMMAND_BUTTON($hWnd, $Msg, $wParam, $lParam)
        Switch BitAND($wParam, 0x0000FFFF)

		Case $idSubmenu1 ;//////// in lại GD
			$indexSelected  	= ControlListView("", "", $Dsgd, "GetSelected")
			$id 				= ControlListView("", "", $Dsgd, "GetText", $indexSelected, 2)
			$sPassword			= InputBox('In lại giao dịch','Nhập mật khẩu để in lại','','*',Default,Default,Default,Default,10)

			If $sPassword = 'tamnguyen' Then
				$merchant_ID		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','merchantID','')
				$business_username	=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','user','')
				$business_pwd		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','pwd','')
				$token				=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','token','')
				$Time=_DateDiff('s', "1970/01/01 00:00:00", _NowCalc())-25200
				$sURL_Delta='https://business.momo.vn/api/transaction/v2/transactions/PAYMENT-'&$id&'?language=vi'
				$sHeaders			='Merchantid: '&$merchant_ID&'|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$business_username&'_'&$Time

				$srqdelta=_HttpRequest(2,$sURL_Delta,'',$token,'',$sHeaders,'GET')
				$oJsonDeltal=_HttpRequest_ParseJSON($srqdelta)
				If $oJsonDeltal.status=0 Then
					$status				=$oJsonDeltal.data.status
					$createdDate		=StringReplace($oJsonDeltal.data.createdDate,"T"," ")
					$customerName		=$oJsonDeltal.data.customerName
					$typeDescription	=$oJsonDeltal.data.paymentMethodDescription
					$totalAmount		=$oJsonDeltal.data.totalAmount
					$customerPhoneNumber=$oJsonDeltal.data.customerPhoneNumber

					# Khởi tạo Template HTML để in
					$text	= 'Phương thức: <b>QR MOMO Tĩnh</b><br>'& _
					'Nguồn tiền: <b>'&$typeDescription&'</b><br>'& _
					'Mã giao dịch: '&$id&'<br>'& _
					'Thời gian: '&$createdDate&'<br><br><div style="border-bottom: 1px dashed #000; height: 1px;">&nbsp;</div>┣➤  <b>Số tiền: '&_number($totalAmount)&' VNĐ</b><br><div style="border-bottom: 1px dashed #000; height: 1px;">&nbsp;</div><br>Ký tên:<br><br><p style="text-align: center;">Sử dụng ngay khi giao dịch thành công, không quy đổi thành tiền mặt<br><br><b>*** IN LẠI LÚC: '&_NowCalc()&' ***</b>'

					$sHTMLTable='<div class="label-list" style="width: 80mm; min-height: 80mm; display: flex; justify-content: space-around; flex-wrap: wrap;"><div class="fr-view print-label" style="break-inside: avoid; width: calc(100% - 2mm); min-height: calc(78mm);'& _
					'padding-top: 0mm; padding-bottom: 0mm; margin: 0px auto;"><div style="border: none;"><span style="font-family: Tahoma, sans-serif;font-size:13px">'& _
					'<table>	<tr><th></th><th>'& _
					'</th>	</tr></table>'& _
					'<p style="text-align: center;"><span style="font-family: Tahoma, sans-serif;font-size:18px"><strong>**** PHIẾU IN LẠI ****</strong></p>'& _
					'<p></p>'&$Text


					FileWrite(@ScriptDir&'\log\GD IN LAI.txt',_NowCalc()&'	'&$id&'	'&$totalAmount&'	'&$typeDescription&'  ** IN LẠI **'&@CRLF)
					sleep(1000)
					$sFile= @ScriptDir & "\log\Temp.html"
					$hFile= FileOpen($sFile,2+128)
					FileWrite($hFile,$sHTMLTable)
					Local $oIE = _IECreate(@ScriptDir & "\log\Temp.html",0,0)
					_IELoadWait($oIE)
					$oIE.execWB(6,2)
					sleep(500)

				Else
				msgbox(0,0,'Có lỗi xảy ra, vui lòng thử lại')
				$User		= GUICtrlRead($GUI_UserName)
				$sPass		= GUICtrlRead($GUI_PWD)
				$get_StoreInfo	=_Login($User,$sPass)
				sleep(10)

				EndIf
			EndIf



		Case $idSubmenu2 ;//////// Chi tiết giao dịch
				$indexSelected  = ControlListView("", "", $Dsgd, "GetSelected")
				$id 			= ControlListView("", "", $Dsgd, "GetText", $indexSelected, 2)
			;ConsoleWrite($id&@LF)
			$merchant_ID		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','merchantID','')
			$store_ID			=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','storeID','')
			$business_username	=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','user','')
			$business_pwd		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','pwd','')
			$store_name			=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','storeName','')
			$store_name			=BinaryToString(StringToBinary($store_name, 1), 4)
			$store_address		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','storeAdress','')
			$store_address		=BinaryToString(StringToBinary($store_address, 1), 4)
			$token				=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','token','')
			$Time=_DateDiff('s', "1970/01/01 00:00:00", _NowCalc())-25200
				$Delta='https://business.momo.vn/api/transaction/v2/transactions/PAYMENT-'&$id&'?language=vi'
				$srqdelta=_HttpRequest(2,$Delta,'',$token,'','Merchantid: '&$merchant_ID&'|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$business_username&'_'&$Time,'GET')
				$oJsonDeltal=_HttpRequest_ParseJSON($srqdelta)
				;ConsoleWrite($oJsonDeltal.toStr(4))
				If $oJsonDeltal.status=0 Then

					$status				=$oJsonDeltal.data.status
					$createdDate		=StringReplace($oJsonDeltal.data.createdDate,"T"," ")
					$customerName		=$oJsonDeltal.data.customerName
					$Description		=$oJsonDeltal.data.description
					$totalAmount		=$oJsonDeltal.data.totalAmount
					$customerPhoneNumber=$oJsonDeltal.data.customerPhoneNumber
					$kenhthanhtoan		=$oJsonDeltal.data.paymentMethodDescription

					if $kenhthanhtoan = 'Chuyển khoản Ngân Hàng' Then
					MsgBox(0,'Chi tiết giao dịch','Mã giao dịch: '&$id&@LF& _
					'Thời gian: '&$createdDate&@LF& _
					'Trạng thái: '&$status&@LF& _
					'Tên khách hàng: '&$customerName&@LF& _
					'Kênh thanh toán: '&StringUpper($kenhthanhtoan)&@LF& _
					'Số tiền: '&_number($totalAmount)&@LF& _
					'Mô tả: '&$Description)

					EndIf

					if $kenhthanhtoan = 'Ví MoMo' Then


					$phone=$oJsonDeltal.data.customerPhoneNumber
					MsgBox(0,'Chi tiết giao dịch','Mã giao dịch: '&$id&@LF&@LF& _
					'Thời gian: '&$createdDate&@LF&@LF& _
					'Trạng thái: '&$status&@LF&@LF& _
					'Tên khách hàng: '&$customerName&@LF&@LF& _
					'Phone: '&$phone&@LF&@LF& _
					'Kênh thanh toán: '&StringUpper($kenhthanhtoan)&@LF&@LF& _
					'Số tiền: '&_number($totalAmount)&@LF&@LF& _
					'Mô tả: '&$Description)


					EndIf



				Else
				msgbox(0,0,'Có lỗi xảy ra, vui lòng thử lại')
				$User		= GUICtrlRead($GUI_UserName)
				$sPass		= GUICtrlRead($GUI_PWD)
				$get_StoreInfo	=_Login($User,$sPass)
				sleep(10)
				EndIf


			Case $ButtonStop
			$Interrupt = 1
			If GUICtrlRead($ButtonStop)="Thoát" Then Exit -1
			GUICtrlSetData($Store,'Vui lòng đăng nhập tài khoản MOMO Business')
			GUICtrlSetData($ButtonStop,"Thoát")
			GUICtrlSetBkColor($ButtonStop,0xE1E1E1)
			GUICtrlSetColor($ButtonStop,Default)
			GUICtrlSetState($ButtonDangChay,$GUI_HIDE)
			GUICtrlSetState($ButtonStart,$GUI_SHOW)
			GUICtrlSetState($GUI_UserName,$GUI_ENABLE )
			GUICtrlSetState($GUI_PWD,$GUI_ENABLE )
		EndSwitch
        Return 'GUI_RUNDEFMSG'
    EndFunc

    GUIRegisterMsg($WM_SYSCOMMAND, "_WM_COMMAND_CLOSEBUTTON")
    Func _WM_COMMAND_CLOSEBUTTON($hWnd, $Msg, $wParam, $lParam)
        If BitAND($wParam, 0x0000FFFF) = 0xF060 Then Exit
        Return 'GUI_RUNDEFMSG'
    EndFunc




While 1
sleep(10)
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case -3
			_Exit()



		Case $ButtonStart
			$Interrupt=0
			If $Interrupt=0 Then





			GUICtrlSetState($ButtonStart,$GUI_HIDE)
			GUICtrlSetState($ButtonStop,$GUI_SHOW)
			GUICtrlSetData($ButtonStop,'Dừng')
			GUICtrlSetState($ButtonDangChay,$GUI_SHOW)
			GUICtrlSetBkColor($ButtonDangChay,0x2E8B57)
			GUICtrlSetColor($ButtonDangChay,$COLOR_WHITE)
			GUICtrlSetData($ButtonDangChay,'Đang chạy')
			GUICtrlSetBkColor($ButtonStop,$color_red)
			GUICtrlSetColor($ButtonStop,$COLOR_WHITE)

			$User		= GUICtrlRead($GUI_UserName)
			$sPass		= GUICtrlRead($GUI_PWD)
			$get_StoreInfo	=_Login($User,$sPass)
			GUICtrlSetData($Store,$get_StoreInfo)

			while $Interrupt=0
				sleep(10)

				_Main()
				sleep(5000)
			WEnd
		EndIf
	EndSwitch
WEnd

Func _Login($User,$sPass)

;------------------   Login , get Token     -----------------
$Data2Send	='{"username":"'&$User&'","password":"'&$sPass&'"}'
$rq= _HttpRequest(2,'https://business.momo.vn/api/authentication/login?language=vi',$Data2Send)
$oJson = _HttpRequest_ParseJSON($rq)

If $oJson.status= 78 then ;(Nếu status = 0 ==> Thành công , 78 ==> Sai mật khẩu)
	MsgBox(0,'Thông báo',$oJson.message)
	$text =$oJson.message
			GUICtrlSetState($ButtonStart,$GUI_SHOW)
			GUICtrlSetState($ButtonStart,$GUI_SHOW)
			GUICtrlSetState($ButtonStop,$GUI_SHOW)
			GUICtrlSetState($ButtonDangChay,$GUI_HIDE)
			GUICtrlSetData($ButtonStop,"Thoát")
			GUICtrlSetBkColor($ButtonStop,0xE1E1E1)
			GUICtrlSetColor($ButtonStop,Default)
			GUICtrlSetState($GUI_UserName,$GUI_FOCUS )
			$Interrupt=1

Else;-------Login thành công------------
			GUICtrlSetState($GUI_UserName,$GUI_DISABLE )
			GUICtrlSetState($GUI_PWD,$GUI_DISABLE )
			$token = $oJson.data.token
			IniWrite(@ScriptDir &'\ENV.ini','StoreInfo','token',$token)
			IniWrite(@ScriptDir &'\ENV.ini','StoreInfo','user',$User)
			IniWrite(@ScriptDir &'\ENV.ini','StoreInfo','pwd',$sPass)
;----------- Get merchantID ---------------
	$Time					=_DateDiff('s', "1970/01/01 00:00:00", _NowCalc())-25200
	$sHeaders	='Null:null|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$User&'_'&$Time
	$sUrl					='https://business.momo.vn/api/profile/v2/merchants?requestType=LOGIN_MERCHANTS&language=vi'
	$srq					=_HttpRequest(2,$sUrl,'',$token,'',$sHeaders,'GET')
	$oJsonRQ				= _HttpRequest_ParseJSON($srq)
	$merchantID	=$oJsonRQ.data.merchantResponseList.index(0).id

		if 	$oJsonRQ.data.merchantResponseList.index(0).stateCode <> '210' Then
			$text='Tài khoản đã bị khóa'
		Else
			IniWrite(@ScriptDir &'\ENV.ini','StoreInfo','merchantID',$merchantID)
			$text='Đăng nhập thành công'
		EndIf

EndIf
Return $text
EndFunc


Func _Main()
Local $merchant_ID,$store_ID,$business_username,$business_pwd,$store_name,$store_address,$token, $rq

#Region Khai báo các biến để thực hiện Request
$merchant_ID		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','merchantID','')
$business_username	=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','user','')
$business_pwd		=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','pwd','')
$token				=IniRead(@ScriptDir &'\ENV.ini','StoreInfo','token','')
$Time=_DateDiff('s', "1970/01/01 00:00:00", _NowCalc())-25200
$Time=$Time&'000'
$today = _NowCalcDate()
$To_mDay=@MDAY+1
$nextday = _DateAdd("D",1,$today)
$today = StringRegExpReplace($today,"/","-")
$nextday = StringRegExpReplace($nextday,"/","-")
$To_mDay=@MDAY+1
If StringLen($To_mDay)=1 Then $To_mDay='0'&$To_mDay
#EndRegion

	;/////// Lấy tổng doanh thu để so sánh
		$sUrl='https://business.momo.vn/api/transaction/v2/transactions/statistics?pageSize=500&fromDate='&$today&'T00:00:00.00&toDate='&$nextday&'T00:00:00.00&status=ALL&merchantId='&$merchant_ID&'&language=vi'
		_HttpRequest_SetTimeout(5000)
		$srq=_HttpRequest(2,$sUrl,'',$token,'','Merchantid: '&$merchant_ID&'|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$business_username&'_'&$Time,'GET')
		$oJsonRQT= _HttpRequest_ParseJSON($srq)

		if Not @Error Then
			If $oJsonRQT.status=0 Then
				$tongdoanhthu			= 	number($oJsonRQT.data.totalSuccessAmount)
				$tonggiaodich			=	$oJsonRQT.data.totalTrans
				$giaodichthanhcong		=	$oJsonRQT.data.totalSuccessTrans
				$giaodichthatbat		=	$oJsonRQT.data.totalFailTrans
				$giaodichchuahoanthanh	=	$oJsonRQT.data.totalPendingTrans
				$New_Balance			=	Int($tongdoanhthu)
				GUICtrlSetData($doanhthu,_number($tongdoanhthu))
				GUICtrlSetData($count,_number($tonggiaodich))
				GUICtrlSetData($success,_number($giaodichthanhcong))
				GUICtrlSetData($error,_number($giaodichthatbat))
				GUICtrlSetData($pending,_number($giaodichchuahoanthanh))

				;///// Xử lý nếu số dư biến động

				if $OldBalance <> $New_Balance Then
				ConsoleWrite('Có biến động số dư '&$OldBalance&' -->'&$New_Balance&@LF)
				_Get_transaction($merchant_ID,$business_username,$business_pwd,$token, $rq , $OldBalance,$New_Balance,$today,$nextday,$Time)
				$OldBalance=$New_Balance
				EndIf ;if $OldBalance <> $New_Balance Then

				if $OldBalance = 0 Then
					$OldBalance=$New_Balance
					_Set_ListView()
				EndIf ;if $OldBalance <> $New_Balance Then
			Else


			EndIf
		Else
			$User		= GUICtrlRead($GUI_UserName)
			$sPass		= GUICtrlRead($GUI_PWD)
			$get_StoreInfo	=_Login($User,$sPass)
			sleep(1000)

		EndIf

_ReduceMemory()
EndFunc


Func _Get_transaction($merchant_ID,$business_username,$business_pwd,$token, $rq , $OldBalance,$New_Balance,$today,$nextday,$Time)
$querytype			= IniRead(@ScriptDir &'\ENV.ini','StoreInfo','querytype','ALL')
;Xóa process IE
$process1 = ProcessList("iexplore.exe")
if $process1[0][0] >= 1 Then
	ProcessClose("iexplore.exe")
	sleep(100)
endIf

; Tạo 1 vòng lặp để lấy tất cả transactions
while 1
	sleep(10)
	_HttpRequest_SetTimeout(5000)
	$sUrl='https://business.momo.vn/api/transaction/v2/transactions?pageSize=20&fromDate='&$today&'T00:00:00.00&toDate='&$nextday&'T00:00:00.00&status='&$querytype&'&merchantId='&$merchant_ID&'&language=vi'
	$srq=_HttpRequest(2,$sUrl,'','','','Merchantid: '&$merchant_ID&'|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$business_username&'_'&$Time,'GET')
	$oJsonRQ= _HttpRequest_ParseJSON($srq)
	if @error Then ; Sever giới hạn 5s/request nên đặt bẫy lỗi để thực hiện lại request
		ConsoleWrite(">>"&_NowCalc()&" "&'Có lỗi'&@LF)
	Else ; Nếu không có lỗi thì thoát vòng lặp
		ExitLoop
	EndIf

WEnd


if $oJsonRQ.status <> 0 then ; //// Kiểm tra Status nếu <> 0 ==> Token hết hạn ==> Login lại
	$User		= GUICtrlRead($GUI_UserName)
	$sPass		= GUICtrlRead($GUI_PWD)
	$get_StoreInfo	=_Login($User,$sPass)
	sleep(10)

	while 1
		sleep(10)
		_HttpRequest_SetTimeout(5000)
		$sUrl='https://business.momo.vn/api/transaction/v2/transactions?pageSize=20&fromDate='&$today&'T00:00:00.00&toDate='&$nextday&'T00:00:00.00&status='&$querytype&'&merchantId='&$merchant_ID&'&language=vi'
		$srq=_HttpRequest(2,$sUrl,'','','','Merchantid: '&$merchant_ID&'|Authorization: Bearer '&$token&'|X-Api-Request-Id: '&$business_username&'_'&$Time,'GET')
		$oJsonRQ= _HttpRequest_ParseJSON($srq)
		if Not @error Then ExitLoop
	WEnd
EndIf


GUICtrlSetData($date,'Dữ liệu từ 00:00:00 đến 23:59:59 999ms ngày '&@MDAY&'/'&@MON&'/'&@YEAR)
$NewCount		=	$oJsonRQ.data.content.length( )

$logmomo 		=@ScriptDir &'\Log\log'&@MON&@MDAY&'.txt'
$refund			=@ScriptDir &'\Log\Refund'&@MON&@MDAY&'.txt'
$RefundMonth	=Number(IniRead(@ScriptDir &'\Log\Refund.ini',@MON,'HoanTien',0))

;/// Đoạn này để kiểm tra từng GD
For $i= $NewCount-1 to 0 step -1 ; Vòng lặp để kiểm tra từng giao dịch
	sleep(10)
	$id=$oJsonRQ.data.content.index($i).coreTranId
	$storeName=$oJsonRQ.data.content.index($i).storeName
	If Not StringInStr(FileRead($logmomo),$id)  Then ;Đọc file $logmomo xem mã giao dịch đã tồn tại hay chưa

		$createdDate		=StringReplace($oJsonRQ.data.content.index($i).createdDate,"T"," ")
		$createdDate		=StringLeft($createdDate,StringInStr($createdDate,'.')-1)
		$typeDescription	=$oJsonRQ.data.content.index($i).paymentMethodDescription
		$totalAmount		=$oJsonRQ.data.content.index($i).totalAmount
		$Count_Trans 		=Number(IniRead(@ScriptDir &'\log\count.ini','list','Count'&@mon&@MDAY,0))
		If $Count_Trans = 0 Then
			$stt =0
		Else
			$stt = $Count_Trans
		EndIf

		;///////////////// Xử lý giao dịch hoàn tiền
		if $oJsonRQ.data.content.index($i).type = 'REFUND' Then

			If Not StringInStr(FileRead($refund),$id) Then ;Nếu GD Refund chưa ghi nhận
			FileWrite($refund,$createdDate&'	'&$id&'	'&$totalAmount&'	'&$typeDescription&$oJsonRQ.data.content.index($i).type&@CRLF)
			IniWrite(@ScriptDir &'\Log\Refund.ini',@MON,'HoanTien',$RefundMonth+$totalAmount)
			$RefundMonth=	Number(IniRead(@ScriptDir &'\Log\Refund.ini',@MON,'HoanTien',0))
			EndIf

		EndIf


		If $oJsonRQ.data.content.index($i).status = 'SUCCESS' Then ;Nếu trạng thái là Thành công
			$logmomo =@ScriptDir &'\Log\log'&@MON&@MDAY&'.txt'
			local $arr = FileReadToArray ( $logmomo )
			$NewCount = (IsArray($arr)) ? UBound($arr)+1 : 1
			IniWrite(@ScriptDir &'\log\count.ini','list','Count'&@mon&@MDAY,$stt+1)
			if $typeDescription = 'Chuyển khoản Ngân Hàng' Then $typeDescription = 'Chuyển khoản'
			FileWrite($logmomo,$stt+1&'	'&$createdDate&'	'&$id&'	'&_number($totalAmount)&'	'&$typeDescription&'	'&$oJsonRQ.data.content.index($i).statusDescription&@CRLF)
			$OldBalance		=	$New_Balance

			#Region Set template HTML để in

			$sHTMLTable='<div class="label-list" style="width: 80mm; min-height: 80mm; display: flex; justify-content: space-around; flex-wrap: wrap;"><div class="fr-view print-label" style="break-inside: avoid; width: calc(100% - 2mm); min-height: calc(78mm);'& _
			'padding-top: 0mm; padding-bottom: 0mm; margin: 0px auto;"><div style="border: none;"><span style="font-family: Tahoma, sans-serif;font-size:13px">'& _
			'<table><tr><th><img src="comeco.png" alt="Logo" width="60" ></th><th><p style="text-align: center;">'& _
			'<span style="font-family: Tahoma, sans-serif;font-size:15px">'& _
			'<strong>'&$storeName&'</strong>'& _
			'</span></p><p style="text-align: center;">'& _
			'<span style="font-family: Tahoma, sans-serif; font-size:10px">'& _
			'</span>'& _
			'</p></th>	</tr></table>'& _
			'<p style="text-align: center;"><span style="font-family: Tahoma, sans-serif;font-size:16px"><b>THANH TOÁN MOMO TĨNH</b></p>'& _
			'<table><tr><td><span style="font-family: Tahoma, sans-serif;font-size:13px">'& _
			'Nguồn tiền: <b>'&$typeDescription&'</b><br>'& _
			'Mã giao dịch: '&$id&'<br>'& _
			'Thời gian: '&$createdDate&'<br><br>Số tiền: <span style="font-family: Tahoma, sans-serif;font-size:14px"><b>'&_number($totalAmount)&' VNĐ</span></b></td><td>'& _
			'</td><td></td></tr></table>Ký tên:<br><br><br><br>Sử dụng ngay khi giao dịch thành công, không quy đổi thành tiền mặt. <b>STT: '&$NewCount-$i&'</b>'
			#EndRegion

			If GUICtrlRead($speech)=1 Then ; Nếu chọn đọc thông báo - Dùng free API từ SoundOfText/Zalo TTS
				#SoundOfText https://soundoftext.com/docs
				$sstring=StringToBinary('Đã nhận '&$totalAmount&' đồng',4)
				$rq=_HttpRequest(2,'https://api.soundoftext.com/sounds','{"engine":"Google","data":{"text":"'&BinaryToString($sstring)&'","voice":"vi-VN"}}')
				$JsonMP3=_HttpRequest_ParseJSON($rq)
				$patch='https://files.soundoftext.com/'&$JsonMP3.id&'.mp3'
				SoundPlay($patch,0)
				#Region Zalo ai
;~ 	 			#Zalo TTS https://ai.zalo.cloud/docs/api/text-to-audio-converter
;~ 				$sstring=StringToBinary('Nhận được '& $totalAmount &' đồng',4)
;~ 				$sstring=StringToBinary('Hãy đưa phiếu cho nhân viên bán hàng, xin cảm ơn',4)
;~ 				$a=_HttpRequest(2,'https://api.zalo.ai/v1/tts/synthesize','speed=1.1&encode_type=0&input='&BinaryToString($sstring),'','','apikey:mg2Y6F1D6TI1OKHbwXv4SqPDQ5aYpWzj','POST')
;~ 				$oJson =_HttpRequest_ParseJSON($a)
;~ 				ConsoleWrite($oJson.toStr(4))
;~ 				if $oJson.error_message='Successful.' Then
;~ 					$url=$oJson.data.url
;~ 					$wmPlayer = ObjCreate("WMPlayer.OCX.7")
;~ 					If Not IsObj($wmPlayer) Then
;~ 						MsgBox(0, "L?i", "Không tìm th?y Windows Media Player!",10)
;~ 						Exit
;~ 					EndIf
;~ 					$wmPlayer.URL = $url
;~ 					$wmPlayer.controls.play()
;~ 					$a=1
;~ 					While 1
;~ 						Sleep(10)
;~ 						If $wmPlayer.playState = 1 Then
;~ 							ExitLoop ; 1 = dã d?ng
;~ 						Else
;~ 							$a=$a+1
;~ 						EndIf

;~ 						if $a=600 then ExitLoop
;~ 					WEnd
;~ 				EndIf
;~ 				SoundPlay(@ScriptDir&'\speech.mp3',0)
				#EndRegion

			Else ; Nếu không chọn đọc thông báo thì sẽ hiện ở Notifications của Windows
				TrayTip('MOMO','Vừa nhận được '&_number($totalAmount) & ' từ MOMO' &@CRLF&'MGD: '&$id,0)
			EndIf

			If GUICtrlRead($printstate) = 1 Then ; Thực hiện lệnh in file HTML bằng execWB https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752087(v=vs.85)
				$sFile= @ScriptDir & "\log\Temp.html"
				$hFile= FileOpen($sFile,2+128)
				FileWrite($hFile,$sHTMLTable)
				Local $oIE = _IECreate(@ScriptDir & "\log\Temp.html",0,0)
				_IELoadWait($oIE)
				$oIE.execWB(6,2)
				sleep(500)
				FileClose($hFile)
			EndIf


		EndIf
	EndIf
Next
;//// Tạo listview
_Set_ListView()


EndFunc


Func _Set_ListView()
$logmomo =@ScriptDir &'\Log\log'&@MON&@MDAY&'.txt'
local $arr = FileReadToArray ( $logmomo )
Local $getdatasql[1][6] = [[0,0,0,0,0,0]]

for $j = UBound($arr) -1 to 0 step -1
	$sArr=StringSplit($arr[$j],'	',1)
	Local $sFill[1][6] = [[$sArr[1],$sArr[2],$sArr[3],$sArr[4],$sArr[5],$sArr[6]]]
	_ArrayAdd($getdatasql, $sFill)
Next

If IsArray($getdatasql) Then
	_ArrayDelete($getdatasql, 0)
	_GUICtrlListView_DeleteAllItems($Dsgd)
	For $i = 0 To UBound($getdatasql) - 1
		$sItem = ""
		For $j = 0 To UBound($getdatasql, 2) - 1
			$sItem &= $getdatasql[$i][$j] & "|"
		Next
		GUICtrlCreateListViewItem($sItem, $Dsgd)
	Next

	for $i =0 to UBound($getdatasql,$UBOUND_COLUMNS) -1
		_GUICtrlListView_SetColumnWidth($Dsgd, $i, $LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($Dsgd, $i, $LVSCW_AUTOSIZE_USEHEADER); <- used thi instead of GUICtrlSendMsg()
	Next

	for $j= 0 to UBound($getdatasql) -1
		$Name= ControlListView("", "", $Dsgd, "GetText", $j, 4)
		GUICtrlSetBKColor(_GUICtrlListView_GetItemParam($Dsgd, $j), $COLOR_LIGHTGOLDENRODYELLOW)
	Next

Else
	_GUICtrlListView_DeleteAllItems($Dsgd)
EndIf

EndFunc



Func _GetPrinter_Name()
#Hàm lấy list name máy In
$wbemFlagReturnImmediately = "&h10"
$wbemFlagForwardOnly = "&h20"
Global $printname,$sogiaodich
$WMI = ObjGet("winmgmts:\\" & @ComputerName & "\root\CIMV2")
$aItems = $WMI.ExecQuery("SELECT * FROM Win32_Printer", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
For $printer in $aItems
 $printname=$printname &'|'&$printer.Name
Next
Return $printname
EndFunc



Func _Exit()
    If MsgBox(1,'Đã dừng Tools', "Tools đang tạm dừng, bạn có muốn thoát không?" & @CRLF &"Bấm 'OK' để thoát Tools" & @CRLF &"Bấm 'Cancel' để tiếp tục") = 1 Then
		Exit
	EndIf
EndFunc

Func _number($NewText)
$Result=StringRegExpReplace($NewText, '\G([+-]?\d+?)(?=(\d{3})+(\D|$))', '$1.')
Return $Result
EndFunc

Func _ReduceMemory($i_PID = -1)
If $i_PID <> -1 Then
Local $ai_Handle = DllCall("kernel32.dll", 'int', 'OpenProcess', 'int', 0x1f0fff, 'int', False, 'int', $i_PID)
$ai_Return = DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'long', $ai_Handle[0])
DllCall('kernel32.dll', 'int', 'CloseHandle', 'int', $ai_Handle[0])
Else
$ai_Return = DllCall("psapi.dll", 'int', 'EmptyWorkingSet', 'long', -1)
EndIf
Return $ai_Return[0]
EndFunc


Func WM_NOTIFY($hWnd, $MsgID, $wParam, $lParam)

  #forceref $hWnd, $MsgID, $wParam
  Local $tagNMHDR = DllStructCreate("int;int;int", $lParam)
		$tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
		$index = DllStructGetData($tInfo, "Index")
        $SubItem = DllStructGetData($tInfo, "SubItem")
  If @error Then Return $GUI_RUNDEFMSG
  Local $code = DllStructGetData($tagNMHDR, 3)
  If $code = $NM_RCLICK Then
    If _GUICtrlListView_GetSelectedIndices($Dsgd) = "" Then
      GUICtrlDelete($idContextmenu)
      $idContextmenu = 0
    ElseIf Not $idContextmenu Then
      $idContextmenu = GUICtrlCreateContextMenu($Dsgd)
      $idSubmenu1 = GUICtrlCreateMenuItem("In lại ", $idContextmenu)
      $idSubmenu2 = GUICtrlCreateMenuItem("Chi tiết", $idContextmenu)
    EndIf
  EndIf
  Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY
