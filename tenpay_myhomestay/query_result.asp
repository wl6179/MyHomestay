<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<meta name="TENCENT_ONELINE_PAYMENT" content="China TENCENT">
<html>
<%
'取返回参数(财付通将会用Post方式请求该页面)
cmdno				= Request("cmdno")
pay_result	= Request("pay_result")
pay_info		= Request("pay_info")
bill_date		= Request("date")
bargainor_id		= Request("bargainor_id")
transaction_id	= Request("transaction_id")
sp_billno		= Request("sp_billno")
total_fee		= Request("total_fee")
fee_type		= Request("fee_type")
attach			= Request("attach")
md5_sign		= Request("sign")

'本地参数
'< !--#include file="tenpay_config.asp"-- >

'返回值定义
Private Const retOK = 0					' 成功					
Private Const invalidSpid = 1			' 商户号错误
Private Const invalidSign = 2			' 签名错误
Private Const tenpayErr	= 3				' 财付通返回支付失败

'验签函数
Function verifyMd5Sign

	origText = "cmdno=" & cmdno & "&pay_result=" & pay_result &_ 
		     "&date=" & bill_date & "&transaction_id=" & transaction_id &_
			   "&sp_billno=" & sp_billno & "&total_fee=" & total_fee &_
			   "&fee_type=" & fee_type & "&attach=" & attach &_
			   "&key=" & sp_key
	
	localSignText = UCase(ASP_MD5(origText)) ' 转换为大写
	verifyMd5Sign = (localSignText = md5_sign)
	
End Function

'返回值
Dim retValue
retValue = retOK

'判断商户号
If bargainor_id <> spid Then
	retValue = invalidSpid
Else 
' 验签
	If verifyMd5Sign <> True Then
		retValue = invalidSign
	Else
' 检查财付通返回值
		If pay_result <> 0 Then
			retValue = tenpayErr
		End If
	End If
End If

' 业务处理,修改数据库,向客户发送通知...
Select Case retValue
	Case retOK			pay_msg = "此单已成功支付!"
	Case invalidSpid	pay_msg = "错误的商户号!"	
	Case invalidSign	pay_msg = "验证MD5签名失败!"
	Case Else			pay_msg = "此单未支付!"
End Select

response.write pay_msg

%>
</html>
