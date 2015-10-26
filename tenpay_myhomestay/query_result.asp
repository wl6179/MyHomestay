<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<meta name="TENCENT_ONELINE_PAYMENT" content="China TENCENT">
<html>
<%
'ȡ���ز���(�Ƹ�ͨ������Post��ʽ�����ҳ��)
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

'���ز���
'< !--#include file="tenpay_config.asp"-- >

'����ֵ����
Private Const retOK = 0					' �ɹ�					
Private Const invalidSpid = 1			' �̻��Ŵ���
Private Const invalidSign = 2			' ǩ������
Private Const tenpayErr	= 3				' �Ƹ�ͨ����֧��ʧ��

'��ǩ����
Function verifyMd5Sign

	origText = "cmdno=" & cmdno & "&pay_result=" & pay_result &_ 
		     "&date=" & bill_date & "&transaction_id=" & transaction_id &_
			   "&sp_billno=" & sp_billno & "&total_fee=" & total_fee &_
			   "&fee_type=" & fee_type & "&attach=" & attach &_
			   "&key=" & sp_key
	
	localSignText = UCase(ASP_MD5(origText)) ' ת��Ϊ��д
	verifyMd5Sign = (localSignText = md5_sign)
	
End Function

'����ֵ
Dim retValue
retValue = retOK

'�ж��̻���
If bargainor_id <> spid Then
	retValue = invalidSpid
Else 
' ��ǩ
	If verifyMd5Sign <> True Then
		retValue = invalidSign
	Else
' ���Ƹ�ͨ����ֵ
		If pay_result <> 0 Then
			retValue = tenpayErr
		End If
	End If
End If

' ҵ����,�޸����ݿ�,��ͻ�����֪ͨ...
Select Case retValue
	Case retOK			pay_msg = "�˵��ѳɹ�֧��!"
	Case invalidSpid	pay_msg = "������̻���!"	
	Case invalidSign	pay_msg = "��֤MD5ǩ��ʧ��!"
	Case Else			pay_msg = "�˵�δ֧��!"
End Select

response.write pay_msg

%>
</html>