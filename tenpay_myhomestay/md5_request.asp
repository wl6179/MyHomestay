<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<html>
<title>MD5ǩ��֧�� Online</title>
<meta http-equiv="Cache-Control" content="no-cache"/> 
<body>
	<% Response.Charset="GB2312" %>
<%
	' ��ȡ���������ڣ���ʽYYYYMMDD
	Function CFTGetServerDate 
		Dim strTmp, iYear,iMonth,iDate 
		iYear = Year(Date) 
		iMonth = Month(Date) 
		iDate = Day(Date) 

		strTmp = CStr(iYear)
		If iMonth < 10 Then 
			strTmp = strTmp & "0" & Cstr(iMonth)
		Else 
			strTmp = strTmp & Cstr(iMonth)
		End If 
		If iDate < 10 Then 
			strTmp = strTmp & "0" & Cstr(iDate) 
		Else 
			strTmp = strTmp & Cstr(iDate) 
		End If 
		CFTGetServerDate = strTmp 
	End Function

	Dim request_text
	Dim md5_sign

	'< !--#include file="tenpay_config.asp"-- >
	
	' �������������
	cmdno		= "1"				' �Ƹ�֧ͨ��Ϊ"1" (��ǰֻ֧�� cmdno=1)	
	bill_date	= CFTGetServerDate	' �������� (yyyymmdd)	
	bank_type	= "0"				' ��������:	0		�Ƹ�ͨ
									'			1001	��������   
									'			1002	�й���������  
									'			1003	�й���������  
									'			1004	�Ϻ��ֶ���չ����   
									'			1005	�й�ũҵ����  
									'			1006	�й���������  
									'			1008	���ڷ�չ����   
									'			1009	��ҵ����   

	desc		= Request.Form("product_desc")		' ��Ʒ����
	purchaser_id = ""				' �û��Ƹ�ͨ�ʺţ����û�п����ÿ�
	bargainor_id = spid				' �̻���
	
	'--------------------Begin
	Dim strBillNo
	strBillNo = CFTGetServerDate_BillNo	'WL�Լ���˾�����̻������ź���
	'--------------------End
	
	'//sp_billno	 = strBillNo		' �̻����ɵĶ�����(���32λ)	������10λ
	sp_billno	 = Request.Form("str_BillNo")		' �̻����ɵĶ�����(���32λ)	������10λ	

	' ��Ҫ:
	' ���׵���(28λ): �̻���(10λ) + ����(8λ) + ��ˮ��(10λ), ���밴�˸�ʽ����, �Ҳ����ظ�
	' ���sp_billno����10λ, ���ȡ���е���ˮ�Ų��ּӵ�transaction_id��(����10λ��0)
	' ���sp_billno����10λ, ����0, �ӵ�transaction_id��
	transaction_id = spid & bill_date & sp_billno 

	total_fee	 = Request.Form("total_money")					' �ܽ��, ��Ϊ��λ
	fee_type	 = "1"					' ��������: 1 �C RMB(�����) 2 - USD(��Ԫ) 3 - HKD(�۱�)
	return_url	 = domain & tenpay_dir &"/notify_handler.asp"	 ' �Ƹ�ͨ�ص�ҳ���ַ, (�255���ַ�)
	attach		 = "WelcomeToHouseShop"	' �̻�˽������, ����ص�ҳ��ʱԭ������

	' ����MD5ǩ��    
	sign_text = "cmdno=" & cmdno & "&date=" & bill_date & "&bargainor_id=" & bargainor_id &_
        "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno &_
        "&total_fee=" & total_fee & "&fee_type=" & fee_type & "&return_url=" & return_url &_
        "&attach=" & attach & "&key=" & sp_key

	md5_sign = UCase(ASP_MD5(sign_text))        ' ת��Ϊ��д 
	 
%>

<form name="form1" id="form1" action="https://www.tenpay.com/cgi-bin/v1.0/pay_gate.cgi">
	<input type=hidden name="cmdno"				value="<%=cmdno%>">
	<input type=hidden name="date"			    value="<%=bill_date%>">
	<input type=hidden name="bank_type"			value="<%=bank_type%>">
	<input type=hidden name="desc"				value="<%=desc%>">
	<input type=hidden name="purchaser_id"		value="<%=purchaser_id%>">
	<input type=hidden name="bargainor_id"		value="<%=bargainor_id%>">
	<input type=hidden name="transaction_id"	value="<%=transaction_id%>">
	<input type=hidden name="sp_billno"			value="<%=sp_billno%>">
	<input type=hidden name="total_fee"			value="<%=total_fee%>">
	<input type=hidden name="fee_type"			value="<%=fee_type%>">
	<input type=hidden name="return_url"		value="<%=return_url%>">
	<input type=hidden name="attach"			value="<%=attach%>">
	<input type=hidden name="sign"				value="<%=md5_sign%>">
	
	<!--<h4>���� ����֧�� Online</h4>
	<div>
	<textarea name="textarea1" cols="100" rows="6" readonly="true">
	cmdno=<%=cmdno%>&date=<%=bill_date%>&bank_type=<%=bank_type%>&
	desc=<%=desc%>&purchaser_id=<%=purchaser_id%>&
	bargainor_id=<%=bargainor_id%>&transaction_id=<%=transaction_id%>&
	sp_billno=<%=sp_billno%>&total_fee=<%=total_fee%>&
	fee_type=<%=fee_type%>&return_url=<%=return_url%>&
	attach=<%=attach%>&sign=<%=md5_sign%>
	</textarea>
	</div>-->
	<!--<input type=submit name=submit value=����֧��Online />-->&nbsp;&nbsp;&nbsp;<!--<input type=button name=close value=�ر� onClick="window.close();" />-->
	
	
</form>
	<script language=javascript type="text/javascript">
		document.form1.submit();
	</script>
</body>
</html>
