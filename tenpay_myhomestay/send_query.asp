<!-- #include file="asp_md5.asp" -->
<!--#include file="tenpay_config.asp"-->

<html>
<title>������ѯ�ӿڲ���</title>
<meta http-equiv="Cache-Control" content="no-cache"/> 
<body>
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
	cmdno			= "2"					' ������ѯҵ��Ϊ2
	bill_date	= "20080516"	' ��Ҫ��ѯ�Ľ������� (yyyymmdd)		

	bargainor_id = spid					' �̻���
	sp_billno	 = "3771163914"		' �̻����ɵĶ�����(���32λ) Ҫ��ѯ���̻������ţ�	

	' ��Ҫ:
	' ��Ҫ��ѯ�Ľ��׵���(28λ): �̻���(10λ) + ����(8λ) + ��ˮ��(10λ), ���밴�˸�ʽ����, �Ҳ����ظ�
	' ���sp_billno����10λ, ���ȡ���е���ˮ�Ų��ּӵ�transaction_id��(����10λ��0)
	' ���sp_billno����10λ, ����0, �ӵ�transaction_id��
	transaction_id = spid & bill_date & sp_billno 
	
	return_url = domain & tenpay_dir &"/query_result.asp" ' �Ƹ�ͨ�ص�ҳ���ַ,֪ͨ������Ϣ,�Ƽ�ʹ��ip��ַ�ķ�ʽ(�255���ַ�)
	attach		 = "Welcome6"	' �̻�˽������, ����ص�ҳ��ʱԭ������

	' ����MD5ǩ��    
	sign_text = "cmdno=" & cmdno & "&date=" & bill_date & "&bargainor_id=" & bargainor_id &_
        "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno &_
        "&return_url=" & return_url &_
        "&attach=" & attach & "&key=" & sp_key

  md5_sign = UCase(ASP_MD5(sign_text))        ' ת��Ϊ��д
  
%>

<form action="http://portal.tenpay.com/cfbiportal/cgi-bin/cfbiqueryorder.cgi">
<input type=hidden name="cmdno"				value="<%=cmdno%>">
<input type=hidden name="date"			    value="<%=bill_date%>">
<input type=hidden name="bargainor_id"		value="<%=bargainor_id%>">
<input type=hidden name="transaction_id"	value="<%=transaction_id%>">
<input type=hidden name="sp_billno"			value="<%=sp_billno%>">
<input type=hidden name="return_url"		value="<%=return_url%>">
<input type=hidden name="attach"			value="<%=attach%>">
<input type=hidden name="sign"				value="<%=md5_sign%>">

<h4>���͸��Ƹ�ͨ������</h4>
<div>
<textarea name="textarea1" cols="100" rows="6" readonly="true">
cmdno=<%=cmdno%>&date=<%=bill_date%>&bargainor_id=<%=bargainor_id%>&
transaction_id=<%=transaction_id%>&sp_billno=<%=sp_billno%>&
return_url=<%=return_url%>&attach=<%=attach%>&sign=<%=md5_sign%>
</textarea>
</div>
<input type=submit name=submit value=��ѯ����>
</form>
</body>

</html>