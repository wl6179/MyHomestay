<!--#include file="inc/inc_syssite.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ʺ�״̬</title>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
	margin-left: 4px;
	margin-top: 20px;
	margin-right: 4px;
	margin-bottom: 4px;
	font-size: 12px;
}
-->
</style>
</head>
<body>
<ul>
<%
dim regusername,user_domain,user_domainroot
regusername=oblog.filt_badstr(trim(request("username")))
user_domain=oblog.filt_badstr(trim(request("domain")))
user_domainroot=oblog.filt_badstr(trim(request("domainroot")))
if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("�û�������Ϊ�գ�Ҳ���ܴ���14С��4���ַ���")
if oblog.chk_regname(regusername) then oblog.adderrstr("�û���ϵͳ������ע�ᣡ")
if oblog.chk_badword(regusername)>0 then oblog.adderrstr("�û����к���ϵͳ��������ַ���")
if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("�û���������ȫ��Ϊ���֣�")
if oblog.chkdomain(regusername)=false then oblog.adderrstr("�û������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
	if user_domain="" or oblog.strLength(user_domain)>20  then oblog.adderrstr("��������Ϊ�գ�Ҳ���ܴ���14���ַ���")
	if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("��������С��4���ַ���")
	if oblog.chk_regname(user_domain) then oblog.adderrstr("������ϵͳ������ע�ᣡ")
	if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
	if oblog.chkdomain(user_domain)=false then oblog.adderrstr("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����ּ��»��ߣ�")
	if user_domainroot="" then oblog.adderrstr("����������Ϊ�գ�")
end if
If oblog.errstr<>"" Then
	dim errmsg,errmsg1,i
	errmsg=Split(oblog.errstr,"_")
	For i=0 to UBound(errmsg)
		If i=0 Then
			errmsg1=errmsg1&"<li>"&errmsg(i)
		Else
			errmsg1=errmsg1&"<br><li>"&errmsg(i)
		End If
	Next
	response.Write(errmsg1)
else
	dim rs
	set rs=oblog.execute("select userid from oblog_user where username='"&regusername&"'")
	if not rs.eof then
		response.Write("<li>�Բ���<strong>"&regusername&"</strong>�Ѵ���,�������</li>")
	else
		response.Write("<li>��ϲ��<strong>"&regusername&"</strong>����ע�ᣡ</li>")
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		set rs=oblog.execute("select userid from oblog_user where user_domain='"&user_domain&"' and user_domainroot='"&user_domainroot&"'")
		if not rs.eof then
			response.Write("<li>�Բ���<strong>"&user_domain&"."&user_domainroot&"</strong>�������Ѵ���,�������</li>")
		else
			response.Write("<li>��ϲ��<strong>"&user_domain&"."&user_domainroot&"</strong>��������ʹ�ã�</li>")
		end if
	end if
End If 
%>
</ul>
<ul style="text-align: center;">
<br>
<br>
<br><input name="Submit" type="submit" value="  �رմ���  " onClick="javascript:window.close()">
</ul>
</center>
</body>
</html>
