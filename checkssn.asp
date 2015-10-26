<!--#include file="inc/inc_syssite.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>检测帐号状态</title>
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
if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("用户名不能为空，也不能大于14小于4个字符！")
if oblog.chk_regname(regusername) then oblog.adderrstr("用户名系统不允许注册！")
if oblog.chk_badword(regusername)>0 then oblog.adderrstr("用户名中含有系统不允许的字符！")
if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("用户名不允许全部为数字！")
if oblog.chkdomain(regusername)=false then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字！")
if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
	if user_domain="" or oblog.strLength(user_domain)>20  then oblog.adderrstr("域名不能为空，也不能大于14个字符！")
	if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
	if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
	if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
	if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字及下划线！")
	if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
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
		response.Write("<li>对不起，<strong>"&regusername&"</strong>已存在,请更换！</li>")
	else
		response.Write("<li>恭喜，<strong>"&regusername&"</strong>可以注册！</li>")
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		set rs=oblog.execute("select userid from oblog_user where user_domain='"&user_domain&"' and user_domainroot='"&user_domainroot&"'")
		if not rs.eof then
			response.Write("<li>对不起，<strong>"&user_domain&"."&user_domainroot&"</strong>此域名已存在,请更换！</li>")
		else
			response.Write("<li>恭喜，<strong>"&user_domain&"."&user_domainroot&"</strong>此域名可使用！</li>")
		end if
	end if
End If 
%>
</ul>
<ul style="text-align: center;">
<br>
<br>
<br><input name="Submit" type="submit" value="  关闭窗口  " onClick="javascript:window.close()">
</ul>
</center>
</body>
</html>
