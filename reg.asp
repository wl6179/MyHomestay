<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
If Request("action")="protocol" Then
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→注册条款<hr />"&oblog.setup(47,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；	
End If
dim action,show_reg,chkregtime
chkregtime=300 '重复注册的时间间隔,单位：秒
if is_ot_user then
	if not IsObject(conn) then link_database
	response.Redirect(ot_regurl)
	response.End()
end if
action=trim(request.Form("action"))
call sysshow()  '首先获得模版~...在往下，就是判断是处理注册 或者 显示出注册表单了；
select case action
	case "chkreg"
	call sub_chkreg()'处理注册;
	case else 
	call sub_showreg()'显示出注册表单了;
end select
show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'将标签替$show_list$换为 该有的内容 show_reg 后 + 底部版权信息代码；
response.Write show'全面输出。
sub sub_showreg()
	dim rs
	dim str_usertype
	
	'show_reg=show_reg&"<div  style='font-size:12px;' CLASS='Register'>"

         

	show_reg=show_reg&"当前位置：<a href='index.asp'>首页</a>→注册新用户<hr />"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"当前系统已关闭注册。"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",0,0)
    str_usertype=str_usertype&"</select><font color=#ff0000> *</font>"
	show_reg=show_reg&"<form name=oblogform method=post action=reg.asp onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>登录用户名：</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=53% colspan=2> <input name=username type=text id=username size=15 maxlength=30><font color=#ff0000 > *</font>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=30% height=25><div align=right >域名：</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td width=53% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>查看用户名、域名是否被占用</a>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr >"& vbcrlf 
    show_reg=show_reg&"<td width=30% height=25><div align=right>输入登录密码：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=password type=password id=password size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>登录密码确认：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=repassword type=password id=repassword size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>密码提示问题：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>密码提示答案：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>地区(省/市)：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&oblog.type_city("","")&"<font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>Email：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>真实姓名：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>Blog类别：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>注册条款：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked>同意　<input class='input_radio' type=radio name=passregtext id=passregtext value='0'>不同意　<a href='reg.asp?action=protocol' target='_blank'>→查看注册条款</a></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf	
	
	if oblog.setup(57,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
    	show_reg=show_reg&"<td width=30% height=25><div align=right>验证码：</div></td>"& vbcrlf
   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
   		show_reg=show_reg&"</tr>"& vbcrlf
	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 ></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' 提交 '>　　　　"& vbcrlf
	show_reg=show_reg&"<input type=reset name=Submit2 value=' 重置 '><input type=hidden name=action value='chkreg'>"& vbcrlf
    show_reg=show_reg&"</div></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"</table>"& vbcrlf
	show_reg=show_reg&"</form>"& vbcrlf
	

	'show_reg=show_reg&"</div>"
    
	end if
	set rs=nothing
end sub

sub sub_chkreg()
	if oblog.ChkPost()=false then
		oblog.adderrstr("系统不允许从外部提交！")
		oblog.showerr
		exit sub
	end if
	chk_regtime()
	if oblog.setup(57,0)=1 then
		if not oblog.codepass then oblog.adderrstr("验证码错误，请刷新后重新输入！"):oblog.showerr
	end if
	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	if oblog.chkiplock() then
		oblog.adderrstr("对不起！你的IP已被锁定,不允许注册！")
		oblog.showerr
		exit sub
	end if
	regusername=oblog.filt_badstr(trim(request("username")))
	if regusername="" then
		response.Redirect "index.asp"
	end if
	regpassword=request("password")
	re_regpassword=request("repassword")
	sex=request("sex")
	email=trim(request("email"))
	question=trim(request("question"))
	answer=trim(request("an"))
	blogname=trim(request("blogname"))
	usertype=clng(request("usertype"))
	'nickname=trim(request("nickname"))
	user_domain=Lcase(trim(request("domain")))
	user_domainroot=trim(request("user_domainroot"))
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("用户名不能为空(不能大于14小于4)！")
	if oblog.chk_regname(regusername) then oblog.adderrstr("用户名系统不允许注册！")
	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("用户名中含有系统不允许的字符！")
	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("用户名不允许全部为数字！")
	if oblog.chkdomain(regusername)=false then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字及下划线！")
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字！")
		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
	end if
	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("密码不能为空(不能大于14小于4)！")
	if re_regpassword="" then oblog.adderrstr("重复密码不能为空！")
	if regpassword<>re_regpassword then oblog.adderrstr("两次输入密码不同！")
	if question="" or oblog.strLength(question)>50 then oblog.adderrstr("找回密码提示问题不能为空(不能大于50)！")
	if answer="" or oblog.strLength(answer)>50 then oblog.adderrstr("找回密码问题答案不能为空(不能大于50)！")
	'if oblog.chk_regname(nickname) then oblog.adderrstr("此昵称系统不允许注册！")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("昵称不能不能大于50字符！")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("真实姓名不能为空(不能大于50字符)！")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("真实姓名中含有系统不允许的字符！")	
	if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("用户名中含有非法字符！")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个昵称存在，请更改昵称！")		
	'end if
	if user_domain<>"" then
		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个域名存在，请更改域名！")
	end if
	if oblog.errstr<>"" then oblog.showerr:exit sub
	if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
	regpassword=md5(regpassword)
	if not IsObject(conn) then link_database
	set rsreg=server.CreateObject("adodb.recordset")
	rsreg.open "select * from [oblog_user] where username='"& oblog.filt_badstr(regusername) &"'",conn,1,3
	if rsreg.eof then
		rsreg.addnew
		rsreg("username")=regusername
		rsreg("password")=regpassword
		if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
			rsreg("user_domain")=user_domain
			rsreg("user_domainroot")=user_domainroot
		end if
		rsreg("question")=question
		rsreg("answer")=md5(answer)
		rsreg("useremail")=email
		rsreg("user_level")=reguserlevel
		rsreg("user_isbest")=0
		rsreg("blogname")=blogname
		rsreg("user_classid")=usertype
		'rsreg("nickname")=nickname
		rsreg("province")=request("province")
		rsreg("city")=request("city")
		rsreg("adddate")=ServerDate(now())
		rsreg("lastloginip")=oblog.userip
		rsreg("lastlogintime")=ServerDate(now())
		rsreg("user_dir")=oblog.setup(30,0)
		rsreg("user_folder")=regusername		
		rsreg.update
		oblog.execute("update oblog_setup set user_count=user_count+1")
		if is_unamedir=0 then
			oblog.execute("update oblog_user set user_folder=userid where username='"&oblog.filt_badstr(regusername)&"'")
		end if
		show_reg="当前位置：<a href='index.asp'>首页</a>→完成注册<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>注册成功，但当前系统设置为需要通过审核，您暂时没有管理权限！</strong></li></ul>"
			show_reg=show_reg&"5秒后自动转到系统首页。"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",5000);"
			show_reg=show_reg&"</script>"
		else
			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
			show_reg=show_reg&"<ul><li><strong>恭喜！您已经注册成功！</strong></li>"
			show_reg=show_reg&"<li><a href='index.asp'>返回首页</a></li>"
			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>点击进入管理后台选择您的喜欢的页面风格(5秒后自动进入管理后台)</a></li></ul>"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
			show_reg=show_reg&"</script>"
		end if
	else
		oblog.adderrstr("系统中已经有这个用户名存在，请更改用户名！")
		oblog.showerr
		exit sub	
	end if
	rsreg.close
	set rsreg=nothing
	'ATFLAG
	'Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=now()	
	Session("chk_regtime")=Now()
end sub

sub chk_regtime()
	dim lasttime
	'ATFLAG
	'lasttime = Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))
	lasttime = Session("chk_regtime")
	If IsDate(lasttime) Then 
		If DateDiff("s",lasttime,Now()) < chkregtime then
			Response.Write("<script language=javascript>alert('"&chkregtime&"秒后才能重复注册。');window.history.back(-1);</script>")
			Response.End
		end if
	end if
end sub
%>
<SCRIPT language="javascript">
<!--
function checkssn(gotoURL) {
   var ssn=oblogform.username.value;
   var domain=oblogform.domain.value;
   var domainroot=oblogform.user_domainroot.value;
	   var open_url = gotoURL + "?username=" + ssn+"&domain="+domain+"&domainroot="+domainroot;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=322,height=200');
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i)!='-')) 
{
return 1;
}
}
return 0;//pass
}
function del_space(s)
{
	for(i=0;i<s.length;++i)
	{
	 if(s.charAt(i)!=" ")
		break;
	}

	for(j=s.length-1;j>=0;--j)
	{
	 if(s.charAt(j)!=" ")
		break;
	}

	return s.substring(i,++j);
}

function VerifySubmit()
{
	 if (document.oblogform.passregtext[1].checked )
     {
        alert("您不同意注册条款，不能注册!");
	return false;
   }	 
   
 
		 
	uid = del_space(document.oblogform.username.value);
     if (uid.length == 0)
     {
        alert("请输入用户名!");
	return false;
     }
	 
	pwd = del_space(document.oblogform.password.value);
     if (pwd.length == 0)
     {
        alert("请输入密码!");
	return false;
     }
	 
	 pwd = del_space(document.oblogform.password.value);
     if (pwd.length < 6)
     {
        alert("密码要大于6个字符！");
	return false;
     }

	pwd2 = del_space(document.oblogform.repassword.value);
     if (pwd2!=pwd)
     {
        alert("两次输入的密码不一致！");
		return false;
     }
	 
	 tishi = del_space(document.oblogform.question.value);
     if (tishi.length == 0)
     {
        alert("请输入密码提示问题");
        return false;
     }
	 
	 tsda = del_space(document.oblogform.an.value);
     if (tsda.length == 0)
     {
        alert("请输入密码提示问题答案！");
        return false;
     }
	 city = del_space(document.all("city").value);
     if (city.length == 0)
     {
        alert("请选择所在城市!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("请输入Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("请输入您的真实姓名!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("请选择您的类别!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>