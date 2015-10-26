<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim uid,blog_password,rs,udir,uname,fromurl,ufolder
uid=clng(trim(request.QueryString("userid")))
blog_password=trim(request.Form("blog_password"))
fromurl=trim(request("fromurl"))
set rs=oblog.execute("select blog_password,user_dir,username,user_folder from [oblog_user] where userid="&uid)
if not rs.eof then
	udir=rs(1)
	uname=rs(2)
	ufolder=rs(3)
	if rs(0)="" then
	set rs=nothing
	response.Redirect blogdir&udir&"/"&ufolder&"/index."&f_ext
	end if
else
	set rs=nothing
	response.Write("无此用户")
	response.End()
end if

if blog_password<>"" then
	blog_password=md5(blog_password)
	if rs(0)=blog_password then
		set rs=nothing
		Response.Cookies(cookies_name)("blogpw") =blog_password
		if fromurl<>"" then
			response.Redirect(replace(fromurl,"$","&"))
		else
			response.Redirect("pwblog.asp?action=blog&userid="&uid)
		end if
		'response.Write ("<script language=javascript>setCookie('blogpassword', '"&blog_password&"');window.location.replace('"&blogdir&udir&"/"&uid&"/index."&f_ext&"');<script>")
		
	else
		set rs=nothing
		oblog.adderrstr("密码错误，请返回重新输入！")
		oblog.showerr
	end if
end if
set rs=nothing
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>加密blog访问验证页面</title>
<style type="text/css">
<!--
body {
color: #333;
background: #fff;
text-align: left;
margin:0;
font-family: 'Century Gothic', Arial, Helvetica, sans-serif;
font-size: 12px;
line-height: 150%;
}
.content {
width:412px;
height:232px;
margin: 80px 0px 0px 0px; 
background: url("images/passwordbg.png") no-repeat top center;
}
#list form {
padding:110px 0px 0px 60px;
}
#list form #password{
border: 0px #694659 dotted;
width:240px;
height:30px;
font-size:30px; 
color:#099;
background: url("images/none.png");
}
#list form #submit {
margin:28px 0px 0px 165px;
}
-->
</style>
</head>
<body>
<center>
<div class="content">
	
	<div id="list">
	<ul>
	<form method="post">
	<input type="password" size="18" maxlength="20" name="blog_password" id="password"/>
	<input type="image" value="提交" src="images/passwordbt.png" alt="Login" id="submit"/>
	</form>
	</ul>
	</div>	
  </div>
</center>
</body>
</html>