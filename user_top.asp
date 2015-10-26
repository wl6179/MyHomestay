<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <meta name="Robots" content="noindex,nofollow" />
	<meta name="author" content="王亮, wangliang6179@163.com" />
	<meta name="copyright" content="www.myhomestay.com.cn,我的住家网" />

<%
Dim t,tName,nowurl,cssfile,actiontype,url
url=request.QueryString("url")
t=Request("t")
actiontype=request("actiontype")
nowurl=lcase(Request.ServerVariables("PATH_INFO"))

select case actiontype
case "left"
	call uleft():response.End()
case "center"
	call ucenter():response.End()
end select

If t="1" Then
	tName="图片"
Else
	t="0"
	tName="文章"
End If

dim oblog,uurl
set oblog=new class_sys
oblog.start
'退出系统
If request("t")="logout" Then
	If cookies_domain <> "" Then
        response.Cookies(cookies_name).domain = cookies_domain
    End If
	Response.Cookies(cookies_name)("username")=oblog.CodeCookie("")
	Response.Cookies(cookies_name)("password")=oblog.CodeCookie("")
	Response.Cookies(cookies_name)("userurl")=oblog.CodeCookie("")
	if request("re")="1" then
		'response.Redirect(oblog.comeurl)
		response.Write("<script language=JavaScript>top.location='"&oblog.comeurl&"';</script>")
	else'退出后回到主页；
		response.Write("<script language=JavaScript>top.location='index.asp';</script>")
	end if
	Set oBlog=Nothing
	response.End()
End If
if not oblog.checkuserlogined() then
	'''response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
	response.Redirect("index.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
else
	if oblog.logined_ulevel=6 then
		oblog.adderrstr("您未通过MyHomestay管理员审核，不能进入后台")
		oblog.showerr
	end if
end if
if request("t")="change0" then change(0)
if request("t")="change1" then change(1)
'判断分栏
if oblog.logined_uframe=1 then
	cssfile="back-ctrl.css"
	if actiontype="" and (instr(nowurl,"HomestayBackctrl-index.asp") and actiontype<>"main") then 
		call frame():response.End()
	end if
	'if instr(nowurl,"HomestayBackctrl-index.asp")=0 and instr(nowurl,"user_top.asp")=0 then
	'	response.Redirect("HomestayBackctrl-index.asp?url="&nowurl)
	'end if
else
	if url<>"" then response.Redirect(url)
	cssfile="back-ctrl.css"
end if
if oblog.setup(12,0)=1 then
	uurl="<a href=""http://"&oblog.logined_udomain&""" target=""_blank"">"&oblog.logined_udomain&"</a>"
else
	uurl="<a href="""&oblog.logined_udir&"/"&oblog.logined_ufolder&"/index."&f_ext&""" target=""_blank"">生成栏目页</a>"
end if
if true_domain=1 and oblog.logined_ucustomdomain<>"" then
	uurl="<a href=""http://"&oblog.logined_ucustomdomain&""" target=""_blank"">"&oblog.logined_ucustomdomain&"</a>"
end if

call utop()


sub utop()%>

	<title>
      <%=oblog.logined_uname%> - 我的住家网 MyHomestay 用户后台管理
    </title>
	<script src="/inc/function.js" type="text/javascript"></script>
    <link rel="stylesheet" media="all" rev="alternate" type="text/css" href="css/<%=cssfile%>" />

	<link rel="stylesheet" type="text/css" href="inc/divCN-EN.css" />

  </head>
  <body>
<%end sub



sub uleft()%>

<%end sub

sub ucenter()%>
<link title="Oblog_Green" href="OblogStyle/OblogStyleAdminMenu.css" rel="stylesheet" media="all" type="text/css" rev="alternate" />
<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("OblogStyle/OblogStyleAdminImages/nav_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,8,*";
		displayBar=false;
		obj.src="OblogStyle/OblogStyleAdminImages/center_open.gif";
		obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.cols="156,8,*";
		displayBar=true;
		obj.src="OblogStyle/OblogStyleAdminImages/center_close.gif";
		obj.title="关闭左边管理菜单";
	}
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" background="oBlogStyle/OblogStyleAdminImages/center_bg.gif">
  <tr>
    <td><img src="OblogStyle/OblogStyleAdminImages/center_close.gif" style="cursor:hand" title="关闭左边管理菜单" onClick="switchBar(this)"></td>
  </tr>
</table>
<%end sub

sub frame()
	if url="" or instr(url,"HomestayBackctrl-index.asp") then url="HomestayBackctrl-index.asp?actiontype=main"
	with response
	.write "<title>"&oblog.logined_uname&"-用户管理后台</title></head>"
	.write "<frameset rows=""*"" cols=""156,8,*"" framespacing=""0"" frameborder=""0"" border=""false"" id=""frame"" scrolling=""yes"">"
	.write "<frame name=""left"" scrolling=""auto"" marginwidth=""0"" marginheight=""0"" src=""user_top.asp?actiontype=left"">"
	.write "<frame src=""user_top.asp?actiontype=center"" scrolling=""no"">"
	.write "<frame name=""main"" scrolling=""auto"" marginwidth=""0"" marginheight=""0"" src="""&url&""">"
	.write "</frameset><noframes></noframes></html>"
	end with
end sub

function chkpm()
	dim rs,lasttime
	lasttime = Session("refutime")
	If IsDate(lasttime) Then 
		If DateDiff("s",lasttime,Now()) < 30 then
			chkpm=Session("newpm")
			exit function
		end if
	end if
	set rs=oblog.execute("select count(id) from oblog_pm where incept='"&oblog.filt_badstr(oblog.logined_uname)&"' and isreaded=0 and delR=0")
	chkpm=rs(0)
	Session("refutime")=now()
	Session("newpm")=chkpm
	set rs=nothing
end function

function chkframe()
	if oblog.logined_uframe=1 then
		chkframe="<a href=""HomestayBackctrl-index.asp?t=change0"">切换全屏</a>"
	else
		chkframe="<a href=""HomestayBackctrl-index.asp?t=change1"">切换分栏</a>"
	end if
end function

sub change(s)
	dim rs,tmp
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_info from oblog_user where userid="&oblog.logined_uid,conn,1,3
	if not rs.eof then
		tmp=rs(0)
		if instr(tmp,"$") then 
			tmp=split(tmp,"$")
			rs(0)=tmp(0)&"$"&cint(s)
		else
			rs(0)="0$"&cint(s)
		end if
		rs.update
	end if
	rs.close
	set rs=nothing
	if s=1 then
		response.Write("<script language=JavaScript>top.location='HomestayBackctrl-index.asp?url="&oblog.comeurl&"';</script>")
	else
		response.Write("<script language=JavaScript>top.location='"&oblog.comeurl&"';</script>")
	end if
end sub
%>