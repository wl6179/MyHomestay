<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
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
	response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
else
	if oblog.logined_ulevel=6 then
		oblog.adderrstr("您未通过管理员审核，不能进入后台")
		oblog.showerr
	end if
end if
if request("t")="change0" then change(0)
if request("t")="change1" then change(1)
'判断分栏
if oblog.logined_uframe=1 then
	cssfile="OblogStyleAdminMainFrame.css"
	if actiontype="" and (instr(nowurl,"HomestayBackctrl-index.asp") and actiontype<>"main") then 
		call frame():response.End()
	end if
	'if instr(nowurl,"HomestayBackctrl-index.asp")=0 and instr(nowurl,"user_top.asp")=0 then
	'	response.Redirect("HomestayBackctrl-index.asp?url="&nowurl)
	'end if
else
	if url<>"" then response.Redirect(url)
	cssfile="OblogStyleAdminMain.css"
end if
if oblog.setup(12,0)=1 then
	uurl="<a href=""http://"&oblog.logined_udomain&""" target=""_blank"">"&oblog.logined_udomain&"</a>"
else
	uurl="<a href="""&oblog.logined_udir&"/"&oblog.logined_ufolder&"/index."&f_ext&""" target=""_blank"">我的首页</a>"
end if
if true_domain=1 and oblog.logined_ucustomdomain<>"" then
	uurl="<a href=""http://"&oblog.logined_ucustomdomain&""" target=""_blank"">"&oblog.logined_ucustomdomain&"</a>"
end if

call utop()

sub utop()%>
<title><%=oblog.logined_uname%>-用户管理后台</title>
<link title="Oblog_Green" href="OblogStyle/<%=cssfile%>" rel="stylesheet" media="all" type="text/css" rev="alternate" />
<script src="inc/function.js" type="text/javascript"></script>
</head>

<body onUnload="chkcopy();">

<div class="none"><p>Oblog用户管理界面（无样式纯文字版，适合手机及PDA，您可以点样式切换回其他界面）</p></div>

<div id="menu">
	<div class="menu_content">
		<div class="menu_navimg"></div>
		<div id="nav">
			<ul id="div1">
				<li class="w100p"><a href="HomestayBackctrl-index.asp" >管理首页</a>
					
          <ul>
            <li><a href="HomestayBackctrl-index.asp">管理首页</a></li>
            <li><a href="user_help.asp">用户管理后台帮助</a></li>
          </ul>
				</li>
				<li><a href="user_blogmanage.asp">文章管理</a>
					
          <ul>
            <li><a href="user_post.asp">添加文章</a></li>
            <li><a href="user_blogmanage.asp?usersearch=5">草稿箱</a></li>
            <li><a href="user_blogmanage.asp">所有文章</a></li>
            <li><a href="user_blogmanage.asp?action=downlog">备份文章</a></li>
            <li><a href="user_subject.asp">文章分类管理</a></li>
          </ul>
				</li>
				<%If EN_PHOTO=1 Then%>
				<li><a href="user_photomanage.asp">相册管理</a>
					
          <ul>
            <li><a href="user_post.asp?t=1">添加图片</a></li>
            <li><a href="user_blogmanage.asp?t=1">所有图片</a></li>
            <li><a href="user_subject.asp?t=1">相册分类管理</a></li>
          </ul>
				</li>
				<%End If%>
				<li  class="w100p"><a href="user_comments.asp">评论留言</a>
					
          <ul>
            <li><a href="user_comments.asp">评论管理</a></li>
            <li><a href="user_messages.asp">留言管理</a></li>
          </ul>
				</li>
				<li><a href="user_template.asp?action=showconfig">模版设置</a>
					
          <ul>
            <li><a href="user_template.asp?action=showconfig">选择模板</a></li>
            <li><a href="user_template.asp?action=modiconfig&amp;editm=1">修改主模板</a></li>
            <li><a href="user_template.asp?action=modiviceconfig&amp;editm=1">修改Blog内容模板</a></li>
            <li><a href="user_template.asp?action=bakskin">备份我当前的模板</a></li>
            <li><a href="user_template.asp?action=good">推荐共享我的模板</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_setting.asp">控制面板</a>
					
          <ul>
            <li><a href="user_setting.asp">站点设置</a></li>
            <li><a href="user_setting.asp?action=placard">修改公告</a></li>
            <li><a href="user_setting.asp?action=links">我的友情连接</a></li>
            <li><a href="user_setting.asp?action=userinfo">我的个人资料</a></li>
            <li><a href="user_setting.asp?action=userpassword">修改我的密码</a></li>
            <li><a href="user_setting.asp?action=blogpassword">加密整站</a></li>
            <li><a href="user_setting.asp?action=blogstar">申请用户明星</a></li>
          </ul>
				</li>
				<li><a href="user_team.asp">团队好友</a>
					
          <ul>
            <li><a href="user_friends.asp?action=add">添加好友</a></li>
            <li><a href="user_friends.asp">好友管理</a></li>
            <li><a href="user_friends.asp?action=add&amp;type=black">添加黑名单</a></li>
            <li><a href="user_friends.asp?usersearch=1">管理黑名单</a></li>
            <li><a href="user_team.asp">用户团队</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_files.asp">文件管理</a>
					
          <ul>
            <li><a href="user_files.asp">所有文件</a></li>
            <li><a href="user_files.asp?usersearch=1">图片文件</a></li>
            <li><a href="user_files.asp?usersearch=2">压缩文件</a></li>
            <li><a href="user_files.asp?usersearch=3">文档文件</a></li>
            <li style="display:none"><a href="user_files.asp?usersearch=4">媒体文件</a></li>
          </ul>
				</li>
			</ul>
		</div>
	</div>
</div>

<div id="top">

<div class="logo"></div>
  <div class="top_button">
	<a href="user_post.asp"><div class="none">发布Blog </div><img src="OblogStyle/OblogStyleAdminImages/space.gif" alt="发布Blog" border="0" class="ico_5" /></a>
		<a href="user_blogmanage.asp">
	<div class="none">管理Blog </div><img class="ico_4" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="管理Blog" /></a>
		<%if en_photo=1 then%>
		<a href="user_photomanage.asp"><div class="none">我的相册</div><img class="ico_7" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="我的相册" /></a>
		<%end if%>
		<a href="user_messages.asp"><div class="none">留言管理</div><img class="ico_3" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="留言管理"/></a>
		<a href="user_comments.asp"><div class="none">评论管理</div><img  class="ico_2" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="评论管理" /></a>
		<a href="user_update.asp"><div class="none">重建首页</div><img class="ico_1" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="重建首页" /></a>
		<a  href="HomestayBackctrl-index.asp?t=logout"><div class="none">退出登陆</div><img class="ico_6" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="退出登陆" /></a>
  </div>
  <div class="top_menu">
  <%=uurl%> | <a href="user_pmmanage.asp">短消息(<%=chkpm()%>)</a> | <a href="index.asp" target="_blank">站点首页</a> | <%=chkframe()%></div>
</div>
<%end sub

sub uleft()%>
<link href="oBlogStyle/OblogStyleAdminMenu.css" rel="stylesheet" type="text/css" />
<script>
function menu(oblog){
if(oblog.style.display=='')oblog.style.display='none';
else oblog.style.display='';
}
</script>
<div id="toplogo"></div>
<div id="sub"><a href="HomestayBackctrl-index.asp" target="_top">管理首页</a> | <a href="user_help.asp" target="main">指引帮助</a></div>
<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 文章管理</div>
<ul id="oblog_1">
<li><a href="user_post.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> 添加文章</a></li>
<li><a href="user_blogmanage.asp?usersearch=5" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/12.gif" align="absmiddle" /> 草稿箱</a></li>
<li><a href="user_blogmanage.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> 所有文章</a></li>
<li><a href="user_blogmanage.asp?action=downlog" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 备份文章</a></li>
<li><a href="user_subject.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/16.gif" align="absmiddle" /> 文章分类管理</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_3)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 评论留言</div>
<ul id="oblog_3">
<li><a href="user_comments.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/31.gif" align="absmiddle" /> 评论管理</a></li>
<li><a href="user_messages.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/32.gif" align="absmiddle" /> 留言管理</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_4)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 模版设置</div>
<ul id="oblog_4" style="display: none;">
<li><a href="user_template.asp?action=showconfig" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/41.gif" align="absmiddle" /> 选择模板</a></li>
<li><a href="user_template.asp?action=modiconfig&amp;editm=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/42.gif" align="absmiddle" /> 修改主模板</a></li>
<li><a href="user_template.asp?action=modiviceconfig&amp;editm=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/43.gif" align="absmiddle" /> 修改内容模板</a></li>
<li><a href="user_template.asp?action=bakskin" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/44.gif" align="absmiddle" /> 备份模板</a></li>
<li><a href="user_template.asp?action=good" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/45.gif" align="absmiddle" /> 推荐我的模板</a></li>
</ul>
<%If EN_PHOTO=1 Then%>
<div class="menu-t1" onClick="menu(oblog_2)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 相册管理</div>
<ul id="oblog_2" style="display: none;">
<li><a href="user_post.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/21.gif" align="absmiddle" /> 添加图片</a></li>
<li><a href="user_blogmanage.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/22.gif" align="absmiddle" /> 所有图片</a></li>
<li><a href="user_subject.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/24.gif" align="absmiddle" /> 相册分类管理</a></li>
</li>
</ul>
<%end if%>
<div class="menu-t1" onClick="menu(oblog_7)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 文件管理</div>
<ul id="oblog_7" style="display: none;">
<li><a href="user_files.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/71.gif" align="absmiddle" /> 所有文件</a></li>
<li><a href="user_files.asp?usersearch=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/72.gif" align="absmiddle" /> 图片文件</a></li>
<li><a href="user_files.asp?usersearch=2" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/73.gif" align="absmiddle" /> 压缩文件</a></li>
<li><a href="user_files.asp?usersearch=3" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/74.gif" align="absmiddle" /> 文档文件</a></li>
<li style="display:none"><a href="user_files.asp?usersearch=4" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/75.gif" align="absmiddle" /> 媒体文件</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_6)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 团队好友</div>
<ul id="oblog_6" style="display: none;">
<li><a href="user_friends.asp?action=add" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/61.gif" align="absmiddle" /> 添加好友</a></li>
<li><a href="user_friends.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/62.gif" align="absmiddle" /> 好友管理</a></li>
<li><a href="user_friends.asp?action=add&amp;type=black" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/63.gif" align="absmiddle" /> 添加黑名单</a></li>
<li><a href="user_friends.asp?usersearch=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/64.gif" align="absmiddle" /> 管理黑名单</a></li>
<li><a href="user_team.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/65.gif" align="absmiddle" /> 用户团队</a></li>
</ul>


<div class="menu-t1" onClick="menu(oblog_5)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 参数设置</div>
<ul id="oblog_5" style="display: none;">
<li><a href="user_setting.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/51.gif" align="absmiddle" /> 站点设置</a></li>
<li><a href="user_setting.asp?action=placard" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/52.gif" align="absmiddle" /> 修改公告</a></li>
<li><a href="user_setting.asp?action=links" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/53.gif" align="absmiddle" /> 友情连接</a></li>
<li><a href="user_setting.asp?action=userinfo" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/54.gif" align="absmiddle" /> 修改个人资料</a></li>
<li><a href="user_setting.asp?action=userpassword" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/55.gif" align="absmiddle" /> 修改登录密码</a></li>
<li><a href="user_setting.asp?action=blogpassword" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/57.gif" align="absmiddle" /> 加密整站</a></li>
<li><a href="user_setting.asp?action=blogstar" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/58.gif" align="absmiddle" /> 申请用户明星</a></li>
</ul>
</body>
</html>
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