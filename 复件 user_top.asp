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
	tName="ͼƬ"
Else
	t="0"
	tName="����"
End If

dim oblog,uurl
set oblog=new class_sys
oblog.start
'�˳�ϵͳ
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
	else'�˳���ص���ҳ��
		response.Write("<script language=JavaScript>top.location='index.asp';</script>")
	end if
	Set oBlog=Nothing
	response.End()
End If
if not oblog.checkuserlogined() then
	response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
else
	if oblog.logined_ulevel=6 then
		oblog.adderrstr("��δͨ������Ա��ˣ����ܽ����̨")
		oblog.showerr
	end if
end if
if request("t")="change0" then change(0)
if request("t")="change1" then change(1)
'�жϷ���
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
	uurl="<a href="""&oblog.logined_udir&"/"&oblog.logined_ufolder&"/index."&f_ext&""" target=""_blank"">�ҵ���ҳ</a>"
end if
if true_domain=1 and oblog.logined_ucustomdomain<>"" then
	uurl="<a href=""http://"&oblog.logined_ucustomdomain&""" target=""_blank"">"&oblog.logined_ucustomdomain&"</a>"
end if

call utop()

sub utop()%>
<title><%=oblog.logined_uname%>-�û������̨</title>
<link title="Oblog_Green" href="OblogStyle/<%=cssfile%>" rel="stylesheet" media="all" type="text/css" rev="alternate" />
<script src="inc/function.js" type="text/javascript"></script>
</head>

<body onUnload="chkcopy();">

<div class="none"><p>Oblog�û�������棨����ʽ�����ְ棬�ʺ��ֻ���PDA�������Ե���ʽ�л����������棩</p></div>

<div id="menu">
	<div class="menu_content">
		<div class="menu_navimg"></div>
		<div id="nav">
			<ul id="div1">
				<li class="w100p"><a href="HomestayBackctrl-index.asp" >������ҳ</a>
					
          <ul>
            <li><a href="HomestayBackctrl-index.asp">������ҳ</a></li>
            <li><a href="user_help.asp">�û������̨����</a></li>
          </ul>
				</li>
				<li><a href="user_blogmanage.asp">���¹���</a>
					
          <ul>
            <li><a href="user_post.asp">�������</a></li>
            <li><a href="user_blogmanage.asp?usersearch=5">�ݸ���</a></li>
            <li><a href="user_blogmanage.asp">��������</a></li>
            <li><a href="user_blogmanage.asp?action=downlog">��������</a></li>
            <li><a href="user_subject.asp">���·������</a></li>
          </ul>
				</li>
				<%If EN_PHOTO=1 Then%>
				<li><a href="user_photomanage.asp">������</a>
					
          <ul>
            <li><a href="user_post.asp?t=1">���ͼƬ</a></li>
            <li><a href="user_blogmanage.asp?t=1">����ͼƬ</a></li>
            <li><a href="user_subject.asp?t=1">���������</a></li>
          </ul>
				</li>
				<%End If%>
				<li  class="w100p"><a href="user_comments.asp">��������</a>
					
          <ul>
            <li><a href="user_comments.asp">���۹���</a></li>
            <li><a href="user_messages.asp">���Թ���</a></li>
          </ul>
				</li>
				<li><a href="user_template.asp?action=showconfig">ģ������</a>
					
          <ul>
            <li><a href="user_template.asp?action=showconfig">ѡ��ģ��</a></li>
            <li><a href="user_template.asp?action=modiconfig&amp;editm=1">�޸���ģ��</a></li>
            <li><a href="user_template.asp?action=modiviceconfig&amp;editm=1">�޸�Blog����ģ��</a></li>
            <li><a href="user_template.asp?action=bakskin">�����ҵ�ǰ��ģ��</a></li>
            <li><a href="user_template.asp?action=good">�Ƽ������ҵ�ģ��</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_setting.asp">�������</a>
					
          <ul>
            <li><a href="user_setting.asp">վ������</a></li>
            <li><a href="user_setting.asp?action=placard">�޸Ĺ���</a></li>
            <li><a href="user_setting.asp?action=links">�ҵ���������</a></li>
            <li><a href="user_setting.asp?action=userinfo">�ҵĸ�������</a></li>
            <li><a href="user_setting.asp?action=userpassword">�޸��ҵ�����</a></li>
            <li><a href="user_setting.asp?action=blogpassword">������վ</a></li>
            <li><a href="user_setting.asp?action=blogstar">�����û�����</a></li>
          </ul>
				</li>
				<li><a href="user_team.asp">�ŶӺ���</a>
					
          <ul>
            <li><a href="user_friends.asp?action=add">��Ӻ���</a></li>
            <li><a href="user_friends.asp">���ѹ���</a></li>
            <li><a href="user_friends.asp?action=add&amp;type=black">��Ӻ�����</a></li>
            <li><a href="user_friends.asp?usersearch=1">���������</a></li>
            <li><a href="user_team.asp">�û��Ŷ�</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_files.asp">�ļ�����</a>
					
          <ul>
            <li><a href="user_files.asp">�����ļ�</a></li>
            <li><a href="user_files.asp?usersearch=1">ͼƬ�ļ�</a></li>
            <li><a href="user_files.asp?usersearch=2">ѹ���ļ�</a></li>
            <li><a href="user_files.asp?usersearch=3">�ĵ��ļ�</a></li>
            <li style="display:none"><a href="user_files.asp?usersearch=4">ý���ļ�</a></li>
          </ul>
				</li>
			</ul>
		</div>
	</div>
</div>

<div id="top">

<div class="logo"></div>
  <div class="top_button">
	<a href="user_post.asp"><div class="none">����Blog </div><img src="OblogStyle/OblogStyleAdminImages/space.gif" alt="����Blog" border="0" class="ico_5" /></a>
		<a href="user_blogmanage.asp">
	<div class="none">����Blog </div><img class="ico_4" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="����Blog" /></a>
		<%if en_photo=1 then%>
		<a href="user_photomanage.asp"><div class="none">�ҵ����</div><img class="ico_7" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="�ҵ����" /></a>
		<%end if%>
		<a href="user_messages.asp"><div class="none">���Թ���</div><img class="ico_3" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="���Թ���"/></a>
		<a href="user_comments.asp"><div class="none">���۹���</div><img  class="ico_2" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="���۹���" /></a>
		<a href="user_update.asp"><div class="none">�ؽ���ҳ</div><img class="ico_1" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="�ؽ���ҳ" /></a>
		<a  href="HomestayBackctrl-index.asp?t=logout"><div class="none">�˳���½</div><img class="ico_6" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="�˳���½" /></a>
  </div>
  <div class="top_menu">
  <%=uurl%> | <a href="user_pmmanage.asp">����Ϣ(<%=chkpm()%>)</a> | <a href="index.asp" target="_blank">վ����ҳ</a> | <%=chkframe()%></div>
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
<div id="sub"><a href="HomestayBackctrl-index.asp" target="_top">������ҳ</a> | <a href="user_help.asp" target="main">ָ������</a></div>
<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ���¹���</div>
<ul id="oblog_1">
<li><a href="user_post.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> �������</a></li>
<li><a href="user_blogmanage.asp?usersearch=5" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/12.gif" align="absmiddle" /> �ݸ���</a></li>
<li><a href="user_blogmanage.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> ��������</a></li>
<li><a href="user_blogmanage.asp?action=downlog" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> ��������</a></li>
<li><a href="user_subject.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/16.gif" align="absmiddle" /> ���·������</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_3)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ��������</div>
<ul id="oblog_3">
<li><a href="user_comments.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/31.gif" align="absmiddle" /> ���۹���</a></li>
<li><a href="user_messages.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/32.gif" align="absmiddle" /> ���Թ���</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_4)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ģ������</div>
<ul id="oblog_4" style="display: none;">
<li><a href="user_template.asp?action=showconfig" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/41.gif" align="absmiddle" /> ѡ��ģ��</a></li>
<li><a href="user_template.asp?action=modiconfig&amp;editm=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/42.gif" align="absmiddle" /> �޸���ģ��</a></li>
<li><a href="user_template.asp?action=modiviceconfig&amp;editm=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/43.gif" align="absmiddle" /> �޸�����ģ��</a></li>
<li><a href="user_template.asp?action=bakskin" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/44.gif" align="absmiddle" /> ����ģ��</a></li>
<li><a href="user_template.asp?action=good" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/45.gif" align="absmiddle" /> �Ƽ��ҵ�ģ��</a></li>
</ul>
<%If EN_PHOTO=1 Then%>
<div class="menu-t1" onClick="menu(oblog_2)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ������</div>
<ul id="oblog_2" style="display: none;">
<li><a href="user_post.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/21.gif" align="absmiddle" /> ���ͼƬ</a></li>
<li><a href="user_blogmanage.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/22.gif" align="absmiddle" /> ����ͼƬ</a></li>
<li><a href="user_subject.asp?t=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/24.gif" align="absmiddle" /> ���������</a></li>
</li>
</ul>
<%end if%>
<div class="menu-t1" onClick="menu(oblog_7)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> �ļ�����</div>
<ul id="oblog_7" style="display: none;">
<li><a href="user_files.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/71.gif" align="absmiddle" /> �����ļ�</a></li>
<li><a href="user_files.asp?usersearch=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/72.gif" align="absmiddle" /> ͼƬ�ļ�</a></li>
<li><a href="user_files.asp?usersearch=2" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/73.gif" align="absmiddle" /> ѹ���ļ�</a></li>
<li><a href="user_files.asp?usersearch=3" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/74.gif" align="absmiddle" /> �ĵ��ļ�</a></li>
<li style="display:none"><a href="user_files.asp?usersearch=4" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/75.gif" align="absmiddle" /> ý���ļ�</a></li>
</ul>

<div class="menu-t1" onClick="menu(oblog_6)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> �ŶӺ���</div>
<ul id="oblog_6" style="display: none;">
<li><a href="user_friends.asp?action=add" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/61.gif" align="absmiddle" /> ��Ӻ���</a></li>
<li><a href="user_friends.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/62.gif" align="absmiddle" /> ���ѹ���</a></li>
<li><a href="user_friends.asp?action=add&amp;type=black" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/63.gif" align="absmiddle" /> ��Ӻ�����</a></li>
<li><a href="user_friends.asp?usersearch=1" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/64.gif" align="absmiddle" /> ���������</a></li>
<li><a href="user_team.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/65.gif" align="absmiddle" /> �û��Ŷ�</a></li>
</ul>


<div class="menu-t1" onClick="menu(oblog_5)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ��������</div>
<ul id="oblog_5" style="display: none;">
<li><a href="user_setting.asp" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/51.gif" align="absmiddle" /> վ������</a></li>
<li><a href="user_setting.asp?action=placard" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/52.gif" align="absmiddle" /> �޸Ĺ���</a></li>
<li><a href="user_setting.asp?action=links" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/53.gif" align="absmiddle" /> ��������</a></li>
<li><a href="user_setting.asp?action=userinfo" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/54.gif" align="absmiddle" /> �޸ĸ�������</a></li>
<li><a href="user_setting.asp?action=userpassword" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/55.gif" align="absmiddle" /> �޸ĵ�¼����</a></li>
<li><a href="user_setting.asp?action=blogpassword" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/57.gif" align="absmiddle" /> ������վ</a></li>
<li><a href="user_setting.asp?action=blogstar" target="main"><img src="oBlogStyle/OblogStyleAdminImages/li/58.gif" align="absmiddle" /> �����û�����</a></li>
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
		obj.title="����߹���˵�";
	}
	else{
		parent.frame.cols="156,8,*";
		displayBar=true;
		obj.src="OblogStyle/OblogStyleAdminImages/center_close.gif";
		obj.title="�ر���߹���˵�";
	}
}
</script>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" background="oBlogStyle/OblogStyleAdminImages/center_bg.gif">
  <tr>
    <td><img src="OblogStyle/OblogStyleAdminImages/center_close.gif" style="cursor:hand" title="�ر���߹���˵�" onClick="switchBar(this)"></td>
  </tr>
</table>
<%end sub

sub frame()
	if url="" or instr(url,"HomestayBackctrl-index.asp") then url="HomestayBackctrl-index.asp?actiontype=main"
	with response
	.write "<title>"&oblog.logined_uname&"-�û������̨</title></head>"
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
		chkframe="<a href=""HomestayBackctrl-index.asp?t=change0"">�л�ȫ��</a>"
	else
		chkframe="<a href=""HomestayBackctrl-index.asp?t=change1"">�л�����</a>"
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