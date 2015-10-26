<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_blog.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
if not oblog.checkuserlogined() then
	response.Redirect("login.asp")
end if
if en_photo=0 then
	oblog.adderrstr("此功能已被系统关闭！")
	oblog.showusererr
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>相册管理页面</title>
<link href="OblogStyle/OblogStyleAdminMain.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			相册管理
	</div>
	
    <div class="content_body">
<%
dim action
action=request("action")
select case action
	case "modifyphoto"
	call modifyphoto
	case "savephoto"
	call savephoto
end select
%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
  <div id="bottom"><%=oblog.user_copyright%></div>
</body>
</html>
<%
sub modifyphoto()
	dim id,rs,sql
	dim restr
	id=trim(request("id"))
	if id="" then
		oblog.adderrstr( "错误：参数不足！")
		oblog.showusererr
		exit sub
	else
		id=Clng(id)
	end if
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_upfile] where fileid=" & id
	else
		sql="select * from [oblog_upfile] where fileid=" & id&" and userid="&oblog.logined_uid
	end if
	set rs=oblog.execute(sql)
	if rs.bof then
		rs.close
		set rs=nothing
		oblog.adderrstr( "错误：找不到指定的文件！")
		oblog.showusererr
		exit sub
	end if
%>
<style type="text/css">
<!--
	#list ul { width: 100%; MARGIN: 0px; } 
	#list ul li.1 { width: 90px;text-align: right;} 
	#list ul li.2 { width: 300px;text-align: left;} 
-->
</style>

<div id="list">
<form action="user_photo.asp?action=savephoto" method="post" name="oblogform">
	<h1>添加图片进相册</h1>
	<ul>
	<li class="1">文件名：</li>
	<li class="2"><%=rs("file_path")%></li>
	</ul>
	<ul>
	<li class="1">图片：</li>
	<li class="2"><%="<img src="""&rs("file_path")&"""  width=""300"" border=""1"" />"%></li>
	</ul>
	<ul> 
    <li class="1">图片说明：<br />(250字内)</li>
	<li class="2"><textarea name="readme" cols="45" rows="5"><%=oblog.filt_html(rs("file_readme"))%></textarea></li>
	</ul>
	<ul><input type="hidden" name="id" value="<%=rs("fileid")%>" />
        <input type="submit"  value="提交" />
	</ul>
</form>
</div>
<%
	rs.close
	set rs=nothing
end sub

sub savephoto()
	dim id,rs,sql,issave,photofile
	id=request("id")
		if id="" then
		oblog.adderrstr( "错误：参数不足！")
		oblog.showusererr
		exit sub
	else
		id=Clng(id)
	end if
	sql="select * from [oblog_upfile] where fileid=" & id&" and userid="&oblog.logined_uid
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		issave=rs("issave")
		photofile=rs("photofile")
		rs("file_readme")=oblog.InterceptStr(oblog.filt_badword(trim(request("readme"))),250)
		rs("isphoto")=1
		rs.update
		rs.close
		dim blog
		set blog=new class_blog
		blog.userid=oblog.logined_uid
		if issave=1 and photofile<>"" then
			set rs=oblog.execute("select sdate,edate,readme from oblog_photo_save where userid="&oblog.logined_uid&" and photofile='"&photofile&"'")
			if not rs.eof then
				oblog.execute("update oblog_upfile set issave=0 where isphoto=1 and userid="&oblog.logined_uid&" and addtime<=#"&dateadd("d",1,rs(1))&"# and addtime>=#"&rs(0)&"#")
				blog.update_album 1,rs(0),rs(1),rs(2)
			end if
		else
			blog.update_album 1,0,0,""
		end if
		set rs=nothing
		set blog=nothing
	end if
end sub

%>
