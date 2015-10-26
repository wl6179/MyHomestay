<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_blog.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>引用通告</title>
<style type="text/css">
<!--
@charset "gb2312";
* {margin: 0;padding: 0;border: 0;}
body {
color: #444;
background: #fff;
text-align: left;
margin: 0;
font-family: 'Century Gothic', Arial, Helvetica, sans-serif;
font-size: 12px;
line-height: 150%;

}
.content {
padding:0px 100px 0px 100px;

}
.content_body {
border-left: 1px #694659 solid;
border-right: 1px #694659 solid;
border-bottom: 1px #694659 solid;
background: #fff;
padding:0px;
}
.content_body h1 {
font-size:14px;
padding:6px 0px 0px 20px;
color:#f30;
font-weight:400;
}
.content_top {
padding:6px 0px 0px 15px;
width:100%;
height:30px;
background: url("images/yinyongmenu.png") no-repeat top right;
color:#fff;
font-size:14px;
font-weight:bold;
border-left: 1px #694659 solid;
border-right: 1px #694659 solid;
}
#list {
background: url("images/menubgline.png");
}
#list h1 {
padding:8px 6px 8px 6px;
font-size:14px;
color:#333;
font-weight:600;
}
#list .list_edit a {
color:#099;
font-weight:bold;
text-decoration: none;
}
#list .list_edit a:hover {
color:#f90;
font-weight:bold;
text-decoration: underline;
word-spacing:15px;
}
#list ul {
padding:1px 0px 3px 15px ;
}
-->
</style>
</head>
<body>
<div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			引用通告
	</div>
	
    <div class="content_body">
<%
dim oblog,logid
set oblog=new class_sys
oblog.start
logid=clng(request("id"))
typetb()
%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</body>
</html>
<%

sub typetb()
dim i,tburl
tburl=oblog.setup(3,0)&"tb.asp?id="&logid
%>
<h1>本文引用地址：<%=tburl%> <input type="button" onclick="window.clipboardData.setData('Text','<%=tburl%>');" value="复制"></h1>
<%
	dim rs
	set rs = oblog.Execute("select * from oblog_trackback where logid="&logid&" order by id desc")
	if rs.eof then 
%>
<ul class="list_edit">该文章没有被引用。</ul>
<%else%>
<h1>该文章有如下引用</h1>
<div id="list">
<%
	i = 1
	while not rs.eof 
	%>
	<h1>  第<%=i%>条:<%="被"&rs("tbuser")&"在"&rs("addtime")&"引用"%></h1>
	<ul class="list_edit">引用标题：<%=oblog.filt_html(rs("topic"))%></ul>
	<ul class="list_edit">引用地址：<a href="<%=rs("url")%>" target="_blank"><%=rs("url")%></a></ul>
      <%
		i = i+1
		rs.movenext
	wend
	set rs=nothing
%> 
</div>
<%	
	end if
end sub
%>
