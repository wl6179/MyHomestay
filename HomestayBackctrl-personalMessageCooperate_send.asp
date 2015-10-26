<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/class_blog.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
if not oblog.checkuserlogined_team() then
	response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))'获得当前短信息详情url地址WL;
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MyHomestay 我的住家网 —— 我的消息</title>
<link href="/css/StyleIndex.css" rel="stylesheet" type="text/css" />
</head>

<style>
body {
	scrollbar-face-color:#ffffff;
	scrollbar-highlight-color: #cccccc;
	scrollbar-shadow-color: #cccccc;
	scrollbar-3d-light-color: #ffffff;
	scrollbar-arrow-color : #cccccc;

	SCROLLBAR-TRACK-COLOR: #ffffff;
	SCROLLBAR-DARKSHADOW-COLOR: #ffffff;
}
</style>
<body>
<div style="width:430px;">
  	
	
    <div style="text-align:left;margin:13px;padding:13px;">
<%
dim action
action=request("action")
select case action
	case "send"
	call send
	case "save"
	call save
	case "read"
	call read
end select
%>
	</div>
	
    
	
  </div>
  <!--<div id="bottom"><%'oblog.user_copyright%></div>-->
</body>
</html>
<%
sub send()
	dim rs
%>
<SCRIPT language=javascript>
function changincept()
{
	document.oblogform.incept.value = document.oblogform.selectincept.value;
}
</SCRIPT>
<style type="text/css">
<!--
	#list ul { width: 100%; MARGIN: 0px; padding:0px;} 
	#list li{MARGIN: 0px; padding:10px 0px;}
	#list ul li.t1 { width: 60px;text-align: left; float:left;} 
	#list ul li.t2 { width: 260px;text-align: left; float:left;} 
-->
</style>
<div id="list">
<form action="HomestayBackctrl-personalMessageCooperate_send.asp?action=save" method="post" name="oblogform">
  <h4>我要向MyHomestay写信</h4>
  <span style="font-size:12px;color:#038ad7">非常感谢您对MyHomestay的关注，您的建议将是我们的重要参考。</span>
  <br />
  <span style="font-size:12px;color:#038ad7">客户咨询热线：010-85493388/13146398085 </span>
  <ul> 
    <li class="t1"><strong>收件人</strong>：</li>
    <li class="t2">
	
	
	<% If oblog.logined_uid=8 Then'如果是管理员用户myhomestay（userid=8）的话，可以正常回复；否则只能单向给管理员发送站内消息！ %>
		<input type="text" name="incept" id="incept" width="20" value="<%=trim(request("incept"))%>" />
	<% Else %>
		<img src="/images/sitemanagers.gif" />我的住家网
		<input type="hidden" name="incept" id="incept" width="20" value="myhomestay" />
	<% End If %>
	<%'<input type="text" name="incept" id="incept" width="20" value="<%=trim(request("incept"))% >" />
'	< select size="1" id="selectincept" onChange='javascript:changincept()' >
'	< option value="" >选择好友< /option >
'
'set rs=oblog.execute("select username from oblog_user,oblog_friend where oblog_user.userid=oblog_friend.friendid and oblog_friend.isblack=0 and oblog_friend.userid="&oblog.logined_uid)
'while not rs.eof
'	response.Write("< option value='"&oblog.filt_html(rs(0))&"'>"&oblog.filt_html(rs(0))&"< / option>")
'	rs.movenext
'wend
'set rs=nothing
'
'	</select >%>
	</li >
  </ul >
  
  
  <ul> 
    <li class="t1"><strong>标题</strong>：</li>
	<li class="t2"><input type="text" style="font-weight:normal;" name="topic" size="35" maxlength="50" value="<%=oblog.filt_html(trim(request("topic")))%>" /></li>
  </ul>
   <ul> 
    <strong>内容</strong>：(1500字内)
	<li class="t2"><textarea name="content" cols="50" rows="8"></textarea></li>
  </ul>
  <ul> 
    <INPUT type="hidden" name="id" value="">
        <input type="submit"  value=" 提交 ">
		　　<input type="button" onClick="window.close();" value=" 关闭 ">
  </ul>
  
</form>
</div>
<%
end sub

sub save()
	dim incept,content,sql,rs,inceptid,topic
	incept=oblog.filt_badstr(trim(request("incept")))
	content=trim(request("content"))
	topic=trim(request("topic"))
	if incept="" then oblog.adderrstr("错误：收件人不能为空")
	if content="" then oblog.adderrstr("错误：消息内容不能为空")
	if topic="" then oblog.adderrstr("错误：消息标题不能为空")
	if oblog.errstr<>"" then
		oblog.showusererr
		exit sub
	end if	
	 
	 If oblog.logined_uid=8 Then'如果是管理员用户myhomestay（userid=8）的话，令myhomestay可以正常回复！否则不是消息管理员myhoemstay则只能给管理员一人发送消息！WL 
		sql="select userid from [oblog_user] where username='"&incept&"'"'WL;
	 Else 
		sql="select userid from [oblog_user] where username='myhomestay'"'WL;
	 End If 
	 
	'''sql="select userid from [oblog_user] where username='"&incept&"'"'WL;
	set rs=oblog.execute(sql)
	if rs.eof then
		response.Write("<ul><li>错误：无此用户，或者系统正在维护，请再试一次...<input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>")
		exit sub
	end if
	inceptid=clng(rs(0))
	set rs=oblog.execute("select id from oblog_friend where isblack=1 and userid="&inceptid&" and friendid="&oblog.logined_uid)
	if not rs.eof then
		response.Write("<ul><li>错误：你在收件人的黑名单中，无法发送消息，<input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>")
		exit sub
	end if
	sql="select top 1 * from oblog_pm"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	rs.addnew
	rs("incept")=oblog.Interceptstr(incept,50)
	rs("topic")=oblog.Interceptstr(oblog.filt_badword(topic),100)
	rs("content")=oblog.Interceptstr(oblog.filt_badword(content),3000)
	rs("sender")=oblog.logined_uname
	rs.update
	rs.close
	set rs=nothing
	response.Write("<ul style='list-style-type:none;'><li><img src='/images/mailok.gif' />消息发送成功!<input type='button' onClick='javascript:window.close();' value=' 关闭窗口 ' /></li></ul>")

end sub

sub read()
	dim id,rs
	id=clng(trim(request("id")))
	
	Dim usernameTemp'WL检测当前帐号名，进而推断是否此人的消息；如果是可以查看，如果不是不允许查看！WL
	usernameTemp= oblog.requestID_selectName("oblog_user","username","Where userid="& oblog.logined_uid )
	
	set rs=oblog.execute("select * from oblog_pm where (incept='"& usernameTemp &"' Or sender='"& usernameTemp &"' ) And id="&id)
	if rs.eof then
		set rs=nothing
		response.Write("<ul style='list-style-type:none;'><li>您的信箱里无此记录！<input type='button' onClick='javascript:window.close();' value=' 关闭窗口 ' /></li></ul>")
		exit sub
	end if	
%>
<style type="text/css">
<!--
	 ul { width: 100%; MARGIN: 0px; list-style-type:none; } 
	 li{MARGIN: 0px; padding:10px 0px;}
	 ul li.t1 { width: 80px;text-align: left; float:left;} 
	 ul li.t2 { width: 380px;text-align: left; float:left;} 
-->
</style>

<div style=" width:500px;">
    <h4 style="text-align:center;"><img src='/images/mail_open.gif' />查看消息</h4>
  <ul> 
    <li class="t1"><strong>收件人</strong>：</li>
    <li class="t2">
	
		<% If rs("incept")="我的住家网" Or rs("incept")="管理员" Or rs("incept")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />我的住家网
		<% Else 			
			Response.Write oblog.filt_html(rs("incept"))
		End If%>
		   
		   <%'rs("incept")%></li>
  </ul>
    <ul> 
    <li class="t1"><strong>发件人</strong>：</li>
    <li class="t2">
		<% If rs("sender")="我的住家网" Or rs("sender")="管理员" Or rs("sender")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />我的住家网
		<% Else 			
			Response.Write oblog.filt_html(rs("sender"))
		   End If%>
		   
		   <%'rs("sender")%>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;【发送时间：&nbsp;<%=rs("addtime")%>】</li>
  </ul>

  <ul> 
    <li class="t1"><strong>标题</strong>：</li>
	<li class="t2"><%=oblog.filt_html(rs("topic"))%></li>
  </ul>
   <ul> 
    <li class="t1"><strong>内容</strong>：</li>
	<li class="t2"><%'oblog.filt_html_b(rs("content"))%><%=rs("content")%></li>
	</ul>
	
<ul>
</ul>
<ul>
</ul>
<ul>
</ul>
    <ul style="text-align:center;">
	<% If rs("sender")= usernameTemp Then'如果是在发件箱中，那么，不显示‘回复’按钮！ %>
	
	<% Else %>
		<input type="button" onClick="window.location='HomestayBackctrl-personalMessageCooperate_send.asp?action=send&incept=<%=rs("sender")%>&topic=<%="回复:"&oblog.filt_html(rs("topic"))%>'" value=" 回复 ">
	<% End If %>	
		
		　　<input type="button" onClick="window.close();" value=" 关闭 ">
  </ul>
</div>
<%
	if oblog.logined_uname=rs("incept") then
		oblog.execute("update oblog_pm set isreaded=1 where id="&id&" and incept='"&oblog.logined_uname&"'")
	end if
	set rs=nothing
end sub
%>
