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
	response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))'��õ�ǰ����Ϣ����url��ַWL;
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MyHomestay �ҵ�ס���� ���� �ҵ���Ϣ</title>
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
  <h4>��Ҫ��MyHomestayд��</h4>
  <span style="font-size:12px;color:#038ad7">�ǳ���л����MyHomestay�Ĺ�ע�����Ľ��齫�����ǵ���Ҫ�ο���</span>
  <br />
  <span style="font-size:12px;color:#038ad7">�ͻ���ѯ���ߣ�010-85493388/13146398085 </span>
  <ul> 
    <li class="t1"><strong>�ռ���</strong>��</li>
    <li class="t2">
	
	
	<% If oblog.logined_uid=8 Then'����ǹ���Ա�û�myhomestay��userid=8���Ļ������������ظ�������ֻ�ܵ��������Ա����վ����Ϣ�� %>
		<input type="text" name="incept" id="incept" width="20" value="<%=trim(request("incept"))%>" />
	<% Else %>
		<img src="/images/sitemanagers.gif" />�ҵ�ס����
		<input type="hidden" name="incept" id="incept" width="20" value="myhomestay" />
	<% End If %>
	<%'<input type="text" name="incept" id="incept" width="20" value="<%=trim(request("incept"))% >" />
'	< select size="1" id="selectincept" onChange='javascript:changincept()' >
'	< option value="" >ѡ�����< /option >
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
    <li class="t1"><strong>����</strong>��</li>
	<li class="t2"><input type="text" style="font-weight:normal;" name="topic" size="35" maxlength="50" value="<%=oblog.filt_html(trim(request("topic")))%>" /></li>
  </ul>
   <ul> 
    <strong>����</strong>��(1500����)
	<li class="t2"><textarea name="content" cols="50" rows="8"></textarea></li>
  </ul>
  <ul> 
    <INPUT type="hidden" name="id" value="">
        <input type="submit"  value=" �ύ ">
		����<input type="button" onClick="window.close();" value=" �ر� ">
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
	if incept="" then oblog.adderrstr("�����ռ��˲���Ϊ��")
	if content="" then oblog.adderrstr("������Ϣ���ݲ���Ϊ��")
	if topic="" then oblog.adderrstr("������Ϣ���ⲻ��Ϊ��")
	if oblog.errstr<>"" then
		oblog.showusererr
		exit sub
	end if	
	 
	 If oblog.logined_uid=8 Then'����ǹ���Ա�û�myhomestay��userid=8���Ļ�����myhomestay���������ظ�����������Ϣ����Աmyhoemstay��ֻ�ܸ�����Աһ�˷�����Ϣ��WL 
		sql="select userid from [oblog_user] where username='"&incept&"'"'WL;
	 Else 
		sql="select userid from [oblog_user] where username='myhomestay'"'WL;
	 End If 
	 
	'''sql="select userid from [oblog_user] where username='"&incept&"'"'WL;
	set rs=oblog.execute(sql)
	if rs.eof then
		response.Write("<ul><li>�����޴��û�������ϵͳ����ά����������һ��...<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>")
		exit sub
	end if
	inceptid=clng(rs(0))
	set rs=oblog.execute("select id from oblog_friend where isblack=1 and userid="&inceptid&" and friendid="&oblog.logined_uid)
	if not rs.eof then
		response.Write("<ul><li>���������ռ��˵ĺ������У��޷�������Ϣ��<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>")
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
	response.Write("<ul style='list-style-type:none;'><li><img src='/images/mailok.gif' />��Ϣ���ͳɹ�!<input type='button' onClick='javascript:window.close();' value=' �رմ��� ' /></li></ul>")

end sub

sub read()
	dim id,rs
	id=clng(trim(request("id")))
	
	Dim usernameTemp'WL��⵱ǰ�ʺ����������ƶ��Ƿ���˵���Ϣ������ǿ��Բ鿴��������ǲ�����鿴��WL
	usernameTemp= oblog.requestID_selectName("oblog_user","username","Where userid="& oblog.logined_uid )
	
	set rs=oblog.execute("select * from oblog_pm where (incept='"& usernameTemp &"' Or sender='"& usernameTemp &"' ) And id="&id)
	if rs.eof then
		set rs=nothing
		response.Write("<ul style='list-style-type:none;'><li>�����������޴˼�¼��<input type='button' onClick='javascript:window.close();' value=' �رմ��� ' /></li></ul>")
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
    <h4 style="text-align:center;"><img src='/images/mail_open.gif' />�鿴��Ϣ</h4>
  <ul> 
    <li class="t1"><strong>�ռ���</strong>��</li>
    <li class="t2">
	
		<% If rs("incept")="�ҵ�ס����" Or rs("incept")="����Ա" Or rs("incept")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />�ҵ�ס����
		<% Else 			
			Response.Write oblog.filt_html(rs("incept"))
		End If%>
		   
		   <%'rs("incept")%></li>
  </ul>
    <ul> 
    <li class="t1"><strong>������</strong>��</li>
    <li class="t2">
		<% If rs("sender")="�ҵ�ס����" Or rs("sender")="����Ա" Or rs("sender")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />�ҵ�ס����
		<% Else 			
			Response.Write oblog.filt_html(rs("sender"))
		   End If%>
		   
		   <%'rs("sender")%>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ʱ�䣺&nbsp;<%=rs("addtime")%>��</li>
  </ul>

  <ul> 
    <li class="t1"><strong>����</strong>��</li>
	<li class="t2"><%=oblog.filt_html(rs("topic"))%></li>
  </ul>
   <ul> 
    <li class="t1"><strong>����</strong>��</li>
	<li class="t2"><%'oblog.filt_html_b(rs("content"))%><%=rs("content")%></li>
	</ul>
	
<ul>
</ul>
<ul>
</ul>
<ul>
</ul>
    <ul style="text-align:center;">
	<% If rs("sender")= usernameTemp Then'������ڷ������У���ô������ʾ���ظ�����ť�� %>
	
	<% Else %>
		<input type="button" onClick="window.location='HomestayBackctrl-personalMessageCooperate_send.asp?action=send&incept=<%=rs("sender")%>&topic=<%="�ظ�:"&oblog.filt_html(rs("topic"))%>'" value=" �ظ� ">
	<% End If %>	
		
		����<input type="button" onClick="window.close();" value=" �ر� ">
  </ul>
</div>
<%
	if oblog.logined_uname=rs("incept") then
		oblog.execute("update oblog_pm set isreaded=1 where id="&id&" and incept='"&oblog.logined_uname&"'")
	end if
	set rs=nothing
end sub
%>
