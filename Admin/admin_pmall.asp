<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MyHomestay 站内消息管理</title>
<link href="images/admin/admin_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<p><%
dim action
action=request("action")
call send()
select case action
	case "save"
	call save()
end select
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
	#list ul { width: 100%; MARGIN: 0px; } 
	#list ul li.1 { width: 90px;text-align: right;} 
	#list ul li.2 { width: 300px;text-align: left;} 
-->
</style>
<br>
</p><p>&nbsp;</p><table align="center" cellpadding="1" cellspacing="1" Class="border">
<form action="admin_pmall.asp?action=save" method="post" name="oblogform">
  <tr class="tdbg">
    <td>收件人：</td>
    <td>
      <input name="u" type="radio" value="1" checked>所有用户
	  <input name="u" type="radio" value="2">所有接待家庭用户
	  <input name="u" type="radio" value="3">所有加盟企业用户
	</td>
  </tr>
 <tr class="tdbg">
    <td>标题：</td>
	<td><input type="text" name="topic" size="45" maxlength="50" /></td>
 </tr>
   <tr class="tdbg">
    <td>内容：<br />(1500字内)</td>
	<td><textarea name="content" cols="45" rows="8"></textarea></td>
  </tr>
 <tr class="tdbg">
   <td></td><td> <INPUT type="hidden" name="id" value="">
        <input type="submit"  value=" 提交 ">
		　<br>
		(此功能比较消耗服务器资源)</td>
</tr>
</form>
</table>
<%
end sub

sub save()
	dim incept,content,sql,rs,inceptid,topic,username,sqlt,rst,u,s
	content=trim(request("content"))
	topic=trim(request("topic"))
	u=clng(request("u"))
	
	Dim tempStrMember'接收者名称；
	select case u
		case 1
			s=""
			tempStrMember= "所有用户"
		case 2
			's=" where user_level=7 "'原vip用户
			s=" where user_level=6 And user_level_team=0 "'所有接待家庭用户；
			tempStrMember= "所有接待家庭用户"
		case 3
			's=" where user_level=6 "'原普通用户
			s=" where user_level=6 And user_level_team=1 "'所有加盟企业用户；
			tempStrMember= "所有加盟企业用户"
	end select
	if content="" then response.write("错误：消息内容不能为空")
	if topic="" then response.write("错误：消息标题不能为空")
	if content<>"" and topic<>"" then
	dim i,kk
	sql="select userid,username from [oblog_user]"&s
	set rs=oblog.execute(sql)
	do while not rs.eof
		username=rs("username")
		sqlt="select * from oblog_pm"
		set rst=server.createobject("adodb.recordset")
		rst.open sqlt,conn,1,3
		rst.addnew
		rst("incept")=oblog.Interceptstr(username,50)
		rst("topic")=oblog.Interceptstr(oblog.filt_badword(topic),100)
		rst("content")=oblog.Interceptstr(oblog.filt_badword(content),3000)
		rst("sender")="我的住家网"
		rst.update
		rst.close
		rs.movenext
	loop
	rs.close
	set rs=nothing
	response.Write("<ul><li>发送给 "& tempStrMember &" 的消息成功！</li></ul>")
end if
end sub

%>
</body>
</html>

