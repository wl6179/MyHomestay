<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>MyHomestay վ����Ϣ����</title>
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
    <td>�ռ��ˣ�</td>
    <td>
      <input name="u" type="radio" value="1" checked>�����û�
	  <input name="u" type="radio" value="2">���нӴ���ͥ�û�
	  <input name="u" type="radio" value="3">���м�����ҵ�û�
	</td>
  </tr>
 <tr class="tdbg">
    <td>���⣺</td>
	<td><input type="text" name="topic" size="45" maxlength="50" /></td>
 </tr>
   <tr class="tdbg">
    <td>���ݣ�<br />(1500����)</td>
	<td><textarea name="content" cols="45" rows="8"></textarea></td>
  </tr>
 <tr class="tdbg">
   <td></td><td> <INPUT type="hidden" name="id" value="">
        <input type="submit"  value=" �ύ ">
		��<br>
		(�˹��ܱȽ����ķ�������Դ)</td>
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
	
	Dim tempStrMember'���������ƣ�
	select case u
		case 1
			s=""
			tempStrMember= "�����û�"
		case 2
			's=" where user_level=7 "'ԭvip�û�
			s=" where user_level=6 And user_level_team=0 "'���нӴ���ͥ�û���
			tempStrMember= "���нӴ���ͥ�û�"
		case 3
			's=" where user_level=6 "'ԭ��ͨ�û�
			s=" where user_level=6 And user_level_team=1 "'���м�����ҵ�û���
			tempStrMember= "���м�����ҵ�û�"
	end select
	if content="" then response.write("������Ϣ���ݲ���Ϊ��")
	if topic="" then response.write("������Ϣ���ⲻ��Ϊ��")
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
		rst("sender")="�ҵ�ס����"
		rst.update
		rst.close
		rs.movenext
	loop
	rs.close
	set rs=nothing
	response.Write("<ul><li>���͸� "& tempStrMember &" ����Ϣ�ɹ���</li></ul>")
end if
end sub

%>
</body>
</html>

