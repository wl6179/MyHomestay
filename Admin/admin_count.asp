<!--#include file="inc/inc_sys.asp"-->
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<%
dim action,rs,sql,trs
action=request("action")
if action="DoUpdate" then
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from oblog_setup"
	rs.Open sql,Conn,1,3
	set trs=oblog.execute("select count(logID) from [oblog_log]")
	if isNull(trs(0)) then
		rs("log_count")=0
	else
		rs("log_count")=trs(0)
	end if
	set trs=oblog.execute("select count(commentID) from oblog_comment")
	if isNull(trs(0)) then
		rs("comment_count")=0
	else
		rs("comment_count")=trs(0)
	end if
	set trs=oblog.execute("select count(messageID) from oblog_message")
	if isNull(trs(0)) then
		rs("message_count")=0
	else
		rs("message_count")=trs(0)
	end if
	set trs=oblog.execute("select count(userID) from [oblog_user]")
	if isNull(trs(0)) then
		rs("user_count")=0
	else
		rs("user_count")=trs(0)
	end if
	rs.update
	rs.close
	set rs=nothing
	set trs=nothing
	oblog.ReloadSetup()
	Application.Lock
	Application(oblog.cache_name&"index_update")=true
	application(oblog.cache_name&"list_update")=true
	application(oblog.cache_name&"class_update")=true
	Application.unLock
	response.Write("<script src="""&blogdir&"index.asp?re=0""></script>")
	response.Write("<br>�Ѿ��ɹ���ϵͳ���ݽ����˸��£�")
	else
	%>
	
<title>����ϵͳ����</title><form name="form1" method="post" action="">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center" class="title">
    <td height="22" colspan="2"><strong><font color="#FFFFFF">�� �� ϵ ͳ �� ��</font></strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><p>˵����<br>
        1�������������¼���ϵͳ�����£����ۣ������������»������ҳ��<br>
        2���������������ķ�������Դ������ϸȷ��ÿһ��������ִ�С�</p></td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="����ϵͳ����">
    <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
  </tr>
</table>
   </form>
<%
end if
%>

