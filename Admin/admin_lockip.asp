<!--#include file="inc/inc_sys.asp"-->
<%
dim action
Action=trim(request("Action"))
if Action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
dim rs
set rs=oblog.execute("select lockip from oblog_setup")	
%>
<html>
<head>
<title>blog���¹�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<form method="POST" action="Admin_lockip.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr> 
      <td height="22" class="topbg"> <strong>IP���ƹ���</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">��������Ӷ������IP��ÿ��IP�ûس��ָ�������IP����д��ʽ��202.152.12.1��������202.152.12.1���IP�ķ��ʣ���202.152.12.*����������202.152.12��ͷ��IP���ʣ�ͬ��*.*.*.*������������IP�ķ��ʡ�����Ӷ��IP��ʱ����ע�����һ��IP�ĺ��治Ҫ�ӻس�</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <textarea name="lockip" cols="35" rows="20" id="lockip">
<%=rs("lockip")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig"> <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " > 
      </td>
    </tr>
  </table>

</form>
</body>
</html>
<%
set rs=nothing
end sub

sub saveconfig()
	dim rs,sql,log_badstr
	if not IsObject(conn) then link_database
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_setup"
	rs.open sql,conn,1,3
	rs("lockip")=trim(request("lockip"))
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	response.Redirect "admin_lockip.asp"
end sub
%>