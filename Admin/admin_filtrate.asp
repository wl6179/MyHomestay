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
dim rs,badstr1,badstr2,badstr
set rs=oblog.execute("select * from oblog_setup")
badstr=rs("log_badstr")
if instr(badstr,"$$$") then
	badstr=split(badstr,"$$$")
	badstr1=badstr(0)
	badstr2=badstr(1)		
else
	badstr1=badstr
	badstr2=""
end if
%>
<html>
<head>
<title>blog���¹�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<form method="POST" action="Admin_filtrate.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr> 
      <td height="22" class="topbg"> <strong>���������ֹ��ˣ�</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">�˴�����Ӱ�쵽���£����ۣ�ר�����֡���ʵ�����֡�ģ�����õĹ��ˡ������ַ������������а��������ַ�������(�ԡ��Ŵ���)��������Ҫ���˵��ַ������룬����ж���ַ��������ûس��ָ�����</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <textarea name="badstr1" cols="35" rows="10" id="badstr1">
<%=badstr1%></textarea>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg">&nbsp;</td>
    </tr>
	    <tr> 
      <td height="22" class="topbg"> <strong>��ֹ�����Ĺؼ��֣�</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">�����ۣ����£������к������¹ؼ��ֽ�����ֹ������������Ҫ��ֹ�������ַ������룬����ж���ַ��������ûس��ָ�����<br>
        �������������ַ����ֹ�������ۡ�</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <textarea name="badstr2" cols="35" rows="10" id="badstr2">
<%=badstr2%></textarea>      </td>
    </tr>

    <tr> 
      <td height="25" class="tdbg">&nbsp; </td>
    </tr>
    <tr> 
      <td height="25" class="topbg"><strong>ע������ַ�</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <p>ע������ַ����������û�ע����������ַ������ݣ�������Ҫ���˵��ַ������룬����ж���ַ��������ûس�������<br>
        </p></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <p class="tdbg"> 
          <textarea name="reg_badstr" cols="35" rows="10" id="reg_badstr"><%=rs("reg_badstr")%></textarea>
        </p></td>
    </tr>
    <tr> 
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " > </td>
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
	rs("log_badstr")=trim(request("badstr1"))&"$$$"&trim(request("badstr2"))
	rs("reg_badstr")=trim(request("reg_badstr"))
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	response.Redirect "admin_filtrate.asp"
end sub
%>