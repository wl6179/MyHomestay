<!--#include file="inc/inc_sys.asp"-->
<%
dim Action,FoundErr,ErrMsg
dim rs,sql
Action=trim(request("Action"))
sql="Select * from oblog_Admin where UserName='" & Admin_Name & "'"
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open sql,conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
else
	if Action="Modify" then
		call ModifyPwd()
	else
		call main()
	end if
end if
rs.close
set rs=nothing
if FoundErr=True then
	call WriteErrMsg()
end if

sub ModifyPwd()
	dim password,PwdConfirm,username
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
		exit sub
	end if
	UserName=rs("UserName")
	if Password<>"" then
		rs("password")=md5(password)
	end if
   	rs.update
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""�����޸ĳɹ��������µ�½��"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub main()
%>
<html>
<head>
<title>�޸Ĺ���Ա����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<script language=javascript>
function check()
{
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}
</script>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body class="bgcolor">
<form method="post" action="admin_adminmodifypwd.asp" name="form1" onSubmit="javascript:return check();">
  <br>
  <br>
  <table width="300" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center" class="style1">�� �� �� �� Ա �� ��</div></td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg">�� �� ����</td>
      <td class="tdbg"><%=rs("UserName")%></td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg">�� �� �룺</td>
      <td class="tdbg"><input type="password" name="Password"> </td>
    </tr>
    <tr> 
      <td width="100" align="right" class="tdbg">ȷ�����룺</td>
      <td class="tdbg"><input type="password" name="PwdConfirm"> </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input  type="submit" name="Submit" value=" ȷ �� " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="reset()" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
end sub

sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub
%>
