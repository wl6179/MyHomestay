<!--#include file="../conn.asp"-->
<!--#include file="../inc/class_sys.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
If request("action")="logout" Then
	session("AdminName")=""
	Response.Redirect "../index.asp"
End If

dim oblog
set oblog=new class_sys
oblog.start
if request("action")<>"login" then
%>
<html>
<head>
<title>����Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/admin/Admin_STYLE.CSS">
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.username.value=="")
	document.Login.username.focus();
else
	document.Login.username.select();
}
function CheckForm()
{
	if(document.Login.username.value=="")
	{
		alert("�������û�����");
		document.Login.username.focus();
		return false;
	}
	if(document.Login.password.value == "")
	{
		alert("���������룡");
		document.Login.password.focus();
		return false;
	}
	if (document.Login.codestr.value==""){
       alert ("������������֤�룡");
       document.Login.GetCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("��ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("��ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>

</head>
<body class="bgcolor">
<p>&nbsp;</p>
<form name="Login" action="administratorLogin.asp?action=login" method="post" target="_parent" onSubmit="return CheckForm();">
    
  <table width="585" border="0" align="center" cellpadding="0" cellspacing="0" >
    <tr> 
      <td width="280" rowspan="2"><div align="right"> </div></td>
      <td width="344" background="Images/entry2.gif"> <table width="100%" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr align="center"> 
            <td height="38" colspan="2" class="style1"><font size="3"><strong>MyHomestay ��̨����Ա��¼</strong></font> </td>
          </tr>
          <tr> 
            <td align="right">�û����ƣ�</td>
            <td><input name="username"  type="text"  id="username" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#D6DFF7';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></td>
          </tr>
          <tr> 
            <td align="right">�û����룺</td>
            <td><input name="password"  type="password" id="password" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#D6DFF7';" onMouseOut="this.style.background='#FFFFFF'" maxlength="20"></td>
          </tr>
          <tr> 
            <td align="right">�� ֤ �룺</td>
            <td><input name="codestr" id="codestr" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onFocus="this.select(); " onMouseOver="this.style.background='#D6DFF7';" onMouseOut="this.style.background='#FFFFFF'" size="6" maxlength="4">
                <%=oblog.getcode%></td>
          </tr>
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input name="Submit"   type="submit"   value=" ȷ&nbsp;�� ">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" ��&nbsp;�� " >
                <br>
                <br>
                ����֤���޷���ʾ�����޸�config.asp<br>
                <strong> </strong> </div></td>
          </tr>
        </table></td>
    </tr>
    <tr><td height="3"></td></tr>
  </table>
  </form>
<script language="JavaScript" type="text/JavaScript">
SetFocus(); 
</script>
</body>
</html>
<%
else
	'�ƺ�������Ӻ�̨�û��ļ���
	Dim level

	dim sql,rs
	dim username,password
	dim founderr,errmsg
	if not oblog.codepass then
		FoundErr=True
		errmsg=errmsg & "<br><li>��֤�����</li>"
	end if
	username=oblog.filt_badstr(trim(request("username")))
	password=trim(request("password"))
	if username="" then
		FoundErr=True
		errmsg=errmsg & "<br><li>�û�������Ϊ�գ�</li>"
	end if
	if password="" then
		FoundErr=True
		errmsg=errmsg & "<br><li>���벻��Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		password=md5(password)
		set rs=server.createobject("adodb.recordset")
		sql="select * from oblog_admin where username='"&username&"'"
		if not IsObject(conn) then link_database
		rs.open sql,conn,1,3
		if rs.bof and rs.eof then
			FoundErr=True
			errmsg=errmsg & "<br><li>�û������������</li>"
		else
			if password<>rs("password") then
				FoundErr=True
				errmsg=errmsg & "<br><li>�û������������</li>"
			else
				rs("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
				rs("LastLoginTime")=now()
				rs("LoginTimes")=rs("LoginTimes")+1
				rs.update
				session.Timeout=60
				session("adminname")=rs("username")
				session("adminpassword")=rs("password")
				
				'�ƺ���������û����𣬸����������ֵ
				level = rs("admin_level")
				
				'''rs.close
				'''set rs=nothing
				
				'�ƺ����������û�������Ӧҳ��
				Select Case level
					Case 0,1,2'WL�����޸ģ�
						Response.Redirect "administrator.asp"
					Case 3
						Response.Redirect "administrator_base.asp"
					Case 4
						Response.Redirect "clinic_index.asp"
					Case Else
						Response.Write "�����ù���Ա����"
						
				End Select
			end if
		end if
		rs.close
		set rs=nothing
	end if
	if founderr=True then
		oblog.sys_err(errmsg)
	end if
end if
%>
