
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
            <td height="38" colspan="2" class="style1"><font size="3"><strong>oBlog��̨����Ա��¼</strong></font> </td>
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
                <img src="/inc/code.asp" style="cursor:hand" onclick="this.src='/inc/code.asp';" alt="������?��һ��" /></td>
          </tr>
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input name="Submit"   type="submit"   value=" ȷ&nbsp;�� ">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" ��&nbsp;�� " >
                <br>
                <br>
                ����֤���޷���ʾ�����޸�config.asp�ļ��е�blogdir·��<br>
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
