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
<title>blog文章过滤设置</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<form method="POST" action="Admin_filtrate.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr> 
      <td height="22" class="topbg"> <strong>文章敏感字过滤：</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">此处设置影响到文章，评论，专题名字、真实姓名字、模版设置的过滤。过滤字符将过滤内容中包含以下字符的内容(以×号代替)，请您将要过滤的字符串添入，如果有多个字符串，请用回车分隔开。</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <textarea name="badstr1" cols="35" rows="10" id="badstr1">
<%=badstr1%></textarea>      </td>
    </tr>
    <tr>
      <td height="25" class="tdbg">&nbsp;</td>
    </tr>
	    <tr> 
      <td height="22" class="topbg"> <strong>禁止发布的关键字：</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg">当评论，文章，留言中含有以下关键字将被禁止发布，请您将要禁止发布的字符串添入，如果有多个字符串，请用回车分隔开。<br>
        可以输入恶意网址来禁止垃圾评论。</td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <textarea name="badstr2" cols="35" rows="10" id="badstr2">
<%=badstr2%></textarea>      </td>
    </tr>

    <tr> 
      <td height="25" class="tdbg">&nbsp; </td>
    </tr>
    <tr> 
      <td height="25" class="topbg"><strong>注册过滤字符</strong></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <p>注册过滤字符将不允许用户注册包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用回车隔开。<br>
        </p></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <p class="tdbg"> 
          <textarea name="reg_badstr" cols="35" rows="10" id="reg_badstr"><%=rs("reg_badstr")%></textarea>
        </p></td>
    </tr>
    <tr> 
      <td height="40" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="saveconfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存设置 " > </td>
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