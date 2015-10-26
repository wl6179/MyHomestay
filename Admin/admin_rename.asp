<!--#include file="inc/inc_sys.asp"-->
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<%
dim action,rs,newname,old,username,sql
action=oblog.filt_badstr(request("action"))
newname=oblog.filt_badstr(trim(request("newname")))
old=oblog.filt_badstr(trim(request("old")))
	if old<>"" then
		set rs=oblog.execute("select userid from oblog_user where username='"&oblog.filt_badstr(old) & "'")
		if rs.eof then 
		Response.Write("<script language=javascript>alert('原用户名不存在！');history.back();</script>")
		Response.End()
	end if
	end if
	if newname<>"" then
		set rs=oblog.execute("select userid from oblog_user where username='"&oblog.filt_badstr(newname) & "'")
		if not rs.eof then 
		Response.Write("<script language=javascript>alert('新用户名已经存在！');history.back();</script>")
		Response.End()
    end if
	end if
		if Instr(newname,"=")>0 or Instr(newname,"%")>0 or Instr(newname,chr(32))>0 or Instr(newname,"?")>0 or Instr(newname,"&")>0 or Instr(newname,";")>0 or Instr(newname,",")>0 or Instr(newname,"'")>0 or Instr(newname,",")>0 or Instr(newname,chr(34))>0 or Instr(newname,chr(9))>0 or Instr(newname,"")>0 or Instr(newname,"$")>0 then 
		Response.Write("<script language=javascript>alert('用户名含有非法字符！');history.back();</script>")
		Response.End()
	end if
if action="DoUpdate" then
	If newname="" Then
		Response.Write("<script language=javascript>alert('新用户名不能为空！');history.back();</script>")
		Response.End()
	End If
	If old="" Then
		Response.Write("<script language=javascript>alert('原用户名不能为空！');history.back();</script>")
		Response.End()
	End If
		If old=newname Then
		Response.Write("<script language=javascript>alert('原用户名与新用户名相同！');history.back();</script>")
		Response.End()
	End If
	oblog.execute("update [oblog_user] set username='"&newname&"' where username='"&old&"'")
	oblog.execute("update [oblog_log] set author='"&newname&"' where author='"&old&"'")
	oblog.execute("update [oblog_comment] set comment_user='"&newname&"' where isguest=0 and comment_user='"&old&"'")
	oblog.execute("update [oblog_message] set message_user='"&newname&"' where isguest=0 and  message_user='"&old&"'")
	oblog.execute("update [oblog_pm] set sender='"&newname&"' where sender='"&old&"'")
	oblog.execute("update [oblog_pm] set incept='"&newname&"' where incept='"&old&"'")
	response.Write("<br>已经成功将用户名进行了更改！")
	else
	%>
	
<title>用户改名</title><form name="form1" method="post" action="">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center" class="title">
    <td height="22" colspan="2"><strong><font color="#FFFFFF">更 新 系 统 数 据</font></strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><p>说明：
        本操作将更改用户名，请慎重操作。</p>
      <p>原用户名：
        <input name="old" type="text" id="old">
         <br>
         新用户名：
         <input name="newname" type="text" id="newname">
         用户名禁止特殊符号<br>
        </p></td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="改名">
    <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
  </tr>
</table>
   </form>
<%
end if
%>

