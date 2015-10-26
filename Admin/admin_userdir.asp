<!--#include file="inc/inc_sys.asp"-->
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css"> 
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<%
dim action
action=request.QueryString("action")
select case action
	case "adduserdir"
		call adduserdir()
	case "del"
		call deluserdir()
	case "setdefault"
		call setdefault()
	case "domain"
		call domain()
	case "savedomain"
		call savedomain
	case else
		call main
end select

sub adduserdir()
	dim userdir,rs
	userdir=oblog.filt_badstr(Trim(Request.Form("userdir")))
	If userdir="" Then
		Response.Write("<script language=javascript>alert('目录不能为空！');window.location.replace('admin_userdir.asp')</script>")
		Response.End()
	End If	 
	oblog.execute("insert into [oblog_userdir] (userdir,is_default) values ('"&userdir&"',0)")
	'oblog.reloaduserdir
    response.Redirect "admin_userdir.asp"
end sub

sub deluserdir
	dim id
	id=clng(request.QueryString("id"))
	oblog.execute("delete  from [oblog_userdir] where id="&id)
	'oblog.reloaduserdir
    response.Redirect "admin_userdir.asp"
end sub

sub setdefault
	dim id,rs
	id=clng(request.QueryString("id"))
	oblog.execute("update [oblog_userdir] set is_default=0")
	oblog.execute("update [oblog_userdir] set is_default=1 where id="&id)
	set rs=oblog.execute("select userdir from oblog_userdir where is_default=1")
	oblog.execute("update oblog_setup set default_userdir='"&rs(0)&"'")
	set rs=nothing
	oblog.reloadsetup
	'oblog.reloaduserdir
    response.Redirect "admin_userdir.asp"
end sub



sub main()
%>
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr valign="middle"> 
    <td height="21" colspan="3" class="topbg"> <strong>添加用户目录</strong></td>
  </tr>
   <tr class="tdbg"><form name="form1" method="post" action="admin_userdir.asp?action=adduserdir">
          <td height="20" colspan="3"><div align="center">目录名： 
          <input name="userdir" type="text" id="userdir" maxlength="20">
          <input type="submit" name="Submit" value=" 添加 ">			             
          </div></td></form>
  </tr>
</table>
		
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr class="topbg" height="25"> 
    <td width="38" > <div align="center">ID</div></td>
    <td width="109"> <div align="center">目录名</div></td>
    <td width="72"> <div align="center">用户数</div></td>
    <td width="110"><div align="center">当前使用目录</div></td>
    <td width="242"> <div align="center">管理操作</div></td>
  </tr>
  <%
dim rs,rstmp
set rs=oblog.execute("select * from oblog_userdir order by id")
while not rs.eof
%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td > <div align="center"><%=rs("id")%></div></td>
    <td><div align="center"><%=rs("userdir")%></div></td>
    <td><div align="center"> 
        <%
set rstmp=oblog.execute("select count(userid) from oblog_user where user_dir='"&rs("userdir")&"'")
response.Write(rstmp(0))
%>
      </div></td>
    <td> <div align="center"> 
        <%if rs("is_default")=1 then response.Write "<font color=red>是</font>" else response.Write("否")%>
      </div></td>
    <td><div align="center"><a href="admin_userdir.asp?action=setdefault&id=<%=rs("id")%>" <%="onClick='return confirm(""确认此目录为用户默认目录吗？"");'"%>>设置为默认</a> 
        | <a href="admin_userdir.asp?action=del&id=<%=rs("id")%>" <%="onClick='return confirm(""删除后，此成员所有文章将不会在你的blog中显示,确定要删除吗？"");'"%>>删除</a></div></td>
  </tr>
  <%
rs.movenext
wend 
%>
</table>
<br>
<%
end sub
sub domain()
dim id,rs
id=clng(request("id"))
set rs=conn.execute("select * from oblog_userdir where id="&id)
%>
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr valign="middle"> 
    <td height="21" colspan="3" class="topbg"> <strong>设置用户目录绑定域名</strong></td>
  </tr>
   <tr class="tdbg"><form name="form1" method="post" action="admin_userdir.asp?action=savedomain&id=<%=id%>">
          <td height="20" colspan="3"><div align="center"> 
          <p>目录名：<%=rs("userdir")%> 　 绑定域名路径： 
            <input name="domain" type="text" id="domain" size="50" maxlength="80" value="<%=rs("dirdomain")%>">
            <font color="#FF0000">不能省略http://和最后的/ </font><br>
            <br>
            <input type="submit" name="Submit" value=" 确认 ">
          </p>
          </div></td></form>
  </tr>
</table>
<%
end sub
sub savedomain()
dim id,domain
domain=oblog.filt_badstr(Trim(Request.Form("domain")))
id=clng(request("id"))
oblog.execute("update [oblog_userdir] set dirdomain='"&domain&"' where id="&id)
oblog.reloaduserdir
response.Redirect("admin_userdir.asp")
end sub
%>
</body>

