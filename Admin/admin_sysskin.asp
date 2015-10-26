<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<title>系统模版管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<%
dim action
Action=trim(request("Action"))

select case Action
	case "saveconfig" 
		call saveconfig()
	case "showskin"
		call showskin()
	case "modiskin"
		call modiskin()
	case "savedefault"
		call savedefault()
	case "delconfig"
		call delconfig()
	case "addskin"
		call addskin()
	case "saveaddskin"
		call saveaddskin()	
end select

sub showskin()
dim rs
set rs=oblog.execute("select * from oblog_sysskin ")
%>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>系统模版管理</strong></div>
  </tr>
</table>
<form name="form2" method="post" action="admin_sysskin.asp?action=savedefault">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
  <tr class="topbg"> 
      <td width="10%" height="25" >
<div align="center">ID</div></td>
    <td width="20%" > <div align="center">名称</div></td>
    <td width="20%" ><div align="center">作者</div></td>
      <td width="10%" >
<div align="center">默认模版</div></td>
      <td width="40%" > 
        <div align="center">模版管理</div></td>
  </tr>

 <% 
while not rs.eof	  
%> 
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
          <td width="10%"> 
              <div align="center"><%= rs("id") %>&nbsp;</div></td>
          <td width="20%" > <div align="center"><%= rs("sysskinname") %></div></td>
          <td width="20%" ><div align="center"><%= rs("skinauthor") %></div></td>
            <td width="10%">              <div align="center"> 
                <input name="radiobutton" type="radio" class="tdbg" value='<%=rs("id")%>' <%if rs("isdefault")=1 then response.Write "checked" %>>
            </div></td>
            <td width="40%"> 
                <div align="center">
				<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">修改主模版</a> 
                        　<a href="../admin_edit.asp?action=modiskin&skintype=sys&t=0&editm=1&skinorder=1&id=<%=rs("id")%>" target="_blank">修改副模版</a>
				<a href="admin_sysskin.asp?action=modiskin&id=<%=rs("id")%>">修改模版(文本方式)</a> 
          　<a href="admin_sysskin.asp?action=delconfig&id=<%=rs("id")%>" onclick=return(confirm("确定要删除这个模版吗？"))>删除模版</a></div></td>
        </tr>
      
      <%
rs.movenext
wend
%>   
    <tr> 
    <td height="40" colspan="5" align="center" class="tdbg"> <div align="center"> 
          <input type="submit" name="Submit" value="保存设置">
        </div></td>
  </tr>
</table>
</form>
<%
	set rs=nothing
end sub

sub saveconfig()
	dim rs,sql
	if trim(request("sysskinname"))="" then oblog.sys_err("模版名不能为空"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("主摸版不能为空"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("副摸版不能为空"):response.End()
	if trim(request("skinshowlog_3"))="" then oblog.sys_err("第三摸版涉及到各种.asp功能页面模板，请不要为空"):response.End()
	if trim(request("skinshowlog_4"))="" then oblog.sys_err("第四摸版专门涉及到登录模块的特殊处理，请不要为空"):response.End()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_sysskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("sysskinname")=trim(request("sysskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=request("skinmain")
	rs("skinshowlog")=request("skinshowlog")
	rs("skinshowlog_3")=request("skinshowlog_3")
	rs("skinshowlog_4")=request("skinshowlog_4")
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	Application(oblog.cache_name&"index_update")=true
	oblog.showok "修改成功",""
end sub
sub savedefault()
	dim isdefaultID
	isdefaultID=clng(trim(request("radiobutton")))
	oblog.execute("update oblog_sysskin set isdefault=0")
	oblog.execute("update oblog_sysskin set isdefault=1 where id="&isdefaultID)
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub delconfig()
	oblog.execute("delete from oblog_sysskin where id="&clng(request.QueryString("id")))
	response.Redirect "admin_sysskin.asp?action=showskin"
end sub

sub saveaddskin()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if trim(request("sysskinname"))="" then oblog.sys_err("模版名不能为空"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("主模版不能为空"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("副模版不能为空"):response.End()
	sql="select * from oblog_sysskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("sysskinname")=trim(request("sysskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=trim(request("skinmain"))
	rs("skinshowlog")=trim(request("skinshowlog"))
	rs.update
	rs.close
	set rs=nothing
	response.Redirect "admin_sysskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_sysskin where id="&clng(request.QueryString("id")))
%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><strong>修改系统模版</strong></td>
    </tr>
    <tr class="tdbg"> 
      
    <td width="253" height="30"><strong>现在修改的模版是：<%=rs("sysskinname")%></strong></td>
      
    <td width="516" height="30">
　　<a href="admin_sysskin.asp?action=showskin">返回管理菜单</a>　　 
      <a href="admin_skin_help.asp" target="_blank"><strong>模版标记帮助</strong></a>
	 </td>
    </tr>
</table>

<form method="POST" action="admin_sysskin.asp?id=<%=clng(request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>模版参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模版名称： 
        <input name="sysskinname" type="text" id="sysskinname" value=<%=rs("sysskinname")%>>
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <strong>主模版：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then response.Write Server.HtmlEncode(rs("skinmain")) else response.Write("")%></textarea>
        <br>
        <br>
        <strong>副模版： </strong><br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then response.Write Server.HtmlEncode(rs("skinshowlog")) else response.Write("")%></textarea>
   		<br>
        <br>
		<strong>第三模版： </strong><br>
        <textarea name="skinshowlog_3" cols="100" rows="12" id="skinshowlog_3"><%if rs("skinshowlog_3")<>"" then response.Write Server.HtmlEncode(rs("skinshowlog_3")) else response.Write("")%></textarea>
		<br>
        <br>
		<strong>第四模版（特殊用途模板，比如reg项不需要登录检测处理，则可以使用此模板优化~WL）： </strong><br>
        <textarea name="skinshowlog_4" cols="100" rows="12" id="skinshowlog_4"><%if rs("skinshowlog_4")<>"" then response.Write Server.HtmlEncode(rs("skinshowlog_4")) else response.Write("")%></textarea>
		</td>
    </tr>
    <tr> 
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " > 
      </div></td>
    </tr>
  </table>

</form>

<%
set rs=nothing
end sub

sub addskin()
%>
  
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="topbg"> 
    <td height="22" align=center><strong>添加系统模版</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td height="30"><div align="center">
	<a href="admin_sysskin.asp?action=showskin"><strong>返回管理菜单</strong></a>　　 
	<a href="admin_skin_help.asp" target="_blank"><strong>模版标记帮助</strong></a>
	</div></td>
  </tr>
</table>

<form method="POST" action="admin_sysskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>模版参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模版名称： 
        <input name="sysskinname" type="text" id="sysskinname">
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor"></td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <strong>主模版：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>副模版： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr> 
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin"> 
          <input name="cmdadd" type="submit" id="cmdadd" value=" 添加 " > 
      </div></td>
    </tr>
  </table>

</form>

<%
end sub%>