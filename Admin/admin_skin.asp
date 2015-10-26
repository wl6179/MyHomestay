<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<title>系统模版管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form.chkAll.checked){
	document.form.chkAll.checked = document.form.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
<%
dim action
Action=trim(request("Action"))

select case Action
	case "outuser" 
		call outuser()
	case "outuserok"
		call outuserok()
	case "inuser1"
		call inuser1()
	case "inuser2"
		call inuser2()
	case "inuserok"
		call inuserok()
	case "outsys" 
		call outsys()
	case "outsysok"
		call outsysok()
	case "insys1"
		call insys1()
	case "insys2"
		call insys2()
	case "insysok"
		call insysok()
end select

sub outuserok()
	dim mdbname,rs,connskin,fso
	dim skinid,i,rsout
	mdbname=trim(request("mdbname"))	
	set fso=Server.CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(Server.MapPath(mdbname))=False then
    Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
    Response.End
    end if
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsout=server.createobject("adodb.recordset")
	rsout.open "select * from skin",connskin,2,3
	skinid=split(request("id"))
	for i=0 to ubound(skinid)
		set rs=oblog.execute("select * from oblog_userskin where id="&clng(skinid(i)))
		rsout.addnew
		rsout("type")="user"
		rsout("skinname")=rs("userskinname")
		rsout("skinmain")=rs("skinmain")
		rsout("skinshowlog")=rs("skinshowlog")
		rsout("skinauthor")=rs("skinauthor")
		rsout("skinauthorurl")=rs("skinauthorurl")
		rsout("skinpic")=rs("skinpic")
		rsout.update
	next
	rsout.close
	set rsout=nothing
	set rs=nothing
	response.Write("导出成功！")
end sub

sub inuserok()
	dim mdbname,rs,connskin
	dim skinid,i,rsin
	mdbname=trim(request("mdbname"))	
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsin=server.createobject("adodb.recordset")
	if not IsObject(conn) then link_database
	rsin.open "select top 1 * from oblog_userskin",conn,2,3
	skinid=split(request("id"))
	for i=0 to ubound(skinid)
		set rs=connskin.execute("select * from skin where type='user' and id="&clng(skinid(i)))
		rsin.addnew
		rsin("userskinname")=rs("skinname")
		rsin("skinmain")=rs("skinmain")
		rsin("skinshowlog")=rs("skinshowlog")
		rsin("skinauthor")=rs("skinauthor")
		rsin("skinauthorurl")=rs("skinauthorurl")
		rsin("skinpic")=rs("skinpic")
		rsin.update
	next
	rsin.close
	set rsin=nothing
	set rs=nothing
	response.Write("导入成功！")
end sub

sub outsysok()
	dim mdbname,rs,connskin,fso
	dim skinid,i,rsout
	mdbname=trim(request("mdbname"))
	set fso=Server.CreateObject("Scripting.FileSystemObject")
    if (fso.FileExists(Server.MapPath(mdbname))) Then
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsout=server.createobject("adodb.recordset")
	rsout.open "select top 1 * from skin",connskin,2,3
	skinid=split(request("id"))
	for i=0 to ubound(skinid)
		set rs=oblog.execute("select * from oblog_sysskin where id="&clng(skinid(i)))
		rsout.addnew
		rsout("type")="sys"
		rsout("skinname")=rs("sysskinname")
		rsout("skinmain")=rs("skinmain")
		rsout("skinshowlog")=rs("skinshowlog")
		rsout("skinauthor")=rs("skinauthor")
		rsout.update
	next
	rsout.close
	set rsout=nothing
	set rs=nothing
	response.Write("导出成功！")
	 Else
    Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
    Response.End
	End if
end sub

sub insysok()
	dim mdbname,rs,connskin
	dim skinid,i,rsin
	mdbname=trim(request("mdbname"))	
	Set connskin = Server.CreateObject("ADODB.Connection")
	connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
	set rsin=server.createobject("adodb.recordset")
	if not IsObject(conn) then link_database
	rsin.open "select * from oblog_sysskin",conn,2,3
	skinid=split(request("id"))
	for i=0 to ubound(skinid)
		set rs=connskin.execute("select * from skin where type='sys' and id="&clng(skinid(i)))
		rsin.addnew
		rsin("sysskinname")=rs("skinname")
		rsin("skinmain")=rs("skinmain")
		rsin("skinshowlog")=rs("skinshowlog")
		rsin("skinauthor")=rs("skinauthor")
		rsin.update
	next
	rsin.close
	set rsin=nothing
	set rs=nothing
	response.Write("导入成功！")
end sub
%>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<%
sub outuser()
dim rs
set rs=oblog.execute("select * from oblog_userskin ")
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>用户模版导出</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=outuserok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <% 
while not rs.eof	  
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("userskinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr> 
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left"> 
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模版　导出数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 导出模版 ">
        </div></td>
    </tr>
  </table>
</form>
<%end sub

sub inuser1()
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>用户模版导入</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=inuser2">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
     <tr> 
      <td width="763" height="40" align="center" class="tdbg"> <div align="left"> 
          　导入数据库名： 
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 下一步 ">
        </div></td>
    </tr>
  </table>
</form>
<%end sub

sub inuser2()
dim connskin,rs,mdbname,fso
mdbname=trim(request("mdbname"))
	set fso=Server.CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(Server.MapPath(mdbname))=False then
    Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
    Response.End
    end if
Set connskin = Server.CreateObject("ADODB.Connection")
connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
set rs=connskin.execute("select * from skin where type='user'")
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>用户模版导入</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=inuserok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <% 
while not rs.eof	  
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("skinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr> 
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left"> 
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模版　 
          <input type="submit" name="Submit" value=" 导入模版 ">
          <input type="hidden" name="mdbname" value="<%=mdbname%>">
        </div></td>
    </tr>
  </table>
</form>
<%end sub%>
<%
sub outsys()
dim rs
set rs=oblog.execute("select * from oblog_sysskin ")
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>系统模版导出</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=outsysok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <% 
while not rs.eof	  
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("sysskinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr> 
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left"> 
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模版　导出数据库名：
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 导出模版 ">
        </div></td>
    </tr>
  </table>
</form>
<%end sub

sub insys1()
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>系统模版导入</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=insys2">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
     <tr> 
      <td width="763" height="40" align="center" class="tdbg"> <div align="left"> 
          　导入数据库名： 
          <input name="mdbname" type="text" id="mdbname" value="../skin/skin.mdb" size="20" maxlength="50">
          <input type="submit" name="Submit" value=" 下一步 ">
        </div></td>
    </tr>
  </table>
</form>
<%end sub

sub insys2()
dim connskin,rs,mdbname,fso
mdbname=trim(request("mdbname"))
set fso=Server.CreateObject("Scripting.FileSystemObject")
    if fso.FileExists(Server.MapPath(mdbname))=False then
    Response.Write("<script language=javascript>alert('“"&mdbname&"”不存在！');history.back();</script>")
    Response.End
    end if
Set connskin = Server.CreateObject("ADODB.Connection")
connskin.open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(mdbname)
set rs=connskin.execute("select * from skin where type='sys'")
%>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>系统模版导入</strong></div>
  </tr>
</table>
<form name="form" method="post" action="admin_skin.asp?action=insysok">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg">
      <td width="63" align="center" >选中</td>
      <td width="120"> <div align="center">ID</div></td>
      <td width="288" > <div align="center">名称</div></td>
      <td width="292" > <div align="center">作者</div></td>
    </tr>
    <% 
while not rs.eof	  
%>
    <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'">
      <td width="63" align="center"><input name='ID' type='checkbox' onclick="unselectall()" id="ID" value='<%=cstr(rs("id"))%>'></td>
      <td width="120"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="288" > <div align="center"><%= rs("skinname") %></div></td>
      <td > <div align="center"><%= rs("skinauthor") %></div></td>
    </tr>
    <%
rs.movenext
wend
%>
    <tr> 
      <td height="40" colspan="4" align="center" class="tdbg"> <div align="left"> 
          <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
          选中所有模版　 
          <input type="submit" name="Submit" value=" 导入模版 ">
          <input type="hidden" name="mdbname" value="<%=mdbname%>">
        </div></td>
    </tr>
  </table>
</form>
<%end sub%>

</body>
</html>
