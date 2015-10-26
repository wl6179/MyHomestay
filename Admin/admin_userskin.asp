<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<title>用户模版管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form2.chkAll.checked){
	document.form2.chkAll.checked = document.form2.chkAll.checked&0;
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
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages,strGuide,ispass
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
	case "passskin"
		call passskin
	case "unpassskin"
		call unpassskin	
end select

sub showskin()
	dim rs,psql,sql,strFileName
	ispass=clng(request("ispass"))
	if ispass=1 then
	    strFileName="admin_userskin.asp?action=showskin&ispass=1"
		psql=" where ispass=1 "
	else
	    strFileName="admin_userskin.asp?action=showskin&ispass=0"
		psql=" where ispass=0 "
	end if
	if request("page")<>"" then
	    currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select id,userskinname,skinauthor,skinauthorurl,isdefault,ispass,skinpic from oblog_userskin "&psql&" order by id desc "
		rs.Open sql,Conn,1,1
	  	if rs.eof and rs.bof then
            showContent(rs)
			strGuide=strGuide & " (共有0个模板)</h1>"
			response.write "<div align='right'>"&strGuide&"</div>"
		else
	    	totalPut=rs.recordcount
			strGuide=strGuide & " (共有" & totalPut & "个模板)</h1>"
			if currentpage<1 then
	       		currentpage=1
	    	end if
	    	if (currentpage-1)*MaxPerPage>totalput then
		   		if (totalPut mod MaxPerPage)=0 then
		     		currentpage= totalPut \ MaxPerPage
			  	else
			      	currentpage= totalPut \ MaxPerPage + 1
		   		end if
	    	end if
		    if currentPage=1 then
	        	Call showContent(rs)
	        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个模板")
	   	 	else
	   	     	if (currentPage-1)*MaxPerPage<totalPut then
	         	   	rs.move  (currentPage-1)*MaxPerPage
	         		dim bookmark
	           		bookmark=rs.bookmark            	
	        	else
		        	currentPage=1           		           		
		    	end if
		    	Call showContent(rs)
		    	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个模板")
			end if
		end if
		rs.Close
		set rs=Nothing
end sub

sub showContent(rs)
	dim i 
	i=0
	
%>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td  height=25 class="topbg" align="left"><strong>用户模版管理　　　<a href="admin_userskin.asp?action=showskin&ispass=1">&gt;&gt;通过审核的模板</a>　　<a href="admin_userskin.asp?action=showskin&ispass=0">&gt;&gt;未通过审核的模版</a></strong>
  </tr>
</table>
<form name="form2" method="post" action="admin_userskin.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg"> 
      <td height="25" colspan="6" ><strong><%if ispass=1 then response.Write "通过审核的模板" else response.write "未通过审核的模版"%></strong></td>
    </tr>
    <tr class="topbg"> 
      <td width="10%" height="25" > <div align="center">ID</div></td>
      <td width="20%" > <div align="center">名称<b>（红色为默认模板）</b></div></td>
      <td width="20%" ><div align="center">作者</div></td>
	  <td width="10%" > <div align="center">审核</div></td>
      <td width="10%" > <div align="center">选中</div></td>
      <td width="40%" > <div align="center">模版管理</div></td>
    </tr>
    <% 
do while not rs.eof	  
dim userskinname
    userskinname=rs("userskinname")
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="10%"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="20%" > <div align="center">
	  <%if rs("isdefault")=1 then
	  response.Write "<font color='red'>"&userskinname&"</red>"
	  else 
	  response.Write userskinname
	  end if
	  %>
	  </div></td>
      <td width="20%" ><div align="center">
	  <%if rs("skinauthorurl")="" or isnull(rs("skinauthorurl")) then
	  response.Write rs("skinauthor") 
	  else
	  response.Write "<a href="""&oblog.filt_html(rs("skinauthorurl"))&""" target='_blank'>"&rs("skinauthor")&"</a>" 
	  end if%>
	  </div></td>
	  <td width="10%" > <div align="center"><%if rs("ispass")=1 then response.Write("已审核") else response.Write("未审核")%></div></td>
      <td width="10%"> <div align="center"> 
          <input name="checkbox" type="checkbox" onClick="unselectall()" id= "checkbox" class="tdbg" value='<%=rs("id")%>'>
        </div></td>
      <td width="40%"> <div align="left">
	<a href="../showskin.asp?id=<%=rs("id")%>" target="_blank">预览</a> 
	<%if ispass=0 then%>
	<a href="admin_userskin.asp?action=passskin&id=<%=rs("id")%>">通过审核</a> 
	<%else%>
	<a href="admin_userskin.asp?action=unpassskin&id=<%=rs("id")%>">取消审核</a> 
	<%end if%>
  <a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">修改主模版</a>
  <a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=1&id=<%=rs("id")%>"  target="_blank">修改副模版</a><br>
	<a href="admin_userskin.asp?action=modiskin&id=<%=rs("id")%>">修改模版(文本方式)</a> 
	　<a href="admin_userskin.asp?action=delconfig&id=<%=rs("id")%>" onclick=return(confirm("确定要删除这个模版吗？"))>删除模版</a></div></td>
    </tr>
    <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>

    <tr> 
      <td height="40" colspan="6" align="center" class="tdbg"> <div align="center"> 
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
	  全选
	 <input type="radio" value="savedefault" name="action" checked>默认模板</option>
	 <%if ispass=0 then%>
	  <input type="radio" value="passskin" name="action" >通过审核</option>
	  <%else%>
	  <input type="radio" value="unpassskin" name="action">取消审核</option>
	  <%end if%>
	   <input type="radio" value="delconfig" name="action" >删除</option>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  <input type="submit" name="Submit" value="保存设置">
        </div></td>
    </tr>
  </table>
</form>
<%
end sub

sub savedefault()
	dim isdefaultID
	isdefaultID=trim(request("checkbox"))
		if instr(isdefaultID,",")>0 then
		Response.Write("<script language=javascript>alert('用户默认模板只可以选择一个！');history.back();</script>")
		Response.End()
	elseif isdefaultID="" then
		Response.Write("<script language=javascript>alert('请指定要设定为默认的模板！');history.back();</script>")
		Response.End()
		exit sub
		end if
	oblog.execute("update oblog_userskin set isdefault=0")
	oblog.execute("update oblog_userskin set isdefault=1 where id="&isdefaultID)
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""修改成功！"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub passskin()
	dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=1 where id in ("&id&")")
	elseif id="" then
	id=clng(request("id"))
	oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	else
    oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	end if
	oblog.showok "通过审核成功",""
end sub

sub unpassskin()
	dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=0 where id in ("&id&")")
	elseif id="" then
	id=clng(request("id"))
	oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	else
	oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	end if
	oblog.showok "取消审核成功",""
end sub

sub saveconfig()
	dim rs,sql
	if trim(request("userskinname"))="" then oblog.sys_err("模版名不能为空"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("主模版不能为空"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("副模版不能为空"):response.End()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_userskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("userskinname")=trim(request("userskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=request("skinmain")
	rs("skinshowlog")=request("skinshowlog")
	rs("skinpic")=trim(request("skinpic"))
	rs("skinauthorurl")=trim(request("skinauthorurl"))
	
	rs("skinmain_Nav")=request("skinmain_Nav")
	
	
	rs.update
	rs.close
	set rs=nothing
	oblog.showok "保存成功",""
end Sub
sub delconfig()
    dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("delete from oblog_userskin where id in ("&id&")")
	elseif id="" then
	id=clng(request.QueryString("id"))
	oblog.execute("delete from oblog_userskin where id="&id)
	else
	oblog.execute("delete from oblog_userskin where id="&id)
	end if
	oblog.showok "删除成功",""
end sub
sub modiconfig()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&clng(request.QueryString("id")))
	End Sub 
sub saveaddskin()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if trim(request("userskinname"))="" then oblog.sys_err("模版名不能为空"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("主模版不能为空"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("副模版不能为空"):response.End()
	sql="select * from oblog_userskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("userskinname")=trim(request("userskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=trim(request("skinmain"))
	rs("skinshowlog")=trim(request("skinshowlog"))
	rs("skinpic")=trim(request("skinpic"))
	rs("skinauthorurl")=trim(request("skinauthorurl"))
	rs.update
	rs.close
	set rs=nothing
	response.Redirect "admin_userskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&clng(request.QueryString("id")))
%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><strong>修改用户模版</strong></td>
    </tr>
    <tr class="tdbg"> 
      
    <td width="253" height="30"><strong>现在修改的模版是：<%=rs("userskinname")%></strong></td>
      
    <td width="516" height="30">
	<a href="admin_userskin.asp?action=modiskin&id=<%=rs("id")%>">修改模版</a>　　<a href="admin_userskin.asp?action=showskin&ispass=1">返回管理菜单</a>　　 
      <a href="admin_skin_help.asp" target="_blank"><strong>模版标记帮助</strong></a></td>
    </tr>
</table>

<form method="POST" action="admin_userskin.asp?id=<%=clng(request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>修改模版</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模版名称： 
        <input name="userskinname" type="text" id="userskinname" value=<%=rs("userskinname")%>>
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
        <br>
        模板作者连接： 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        模版预览图片<strong>： 
        <input name="skinpic" type="text" id="skinpic" size="50" value="<%=rs("skinpic")%>">
        </td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <strong>主模版：</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then response.Write Server.HtmlEncode(rs("skinmain")) else response.Write("")%></textarea>
        <br>
        <br>
        <strong>副模版： <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then response.Write Server.HtmlEncode(rs("skinshowlog")) else response.Write("")%></textarea>
        </strong>
		<br>
        <br>
        <strong>栏目级模版： <br>
        <textarea cols="60" rows="20" name="skinmain_Nav" id="edit"><%=Server.HtmlEncode(rs("skinmain_Nav"))%></textarea>
        </strong>
		</td>
    </tr>
		
	 
    <tr> 
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig"> <br />
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " > &nbsp;<br />&nbsp;
      </div>
	  </td>
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
    <td height="22" align=center><strong>添加用户模版</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td height="30"><div align="center"><a href="admin_userskin.asp?action=showskin"><strong>返回管理菜单</strong></a>　　 <a href="admin_skin_help.asp" target="_blank"><strong>模版标记帮助</strong></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	 <a href="admin_adduserskin.asp?add=first"><strong>2.52模式</strong></a>
	</div></td>
  </tr>
</table>

<form method="POST" action="admin_userskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>模版参数</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">模版名称： 
        <input name="userskinname" type="text" id="userskinname">
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor">
        <br>
        模板作者连接<strong>： 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="">
        </strong> <br>
        模版预览图片<strong>： 
        <input name="skinpic" type="text" id="skinpic" size="50">
        </strong> </td>
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
end sub
%>