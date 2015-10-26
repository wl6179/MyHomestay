<!--#include file="inc/inc_sys.asp"-->
<%
dim Action,ParentID,i,FoundErr,ErrMsg,t,tName
dim SkinCount,LayoutCount
Action=trim(Request("Action"))
ParentID=trim(request("ParentID"))
t=trim(request("t"))
if ParentID="" then
	ParentID=0
else
	ParentID=CLng(ParentID)
end if

If t="1" Then
	tName="相册"
Else
	t="0"
	tName="文章"
End If
%>
<html>
<head>
<title><%=tName%>分类管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_Style.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style1 {color: #FF6600}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr>
	<td colspan="2" align="center" class="title"><strong><%=tName%>分类管理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="80" height="30"><strong>管理导航：</strong></td>
    <td height="30">
    	<a href="admin_logclass.asp?t=<%=t%>"><%=tName%>分类管理首页</a> | 
    	<a href="admin_logclass.asp?Action=Add&t=<%=t%>">添加<%=tName%>分类</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Order&t=<%=t%>">一级分类排序</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=OrderN&t=<%=t%>">N级分类排序</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Reset&t=<%=t%>">复位所有<%=tName%>分类</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Unite&t=<%=t%>"><%=tName%>分类合并</a></td>
  </tr>
</table><br>

<%
if not IsObject(conn) then link_database
if Action="Add" then
	call AddClass()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Move" then
	call MoveClass()
elseif Action="SaveMove" then
	call SaveMove()
elseif Action="Del" then
	call DeleteClass()
elseif Action="UpOrder" then 
	call UpOrder() 
elseif Action="DownOrder" then 
	call DownOrder() 
elseif Action="Order" then
	call Order()
elseif Action="UpOrderN" then 
	call UpOrderN() 
elseif Action="DownOrderN" then 
	call DownOrderN() 
elseif Action="OrderN" then
	call OrderN()
elseif Action="Reset" then
	call Reset()
elseif Action="SaveReset" then
	call SaveReset()
elseif Action="Unite" then
	call Unite()
elseif Action="SaveUnite" then
	call SaveUnite()
	
ElseIf Action="isDisplay" Then
	Call isDisplay()
	
ElseIf Action="ClearBinding" Then
	Call ClearBinding()
	
else
	call main()
end if
if FoundErr=True then
	call oblog.sys_err(errmsg)
end if
''call CloseConn() 'shiyu


sub main()
	dim arrShowLine(10)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	dim sqlClass,rsClass,i,iDepth
	'sqlClass="select * From oblog_logclass Where idType=" & t & " order by RootID,OrderID"
	sqlClass="Select classid,id,classname,classlognum,ordernum,orderid,parentid,parentpath,"&_
					"depth,rootid,child,readme,previd,nextid,idType,userid_Nav,username,EName "&_
					" ,root_link,isDisplay "&_
				 "From oblog_logclass As logclass "&_
					 "Left Join oblog_user As oblogUser On ( logclass.userid_Nav = oblogUser.userid) "&_
				 "Where logclass.idType=" & t & " Order By logclass.RootID,logclass.OrderID"
			 
	set rsClass=server.CreateObject("adodb.recordset")
	rsClass.open sqlClass,conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" align="center"><strong><%=tName%>分类名称</strong></td>
	<td height="12" align="center"><strong title="请在 栏目空间管理+发布 的 ‘绑定分类’处进行绑定！" style="cursor:hand;">捆绑的栏目空间</strong></td>
    <td width="280" height="22" align="center"><strong>操作选项</strong></td>
  </tr>
  <% 
do while not rsClass.eof 
%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td> 
      <% 
	iDepth=rsClass("Depth")
	if rsClass("NextID")>0 then
		arrShowLine(iDepth)=True
	else
		arrShowLine(iDepth)=False
	end if
	if iDepth>0 then
	  	for i=1 to iDepth 
			if i=iDepth then 
				if rsClass("NextID")>0 then 
					response.write "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>" 
				else 
					response.write "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>" 
				end if 
			else 
				if arrShowLine(i)=True then
					response.write "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>" 
				else
					response.write "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>" 
				end if
			end if 
	  	next 
	  end if 
	  if rsClass("Child")>0 then 
	  	'response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	  else 
	  	'response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	  end if 
	  
	  if rsClass("Depth")=0 then 
	  	response.write "<b>" 
	  end if 
	  response.write "<a href='admin_logclass.asp?Action=Modify&id=" & rsClass("id") & "&t=" & t & "' title='" & rsClass("ReadMe")&"["& rsClass("classid") & "]'>" & rsClass("classname") & "</a>"
	  if rsClass("Child")>0 then 
	  	response.write "（" & rsClass("Child") & "）" 
	  end if 
	  
	  
	  Response.write "<font color=#666666 style='font-weight:normal;'>&nbsp;&nbsp;[&nbsp;"& rsClass("EName") &"&nbsp;]</font>"
	  if Cint(rsClass("userid_Nav"))>0 then
		Response.write "<font style='cursor:hand;font-weight:normal;' title='已绑定栏目空间' color=green>&nbsp; √已绑定</font>"
	  end if 
	  
	  If rsClass("root_link")<>"" And Len(rsClass("root_link"))>6 Then
		Response.write "<font style='cursor:hand;font-weight:normal;' title='"& rsClass("classname") &" 绑定了固定链接:"& rsClass("root_link") &"' color=#038ad7> √已固定</font>"
	  end if
	  
	  'response.write "&nbsp;&nbsp;" & rsClass("id") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
	  %>
    </td>
	<td height="22" align="center" style="background:#FFFFFF; border:solid 1px #BFE0FB;">
	<% If rsClass("root_link")<>"" And Len(rsClass("root_link"))>6 Then %>
	
		<a href="<%=rsClass("root_link")%>" target="_blank">
		<font style='cursor:hand;font-weight:normal;' title='<%=rsClass("classname")%> 绑定了固定链接:<%=rsClass("root_link")%>' color=#038ad7><%=rsClass("root_link")%></font></a>
		
	<% Else %>
	
		<a href="/blog.asp?name=<%=rsClass("username")%>" target="_blank">
		<font style='cursor:hand;font-weight:normal;' title='<%=rsClass("classname")%> 正在绑定的栏目空间:<%=rsClass("username")%>' color=green><%=rsClass("username")%></font></a>
		
	<% End If %>
	
	<% If rsClass("username")<>"" And Not(rsClass("root_link")<>"" And Len(rsClass("root_link"))>6) Then
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=gouser2&username=" & rsClass("username") & "' target='blank'><img src='/images/ajax-loader.gif' title='点击，进入栏目发布文章' /></a>"
	   End If
	%>
	</td>
	
    <td align="center">
	<a href="admin_logclass.asp?Action=isDisplay&isDisplay=<%=rsClass("isDisplay")%>&classid=<%=rsClass("classid")%>" onClick="confirm('确定转变前台状态吗？')">
	<% If rsClass("isDisplay")=1 Then
	Response.Write("<font color=green title='正处于我的住家网前台页面显示'>处于显示</font>")
	ElseIf rsClass("isDisplay")=0 Then
	Response.Write("<font color=red title='隐藏，并且不在我的住家网前台页面显示！'>处于隐藏</font>")
	End If %></a> | 
    	<a href="admin_logclass.asp?Action=Add&ParentID=<%=rsClass("id")%>&t=<%=t%>">添加子分类</a> | 
    	<a href="admin_logclass.asp?Action=Modify&id=<%=rsClass("id")%>&t=<%=t%>">修改</a> | 
      	<a href="admin_logclass.asp?Action=Move&id=<%=rsClass("id")%>&t=<%=t%>">移动分类</a> | 
      	<a href="admin_logclass.asp?Action=Del&id=<%=rsClass("id")%>&t=<%=t%>" onClick="<%if rsClass("Child")>0 then%>return ConfirmDel1();<%else%>return ConfirmDel2();<%end if%>">删除</a>
		<% If rsClass("userid_Nav")>0 Then %> | 
      	<a href="admin_logclass.asp?Action=ClearBinding&id=<%=rsClass("classid")%>" onClick="return confirm('确定强制清除绑定状态吗？');"><img src="/images/del.gif" title="强制删除和栏目空间的绑定状态！！" /></a>
		<% End If %></td>
  </tr>
  <% 
	rsClass.movenext 
loop 
%>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmDel1()
{
   alert("此分类下还有子分类，必须先删除下属子分类后才能删除此分类！");
   return false;
}

function ConfirmDel2()
{
   if(confirm("删除分类将不能恢复！确定要删除此分类吗？"))
     return true;
   else
     return false;
	 
}
</script>
<%
end sub

sub AddClass()
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="admin_logclass.asp?t=<%=t%>" onSubmit="return check()">  	
    <tr> 
      <td colspan="3" align="center" class="title"><strong>添 加 分 类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属分类：</strong><br>
        不能指定为外部分类 </td>
      <td> <select name="ParentID">
          <%call Admin_ShowClass_Option(0,ParentID)%>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>分类名称：</strong></td>
      <td><input name="classname" type="text" size="37" maxlength="20"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="350"><font color="blue"><strong>英文名称：</strong></font></td>
      <td><input name="EName" type="text" size="37" maxlength="40"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>分类说明：<br>
        </strong> 鼠标移至分类名称上时将显示设定的说明文字（不支持HTML）</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"></textarea></td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>我要固定链接指向一个地址：</strong></font> 例如：/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100"> (不需要固定链接请留空)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>是否在前台显示：</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1">显示 &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" checked="checked">隐藏</td>
    </tr>
	
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name="Add" type="submit" value=" 添&nbsp;&nbsp;加 " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">         
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.classname.value=="")
  {
    alert("分类名称不能为空！");
	document.form1.classname.focus();
	return false;
  }
}
</script>
<%
end sub

sub Modify()
	dim id,sql,rsClass,i
	id=trim(request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		response.Write "<br><li>找不到指定的分类！</li>"
		response.End()
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="admin_logclass.asp?t=<%=t%>" onSubmit="return check()">
    <tr> 
      <td colspan="3" align="center" class="title"><strong>修 改 分 类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属分类：</strong><br>
        如果你想改变所属分类，请<a href="admin_logclass.asp?Action=Move&id=<%=id%>&t=<%=t%>">点此移动分类</a></td>
      <td> 
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "无（作为一级分类）"
	else
    	dim rsParentClass,sqlParentClass
		sqlParentClass="Select * From oblog_logclass where id in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParentClass=server.CreateObject("adodb.recordset")
		rsParentClass.open sqlParentClass,conn,1,1
		do while not rsParentClass.eof
			for i=1 to rsParentClass("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParentClass("Depth")>0 then
				response.write "└"
			end if
			response.write "&nbsp;" & rsParentClass("classname") & "<br>"
			rsParentClass.movenext
		loop
		rsParentClass.close
		set rsParentClass=nothing
	end if
	%></select>
        </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>分类名称：</strong></td>
      <td><input name="classname" type="text" value="<%=rsClass("classname")%>" size="37" maxlength="20"> <input name="id" type="hidden" id="id" value="<%=rsClass("id")%>"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="350"><font color="blue"><strong>英文名称：</strong></font></td>
      <td><input name="EName" type="text" value="<%=rsClass("EName")%>" size="37" maxlength="40"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>分类说明：<br>
        </strong> 鼠标移至分类名称上时将显示设定的说明文字（不支持HTML）</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"><%=rsClass("ReadMe")%></textarea></td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>我要固定链接指向一个地址：</strong></font> 例如：/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100" value="<%=rsClass("root_link")%>"> (不需要固定链接请留空)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>是否在前台显示：</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1" <% If rsClass("isDisplay")=1 Then Response.write "checked='checked'"%>>显示 &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" <% If rsClass("isDisplay")=0 Then Response.write "checked='checked'"%>>隐藏</td>
    </tr>
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name="Submit" type="submit" value=" &nbsp;保存修改结果&nbsp; " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.classname.value=="")
  {
    alert("分类名称不能为空！");
	document.form1.classname.focus();
	return false;
  }
}
</script>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub MoveClass()
	dim id,sql,rsClass,i
	id=trim(request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
<form name="form1" method="post" action="admin_logclass.asp?&t=<%=t%>">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>移 动 分 类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>分类名称：</strong></td>
      <td><%=rsClass("classname")%> <input name="id" type="hidden" id="id" value="<%=rsClass("id")%>"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="200"><strong>英文名称：</strong></td>
      <td><%=rsClass("EName")%></td>
    </tr>
    <tr class="tdbg">
      <td width="200"><strong>当前所属分类：</strong></td>
      <td>
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "无（作为一级分类）"
	else
    	dim rsParent,sqlParent
		sqlParent="Select * From oblog_logclass where id in (" & rsClass("ParentPath") & ") order by Depth"
		set rsParent=server.CreateObject("adodb.recordset")
		rsParent.open sqlParent,conn,1,1
		do while not rsParent.eof
			for i=1 to rsParent("Depth")
				response.write "&nbsp;&nbsp;&nbsp;"
			next
			if rsParent("Depth")>0 then
				response.write "└"
			end if
			response.write "&nbsp;" & rsParent("classname") & "<br>"
			rsParent.movenext
		loop
		rsParent.close
		set rsParent=nothing
	end if
	%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>移动到：</strong><br>
        不能指定为当前分类的下属子分类<br>
        不能指定为外部分类</td>
      <td><select name="ParentID" size="2" style="height:300px;width:500px;"><%call Admin_ShowClass_Option(0,rsClass("ParentID"))%></select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveMove"> 
        <input name="Submit" type="submit" value=" &nbsp;保存移动结果&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='admin_logclass.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
	 </td>
   </tr>
</form>  
</table>
<%
	end if
	rsClass.close
	set rsClass=nothing
end sub

sub Order() 
	dim sqlClass,rsClass,i,iCount,j 
	sqlClass="select * From oblog_logclass where ParentID=0 and idtype="&t&" order by RootID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
	iCount=rsClass.recordcount 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>一 级 分 类 排 序</strong></td> 
  </tr> 
  <% 
j=1 
do while not rsClass.eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="200">&nbsp;<%=rsClass("classname")%></td> 
<% 
	if j>1 then 
  		response.write "<form action='admin_logclass.asp?Action=UpOrder&t=" & t & "' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=id value="&rsClass("id")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
	if iCount>j then 
  		response.write "<form action='admin_logclass.asp?Action=DownOrder&t=" & t & "' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=id value="&rsClass("id")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	j=j+1 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub OrderN() 
	dim sqlClass,rsClass,i,iCount,trs,UpMoveNum,DownMoveNum 
	sqlClass="select * From oblog_logclass where idtype="&t&" order by RootID,OrderID" 
	set rsClass=server.CreateObject("adodb.recordset") 
	rsClass.open sqlClass,conn,1,1 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>N 级 分 类 排 序</strong></td> 
  </tr> 
  <% 
do while not rsClass.eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="300"> 
	  <% 
	for i=1 to rsClass("Depth") 
	  	response.write "&nbsp;&nbsp;&nbsp;" 
	next 
	if rsClass("Child")>0 then 
		response.write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	else 
	  	response.write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	end if 
	if rsClass("ParentID")=0 then 
		response.write "<b>" 
	end if 
	response.write rsClass("classname") 
	if rsClass("Child")>0 then 
		response.write "(" & rsClass("Child") & ")" 
	end if 
	%></td> 
<% 
	if rsClass("ParentID")>0 then   '如果不是一级分类，则算出相同深度的分类数目，得到该分类在相同深度的分类中所处位置（之上或者之下的分类数） 
		'所能提升最大幅度应为For i=1 to 该版之上的版面数 
		set trs=conn.execute("select count(id) From oblog_logclass where ParentID="&rsClass("ParentID")&" and OrderID<"&rsClass("OrderID")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		if UpMoveNum>0 then 
  			response.write "<form action='admin_logclass.asp?Action=UpOrderN&t=" & t & "' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
			for i=1 to UpMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
		'所能降低最大幅度应为For i=1 to 该版之下的版面数 
		set trs=conn.execute("select count(id) From oblog_logclass where ParentID="&rsClass("ParentID")&" and orderID>"&rsClass("orderID")&"") 
		DownMoveNum=trs(0) 
		if isnull(DownMoveNum) then DownMoveNum=0 
		if DownMoveNum>0 then 
  			response.write "<form action='admin_logclass.asp?Action=DownOrderN&t=" & t & "'' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
			for i=1 to DownMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
	else 
		response.write "<td colspan=2>&nbsp;</td>" 
	end if 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	UpMoveNum=0 
	DownMoveNum=0 
	rsClass.movenext 
loop 
%> 
</table> 
<% 
	rsClass.close 
	set rsClass=nothing 
end sub 

sub Reset() 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="admin_logclass.asp?Action=SaveReset&t=<%=t%>"> 
	<tr>
	  <td colspan="3" align="center" class="title"><strong>复 位 所 有 分 类</strong></td> 
  </tr> 
    <tr class="tdbg">  
    <td align="center">  
        <table width="80%" border="0" cellspacing="1" cellpadding="1"> 
          <tr class="tdbg">  
            <td height="150"><span class="style1"><strong>注意：</strong></span><br> 
            &nbsp;&nbsp;&nbsp;&nbsp;如果选择复位所有分类，则所有分类都将作为一级分类，这时您需要重新对各个分类进行归属的基本设置。不要轻易使用该功能，仅在做出了错误的设置而无法复原分类之间的关系和排序的时候使用。          
		    </td> 
          </tr> 
        </table> 
	 <tr class="tdbg">  
    <td align="center">  
        <input type="submit" name="Submit" value="&nbsp;复位所有分类&nbsp;" style="cursor: hand;background-color: #cccccc;"> &nbsp;&nbsp;&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='admin_logclass.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
      </td>
    </tr>
	</form>
</table>
<%
end sub

sub Unite()
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
<form name="myform" method="post" action="admin_logclass.asp?t=<%=t%>" onSubmit="return ConfirmUnite();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>分 类 合 并</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;将分类 
        <select name="id" id="id">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
        合并到
        <select name="Targetid" id="Targetid">
        <%call Admin_ShowClass_Option(4,0)%>
        </select>
		</td>
		</tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="SaveUnite">
        <input type="submit" name="Submit" value=" &nbsp;合并分类&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">

	</td>
  </tr>
  <tr class="tdbg"> 
    <td height="60"><span class="style1"><strong>注意事项：</strong></span><br>
      &nbsp;&nbsp;&nbsp;&nbsp;所有操作不可逆，请慎重操作！！！<br>
      &nbsp;&nbsp;&nbsp;&nbsp;不能在同一个分类内进行操作，不能将一个分类合并到其下属分类中。目标分类中不能含有子分类。<br>
        &nbsp;&nbsp;&nbsp;&nbsp;合并后您所指定的分类（或者包括其下属分类）将被删除，所有用户将转移到目标分类中。</td>
  </tr>
 </form>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmUnite()
{
  if (document.myform.id.value==document.myform.Targetid.value)
  {
    alert("请不要在相同分类内进行操作！");
	document.myform.Targetid.focus();
	return false;
  }
  if (document.myform.Targetid.value=="")
  {
    alert("目标分类不能指定为含有子分类的分类！");
	document.myform.Targetid.focus();
	return false;
  }
}
</script>
<% 
end sub 
%> 
</body> 
</html> 
<% 

sub SaveAdd()
	dim id,classname,EName,Readme,PrevOrderID
	dim sql,rs,trs
	dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,Maxid,MaxRootID
	dim PrevID,NextID,Child
	Dim root_link,isDisplay

	classname=trim(request("classname"))
	EName=trim(request("EName"))
	Readme=trim(request("Readme"))
	
	root_link= Trim(Request("root_link"))
	isDisplay= Trim(Request("isDisplay"))
	
	
	if classname="" then
		response.Write "<br><li>分类名称不能为空！</li>"
		response.End()
	end if
	If Len(root_link)>100 Then
		Response.Write "<br><li>固定链接字符超过了100位长度！</li>"
		Response.End()
	End If
	
	set rs = conn.execute("select Max(id) From oblog_logclass")
	Maxid=rs(0)
	if isnull(Maxid) then
		Maxid=0
	end if
	rs.close
	id=Maxid+1
	set rs=conn.execute("select max(rootid) From oblog_logclass")
	MaxRootID=rs(0)
	if isnull(MaxRootID) then
		MaxRootID=0
	end if
	rs.close
	RootID=MaxRootID+1
	
	if ParentID>0 then
		sql="select * From oblog_logclass where id=" & ParentID & ""
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属分类已经被删除！</li>"
		end if
		if FoundErr=True then
			rs.close
			set rs=nothing
			exit sub
		else	
			RootID=rs("RootID")
			ParentName=rs("classname")
			ParentDepth=rs("Depth")
			ParentPath=rs("ParentPath")
			Child=rs("Child")
			ParentPath=ParentPath & "," & ParentID     '得到此分类的父级分类路径
			PrevOrderID=rs("OrderID")
			if Child>0 then		
				dim rsPrevOrderID
				'得到与本分类同级的最后一个分类的OrderID
				set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				set trs=conn.execute("select id From oblog_logclass where ParentID=" & ParentID & " and OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
				set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentPath like '" & ParentPath & ",%'")
				if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
					if not IsNull(rsPrevOrderID(0))  then
				 		if rsPrevOrderID(0)>PrevOrderID then
							PrevOrderID=rsPrevOrderID(0)
						end if
					end if
				end if
			else
				PrevID=0
			end if

		end if
		rs.close
	else
		if MaxRootID>0 then
			set trs=conn.execute("select id From oblog_logclass where RootID=" & MaxRootID & " and Depth=0")
			PrevID=trs(0)
			trs.close
		else
			PrevID=0
		end if
		PrevOrderID=0
		ParentPath="0"
	end if

	sql="Select * From oblog_logclass Where ParentID=" & ParentID & " AND classname='" & classname & "'"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		if ParentID=0 then
			ErrMsg=ErrMsg & "<br><li>已经存在一级分类：" & classname & "</li>"
		else
			ErrMsg=ErrMsg & "<br><li>“" & ParentName & "”中已经存在子分类“" & classname & "”！</li>"
		end if
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	
	sql="Select top 1 * From oblog_logclass"
	rs.open sql,conn,2,2
    rs.addnew
	rs("id")=id
   	rs("classname")=classname
	
	rs("EName")=EName
	rs("RootID")=RootID
	rs("ParentID")=ParentID
	if ParentID>0 then
		rs("Depth")=ParentDepth+1
	else
		rs("Depth")=0
	end if
	rs("ParentPath")=ParentPath
	rs("OrderID")=PrevOrderID
	rs("Child")=0
	rs("Readme")=Readme
	rs("PrevID")=PrevID
	rs("NextID")=0
	rs("idType")=t
	
	rs("root_link")= root_link
	rs("isDisplay")= Cint(isDisplay)
	
	rs.update
	rs.Close
    set rs=Nothing
	
	'更新与本分类同一父分类的上一个分类的“NextID”字段值
	if PrevID>0 then
		conn.execute("update oblog_logclass set NextID=" & id & " where id=" & PrevID)
	end if
	
	if ParentID>0 then
		'更新其父类的子分类数
		conn.execute("update oblog_logclass set child=child+1 where id="&ParentID)
		
		'更新该分类排序以及大于本需要和同在本分类下的分类排序序号
		conn.execute("update oblog_logclass set OrderID=OrderID+1 where rootid=" & rootid & " and OrderID>" & PrevOrderID)
		conn.execute("update oblog_logclass set OrderID=" & PrevOrderID & "+1 where id=" & id)
	end if
	
    'call CloseConn()
	'Response.Redirect "admin_logclass.asp?t=" & t
	oblog.showok "保存成功","admin_logclass.asp?t=" & t
end sub

sub SaveModify()
	dim classname,EName,Readme,IsElite,ShowOnTop,Setting,ClassMaster,ClassPicUrl,LinkUrl,SkinID,LayoutID,BrowsePurview,AddPurview
	dim trs,rs
	dim id,sql,rsClass,i
	dim SkinCount,LayoutCount
	Dim root_link,isDisplay
	
	id=trim(request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CLng(id)
	end if
	classname=trim(request("classname"))
	EName=trim(request("EName"))
	Readme=trim(request("Readme"))
	
	root_link= Trim(Request("root_link"))
	isDisplay= Trim(Request("isDisplay"))
	
	if classname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>分类名称不能为空！</li>"
	end if
	
	If Len(root_link) >100 Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>固定链接字符长度不能超过100！</li>"
	End If
	
	if FoundErr=True then
		exit sub
	end if	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
   	rsClass("classname")=classname
	rsClass("EName")=EName
	rsClass("Readme")=Readme
	
	rsClass("root_link")= root_link
	rsClass("isDisplay")= Cint(isDisplay)
	
	rsClass.update
	rsClass.close
	set rsClass=nothing
	
	set rs=nothing
	set trs=nothing	
    'call CloseConn()
	'Response.Redirect "admin_logclass.asp?t=" & t
	oblog.showok "保存成功","admin_logclass.asp?t=" & t
end sub


sub DeleteClass()
	dim sql,rs,PrevID,NextID,id
	id=trim(Request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select id,RootID,Depth,ParentID,Child,PrevID,NextID From oblog_logclass where id="&id
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>分类不存在，或者已经被删除</li>"
	else
		if rs("Child")>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>该分类含有子分类，请删除其子分类后再进行删除本分类的操作</li>"
		end if
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	if rs("Depth")>0 then
		conn.execute("update oblog_logclass set child=child-1 where id=" & rs("ParentID"))
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	'call CloseConn()
	response.redirect "admin_logclass.asp?t=" & t
		
end sub


sub SaveMove()
	dim id,sql,rsClass,i
	dim rParentID
	dim trs,rs
	dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID
	id=trim(request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
		rsClass.close
		set rsClass=nothing
		exit sub
	end if

	rParentID=trim(request("ParentID"))
	if rParentID="" then
		rParentID=0
	else
		rParentID=CLng(rParentID)
	end if
	
	if rsClass("ParentID")<>rParentID then   '更改了所属分类，则要做一系列检查
		if rParentID=rsClass("id") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属分类不能为自己！</li>"
		end if
		'判断所指定的分类是否为外部分类或本分类的下属分类
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=conn.execute("select rootid From oblog_logclass where id="&rParentID)
				if trs.bof and trs.eof then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>不能指定外部分类为所属分类</li>"
				else
					if rsClass("rootid")=trs(0) then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>不能指定该分类的下属分类作为所属分类</li>"
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=conn.execute("select id From oblog_logclass where ParentPath like '"&rsClass("ParentPath")&"," & rsClass("id") & "%' and id="&rParentID)
			if not (trs.eof and trs.bof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>您不能指定该分类的下属分类作为所属分类</li>"
			end if
			trs.close
			set trs=nothing
		end if
		
	end if

	if FoundErr=True then
		rsClass.close
		set rsClass=nothing
		exit sub
	end if
	
	if rsClass("ParentID")=0 then
		ParentID=rsClass("id")
		iParentID=0
	else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	end if
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.close
	set rsClass=nothing
	
	
  '假如更改了所属分类
  '需要更新其原来所属分类信息，包括深度、父级ID、分类数、排序、继承版主等数据
  '需要更新当前所属分类信息
  '继承版主数据需要另写函数进行更新--取消，在前台可用id in ParentPath来获得
  dim mrs,MaxRootID
  set mrs=conn.execute("select max(rootid) From oblog_logclass")
  MaxRootID=mrs(0)
  set mrs=nothing
  if isnull(MaxRootID) then
	MaxRootID=0
  end if
  dim k,nParentPath,mParentPath
  dim ParentSql,ClassCount
  dim rsPrevOrderID
  if clng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '假如更改了所属分类
	'更新原来同一父分类的上一个分类的NextID和下一个分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	if iParentID>0 and rParentID=0 then  	'如果原来不是一级分类改成一级分类
		'得到上一个一级分类分类
		sql="select id,NextID From oblog_logclass where RootID=" & MaxRootID & " and Depth=0"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '得到新的PrevID
		rs(1)=id     '更新上一个一级分类分类的NextID的值
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'更新当前分类数据
		conn.execute("update oblog_logclass set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where id="&id)
		'如果有下属分类，则更新其下属分类数据。下属分类的排序不需考虑，只需更新下属分类深度和一级排序ID(rootid)数据
		if child>0 then
			i=0
			ParentPath=ParentPath & ","
			set rs=conn.execute("select * From oblog_logclass where ParentPath like '%"&ParentPath & id&"%'")
			do while not rs.eof
				i=i+1
				mParentPath=replace(rs("ParentPath"),ParentPath,"")
				conn.execute("update oblog_logclass set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' where id="&rs("id"))
				rs.movenext
			loop
			rs.close
			set rs=nothing
		end if
		
		'更新其原来所属分类的分类数，排序相当于剪枝而不需考虑
		conn.execute("update oblog_logclass set child=child-1 where id="&iParentID)
		
	elseif iParentID>0 and rParentID>0 then    '如果是将一个分分类移动到其他分分类下
		'得到当前分类的下属子分类数
		ParentPath=ParentPath & ","
		set rs=conn.execute("select count(*) From oblog_logclass where ParentPath like '%"&ParentPath & id&"%'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From oblog_logclass where id="&rParentID)
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			'得到与本分类同级的最后一个分类的id
			sql="select id,NextID From oblog_logclass where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '得到新的PrevID
			rs(1)=id     '更新上一个分类的NextID的值
			rs.update
			rs.close
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentPath like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
		
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update oblog_logclass set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		'更新当前分类数据
		conn.execute("update oblog_logclass set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("id") & "',PrevID=" & PrevID & ",NextID=0 where id="&id)
		
		'如果有子分类则更新子分类数据，深度为原来的相对深度加上当前所属分类的深度
		set rs=conn.execute("select * From oblog_logclass where ParentPath like '%"&ParentPath&id&"%' order by OrderID")
		i=1
		do while not rs.eof
			i=i+1
			iParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.execute("update oblog_logclass set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' where id="&rs("id"))
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		
		'更新所指向的上级分类的子分类数
		conn.execute("update oblog_logclass set child=child+1 where id="&rParentID)
		
		'更新其原父类的子分类数			
		conn.execute("update oblog_logclass set child=child-1 where id="&iParentID)
	else    '如果原来是一级分类改成其他分类的下属分类
		'得到移动的分类总数
		set rs=conn.execute("select count(*) From oblog_logclass where rootid="&rootid)
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'获得目标分类的相关信息		
		set trs=conn.execute("select * From oblog_logclass where id="&rParentID)
		if trs("Child")>0 then		
			'得到与本分类同级的最后一个分类的OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			sql="select id,NextID From oblog_logclass where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=id
			rs.update
			set rs=nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentPath like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			if (not(rsPrevOrderID.bof and rsPrevOrderID.eof)) then
				if not IsNull(rsPrevOrderID(0))  then
			 		if rsPrevOrderID(0)>PrevOrderID then
						PrevOrderID=rsPrevOrderID(0)
					end if
				end if
			end if
		else
			PrevID=0
			PrevOrderID=trs("OrderID")
		end if
	
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.execute("update oblog_logclass set OrderID=OrderID+" & ClassCount &"+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		conn.execute("update oblog_logclass set PrevID=" & PrevID & ",NextID=0 where id=" & id)
		set rs=conn.execute("select * From oblog_logclass where rootid="&rootid&" order by OrderID")
		i=0
		do while not rs.eof
			i=i+1
			if rs("parentid")=0 then
				ParentPath=trs("ParentPath") & "," & trs("id")
				conn.execute("update oblog_logclass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" where id="&rs("id"))
			else
				ParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),"0,","")
				conn.execute("update oblog_logclass set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' where id="&rs("id"))
			end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
		trs.close
		set trs=nothing
		'更新所指向的上级分类分类数		
		conn.execute("update oblog_logclass set child=child+1 where id="&rParentID)

	end if
  end if
	
  'call CloseConn()
  Response.Redirect "admin_logclass.asp?t=" & t
end sub

sub UpOrder()
	dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=trim(request("id"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CLng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From oblog_logclass where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From oblog_logclass")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update oblog_logclass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前分类以上的分类的RootID依次加一，范围为要提升的数字
	sqlOrder="select * From oblog_logclass where ParentID=0 and RootID<" & cRootID & " order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最上面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		i=i+1
		if i>MoveNum then
			rsOrder("PrevID")=id
			rsOrder.update
			conn.execute("update oblog_logclass set NextID=" & rsOrder("id") & " where id=" & id)
			conn.execute("update oblog_logclass set RootID=RootID+1 where RootID=" & tRootID)
			exit do
		end if
		conn.execute("update oblog_logclass set RootID=RootID+1 where RootID=" & tRootID)
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update oblog_logclass set PrevID=0 where id=" & id)
	else
		rsOrder("NextID")=id
		rsOrder.update
		conn.execute("update oblog_logclass set PrevID=" & rsOrder("id") & " where id=" & id)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update oblog_logclass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "admin_logclass.asp?Action=Order&t=" & t
end sub

sub DownOrder()
	dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=trim(request("id"))
	cRootID=Trim(request("cRootID"))
	MoveNum=trim(request("MoveNum"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	else
		id=CLng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到本分类的PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From oblog_logclass where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From oblog_logclass")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.execute("update oblog_logclass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'然后将位于当前分类以下的分类的RootID依次减一，范围为要下降的数字
	sqlOrder="select * From oblog_logclass where ParentID=0 and idType =" & t & " And RootID>" & cRootID & " order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,2,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '如果当前分类已经在最下面，则无需移动
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		
		i=i+1
		if i>MoveNum then
			rsOrder("NextID")=id
			rsOrder.update
			conn.execute("update oblog_logclass set PrevID=" & rsOrder("id") & " where id=" & id)
			conn.execute("update oblog_logclass set RootID=RootID-1 where RootID=" & tRootID)
			exit do
		end if
		conn.execute("update oblog_logclass set RootID=RootID-1 where RootID=" & tRootID)
		rsOrder.movenext
	loop
	rsOrder.movenext
	if rsOrder.eof then
		conn.execute("update oblog_logclass set NextID=0 where id=" & id)
	else
		rsOrder("PrevID")=id
		rsOrder.update
		conn.execute("update oblog_logclass set NextID=" & rsOrder("id") & " where id=" & id)
	end if	
	rsOrder.close
	set rsOrder=nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.execute("update oblog_logclass set RootID=" & tRootID & " where RootID=" & MaxRootID)
	'call CloseConn()
	response.Redirect "admin_logclass.asp?Action=Order&t=" & t
end sub

sub UpOrderN()
	dim sqlOrder,rsOrder,MoveNum,id,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=Trim(request("id"))
	MoveNum=trim(request("MoveNum"))
	if id="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		id=CLng(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From oblog_logclass where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing
	if child>0 then
		set rs=conn.execute("select count(*) From oblog_logclass where ParentPath like '%"&ParentPath&"%'")
		oldorders=rs(0)
		rs.close
		set rs=nothing
	else
		oldorders=0
	end if
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'和该分类同级且排序在其之上的分类------更新其排序，范围为要提升的数字
	sql="select id,OrderID,child,ParentPath,PrevID,NextID From oblog_logclass where ParentID="&ParentID&" and OrderID<"&OrderID&" order by OrderID desc"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=1
	do while not rs.eof
		tOrderID=rs(1)
		
		if rs(2)>0 then
			ii=i+1
			set trs=conn.execute("select id,OrderID From oblog_logclass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					conn.execute("update oblog_logclass set OrderID="&tOrderID+oldorders+ii&" where id="&trs(0))
					ii=ii+1
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		i=i+1
		if i>MoveNum then
			rs(4)=id
			rs.update
			conn.execute("update oblog_logclass set NextID=" & rs(0) & " where id=" & id)
			conn.execute("update oblog_logclass set OrderID="&tOrderID+oldorders+i-1&" where id="&rs(0))		
			exit do
		end if
		conn.execute("update oblog_logclass set OrderID="&tOrderID+oldorders+i-1&" where id="&rs(0))
		rs.movenext
	loop
	if not rs.eof then
	rs.movenext
	end if
	if rs.eof then
		conn.execute("update oblog_logclass set PrevID=0 where id=" & id)
	else
		rs(5)=id
		rs.update
		conn.execute("update oblog_logclass set PrevID=" & rs(0) & " where id=" & id)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update oblog_logclass set OrderID="&tOrderID&" where id="&id)
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select id From oblog_logclass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update oblog_logclass set OrderID="&tOrderID+i&" where id="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "admin_logclass.asp?Action=OrderN&t=" & t
end sub

sub DownOrderN()
	dim sqlOrder,rsOrder,MoveNum,id,i
	dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=Trim(request("id"))
	MoveNum=trim(request("MoveNum"))
	if id="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		id=Cint(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		exit sub
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要下降的数字！</li>"
			exit sub
		end if
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From oblog_logclass where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing

	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'和该分类同级且排序在其之下的分类------更新其排序，范围为要下降的数字
	sql="select id,OrderID,child,ParentPath,PrevID,NextID From oblog_logclass where ParentID="&ParentID&" and OrderID>"&OrderID&" order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      '同级分类
	ii=0     '同级分类和子分类
	do while not rs.eof
		'conn.execute("update oblog_logclass set OrderID="&OrderID+ii&" where id="&rs(0))
		if rs(2)>0 then
			set trs=conn.execute("select id,OrderID From oblog_logclass where ParentPath like '%"&rs(3)&","&rs(0)&"%' order by OrderID")
			if not (trs.eof and trs.bof) then
				do while not trs.eof
					ii=ii+1
					conn.execute("update oblog_logclass set OrderID="&OrderID+ii&" where id="&trs(0))
					trs.movenext
				loop
			end if
			trs.close
			set trs=nothing
		end if
		ii=ii+1
		i=i+1
		if i>=MoveNum then
			rs(5)=id
			rs.update
			conn.execute("update oblog_logclass set PrevID=" & rs(0) & " where id=" & id)
			conn.execute("update oblog_logclass set OrderID="&OrderID+ii-1&" where id="&rs(0))		
			exit do
		end if
		conn.execute("update oblog_logclass set OrderID="&OrderID+ii-1&" where id="&rs(0))
		rs.movenext
	loop
	rs.movenext
	if rs.eof then
		conn.execute("update oblog_logclass set NextID=0 where id=" & id)
	else
		rs(4)=id
		rs.update
		conn.execute("update oblog_logclass set NextID=" & rs(0) & " where id=" & id)
	end if	
	rs.close
	set rs=nothing
	
	'更新所要排序的分类的序号
	conn.execute("update oblog_logclass set OrderID="&OrderID+ii&" where id="&id)
	'如果有下属分类，则更新其下属分类排序
	if child>0 then
		i=1
		set rs=conn.execute("select id From oblog_logclass where ParentPath like '%"&ParentPath&"%' order by OrderID")
		do while not rs.eof
			conn.execute("update oblog_logclass set OrderID="&OrderID+ii+i&" where id="&rs(0))
			i=i+1
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if
	'call CloseConn()
	response.Redirect "admin_logclass.asp?Action=OrderN&t=" & t
end sub

sub SaveReset()
	dim i,sql,rs,SuccessMsg,iCount,PrevID,NextID
	sql="select id From oblog_logclass Where idType=" & t & " order by RootID,OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,1
	iCount=rs.recordcount
	i=1
	PrevID=0
	do while not rs.eof
		rs.movenext
		if rs.eof then
			NextID=0
		else
			NextID=rs(0)
		end if
		rs.moveprevious
		conn.execute("update oblog_logclass set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & " where id=" & rs(0))
		PrevID=rs(0)
		i=i+1
		rs.movenext
	loop
	rs.close
	set rs=nothing	
	
	response.Write "复位成功！请返回<a href='admin_logclass.asp?t=" & t & "'>分类管理首页</a>做分类的归属设置。"
end sub

sub SaveUnite()
	dim id,Targetid,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i
	id=trim(request("id"))
	Targetid=trim(request("Targetid"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要合并的分类！</li>"
	else
		id=CLng(id)
	end if
	if Targetid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标分类！</li>"
	else
		Targetid=CLng(Targetid)
	end if
	if id=Targetid then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请不要在相同分类内进行操作</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	'判断目标分类是否有子分类，如果有，则报错。
	set rs=conn.execute("select Child From oblog_logclass where id=" & Targetid)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>目标分类不存在，可能已经被删除！</li>"
	else
		if rs(0)>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>目标分类中含有子分类，不能合并！</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'得到当前分类信息
	set rs=conn.execute("select id,ParentID,ParentPath,PrevID,NextID,Depth From oblog_logclass where id="&id)
	iParentID=rs(1)
	Depth=rs(5)
	if iParentID=0 then
		ParentPath=rs(0)
	else
		ParentPath=rs(2) & "," & rs(0)
	end if
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'判断是否是合并到其下属分类中
	set rs=conn.execute("select id From oblog_logclass where id="&Targetid&" and ParentPath like '"&ParentPath&"%'")
	if not (rs.eof and rs.bof) then
		response.Write "<br><li>不能将一个分类合并到其下属子分类中</li>"
		exit sub
	end if
	
	'得到当前分类的下属分类ID
	set rs=conn.execute("select id From oblog_logclass where ParentPath like '"&ParentPath&"%'")
	i=0
	if not (rs.eof and rs.bof) then
		do while not rs.eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.movenext
		loop
	end if
	if i>0 then
		ParentPath=iParentPath
	else
		ParentPath=id
	end if
	
	'先修改上一分类的NextID和下一分类的PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'更新log所属分类
	conn.execute("update [oblog_log] set classid="&Targetid&" where classid in ("&ParentPath&")")
	
	'删除被合并分类及其下属分类
	conn.execute("delete From oblog_logclass where id in ("&ParentPath&")")
	
	'更新其原来所属分类的子分类数，排序相当于剪枝而不需考虑
	if Depth>0 then
		conn.execute("update oblog_logclass set Child=Child-1 where id="&iParentID)
	end if
	
	response.Write "分类合并成功！已经将被合并分类及其下属子分类的所有数据转入目标分类中。<br><br>同时删除了被合并的分类及其子分类。"
	set rs=nothing
	set trs=nothing
end sub

sub Admin_ShowClass_Option(ShowType,CurrentID)
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
						't="0"   tName="文章     /相册
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value='" & rsClass("id") & "'"
			if CurrentID>0 and rsClass("id")=CurrentID then
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub



Sub isDisplay()
	dim sql,rs,classid
	Dim isDisplay
	
	classid=trim(Request("classid"))
	isDisplay= Cint(Request("isDisplay"))
	
	if classid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		classid=CLng(classid)
	end if
	

	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if

	If isDisplay= 1 Then
		conn.execute("update oblog_logclass set isDisplay=0 where classid=" & classid )
	ElseIf isDisplay= 0 Then
		conn.execute("update oblog_logclass set isDisplay=1 where classid=" & classid )
	End If

	response.redirect "admin_logclass.asp"
		
end sub



Sub ClearBinding()'强制清除绑定状态；
	dim sql,rs,id
	
	
	id=trim(Request("id"))
	
	
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	

	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if

	
		conn.execute("Update oblog_logclass Set userid_Nav=0 Where classid=" & id )


	response.redirect "admin_logclass.asp"
		
end sub

%> 
