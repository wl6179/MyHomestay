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
	tName="���"
Else
	t="0"
	tName="����"
End If
%>
<html>
<head>
<title><%=tName%>�������</title>
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
	<td colspan="2" align="center" class="title"><strong><%=tName%>�������</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="80" height="30"><strong>��������</strong></td>
    <td height="30">
    	<a href="admin_logclass.asp?t=<%=t%>"><%=tName%>���������ҳ</a> | 
    	<a href="admin_logclass.asp?Action=Add&t=<%=t%>">���<%=tName%>����</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Order&t=<%=t%>">һ����������</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=OrderN&t=<%=t%>">N����������</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Reset&t=<%=t%>">��λ����<%=tName%>����</a>&nbsp;|&nbsp;
    	<a href="admin_logclass.asp?Action=Unite&t=<%=t%>"><%=tName%>����ϲ�</a></td>
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
    <td height="22" align="center"><strong><%=tName%>��������</strong></td>
	<td height="12" align="center"><strong title="���� ��Ŀ�ռ����+���� �� ���󶨷��࡯�����а󶨣�" style="cursor:hand;">�������Ŀ�ռ�</strong></td>
    <td width="280" height="22" align="center"><strong>����ѡ��</strong></td>
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
	  	response.write "��" & rsClass("Child") & "��" 
	  end if 
	  
	  
	  Response.write "<font color=#666666 style='font-weight:normal;'>&nbsp;&nbsp;[&nbsp;"& rsClass("EName") &"&nbsp;]</font>"
	  if Cint(rsClass("userid_Nav"))>0 then
		Response.write "<font style='cursor:hand;font-weight:normal;' title='�Ѱ���Ŀ�ռ�' color=green>&nbsp; ���Ѱ�</font>"
	  end if 
	  
	  If rsClass("root_link")<>"" And Len(rsClass("root_link"))>6 Then
		Response.write "<font style='cursor:hand;font-weight:normal;' title='"& rsClass("classname") &" ���˹̶�����:"& rsClass("root_link") &"' color=#038ad7> ���ѹ̶�</font>"
	  end if
	  
	  'response.write "&nbsp;&nbsp;" & rsClass("id") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
	  %>
    </td>
	<td height="22" align="center" style="background:#FFFFFF; border:solid 1px #BFE0FB;">
	<% If rsClass("root_link")<>"" And Len(rsClass("root_link"))>6 Then %>
	
		<a href="<%=rsClass("root_link")%>" target="_blank">
		<font style='cursor:hand;font-weight:normal;' title='<%=rsClass("classname")%> ���˹̶�����:<%=rsClass("root_link")%>' color=#038ad7><%=rsClass("root_link")%></font></a>
		
	<% Else %>
	
		<a href="/blog.asp?name=<%=rsClass("username")%>" target="_blank">
		<font style='cursor:hand;font-weight:normal;' title='<%=rsClass("classname")%> ���ڰ󶨵���Ŀ�ռ�:<%=rsClass("username")%>' color=green><%=rsClass("username")%></font></a>
		
	<% End If %>
	
	<% If rsClass("username")<>"" And Not(rsClass("root_link")<>"" And Len(rsClass("root_link"))>6) Then
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=gouser2&username=" & rsClass("username") & "' target='blank'><img src='/images/ajax-loader.gif' title='�����������Ŀ��������' /></a>"
	   End If
	%>
	</td>
	
    <td align="center">
	<a href="admin_logclass.asp?Action=isDisplay&isDisplay=<%=rsClass("isDisplay")%>&classid=<%=rsClass("classid")%>" onClick="confirm('ȷ��ת��ǰ̨״̬��')">
	<% If rsClass("isDisplay")=1 Then
	Response.Write("<font color=green title='�������ҵ�ס����ǰ̨ҳ����ʾ'>������ʾ</font>")
	ElseIf rsClass("isDisplay")=0 Then
	Response.Write("<font color=red title='���أ����Ҳ����ҵ�ס����ǰ̨ҳ����ʾ��'>��������</font>")
	End If %></a> | 
    	<a href="admin_logclass.asp?Action=Add&ParentID=<%=rsClass("id")%>&t=<%=t%>">����ӷ���</a> | 
    	<a href="admin_logclass.asp?Action=Modify&id=<%=rsClass("id")%>&t=<%=t%>">�޸�</a> | 
      	<a href="admin_logclass.asp?Action=Move&id=<%=rsClass("id")%>&t=<%=t%>">�ƶ�����</a> | 
      	<a href="admin_logclass.asp?Action=Del&id=<%=rsClass("id")%>&t=<%=t%>" onClick="<%if rsClass("Child")>0 then%>return ConfirmDel1();<%else%>return ConfirmDel2();<%end if%>">ɾ��</a>
		<% If rsClass("userid_Nav")>0 Then %> | 
      	<a href="admin_logclass.asp?Action=ClearBinding&id=<%=rsClass("classid")%>" onClick="return confirm('ȷ��ǿ�������״̬��');"><img src="/images/del.gif" title="ǿ��ɾ������Ŀ�ռ�İ�״̬����" /></a>
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
   alert("�˷����»����ӷ��࣬������ɾ�������ӷ�������ɾ���˷��࣡");
   return false;
}

function ConfirmDel2()
{
   if(confirm("ɾ�����ཫ���ָܻ���ȷ��Ҫɾ���˷�����"))
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
      <td colspan="3" align="center" class="title"><strong>�� �� �� ��</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>�������ࣺ</strong><br>
        ����ָ��Ϊ�ⲿ���� </td>
      <td> <select name="ParentID">
          <%call Admin_ShowClass_Option(0,ParentID)%>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>�������ƣ�</strong></td>
      <td><input name="classname" type="text" size="37" maxlength="20"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="350"><font color="blue"><strong>Ӣ�����ƣ�</strong></font></td>
      <td><input name="EName" type="text" size="37" maxlength="40"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>����˵����<br>
        </strong> �����������������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"></textarea></td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>��Ҫ�̶�����ָ��һ����ַ��</strong></font> ���磺/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100"> (����Ҫ�̶�����������)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>�Ƿ���ǰ̨��ʾ��</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1">��ʾ &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" checked="checked">����</td>
    </tr>
	
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name="Add" type="submit" value=" ��&nbsp;&nbsp;�� " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">         
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.classname.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		id=CLng(id)
	end if
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		response.Write "<br><li>�Ҳ���ָ���ķ��࣡</li>"
		response.End()
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="admin_logclass.asp?t=<%=t%>" onSubmit="return check()">
    <tr> 
      <td colspan="3" align="center" class="title"><strong>�� �� �� ��</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>�������ࣺ</strong><br>
        �������ı��������࣬��<a href="admin_logclass.asp?Action=Move&id=<%=id%>&t=<%=t%>">����ƶ�����</a></td>
      <td> 
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "�ޣ���Ϊһ�����ࣩ"
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
				response.write "��"
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
      <td width="350"><strong>�������ƣ�</strong></td>
      <td><input name="classname" type="text" value="<%=rsClass("classname")%>" size="37" maxlength="20"> <input name="id" type="hidden" id="id" value="<%=rsClass("id")%>"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="350"><font color="blue"><strong>Ӣ�����ƣ�</strong></font></td>
      <td><input name="EName" type="text" value="<%=rsClass("EName")%>" size="37" maxlength="40"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>����˵����<br>
        </strong> �����������������ʱ����ʾ�趨��˵�����֣���֧��HTML��</td>
      <td><textarea name="Readme" cols="37" rows="4" id="Readme"><%=rsClass("ReadMe")%></textarea></td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>��Ҫ�̶�����ָ��һ����ַ��</strong></font> ���磺/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100" value="<%=rsClass("root_link")%>"> (����Ҫ�̶�����������)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>�Ƿ���ǰ̨��ʾ��</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1" <% If rsClass("isDisplay")=1 Then Response.write "checked='checked'"%>>��ʾ &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" <% If rsClass("isDisplay")=0 Then Response.write "checked='checked'"%>>����</td>
    </tr>
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name="Submit" type="submit" value=" &nbsp;�����޸Ľ��&nbsp; " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.classname.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���ķ��࣡</li>"
	else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
<form name="form1" method="post" action="admin_logclass.asp?&t=<%=t%>">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>�� �� �� ��</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>�������ƣ�</strong></td>
      <td><%=rsClass("classname")%> <input name="id" type="hidden" id="id" value="<%=rsClass("id")%>"></td>
    </tr>
	<tr class="tdbg"> 
      <td width="200"><strong>Ӣ�����ƣ�</strong></td>
      <td><%=rsClass("EName")%></td>
    </tr>
    <tr class="tdbg">
      <td width="200"><strong>��ǰ�������ࣺ</strong></td>
      <td>
        <%
	if rsClass("ParentID")<=0 then
	  	response.write "�ޣ���Ϊһ�����ࣩ"
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
				response.write "��"
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
      <td width="200"><strong>�ƶ�����</strong><br>
        ����ָ��Ϊ��ǰ����������ӷ���<br>
        ����ָ��Ϊ�ⲿ����</td>
      <td><select name="ParentID" size="2" style="height:300px;width:500px;"><%call Admin_ShowClass_Option(0,rsClass("ParentID"))%></select></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveMove"> 
        <input name="Submit" type="submit" value=" &nbsp;�����ƶ����&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='admin_logclass.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
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
	  <td colspan="4" align="center" class="title"><strong>һ �� �� �� �� ��</strong></td> 
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
		response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
		for i=1 to j-1 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=id value="&rsClass("id")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
		response.write "</td></form>" 
	else 
		response.write "<td width='150'>&nbsp;</td>" 
	end if 
	if iCount>j then 
  		response.write "<form action='admin_logclass.asp?Action=DownOrder&t=" & t & "' method='post'><td width='150'>" 
		response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
		for i=1 to iCount-j 
			response.write "<option value="&i&">"&i&"</option>" 
		next 
		response.write "</select>" 
		response.write "<input type=hidden name=id value="&rsClass("id")&">"
		response.write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
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
	  <td colspan="4" align="center" class="title"><strong>N �� �� �� �� ��</strong></td> 
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
	if rsClass("ParentID")>0 then   '�������һ�����࣬�������ͬ��ȵķ�����Ŀ���õ��÷�������ͬ��ȵķ���������λ�ã�֮�ϻ���֮�µķ������� 
		'��������������ӦΪFor i=1 to �ð�֮�ϵİ����� 
		set trs=conn.execute("select count(id) From oblog_logclass where ParentID="&rsClass("ParentID")&" and OrderID<"&rsClass("OrderID")&"") 
		UpMoveNum=trs(0) 
		if isnull(UpMoveNum) then UpMoveNum=0 
		if UpMoveNum>0 then 
  			response.write "<form action='admin_logclass.asp?Action=UpOrderN&t=" & t & "' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
			for i=1 to UpMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
			response.write "</td></form>" 
		else 
			response.write "<td width='150'>&nbsp;</td>" 
		end if 
		trs.close 
		'���ܽ���������ӦΪFor i=1 to �ð�֮�µİ����� 
		set trs=conn.execute("select count(id) From oblog_logclass where ParentID="&rsClass("ParentID")&" and orderID>"&rsClass("orderID")&"") 
		DownMoveNum=trs(0) 
		if isnull(DownMoveNum) then DownMoveNum=0 
		if DownMoveNum>0 then 
  			response.write "<form action='admin_logclass.asp?Action=DownOrderN&t=" & t & "'' method='post'><td width='150'>" 
			response.write "<select name=MoveNum size=1><option value=0>�����ƶ�</option>" 
			for i=1 to DownMoveNum 
				response.write "<option value="&i&">"&i&"</option>" 
			next 
			response.write "</select>" 
			response.write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=��&nbsp;�� style='cursor: hand;background-color: #cccccc;'>" 
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
	  <td colspan="3" align="center" class="title"><strong>�� λ �� �� �� ��</strong></td> 
  </tr> 
    <tr class="tdbg">  
    <td align="center">  
        <table width="80%" border="0" cellspacing="1" cellpadding="1"> 
          <tr class="tdbg">  
            <td height="150"><span class="style1"><strong>ע�⣺</strong></span><br> 
            &nbsp;&nbsp;&nbsp;&nbsp;���ѡ��λ���з��࣬�����з��඼����Ϊһ�����࣬��ʱ����Ҫ���¶Ը���������й����Ļ������á���Ҫ����ʹ�øù��ܣ����������˴�������ö��޷���ԭ����֮��Ĺ�ϵ�������ʱ��ʹ�á�          
		    </td> 
          </tr> 
        </table> 
	 <tr class="tdbg">  
    <td align="center">  
        <input type="submit" name="Submit" value="&nbsp;��λ���з���&nbsp;" style="cursor: hand;background-color: #cccccc;"> &nbsp;&nbsp;&nbsp; 
		<input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='admin_logclass.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
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
	  <td colspan="3" align="center" class="title"><strong>�� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;������ 
        <select name="id" id="id">
        <%call Admin_ShowClass_Option(1,0)%>
        </select>
        �ϲ���
        <select name="Targetid" id="Targetid">
        <%call Admin_ShowClass_Option(4,0)%>
        </select>
		</td>
		</tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="SaveUnite">
        <input type="submit" name="Submit" value=" &nbsp;�ϲ�����&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" ȡ&nbsp;&nbsp;�� " onClick="window.location.href='admin_logclass.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">

	</td>
  </tr>
  <tr class="tdbg"> 
    <td height="60"><span class="style1"><strong>ע�����</strong></span><br>
      &nbsp;&nbsp;&nbsp;&nbsp;���в��������棬�����ز���������<br>
      &nbsp;&nbsp;&nbsp;&nbsp;������ͬһ�������ڽ��в��������ܽ�һ������ϲ��������������С�Ŀ������в��ܺ����ӷ��ࡣ<br>
        &nbsp;&nbsp;&nbsp;&nbsp;�ϲ�������ָ���ķ��ࣨ���߰������������ࣩ����ɾ���������û���ת�Ƶ�Ŀ������С�</td>
  </tr>
 </form>
</table> 
<script language="JavaScript" type="text/JavaScript">
function ConfirmUnite()
{
  if (document.myform.id.value==document.myform.Targetid.value)
  {
    alert("�벻Ҫ����ͬ�����ڽ��в�����");
	document.myform.Targetid.focus();
	return false;
  }
  if (document.myform.Targetid.value=="")
  {
    alert("Ŀ����಻��ָ��Ϊ�����ӷ���ķ��࣡");
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
		response.Write "<br><li>�������Ʋ���Ϊ�գ�</li>"
		response.End()
	end if
	If Len(root_link)>100 Then
		Response.Write "<br><li>�̶������ַ�������100λ���ȣ�</li>"
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
			ErrMsg=ErrMsg & "<br><li>���������Ѿ���ɾ����</li>"
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
			ParentPath=ParentPath & "," & ParentID     '�õ��˷���ĸ�������·��
			PrevOrderID=rs("OrderID")
			if Child>0 then		
				dim rsPrevOrderID
				'�õ��뱾����ͬ�������һ�������OrderID
				set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				set trs=conn.execute("select id From oblog_logclass where ParentID=" & ParentID & " and OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'�õ�ͬһ�����൫�ȱ����༶������ӷ�������OrderID�������ǰһ��ֵ����������ֵ��
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
			ErrMsg=ErrMsg & "<br><li>�Ѿ�����һ�����ࣺ" & classname & "</li>"
		else
			ErrMsg=ErrMsg & "<br><li>��" & ParentName & "�����Ѿ������ӷ��ࡰ" & classname & "����</li>"
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
	
	'�����뱾����ͬһ���������һ������ġ�NextID���ֶ�ֵ
	if PrevID>0 then
		conn.execute("update oblog_logclass set NextID=" & id & " where id=" & PrevID)
	end if
	
	if ParentID>0 then
		'�����丸����ӷ�����
		conn.execute("update oblog_logclass set child=child+1 where id="&ParentID)
		
		'���¸÷��������Լ����ڱ���Ҫ��ͬ�ڱ������µķ����������
		conn.execute("update oblog_logclass set OrderID=OrderID+1 where rootid=" & rootid & " and OrderID>" & PrevOrderID)
		conn.execute("update oblog_logclass set OrderID=" & PrevOrderID & "+1 where id=" & id)
	end if
	
    'call CloseConn()
	'Response.Redirect "admin_logclass.asp?t=" & t
	oblog.showok "����ɹ�","admin_logclass.asp?t=" & t
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
		ErrMsg=ErrMsg & "<br><li>�������Ʋ���Ϊ�գ�</li>"
	end if
	
	If Len(root_link) >100 Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�̶������ַ����Ȳ��ܳ���100��</li>"
	End If
	
	if FoundErr=True then
		exit sub
	end if	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���ķ��࣡</li>"
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
	oblog.showok "����ɹ�","admin_logclass.asp?t=" & t
end sub


sub DeleteClass()
	dim sql,rs,PrevID,NextID,id
	id=trim(Request("id"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select id,RootID,Depth,ParentID,Child,PrevID,NextID From oblog_logclass where id="&id
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���಻���ڣ������Ѿ���ɾ��</li>"
	else
		if rs("Child")>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�÷��ຬ���ӷ��࣬��ɾ�����ӷ�����ٽ���ɾ��������Ĳ���</li>"
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
	'�޸���һ�����NextID����һ�����PrevID
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	sql="select * From oblog_logclass where id=" & id
	set rsClass=server.CreateObject ("Adodb.recordset")
	rsClass.open sql,conn,1,3
	if rsClass.bof and rsClass.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���ķ��࣡</li>"
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
	
	if rsClass("ParentID")<>rParentID then   '�������������࣬��Ҫ��һϵ�м��
		if rParentID=rsClass("id") then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�������಻��Ϊ�Լ���</li>"
		end if
		'�ж���ָ���ķ����Ƿ�Ϊ�ⲿ����򱾷������������
		if rsClass("ParentID")=0 then
			if rParentID>0 then
				set trs=conn.execute("select rootid From oblog_logclass where id="&rParentID)
				if trs.bof and trs.eof then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>����ָ���ⲿ����Ϊ��������</li>"
				else
					if rsClass("rootid")=trs(0) then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>����ָ���÷��������������Ϊ��������</li>"
					end if
				end if
				trs.close
				set trs=nothing
			end if
		else
			set trs=conn.execute("select id From oblog_logclass where ParentPath like '"&rsClass("ParentPath")&"," & rsClass("id") & "%' and id="&rParentID)
			if not (trs.eof and trs.bof) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������ָ���÷��������������Ϊ��������</li>"
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
	
	
  '�����������������
  '��Ҫ������ԭ������������Ϣ��������ȡ�����ID�������������򡢼̳а���������
  '��Ҫ���µ�ǰ����������Ϣ
  '�̳а���������Ҫ��д�������и���--ȡ������ǰ̨����id in ParentPath�����
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
  if clng(parentid)<>rParentID and not (iParentID=0 and rParentID=0) then  '�����������������
	'����ԭ��ͬһ���������һ�������NextID����һ�������PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	if iParentID>0 and rParentID=0 then  	'���ԭ������һ������ĳ�һ������
		'�õ���һ��һ���������
		sql="select id,NextID From oblog_logclass where RootID=" & MaxRootID & " and Depth=0"
		set rs=server.CreateObject("Adodb.recordset")
		rs.open sql,conn,1,3
		PrevID=rs(0)      '�õ��µ�PrevID
		rs(1)=id     '������һ��һ����������NextID��ֵ
		rs.update
		rs.close
		set rs=nothing
		
		MaxRootID=MaxRootID+1
		'���µ�ǰ��������
		conn.execute("update oblog_logclass set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 where id="&id)
		'������������࣬������������������ݡ���������������迼�ǣ�ֻ���������������Ⱥ�һ������ID(rootid)����
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
		
		'������ԭ����������ķ������������൱�ڼ�֦�����迼��
		conn.execute("update oblog_logclass set child=child-1 where id="&iParentID)
		
	elseif iParentID>0 and rParentID>0 then    '����ǽ�һ���ַ����ƶ��������ַ�����
		'�õ���ǰ����������ӷ�����
		ParentPath=ParentPath & ","
		set rs=conn.execute("select count(*) From oblog_logclass where ParentPath like '%"&ParentPath & id&"%'")
		ClassCount=rs(0)
		if isnull(ClassCount) then
			ClassCount=1
		end if
		rs.close
		set rs=nothing
		
		'���Ŀ�����������Ϣ		
		set trs=conn.execute("select * From oblog_logclass where id="&rParentID)
		if trs("Child")>0 then		
			'�õ��뱾����ͬ�������һ�������OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			'�õ��뱾����ͬ�������һ�������id
			sql="select id,NextID From oblog_logclass where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)    '�õ��µ�PrevID
			rs(1)=id     '������һ�������NextID��ֵ
			rs.update
			rs.close
			set rs=nothing
			
			'�õ�ͬһ�����൫�ȱ����༶������ӷ�������OrderID�������ǰһ��ֵ����������ֵ��
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
		
		'�ڻ���ƶ������ķ����������������ָ������֮��ķ�����������
		conn.execute("update oblog_logclass set OrderID=OrderID+" & ClassCount & "+1 where rootid=" & trs("rootid") & " and OrderID>" & PrevOrderID)
		
		'���µ�ǰ��������
		conn.execute("update oblog_logclass set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("id") & "',PrevID=" & PrevID & ",NextID=0 where id="&id)
		
		'������ӷ���������ӷ������ݣ����Ϊԭ���������ȼ��ϵ�ǰ������������
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
		
		'������ָ����ϼ�������ӷ�����
		conn.execute("update oblog_logclass set child=child+1 where id="&rParentID)
		
		'������ԭ������ӷ�����			
		conn.execute("update oblog_logclass set child=child-1 where id="&iParentID)
	else    '���ԭ����һ������ĳ������������������
		'�õ��ƶ��ķ�������
		set rs=conn.execute("select count(*) From oblog_logclass where rootid="&rootid)
		ClassCount=rs(0)
		rs.close
		set rs=nothing
		
		'���Ŀ�����������Ϣ		
		set trs=conn.execute("select * From oblog_logclass where id="&rParentID)
		if trs("Child")>0 then		
			'�õ��뱾����ͬ�������һ�������OrderID
			set rsPrevOrderID=conn.execute("select Max(OrderID) From oblog_logclass where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			sql="select id,NextID From oblog_logclass where ParentID=" & trs("id") & " and OrderID=" & PrevOrderID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=id
			rs.update
			set rs=nothing
			
			'�õ�ͬһ�����൫�ȱ����༶������ӷ�������OrderID�������ǰһ��ֵ����������ֵ��
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
	
		'�ڻ���ƶ������ķ����������������ָ������֮��ķ�����������
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
		'������ָ����ϼ����������		
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	else
		id=CLng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ��������PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From oblog_logclass where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'���޸���һ�����NextID����һ�����PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From oblog_logclass")
	MaxRootID=mrs(0)+1
	'�Ƚ���ǰ����������󣬰����ӷ���
	conn.execute("update oblog_logclass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'Ȼ��λ�ڵ�ǰ�������ϵķ����RootID���μ�һ����ΧΪҪ����������
	sqlOrder="select * From oblog_logclass where ParentID=0 and RootID<" & cRootID & " order by RootID desc"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,1,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '�����ǰ�����Ѿ��������棬�������ƶ�
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '�õ�Ҫ����λ�õ�RootID�������ӷ���
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
	
	'Ȼ���ٽ���ǰ���������Ƶ���Ӧλ�ã������ӷ���
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	else
		id=CLng(id)
	end if
	if cRootID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		cRootID=Cint(cRootID)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ��������PrevID,NextID
	set rs=conn.execute("select PrevID,NextID From oblog_logclass where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.close
	set rs=nothing
	'���޸���һ�����NextID����һ�����PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if

	dim mrs,MaxRootID
	set mrs=conn.execute("select max(rootid) From oblog_logclass")
	MaxRootID=mrs(0)+1
	'�Ƚ���ǰ����������󣬰����ӷ���
	conn.execute("update oblog_logclass set RootID=" & MaxRootID & " where RootID=" & cRootID)
	
	'Ȼ��λ�ڵ�ǰ�������µķ����RootID���μ�һ����ΧΪҪ�½�������
	sqlOrder="select * From oblog_logclass where ParentID=0 and idType =" & t & " And RootID>" & cRootID & " order by RootID"
	set rsOrder=server.CreateObject("adodb.recordset")
	rsOrder.open sqlOrder,conn,2,3
	if rsOrder.bof and rsOrder.eof then
		exit sub        '�����ǰ�����Ѿ��������棬�������ƶ�
	end if
	i=1
	do while not rsOrder.eof
		tRootID=rsOrder("RootID")       '�õ�Ҫ����λ�õ�RootID�������ӷ���
		
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
	
	'Ȼ���ٽ���ǰ���������Ƶ���Ӧλ�ã������ӷ���
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
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		id=CLng(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ���������֣�</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'Ҫ�ƶ��ķ�����Ϣ
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
	'���޸���һ�����NextID����һ�����PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'�͸÷���ͬ������������֮�ϵķ���------���������򣬷�ΧΪҪ����������
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
	
	'������Ҫ����ķ�������
	conn.execute("update oblog_logclass set OrderID="&tOrderID&" where id="&id)
	'������������࣬�������������������
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
		ErrMsg=ErrMsg & "<br><li>���������</li>"
		exit sub
	else
		id=Cint(id)
	end if
	if MoveNum="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>���������</li>"
		exit sub
	else
		MoveNum=Cint(MoveNum)
		if MoveNum=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�½������֣�</li>"
			exit sub
		end if
	end if

	dim sql,rs,oldorders,ii,trs,tOrderID
	'Ҫ�ƶ��ķ�����Ϣ
	set rs=conn.execute("select ParentID,OrderID,ParentPath,child,PrevID,NextID From oblog_logclass where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.close
	set rs=nothing

	'���޸���һ�����NextID����һ�����PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'�͸÷���ͬ������������֮�µķ���------���������򣬷�ΧΪҪ�½�������
	sql="select id,OrderID,child,ParentPath,PrevID,NextID From oblog_logclass where ParentID="&ParentID&" and OrderID>"&OrderID&" order by OrderID"
	set rs=server.CreateObject("adodb.recordset")
	rs.open sql,conn,1,3
	i=0      'ͬ������
	ii=0     'ͬ��������ӷ���
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
	
	'������Ҫ����ķ�������
	conn.execute("update oblog_logclass set OrderID="&OrderID+ii&" where id="&id)
	'������������࣬�������������������
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
	
	response.Write "��λ�ɹ����뷵��<a href='admin_logclass.asp?t=" & t & "'>���������ҳ</a>������Ĺ������á�"
end sub

sub SaveUnite()
	dim id,Targetid,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	dim rs,trs,i
	id=trim(request("id"))
	Targetid=trim(request("Targetid"))
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ϲ��ķ��࣡</li>"
	else
		id=CLng(id)
	end if
	if Targetid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ����࣡</li>"
	else
		Targetid=CLng(Targetid)
	end if
	if id=Targetid then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�벻Ҫ����ͬ�����ڽ��в���</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	'�ж�Ŀ������Ƿ����ӷ��࣬����У��򱨴�
	set rs=conn.execute("select Child From oblog_logclass where id=" & Targetid)
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>Ŀ����಻���ڣ������Ѿ���ɾ����</li>"
	else
		if rs(0)>0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>Ŀ������к����ӷ��࣬���ܺϲ���</li>"
		end if
	end if
	if FoundErr=True then
		exit sub
	end if

	'�õ���ǰ������Ϣ
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
	
	'�ж��Ƿ��Ǻϲ���������������
	set rs=conn.execute("select id From oblog_logclass where id="&Targetid&" and ParentPath like '"&ParentPath&"%'")
	if not (rs.eof and rs.bof) then
		response.Write "<br><li>���ܽ�һ������ϲ����������ӷ�����</li>"
		exit sub
	end if
	
	'�õ���ǰ�������������ID
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
	
	'���޸���һ�����NextID����һ�����PrevID
	if PrevID>0 then
		conn.execute "update oblog_logclass set NextID=" & NextID & " where id=" & PrevID
	end if
	if NextID>0 then
		conn.execute "update oblog_logclass set PrevID=" & PrevID & " where id=" & NextID
	end if
	
	'����log��������
	conn.execute("update [oblog_log] set classid="&Targetid&" where classid in ("&ParentPath&")")
	
	'ɾ�����ϲ����༰����������
	conn.execute("delete From oblog_logclass where id in ("&ParentPath&")")
	
	'������ԭ������������ӷ������������൱�ڼ�֦�����迼��
	if Depth>0 then
		conn.execute("update oblog_logclass set Child=Child-1 where id="&iParentID)
	end if
	
	response.Write "����ϲ��ɹ����Ѿ������ϲ����༰�������ӷ������������ת��Ŀ������С�<br><br>ͬʱɾ���˱��ϲ��ķ��༰���ӷ��ࡣ"
	set rs=nothing
	set trs=nothing
end sub

sub Admin_ShowClass_Option(ShowType,CurrentID)
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
						't="0"   tName="����     /���
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
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
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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



Sub ClearBinding()'ǿ�������״̬��
	dim sql,rs,id
	
	
	id=trim(Request("id"))
	
	
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
