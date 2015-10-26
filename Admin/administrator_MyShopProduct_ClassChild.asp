<!--#include file="inc/inc_sys.asp"-->
<%
Dim Action,ParentID,i,FoundErr,ErrMsg,t,tName
Dim SkinCount,LayoutCount
Action 		= Trim(Request("Action"))
ParentID 	= Trim(Request("ParentID"))
t 			= Trim(Request("t"))
If ParentID="" Then
	ParentID = 0
Else
	ParentID = CLng(ParentID)
End If

If t="" Then
	t = 0
	tName = "产品 小"
Else
	t = Trim(t)
	tName = "产品 小"	'如果有参数则接收参数！
End If

t = "0"		'默认为非相册模式；

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
    	<a href="administrator_MyShopProduct_ClassChild.asp?t=<%=t%>"><%=tName%>分类管理首页</a> | 
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Add&t=<%=t%>">添加<%=tName%>分类</a>&nbsp;|&nbsp;
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Order&t=<%=t%>">一级分类排序</a>&nbsp;|&nbsp;
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=OrderN&t=<%=t%>">N级分类排序</a>&nbsp;|&nbsp;
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Reset&t=<%=t%>">复位所有<%=tName%>分类</a>&nbsp;|&nbsp;
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Unite&t=<%=t%>"><%=tName%>分类合并</a></td>
  </tr>
</table><br>

<%
If Not IsObject(conn) Then link_database
If Action="Add" Then
	Call AddClass()
ElseIf Action="SaveAdd" Then
	Call SaveAdd()
ElseIf Action="Modify" Then
	Call Modify()
ElseIf Action="SaveModify" Then
	Call SaveModify()
ElseIf Action="Move" Then
	Call MoveClass()
ElseIf Action="SaveMove" Then
	Call SaveMove()
ElseIf Action="Del" Then
	Call DeleteClass()
ElseIf Action="UpOrder" Then 
	Call UpOrder() 
ElseIf Action="DownOrder" Then 
	Call DownOrder() 
ElseIf Action="Order" Then
	Call Order()
ElseIf Action="UpOrderN" Then 
	Call UpOrderN() 
ElseIf Action="DownOrderN" Then 
	Call DownOrderN() 
ElseIf Action="OrderN" Then
	Call OrderN()
ElseIf Action="Reset" Then
	Call Reset()
ElseIf Action="SaveReset" Then
	Call SaveReset()
ElseIf Action="Unite" Then
	Call Unite()
ElseIf Action="SaveUnite" Then
	Call SaveUnite()
	
'ElseIf Action="isDisplay" Then
'	Call isDisplay()
	
'ElseIf Action="ClearBinding" Then
'	Call ClearBinding()
	
Else
	Call main()
End If
If FoundErr=True Then
	Call oblog.sys_err(errmsg)
End If
''Call CloseConn() 'shiyu


Sub main()
	Dim arrShowLine(10)
	For i=0 to Ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	Dim sqlClass,rsClass,i,iDepth
	'sqlClass="Select * From Product_ClassChild Where idType=" & t & " Order By RootID,OrderID"
	sqlClass="Select classid,id,classname,classlognum,ordernum,orderid,parentid,parentpath,"&_
					"depth,rootid,child,readme,previd,nextid,idType,EName "&_

				 " From Product_ClassChild "&_
				 " Where idType=" & t & " Order By RootID, OrderID "
			 
	Set rsClass=Server.CreateObject("Adodb.RecordSet")
	rsClass.Open sqlClass,conn,1,1
%>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22" align="center"><strong><%=tName%>分类名称</strong></td>
	<td height="12" align="center"><strong title="请在 栏目空间管理+发布 的 ‘绑定分类’处进行绑定！" style="cursor:hand;">分类说明</strong></td>
    <td width="280" height="22" align="center"><strong>操作选项</strong></td>
  </tr>
  <% 
Do While Not rsClass.Eof 
%>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
    <td> 
      <% 
	iDepth=rsClass("Depth")
	If rsClass("NextID")>0 Then
		arrShowLine(iDepth)=True	'用数组记录下状态
	Else
		arrShowLine(iDepth)=False
	End If
	If iDepth>0 Then
	  	For i=1 to iDepth 
			If i=iDepth Then 
				If rsClass("NextID")>0 Then 
					Response.Write "<img src='images/tree_line1.gif' width='17' height='16' valign='abvmiddle'>" 
				Else 
					Response.Write "<img src='images/tree_line2.gif' width='17' height='16' valign='abvmiddle'>" 
				End If 
			Else 
				If arrShowLine(i)=True Then
					Response.Write "<img src='images/tree_line3.gif' width='17' height='16' valign='abvmiddle'>" 
				Else
					Response.Write "<img src='images/tree_line4.gif' width='17' height='16' valign='abvmiddle'>" 
				End If
			End If 
	  	Next 
	  End If 
	  If rsClass("Child")>0 Then 
	  	Response.Write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	  Else 
	  	Response.Write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	  End If 
	  
	  If rsClass("Depth")=0 Then 
	  	Response.Write "<b>" 
	  End If 
	  Response.Write "<a href='administrator_MyShopProduct_ClassChild.asp?Action=Modify&id=" & rsClass("id") & "&t=" & t & "' title='" & rsClass("ReadMe")&"["& rsClass("classid") & "]'>" & rsClass("classname") & "</a>"
	  If rsClass("Child")>0 Then 
	  	Response.Write "（" & rsClass("Child") & "）" 
	  End If 
	  
	  
	  Response.Write "<font color=#666666 style='font-weight:normal;'>&nbsp;&nbsp;[&nbsp;"& rsClass("EName") &"&nbsp;]</font>"
	  
	  'Response.Write "&nbsp;&nbsp;" & rsClass("id") & "," & rsClass("PrevID") & "," & rsClass("NextID") & "," & rsClass("ParentID") & "," & rsClass("RootID")
	  %>
    </td>
	<td height="22" align="center" style="background:#FFFFFF; border:solid 1px #BFE0FB;">
	<%=rsClass("readme")%>
	
	</td>
	
    <td align="center">
	
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Add&ParentID=<%=rsClass("id")%>&t=<%=t%>">添加子分类</a> | 
    	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Modify&id=<%=rsClass("id")%>&t=<%=t%>">修改</a> | 
      	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Move&id=<%=rsClass("id")%>&t=<%=t%>">移动分类</a> | 
      	<a href="administrator_MyShopProduct_ClassChild.asp?Action=Del&id=<%=rsClass("id")%>&t=<%=t%>" onClick="<%If rsClass("Child")>0 Then%>return ConfirmDel1();<%Else%>return ConfirmDel2();<%End If%>">删除</a>
	
	</td>
  </tr>
  <% 
	rsClass.MoveNext
Loop 
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
   If(confirm("删除分类将不能恢复！确定要删除此分类吗？"))
     return true;
   Else
     return false;
	 
}
</script>
<%
End Sub


Sub AddClass()
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="administrator_MyShopProduct_ClassChild.asp?t=<%=t%>" onSubmit="return check()">  	
    <tr> 
      <td colspan="3" align="center" class="title"><strong>添 加 分 类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属分类：</strong><br>
        不能指定为外部分类 </td>
      <td> <select name="ParentID">
          <%Call Admin_ShowClass_Option(0,ParentID)%>
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
	
	<!--<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>我要固定链接指向一个地址：</strong></font> 例如：/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100"> (不需要固定链接请留空)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>是否在前台显示：</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1">显示 &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" checked="checked">隐藏</td>
    </tr>-->
	
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name="Add" type="submit" value=" 添&nbsp;&nbsp;加 " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='administrator_MyShopProduct_ClassChild.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">         
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  If (document.form1.classname.value=="")
  {
    alert("分类名称不能为空！");
	document.form1.classname.focus();
	return false;
  }
}
</script>
<%
End Sub


Sub Modify()
	Dim id,sql,rsClass,i
	id=Trim(Request("id"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CLng(id)
	End If
	sql="Select * From Product_ClassChild Where id=" & id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	rsClass.Open sql,conn,1,3
	If rsClass.Bof And rsClass.Eof Then
		Response.Write "<br><li>找不到指定的分类！</li>"
		Response.End()
	Else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="administrator_MyShopProduct_ClassChild.asp?t=<%=t%>" onSubmit="return check()">
    <tr> 
      <td colspan="3" align="center" class="title"><strong>修 改 分 类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>所属分类：</strong><br>
        如果你想改变所属分类，请<a href="administrator_MyShopProduct_ClassChild.asp?Action=Move&id=<%=id%>&t=<%=t%>">点此移动分类</a></td>
      <td> 
        <%
	If rsClass("ParentID")<=0 Then
	  	Response.Write "无（作为一级分类）"
	Else
    	Dim rsParentClass,sqlParentClass
		sqlParentClass="Select * From Product_ClassChild Where id in ("& rsClass("ParentPath") &") Order By Depth"
		Set rsParentClass=Server.CreateObject("Adodb.RecordSet")
		rsParentClass.Open sqlParentClass,conn,1,1
		Do While Not rsParentClass.Eof
			For i=1 to rsParentClass("Depth")
				Response.Write "&nbsp;&nbsp;&nbsp;"
			Next
			If rsParentClass("Depth")>0 Then
				Response.Write "└"
			End If
			Response.Write "&nbsp;" & rsParentClass("classname") & "<br>"
			rsParentClass.MoveNext
		Loop
		rsParentClass.Close
		Set rsParentClass=Nothing
	End If
	%><!--</select>-->
        </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>分类名称：</strong></td>
      <td><input name="classname" type="text" value="<%=rsClass("classname")%>" size="37" maxlength="20">
	  		
			<input name="id" type="hidden" id="id" value="<%=rsClass("id")%>"></td>
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
	
	<!--<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>我要固定链接指向一个地址：</strong></font> 例如：/abc.html</td>
      <td><input name="root_link" type="text" size="37" maxlength="100" value="<%'=rsClass("root_link")%>"> (不需要固定链接请留空)</td>
    </tr>
	
	<tr class="tdbg"> 
      <td width="350"><font color="#038ad7"><strong>是否在前台显示：</strong></font></td>
      <td><input name="isDisplay" type="radio" value="1" <%' If rsClass("isDisplay")=1 Then Response.Write "checked='checked'"%>>显示 &nbsp;&nbsp;<input name="isDisplay" type="radio" value="0" <%' If rsClass("isDisplay")=0 Then Response.Write "checked='checked'"%>>隐藏</td>
    </tr>-->
	
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name="Submit" type="submit" value=" &nbsp;保存修改结果&nbsp; " > 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='administrator_MyShopProduct_ClassChild.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;"> 
      </td>
    </tr>
  </form>
</table>
<script language="JavaScript" type="text/JavaScript">
function check()
{
  If (document.form1.classname.value=="")
  {
    alert("分类名称不能为空！");
	document.form1.classname.focus();
	return false;
  }
}
</script>
<%
	End If
	rsClass.Close
	Set rsClass=Nothing
End Sub


Sub MoveClass()
	Dim id,sql,rsClass,i
	id=Trim(Request("id"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CLng(id)
	End If
	
	sql="Select * From Product_ClassChild Where id=" & id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	rsClass.Open sql,conn,1,3
	If rsClass.Bof And rsClass.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
	Else
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
<form name="form1" method="post" action="administrator_MyShopProduct_ClassChild.asp?&t=<%=t%>">
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
	If rsClass("ParentID")<=0 Then
	  	Response.Write "无（作为一级分类）"
	Else
    	Dim rsParent,sqlParent
		sqlParent="Select * From Product_ClassChild Where id in (" & rsClass("ParentPath") & ") Order By Depth"
		Set rsParent=Server.CreateObject("Adodb.RecordSet")
		rsParent.Open sqlParent,conn,1,1
		Do While Not rsParent.Eof
			For i=1 to rsParent("Depth")
				Response.Write "&nbsp;&nbsp;&nbsp;"
			Next
			If rsParent("Depth")>0 Then
				Response.Write "└"
			End If
			Response.Write "&nbsp;" & rsParent("classname") & "<br>"
			rsParent.MoveNext
		Loop
		rsParent.Close
		Set rsParent=Nothing
	End If
	%>
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="200"><strong>移动到：</strong><br>
        不能指定为当前分类的下属子分类<br>
        不能指定为外部分类</td>
      <td><select name="ParentID" size="2" style="height:300px;width:500px;">
	   			<%Call Admin_ShowClass_Option(0,rsClass("ParentID"))%>
		  </select>
	  </td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveMove"> 
        <input name="Submit" type="submit" value=" &nbsp;保存移动结果&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='administrator_MyShopProduct_ClassChild.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
	 </td>
   </tr>
</form>  
</table>
<%
	End If
	rsClass.Close
	Set rsClass=Nothing
End Sub


Sub Order() 
	Dim sqlClass,rsClass,i,iCount,j 
	sqlClass="Select * From Product_ClassChild Where ParentID=0 And idtype="&t&" Order By RootID" 
	Set rsClass=Server.CreateObject("Adodb.RecordSet") 
	rsClass.Open sqlClass,conn,1,1 
	iCount=rsClass.recordcount 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>一 级 分 类 排 序</strong></td> 
  </tr> 
  <% 
j=1 
Do While Not rsClass.Eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="200">&nbsp;<%=rsClass("classname")%></td> 
<% 
	If j>1 Then 
  		Response.Write "<form action='administrator_MyShopProduct_ClassChild.asp?Action=UpOrder&t=" & t & "' method='post'><td width='150'>" 
		Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
		For i=1 to j-1 
			Response.Write "<option value="&i&">"&i&"</option>" 
		Next 
		Response.Write "</select>" 
		Response.Write "<input type=hidden name=id value="&rsClass("id")&">"
		Response.Write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
		Response.Write "</td></form>" 
	Else 
		Response.Write "<td width='150'>&nbsp;</td>" 
	End If 
	If iCount>j Then 
  		Response.Write "<form action='administrator_MyShopProduct_ClassChild.asp?Action=DownOrder&t=" & t & "' method='post'><td width='150'>" 
		Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
		For i=1 to iCount-j 
			Response.Write "<option value="&i&">"&i&"</option>" 
		Next 
		Response.Write "</select>" 
		Response.Write "<input type=hidden name=id value="&rsClass("id")&">"
		Response.Write "<input type=hidden name=cRootID value="&rsClass("RootID")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
		Response.Write "</td></form>" 
	Else 
		Response.Write "<td width='150'>&nbsp;</td>" 
	End If 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	j=j+1 
	rsClass.MoveNext 
Loop 
%> 
</table> 
<% 
	rsClass.Close 
	Set rsClass=Nothing 
End Sub 


Sub OrderN() 
	Dim sqlClass,rsClass,i,iCount,trs,UpMoveNum,DownMoveNum 
	sqlClass="Select * From Product_ClassChild Where idtype="&t&" Order By RootID,OrderID" 
	Set rsClass=Server.CreateObject("Adodb.RecordSet") 
	rsClass.Open sqlClass,conn,1,1 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
	<tr>
	  <td colspan="4" align="center" class="title"><strong>N 级 分 类 排 序</strong></td> 
  </tr> 
  <% 
Do While Not rsClass.Eof 
%> 
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;"> 
      <td width="300"> 
	  <% 
	For i=1 to rsClass("Depth") 
	  	Response.Write "&nbsp;&nbsp;&nbsp;" 
	Next 
	If rsClass("Child")>0 Then 
		Response.Write "<img src='Images/tree_folder4.gif' width='15' height='15' valign='abvmiddle'>" 
	Else 
	  	Response.Write "<img src='Images/tree_folder3.gif' width='15' height='15' valign='abvmiddle'>" 
	End If 
	If rsClass("ParentID")=0 Then 
		Response.Write "<b>" 
	End If 
	Response.Write rsClass("classname") 
	If rsClass("Child")>0 Then 
		Response.Write "(" & rsClass("Child") & ")" 
	End If 
	%></td> 
<% 
	If rsClass("ParentID")>0 Then   '如果不是一级分类，则算出相同深度的分类数目，得到该分类在相同深度的分类中所处位置（之上或者之下的分类数） 
		'所能提升最大幅度应为For i=1 to 该版之上的版面数 
		Set trs=conn.Execute("Select Count(id) From Product_ClassChild Where ParentID="&rsClass("ParentID")&" And OrderID<"&rsClass("OrderID")&"") 
		UpMoveNum=trs(0) 
		If isNull(UpMoveNum) Then UpMoveNum=0 
		If UpMoveNum>0 Then 
  			Response.Write "<form action='administrator_MyShopProduct_ClassChild.asp?Action=UpOrderN&t=" & t & "' method='post'><td width='150'>" 
			Response.Write "<select name=MoveNum size=1><option value=0>向上移动</option>" 
			For i=1 to UpMoveNum 
				Response.Write "<option value="&i&">"&i&"</option>" 
			Next 
			Response.Write "</select>" 
			Response.Write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
			Response.Write "</td></form>" 
		Else 
			Response.Write "<td width='150'>&nbsp;</td>" 
		End If 
		trs.Close 
		'所能降低最大幅度应为For i=1 to 该版之下的版面数 
		Set trs=conn.Execute("Select Count(id) From Product_ClassChild Where ParentID="&rsClass("ParentID")&" And orderID>"&rsClass("orderID")&"") 
		DownMoveNum=trs(0) 
		If isNull(DownMoveNum) Then DownMoveNum=0 
		If DownMoveNum>0 Then 
  			Response.Write "<form action='administrator_MyShopProduct_ClassChild.asp?Action=DownOrderN&t=" & t & "'' method='post'><td width='150'>" 
			Response.Write "<select name=MoveNum size=1><option value=0>向下移动</option>" 
			For i=1 to DownMoveNum 
				Response.Write "<option value="&i&">"&i&"</option>" 
			Next 
			Response.Write "</select>" 
			Response.Write "<input type=hidden name=id value="&rsClass("id")&">&nbsp;<input type=submit name=Submit value=修&nbsp;改 style='cursor: hand;background-color: #cccccc;'>" 
			Response.Write "</td></form>" 
		Else 
			Response.Write "<td width='150'>&nbsp;</td>" 
		End If 
		trs.Close 
	Else 
		Response.Write "<td colspan=2>&nbsp;</td>" 
	End If 
%> 
      <td>&nbsp;</td>
  </tr> 
  <% 
	UpMoveNum=0 
	DownMoveNum=0 
	rsClass.MoveNext 
Loop 
%> 
</table> 
<% 
	rsClass.Close 
	Set rsClass=Nothing 
End Sub 


Sub Reset() 
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
  <form name="form1" method="post" action="administrator_MyShopProduct_ClassChild.asp?Action=SaveReset&t=<%=t%>"> 
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
		<input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='administrator_MyShopProduct_ClassChild.asp?&t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">
      </td>
    </tr>
	</form>
</table>
<%
End Sub


Sub Unite()
%>
<table cellpadding="0" cellspacing="1" border="0" width="98%" class="border" align=center>
<form name="myform" method="post" action="administrator_MyShopProduct_ClassChild.asp?t=<%=t%>" onSubmit="return ConfirmUnite();">
	<tr>
	  <td colspan="3" align="center" class="title"><strong>分 类 合 并</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;将分类 
        <select name="id" id="id">
        <%Call Admin_ShowClass_Option(1,0)%>
        </select>
        合并到
        <select name="Targetid" id="Targetid">
        <%Call Admin_ShowClass_Option(4,0)%>
        </select>
		</td>
		</tr>
  <tr class="tdbg"> 
    <td align="center">
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input name="Action" type="hidden" id="Action" value="SaveUnite">
        <input type="submit" name="Submit" value=" &nbsp;合并分类&nbsp; " style="cursor: hand;background-color: #cccccc;">
        &nbsp;&nbsp; 
        <input name="Cancel" type="button" id="Cancel" value=" 取&nbsp;&nbsp;消 " onClick="window.location.href='administrator_MyShopProduct_ClassChild.asp?t=<%=t%>'" style="cursor: hand;background-color: #cccccc;">

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
  If (document.myform.id.value==document.myform.Targetid.value)
  {
    alert("请不要在相同分类内进行操作！");
	document.myform.Targetid.focus();
	return false;
  }
  If (document.myform.Targetid.value=="")
  {
    alert("目标分类不能指定为含有子分类的分类！");
	document.myform.Targetid.focus();
	return false;
  }
}
</script>
<% 
End Sub 
%> 
</body> 
</html> 
<% 

Sub SaveAdd()
	Dim id,classname,EName,Readme,PrevOrderID
	Dim sql,rs,trs
	Dim RootID,ParentDepth,ParentPath,ParentStr,ParentName,Maxid,MaxRootID
	Dim PrevID,NextID,Child
'	Dim root_link,isDisplay

	classname=Trim(Request("classname"))
	EName=Trim(Request("EName"))
	Readme=Trim(Request("Readme"))
	
'	root_link= Trim(Request("root_link"))
'	isDisplay= Trim(Request("isDisplay"))
	
	
	If classname="" Then
		Response.Write "<br><li>分类名称不能为空！</li>"
		Response.End()
	End If
'	If Len(root_link)>100 Then
'		Response.Write "<br><li>固定链接字符超过了100位长度！</li>"
'		Response.End()
'	End If
	
	Set rs = conn.Execute("Select Max(id) From Product_ClassChild")
	Maxid=rs(0)
	If isNull(Maxid) Then
		Maxid=0
	End If
	rs.Close
	id=Maxid+1
	Set rs=conn.Execute("Select Max(rootid) From Product_ClassChild")
	MaxRootID=rs(0)
	If isNull(MaxRootID) Then
		MaxRootID=0
	End If
	rs.Close
	RootID=MaxRootID+1
	
	If ParentID>0 Then
		sql="Select * From Product_ClassChild Where id=" & ParentID & ""
		rs.Open sql,conn,1,1
		If rs.Bof And rs.Eof Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属分类已经被删除！</li>"
		End If
		If FoundErr=True Then
			rs.Close
			Set rs=Nothing
			Exit Sub
		Else	
			RootID=rs("RootID")
			ParentName=rs("classname")
			ParentDepth=rs("Depth")
			ParentPath=rs("ParentPath")
			Child=rs("Child")
			ParentPath=ParentPath & "," & ParentID     '得到此分类的父级分类路径
			PrevOrderID=rs("OrderID")
			If Child>0 Then		
				Dim rsPrevOrderID
				'得到与本分类同级的最后一个分类的OrderID
				Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentID=" & ParentID)
				PrevOrderID=rsPrevOrderID(0)
				Set trs=conn.Execute("Select id From Product_ClassChild Where ParentID=" & ParentID & " And OrderID=" & PrevOrderID)
				PrevID=trs(0)
				
				'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
				Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentPath Like '" & ParentPath & ",%'")
				If (Not(rsPrevOrderID.Bof And rsPrevOrderID.Eof)) Then
					If Not IsNull(rsPrevOrderID(0))  Then
				 		If rsPrevOrderID(0)>PrevOrderID Then
							PrevOrderID=rsPrevOrderID(0)
						End If
					End If
				End If
			Else
				PrevID=0
			End If

		End If
		rs.Close
	Else
		If MaxRootID>0 Then
			Set trs=conn.Execute("Select id From Product_ClassChild Where RootID=" & MaxRootID & " And Depth=0")
			PrevID=trs(0)
			trs.Close
		Else
			PrevID=0
		End If
		PrevOrderID=0
		ParentPath="0"
	End If

	sql="Select * From Product_ClassChild Where ParentID=" & ParentID & " AND classname='" & classname & "'"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
	If Not(rs.Bof And rs.Eof) Then
		FoundErr=True
		If ParentID=0 Then
			ErrMsg=ErrMsg & "<br><li>已经存在一级分类：" & classname & "</li>"
		Else
			ErrMsg=ErrMsg & "<br><li>“" & ParentName & "”中已经存在子分类“" & classname & "”！</li>"
		End If
		rs.Close
		Set rs=Nothing
		Exit Sub
	End If
	rs.Close
	
	sql="Select top 1 * From Product_ClassChild"
	rs.Open sql,conn,2,2
    rs.Addnew
		rs("id")=id
		rs("classname")=classname
		
		rs("EName")=EName
		rs("RootID")=RootID
		rs("ParentID")=ParentID
		If ParentID>0 Then
			rs("Depth")=ParentDepth+1
		Else
			rs("Depth")=0
		End If
		rs("ParentPath")=ParentPath
		rs("OrderID")=PrevOrderID
		rs("Child")=0
		rs("Readme")=Readme
		rs("PrevID")=PrevID
		rs("NextID")=0
		rs("idType")=t
		
	'	rs("root_link")= root_link
	'	rs("isDisplay")= Cint(isDisplay)
	
	rs.Update
	rs.Close
    Set rs=Nothing
	
	'更新与本分类同一父分类的上一个分类的“NextID”字段值
	If PrevID>0 Then
		conn.Execute("Update Product_ClassChild Set NextID=" & id & " Where id=" & PrevID)
	End If
	
	If ParentID>0 Then
		'更新其父类的子分类数
		conn.Execute("Update Product_ClassChild Set child=child+1 Where id="&ParentID)
		
		'更新该分类排序以及大于本需要和同在本分类下的分类排序序号
		conn.Execute("Update Product_ClassChild Set OrderID=OrderID+1 Where rootid=" & rootid & " And OrderID>" & PrevOrderID)
		conn.Execute("Update Product_ClassChild Set OrderID=" & PrevOrderID & "+1 Where id=" & id)
	End If
	
    'Call CloseConn()
	'Response.Redirect "administrator_MyShopProduct_ClassChild.asp?t=" & t
	oblog.showok "保存成功","administrator_MyShopProduct_ClassChild.asp?t=" & t
End Sub


Sub SaveModify()
	Dim classname,EName,Readme,IsElite,ShowOnTop,Setting,ClassMaster,ClassPicUrl,LinkUrl,SkinID,LayoutID,BrowsePurview,AddPurview
	Dim trs,rs
	Dim id,sql,rsClass,i
	Dim SkinCount,LayoutCount
'	Dim root_link,isDisplay
	
	id=Trim(Request("id"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	Else
		id=CLng(id)
	End If
	classname=Trim(Request("classname"))
	EName=Trim(Request("EName"))
	Readme=Trim(Request("Readme"))
	
'	root_link= Trim(Request("root_link"))
'	isDisplay= Trim(Request("isDisplay"))
	
	If classname="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>分类名称不能为空！</li>"
	End If
	
'	If Len(root_link) >100 Then
'		FoundErr=True
'		ErrMsg=ErrMsg & "<br><li>固定链接字符长度不能超过100！</li>"
'	End If
	
	If FoundErr=True Then
		Exit Sub
	End If	
	sql="Select * From Product_ClassChild Where id=" & id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	rsClass.Open sql,conn,1,3
	If rsClass.Bof And rsClass.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
		rsClass.Close
		Set rsClass=Nothing
		Exit Sub
	End If

	If FoundErr=True Then
		rsClass.Close
		Set rsClass=Nothing
		Exit Sub
	End If
	
   	rsClass("classname")=classname
	rsClass("EName")=EName
	rsClass("Readme")=Readme
	
'	rsClass("root_link")= root_link
'	rsClass("isDisplay")= Cint(isDisplay)
	
	rsClass.Update
	rsClass.Close
	Set rsClass=Nothing
	
	Set rs=Nothing
	Set trs=Nothing	
    'Call CloseConn()
	'Response.Redirect "administrator_MyShopProduct_ClassChild.asp?t=" & t
	oblog.showok "保存成功","administrator_MyShopProduct_ClassChild.asp?t=" & t
End Sub


Sub DeleteClass()
	Dim sql,rs,PrevID,NextID,id
	id=Trim(Request("id"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CLng(id)
	End If
	
	sql="Select id,RootID,Depth,ParentID,Child,PrevID,NextID From Product_ClassChild Where id="&id
	Set rs=Server.CreateObject ("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	If rs.Bof And rs.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>分类不存在，或者已经被删除</li>"
	Else
		If rs("Child")>0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>该分类含有子分类，请删除其子分类后再进行删除本分类的操作</li>"
		End If
	End If
	If FoundErr=True Then
		rs.Close
		Set rs=Nothing
		Exit Sub
	End If
	PrevID=rs("PrevID")
	NextID=rs("NextID")
	If rs("Depth")>0 Then
		conn.Execute("Update Product_ClassChild Set child=child-1 Where id=" & rs("ParentID"))
	End If
	
	rs.Delete	'WL 删除
	
	rs.Update
	rs.Close
	Set rs=Nothing
	'修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If
	'Call CloseConn()
	Response.redirect "administrator_MyShopProduct_ClassChild.asp?t=" & t
		
End Sub


Sub SaveMove()
	Dim id,sql,rsClass,i
	Dim rParentID
	Dim trs,rs
	Dim ParentID,RootID,Depth,Child,ParentPath,ParentName,iParentID,iParentPath,PrevOrderID,PrevID,NextID
	id=Trim(Request("id"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		id=CLng(id)
	End If
	
	sql="Select * From Product_ClassChild Where id=" & id
	Set rsClass=Server.CreateObject ("Adodb.RecordSet")
	rsClass.Open sql,conn,1,3
	If rsClass.Bof And rsClass.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的分类！</li>"
		rsClass.Close
		Set rsClass=Nothing
		Exit Sub
	End If

	rParentID=Trim(Request("ParentID"))
	If rParentID="" Then
		rParentID=0
	Else
		rParentID=CLng(rParentID)
	End If
	
	If rsClass("ParentID")<>rParentID Then   '更改了所属分类，则要做一系列检查
		If rParentID=rsClass("id") Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>所属分类不能为自己！</li>"
		End If
		'判断所指定的分类是否为外部分类或本分类的下属分类
		If rsClass("ParentID")=0 Then
			If rParentID>0 Then
				Set trs=conn.Execute("Select rootid From Product_ClassChild Where id="&rParentID)
				If trs.Bof And trs.Eof Then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>不能指定外部分类为所属分类</li>"
				Else
					If rsClass("rootid")=trs(0) Then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>不能指定该分类的下属分类作为所属分类</li>"
					End If
				End If
				trs.Close
				Set trs=Nothing
			End If
		Else
			Set trs=conn.Execute("Select id From Product_ClassChild Where ParentPath Like '"&rsClass("ParentPath")&"," & rsClass("id") & "%' And id="&rParentID)
			If Not (trs.Eof And trs.Bof) Then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>您不能指定该分类的下属分类作为所属分类</li>"
			End If
			trs.Close
			Set trs=Nothing
		End If
		
	End If

	If FoundErr=True Then
		rsClass.Close
		Set rsClass=Nothing
		Exit Sub
	End If
	
	If rsClass("ParentID")=0 Then
		ParentID=rsClass("id")
		iParentID=0
	Else
		ParentID=rsClass("ParentID")
		iParentID=rsClass("ParentID")
	End If
	Depth=rsClass("Depth")
	Child=rsClass("Child")
	RootID=rsClass("RootID")
	ParentPath=rsClass("ParentPath")
	PrevID=rsClass("PrevID")
	NextID=rsClass("NextID")
	rsClass.Close
	Set rsClass=Nothing
	
	
  '假如更改了所属分类
  '需要更新其原来所属分类信息，包括深度、父级ID、分类数、排序、继承版主等数据
  '需要更新当前所属分类信息
  '继承版主数据需要另写函数进行更新--取消，在前台可用id in ParentPath来获得
  Dim mrs,MaxRootID
  Set mrs=conn.Execute("Select Max(rootid) From Product_ClassChild")
  MaxRootID=mrs(0)
  Set mrs=Nothing
  If isNull(MaxRootID) Then
	MaxRootID=0
  End If
  Dim k,nParentPath,mParentPath
  Dim ParentSql,ClassCount
  Dim rsPrevOrderID
  If clng(parentid)<>rParentID And Not (iParentID=0 And rParentID=0) Then  '假如更改了所属分类
	'更新原来同一父分类的上一个分类的NextID和下一个分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If
	
	If iParentID>0 And rParentID=0 Then  	'如果原来不是一级分类改成一级分类
		'得到上一个一级分类分类
		sql="Select id,NextID From Product_ClassChild Where RootID=" & MaxRootID & " And Depth=0"
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open sql,conn,1,3
		PrevID=rs(0)      '得到新的PrevID
		rs(1)=id     '更新上一个一级分类分类的NextID的值
		rs.Update
		rs.Close
		Set rs=Nothing
		
		MaxRootID=MaxRootID+1
		'更新当前分类数据
		conn.Execute("Update Product_ClassChild Set depth=0,OrderID=0,rootid="&maxrootid&",parentid=0,ParentPath='0',PrevID=" & PrevID & ",NextID=0 Where id="&id)
		'如果有下属分类，则更新其下属分类数据。下属分类的排序不需考虑，只需更新下属分类深度和一级排序ID(rootid)数据
		If child>0 Then
			i=0
			ParentPath=ParentPath & ","
			Set rs=conn.Execute("Select * From Product_ClassChild Where ParentPath Like '%"&ParentPath & id&"%'")
			Do While Not rs.Eof
				i=i+1
				mParentPath=replace(rs("ParentPath"),ParentPath,"")
				conn.Execute("Update Product_ClassChild Set depth=depth-"&depth&",rootid="&maxrootid&",ParentPath='"&mParentPath&"' Where id="&rs("id"))
				rs.MoveNext
			Loop
			rs.Close
			Set rs=Nothing
		End If
		
		'更新其原来所属分类的分类数，排序相当于剪枝而不需考虑
		conn.Execute("Update Product_ClassChild Set child=child-1 Where id="&iParentID)
		
	ElseIf iParentID>0 And rParentID>0 Then    '如果是将一个分分类移动到其他分分类下
		'得到当前分类的下属子分类数
		ParentPath=ParentPath & ","
		Set rs=conn.Execute("Select Count(*) From Product_ClassChild Where ParentPath Like '%"&ParentPath & id&"%'")
		ClassCount=rs(0)
		If isNull(ClassCount) Then
			ClassCount=1
		End If
		rs.Close
		Set rs=Nothing
		
		'获得目标分类的相关信息		
		Set trs=conn.Execute("Select * From Product_ClassChild Where id="&rParentID)
		If trs("Child")>0 Then		
			'得到与本分类同级的最后一个分类的OrderID
			Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			'得到与本分类同级的最后一个分类的id
			sql="Select id,NextID From Product_ClassChild Where ParentID=" & trs("id") & " And OrderID=" & PrevOrderID
			Set rs=Server.createobject("Adodb.RecordSet")
			rs.Open sql,conn,1,3
			PrevID=rs(0)    '得到新的PrevID
			rs(1)=id     '更新上一个分类的NextID的值
			rs.Update
			rs.Close
			Set rs=Nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentPath Like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			If (Not(rsPrevOrderID.Bof And rsPrevOrderID.Eof)) Then
				If Not IsNull(rsPrevOrderID(0))  Then
			 		If rsPrevOrderID(0)>PrevOrderID Then
						PrevOrderID=rsPrevOrderID(0)
					End If
				End If
			End If
		Else
			PrevID=0
			PrevOrderID=trs("OrderID")
		End If
		
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.Execute("Update Product_ClassChild Set OrderID=OrderID+" & ClassCount & "+1 Where rootid=" & trs("rootid") & " And OrderID>" & PrevOrderID)
		
		'更新当前分类数据
		conn.Execute("Update Product_ClassChild Set depth="&trs("depth")&"+1,OrderID="&PrevOrderID&"+1,rootid="&trs("rootid")&",ParentID="&rParentID&",ParentPath='" & trs("ParentPath") & "," & trs("id") & "',PrevID=" & PrevID & ",NextID=0 Where id="&id)
		
		'如果有子分类则更新子分类数据，深度为原来的相对深度加上当前所属分类的深度
		Set rs=conn.Execute("Select * From Product_ClassChild Where ParentPath Like '%"&ParentPath&id&"%' Order By OrderID")
		i=1
		Do While Not rs.Eof
			i=i+1
			iParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),ParentPath,"")
			conn.Execute("Update Product_ClassChild Set depth=depth-"&depth&"+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&iParentPath&"' Where id="&rs("id"))
			rs.MoveNext
		Loop
		rs.Close
		Set rs=Nothing
		trs.Close
		Set trs=Nothing
		
		'更新所指向的上级分类的子分类数
		conn.Execute("Update Product_ClassChild Set child=child+1 Where id="&rParentID)
		
		'更新其原父类的子分类数			
		conn.Execute("Update Product_ClassChild Set child=child-1 Where id="&iParentID)
	Else    '如果原来是一级分类改成其他分类的下属分类
		'得到移动的分类总数
		Set rs=conn.Execute("Select Count(*) From Product_ClassChild Where rootid="&rootid)
		ClassCount=rs(0)
		rs.Close
		Set rs=Nothing
		
		'获得目标分类的相关信息		
		Set trs=conn.Execute("Select * From Product_ClassChild Where id="&rParentID)
		If trs("Child")>0 Then		
			'得到与本分类同级的最后一个分类的OrderID
			Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentID=" & trs("id"))
			PrevOrderID=rsPrevOrderID(0)
			sql="Select id,NextID From Product_ClassChild Where ParentID=" & trs("id") & " And OrderID=" & PrevOrderID
			Set rs=Server.createobject("Adodb.RecordSet")
			rs.Open sql,conn,1,3
			PrevID=rs(0)
			rs(1)=id
			rs.Update
			Set rs=Nothing
			
			'得到同一父分类但比本分类级数大的子分类的最大OrderID，如果比前一个值大，则改用这个值。
			Set rsPrevOrderID=conn.Execute("Select Max(OrderID) From Product_ClassChild Where ParentPath Like '" & trs("ParentPath") & "," & trs("id") & ",%'")
			If (Not(rsPrevOrderID.Bof And rsPrevOrderID.Eof)) Then
				If Not IsNull(rsPrevOrderID(0))  Then
			 		If rsPrevOrderID(0)>PrevOrderID Then
						PrevOrderID=rsPrevOrderID(0)
					End If
				End If
			End If
		Else
			PrevID=0
			PrevOrderID=trs("OrderID")
		End If
	
		'在获得移动过来的分类数后更新排序在指定分类之后的分类排序数据
		conn.Execute("Update Product_ClassChild Set OrderID=OrderID+" & ClassCount &"+1 Where rootid=" & trs("rootid") & " And OrderID>" & PrevOrderID)
		
		conn.Execute("Update Product_ClassChild Set PrevID=" & PrevID & ",NextID=0 Where id=" & id)
		Set rs=conn.Execute("Select * From Product_ClassChild Where rootid="&rootid&" Order By OrderID")
		i=0
		Do While Not rs.Eof
			i=i+1
			If rs("parentid")=0 Then
				ParentPath=trs("ParentPath") & "," & trs("id")
				conn.Execute("Update Product_ClassChild Set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"',parentid="&rParentID&" Where id="&rs("id"))
			Else
				ParentPath=trs("ParentPath") & "," & trs("id") & "," & replace(rs("ParentPath"),"0,","")
				conn.Execute("Update Product_ClassChild Set depth=depth+"&trs("depth")&"+1,OrderID="&PrevOrderID&"+"&i&",rootid="&trs("rootid")&",ParentPath='"&ParentPath&"' Where id="&rs("id"))
			End If
			rs.MoveNext
		Loop
		rs.Close
		Set rs=Nothing
		trs.Close
		Set trs=Nothing
		'更新所指向的上级分类分类数		
		conn.Execute("Update Product_ClassChild Set child=child+1 Where id="&rParentID)

	End If
  End If
	
  'Call CloseConn()
  Response.Redirect "administrator_MyShopProduct_ClassChild.asp?t=" & t
End Sub


Sub UpOrder()
	Dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=Trim(Request("id"))
	cRootID=Trim(Request("cRootID"))
	MoveNum=Trim(Request("MoveNum"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	Else
		id=CLng(id)
	End If
	If cRootID="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		cRootID=Cint(cRootID)
	End If
	If MoveNum="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		MoveNum=Cint(MoveNum)
		If MoveNum=0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		End If
	End If
	If FoundErr=True Then
		Exit Sub
	End If

	'得到本分类的PrevID,NextID
	Set rs=conn.Execute("Select PrevID,NextID From Product_ClassChild Where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.Close
	Set rs=Nothing
	'先修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If

	Dim mrs,MaxRootID
	Set mrs=conn.Execute("Select Max(rootid) From Product_ClassChild")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.Execute("Update Product_ClassChild Set RootID=" & MaxRootID & " Where RootID=" & cRootID)
	
	'然后将位于当前分类以上的分类的RootID依次加一，范围为要提升的数字
	sqlOrder="Select * From Product_ClassChild Where ParentID=0 And RootID<" & cRootID & " Order By RootID desc"
	Set rsOrder=Server.CreateObject("Adodb.RecordSet")
	rsOrder.Open sqlOrder,conn,1,3
	If rsOrder.Bof And rsOrder.Eof Then
		Exit Sub        '如果当前分类已经在最上面，则无需移动
	End If
	i=1
	Do While Not rsOrder.Eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		i=i+1
		If i>MoveNum Then
			rsOrder("PrevID")=id
			rsOrder.Update
			conn.Execute("Update Product_ClassChild Set NextID=" & rsOrder("id") & " Where id=" & id)
			conn.Execute("Update Product_ClassChild Set RootID=RootID+1 Where RootID=" & tRootID)
			Exit Do
		End If
		conn.Execute("Update Product_ClassChild Set RootID=RootID+1 Where RootID=" & tRootID)
		rsOrder.MoveNext
	Loop
	rsOrder.MoveNext
	If rsOrder.Eof Then
		conn.Execute("Update Product_ClassChild Set PrevID=0 Where id=" & id)
	Else
		rsOrder("NextID")=id
		rsOrder.Update
		conn.Execute("Update Product_ClassChild Set PrevID=" & rsOrder("id") & " Where id=" & id)
	End If	
	rsOrder.Close
	Set rsOrder=Nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.Execute("Update Product_ClassChild Set RootID=" & tRootID & " Where RootID=" & MaxRootID)
	'Call CloseConn()
	Response.Redirect "administrator_MyShopProduct_ClassChild.asp?Action=Order&t=" & t
End Sub


Sub DownOrder()
	Dim id,sqlOrder,rsOrder,MoveNum,cRootID,tRootID,i,rs,PrevID,NextID
	id=Trim(Request("id"))
	cRootID=Trim(Request("cRootID"))
	MoveNum=Trim(Request("MoveNum"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	Else
		id=CLng(id)
	End If
	If cRootID="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		cRootID=Cint(cRootID)
	End If
	If MoveNum="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		MoveNum=Cint(MoveNum)
		If MoveNum=0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		End If
	End If
	If FoundErr=True Then
		Exit Sub
	End If

	'得到本分类的PrevID,NextID
	Set rs=conn.Execute("Select PrevID,NextID From Product_ClassChild Where id=" & id)
	PrevID=rs(0)
	NextID=rs(1)
	rs.Close
	Set rs=Nothing
	'先修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If

	Dim mrs,MaxRootID
	Set mrs=conn.Execute("Select Max(rootid) From Product_ClassChild")
	MaxRootID=mrs(0)+1
	'先将当前分类移至最后，包括子分类
	conn.Execute("Update Product_ClassChild Set RootID=" & MaxRootID & " Where RootID=" & cRootID)
	
	'然后将位于当前分类以下的分类的RootID依次减一，范围为要下降的数字
	sqlOrder="Select * From Product_ClassChild Where ParentID=0 And idType =" & t & " And RootID>" & cRootID & " Order By RootID"
	Set rsOrder=Server.CreateObject("Adodb.RecordSet")
	rsOrder.Open sqlOrder,conn,2,3
	If rsOrder.Bof And rsOrder.Eof Then
		Exit Sub        '如果当前分类已经在最下面，则无需移动
	End If
	i=1
	Do While Not rsOrder.Eof
		tRootID=rsOrder("RootID")       '得到要提升位置的RootID，包括子分类
		
		i=i+1
		If i>MoveNum Then
			rsOrder("NextID")=id
			rsOrder.Update
			conn.Execute("Update Product_ClassChild Set PrevID=" & rsOrder("id") & " Where id=" & id)
			conn.Execute("Update Product_ClassChild Set RootID=RootID-1 Where RootID=" & tRootID)
			Exit Do
		End If
		conn.Execute("Update Product_ClassChild Set RootID=RootID-1 Where RootID=" & tRootID)
		rsOrder.MoveNext
	Loop
	rsOrder.MoveNext
	If rsOrder.Eof Then
		conn.Execute("Update Product_ClassChild Set NextID=0 Where id=" & id)
	Else
		rsOrder("PrevID")=id
		rsOrder.Update
		conn.Execute("Update Product_ClassChild Set NextID=" & rsOrder("id") & " Where id=" & id)
	End If	
	rsOrder.Close
	Set rsOrder=Nothing
	
	'然后再将当前分类从最后移到相应位置，包括子分类
	conn.Execute("Update Product_ClassChild Set RootID=" & tRootID & " Where RootID=" & MaxRootID)
	'Call CloseConn()
	Response.Redirect "administrator_MyShopProduct_ClassChild.asp?Action=Order&t=" & t
End Sub


Sub UpOrderN()
	Dim sqlOrder,rsOrder,MoveNum,id,i
	Dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=Trim(Request("id"))
	MoveNum=Trim(Request("MoveNum"))
	If id="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		id=CLng(id)
	End If
	If MoveNum="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	Else
		MoveNum=Cint(MoveNum)
		If MoveNum=0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要提升的数字！</li>"
		End If
	End If
	If FoundErr=True Then
		Exit Sub
	End If

	Dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	Set rs=conn.Execute("Select ParentID,OrderID,ParentPath,child,PrevID,NextID From Product_ClassChild Where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.Close
	Set rs=Nothing
	If child>0 Then
		Set rs=conn.Execute("Select Count(*) From Product_ClassChild Where ParentPath Like '%"&ParentPath&"%'")
		oldorders=rs(0)
		rs.Close
		Set rs=Nothing
	Else
		oldorders=0
	End If
	'先修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If
	
	'和该分类同级且排序在其之上的分类------更新其排序，范围为要提升的数字
	sql="Select id,OrderID,child,ParentPath,PrevID,NextID From Product_ClassChild Where ParentID="&ParentID&" And OrderID<"&OrderID&" Order By OrderID desc"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	i=1
	Do While Not rs.Eof
		tOrderID=rs(1)
		
		If rs(2)>0 Then
			ii=i+1
			Set trs=conn.Execute("Select id,OrderID From Product_ClassChild Where ParentPath Like '%"&rs(3)&","&rs(0)&"%' Order By OrderID")
			If Not (trs.Eof And trs.Bof) Then
				Do While Not trs.Eof
					conn.Execute("Update Product_ClassChild Set OrderID="&tOrderID+oldorders+ii&" Where id="&trs(0))
					ii=ii+1
					trs.MoveNext
				Loop
			End If
			trs.Close
			Set trs=Nothing
		End If
		i=i+1
		If i>MoveNum Then
			rs(4)=id
			rs.Update
			conn.Execute("Update Product_ClassChild Set NextID=" & rs(0) & " Where id=" & id)
			conn.Execute("Update Product_ClassChild Set OrderID="&tOrderID+oldorders+i-1&" Where id="&rs(0))		
			Exit Do
		End If
		conn.Execute("Update Product_ClassChild Set OrderID="&tOrderID+oldorders+i-1&" Where id="&rs(0))
		rs.MoveNext
	Loop
	If Not rs.Eof Then
	rs.MoveNext
	End If
	If rs.Eof Then
		conn.Execute("Update Product_ClassChild Set PrevID=0 Where id=" & id)
	Else
		rs(5)=id
		rs.Update
		conn.Execute("Update Product_ClassChild Set PrevID=" & rs(0) & " Where id=" & id)
	End If	
	rs.Close
	Set rs=Nothing
	
	'更新所要排序的分类的序号
	conn.Execute("Update Product_ClassChild Set OrderID="&tOrderID&" Where id="&id)
	'如果有下属分类，则更新其下属分类排序
	If child>0 Then
		i=1
		Set rs=conn.Execute("Select id From Product_ClassChild Where ParentPath Like '%"&ParentPath&"%' Order By OrderID")
		Do While Not rs.Eof
			conn.Execute("Update Product_ClassChild Set OrderID="&tOrderID+i&" Where id="&rs(0))
			i=i+1
			rs.MoveNext
		Loop
		rs.Close
		Set rs=Nothing
	End If
	'Call CloseConn()
	Response.Redirect "administrator_MyShopProduct_ClassChild.asp?Action=OrderN&t=" & t
End Sub


Sub DownOrderN()
	Dim sqlOrder,rsOrder,MoveNum,id,i
	Dim ParentID,OrderID,ParentPath,Child,PrevID,NextID
	id=Trim(Request("id"))
	MoveNum=Trim(Request("MoveNum"))
	If id="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		Exit Sub
	Else
		id=Cint(id)
	End If
	If MoveNum="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
		Exit Sub
	Else
		MoveNum=Cint(MoveNum)
		If MoveNum=0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请选择要下降的数字！</li>"
			Exit Sub
		End If
	End If

	Dim sql,rs,oldorders,ii,trs,tOrderID
	'要移动的分类信息
	Set rs=conn.Execute("Select ParentID,OrderID,ParentPath,child,PrevID,NextID From Product_ClassChild Where id="&id)
	ParentID=rs(0)
	OrderID=rs(1)
	ParentPath=rs(2) & "," & id
	child=rs(3)
	PrevID=rs(4)
	NextID=rs(5)
	rs.Close
	Set rs=Nothing

	'先修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If
	
	'和该分类同级且排序在其之下的分类------更新其排序，范围为要下降的数字
	sql="Select id,OrderID,child,ParentPath,PrevID,NextID From Product_ClassChild Where ParentID="&ParentID&" And OrderID>"&OrderID&" Order By OrderID"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	i=0      '同级分类
	ii=0     '同级分类和子分类
	Do While Not rs.Eof
		'conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii&" Where id="&rs(0))
		If rs(2)>0 Then
			Set trs=conn.Execute("Select id,OrderID From Product_ClassChild Where ParentPath Like '%"&rs(3)&","&rs(0)&"%' Order By OrderID")
			If Not (trs.Eof And trs.Bof) Then
				Do While Not trs.Eof
					ii=ii+1
					conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii&" Where id="&trs(0))
					trs.MoveNext
				Loop
			End If
			trs.Close
			Set trs=Nothing
		End If
		ii=ii+1
		i=i+1
		If i>=MoveNum Then
			rs(5)=id
			rs.Update
			conn.Execute("Update Product_ClassChild Set PrevID=" & rs(0) & " Where id=" & id)
			conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii-1&" Where id="&rs(0))		
			Exit Do
		End If
		conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii-1&" Where id="&rs(0))
		rs.MoveNext
	Loop
	rs.MoveNext
	If rs.Eof Then
		conn.Execute("Update Product_ClassChild Set NextID=0 Where id=" & id)
	Else
		rs(4)=id
		rs.Update
		conn.Execute("Update Product_ClassChild Set NextID=" & rs(0) & " Where id=" & id)
	End If	
	rs.Close
	Set rs=Nothing
	
	'更新所要排序的分类的序号
	conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii&" Where id="&id)
	'如果有下属分类，则更新其下属分类排序
	If child>0 Then
		i=1
		Set rs=conn.Execute("Select id From Product_ClassChild Where ParentPath Like '%"&ParentPath&"%' Order By OrderID")
		Do While Not rs.Eof
			conn.Execute("Update Product_ClassChild Set OrderID="&OrderID+ii+i&" Where id="&rs(0))
			i=i+1
			rs.MoveNext
		Loop
		rs.Close
		Set rs=Nothing
	End If
	'Call CloseConn()
	Response.Redirect "administrator_MyShopProduct_ClassChild.asp?Action=OrderN&t=" & t
End Sub


Sub SaveReset()
	Dim i,sql,rs,SuccessMsg,iCount,PrevID,NextID
	sql="Select id From Product_ClassChild Where idType=" & t & " Order By RootID,OrderID"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
	i=1
	PrevID=0
	Do While Not rs.Eof
		rs.MoveNext
		If rs.Eof Then
			NextID=0
		Else
			NextID=rs(0)
		End If
		rs.moveprevious
		conn.Execute("Update Product_ClassChild Set RootID=" & i & ",OrderID=0,ParentID=0,Child=0,ParentPath='0',Depth=0,PrevID=" & PrevID & ",NextID=" & NextID & " Where id=" & rs(0))
		PrevID=rs(0)
		i=i+1
		rs.MoveNext
	Loop
	rs.Close
	Set rs=Nothing	
	
	Response.Write "复位成功！请返回<a href='administrator_MyShopProduct_ClassChild.asp?t=" & t & "'>分类管理首页</a>做分类的归属设置。"
End Sub


Sub SaveUnite()
	Dim id,Targetid,ParentPath,iParentPath,Depth,iParentID,Child,PrevID,NextID
	Dim rs,trs,i
	id=Trim(Request("id"))
	Targetid=Trim(Request("Targetid"))
	If id="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要合并的分类！</li>"
	Else
		id=CLng(id)
	End If
	If Targetid="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定目标分类！</li>"
	Else
		Targetid=CLng(Targetid)
	End If
	If id=Targetid Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请不要在相同分类内进行操作</li>"
	End If
	If FoundErr=True Then
		Exit Sub
	End If
	'判断目标分类是否有子分类，如果有，则报错。
	Set rs=conn.Execute("Select Child From Product_ClassChild Where id=" & Targetid)
	If rs.Bof And rs.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>目标分类不存在，可能已经被删除！</li>"
	Else
		If rs(0)>0 Then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>目标分类中含有子分类，不能合并！</li>"
		End If
	End If
	If FoundErr=True Then
		Exit Sub
	End If

	'得到当前分类信息
	Set rs=conn.Execute("Select id,ParentID,ParentPath,PrevID,NextID,Depth From Product_ClassChild Where id="&id)
	iParentID=rs(1)
	Depth=rs(5)
	If iParentID=0 Then
		ParentPath=rs(0)
	Else
		ParentPath=rs(2) & "," & rs(0)
	End If
	iParentPath=rs(0)
	PrevID=rs(3)
	NextID=rs(4)
	
	'判断是否是合并到其下属分类中
	Set rs=conn.Execute("Select id From Product_ClassChild Where id="&Targetid&" And ParentPath Like '"&ParentPath&"%'")
	If Not (rs.Eof And rs.Bof) Then
		Response.Write "<br><li>不能将一个分类合并到其下属子分类中</li>"
		Exit Sub
	End If
	
	'得到当前分类的下属分类ID
	Set rs=conn.Execute("Select id From Product_ClassChild Where ParentPath Like '"&ParentPath&"%'")
	i=0
	If Not (rs.Eof And rs.Bof) Then
		Do While Not rs.Eof
			iParentPath=iParentPath & "," & rs(0)
			i=i+1
			rs.MoveNext
		Loop
	End If
	If i>0 Then
		ParentPath=iParentPath
	Else
		ParentPath=id
	End If
	
	'先修改上一分类的NextID和下一分类的PrevID
	If PrevID>0 Then
		conn.Execute "Update Product_ClassChild Set NextID=" & NextID & " Where id=" & PrevID
	End If
	If NextID>0 Then
		conn.Execute "Update Product_ClassChild Set PrevID=" & PrevID & " Where id=" & NextID
	End If
	
	'更新log所属分类
	conn.Execute("Update [oblog_log] Set classid="&Targetid&" Where classid in ("&ParentPath&")")
	
	'删除被合并分类及其下属分类
	conn.Execute("Delete From Product_ClassChild Where id in ("&ParentPath&")")
	
	'更新其原来所属分类的子分类数，排序相当于剪枝而不需考虑
	If Depth>0 Then
		conn.Execute("Update Product_ClassChild Set Child=Child-1 Where id="&iParentID)
	End If
	
	Response.Write "分类合并成功！已经将被合并分类及其下属子分类的所有数据转入目标分类中。<br><br>同时删除了被合并的分类及其子分类。"
	Set rs=Nothing
	Set trs=Nothing
End Sub


Sub Admin_ShowClass_Option(ShowType,CurrentID)
	If ShowType=0 Then
	    Response.Write "<option value='0'"
		If CurrentID=0 Then Response.Write " selected"
		Response.Write ">无（作为一级栏目）</option>"
	End If
	Dim rsClass,sqlClass,strTemp,tmpDepth,i
	Dim arrShowLine(20)
	For i=0 to Ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	sqlClass="Select * From Product_ClassChild Where idType=" & t &" Order By RootID,OrderID"
						't="0"   tName="产品     /相册
	Set rsClass=conn.Execute(sqlClass)
	If rsClass.Bof And rsClass.Eof Then 
		Response.Write "<option value=''>请先添加栏目</option>"
	Else
		Do While Not rsClass.Eof
			tmpDepth=rsClass("Depth")
			If rsClass("NextID")>0 Then
				arrShowLine(tmpDepth)=True
			Else
				arrShowLine(tmpDepth)=False
			End If
				strTemp="<option value='" & rsClass("id") & "'"		'WL
			If CurrentID>0 And rsClass("id")=CurrentID Then		'WL
				 strTemp=strTemp & " selected"
			End If
			strTemp=strTemp & ">"
			
			If tmpDepth>0 Then
				For i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					If i=tmpDepth Then
						If rsClass("NextID")>0 Then
							strTemp=strTemp & "├&nbsp;"
						Else
							strTemp=strTemp & "└&nbsp;"
						End If
					Else
						If arrShowLine(i)=True Then
							strTemp=strTemp & "│"
						Else
							strTemp=strTemp & "&nbsp;"
						End If
					End If
				Next
			End If
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			Response.Write strTemp
			rsClass.MoveNext
		Loop
	End If
	rsClass.Close
	Set rsClass=Nothing
End Sub


Sub Admin_ShowClass_Option_classid(ShowType,CurrentID)
	If ShowType=0 Then
	    Response.Write "<option value='0'"
		If CurrentID=0 Then Response.Write " selected"
		Response.Write ">无（作为一级栏目）</option>"
	End If
	Dim rsClass,sqlClass,strTemp,tmpDepth,i
	Dim arrShowLine(20)
	For i=0 to Ubound(arrShowLine)
		arrShowLine(i)=False
	Next
	sqlClass="Select * From Product_ClassChild Where idType=" & t &" Order By RootID,OrderID"
						't="0"   tName="产品     /相册
	Set rsClass=conn.Execute(sqlClass)
	If rsClass.Bof And rsClass.Eof Then 
		Response.Write "<option value=''>请先添加栏目</option>"
	Else
		Do While Not rsClass.Eof
			tmpDepth=rsClass("Depth")
			If rsClass("NextID")>0 Then
				arrShowLine(tmpDepth)=True
			Else
				arrShowLine(tmpDepth)=False
			End If
				strTemp="<option value='" & rsClass("classid") & "'"		'WL
			If CurrentID>0 And rsClass("classid")=CurrentID Then		'WL
				 strTemp=strTemp & " selected"
			End If
			strTemp=strTemp & ">"
			
			If tmpDepth>0 Then
				For i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					If i=tmpDepth Then
						If rsClass("NextID")>0 Then
							strTemp=strTemp & "├&nbsp;"
						Else
							strTemp=strTemp & "└&nbsp;"
						End If
					Else
						If arrShowLine(i)=True Then
							strTemp=strTemp & "│"
						Else
							strTemp=strTemp & "&nbsp;"
						End If
					End If
				Next
			End If
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			Response.Write strTemp
			rsClass.MoveNext
		Loop
	End If
	rsClass.Close
	Set rsClass=Nothing
End Sub


'
'Sub isDisplay()
'	Dim sql,rs,classid
'	Dim isDisplay
'	
'	classid=Trim(Request("classid"))
'	isDisplay= Cint(Request("isDisplay"))
'	
'	If classid="" Then
'		FoundErr=True
'		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
'		Exit Sub
'	Else
'		classid=CLng(classid)
'	End If
'	
'
'	If FoundErr=True Then
'		rs.Close
'		Set rs=Nothing
'		Exit Sub
'	End If
'
'	If isDisplay= 1 Then
'		conn.Execute("Update Product_ClassChild Set isDisplay=0 Where classid=" & classid )
'	ElseIf isDisplay= 0 Then
'		conn.Execute("Update Product_ClassChild Set isDisplay=1 Where classid=" & classid )
'	End If
'
'	Response.redirect "administrator_MyShopProduct_ClassChild.asp"
'		
'End Sub
'
'
'
'Sub ClearBinding()'强制清除绑定状态；
'	Dim sql,rs,id
'	
'	
'	id=Trim(Request("id"))
'	
'	
'	If id="" Then
'		FoundErr=True
'		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
'		Exit Sub
'	Else
'		id=CLng(id)
'	End If
'	
'
'	If FoundErr=True Then
'		rs.Close
'		Set rs=Nothing
'		Exit Sub
'	End If
'
'	
'		conn.Execute("Update Product_ClassChild Set userid_Nav=0 Where classid=" & id )
'
'
'	Response.redirect "administrator_MyShopProduct_ClassChild.asp"
'		
'End Sub

%> 

