<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
Const MaxPerPage=20
Dim strFileName
Dim totalPut,TotalPages
Dim rs, sql
Dim CoverID,UserSearch,Keyword,strField
Dim Action,FoundErr,ErrMsg
Dim tmpDays
keyword=Trim(Request("keyword"))
If keyword<>"" Then 
	keyword=oblog.filt_badstr(keyword)
End If
strField=Trim(Request("Field"))
UserSearch=Trim(Request("UserSearch"))
Action=Trim(Request("Action"))
CoverID=Trim(Request("CoverID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

If UserSearch="" Then
	UserSearch=0
Else
	UserSearch=Clng(UserSearch)
End If
strFileName="Administrator_CoverProducts.asp?UserSearch=" & UserSearch
If strField<>"" Then
	strFileName=strFileName&"&Field="&strField
End If
If keyword<>"" Then
	strFileName=strFileName&"&keyword="&keyword
End If
If Request("page")<>"" Then
    currentPage=Cint(Request("page"))
Else
	currentPage=1
End If

%>
<html>
<head>
<title>注册促销礼包管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function unselectall()
{
    If(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    If (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>促 销 礼 包 管 理</strong></td>
  </tr>
  <form name="form1" action="Administrator_CoverProducts.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>快速查找礼包：</strong></td>
      <td width="687" height="30">
	   <select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="0">最后注册的500个礼包</option>
    
		  <option value="3">VIP礼包</option>

        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="Administrator_CoverProducts.asp">促销礼包管理首页</a>&nbsp;|&nbsp;<a href="?Action=Add">添加新礼包</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="Administrator_CoverProducts.asp">
  <tr class="tdbg">
    <td width="120"><strong>礼包高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
		  <!--<option value="UserName" Selected>礼包名</option>-->
		  <option value="CoverName" Selected>按礼包名</option>
		  <option value="CoverID" >按礼包ID</option>
		  <option value="address" >按地址</option>
		  <option value="beizhu" >按备注</option>
		  <option value="xuwei_beizhu" >按明细</option>

      </select>
	  
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  若为空，则查询所有礼包</td>
  </tr>
</form>
</table>
<%
If Action="Add" Then
	Call AddCover()
ElseIf Action="SaveAdd" Then
	Call SaveAdd()
ElseIf Action="Modify" Then
	Call Modify()
ElseIf Action="SaveModify" Then
	Call SaveModify()
ElseIf Action="Del" Then
	Call DelUser()
ElseIf Action="Lock" Then
	Call LockUser()
ElseIf Action="UnLock" Then
	Call UnLockUser()
ElseIf Action="Move" Then
	Call MoveUser()
ElseIf Action="Update" Then
	Call UpdateUser()
ElseIf Action="DoUpdate" Then
	Call DoUpdate()
ElseIf Action="DoUpdatelog" Then
	Call DoUpdatelog()
ElseIf Action="gouser1" Then
	Call gouser1()
ElseIf Action="gouser2" Then
	Call gouser2()
ElseIf Action="ShowInfo" Then
	Call ShowInfo()


ElseIf Action="Discontinued" Then
	Call Discontinued()
ElseIf Action="UnDiscontinued" Then
	Call UnDiscontinued()
	
	
Else
	Call main()
End If

If FoundErr=True Then
	Call WriteErrMsg()
End If


Sub main()
	Dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='Administrator_CoverProducts.asp'>注册促销礼包管理</a>&nbsp;&gt;&gt;&nbsp;"
	Select Case UserSearch
		Case 0
			sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 Order By CoverID Desc"
			strGuide=strGuide & "最后注册的500个礼包"
		Case 1
			sql="Select Top 100 * From Product_CoverProducts Where Discontinued=0 Order By log_count Desc"
			strGuide=strGuide & "发布文章最多的前100个礼包"
		Case 2
			sql="Select Top 100 * From  Product_CoverProducts Where Discontinued=0 Order By log_count"
			strGuide=strGuide & "发布文章最少的100个礼包"
		Case 3
			sql="Select * From  Product_CoverProducts Where Discontinued=1 Order By CoverID Desc"
			strGuide=strGuide & "前500个放弃的礼包"
		Case 4
			sql="Select * From  Product_CoverProducts Where Discontinued=0 And user_level=9 Order By CoverID Desc"
			strGuide=strGuide & "前台管理员"
		Case 5
			sql="Select * From  Product_CoverProducts Where Discontinued=0 And user_isbest=1 Order By CoverID Desc"
			strGuide=strGuide & "推荐礼包"
		Case 6
			sql="Select * From Product_CoverProducts Where Discontinued=0 And User_Level=6 Order By CoverID  Desc"
			strGuide=strGuide & "等待管理认证证的礼包"
		Case 7
			sql="Select * From Product_CoverProducts Where Discontinued=0 And  LockUser =1 Order By CoverID  Desc"
			strGuide=strGuide & "所有被锁住的礼包"
		Case 10
			If Keyword="" Then
				sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 Order By CoverID Desc"
				strGuide=strGuide & "所有礼包"
			Else
				Select Case strField
				Case "CoverID"

					If IsNumeric(Keyword)=false Then
						FoundErr=True
						ErrMsg=ErrMsg & "<br><li>礼包ID必须是整数！</li>"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverID =" & Clng(Keyword)
						strGuide=strGuide & "礼包ID等于<font color=red> " & Clng(Keyword) & " </font>的礼包"
					End If
				Case "CoverName"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverName like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "礼包名中含有“ <font color=red>" & Keyword & "</font> ”的礼包"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And CoverName= '" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "礼包名等于“ <font color=red>" & Keyword & "</font> ”的礼包"
					End If
					
				Case "address"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And address like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "通讯地址中含有“ <font color=red>" & Keyword & "</font> ”的礼包"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And address='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "通讯地址等于“ <font color=red>" & Keyword & "</font> ”的礼包"
					End If
				Case "beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And beizhu like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "备注中含有“ <font color=red>" & Keyword & "</font> ”的礼包"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And beizhu='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "备注等于“ <font color=red>" & Keyword & "</font> ”的礼包"
					End If
				Case "xuwei_beizhu"
					If is_sqldata=1 Then
						sql="Select * From Product_CoverProducts Where Discontinued=0 And xuwei_beizhu like '%" & Keyword & "%' Order By CoverID  Desc"
						strGuide=strGuide & "明细中含有“ <font color=red>" & Keyword & "</font> ”的礼包"
					Else
						sql="Select * From Product_CoverProducts Where Discontinued=0 And xuwei_beizhu='" & Keyword & "' Order By CoverID  Desc"
						strGuide=strGuide & "明细等于“ <font color=red>" & Keyword & "</font> ”的礼包"
					End If
				Case "logcount"
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And log_count < " & Clng(Keyword) & " Order By CoverID  Desc"
					strGuide=strGuide & "文章数小于“ <font color=red>" & Keyword & "</font> ”的礼包"
				Case "logintimes"
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And logintimes < " & Clng(Keyword) & " Order By CoverID  Desc"
					strGuide=strGuide & "登录次数小于“ <font color=red>" & Keyword & "</font> ”的礼包"
				Case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And datediff(d,lastlogintime,getdate())>"&Clng(Keyword)&" Order By CoverID  Desc"
				   Else 
				   	sql="Select Top 500 * From Product_CoverProducts Where Discontinued=0 And datediff('d',lastlogintime,now())>"&Clng(Keyword)&" Order By CoverID  Desc"
				   End If 
					strGuide=strGuide & "“ <font color=red>" & Keyword & "</font> ”天内未登录的礼包"
				End Select
			End If
		Case Else
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	End Select
	strGuide=strGuide & "</td><td align='right'>"
	If FoundErr=True Then Exit Sub
	If Not IsObject(conn) Then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
'	Response.Write sql
'	Response.End 
	rs.Open sql,Conn,1,1
  	If rs.Eof And rs.Bof Then
		strGuide=strGuide & "共找到 <font color=red>0</font> 个礼包</td></tr></table>"
		Response.Write strGuide
	Else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个礼包</td></tr></table>"
		Response.Write strGuide
		If currentpage<1 Then
       		currentpage=1
    	End If
    	If (currentpage-1)*MaxPerPage>totalput Then
	   		If (totalPut mod MaxPerPage)=0 Then
	     		currentpage= totalPut \ MaxPerPage
		  	Else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		End If

    	End If
	    If currentPage=1 Then
        	showContent
        	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"个礼包")
   	 	Else
   	     	If (currentPage-1)*MaxPerPage<totalPut Then
         	   	rs.Move  (currentPage-1)*MaxPerPage
         		Dim Bookmark
           		Bookmark=rs.Bookmark
            	showContent
            	Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"个礼包")
        	Else
	        	currentPage=1
           		showContent
           		Response.Write oblog.showpage(strFileName,totalput,MaxPerPage,True,True,"个礼包")
	    	End If
		End If
	End If
	rs.Close
	Set rs=Nothing
End Sub


Sub showContent()
   	Dim i
    i=0
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="Administrator_CoverProducts.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="39" align="center"><font color="#FFFFFF">选中</font></td>
            <td width="34" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="181" height="22" align="center"><font color="#FFFFFF">礼包名称</font></td>
            <td width="225" height="22" align="center"><font color="#FFFFFF">礼包描述</font></td>
            <td width="130" align="center"><font color="#FFFFFF">包含的产品IDs (多选)</font></td>
			<td width="125" height="22" align="center"><font color="#FFFFFF">所属大类型 (多选)</font></td>
            
			<td width="85" align="center"><font color="#cccccc">最大可选数量</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF">总体优惠价格</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF">操作</font></td>
          </tr>
          <%do while Not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td  align="center"><input name='CoverID' type='checkbox' onClick="unselectall()" id="CoverID" value='<%=Cstr(rs("CoverID"))%>' /></td>
            <td  align="center"><%=rs("CoverID")%></td>
            <td  align="center"><%
			Response.Write "<a href='?action=Modify&CoverID="&rs("CoverID")&"'" 			
			Response.Write """  title='[礼包图片地址:"&rs("CoverPhoto")&"]'>" & rs("CoverName") & "</a>"
			%> </td>
            <td align="center" title="<%=rs("CoverDesc")%>"> <%=Left(rs("CoverDesc") ,20)%> ...</td>
            
			<td align="center"> <%=rs("Iclude_ProductIDs")%></td>
			
			<td align="center"> <%

		Response.Write rs("ClassIDs")& "&nbsp;"
	

	%> </td>
            
			<td align="center"> <%=rs("MaxNumber")%> 件套</td>
			
            <td align="center"><%=FormatCurrency(rs("CoverPrice"))%></td>
            <td align="center">
			

			
			<%
			
			
			If rs("Discontinued")=1 Then
				Response.Write "<a href='Administrator_CoverProducts.asp?Action=Discontinued&CoverID=" & rs("CoverID") & "' title=''>"
				Response.Write("<font color=red title='此礼包已经停止促销，我现在想开始促销！'>已停止</font>")
			ElseIf rs("Discontinued")=0 Then
				Response.Write "<a href='Administrator_CoverProducts.asp?Action=UnDiscontinued&CoverID=" & rs("CoverID") & "' title=''>"
				Response.Write("<font color=green title='此礼包正在热销中，我现在想停止其促销！'>促销中</font>")
			End If
			Response.Write "</font></a>&nbsp;"
			
		Response.Write "<a href='Administrator_CoverProducts.asp?Action=Modify&CoverID=" & rs("CoverID") & "'>修改</a>&nbsp;"
		
		'Response.Write "<a href='Administrator_CoverProducts.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>进后台</a>&nbsp;"

        Response.Write "<a href='Administrator_CoverProducts.asp?Action=Del&CoverID=" & rs("CoverID") & "' onClick='return confirm(""确定要删除此礼包吗？"");'>删除</a>&nbsp;"
		
		%> </td>
          </tr>
          <%
	i=i+1
	If i>=MaxPerPage Then Exit do
	rs.movenext
loop
%>
        </table>  
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
              选中本页显示的所有礼包</td>
            <td> <strong>操作：</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=True" />
              删除&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">移动到
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">促销礼包</option>
                <option value="7">注册栏目发布</option>
                <option value="8">VIP栏目发布</option>
                <option value="9">前台管理发布</option>
              </select>-->
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" 执 行 " /> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
End Sub




Sub AddCover()

%>
<FORM name="Form1" Action="Administrator_CoverProducts.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">注册新 促销礼包</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">礼包名：</TD>
      <TD width="60%"><input name="CoverName" type="text" value="" size=20 maxlength=50 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>礼包描述：</td>
      <td><textarea name="CoverDesc" cols="38" rows="8"></textarea></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>包含的产品有：(多选)</td>
      <td><!--<Select name="Iclude_ProductIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_ClassChild")	大型多选select控件.%>
			
		  </select>-->
	  		<input name="Iclude_ProductIDs" type="text" value="" size=6 maxlength=30 />
	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>所属大类型：(多选)</td>
      <td><!--<Select name="ClassIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_Class")	大型多选select控件.%>
			
		  </select>-->
			<input name="ClassIDs" type="text" value="" size=6 maxlength=30 />  
	  </td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>本礼包最大可选商品数量：</td>
      <td><input name="MaxNumber" type="text" value="" size=3 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>总体优惠价格：</td>
      <td><input name="CoverPrice" type="text" value="" size=10 maxlength=20 style="text-align:center;" />元RMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>礼包的展示图片：</TD>
      <TD><input name="CoverPhoto" type="text" value="" size=50 maxlength=500 /> 
	  
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">本礼包e促销状态：</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" checked="checked" />促销中&nbsp;&nbsp;&nbsp;<INPUT name="Discontinued" type=radio value="1" />停止促销
      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="保存新礼包"></TD>
    </TR>
  </TABLE>
</form>
<%

End Sub




Sub Modify()
	Dim CoverID
	Dim rsCover,sqlCover
	CoverID=Trim(Request("CoverID"))
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		CoverID=Clng(CoverID)
	End If
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts Where CoverID=" & CoverID
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3
	If rsCover.Bof And rsCover.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的礼包！</li>"
		rsCover.Close
		Set rsCover=Nothing
		Exit Sub
	End If
%>
<FORM name="Form1" Action="Administrator_CoverProducts.asp" method="post">
  <table width="500" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">注册新 促销礼包</font></b></TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">礼包名：</TD>
      <TD width="60%"><input name="CoverName" type="text" value="<%=rsCover("CoverName")%>" size=20 maxlength=50 /></TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>礼包描述：</td>
      <td><textarea name="CoverDesc" cols="38" rows="8"><%=rsCover("CoverDesc")%></textarea></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>包含的产品有：(多选)</td>
      <td><!--<Select name="Iclude_ProductIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0,0,"Product_ClassChild")		大型控件%>
			
		  </select>-->
		  
		  <input name="Iclude_ProductIDs" type="text" value="<%=rsCover("Iclude_ProductIDs")%>" size=6 maxlength=30 />
	  </td>
    </tr>
	
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>所属大类型：(多选)</td>
      <td><!--<Select name="ClassIDs" size="6" multiple>
				  
	  		<%'Call Admin_ShowClass_Option_classid(0, rsCover("ClassIDs"), "Product_Class")		大型控件%>
			
		  </select>-->
	  		<input name="ClassIDs" type="text" value="<%=rsCover("ClassIDs")%>" size=6 maxlength=30 />
	  </td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>本礼包最大可选商品数量：</td>
      <td><input name="MaxNumber" type="text" value="<%=rsCover("MaxNumber")%>" size=3 maxlength=10 /></td>
    </tr>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>总体优惠价格：</td>
      <td><input name="CoverPrice" type="text" value="<%=FormatCurrency(rsCover("CoverPrice"))%>" size=10 maxlength=20 style="text-align:center;" />元RMB</td>
    </tr>
	
    <TR class="tdbg" > 
      <TD>礼包的展示图片：</TD>
      <TD><input name="CoverPhoto" type="text" value="<%=rsCover("CoverPhoto")%>" size=50 maxlength=500 /> 
	  
	  </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">本礼包e促销状态：</TD>
      <TD width="60%"><INPUT name="Discontinued" type=radio value="0" <% If rsCover("Discontinued")=0 Then Response.Write "checked" %> />促销中&nbsp;&nbsp;&nbsp;<INPUT name="Discontinued" type=radio value="1" <% If rsCover("Discontinued")=1 Then Response.Write "checked" %> />停止促销
      </TD>
    </TR>
   
  
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="保存礼包修改结果"> <input name="CoverID" type="hidden" id="CoverID" value="<%=rsCover("CoverID")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsCover.Close
	Set rsCover=Nothing
End Sub



Sub gouser1()
If is_ot_user=1 Then
	Response.Write("整合外部数据库礼包，不能使用此功能"):Response.End()
End If
%>
<FORM name="Form1" action="Administrator_CoverProducts.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">登录到栏目管理后台</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          本操作供管理员登录到栏目的管理界面进行管理。<br>
          当用栏目操作出现障碍时，可进入该礼包后台。<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">输入栏目名称：</td>
    <td height="25"><input name="username" type="text" id="username" value="" size="30" maxlength="50"></td>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value=" 提交 "></td>
  </tr>
</table>
</form>
  <%
End Sub
%>
</body>
</html>
<%



Sub SaveAdd()
	Dim rsCover,sqlCover
	Dim CoverName,CoverDesc,Iclude_ProductIDs,ClassIDs,MaxNumber,CoverPrice,CoverPhoto,Discontinued
	
	CoverName			=Trim(Request("CoverName"))
	CoverDesc			=Trim(Request("CoverDesc"))
	Iclude_ProductIDs	=Trim(Request("Iclude_ProductIDs"))
	ClassIDs			=Trim(Request("ClassIDs"))
	MaxNumber			=Trim(Request("MaxNumber"))
	CoverPrice			=Trim(Request("CoverPrice"))
	CoverPhoto			=Trim(Request("CoverPhoto"))
	Discontinued		=Trim(Request("Discontinued"))
	

	If CoverName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包名称不能为空</li>"
	Else
		CoverName = Trim(CoverName)
	End If

	If CoverDesc="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包描述不能为空</li>"
	Else
		CoverDesc = Trim(CoverDesc)
	End If
	
	If Not Instr(Iclude_ProductIDs,",")>0 Then
		Errmsg=Errmsg & "<br><li>请选择2个或者2个以上的产品，作为礼包！</li>"
		founderr=True
	Else
		Iclude_ProductIDs = Trim(Iclude_ProductIDs)
	End If

	If Not Len(ClassIDs)>0 Then
		Errmsg=Errmsg & "<br><li>请选择至少1个的所属大分类！</li>"
		founderr=True
	Else
		ClassIDs = Trim(ClassIDs)
	End If
	
	If MaxNumber="" Or Not isNumeric(MaxNumber) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包 可选择的最大产品数量 不能为空，并且要为数字！</li>"
	Else
		'UnitPrice = FormatCurrency(UnitPrice)
		MaxNumber = Cint(MaxNumber)
	End If
	
	If CoverPrice="" Or Not isNumeric(CoverPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包 总体优惠价格 不能为空，并且要为货币数字！</li>"
	Else
		CoverPrice = FormatCurrency(CoverPrice)
	End If
	
	
	If CoverPhoto="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包图片上传不能为空</li>"
	Else
		CoverPhoto = Trim(CoverPhoto)
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>请选择 是否已经放弃此礼包的 促销状态！</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rs=Nothing
		Exit Sub
	End If
	
	
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts"
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3

	rsCover.Addnew
	
		If Not CoverName="" Then  			rsCover("CoverName")			=CoverName
		If Not CoverDesc="" Then  			rsCover("CoverDesc")			=CoverDesc
		If Not Iclude_ProductIDs="" Then  	rsCover("Iclude_ProductIDs")	=Iclude_ProductIDs
		If Not ClassIDs="" Then  			rsCover("ClassIDs")				=ClassIDs
		If Not MaxNumber="" Then  			rsCover("MaxNumber")			=MaxNumber
		If Not CoverPrice="" Then  			rsCover("CoverPrice")			=CoverPrice
		If Not CoverPhoto="" Then  			rsCover("CoverPhoto")			=CoverPhoto
		If Not Discontinued="" Then  		rsCover("Discontinued")			=Discontinued

		
	rsCover.Update
	
	rsCover.Close
	Set rsCover=Nothing
	
	oblog.showok "注册新礼包成功!",""
End Sub



Sub gouser2()
	If is_ot_user=1 Then
		Response.Write("整合外部数据库礼包，不能使用此功能"):Response.End()
	End If
	Dim rs,username
	username=oblog.filt_badstr(Trim(Request("username")))
	If username="" Then Response.Write("礼包名不能为空"):Response.End()
	If is_ot_user=0 Then
		Set rs=oblog.Execute("Select username,password From Product_CoverProducts Where username='"&username&"'")
	Else
		Set rs=ot_conn.Execute("Select "&ot_username&","&ot_password&" From "&ot_usertable&" Where "&ot_username&"='"&username&"'")
	End If
	If Not rs.Eof Then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		Set rs=Nothing
		Response.Redirect("../HomestayBackctrl-indexHome.asp")
	Else
		Set rs=Nothing
		Response.Write("无此礼包"):Response.End()
	End If
	
End Sub


Sub SaveModify()
	Dim rsCover,sqlCover
	
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		Exit Sub
	Else
		CoverID=Clng(CoverID)
	End If
	
	Dim CoverName,CoverDesc,Iclude_ProductIDs,ClassIDs,MaxNumber,CoverPrice,CoverPhoto,Discontinued
	
	CoverName			=Trim(Request("CoverName"))
	CoverDesc			=Trim(Request("CoverDesc"))
	Iclude_ProductIDs	=Trim(Request("Iclude_ProductIDs"))
	ClassIDs			=Trim(Request("ClassIDs"))
	MaxNumber			=Trim(Request("MaxNumber"))
	CoverPrice			=Trim(Request("CoverPrice"))
	CoverPhoto			=Trim(Request("CoverPhoto"))
	Discontinued		=Trim(Request("Discontinued"))
	

	If CoverName="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包名称不能为空</li>"
	Else
		CoverName = Trim(CoverName)
	End If

	If CoverDesc="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包描述不能为空</li>"
	Else
		CoverDesc = Trim(CoverDesc)
	End If
	
	If Not Instr(Iclude_ProductIDs,",")>0 Then
		Errmsg=Errmsg & "<br><li>请选择2个或者2个以上的产品，作为礼包！</li>"
		founderr=True
	Else
		Iclude_ProductIDs = Trim(Iclude_ProductIDs)
	End If

	If Not Len(ClassIDs)>0 Then
		Errmsg=Errmsg & "<br><li>请选择至少1个的所属大分类！</li>"
		founderr=True
	Else
		ClassIDs = Trim(ClassIDs)
	End If
	
	If MaxNumber="" Or Not isNumeric(MaxNumber) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包 可选择的最大产品数量 不能为空，并且要为数字！</li>"
	Else
		'UnitPrice = FormatCurrency(UnitPrice)
		MaxNumber = Cint(MaxNumber)
	End If
	
	If CoverPrice="" Or Not isNumeric(CoverPrice) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包 总体优惠价格 不能为空，并且要为货币数字！</li>"
	Else
		CoverPrice = FormatCurrency(CoverPrice)
	End If
	
	
	If CoverPhoto="" Then
		founderr=True
		Errmsg=Errmsg & "<br><li>礼包图片上传不能为空</li>"
	Else
		CoverPhoto = Trim(CoverPhoto)
	End If
	
	If Not isNumeric(Discontinued) Then
		founderr=True
		Errmsg=Errmsg & "<br><li>请选择 是否已经放弃此礼包的 促销状态！</li>"
	Else
		Discontinued = Cint(Discontinued)
	End If
	

	If founderr=True Then
		Set rsuser=Nothing
		Exit Sub
	End If
	
	Set rsCover=Server.CreateObject("Adodb.RecordSet")
	sqlCover="Select * From Product_CoverProducts Where CoverID=" & CoverID
	If Not IsObject(conn) Then link_database
	rsCover.Open sqlCover,Conn,1,3
	If rsCover.Bof And rsCover.Eof Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的礼包！</li>"
		rsCover.Close
		Set rsCover=Nothing
		Exit Sub
	End If
	
	
	If Not CoverName="" Then  			rsCover("CoverName")			=CoverName
	If Not CoverDesc="" Then  			rsCover("CoverDesc")			=CoverDesc
	If Not Iclude_ProductIDs="" Then  	rsCover("Iclude_ProductIDs")	=Iclude_ProductIDs
	If Not ClassIDs="" Then  			rsCover("ClassIDs")				=ClassIDs
	If Not MaxNumber="" Then  			rsCover("MaxNumber")			=MaxNumber
	If Not CoverPrice="" Then  			rsCover("CoverPrice")			=CoverPrice
	If Not CoverPhoto="" Then  			rsCover("CoverPhoto")			=CoverPhoto
	If Not Discontinued="" Then  		rsCover("Discontinued")			=Discontinued
		
	
	rsCover.Update
	rsCover.Close
	Set rsCover=Nothing
	oblog.showok "修改成功!",""
End Sub


Sub DelUser()
	Dim rs,i
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的礼包</li>"
		Exit Sub
	End If
	If Instr(CoverID,",")>0 Then
		CoverID=Split(CoverID,",")
		for i=0 to Ubound(CoverID)
			deloneuser(CoverID(i))
		next
	Else
		deloneuser(CoverID)
	End If
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub

Sub deloneuser(CoverID)
	CoverID=Clng(CoverID)
'	Dim rs,fso,f,uname,udir
'	Set rs=oblog.Execute("Select CoverID From Product_CoverProducts Where CoverID="&CoverID)
'	If Not rs.Eof Then
'		udir=rs(0)
'		uname=rs(1)
'		Set fso=Server.CreateObject("scripting.filesystemobject")
'		If fso.FolderExists(Server.MapPath(blogdir & udir&"/"&rs("user_folder"))) Then 
'			Set f = fso.GetFolder(Server.MapPath(blogdir & udir&"/"&rs("user_folder")))
'			f.Delete True
'		End If
'		Set f=Nothing
'		Set fso=Nothing
'		oblog.Execute("Delete From oblog_log Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_comment Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_message Where CoverID="&CoverID)
'		oblog.Execute("Delete From oblog_subject Where CoverID="&CoverID)
'		oblog.Execute("Delete From Product_CoverProducts Where CoverID=" & CoverID)
'		oblog.Execute("Delete From oblog_upfile Where CoverID=" & CoverID)
		oblog.Execute("Delete From Product_CoverProducts Where CoverID=" & CoverID)
'		oblog.Execute("Update Product_CoverProducts Set Discontinued=1 Where CoverID=" & CoverID)
'	End If
'	Set rs=Nothing
End Sub

Sub LockUser()
	Dim rs,udir
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的礼包</li>"
		Exit Sub
	End If
	CoverID=Clng(CoverID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_CoverProducts Where CoverID="&CoverID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=1
		udir=rs(1)
		rs.Update
		strtmp="<script language=javascript>"&vbcrlf
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=此礼包已经被锁定，请联系管理员!');"&vbcrlf
		strtmp=strtmp&"</script>"&vbcrlf
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&CoverID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&CoverID&"/inc/chkblogpassword.htm",strtmp)								
			End If
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "锁定礼包成功",""
End Sub

Sub UnLockUser()
	Dim rs,udir
	If CoverID="" Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的礼包</li>"
		Exit Sub
	End If
	CoverID=Clng(CoverID)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select lockuser,user_dir From Product_CoverProducts Where CoverID="&CoverID,conn,1,3
	If Not rs.Eof Then
		Dim strtmp,fso
		rs(0)=0
		udir=rs(1)
		rs.Update
		strtmp=" "
		If udir<>"" Then
		Set fso=Server.CreateObject("scripting.filesystemobject")
			udir=Server.MapPath(udir)
			If fso.FolderExists(udir&"/"&CoverID&"/inc") Then
				Call oblog.BuildFile(udir&"/"&CoverID&"/inc/chkblogpassword.htm",strtmp)				
			End If			
			Set fso=Nothing
		End If
	End If
	rs.Close
	Set rs=Nothing
	oblog.showok "解锁礼包成功",""
End Sub



Sub Discontinued()
	CoverID=Clng(CoverID)
	oblog.Execute("Update Product_CoverProducts Set Discontinued=1 Where CoverID="& CoverID)
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub


Sub UnDiscontinued()
	CoverID=Clng(CoverID)
	oblog.Execute("Update Product_CoverProducts Set Discontinued=0 Where CoverID="& CoverID)
	Response.Redirect "Administrator_CoverProducts.asp"
End Sub



'使用例子：Call Admin_ShowClass_Option_classid(0,ParentID,"Product_Class")		'ParentID=classid.
Sub Admin_ShowClass_Option_classid(ShowType,CurrentID,Table_Class)
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
	sqlClass="Select * From "& Table_Class &" Where idType=" & t &" Order By RootID,OrderID"
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


Sub WriteErrMsg()
	Dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	Response.Write strErr
End Sub





%>
