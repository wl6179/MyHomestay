<!--#include file="inc/inc_sys.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim rs, sql
dim id,UserSearch,Keyword,strField
dim Action,FoundErr,ErrMsg
dim tmpDays
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
UserSearch=trim(request("UserSearch"))
Action=trim(request("Action"))
id=trim(Request("id"))

if UserSearch="" then
	UserSearch=0
else
	UserSearch=Clng(UserSearch)
end if
strFileName="admin_blogstar.asp?UserSearch=" & UserSearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>用户明星管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
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
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>博 客 之 星 管 理</strong></td>
  </tr>
  <form name="form1" action="admin_blogstar.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>快速查找：</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="0">最后500个用户明星</option>
          <option value="1">通过审核的用户明星</option>
          <option value="2">未通过审核的用户明星</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_blogstar.asp">用户明星管理首页</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_blogstar.asp">
  <tr class="tdbg">
      <td width="120"><strong>高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="blogname" selected>用户明星名</option>
      <option value="UserID" >用户明星ID</option>
  
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
        若为空，则查询所有</td>
  </tr>
</form>
</table>
<%
if Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelUser()
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='admin_blogstar.asp'>用户明星管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_blogstar order by id desc"
			strGuide=strGuide & "最后500个用户明星"
		case 1
			sql="select * from oblog_blogstar where ispass=1 order by id desc"
			strGuide=strGuide & "通过审核的用户明星"
		case 2
			sql="select * from oblog_blogstar where ispass=0 order by id desc"
			strGuide=strGuide & "未通过审核的用户明星"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_blogstar order by id desc"
				strGuide=strGuide & "所有用户明星"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID必须是整数！</li>"
					else
						sql="select * from oblog_blogstar where id =" & Clng(Keyword)
						strGuide=strGuide & "用户明星ID等于<font color=red> " & Clng(Keyword) & " </font>的用户明星"
					end if
				case "blogname"
					sql="select * from oblog_blogstar where blogname like '%" & Keyword & "%' order by id  desc"
					strGuide=strGuide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的用户明星"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>错误的参数！</li>"
	end select
	strGuide=strGuide & "</td><td align='right'>"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "共找到 <font color=red>0</font> 个用户明星</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个用户明星</td></tr></table>"
		response.write strGuide
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
        	showContent
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户明星")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户明星")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户明星")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="admin_blogstar.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> 用户名</font></td>
            <td height="22" align="center"><font color="#FFFFFF">状态</font></td>
            <td align="center"><font color="#FFFFFF">申请时间</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF">简介</font></td>
            <td width="50" height="22" align="center"><font color="#FFFFFF"> 图片</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF"> 
              操作</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><%=rs("id")%></td>
            <td width="80" align="center"><%="<a href='"& rs("userurl")&"' target='_blank'>"&rs("blogname")&"</a>"%> </td>
            <td align="center"> <%
			select case rs("ispass")
				case 0
					response.write "<font color=red>等待审核</font>"
				case 1
					response.write "通过审核"
			end select
			%> </td>
            <td align="center"> <%
	if rs("addtime")<>"" then
		response.write rs("addtime")
	else
		response.write "&nbsp;"
	end if
	%> </td>
            <td width="250" align="center"> <%=oblog.filt_html(rs("info"))%> </td>
            <td width="50" align="center"><%="<a href='"&rs("picurl")&"' target='_blank'>点击察看</a>"%></td>
            <td width="120" align="center"><%
		response.write "<a href='admin_blogstar.asp?Action=Modify&id=" & rs("id") & "'>修改</a>&nbsp;"
        response.write "<a href='admin_blogstar.asp?Action=Del&id=" & rs("id") & "' onClick='return confirm(""确定要删除此用户明星吗？"");'>删除</a>&nbsp;"
		%> </td>
          </tr>
          <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
        </table>  
</td>
</form></tr></table>
<%
end sub


sub Modify()
	dim rsUser,sqlUser
	id=clng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户明星！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_blogstar.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">修改用户明星信息</font></b></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>真实姓名：</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="40%">连接地址：</TD>
      <TD width="60%"> <INPUT name="userurl" value="<%=rsUser("userurl")%>" size=50   maxLength=250> <a href="<%=rsuser("userurl")%>" target="_blank">查看</a>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"> 图片连接(<strong><font color="#FF0000">请将大图片手工改为合适的尺寸</font></strong>)：</TD>
      <TD width="60%"> <INPUT name=picurl value="<%=rsUser("picurl")%>" size=50 maxLength=250><a href="<%=rsuser("picurl")%>" target="_blank">查看</a></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">简介：</TD>
      <TD width="60%"><textarea name="bloginfo" cols="40" rows="5"><%=oblog.filt_html(rsuser("info"))%></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">状态：</TD>
      <TD width="60%"><input type="radio" name="ispass" value=0 <%if rsUser("ispass")=0 then response.write "checked"%>>
        未通过审核&nbsp;&nbsp; <input type="radio" name="ispass" value=1 <%if rsUser("ispass")=1 then response.write "checked"%>>
        已通过审核</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="id" type="hidden" id="id" value="<%=rsUser("id")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsUser.close
	set rsUser=nothing
end sub


sub SaveModify()
	dim rsuser,sqlUser
	id=clng(id)
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_blogstar where id=" & id
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	rsUser("blogname")=trim(request("blogname"))
	rsUser("userurl")=trim(request("userurl"))
	rsUser("picurl")=trim(request("picurl"))
	rsUser("ispass")=trim(request("ispass"))
	rsUser("info")=trim(request("bloginfo"))
	rsUser("addtime")=now()
	rsUser.update
	rsUser.Close
	set rsUser=nothing
	oblog.showok "修改成功!",""
end sub


sub DelUser()
	id=clng(id)
	oblog.execute("delete from oblog_blogstar where id="&id)
	oblog.showok "删除成功！",""
end sub


sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

%>