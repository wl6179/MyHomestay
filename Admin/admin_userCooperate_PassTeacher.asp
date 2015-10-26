<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
Dim tName,t
tName = "外教资料"

Const MaxPerPage = 20
Dim strFileName
Dim totalPut, TotalPages
Dim rs, sql
Dim id, usersearch, Keyword, strField, uid, action
Dim selectsub, substr, wsql, allsub

t = Trim(request("t"))
Keyword = Trim(request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
strField = Trim(request("Field"))
usersearch = Trim(request("usersearch"))
selectsub = Trim(request("selectsub"))
action = Trim(request("action"))
id = oblog.filt_badstr(Trim(request("id")))
uid = oblog.filt_badstr(Trim(request("uid")))
strFileName = "admin_userCooperate_PassTeacher.asp?t=" & t
If usersearch = "" Then
    usersearch = 0
Else
    usersearch = CLng(usersearch)
    strFileName = "admin_userCooperate_PassTeacher.asp?usersearch=" & usersearch & "&t=" & t
End If
If selectsub <> "" Then
    selectsub = CLng(selectsub)
    strFileName = "admin_userCooperate_PassTeacher.asp?usersearch=10&selectsub=" & selectsub & "&t=" & t
Else
    selectsub = 0
End If
If Keyword <> "" Then
    strFileName="admin_userCooperate_PassTeacher.asp?usersearch=10&keyword="&Keyword&"&Field="&strField & "&t="  & t
End If
If request("page") <> "" Then CurrentPage = CInt(request("page")) Else CurrentPage = 1
if oblog.logined_ulevel<>9 then wsql=" and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" )"




dim FoundErr,ErrMsg
dim tmpDays
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>通过审核的外教资料管理</title>
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
    <td height="22" colspan=2 align=center><strong>已通过我方审核的 外教资料管理</strong></td>
  </tr>
  <form name="form1" action="admin_userCooperate_PassTeacher.asp" method="get">
    <tr class="tdbg"> 
      <td width="120" height="30"><strong>快速查找外教资料：</strong></td>
      <td width="507" height="30"><input type="hidden" name="t" value="<%=t%>">
	  
	  <select size=1 name="usersearch" onChange="javascript:submit()">
          <option value="10">请选择类型</option>
          <option value="0">列出所有<%=tName%></option>
          <option value="1" <% If usersearch=1 Then Response.Write("selected")%>>列出没有关联发布地址URL的<%=tName%></option>
		  <option value="2" <% If usersearch=2 Then Response.Write("selected")%>>列出已经填写关联发布地址URL的<%=tName%></option>
          <option value="3" <% If usersearch=3 Then Response.Write("selected")%>>在草稿箱里的<%=tName%></option>
			<option value="4" <% If usersearch=4 Then Response.Write("selected")%>>强制查看已被删除的<%=tName%></option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_userCooperate_PassTeacher.asp">通过审核的外教资料管理首页</a></td>
    </tr>
  </form>
  
  
  <form name="form2" method="post" action="admin_userCooperate_PassTeacher.asp">
  <tr class="tdbg">
    <td width="120"><strong>外教资料高级查询：</strong></td>
    <td >
      <select name="Field" id="Field">

            <option value="logid" <% If strField="logid" Then Response.Write("selected")%>>按<%=tName%>ID号查询</option>
            <option value="topic" <% If strField<>"logid" Then Response.Write("selected")%>>按标题注释查询</option>
			<option value="author" <% If strField="author" Then Response.Write("selected")%>>指定某加盟帐号的<%=tName%></option>
			<option value="logtext" <% If strField="logtext" Then Response.Write("selected")%>>按提交内容查询</option>
			
         </select>	<input name="usersearch" type="hidden" value="10">
		 
         <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
		 
         <input type="submit" name="Submit2" value=" 搜索 "> 
	  若为空，则查询所有通过审核的外教资料</td>
  </tr>
</form>
</table>
<%
if Action="Add" then
	call AddUser()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
	
elseif Action="del" then
    	Call delblog
elseif Action="state_shenhe" then
    	Call change_state_shenhe
		
elseif Action="Lock" then
	call LockUser()
elseif Action="UnLock" then
	call UnLockUser()
elseif Action="Move" then
	call MoveUser()
elseif Action="Update" then
	call UpdateUser()
elseif Action="DoUpdate" then
	call DoUpdate()
elseif Action="DoUpdatelog" then
	call DoUpdatelog()
elseif Action="gouser1" then
	call gouser1()
elseif Action="gouser2" then
	call gouser2()
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='admin_userCooperate_PassTeacher.asp'>通过审核的外教资料管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 order by addtime desc"
			strGuide=strGuide & "最后500个通过审核的外教资料"
		case 1'列出没有发布URL的外教资料；
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 And logPublic_path='' order by addtime desc"
			strGuide=strGuide & "没有填写发布地址URL的前500个外教资料"
		case 2
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 And logPublic_path<>'' order by addtime desc"
			strGuide=strGuide & "已经填写了发布地址URL的前500个外教资料"
		case 3
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft=1 order by addtime desc"
			strGuide=strGuide & "尚保存在加盟用户草稿箱里的前500个外教资料"
		case 4
			sql="select * from  [oblog_logCooperateSubmit] where isDelete=1 And state_shenhe=1 And isdraft<>1 order by addtime desc"
			strGuide=strGuide & "强制查看已被删除的前500个外教资料"
		case 5
			sql="select * from  [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and user_isbest=1 order by addtime desc"
			strGuide=strGuide & "推荐用户"
		case 6
			sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and User_Level=6 order by userID  desc"
			strGuide=strGuide & "等待管理认证证的用户"
		case 7
			sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and  LockUser =1 order by userID  desc"
			strGuide=strGuide & "所有被锁住的用户"
		case 10
		
			if Keyword="" then
				sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 order by addtime desc"
				strGuide=strGuide & "所有外教资料"
			else
				select case strField
				case "logid"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>外教资料ID必须是整数！</li>"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logid =" & Clng(Keyword)
						strGuide=strGuide & "外教资料ID等于<font color=red> " & Clng(Keyword) & " </font>的外教资料"
					end if
				case "topic"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and topic like '%" & Keyword & "%' order by addtime  desc"
						strGuide=strGuide & "外教资料的标题注释中含有“ <font color=red>" & Keyword & "</font> ”的外教资料"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and topic= '" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "外教资料的标题注释等于“ <font color=red>" & Keyword & "</font> ”的外教资料"
					end if
					
				case "author"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and author like '%" & Keyword & "%' order by addtime  desc"
						strGuide=strGuide & "加盟帐户名含有：“ <font color=red>" & Keyword & "</font> ”的帐号 所填写的所有外教资料！"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and author='" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "加盟帐户等于“ <font color=red>" & Keyword & "</font> ”的帐号 所填写的所有外教资料"
					end if
				case "logtext"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logtext like '%" & Keyword & "%' order by addtime desc"
						strGuide=strGuide & "提交内容中含有“ <font color=red>" & Keyword & "</font> ”的外教资料"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logtext='" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "提交内容等于“ <font color=red>" & Keyword & "</font> ”的外教资料"
					end if
					
					
					
					
				case "blogname"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and blogname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "真实姓名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and blogname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "真实姓名等于“ <font color=red>" & Keyword & "</font> ”的用户"
					end if
				case "logcount"
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and log_count < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "文章数小于“ <font color=red>" & Keyword & "</font> ”的用户"
				case "logintimes"
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and logintimes < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "登录次数小于“ <font color=red>" & Keyword & "</font> ”的用户"
				case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and datediff(d,lastlogintime,getdate())>"&clng(Keyword)&" order by userID  desc"
				   Else 
				   	sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and datediff('d',lastlogintime,now())>"&clng(Keyword)&" order by userID  desc"
				   End If 
					strGuide=strGuide & "“ <font color=red>" & Keyword & "</font> ”天内未登录的用户"
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
'	Response.Write sql
'	Response.End 
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "共找到 <font color=red>0</font> 个外教资料</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个外教资料</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个外教资料")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个外教资料")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个外教资料")
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
  <form name="myform" method="Post" action="admin_userCooperate_PassTeacher.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">选中</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="180" height="22" align="center"><font color="#FFFFFF"> 标题注释</font></td>
            <td height="22" align="center"><font color="#FFFFFF">提交者(加盟商)</font></td>
            
			
            <td align="center"><font color="#FFFFFF">提交时间</font></td>
            
            <td width="80" height="22" align="center"><font color="#FFFFFF"> 审核操作</font></td>
            <td width="150" height="22" align="center"><font color="#FFFFFF"> 
              操作</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("logid"))%>'></td>
            <td width="30" align="center"><%=rs("logid")%></td>
            <td width="180" align="center"><%
			response.write "" & rs("topic") & ""
			If rs("isdraft")=1 Then response.write "<font color=red title='该外教资料已被加盟商设为了草稿'>*草稿</font>"
			If rs("logPublic_path")<>"" And Len(rs("logPublic_path"))>6 Then
				response.write "<br />URL：<a href='"& rs("logPublic_path") &"'><font color=#038ad7>"& rs("logPublic_path") &"</font></a>&nbsp;"
			Else
				
			End If
			%> </td>
            <td align="center"> <%
			Response.Write rs("author")

			%> </td>
            
			
            <td align="center"> <%
	if rs("addtime")<>"" then
		response.write rs("addtime")
	else
		response.write "&nbsp;"
	end if
	%> </td>
            
            <td width="80" align="center"><input type="button" value=" 取 消 " onClick="if (confirm('确定要取消该外教资料的审核状态吗？') == false)return false;window.location.href='admin_userCooperate_PassTeacher.asp?action=state_shenhe&state_shenhe=<%=rs("state_shenhe")%>&logPublic_path=<%=rs("logPublic_path")%>&logid=<%=rs("logid")%>&userid=<%=rs("userid")%>'"></td>
			
            <td width="150" align="center"><%
		response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=gouser2&userid=" & rs("userid") & "&logid="& rs("logid") &"&actionTemp=preview' target='blank'>预览资料</a>&nbsp;"
		
		If rs("logPublic_path")<>"" And Len(rs("logPublic_path"))>6 Then
			response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=Modify&logid=" & rs("logid") & "' title='"& rs("logPublic_path") &"'><font color=#038ad7>已关联URL√</font></a>&nbsp;"
		Else
			response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=Modify&logid=" & rs("logid") & "' title='请在审核之前，先关联好此外教的站点发布地址URL！'><font color=red>请关联发布地址</font></a>&nbsp;"
		End If
		
		response.write "<br /><a href='admin_userCooperate_PassTeacher.asp?Action=gouser2&userid=" & rs("userid") & "&logid="& rs("logid") &"' title='您可以强制修改加盟用户提交的外教资料内容' target='blank'>强制修改</a>&nbsp;"
		
        response.write "<a href='admin_userCooperate_PassTeacher.asp?action=del&id=" & rs("logid") & "&t=" & t & "' onClick='return confirm(""确定要删除此外教资料吗？"");'>删除</a>&nbsp;"
		%> </td>
          </tr>
          <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
        </table>  
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的外教资料</td>
            <td> <strong>操作：</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=true">
              删除&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">移动到
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">中国家庭or合作伙伴用户</option>
                <option value="7">注册栏目发布</option>
                <option value="8">VIP栏目发布</option>
                <option value="9">前台管理发布</option>
              </select>
              &nbsp;&nbsp; -->
              <input type="submit" name="Submit" value=" 执 行 "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub


sub Modify()
	dim logid
	dim rslog,sqllog
	logid=trim(request("logid"))
	if logid="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		logid=Clng(logid)
	end if
	Set rslog=Server.CreateObject("Adodb.RecordSet")
	sqllog="select * from [oblog_logCooperateSubmit] where logid=" & logid
	if not IsObject(conn) then link_database
	rslog.Open sqllog,Conn,1,3
	if rslog.bof and rslog.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的外教资料！</li>"
		rslog.close
		set rslog=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">该外教资料的 站点发布地址URL</font></b></TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">标题注释：</TD>
      <TD width="60%"><%=rslog("topic")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">提交者(加盟商)：</TD>
      <TD width="60%"><%=rslog("author")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">发布地址URL：</TD>
      <TD width="60%"><input name="logPublic_path" type="text" id="logPublic_path" value="<%=rslog("logPublic_path")%>" size="50">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
   
   
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
	  <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'>
	  <input name="logid" type="hidden" id="logid" value="<%=rslog("logid")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rslog.close
	set rslog=nothing
end sub


sub UpdateUser()
%>
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=DoUpdate" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">更新用户静态页面</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成用户静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3 、本操作根据用户ｉｄ更新。 </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">开始用户ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      用户ID，可以填写您想从哪一个ID号开始进行更新</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">结束用户ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="生成静态页面"></td>
  </tr>
</table>
</form>

<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=DoUpdatelog" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">更新文章静态页面</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成用户静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3、本操作根据文章ｉｄ更新。</p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">开始文章ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      用户ID，可以填写您想从哪一个ID号开始进行更新</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">结束文章ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      将更新开始到结束ID之间的文章页面，之间的数值最好不要选择过大</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="生成文章静态页面"></td>
  </tr>
</table>
</form>
  <%
end sub

sub gouser1()
if is_ot_user=1 then
	response.Write("整合外部数据库用户，不能使用此功能"):response.End()
end if
%>
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">登录到用户管理后台</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          本操作供管理员登录到用户的管理界面进行管理。<br>
          当用户操作出现障碍时，可进入该用户后台，协助用户进行操作。<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">用户账号：</td>
    <td height="25"><input name="username" type="text" id="username" value="" size="30" maxlength="50"></td>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value=" 提交 "></td>
  </tr>
</table>
</form>
  <%
end sub
%>
</body>
</html>
<%

sub gouser2()
	if is_ot_user=1 then
		response.Write("整合外部数据库用户，不能使用此功能"):response.End()
	end if
	dim rs,username,userid
	userid=oblog.filt_badstr(trim(request("userid")))
	if userid="" then response.Write("参数为空"):response.End()
	userid=Cint(request("userid"))
	if is_ot_user=0 then
		set rs=oblog.execute("select username,password from [oblog_user] where userid='"&userid&"'")
	else
		'''set rs=ot_conn.execute("select "&ot_username&","&ot_password&" from "&ot_usertable&" where "&ot_username&"='"&username&"'")
	end if
	if not rs.eof then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		set rs=nothing
		
		If Request("actionTemp")="preview" Then'加入是预览操作；
		response.Redirect("../HomestayBackctrl-logPreviewCooperate.asp?logid="& Request("logid") &"&t=0")
		Else
		response.Redirect("../HomestayBackctrl-PostSubmitCooperate.asp?logid="& Request("logid") &"&t=0")
		End If
	else
		set rs=nothing
		response.Write("无此用户"):response.End()
	end if
	
end sub


sub SaveModify()
	dim UserID,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,User_Level,LockUser,isbest
	dim rsUser,sqlUser
	dim blogname,usertype,user_upfiles_max,upfiles_size,user_domain,user_domainroot
	
	Dim logid,logPublic_path
	
	logPublic_path=trim(request.Form("logPublic_path"))
	logid=trim(request("logid"))
	
	if logid="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		logid=Clng(logid)
	end if
	
	
'	if User_Level="" then
'		FoundErr=true
'		ErrMsg=ErrMsg & "<br><li>请指定用户级别！</li>"
'	else
'		User_Level=CLng(User_Level)
'	end if
	
	

	if founderr=true then
		''set rsuser=nothing
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [oblog_logCooperateSubmit] where logid=" & logid
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的外教资料！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	

	rsUser("logPublic_path")=logPublic_path


	rsUser.update
	rsUser.Close
	set rsUser=nothing
	oblog.showok "修改成功!",""
end sub


sub DelUser()
	dim rs,i
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定要删除的外教资料</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=split(UserID,",")
		for i=0 to ubound(userid)
			deloneuser(userid(i))
		next
	else
		deloneuser(userid)
	end if
	response.redirect "admin_userCooperate_PassTeacher.asp"
end sub

sub deloneuser(logid)
	'''Dim logid
	logid=Clng(logid)
	


		oblog.execute("delete from oblog_logCooperateSubmit where logid="&logid)

end sub

sub LockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的用户</li>"
		exit sub
	end if
	userid=clng(userid)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select lockuser,user_dir from [oblog_logCooperateSubmit] where userid="&userid,conn,1,3
	if not rs.eof then
		dim strtmp,fso
		rs(0)=1
		udir=rs(1)
		rs.update
		strtmp="<script language=javascript>"&vbcrlf
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=此用户已经被锁定，请联系管理员!');"&vbcrlf
		strtmp=strtmp&"</script>"&vbcrlf
		if udir<>"" then
		set fso=server.createobject("scripting.filesystemobject")
			udir=server.MapPath(udir)
			if fso.FolderExists(udir&"/"&userid&"/inc") then
				Call oblog.BuildFile(udir&"/"&userid&"/inc/chkblogpassword.htm",strtmp)								
			end if
			set fso=nothing
		end if
	end if
	rs.close
	set rs=nothing
	oblog.showok "锁定用户成功",""
end sub

sub UnLockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的用户</li>"
		exit sub
	end if
	userid=clng(userid)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select lockuser,user_dir from [oblog_logCooperateSubmit] where userid="&userid,conn,1,3
	if not rs.eof then
		dim strtmp,fso
		rs(0)=0
		udir=rs(1)
		rs.update
		strtmp=" "
		if udir<>"" then
		set fso=server.createobject("scripting.filesystemobject")
			udir=server.MapPath(udir)
			if fso.FolderExists(udir&"/"&userid&"/inc") then
				Call oblog.BuildFile(udir&"/"&userid&"/inc/chkblogpassword.htm",strtmp)				
			end if			
			set fso=nothing
		end if
	end if
	rs.close
	set rs=nothing
	oblog.showok "解锁用户成功",""
end sub

sub MoveUser()
	dim msg
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定要移动的用户</li>"
		exit sub
	end if
	dim User_Level
	User_Level=trim(request("User_Level"))
	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定目标用户组</li>"
		exit sub
	else
		User_Level=Clng(User_Level)
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>中国家庭or合作伙伴用户</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=6 where userID in (" & UserID & ")"
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>注册栏目发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=7 where userID in (" & UserID & ")"
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>VIP栏目发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=8 where userID in (" & UserID & ")"
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>前台管理发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=9 where userID in (" & UserID & ")"
		end if
	else
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>中国家庭or合作伙伴用户</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=6 where userID=" & UserID 
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>注册栏目发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=7 where userID="& UserID
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>VIP栏目发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=8 where userID=" & UserID 
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定用户设为“<font color=blue>前台管理发布</font>”！"
			sql="Update [oblog_logCooperateSubmit] set User_Level=9 where userID=" & UserID
		end if
	end if
	Conn.Execute sql
	response.Redirect "admin_userCooperate_PassTeacher.asp"
	'call WriteSuccessMsg(msg)
end sub

sub DoUpdate()
	Server.ScriptTimeOut=999999999
	dim BeginID,EndID,p1,rsUser,blog,i
	BeginID=trim(request("BeginID"))
	EndID=trim(request("EndID"))
	if BeginID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定开始ID</li>"
	else
		BeginID=Clng(BeginID)
	end if
	if EndID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定结束ID</li>"
	else
		EndID=Clng(EndID)
	end if
	if FoundErr=true then exit sub
	set rsuser=oblog.execute("select count(userid) from [oblog_logCooperateSubmit] where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID))
	p1=rsuser(0)
	set rsuser=oblog.execute("select userid from [oblog_logCooperateSubmit] where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID)&" order by userid")
	set blog=new class_blog
	response.Write("<div style=""text-align: center;"">")
	response.Write("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
	i=1
	blog.progress_init
	do while not rsUser.eof
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""全部进度：当前用户ID:"&rsuser(0)&""";</script>" & VbCrLf
		Response.Flush
		'blog.userid=rsUser(0)
		blog.update_alllog_admin rsUser(0)
		If P_BLOG_UPDATEPAUSE>0 Then Call Pause()
		rsUser.movenext
		i=i+1
	loop
	response.Write("</div>")
	set rsUser=nothing
	set blog=nothing
end sub

sub DoUpdatelog()
	Server.ScriptTimeOut=999999999
	dim BeginID,EndID,p1,rs,blog,i
	BeginID=trim(request("BeginID"))
	EndID=trim(request("EndID"))
	if BeginID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定开始ID</li>"
	else
		BeginID=Clng(BeginID)
	end if
	if EndID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定结束ID</li>"
	else
		EndID=Clng(EndID)
	end if
	if FoundErr=true then exit sub
	set rs=oblog.execute("select count(logid) from oblog_log where logid>=" & clng(BeginID) & " and logid<=" & clng(EndID))
	p1=rs(0)
	set rs=oblog.execute("select logid,userid from oblog_log where logid>=" & clng(BeginID) & " and logid<=" & clng(EndID)&" order by logid")
	set blog=new class_blog
	response.Write("<div style=""text-align: center;"">")
	response.Write("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
	i=1
	'blog.progress_init
	do while not rs.eof
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""进度：当前文章ID:"&rs(0)&""";</script>" & VbCrLf
		Response.Flush
		blog.userid=rs(1)
		blog.update_log rs(0),0
		If P_BLOG_UPDATEPAUSE>0 Then Call Pause()
		rs.movenext
		i=i+1
	loop
	response.Write("</div>")
	set rs=nothing
	set blog=nothing
end sub




Sub delblog()'删除外教资料！
	id = Request("id")
    If id = "" Then
        oblog.adderrstr ("请指定要删除的" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id=FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            deloneblog (n(i))
        Next
    Else
        deloneblog (id)
    End If
    oblog.showok "删除" & tName & "成功！", ""
End Sub

Sub deloneblog(logid)
    logid = CLng(logid)
    Dim uid, delname,rst,fso, sid,rst1
    Set rst = server.CreateObject("adodb.recordset")
	If Not IsObject(conn) Then link_database
    rst.open "select userid,logfile,subjectid,logtype from oblog_logCooperateSubmit where logid="&logid,conn,1,3
    If Not rst.EOF Then
        uid = rst(0)
        delname = Trim(rst(1))
        sid = rst(2)
        '清理图片记录
        If rst("logtype")=1 Then
        	Call DeletePhotos(logid)
   	 	End If
		'真实域名需要重新整理文件数据
		if true_domain=1 and delname<>"" then
			if instr(delname,"archives") then
				delname=right(delname,len(delname)-InstrRev(delname,"archives")+1)
			else
				delname=right(delname,len(delname)-InstrRev(delname,"/"))
			end if
			if oblog.logined_ulevel=9 then
				set rst1=oblog.execute("select user_dir,user_folder from oblog_user where userid="&clng(uid))
				if not rst1.eof then
					delname=rst1(0)&"/"&rst1(1)&"/"&delname
				end if
				set rst1=nothing
			else
				delname=oblog.logined_udir&"/"&oblog.logined_ufolder&"/"&delname
			end if
			'response.write(delname)
			'response.end
		end if
'''        If delname <> "" Then
'''            Set fso = server.CreateObject("Scripting.FileSystemObject")
'''            If fso.FileExists(server.MapPath(delname)) Then fso.DeleteFile server.MapPath(delname)
'''        End If
        '''rst.Delete'''WL;不删除记录，改为令其’隐藏‘！
			oblog.execute ("update oblog_logCooperateSubmit set isDelete=1 where logid=" & CLng(logid))'是否已被用户删除！表示此外教资料过期！！好让我们对照处理...WL;
			oblog.execute ("update oblog_logCooperateSubmit set isAdminDeleted=1 where logid=" & CLng(logid))
        rst.Close
        '---------------------------------
'''		Call Tags_UserDelete(logid)
        '--------------------------------------------
        '此处需要进一步修改，计数器
        oblog.execute ("update oblog_user set log_count=log_count-1 where userid=" & uid)
        oblog.execute ("delete from oblog_comment where mainid=" & CLng(logid))
        oblog.execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & CLng(sid))
        '--------------------------------------------
'''        Dim blog
'''        Set blog = New class_blog
'''        blog.userid = uid
'''        blog.Update_Subject uid
'''        blog.update_index 0
'''        blog.update_newblog (uid)
'''        Set blog = Nothing
        Set fso = Nothing
        Set rst = Nothing  
		
		
    Else
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
End Sub



Sub change_state_shenhe()'审核操作
	dim sql,rs,id,userid
	Dim isValidate
	Dim templogPublic_path
	
	id=trim(Request("logid"))
	userid= Trim(Request("userid"))
	isValidate= Cint(Request("state_shenhe"))
	
	templogPublic_path= Trim(Request("logPublic_path"))
	
	if templogPublic_path="" or Len(templogPublic_path)<6 then
		'FoundErr=True
		'ErrMsg=ErrMsg & "<br><li>请先 关联发布地址URL ,才能通过审核 ！</li>"
		'exit sub
	else
		id=CLng(id)
	end if
	
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	if userid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>userid参数不足！</li>"
		exit sub
	else
		userid= CLng(userid)
	end if
	

	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if

Dim TempWord
	If isValidate= 1 Then
		conn.execute("update oblog_logCooperateSubmit set state_shenhe=0 where logid=" & id )
		
		conn.execute("Update oblog_user Set passTeacher_count = passTeacher_count-1 Where userid=" & userid )
		TempWord = "成功通过审核！"
	ElseIf isValidate= 0 Then
		conn.execute("update oblog_logCooperateSubmit set state_shenhe=1 where logid=" & id )
		
		conn.execute("Update oblog_user Set passTeacher_count = passTeacher_count+1 Where userid=" & userid )
		TempWord = "已取消审核！"
	End If

	oblog.showok "已取消了审核！","admin_userCooperate_PassTeacher.asp"
	
End Sub


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