<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,TotalPages
dim rs, sql
dim UserID,UserSearch,Keyword,strField
dim Action,FoundErr,ErrMsg
dim tmpDays
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
UserSearch=trim(request("UserSearch"))
Action=trim(request("Action"))
UserID=trim(Request("UserID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")

if UserSearch="" then
	UserSearch=0
else
	UserSearch=Clng(UserSearch)
end if
strFileName="admin_UserLogclass_Nav.asp?UserSearch=" & UserSearch
if strField<>"" then
	strFileName=strFileName&"&Field="&strField
end if
if keyword<>"" then
	strFileName=strFileName&"&keyword="&keyword
end if
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>注册</title>
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
    <td height="22" colspan=2 align=center><strong>注 册 栏目空间 管 理</strong></td>
  </tr>
  <form name="form1" action="admin_UserLogclass_Nav.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>栏目空间配置：</strong></td>
      <td width="687" height="30"><select style="display:none;" size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>请选择查询条件</option>
		  <option value="0">最后注册的500个栏目空间</option>
          <option value="1">文章最多TOP100</option>
          <option value="2">文章最少的100个栏目空间</option>
		  <option value="3">VIP栏目空间</option>
		  <option value="4">前台管理员</option>
		  <option value="5">推荐栏目空间</option>
          <option value="6">等待管理员认证的栏目空间</option>
          <option value="7">所有被锁住的栏目空间</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_UserLogclass_Nav.asp">栏目空间管理首页</a>&nbsp;|&nbsp;<a href="admin_UserLogclass_Nav.asp?Action=Add">添加新栏目空间</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_UserLogclass_Nav.asp">
  <tr class="tdbg">
    <td width="120"><strong>栏目空间查询：</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="UserName" selected>栏目空间名</option>
      <option value="UserID" >栏目空间ID</option>
	 <!-- <option value="nickname" >栏目空间昵称</option>-->
      <option value="ip" >最后登录ip</option>
	  <option value="blogname" >中文名称</option>
	  <option value="lastlogintime" >多少天内未登录</option>
	  <option value="logcount" >栏目文章数小于</option>
	  <option value="logintimes" >栏目登录次数小于</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  若为空，则查询所有栏目空间</td>
  </tr>
</form>
</table>
<%
if Action="Add" then
	call AddUser()
elseif Action="SaveAdd" then
	call sub_chkreg()
elseif Action="Sheaf" then
	call Sheaf()	
elseif Action="SaveSheaf" then
	call SaveSheaf()
elseif Action="Modify" then
	call Modify()
elseif Action="SaveModify" then
	call SaveModify()
elseif Action="Del" then
	call DelUser()
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
	
ElseIf Action="ClearBinding" Then
	Call ClearBinding()
	
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>您现在的位置：<a href='admin_UserLogclass_Nav.asp'>注册栏目空间管理</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_user where user_level>6 order by UserID desc"
			strGuide=strGuide & "最后注册的500个栏目空间"
		case 1
			sql="select top 100 * from oblog_user where user_level>6 order by log_count desc"
			strGuide=strGuide & "发布文章最多的前100个栏目空间"
		case 2
			sql="select top 100 * from  oblog_user where user_level>6 order by log_count"
			strGuide=strGuide & "发布文章最少的100个栏目空间"
		case 3
			sql="select * from  oblog_user where user_level=8 order by userid desc"
			strGuide=strGuide & "所有VIP栏目空间"
		case 4
			sql="select * from  oblog_user where user_level=9 order by userid desc"
			strGuide=strGuide & "前台管理发布"
		case 5
			sql="select * from  oblog_user where user_isbest=1 order by userid desc"
			strGuide=strGuide & "推荐栏目空间"
		case 6
			sql="select * from oblog_user where User_Level=6 order by userID  desc"
			strGuide=strGuide & "等待管理认证证的栏目空间"
		case 7
			sql="select * from oblog_user where  LockUser =1 order by userID  desc"
			strGuide=strGuide & "所有被锁住的栏目空间"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_user where user_level>6 order by userID desc"
				strGuide=strGuide & "所有栏目空间"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>栏目空间ID必须是整数！</li>"
					else
						sql="select * from oblog_user where user_level>6 and userID =" & Clng(Keyword)
						strGuide=strGuide & "栏目空间ID等于<font color=red> " & Clng(Keyword) & " </font>的栏目空间"
					end if
				case "UserName"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and username like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "栏目空间名中含有“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					else
						sql="select * from oblog_user where user_level>6 and username= '" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "栏目空间名等于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					end if
					
				case "nickname"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and nickname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "栏目空间昵称中含有“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					else
						sql="select * from oblog_user where user_level>6 and nickname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "栏目空间昵称等于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					end if
				case "ip"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and lastloginip like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "最后登录ip中含有“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					else
						sql="select * from oblog_user where user_level>6 and lastloginip='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "最后登录ip等于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					end if
				case "blogname"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and blogname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "栏目名中含有“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					else
						sql="select * from oblog_user where user_level>6 and blogname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "栏目名等于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
					end if
				case "logcount"
					sql="select top 500 * from oblog_user where user_level>6 and log_count < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "文章数小于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
				case "logintimes"
					sql="select top 500 * from oblog_user where user_level>6 and logintimes < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "登录次数小于“ <font color=red>" & Keyword & "</font> ”的栏目空间"
				case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="select top 500 * from oblog_user where user_level>6 and datediff(d,lastlogintime,getdate())>"&clng(Keyword)&" order by userID  desc"
				   Else 
				   	sql="select top 500 * from oblog_user where user_level>6 and datediff('d',lastlogintime,now())>"&clng(Keyword)&" order by userID  desc"
				   End If 
					strGuide=strGuide & "“ <font color=red>" & Keyword & "</font> ”天内未登录的栏目空间"
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
		strGuide=strGuide & "共找到 <font color=red>0</font> 个栏目空间</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个栏目空间</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个栏目空间")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个栏目空间")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个栏目空间")
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
  <form name="myform" method="Post" action="admin_UserLogclass_Nav.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">选中</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> 栏目空间名</font></td>
            <td height="22" align="center"><font color="#FFFFFF">所属栏目空间组</font></td>
            <td height="22" align="center"><font color="#FFFFFF">最后登录IP</font></td>
            <td align="center"><font color="#FFFFFF">最后登录时间</font></td>
            <td width="60" height="22" align="center"><font color="#FFFFFF">登录次数</font></td>
            <td width="40" height="22" align="center"><font color="#FFFFFF"> 状态</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF"> 
              操作</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='UserID' type='checkbox' onClick="unselectall()" id="UserID" value='<%=cstr(rs("userID"))%>'></td>
            <td width="30" align="center"><%=rs("userID")%></td>
            <td width="80" align="center"><%
			if Cint(rs("logclassid_Nav"))>0 then
				response.write "<font style='cursor:hand;' title='已绑定栏目空间' color=red>√&nbsp;</font>"
			end if
			response.write "<a href='../blog.asp?name=" & rs("userName") & "'" 			
			response.write """ target='_blank'>" & rs("userName") & "</a>"
			%> </td>
            <td align="center"> <%
			select case rs("User_Level")
				case 6
					response.write "<font color=green>注册的家庭用户</font>"
				case 7
					response.write "注册栏目发布"
				case 8
					response.write "<font color=blue>VIP栏目发布</font>"
				case 9
					response.write "<font color=blue>前台管理发布</font>"
				case else
					response.write "<font color=red>异常注册</font>"
			end select
			%> </td>
            <td align="center"> <%
	if rs("LastLoginIP")<>"" then
		response.write rs("LastLoginIP")
	else
		response.write "&nbsp;"
	end if
	%> </td>
            <td align="center"> <%
	if rs("LastLoginTime")<>"" then
		response.write rs("LastLoginTime")
	else
		response.write "&nbsp;"
	end if
	%> </td>
            <td width="60" align="center"> <%
	if rs("LoginTimes")<>"" then
		response.write rs("LoginTimes")
	else
		response.write "0"
	end if
	%> </td>
            <td width="40" align="center"><%
	  if rs("LockUser")=1 then
	  	response.write "<font color=red>已锁定</font>"
	  else
	  	response.write "正常"
	  end if
	  %></td>
            <td width="120" align="center"><%
		response.write "<!--<a href='admin_UserLogclass_Nav.asp?Action=Modify&UserID=" & rs("userID") & "'>修改</a>-->&nbsp;"
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>进入栏目发布文章</a><br />&nbsp;"
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=Sheaf&userid=" & rs("userid") & "'><font color='#038ad7'>绑定分类</font></a>&nbsp;|"
		if rs("LockUser")=0 then
			response.write "<a href='admin_UserLogclass_Nav.asp?Action=Lock&UserID=" & rs("userID") & "'>锁定</a>&nbsp;"
		else
            response.write "<a href='admin_UserLogclass_Nav.asp?Action=UnLock&UserID=" & rs("userID") & "'>解锁</a>&nbsp;"
		end if
        response.write "<a href='admin_UserLogclass_Nav.asp?Action=Del&UserID=" & rs("userID") & "' onClick='return confirm(""确定要删除此栏目空间吗？"");'>删除</a>&nbsp;"
		%>
		<% If rs("logclassid_Nav")>0 Then %>
		 | <a href="admin_UserLogclass_Nav.asp?Action=ClearBinding&id=<%=rs("userid")%>" onClick="return confirm('确定强制清除绑定状态吗？');"><img src="/images/del.gif" title="强制删除和文章分类的绑定状态！！" /></a>
		 <% End If %></td>
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
              选中本页显示的所有栏目空间</td>
            <td> <strong>操作：</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=true">
              删除&nbsp;&nbsp;&nbsp;&nbsp; 
              <input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">移动到
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">注册的家庭用户</option>
                <option value="7">注册栏目发布</option>
                <option value="8">VIP栏目发布</option>
                <option value="9">前台管理发布</option>
              </select>
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" 执 行 "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub




sub AddUser()


	if not IsObject(conn) then link_database

%>
<FORM name="oblogform" action="admin_UserLogclass_Nav.asp" method="post" onSubmit='return VerifySubmit()'>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">添加新栏目空间信息</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><font color="red">*</font>&nbsp;新建栏目名（即文件夹名，例如 about_us ）：</TD>
      <TD width="60%"><input name="username" id="username" type="text" value="" size=15 maxlength=30 />&nbsp;&nbsp;<a href="javascript:checkssn('../checkssn_Nav.asp')";>查看此栏目名是否可用</a></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" <% if trim(oblog.setup(4,0))="" or oblog.setup(12,0)<>1 then response.Write("style='display:none;'")%>>
      <td>栏目空间域名：</td>
      <td>
	  <%
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		response.Write "<input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>" 
	else
		response.Write "<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"
	end if
	%>
	  
	  </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td><font color="red">*</font>&nbsp;栏目中文名（最好与文章分类的中文名一致！）：</td>
      <td><input name=blogname   type=text id="blogname" value="" size=30 maxlength=20></td>
    </tr>
	<tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td><font color="blue">英文名</font>（如"About Us"，可以不填）：</td>
      <td><input name=EnglishName   type=text id="EnglishName" value="" size=30 maxlength=20></td>
    </tr>
	
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>栏目空间绑定的顶级域名：</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="" size=30 maxlength=20></td>
    </tr>
	<%end if%>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>blog类别：</td>
      <td><select name="usertype" id="usertype">
          <option value="0" selected="selected">0</option>
        </select></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="40%"><font color="red">*</font>&nbsp;密码(至少6位)：<BR>
        请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
      <TD width="60%"> <input name=password type=password id=password size=15 maxlength=30> <font color="#FF0000">对此栏目留一个管理密码</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD><font color="red">*</font>&nbsp;确认密码(至少6位)：<br>
        请再输一遍确认</TD>
      <TD><input name=repassword type=password id=repassword size=15 maxlength=30> <font color="#FF0000">对此栏目留一个管理密码</font> </TD>
    </TR>
    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">密码问题：<br>
        忘记密码的提示问题</TD>
      <TD width="60%"> <input value="MyHomestay" name=question type=text id=question size=30 maxlength=30> 
      </TD>
    </TR>
    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">问题答案：<BR>
        忘记密码的提示问题答案，用于取回密码</TD>
      <TD width="60%"> <input value="MyHomestay" name=an type=text id=an size=30 maxlength=30> <font color="#FF0000"></font></TD>
    </TR>

    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">Email地址：</TD>
      <TD width="60%"> <input value="service@myhomestay.com.cn" name=email type=text size=15 maxlength=30> 
        
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">确认同意：</TD>
      <TD width="60%"> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked>同意　<input class='input_radio' type=radio name=passregtext id=passregtext value='0'>不同意</TD>
    </TR>

    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="添加新栏目空间"> </TD>
    </TR>
  </TABLE>
</form>
<%

end sub




sub Sheaf()'绑定分类&栏目空间
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_user where userID=" & UserID
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目空间！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">栏目空间—绑定—>文章分类</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">栏目空间名：</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
   
   
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>捆绑文章分类：</td>
      <td><select name="classid" id="classid" <%If rsUser("logclassid_Nav")>0 Then Response.Write("disabled='disabled'")%>>
	<%=oblog.show_class2("log",Cint(rsUser("logclassid_Nav")),0)%> <%''"log",classid,t=1图%>
	</select></td>
    </tr>
   
   
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveSheaf"> <input name=Submit   type=submit id="Submit" value="保存绑定"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("userID")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsUser.close
	set rsUser=nothing
end sub



sub SaveSheaf()
	dim UserID,classid
	dim rsUser,sqlUser
	classid=Cint(request("classid"))
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	
	if classid=0 then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请选择要绑定的文章分类！</li>"
		exit sub
	else
		classid=Clng(classid)
	end if
	
	
	if classid>0 then
		dim sql_temp,rs_temp
		sql_temp="select parentid,userid_Nav from oblog_logclass where classid=" & classid
		Set rs_temp=Server.CreateObject("Adodb.RecordSet")
		if not IsObject(conn) then link_database
		rs_temp.Open sql_temp,Conn,1,3
		
		if rs_temp.bof and rs_temp.eof then
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>找不到所选的文章分类！</li>"
			rs_temp.close
			set rs_temp=nothing
			exit sub
		else'如果不出意外，立刻分析其记录是否为 主分类——一级分类！
			If rs_temp(0)>0 Then'如果 不是主分类(一级分类)，报错；
				founderr=true
				ErrMsg=ErrMsg & "<br><li>请您选择 一级文章分类！</li>"
				exit sub
			elseIf rs_temp(1)>0 Then'如果已经有绑定了，报错；
				founderr=true
				ErrMsg=ErrMsg & "<br><li>此文章分类已被绑定，请选择其它文章分类！</li>"
				exit sub
			End If
		end if
				
	else'如果两项检测均没有问题，则通过，赋予其数值类型；
		classid=Clng(classid)
	end if
	
	
	if founderr=true then
		set rsUser=nothing
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_user where userID=" & UserID
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目空间！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	
	Dim Update_sql
	Update_sql="Update oblog_logclass set userid_Nav=0 where classid=" & rsUser("logclassid_Nav")
	Conn.Execute Update_sql
	
	rsUser("logclassid_Nav")=classid

	rsUser.update
	rsUser.Close
	set rsUser=nothing
	
	sql="Update oblog_logclass set userid_Nav="& UserID &" where classid=" & classid
	Conn.Execute sql
	
	oblog.showok "栏目空间捆绑成功!",""
end sub



sub Modify()
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_user where userID=" & UserID
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目空间！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">修改注册栏目空间信息</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">栏目空间名：</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>栏目空间域名：</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <select name="user_domainroot" ><%=oblog.type_domainroot(rsuser("user_domainroot"))%></select></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>栏目名：</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>栏目空间绑定的顶级域名：</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="<%=rsuser("custom_domain")%>" size=30 maxlength=20></td>
    </tr>
	<%end if%>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>blog类别：</td>
      <td><select name="usertype" id="usertype">
          <%if rsUser("user_classid")<>"" then
	  response.Write(oblog.show_class("user",rsUser("user_classid"),0))
	  else
	  response.Write(oblog.show_class("user",0,0))
	  end if
	  %>
        </select></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="40%">密码(至少6位)：<BR>
        请输入密码，区分大小写。 请不要使用任何类似 '*'、' ' 或 HTML 字符 </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">如果不想修改，请留空(整合栏目空间请到论坛修改)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD>确认密码(至少6位)：<br>
        请再输一遍确认</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12> <font color="#FF0000">如果不想修改，请留空(整合栏目空间请到论坛修改)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">密码问题：<br>
        忘记密码的提示问题</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30>(整合栏目空间请到论坛修改) 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">问题答案：<BR>
        忘记密码的提示问题答案，用于取回密码</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">如果不想修改，请留空(整合栏目空间请到论坛修改)</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">性别：</TD>
      <TD width="60%"> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then response.write "CHECKED"%>>
        男 &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then response.write "CHECKED"%>>
        女</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">Email地址：</TD>
      <TD width="60%"> <INPUT name=Email value="<%=rsUser("userEmail")%>" size=30   maxLength=50> 
        <a href="mailto:<%=rsUser("userEmail")%>">给此栏目空间发一封电子邮件</a> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">OICQ号码：</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">MSN：</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsUser("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">栏目空间级别：</TD>
      <TD width="60%"><select name="User_Level" id="User_Level">
          <option value="6" <%if clng(rsUser("user_level"))=6 then response.Write("selected")%>>注册的家庭用户</option>
          <option value="7" <%if clng(rsUser("user_level"))=7 then response.Write("selected")%>>注册栏目发布</option>
          <option value="8" <%if clng(rsUser("user_level"))=8 then response.Write("selected")%>>vip栏目发布</option>
          <option value="9" <%if clng(rsUser("user_level"))=9 then response.Write("selected")%>>前台管理发布</option>
        </select></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>可上传空间(kb)：</td>
      <td><input name=user_upfiles_max   type=text value="<%=rsuser("user_upfiles_max")%>" size=20 maxlength=20>
        为零时为系统默认设置，如需单独设置，请输入kb值</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>已上传字节(字节)：</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=20 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD>是否为推荐栏目空间：</TD>
      <TD><input type="radio" name="isbest" value=1 <%if rsUser("user_isbest")=1 then response.write "checked"%>>
        是 &nbsp;&nbsp; <input type="radio" name="isbest" value=0 <%if rsUser("user_isbest")<>1 then response.write "checked"%>>
        否</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">栏目空间目录：</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsUser("user_dir")%>" size=30 maxLength=50>
        如无必要请不要修改，否则将造成栏目空间目录混乱</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">栏目空间状态：</TD>
      <TD width="60%"><input type="radio" name="LockUser" value=0 <%if rsUser("LockUser")=0 then response.write "checked"%>>
        正常&nbsp;&nbsp; <input type="radio" name="LockUser" value=1 <%if rsUser("LockUser")=1 then response.write "checked"%>>
        锁定</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("userID")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rsUser.close
	set rsUser=nothing
end sub


sub UpdateUser()
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp?action=DoUpdate" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">更新栏目空间静态页面</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成栏目空间静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3 、本操作根据栏目空间ｉｄ更新。 </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">开始栏目空间ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      栏目空间ID，可以填写您想从哪一个ID号开始进行更新</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">结束栏目空间ID：</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      将更新开始到结束ID之间的栏目空间数据，之间的数值最好不要选择过大</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="生成静态页面"></td>
  </tr>
</table>
</form>

<FORM name="Form1" action="admin_UserLogclass_Nav.asp?action=DoUpdatelog" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">更新文章静态页面</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          1、本操作将重新生成栏目空间静态页面。<br>
          2、本操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。<br>
      3、本操作根据文章ｉｄ更新。</p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">开始文章ID：</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      栏目空间ID，可以填写您想从哪一个ID号开始进行更新</td>
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
	response.Write("整合外部数据库栏目空间，不能使用此功能"):response.End()
end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">登录到栏目空间管理后台</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>说明：<br>
          本操作供管理员登录到栏目空间的管理界面进行管理。<br>
          当栏目空间操作出现障碍时，可进入该栏目空间后台，协助栏目空间进行操作。<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">栏目空间账号：</td>
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
		response.Write("整合外部数据库栏目空间，不能使用此功能"):response.End()
	end if
	dim rs,username
	username=oblog.filt_badstr(trim(request("username")))
	if username="" then response.Write("栏目空间名不能为空"):response.End()
	if is_ot_user=0 then
		set rs=oblog.execute("select username,password from oblog_user where username='"&username&"'")
	else
		set rs=ot_conn.execute("select "&ot_username&","&ot_password&" from "&ot_usertable&" where "&ot_username&"='"&username&"'")
	end if
	if not rs.eof then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		set rs=nothing
		response.Write(oblog.logined_uid)
		If Request("username")="myhomestay" Then '如果是消息管理员帐号myhomestay，那么，直接进入消息管理页面user_pmmanage.asp
			response.Redirect("../user_pmmanage.asp")
		Else
			response.Redirect("../HomestayBackctrl-index.asp")
		End If
	else
		set rs=nothing
		response.Write("无此栏目空间"):response.End()
	end if
	
end sub
sub SaveModify()
	dim UserID,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,User_Level,LockUser,isbest
	dim rsUser,sqlUser
	dim blogname,usertype,user_upfiles_max,upfiles_size,user_domain,user_domainroot
	Action=trim(request("Action"))
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	Password=trim(request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	Question=trim(request("Question"))
	Answer=trim(request("Answer"))
	Sex=trim(Request("Sex"))
	Email=trim(request("Email"))
	Homepage=trim(request("Homepage"))
	OICQ=trim(request("OICQ"))
	MSN=trim(request("MSN"))
	User_Level=trim(request("User_Level"))
	isbest=trim(request("isbest"))
	LockUser=trim(request("LockUser"))
	blogname=trim(request("blogname"))
	usertype=trim(request("usertype"))
	user_upfiles_max=trim(request("user_upfiles_max"))
	upfiles_size=trim(request("upfiles_size"))
	user_domain=trim(request("user_domain"))
	user_domainroot=trim(request("user_domainroot"))
	if is_ot_user=0 then
		if Password<>"" then
			if oblog.strLength(Password)>12 or oblog.strLength(Password)<6 then
				founderr=true
				errmsg=errmsg & "<br><li>密码不能大于12小于6，如果你不想修改密码，请保持为空。</li>"
			end if
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>密码中含有非法字符，如果你不想修改密码，请保持为空。</li>"
				founderr=true
			end if
		end if
		if Password<>PwdConfirm then
			founderr=true
			errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
		end if
		if Question="" then
			founderr=true
			errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
		end if
	end if
	if Sex="" then
		founderr=true
		errmsg=errmsg & "<br><li>性别不能为空</li>"
	else
		sex=cint(sex)
		if Sex<>0 and Sex<>1 then
			Sex=1
		end if
	end if
	if is_ot_user=0 then
		if Email="" then
			founderr=true
			errmsg=errmsg & "<br><li>Email不能为空</li>"
		else
			if oblog.IsValidEmail(Email)=false then
				errmsg=errmsg & "<br><li>您的Email有错误</li>"
				founderr=true
			end if
		end if
	end if
	if OICQ<>"" then
		if not isnumeric(OICQ) or len(cstr(OICQ))>10 then
			errmsg=errmsg & "<br><li>OICQ号码只能是4-10位数字，您可以选择不输入。</li>"
			founderr=true
		end if
	end if
	if MSN<>"" then
		if oblog.IsValidEmail(MSN)=false then
			errmsg=errmsg & "<br><li>你的MSN有误。</li>"
			founderr=true
		end if
	end if

	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定栏目空间级别！</li>"
	else
		User_Level=CLng(User_Level)
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		set rsuser=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"' and userid<>"&userid)
		if not rsuser.eof then 
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>此域名已经被其他人使用！</li>"
		end if
	end if

	if founderr=true then
		set rsuser=nothing
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from oblog_user where userID=" & UserID
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目空间！</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	if Password<>"" then
		rsUser("password")=md5(Password)
	end if
	rsUser("Question")=Question
	if Answer<>"" then
		rsUser("Answer")=md5(Answer)
	end if
	rsUser("Sex")=Sex
	rsUser("userEmail")=Email
	if oicq="" then
		oicq=0
	end if
	rsUser("qq")=OICQ
	rsUser("Msn")=MSN
	rsUser("User_Level")=User_Level
	rsUser("LockUser")=LockUser
	rsuser("user_isbest")=isbest
	rsuser("blogname")=blogname
	rsuser("user_classid")=usertype
	rsuser("user_upfiles_max")=user_upfiles_max
	rsuser("user_upfiles_size")=upfiles_size
	rsuser("user_dir")=trim(request("user_dir"))
	if true_domain=1 then
		rsuser("custom_domain")=trim(request("custom_domain"))
	end if
	if trim(request("user_domain"))<>"" then
		rsuser("user_domain")=trim(request("user_domain"))
	else
		rsuser("user_domain")=" "
	end if
	rsuser("user_domainroot")=trim(request("user_domainroot"))
	rsUser.update
	rsUser.Close
	set rsUser=nothing
	oblog.showok "修改成功!",""
end sub


sub DelUser()
	dim rs,i
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定要删除的栏目空间</li>"
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
	response.redirect "admin_UserLogclass_Nav.asp"
end sub

sub deloneuser(userid)
	userid=clng(userid)
	dim rs,fso,f,uname,udir
	set rs=oblog.execute("select user_dir,username,user_folder from oblog_user where userid="&userid)
	if not rs.eof then
		udir=rs(0)
		uname=rs(1)
		set fso=server.createobject("scripting.filesystemobject")
		if fso.FolderExists(server.MapPath(blogdir & udir&"/"&rs("user_folder"))) then 
			Set f = fso.GetFolder(server.MapPath(blogdir & udir&"/"&rs("user_folder")))
			f.delete true
		end if
		set f=nothing
		set fso=nothing
		oblog.execute("delete from oblog_log where userid="&userid)
		oblog.execute("delete from oblog_comment where userid="&userid)
		oblog.execute("delete from oblog_message where userid="&userid)
		oblog.execute("delete from oblog_subject where userid="&userid)
		oblog.execute("delete from oblog_user where userid=" & userid)
		oblog.execute("delete from oblog_upfile where userid=" & userid)
		oblog.execute("delete from oblog_friend where userid=" & userid)
		oblog.execute("update oblog_pm set dels=1 where sender='" &uname&"'")
	end if
	set rs=nothing
end sub

sub LockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的栏目空间</li>"
		exit sub
	end if
	userid=clng(userid)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select lockuser,user_dir from oblog_user where userid="&userid,conn,1,3
	if not rs.eof then
		dim strtmp,fso
		rs(0)=1
		udir=rs(1)
		rs.update
		strtmp="<script language=javascript>"&vbcrlf
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=此栏目空间已经被锁定，请联系管理员!');"&vbcrlf
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
	oblog.showok "锁定栏目空间成功",""
end sub

sub UnLockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请选择要锁定的栏目空间</li>"
		exit sub
	end if
	userid=clng(userid)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select lockuser,user_dir from oblog_user where userid="&userid,conn,1,3
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
	oblog.showok "解锁栏目空间成功",""
end sub

sub MoveUser()
	dim msg
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定要移动的栏目空间</li>"
		exit sub
	end if
	dim User_Level
	User_Level=trim(request("User_Level"))
	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>请指定目标栏目空间组</li>"
		exit sub
	else
		User_Level=Clng(User_Level)
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>注册的家庭用户</font>”！"
			sql="Update oblog_user set User_Level=6 where userID in (" & UserID & ")"
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>注册栏目发布</font>”！"
			sql="Update oblog_user set User_Level=7 where userID in (" & UserID & ")"
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>VIP栏目发布</font>”！"
			sql="Update oblog_user set User_Level=8 where userID in (" & UserID & ")"
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>前台管理发布</font>”！"
			sql="Update oblog_user set User_Level=9 where userID in (" & UserID & ")"
		end if
	else
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>注册的家庭用户</font>”！"
			sql="Update oblog_user set User_Level=6 where userID=" & UserID 
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>注册栏目发布</font>”！"
			sql="Update oblog_user set User_Level=7 where userID="& UserID
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>VIP栏目发布</font>”！"
			sql="Update oblog_user set User_Level=8 where userID=" & UserID 
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;已经成功将选定栏目空间设为“<font color=blue>前台管理发布</font>”！"
			sql="Update oblog_user set User_Level=9 where userID=" & UserID
		end if
	end if
	Conn.Execute sql
	response.Redirect "admin_UserLogclass_Nav.asp"
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
	set rsuser=oblog.execute("select count(userid) from oblog_user where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID))
	p1=rsuser(0)
	set rsuser=oblog.execute("select userid from oblog_user where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID)&" order by userid")
	set blog=new class_blog
	response.Write("<div style=""text-align: center;"">")
	response.Write("<div class=""progress1""><div class=""progress2"" id=""progress1""></div></div><span id=""pstr1""></span><br><br>")
	i=1
	blog.progress_init
	do while not rsUser.eof
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""全部进度：当前栏目空间ID:"&rsuser(0)&""";</script>" & VbCrLf
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






sub sub_chkreg()
	if oblog.ChkPost()=false then
		oblog.adderrstr("系统不允许从外部提交！")
		oblog.showerr
		exit sub
	end if

	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	if oblog.chkiplock() then
		oblog.adderrstr("对不起！你的IP已被锁定,不允许注册！")
		oblog.showerr
		exit sub
	end if
	regusername=oblog.filt_badstr(trim(request("username")))
	if regusername="" then
		response.write "请输入栏目名"
	end if
	regpassword=request("password")
	re_regpassword=request("repassword")
	sex=request("sex")
	email=trim(request("email"))
	question=trim(request("question"))
	answer=trim(request("an"))
	blogname=trim(request("blogname"))
	usertype=clng(request("usertype"))
	'nickname=trim(request("nickname"))
	user_domain=Lcase(trim(request("domain")))
	user_domainroot=trim(request("user_domainroot"))
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("栏目空间名不能为空(不能大于14小于4)！")
	if oblog.chk_regname(regusername) then oblog.adderrstr("栏目空间名系统不允许注册！")
	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("栏目空间名中含有系统不允许的字符！")
	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("栏目空间名不允许全部为数字！")
	if oblog.chkdomain(regusername)=false then oblog.adderrstr("栏目空间名不合规范，只能使用小写字母，数字及下划线！")
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字！")
		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
	end if
	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("密码不能为空(不能大于14小于4)！")
	if re_regpassword="" then oblog.adderrstr("重复密码不能为空！")
	if regpassword<>re_regpassword then oblog.adderrstr("两次输入密码不同！")
	if question="" or oblog.strLength(question)>50 then oblog.adderrstr("找回密码提示问题不能为空(不能大于50)！")
	
	if answer="" or oblog.strLength(answer)>50 then
		oblog.adderrstr("找回密码问题答案不能为空(不能大于50)！")
	elseif answer="MyHomestay" then
		answer="49ba59abbe56e057"
	end if
	
	'if oblog.chk_regname(nickname) then oblog.adderrstr("此昵称系统不允许注册！")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("昵称不能不能大于50字符！")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("栏目中文名不能为空(不能大于50字符)！")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("栏目中文名中含有系统不允许的字符！")	
	if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("栏目空间名中含有非法字符！")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个昵称存在，请更改昵称！")		
	'end if
	if user_domain<>"" then
		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个域名存在，请更改域名！")
	end if
	
	if oblog.errstr<>"" then response.Write("<ul><li>错误："& oblog.errstr &"<input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>"):exit sub
	
	

	
	if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
	regpassword=md5(regpassword)
	if not IsObject(conn) then link_database
	set rsreg=server.CreateObject("adodb.recordset")
	rsreg.open "select * from [oblog_user] where username='"& oblog.filt_badstr(regusername) &"'",conn,1,3
	if rsreg.eof then
		rsreg.addnew
		rsreg("username")=regusername
		rsreg("password")=regpassword
		if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
			rsreg("user_domain")=user_domain
			rsreg("user_domainroot")=user_domainroot
		end if
		rsreg("question")=question
		rsreg("answer")=answer
		rsreg("useremail")=email
		rsreg("user_level")=7
		rsreg("user_isbest")=0
		rsreg("blogname")=blogname
		rsreg("nickname")=blogname
		rsreg("truename")=blogname
			'rsreg("EnglishName")=Request("EnglishName")
		
		rsreg("user_classid")=usertype
		'rsreg("nickname")=nickname
		rsreg("province")=request("province")
		rsreg("city")=request("city")
		rsreg("adddate")=ServerDate(now())
		rsreg("lastloginip")=oblog.userip
		rsreg("lastlogintime")=ServerDate(now())
		rsreg("user_dir")=oblog.setup(30,0)
		rsreg("user_folder")=regusername		
		rsreg.update
		oblog.execute("update oblog_setup set user_count=user_count+1")
		if is_unamedir=0 then
			oblog.execute("update oblog_user set user_folder=userid where username='"&oblog.filt_badstr(regusername)&"'")
		end if
		dim show_reg
		show_reg="→完成注册<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>新栏目"&blogname&"注册成功！</strong></li></ul>"
			show_reg=show_reg&"5秒后自动转返回。"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""history.go(-2)"",5000);"
			show_reg=show_reg&"</script>"
		else
			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
			show_reg=show_reg&"<ul><li><strong>恭喜！您已经注册成功！</strong></li>"
			show_reg=show_reg&"<li><a href='index.asp'>返回首页</a></li>"
			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>点击进入管理后台选择您的喜欢的页面风格(5秒后自动进入管理后台)</a></li></ul>"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
			show_reg=show_reg&"</script>"
		end if
	else
		response.Write("系统中已经有这个栏目空间名存在，请更改栏目空间名！<input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'>")
		
		exit sub	
	end if
	rsreg.close
	set rsreg=nothing
	'ATFLAG
	'Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=now()	
	Session("chk_regtime")=Now()
	response.Write(show_reg)
end sub

sub chk_regtime()
	dim lasttime
	'ATFLAG
	'lasttime = Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))
	lasttime = Session("chk_regtime")
	If IsDate(lasttime) Then 
		If DateDiff("s",lasttime,Now()) < chkregtime then
			Response.Write("<script language=javascript>alert('"&chkregtime&"秒后才能重复注册。');window.history.back(-1);</script>")
			Response.End
		end if
	end if
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

	
		conn.execute("Update oblog_user Set logclassid_Nav=0 Where userid=" & id )


	response.redirect "admin_UserLogclass_Nav.asp"
		
end sub


%>

<SCRIPT language="javascript">
<!--
function checkssn(gotoURL) {
   var ssn=oblogform.username.value;
   var domain=oblogform.domain.value;
   var domainroot=oblogform.user_domainroot.value;
	   var open_url = gotoURL + "?username=" + ssn+"&domain="+domain+"&domainroot="+domainroot;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=322,height=200');
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i)!='-')) 
{
return 1;
}
}
return 0;//pass
}
function del_space(s)
{
	for(i=0;i<s.length;++i)
	{
	 if(s.charAt(i)!=" ")
		break;
	}

	for(j=s.length-1;j>=0;--j)
	{
	 if(s.charAt(j)!=" ")
		break;
	}

	return s.substring(i,++j);
}

function VerifySubmit()
{
	 if (document.oblogform.passregtext[1].checked )
     {
        alert("您不同意注册条款，不能注册!");
	return false;
   }	 
   
 
		 
	uid = del_space(document.oblogform.username.value);
     if (uid.length == 0)
     {
        alert("请输入栏目空间名!");
	return false;
     }
	 
	pwd = del_space(document.oblogform.password.value);
     if (pwd.length == 0)
     {
        alert("请输入密码!");
	return false;
     }
	 
	 pwd = del_space(document.oblogform.password.value);
     if (pwd.length < 6)
     {
        alert("密码要大于6个字符！");
	return false;
     }

	pwd2 = del_space(document.oblogform.repassword.value);
     if (pwd2!=pwd)
     {
        alert("两次输入的密码不一致！");
		return false;
     }
	 
	 tishi = del_space(document.oblogform.question.value);
     if (tishi.length == 0)
     {
        alert("请输入密码提示问题");
        return false;
     }
	 
	 tsda = del_space(document.oblogform.an.value);
     if (tsda.length == 0)
     {
        alert("请输入密码提示问题答案！");
        return false;
     }
	 city = del_space(document.all("city").value);
     if (city.length == 0)
     {
        alert("请选择所在城市!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("请输入Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("请输入您的栏目名!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("请选择您的类别!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>

