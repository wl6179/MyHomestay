<!--#include file="inc/inc_sys.asp"-->
<%
dim rs, sql
dim Action,iCount,adminname,strPara
strPara=LCase(Request.QueryString)
Action=Trim(request("Action"))
adminname=session("adminname")
CheckSafePath(0)
%>
<html>
<head>
<title>管理员管理</title>
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
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}

function CheckAdd()
{
  if(document.form1.username.value=="")
    {
      alert("用户名不能为空！");
	  document.form1.username.focus();
      return false;
    }
    
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
  if (document.form1.Purview[1].checked==true){
	GetClassPurview();
  }
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}

</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="tdbg" style="background:#ffffff">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>管 理 员 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="admin_admin.asp">管理员管理首页</a>&nbsp;|&nbsp;<a href="admin_admin.asp?Action=Add">新增管理员</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddAdmin()
elseif Action="SaveAdd" then
	If CheckSafePath(0) Then call SaveAdd()
elseif Action="Del" then
	If CheckSafePath(0) Then call DelAdmin()
else
	call main()
end if

sub main()
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from oblog_admin order by id"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<br>
<table width='100%' border="0" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="admin_admin.asp" onSubmit="return confirm('确定要删除选中的管理员吗？');">
     <td>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr align="center" class="title"> 
            <td  width="38"><font color="#FFFFFF"><strong>选中</strong></font></td>
            <td  width="38" height="22"><font color="#FFFFFF"><strong> 序号</strong></font></td>
            <td width="210" height="22"><font color="#FFFFFF"><strong> 用 户 名</strong></font></td>
            <td width="226"><font color="#FFFFFF"><strong>最后登录IP</strong></font></td>
            <td width="159"><font color="#FFFFFF"><strong>最后登录时间</strong></font></td>
            <td  width="78"><font color="#FFFFFF"><strong>登录次数</strong></font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="38"><input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" <%if rs("UserName")=AdminName then response.write " disabled"%> onClick="unselectall()"></td>
            <td width="38"><%=rs("ID")%></td>
            <td> 
              <%
	if rs("username")=AdminName then
		response.write "<font color=red><b>" & rs("UserName") & "</b></font>"
	else
		response.write rs("UserName")
	end if
	%>
            </td>
            <td width="226"> 
              <%
	if rs("LastLoginIP")<>"" then
		response.write rs("LastLoginIP")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="159"> 
              <%
	if rs("LastLoginTime")<>"" then
		response.write rs("LastLoginTime")
	else
		response.write "&nbsp;"
	end if
	%>
            </td>
            <td width="78"> 
<%
    If Not IsNull(rs("LoginTimes")) Then
		If rs("LoginTimes")<>"" Then
			response.write rs("LoginTimes")
		Else
			response.write 0
			oblog.execute ("update [oblog_admin]  set LoginTimes=0 where id="&uid)
		End If        
	Else
		oblog.execute ("update [oblog_admin]  set LoginTimes=0")
	End if
	%>  
            </td>
          </tr>
          <%
	rs.MoveNext
loop
  %>
        </table>  
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="200" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有管理员</td>
            <td><input name="Action" type="hidden" id="Action" value="Del">
              <input name="Submit" type="submit" id="Submit" value="删除选中的管理员"></td>
  </tr>
</table>
</td>
</form></tr></table>
<%
	rs.Close
	set rs=Nothing
end sub

sub AddAdmin()
%>
<form method="post" action="admin_admin.asp" name="form1" onSubmit="javascript:return CheckAdd();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong><font color="#FFFFFF">新 
          增 管 理 员</font></strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><div align="right">用 户 名：</div></td>
      <td width="65%" class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><div align="right">初始密码： </div></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="Password">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><div align="right">确认密码：</div></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2"><div align="center">
          <input name="Action" type="hidden" id="Action" value="SaveAdd">
          <input  type="submit" name="Submit" value=" 添 加 " style="cursor:hand;">
          &nbsp; 
          <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;">
        </div></td>
    </tr>
  </table>
</form>
<%
end sub
%>
<%
sub SaveAdd()
	dim username, password,PwdConfirm
	If Instr(strPara,"password") Then		
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""注意：外部恶意连接！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		Exit Sub
	End If
	username=trim(Request("username"))
	password=trim(Request("Password"))
	sql="Select * from oblog_admin where username='"&username&"'"
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not (rs.bof and rs.EOF) then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""对不起！此用管理员已经存在，请更换用户名再进行注册！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		rs.close
		set rs=nothing
		exit sub
	end if
   	rs.addnew
 	rs("username")=username
   	rs("password")=md5(password)
	rs.update
    rs.Close
	set rs=Nothing
	Call main()
end sub


sub DelAdmin()
	dim UserID
	If Instr(strPara,"del") Then		
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""注意：外部恶意连接！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		Exit Sub
	End If
	UserID=trim(Request("ID"))
	if UserID="" then
		Response.Write"<script language=JavaScript>"
		Response.Write"alert(""请选择要删除的管理员！"");"
		Response.Write"window.history.go(-1);"
		Response.Write"</script>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Select * from oblog_Admin where ID in (" & UserID & ")"
	else
		UserID=clng(UserID)
		sql="select * from oblog_Admin where ID=" & UserID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	do while not rs.eof
		rs.delete
		rs.update
		rs.movenext
	loop
	rs.close
	set rs=nothing
	call main()
end sub

%>