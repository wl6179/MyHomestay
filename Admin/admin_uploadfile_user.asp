<!--#include file="inc/inc_sys.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim rs, sql
dim userid,UserSearch,Keyword,strField
dim Action
dim tmpDays
dim uppath,fso,thefile
dim del,moreid,delmore
moreid=trim(request("moreid"))
response.Write moreid
del=trim(request.QueryString("del"))
userid=trim(request.QueryString("userid"))
delmore=trim(request.QueryString("delmore"))
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
UserSearch=trim(request("UserSearch"))
Action=trim(request("Action"))
if UserSearch="" then
	UserSearch=0
else
	UserSearch=Clng(UserSearch)
end if
strFileName="admin_uploadfile_user.asp?UserSearch=" & UserSearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>上传文件管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>上传文件管理(用户列表)</strong></td>
  </tr>
  <form name="form1" action="admin_user.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>管理导航：</strong></td>
      <td width="687" height="30"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_uploadfile.asp">上传文件管理用户列表</a> <a href="admin_uploadfile.asp">上传文件管理文件列表</a></td>
    </tr>
  </form>
</table>
<br>
<%
if del<>"" then
	Delfile del
else
	call main()
end if

sub main()
	dim theFolder,filecount,totalsize,upstr
	upstr=" where user_upfiles_size>0"
	dim strGuide
	strGuide="<table width='98%' align='center'><tr><td align='left'>您现在的位置：<a href='Admin_Uploadfile_user.asp'>上传文件管理-用户列表</a>&nbsp;&gt;&gt;&nbsp;"

	if Keyword="" then
		sql="select top 500 user_upfiles_max,user_upfiles_size,username,userid,user_level,user_upfiles_num from [oblog_user] "&upstr&" order by user_upfiles_size desc"
		strGuide=strGuide & "前500个用户"
	else				
		sql="select top 500 user_upfiles_max,user_upfiles_size,username,userid,user_level,user_upfiles_num from [oblog_user] "&upstr&" and userName like '%" & Keyword & "%' order by user_upfiles_size  desc"
		strGuide=strGuide & "用户名中含有“ <font color=red>" & Keyword & "</font> ”的用户"
	end if		
	
	strGuide=strGuide & "</td><td align='right'>"
	'response.Write(sql)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "共找到 <font color=red>0</font> 个有上传文件的用户</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "共找到 <font color=red>" & totalPut & "</font> 个有上传文件的用户</td></tr></table>"
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
        	response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户")
        	else
	        	currentPage=1
           		showContent
           		response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个用户")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
	call ShowSearch()

end sub

sub showContent()
	dim i
	dim user_maxsize,vip_maxsize,admin_maxsize,umix
    i=0
	user_maxsize=oblog.setup(36,0)
	vip_maxsize=oblog.setup(40,0)
	admin_maxsize=oblog.setup(44,0)
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="admin_uploadfile_user.asp?delmore=true" onsubmit="return confirm('确定要执行选定的操作吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="96" align="center"><font color="#FFFFFF">用户</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">上传文件数</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">总计大小</font></td>
            <td width="68" align="center"><font color="#FFFFFF">剩余空间</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">分配空间</font></td>
            <td width="275" align="center"><font color="#FFFFFF">百分比</font></td>
            <td width="70" height="22" align="center"><font color="#FFFFFF"> 
              操作</font></td>
          </tr>
<%do while not rs.EOF
if rs("user_upfiles_max")=0 or rs("user_upfiles_max")="" then
	select case rs("user_level")
		case 7
		umix=user_maxsize
		case 8
		umix=vip_maxsize
		case 9
		umix=admin_maxsize
	end select
else
	umix=rs("user_upfiles_max")		
end if

%>
          <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
            <td width="96" align="center"><a href=../blog.asp?name=<%=rs("username")%> target="_blank"><%=rs("username")%></a></td>
            <td width="68" align="center"><%=rs("user_upfiles_num")%></td>
            <td align="center"><%=round(rs("user_upfiles_size")/1024)&"KB"%></td>
            <td align="center"><%=round(umix-rs("user_upfiles_size")/1024)&"KB"%></td>
            <td align="center"><%=umix&"KB"%> </td>
            <td align="center"><div align="left"><img src="images/bar.gif" width=<%=((rs("user_upfiles_size")/1024)/umix)*100&"%"%> height=10></div></td>
            <td width="70" align="center"><%		
        response.write "<a href='admin_uploadfile.asp?usermore="&rs("userid")&"'>详细</a>&nbsp;"
        response.write "<a href='admin_uploadfile_user.asp?del="&rs("userid")&"'>清空</a>"

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

sub showsearch()
%>
<form name="form2" method="post" action="admin_uploadfile_user.asp">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="tdbg">
    <td width="184">按用户查询上传文件<strong>：</strong></td>
    <td width="236">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" 查 询 ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
	</td>
    <td width="346">若为空，则查询所有用户</td>
  </tr>
</table>
</form>
<%end sub

sub delfile(uid)
	uid=clng(uid)
	dim rs,fs,oFolder,file
	set rs=oblog.execute("select user_dir from oblog_user where userid="&uid)
	if not rs.eof then
		set fs=createobject("scripting.filesystemobject")
		Set oFolder = fs.GetFolder(server.mappath(rs(0)&"/"&uid&"/upload"))		
		for each file in oFolder.Files
			fs.DeleteFile(server.mappath(rs(0)&"/"&uid&"/upload") & "/" & file.Name)
		next
		set rs=nothing
		set fs=nothing
		set oFolder=nothing
		oblog.execute("delete from [oblog_upfile] where userid="&uid)
		oblog.execute("update [oblog_user] set user_upfiles_size=0,user_upfiles_num=0 where userid="&uid)
	end if	
	response.Redirect("admin_uploadfile_user.asp")
end sub

%>
</p>
</body>
</html>
