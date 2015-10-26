<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<%
dim action,old
action=request("action")
old=trim(request("old"))
if action="DoUpdate" then
	Server.ScriptTimeOut=999999999
	dim fso,i,s,fileinfo,blog,p,j,q,rstmp,rs
	set fso=CreateObject("Scripting.FileSystemObject")
	set rs=server.CreateObject("adodb.recordset")
	set blog=new class_blog
	j=1
	blog.progress_init
	rs.open "select * from oblog_upfile",conn,1,3
	p=rs.recordcount
	while not rs.eof 
		blog.progress int(j/p*100),"检查ID为"&rs("fileid")&"的文件..."
		if fso.FileExists(Server.MapPath(rs("file_path"))) then
			Set fileinfo = fso.GetFile(Server.MapPath(rs("file_path")))
			rs("file_size")=fileinfo.size
			rs("addtime")=fileinfo.DateLastModified
			rs.update
		else
			if instr(old,",")>0 then
				q=0
				s=split(old,",")
				for i=0 to ubound(s)
					if trim(s(i))<>"" then
						if fso.FileExists(Server.MapPath(trim(s(i))&"/"&rs("file_name"))) then
							Set fileinfo = fso.GetFile(Server.MapPath(trim(s(i))&"/"&rs("file_name")))
							rs("file_size")=fileinfo.size
							rs("addtime")=fileinfo.DateLastModified
							rs.update
							q=1
						end if					
					end if
				next
				if q<>1 then
				rs.delete
				rs.update
				end if
			else
				if fso.FileExists(Server.MapPath(old&"/"&rs("file_name"))) then
					Set fileinfo = fso.GetFile(Server.MapPath(old&"/"&rs("file_name")))
					rs("file_size")=fileinfo.size
					rs("addtime")=fileinfo.DateLastModified
					rs.update
				else
					rs.delete
					rs.update
				end if
			end if
		end if
		rs.movenext
		j=j+1
	wend
	j=1
	rs.close
	rs.open "select userid,user_upfiles_size from oblog_user",conn,1,3
	p=rs.recordcount
	while not rs.eof
		blog.progress int(j/p*100),"更新ID为"&rs("userid")&"的用户..."
		set rstmp=oblog.execute("select sum(file_size) from oblog_upfile where userid="&rs(0))
		if isnull(rstmp(0)) then
			rs(1)=0
		else
			rs(1)=rstmp(0)
		end if
		rs.update
		rs.movenext
		j=j+1
	wend
	rs.close
	'oblog.execute("update oblog_user innner join oblog_upfile on  oblog_user.userid=oblog_upfile.userid set oblog_user.user_upfiles_size=oblog_upfile.file_size")
	set rs=nothing
	set rstmp=nothing
	set fso=nothing
	set blog=nothing
	response.Write("上传文件表全部整理完成！")	
else
	%>
	
<title>清理用户上传文件</title>
<form name="form1" method="post" action="">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center" class="title">
      <td height="22" colspan="2"><strong><font color="#FFFFFF">清理用户上传文件</font></strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><p>说明： 本操作将重新整理用户上传文件表，删除数据库中的无效纪录，重新计算用户上传字节，执行时间较长，请慎重操作。</p>
        <p>原2.52版本曾使用上传目录： 
          <input name="old" type="text" id="old" value="uploadfile" size="50" maxlength="200">
          <br>
          (如使用过多个目录，请使用,号隔开，如使用过uploadfile和uploadfile1两个目录，则填写“uploadfile,uploadfile1”，从2.52版本升级过来的用户第一次执行一定要认真填写，以后整理可以不填写。) <br>
        </p>
        </td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="执行">
    <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
  </tr>
</table>
   </form>
<%
end if
%>

