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
		blog.progress int(j/p*100),"���IDΪ"&rs("fileid")&"���ļ�..."
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
		blog.progress int(j/p*100),"����IDΪ"&rs("userid")&"���û�..."
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
	response.Write("�ϴ��ļ���ȫ��������ɣ�")	
else
	%>
	
<title>�����û��ϴ��ļ�</title>
<form name="form1" method="post" action="">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center" class="title">
      <td height="22" colspan="2"><strong><font color="#FFFFFF">�����û��ϴ��ļ�</font></strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="2"><p>˵���� �����������������û��ϴ��ļ���ɾ�����ݿ��е���Ч��¼�����¼����û��ϴ��ֽڣ�ִ��ʱ��ϳ��������ز�����</p>
        <p>ԭ2.52�汾��ʹ���ϴ�Ŀ¼�� 
          <input name="old" type="text" id="old" value="uploadfile" size="50" maxlength="200">
          <br>
          (��ʹ�ù����Ŀ¼����ʹ��,�Ÿ�������ʹ�ù�uploadfile��uploadfile1����Ŀ¼������д��uploadfile,uploadfile1������2.52�汾�����������û���һ��ִ��һ��Ҫ������д���Ժ�������Բ���д��) <br>
        </p>
        </td>
  </tr>
  <tr class="tdbg">
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="ִ��">
    <input name="Action" type="hidden" id="Action" value="DoUpdate"></td>
  </tr>
</table>
   </form>
<%
end if
%>

