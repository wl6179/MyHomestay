<!--#include file="inc/inc_sys.asp"-->
<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
dim dbpath
dim ObjInstalled
if not IsObject(conn) then link_database
if is_sqldata=0 then dbpath=server.mappath(db)
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

%>
<html>
<head>
<title>���ݿⱸ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="images/admin/Admin_STYLE.CSS">
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>�� �� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="admin_database.asp?Action=Backup">�������ݿ�</a> | <a href="admin_database.asp?Action=Restore">�ָ����ݿ�</a> 
      | <a href="admin_database.asp?Action=Compact">ѹ�����ݿ�</a> | <a href="admin_database.asp?Action=SpaceSize">ϵͳ�ռ�ռ�����</a></td>
  </tr>
</table>
<%
if Action="Backup" or Action="BackupData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowBackup()
elseif Action="Compact" or Action="CompactData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowCompact()
elseif Action="Restore" or Action="RestoreData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowRestore()
elseif Action="SpaceSize" then
	call SpaceSize()
	if isobject(conn) then conn.close:set conn=nothing
else
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>���������</li>"
	if isobject(conn) then conn.close:set conn=nothing
end if
if FoundErr=True then
	call WriteErrMsg()
end if

sub ShowBackup()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=BackupData">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
  <tr class="title"> 
      <td align="center" height="22" valign="middle"><span class="style1">�� �� �� �� ��</span></td>
  </tr>
  <tr class="tdbg"> 
    <td height="150" align="center" valign="middle">
<%    
if request("action")="BackupData" then
	call backupdata()
else
%>
        <table cellpadding="3" cellspacing="1" border="0" width="100%">
          <tr> 
            <td width="200" height="33" align="right">����Ŀ¼��</td>
            <td><input type=text size=20 name=bkfolder value=Databackup></td>
            <td>���·��Ŀ¼����Ŀ¼�����ڣ����Զ�����</td>
          </tr>
          <tr> 
            <td width="200" height="34" align="right">�������ƣ�</td>
            <td height="34"><input type=text size=20 name=bkDBname value="<%=date()%>"></td>
            <td height="34">���������ļ�����׺��Ĭ��Ϊ��.asa����������ͬ���ļ���������</td>
          </tr>
          <tr align="center"> 
            <td height="40" colspan="3"><input name="submit" type=submit value=" ��ʼ���� " <%If ObjInstalled=false Then response.Write "disabled"%>></td>
          </tr>
        </table>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
    </td>
 </tr>
</table>
</form>
<%
end sub

sub ShowCompact()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=CompactData">
<table class="border" width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="title"> 
    <td align="center" height="22" valign="middle"><span class="style1">���ݿ�����ѹ��</span></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center" height="150" valign="middle"> 
      <%    
if request("action")="CompactData" then
	call CompactData()
else
%>
      <br> <br> <br>
      ѹ��ǰ�������ȱ������ݿ⣬���ⷢ��������� <br> <br> <br> <input name="submit2" type=submit value=" ѹ�����ݿ� " <%If ObjInstalled=false Then response.Write "disabled"%>> <br> <br> 
      <%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
    </td>
  </tr>
</table>
</form>
<%
end sub

sub ShowRestore()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=RestoreData">
	<table width="98%" class="border" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr class="title"> 
      		<td align="center" height="22" valign="middle"><span class="style1">���ݿ�ָ�</span></td>
        </tr>
        <tr class="tdbg">
            <td align="center" height="150" valign="middle"> 
        <%
if request("action")="RestoreData" then
	call RestoreData()
else
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="200" height="30" align="right">�������ݿ�·������ԣ���</td>
            <td height="30"><input name=backpath type=text id="backpath" value="DataBackup\oblog.mdb" size=50 maxlength="200"></td>
          </tr>
          <tr align="center"> 
            <td height="40" colspan="2"><input name="submit" type=submit value=" �ָ����� " <%If ObjInstalled=false Then response.Write "disabled"%>></td>
          </tr>
        </table>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
	end if
end if
%>
            </td>
        </tr>
  </table>
</form>
<%
end sub


sub SpaceSize()
	on error resume next
	dim rs
	set rs=oblog.execute("select userdir from oblog_userdir")
%>
<br>
<table class="border" width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr class="title"> 
		<td align="center" height="22" valign="middle"><span class="style1">ϵͳ�ռ�ռ�����</span></td>
	</tr>
  <tr class="tdbg"> 
    <td width="100%" height="150" valign="middle">
	<blockquote><br>
        ϵͳ����ռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpaceinfo("data")%>
      <br>
      <br>
        ��������ռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpaceinfo("databackup")%>
      <br>
      <br>
        �����ļ�ռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpecialSpaceinfo("Program")%>
      <br>
      <br>
        ��ɫģ��ռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpaceinfo("skin")%>
      <br>
      <br>
        ϵͳͼƬռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpaceinfo("images")%>
	  <%while not rs.eof %>
      <br>
      <br>
      <%=rs(0)%>Ŀ¼ռ�ÿռ䣺&nbsp;&nbsp; 
        <%showSpaceinfo(rs(0))%>
	  <%rs.movenext
	  wend%>
      <br>
      <br>
        ϵͳռ�ÿռ��ܼƣ� 
        <%showspecialspaceinfo("All")%>
	</blockquote> 	
    </td>
  </tr>
</table>
<br>
<%
end sub
sub sqldata_readme
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="98%" class="border">			<tr>
  					
    <th height=25> &nbsp;<font color="#FFFFFF">&nbsp;SQL���ݿ����ݴ���˵��</font> </th>
  				</tr> 	
 				<tr class="tdbg">
 					<td > 			
 			<blockquote>
<B>һ���������ݿ�</B>
<BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
3��ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ�����ӣ����ԭ��û��·����������ֱ��ѡ����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б���
<BR><BR>
<B>������ԭ���ݿ�</B><BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
3������½��õ����ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->�����-->Ȼ��ѡ����ı����ļ���-->��Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ�����������µ�ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����bbs_data.mdf�����ڵ����ݿ���forum���͸ĳ�forum_data.mdf�������º������ļ���Ҫ���������ķ�ʽ����صĸĶ������µ��ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\bbs_data.mdf����d:\sqldata\bbs_log.ldf��������ָ�������<BR>
6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ�������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR><BR>

<B>�����������ݿ�</B><BR><BR>
һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ�������������´�С��Ӧ�����ڽ��д˲����������ݿ����¹���<BR>
1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
3��<font color=blue>�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ������һЩ�쳣����������ǻָ����ݿ����Ҫ����</font>
<BR><BR>

<B>�ġ��趨ÿ���Զ��������ݿ�</B><BR><BR>
<font color=red>ǿ�ҽ������������û����д˲�����</font><BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
6����һ��ָ���������±��ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı���һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ�����
<BR><BR>
�޸ļƻ���<BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������
<BR><BR>
<B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR><BR>
һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ�������⣬������ѯ���������Ա���߲�ѯ��������<BR>
1����ԭ���ݿ�����б��洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%end sub%>
</body>
</html>
<%
sub BackupData()
	dim bkfolder,bkdbname,fso
	bkfolder=trim(request("bkfolder"))
	bkdbname=trim(request("bkdbname"))
	if bkfolder="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ������Ŀ¼��</li>"
	end if
	if bkdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ�������ļ���</li>"
	end if
	if FoundErr=True then exit sub
	bkfolder=server.MapPath(bkfolder)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.FileExists(dbpath) then
		If fso.FolderExists(bkfolder)=false Then
			fso.CreateFolder(bkfolder)
		end if
		fso.copyfile dbpath,bkfolder & "\" & bkdbname & ".asa"
		response.write "<center>�������ݿ�ɹ������ݵ����ݿ�Ϊ " & bkfolder & "\" & bkdbname & ".asa</center>"
	Else
		response.write "<center>�Ҳ���Դ���ݿ��ļ�������inc/conn.asp�е����á�</center>"
	End if
end sub

sub CompactData()
	Dim fso, Engine, strDBPath
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(dbPath) Then
		Set Engine = CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath," Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
		fso.CopyFile strDBPath & "temp.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		Set fso = nothing
		Set Engine = nothing
		response.write "���ݿ�ѹ���ɹ�!"
	Else
		response.write "���ݿ�û���ҵ�!"
	End If
end sub

sub RestoreData()
	dim backpath,fso
	backpath=request.form("backpath")
	if backpath="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ָ��ԭ���ݵ����ݿ��ļ�����<li>"
		exit sub	
	end if
	backpath=server.mappath(backpath)
	Set Fso=server.createobject("scripting.filesystemobject")
	If Right(LCase(backpath),4)=".asa" OR Right(LCase(backpath),4)=".asp" Or Right(LCase(backpath),4)=".mdb" Then		
		if fso.fileexists(backpath) then  					
			fso.copyfile Backpath,Dbpath
			response.write "�ɹ��ָ����ݣ�"
		else
			response.write "�Ҳ���ָ���ı����ļ���"	
		end if
	Else
		response.write "����ı����ļ���"	
	End If
	Set Fso=Nothing
end sub

Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	

Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("pic")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath) 		
	
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if	
	
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	drvpath=server.mappath(drvpath) 		
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	

Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next	
	
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function 	

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>