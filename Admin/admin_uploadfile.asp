<!--#include file="inc/inc_sys.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim rs, sql
dim userid,UserSearch,Keyword,strField
dim Action,usermore,del
del=trim(request.QueryString("del"))
userid=trim(request.QueryString("userid"))
usermore=trim(request.QueryString("usermore"))
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
if usermore<>"" then
strFileName="admin_uploadfile.asp?usermore=" & Usermore
else
strFileName="admin_uploadfile.asp?UserSearch=" & UserSearch
end if
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>�ϴ��ļ�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>�ϴ��ļ�����(�ļ��б�)</strong></td>
  </tr>
  <form name="form1" action="admin_user.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>��������</strong></td>
      <td width="687" height="30"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_uploadfile.asp">�ϴ��ļ������û��б�</a> <a href="admin_uploadfile.asp">�ϴ��ļ������ļ��б�</a></td>
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
	dim strGuide
	strGuide="<table width='98%' align='center'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_uploadfile.asp'>�ϴ��ļ�����-�û��б�</a>&nbsp;&gt;&gt;&nbsp;"

	if Keyword="" then
		if usermore<>"" then
			sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid and oblog_upfile.userid="&usermore&" order by fileid desc"
			strGuide=strGuide & "�û�idΪ"&usermore&"���û��ϴ��ļ�"
		else
			sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid order by fileid desc"
			strGuide=strGuide & "ǰ500���ļ�"
		end if
	else				
		sql="select top 500 file_name,file_path,file_size,isphoto,username,fileid from [oblog_upfile],oblog_user where oblog_upfile.userid=oblog_user.userid and file_name like '%" & Keyword & "%' order by fileid  desc"
		strGuide=strGuide & "�ļ����к��С� <font color=red>" & Keyword & "</font> �����ļ�"
	end if		
	
	strGuide=strGuide & "</td><td align='right'>"
	'response.Write(sql)
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ���ϴ��ļ�</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ���ϴ��ļ�</td></tr></table>"
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
			if usermore<>"" then
			response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ϴ��ļ�")
			else
        	response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
			end if
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
			    if usermore<>"" then
			    response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ϴ��ļ�")
			    else
            	response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
				end if
        	else
	        	currentPage=1
           		showContent
				if usermore<>"" then
			    response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ϴ��ļ�")
			    else
           		response.Write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
				end if
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
	call ShowSearch()

end sub

sub showContent()
	dim i
    i=0
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="admin_uploadfile.asp?delmore=true" onsubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="96" align="center"><font color="#FFFFFF">�û�</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">�ļ���</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">�ļ�·��</font></td>
            <td width="68" align="center"><font color="#FFFFFF">�ļ���С</font></td>
            <td width="68" height="22" align="center"><font color="#FFFFFF">���</font></td>
            <td width="70" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
<%do while not rs.EOF%>
          <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
            <td width="96" align="center"><a href=../blog.asp?name=<%=rs("username")%> target="_blank"><%=rs("username")%></a></td>
            <td width="68" align="center"><%=rs("file_name")%></td>
            <td align="center"><a href="<%=blogdir & rs("file_path")%>" target="_blank"><%=rs("file_path")%></a></td>
            <td align="center"><%if rs("file_size")=0 or isnull(rs("file_size")) then response.Write("0") else response.Write round(rs("file_size")/1024)&"KB"%></td>
            <td align="center"><%if rs("isphoto")=1 then response.Write("��") else response.Write("��")%> </td>
            <td width="70" align="center"><%		
        response.write "<a href='admin_uploadfile.asp?del="&rs("fileid")&"'>ɾ��</a>"

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
<form name="form2" method="post" action="admin_uploadfile.asp">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="tdbg">
      <td width="184">���ļ�����ѯ�ϴ��ļ�<strong>��</strong></td>
    <td width="236">
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
	</td>
      <td width="346">&nbsp;</td>
  </tr>
</table>
</form>
<%end sub

sub delfile(fid)
	fid=clng(fid)
	dim rs,fs,fsize,uid,imgsrc
	set rs=oblog.execute("select file_path,file_size,userid from oblog_upfile where fileid="&fid)
	if not rs.eof then
		fsize=rs(1)
		uid=rs(2)
		imgsrc=rs(0)
		set fs=createobject("scripting.filesystemobject")
		if instr("jpg,bmp,gif,png,pcx",right(imgsrc,3))>0 then
			imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
			imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
			if  fs.FileExists(Server.MapPath(imgsrc)) then
				fs.DeleteFile Server.MapPath(imgsrc)
			end if
		end if
		if fs.FileExists(Server.MapPath(imgsrc)) then 
			fs.DeleteFile(server.mappath(imgsrc))
		end if
		set fs=nothing
		oblog.execute("delete from [oblog_upfile] where fileid="&fid)
		oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&fsize&",user_upfiles_num=user_upfiles_num-1 where userid="&uid)
		set rs=nothing
	end if	
	oblog.showok "ɾ���ɹ���",""
end sub

%>
</p>
</body>
</html>
