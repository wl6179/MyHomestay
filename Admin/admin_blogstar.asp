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
<title>�û����ǹ���</title>
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
    <td height="22" colspan=2 align=center><strong>�� �� ֮ �� �� ��</strong></td>
  </tr>
  <form name="form1" action="admin_blogstar.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ��ң�</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���500���û�����</option>
          <option value="1">ͨ����˵��û�����</option>
          <option value="2">δͨ����˵��û�����</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_blogstar.asp">�û����ǹ�����ҳ</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_blogstar.asp">
  <tr class="tdbg">
      <td width="120"><strong>�߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="blogname" selected>�û�������</option>
      <option value="UserID" >�û�����ID</option>
  
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10">
        ��Ϊ�գ����ѯ����</td>
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
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_blogstar.asp'>�û����ǹ���</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_blogstar order by id desc"
			strGuide=strGuide & "���500���û�����"
		case 1
			sql="select * from oblog_blogstar where ispass=1 order by id desc"
			strGuide=strGuide & "ͨ����˵��û�����"
		case 2
			sql="select * from oblog_blogstar where ispass=0 order by id desc"
			strGuide=strGuide & "δͨ����˵��û�����"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_blogstar order by id desc"
				strGuide=strGuide & "�����û�����"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>ID������������</li>"
					else
						sql="select * from oblog_blogstar where id =" & Clng(Keyword)
						strGuide=strGuide & "�û�����ID����<font color=red> " & Clng(Keyword) & " </font>���û�����"
					end if
				case "blogname"
					sql="select * from oblog_blogstar where blogname like '%" & Keyword & "%' order by id  desc"
					strGuide=strGuide & "�û����к��С� <font color=red>" & Keyword & "</font> �����û�����"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select
	strGuide=strGuide & "</td><td align='right'>"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ���û�����</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ���û�����</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�����")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�����")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�����")
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
  <form name="myform" method="Post" action="admin_blogstar.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> �û���</font></td>
            <td height="22" align="center"><font color="#FFFFFF">״̬</font></td>
            <td align="center"><font color="#FFFFFF">����ʱ��</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF">���</font></td>
            <td width="50" height="22" align="center"><font color="#FFFFFF"> ͼƬ</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><%=rs("id")%></td>
            <td width="80" align="center"><%="<a href='"& rs("userurl")&"' target='_blank'>"&rs("blogname")&"</a>"%> </td>
            <td align="center"> <%
			select case rs("ispass")
				case 0
					response.write "<font color=red>�ȴ����</font>"
				case 1
					response.write "ͨ�����"
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
            <td width="50" align="center"><%="<a href='"&rs("picurl")&"' target='_blank'>����쿴</a>"%></td>
            <td width="120" align="center"><%
		response.write "<a href='admin_blogstar.asp?Action=Modify&id=" & rs("id") & "'>�޸�</a>&nbsp;"
        response.write "<a href='admin_blogstar.asp?Action=Del&id=" & rs("id") & "' onClick='return confirm(""ȷ��Ҫɾ�����û�������"");'>ɾ��</a>&nbsp;"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û����ǣ�</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_blogstar.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�޸��û�������Ϣ</font></b></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��ʵ������</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="40%">���ӵ�ַ��</TD>
      <TD width="60%"> <INPUT name="userurl" value="<%=rsUser("userurl")%>" size=50   maxLength=250> <a href="<%=rsuser("userurl")%>" target="_blank">�鿴</a>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"> ͼƬ����(<strong><font color="#FF0000">�뽫��ͼƬ�ֹ���Ϊ���ʵĳߴ�</font></strong>)��</TD>
      <TD width="60%"> <INPUT name=picurl value="<%=rsUser("picurl")%>" size=50 maxLength=250><a href="<%=rsuser("picurl")%>" target="_blank">�鿴</a></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��飺</TD>
      <TD width="60%"><textarea name="bloginfo" cols="40" rows="5"><%=oblog.filt_html(rsuser("info"))%></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">״̬��</TD>
      <TD width="60%"><input type="radio" name="ispass" value=0 <%if rsUser("ispass")=0 then response.write "checked"%>>
        δͨ�����&nbsp;&nbsp; <input type="radio" name="ispass" value=1 <%if rsUser("ispass")=1 then response.write "checked"%>>
        ��ͨ�����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="id" type="hidden" id="id" value="<%=rsUser("id")%>"></TD>
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
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
	oblog.showok "�޸ĳɹ�!",""
end sub


sub DelUser()
	id=clng(id)
	oblog.execute("delete from oblog_blogstar where id="&id)
	oblog.showok "ɾ���ɹ���",""
end sub


sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

%>