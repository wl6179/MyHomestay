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
<title>ע��</title>
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
    <td height="22" colspan=2 align=center><strong>ע �� ��Ŀ�ռ� �� ��</strong></td>
  </tr>
  <form name="form1" action="admin_UserLogclass_Nav.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>��Ŀ�ռ����ã�</strong></td>
      <td width="687" height="30"><select style="display:none;" size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���ע���500����Ŀ�ռ�</option>
          <option value="1">�������TOP100</option>
          <option value="2">�������ٵ�100����Ŀ�ռ�</option>
		  <option value="3">VIP��Ŀ�ռ�</option>
		  <option value="4">ǰ̨����Ա</option>
		  <option value="5">�Ƽ���Ŀ�ռ�</option>
          <option value="6">�ȴ�����Ա��֤����Ŀ�ռ�</option>
          <option value="7">���б���ס����Ŀ�ռ�</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_UserLogclass_Nav.asp">��Ŀ�ռ������ҳ</a>&nbsp;|&nbsp;<a href="admin_UserLogclass_Nav.asp?Action=Add">�������Ŀ�ռ�</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_UserLogclass_Nav.asp">
  <tr class="tdbg">
    <td width="120"><strong>��Ŀ�ռ��ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	  <option value="UserName" selected>��Ŀ�ռ���</option>
      <option value="UserID" >��Ŀ�ռ�ID</option>
	 <!-- <option value="nickname" >��Ŀ�ռ��ǳ�</option>-->
      <option value="ip" >����¼ip</option>
	  <option value="blogname" >��������</option>
	  <option value="lastlogintime" >��������δ��¼</option>
	  <option value="logcount" >��Ŀ������С��</option>
	  <option value="logintimes" >��Ŀ��¼����С��</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  ��Ϊ�գ����ѯ������Ŀ�ռ�</td>
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
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_UserLogclass_Nav.asp'>ע����Ŀ�ռ����</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_user where user_level>6 order by UserID desc"
			strGuide=strGuide & "���ע���500����Ŀ�ռ�"
		case 1
			sql="select top 100 * from oblog_user where user_level>6 order by log_count desc"
			strGuide=strGuide & "������������ǰ100����Ŀ�ռ�"
		case 2
			sql="select top 100 * from  oblog_user where user_level>6 order by log_count"
			strGuide=strGuide & "�����������ٵ�100����Ŀ�ռ�"
		case 3
			sql="select * from  oblog_user where user_level=8 order by userid desc"
			strGuide=strGuide & "����VIP��Ŀ�ռ�"
		case 4
			sql="select * from  oblog_user where user_level=9 order by userid desc"
			strGuide=strGuide & "ǰ̨������"
		case 5
			sql="select * from  oblog_user where user_isbest=1 order by userid desc"
			strGuide=strGuide & "�Ƽ���Ŀ�ռ�"
		case 6
			sql="select * from oblog_user where User_Level=6 order by userID  desc"
			strGuide=strGuide & "�ȴ�������֤֤����Ŀ�ռ�"
		case 7
			sql="select * from oblog_user where  LockUser =1 order by userID  desc"
			strGuide=strGuide & "���б���ס����Ŀ�ռ�"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_user where user_level>6 order by userID desc"
				strGuide=strGuide & "������Ŀ�ռ�"
			else
				select case strField
				case "UserID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>��Ŀ�ռ�ID������������</li>"
					else
						sql="select * from oblog_user where user_level>6 and userID =" & Clng(Keyword)
						strGuide=strGuide & "��Ŀ�ռ�ID����<font color=red> " & Clng(Keyword) & " </font>����Ŀ�ռ�"
					end if
				case "UserName"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and username like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��Ŀ�ռ����к��С� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					else
						sql="select * from oblog_user where user_level>6 and username= '" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��Ŀ�ռ������ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					end if
					
				case "nickname"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and nickname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��Ŀ�ռ��ǳ��к��С� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					else
						sql="select * from oblog_user where user_level>6 and nickname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��Ŀ�ռ��ǳƵ��ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					end if
				case "ip"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and lastloginip like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "����¼ip�к��С� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					else
						sql="select * from oblog_user where user_level>6 and lastloginip='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "����¼ip���ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					end if
				case "blogname"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level>6 and blogname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��Ŀ���к��С� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					else
						sql="select * from oblog_user where user_level>6 and blogname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��Ŀ�����ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
					end if
				case "logcount"
					sql="select top 500 * from oblog_user where user_level>6 and log_count < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "������С�ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
				case "logintimes"
					sql="select top 500 * from oblog_user where user_level>6 and logintimes < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "��¼����С�ڡ� <font color=red>" & Keyword & "</font> ������Ŀ�ռ�"
				case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="select top 500 * from oblog_user where user_level>6 and datediff(d,lastlogintime,getdate())>"&clng(Keyword)&" order by userID  desc"
				   Else 
				   	sql="select top 500 * from oblog_user where user_level>6 and datediff('d',lastlogintime,now())>"&clng(Keyword)&" order by userID  desc"
				   End If 
					strGuide=strGuide & "�� <font color=red>" & Keyword & "</font> ������δ��¼����Ŀ�ռ�"
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
'	Response.Write sql
'	Response.End 
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ����Ŀ�ռ�</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ����Ŀ�ռ�</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����Ŀ�ռ�")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����Ŀ�ռ�")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����Ŀ�ռ�")
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
  <form name="myform" method="Post" action="admin_UserLogclass_Nav.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> ��Ŀ�ռ���</font></td>
            <td height="22" align="center"><font color="#FFFFFF">������Ŀ�ռ���</font></td>
            <td height="22" align="center"><font color="#FFFFFF">����¼IP</font></td>
            <td align="center"><font color="#FFFFFF">����¼ʱ��</font></td>
            <td width="60" height="22" align="center"><font color="#FFFFFF">��¼����</font></td>
            <td width="40" height="22" align="center"><font color="#FFFFFF"> ״̬</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='UserID' type='checkbox' onClick="unselectall()" id="UserID" value='<%=cstr(rs("userID"))%>'></td>
            <td width="30" align="center"><%=rs("userID")%></td>
            <td width="80" align="center"><%
			if Cint(rs("logclassid_Nav"))>0 then
				response.write "<font style='cursor:hand;' title='�Ѱ���Ŀ�ռ�' color=red>��&nbsp;</font>"
			end if
			response.write "<a href='../blog.asp?name=" & rs("userName") & "'" 			
			response.write """ target='_blank'>" & rs("userName") & "</a>"
			%> </td>
            <td align="center"> <%
			select case rs("User_Level")
				case 6
					response.write "<font color=green>ע��ļ�ͥ�û�</font>"
				case 7
					response.write "ע����Ŀ����"
				case 8
					response.write "<font color=blue>VIP��Ŀ����</font>"
				case 9
					response.write "<font color=blue>ǰ̨������</font>"
				case else
					response.write "<font color=red>�쳣ע��</font>"
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
	  	response.write "<font color=red>������</font>"
	  else
	  	response.write "����"
	  end if
	  %></td>
            <td width="120" align="center"><%
		response.write "<!--<a href='admin_UserLogclass_Nav.asp?Action=Modify&UserID=" & rs("userID") & "'>�޸�</a>-->&nbsp;"
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>������Ŀ��������</a><br />&nbsp;"
		response.write "<a href='admin_UserLogclass_Nav.asp?Action=Sheaf&userid=" & rs("userid") & "'><font color='#038ad7'>�󶨷���</font></a>&nbsp;|"
		if rs("LockUser")=0 then
			response.write "<a href='admin_UserLogclass_Nav.asp?Action=Lock&UserID=" & rs("userID") & "'>����</a>&nbsp;"
		else
            response.write "<a href='admin_UserLogclass_Nav.asp?Action=UnLock&UserID=" & rs("userID") & "'>����</a>&nbsp;"
		end if
        response.write "<a href='admin_UserLogclass_Nav.asp?Action=Del&UserID=" & rs("userID") & "' onClick='return confirm(""ȷ��Ҫɾ������Ŀ�ռ���"");'>ɾ��</a>&nbsp;"
		%>
		<% If rs("logclassid_Nav")>0 Then %>
		 | <a href="admin_UserLogclass_Nav.asp?Action=ClearBinding&id=<%=rs("userid")%>" onClick="return confirm('ȷ��ǿ�������״̬��');"><img src="/images/del.gif" title="ǿ��ɾ�������·���İ�״̬����" /></a>
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
              ѡ�б�ҳ��ʾ��������Ŀ�ռ�</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=true">
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">�ƶ���
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">ע��ļ�ͥ�û�</option>
                <option value="7">ע����Ŀ����</option>
                <option value="8">VIP��Ŀ����</option>
                <option value="9">ǰ̨������</option>
              </select>
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" ִ �� "> </td>
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
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�������Ŀ�ռ���Ϣ</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><font color="red">*</font>&nbsp;�½���Ŀ�������ļ����������� about_us ����</TD>
      <TD width="60%"><input name="username" id="username" type="text" value="" size=15 maxlength=30 />&nbsp;&nbsp;<a href="javascript:checkssn('../checkssn_Nav.asp')";>�鿴����Ŀ���Ƿ����</a></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''" <% if trim(oblog.setup(4,0))="" or oblog.setup(12,0)<>1 then response.Write("style='display:none;'")%>>
      <td>��Ŀ�ռ�������</td>
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
      <td><font color="red">*</font>&nbsp;��Ŀ����������������·����������һ�£�����</td>
      <td><input name=blogname   type=text id="blogname" value="" size=30 maxlength=20></td>
    </tr>
	<tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td><font color="blue">Ӣ����</font>����"About Us"�����Բ����</td>
      <td><input name=EnglishName   type=text id="EnglishName" value="" size=30 maxlength=20></td>
    </tr>
	
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ŀ�ռ�󶨵Ķ���������</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="" size=30 maxlength=20></td>
    </tr>
	<%end if%>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>blog���</td>
      <td><select name="usertype" id="usertype">
          <option value="0" selected="selected">0</option>
        </select></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="40%"><font color="red">*</font>&nbsp;����(����6λ)��<BR>
        ���������룬���ִ�Сд�� �벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ� </TD>
      <TD width="60%"> <input name=password type=password id=password size=15 maxlength=30> <font color="#FF0000">�Դ���Ŀ��һ����������</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD><font color="red">*</font>&nbsp;ȷ������(����6λ)��<br>
        ������һ��ȷ��</TD>
      <TD><input name=repassword type=password id=repassword size=15 maxlength=30> <font color="#FF0000">�Դ���Ŀ��һ����������</font> </TD>
    </TR>
    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">�������⣺<br>
        �����������ʾ����</TD>
      <TD width="60%"> <input value="MyHomestay" name=question type=text id=question size=30 maxlength=30> 
      </TD>
    </TR>
    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">����𰸣�<BR>
        �����������ʾ����𰸣�����ȡ������</TD>
      <TD width="60%"> <input value="MyHomestay" name=an type=text id=an size=30 maxlength=30> <font color="#FF0000"></font></TD>
    </TR>

    <TR class="tdbg" style="display:none;"> 
      <TD width="40%">Email��ַ��</TD>
      <TD width="60%"> <input value="service@myhomestay.com.cn" name=email type=text size=15 maxlength=30> 
        
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">ȷ��ͬ�⣺</TD>
      <TD width="60%"> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked>ͬ�⡡<input class='input_radio' type=radio name=passregtext id=passregtext value='0'>��ͬ��</TD>
    </TR>

    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> <input name=Submit   type=submit id="Submit" value="�������Ŀ�ռ�"> </TD>
    </TR>
  </TABLE>
</form>
<%

end sub




sub Sheaf()'�󶨷���&��Ŀ�ռ�
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ�ռ䣡</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">��Ŀ�ռ䡪�󶨡�>���·���</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ŀ�ռ�����</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
   
   
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�������·��ࣺ</td>
      <td><select name="classid" id="classid" <%If rsUser("logclassid_Nav")>0 Then Response.Write("disabled='disabled'")%>>
	<%=oblog.show_class2("log",Cint(rsUser("logclassid_Nav")),0)%> <%''"log",classid,t=1ͼ%>
	</select></td>
    </tr>
   
   
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveSheaf"> <input name=Submit   type=submit id="Submit" value="�����"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("userID")%>"></TD>
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	
	if classid=0 then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�󶨵����·��࣡</li>"
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
			ErrMsg=ErrMsg & "<br><li>�Ҳ�����ѡ�����·��࣡</li>"
			rs_temp.close
			set rs_temp=nothing
			exit sub
		else'����������⣬���̷������¼�Ƿ�Ϊ �����ࡪ��һ�����࣡
			If rs_temp(0)>0 Then'��� ����������(һ������)������
				founderr=true
				ErrMsg=ErrMsg & "<br><li>����ѡ�� һ�����·��࣡</li>"
				exit sub
			elseIf rs_temp(1)>0 Then'����Ѿ��а��ˣ�����
				founderr=true
				ErrMsg=ErrMsg & "<br><li>�����·����ѱ��󶨣���ѡ���������·��࣡</li>"
				exit sub
			End If
		end if
				
	else'����������û�����⣬��ͨ������������ֵ���ͣ�
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ�ռ䣡</li>"
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
	
	oblog.showok "��Ŀ�ռ�����ɹ�!",""
end sub



sub Modify()
	dim UserID
	dim rsUser,sqlUser
	UserID=trim(request("UserID"))
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ�ռ䣡</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�޸�ע����Ŀ�ռ���Ϣ</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ŀ�ռ�����</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>��Ŀ�ռ�������</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <select name="user_domainroot" ><%=oblog.type_domainroot(rsuser("user_domainroot"))%></select></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ŀ����</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��Ŀ�ռ�󶨵Ķ���������</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="<%=rsuser("custom_domain")%>" size=30 maxlength=20></td>
    </tr>
	<%end if%>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>blog���</td>
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
      <TD width="40%">����(����6λ)��<BR>
        ���������룬���ִ�Сд�� �벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ� </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������(������Ŀ�ռ��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD>ȷ������(����6λ)��<br>
        ������һ��ȷ��</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12> <font color="#FF0000">��������޸ģ�������(������Ŀ�ռ��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�������⣺<br>
        �����������ʾ����</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30>(������Ŀ�ռ��뵽��̳�޸�) 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">����𰸣�<BR>
        �����������ʾ����𰸣�����ȡ������</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������(������Ŀ�ռ��뵽��̳�޸�)</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�Ա�</TD>
      <TD width="60%"> <INPUT type=radio value="1" name=sex <%if rsUser("Sex")=1 then response.write "CHECKED"%>>
        �� &nbsp;&nbsp; <INPUT type=radio value="0" name=sex <%if rsUser("Sex")=0 then response.write "CHECKED"%>>
        Ů</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">Email��ַ��</TD>
      <TD width="60%"> <INPUT name=Email value="<%=rsUser("userEmail")%>" size=30   maxLength=50> 
        <a href="mailto:<%=rsUser("userEmail")%>">������Ŀ�ռ䷢һ������ʼ�</a> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">OICQ���룺</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">MSN��</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsUser("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ŀ�ռ伶��</TD>
      <TD width="60%"><select name="User_Level" id="User_Level">
          <option value="6" <%if clng(rsUser("user_level"))=6 then response.Write("selected")%>>ע��ļ�ͥ�û�</option>
          <option value="7" <%if clng(rsUser("user_level"))=7 then response.Write("selected")%>>ע����Ŀ����</option>
          <option value="8" <%if clng(rsUser("user_level"))=8 then response.Write("selected")%>>vip��Ŀ����</option>
          <option value="9" <%if clng(rsUser("user_level"))=9 then response.Write("selected")%>>ǰ̨������</option>
        </select></TD>
    </TR>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ռ�(kb)��</td>
      <td><input name=user_upfiles_max   type=text value="<%=rsuser("user_upfiles_max")%>" size=20 maxlength=20>
        Ϊ��ʱΪϵͳĬ�����ã����赥�����ã�������kbֵ</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ֽ�(�ֽ�)��</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=20 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD>�Ƿ�Ϊ�Ƽ���Ŀ�ռ䣺</TD>
      <TD><input type="radio" name="isbest" value=1 <%if rsUser("user_isbest")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp; <input type="radio" name="isbest" value=0 <%if rsUser("user_isbest")<>1 then response.write "checked"%>>
        ��</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ŀ�ռ�Ŀ¼��</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsUser("user_dir")%>" size=30 maxLength=50>
        ���ޱ�Ҫ�벻Ҫ�޸ģ����������Ŀ�ռ�Ŀ¼����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">��Ŀ�ռ�״̬��</TD>
      <TD width="60%"><input type="radio" name="LockUser" value=0 <%if rsUser("LockUser")=0 then response.write "checked"%>>
        ����&nbsp;&nbsp; <input type="radio" name="LockUser" value=1 <%if rsUser("LockUser")=1 then response.write "checked"%>>
        ����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("userID")%>"></TD>
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
      <td height="22" colspan="2"><strong><font color="#FFFFFF">������Ŀ�ռ侲̬ҳ��</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          1��������������������Ŀ�ռ侲̬ҳ�档<br>
          2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
      3 ��������������Ŀ�ռ�����¡� </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��ʼ��Ŀ�ռ�ID��</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      ��Ŀ�ռ�ID��������д�������һ��ID�ſ�ʼ���и���</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">������Ŀ�ռ�ID��</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      �����¿�ʼ������ID֮�����Ŀ�ռ����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="���ɾ�̬ҳ��"></td>
  </tr>
</table>
</form>

<FORM name="Form1" action="admin_UserLogclass_Nav.asp?action=DoUpdatelog" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">�������¾�̬ҳ��</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          1��������������������Ŀ�ռ侲̬ҳ�档<br>
          2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
      3���������������£����¡�</p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��ʼ����ID��</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      ��Ŀ�ռ�ID��������д�������һ��ID�ſ�ʼ���и���</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��������ID��</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      �����¿�ʼ������ID֮�������ҳ�棬֮�����ֵ��ò�Ҫѡ�����</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="�������¾�̬ҳ��"></td>
  </tr>
</table>
</form>
  <%
end sub

sub gouser1()
if is_ot_user=1 then
	response.Write("�����ⲿ���ݿ���Ŀ�ռ䣬����ʹ�ô˹���"):response.End()
end if
%>
<FORM name="Form1" action="admin_UserLogclass_Nav.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼����Ŀ�ռ�����̨</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          ������������Ա��¼����Ŀ�ռ�Ĺ��������й���<br>
          ����Ŀ�ռ���������ϰ�ʱ���ɽ������Ŀ�ռ��̨��Э����Ŀ�ռ���в�����<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��Ŀ�ռ��˺ţ�</td>
    <td height="25"><input name="username" type="text" id="username" value="" size="30" maxlength="50"></td>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value=" �ύ "></td>
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
		response.Write("�����ⲿ���ݿ���Ŀ�ռ䣬����ʹ�ô˹���"):response.End()
	end if
	dim rs,username
	username=oblog.filt_badstr(trim(request("username")))
	if username="" then response.Write("��Ŀ�ռ�������Ϊ��"):response.End()
	if is_ot_user=0 then
		set rs=oblog.execute("select username,password from oblog_user where username='"&username&"'")
	else
		set rs=ot_conn.execute("select "&ot_username&","&ot_password&" from "&ot_usertable&" where "&ot_username&"='"&username&"'")
	end if
	if not rs.eof then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		set rs=nothing
		response.Write(oblog.logined_uid)
		If Request("username")="myhomestay" Then '�������Ϣ����Ա�ʺ�myhomestay����ô��ֱ�ӽ�����Ϣ����ҳ��user_pmmanage.asp
			response.Redirect("../user_pmmanage.asp")
		Else
			response.Redirect("../HomestayBackctrl-index.asp")
		End If
	else
		set rs=nothing
		response.Write("�޴���Ŀ�ռ�"):response.End()
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
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
				errmsg=errmsg & "<br><li>���벻�ܴ���12С��6������㲻���޸����룬�뱣��Ϊ�ա�</li>"
			end if
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>�����к��зǷ��ַ�������㲻���޸����룬�뱣��Ϊ�ա�</li>"
				founderr=true
			end if
		end if
		if Password<>PwdConfirm then
			founderr=true
			errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
		end if
		if Question="" then
			founderr=true
			errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
		end if
	end if
	if Sex="" then
		founderr=true
		errmsg=errmsg & "<br><li>�Ա���Ϊ��</li>"
	else
		sex=cint(sex)
		if Sex<>0 and Sex<>1 then
			Sex=1
		end if
	end if
	if is_ot_user=0 then
		if Email="" then
			founderr=true
			errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
		else
			if oblog.IsValidEmail(Email)=false then
				errmsg=errmsg & "<br><li>����Email�д���</li>"
				founderr=true
			end if
		end if
	end if
	if OICQ<>"" then
		if not isnumeric(OICQ) or len(cstr(OICQ))>10 then
			errmsg=errmsg & "<br><li>OICQ����ֻ����4-10λ���֣�������ѡ�����롣</li>"
			founderr=true
		end if
	end if
	if MSN<>"" then
		if oblog.IsValidEmail(MSN)=false then
			errmsg=errmsg & "<br><li>���MSN����</li>"
			founderr=true
		end if
	end if

	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ����Ŀ�ռ伶��</li>"
	else
		User_Level=CLng(User_Level)
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		set rsuser=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"' and userid<>"&userid)
		if not rsuser.eof then 
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>�������Ѿ���������ʹ�ã�</li>"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ�ռ䣡</li>"
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
	oblog.showok "�޸ĳɹ�!",""
end sub


sub DelUser()
	dim rs,i
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ������Ŀ�ռ�</li>"
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
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ��������Ŀ�ռ�</li>"
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
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=����Ŀ�ռ��Ѿ�������������ϵ����Ա!');"&vbcrlf
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
	oblog.showok "������Ŀ�ռ�ɹ�",""
end sub

sub UnLockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ��������Ŀ�ռ�</li>"
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
	oblog.showok "������Ŀ�ռ�ɹ�",""
end sub

sub MoveUser()
	dim msg
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ƶ�����Ŀ�ռ�</li>"
		exit sub
	end if
	dim User_Level
	User_Level=trim(request("User_Level"))
	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ����Ŀ�ռ���</li>"
		exit sub
	else
		User_Level=Clng(User_Level)
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ע��ļ�ͥ�û�</font>����"
			sql="Update oblog_user set User_Level=6 where userID in (" & UserID & ")"
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update oblog_user set User_Level=7 where userID in (" & UserID & ")"
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update oblog_user set User_Level=8 where userID in (" & UserID & ")"
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update oblog_user set User_Level=9 where userID in (" & UserID & ")"
		end if
	else
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ע��ļ�ͥ�û�</font>����"
			sql="Update oblog_user set User_Level=6 where userID=" & UserID 
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update oblog_user set User_Level=7 where userID="& UserID
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update oblog_user set User_Level=8 where userID=" & UserID 
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ����Ŀ�ռ���Ϊ��<font color=blue>ǰ̨������</font>����"
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
		ErrMsg=ErrMsg & "<br><li>��ָ����ʼID</li>"
	else
		BeginID=Clng(BeginID)
	end if
	if EndID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ������ID</li>"
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
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""ȫ�����ȣ���ǰ��Ŀ�ռ�ID:"&rsuser(0)&""";</script>" & VbCrLf
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
		ErrMsg=ErrMsg & "<br><li>��ָ����ʼID</li>"
	else
		BeginID=Clng(BeginID)
	end if
	if EndID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ������ID</li>"
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
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""���ȣ���ǰ����ID:"&rs(0)&""";</script>" & VbCrLf
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






sub sub_chkreg()
	if oblog.ChkPost()=false then
		oblog.adderrstr("ϵͳ��������ⲿ�ύ��")
		oblog.showerr
		exit sub
	end if

	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	if oblog.chkiplock() then
		oblog.adderrstr("�Բ������IP�ѱ�����,������ע�ᣡ")
		oblog.showerr
		exit sub
	end if
	regusername=oblog.filt_badstr(trim(request("username")))
	if regusername="" then
		response.write "��������Ŀ��"
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
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("��Ŀ�ռ�������Ϊ��(���ܴ���14С��4)��")
	if oblog.chk_regname(regusername) then oblog.adderrstr("��Ŀ�ռ���ϵͳ������ע�ᣡ")
	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("��Ŀ�ռ����к���ϵͳ��������ַ���")
	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("��Ŀ�ռ���������ȫ��Ϊ���֣�")
	if oblog.chkdomain(regusername)=false then oblog.adderrstr("��Ŀ�ռ������Ϲ淶��ֻ��ʹ��Сд��ĸ�����ּ��»��ߣ�")
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("��������Ϊ��(���ܴ���14���ַ�)��")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("��������С��4���ַ���")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("������ϵͳ������ע�ᣡ")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
		if user_domainroot="" then oblog.adderrstr("����������Ϊ�գ�")
	end if
	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("���벻��Ϊ��(���ܴ���14С��4)��")
	if re_regpassword="" then oblog.adderrstr("�ظ����벻��Ϊ�գ�")
	if regpassword<>re_regpassword then oblog.adderrstr("�����������벻ͬ��")
	if question="" or oblog.strLength(question)>50 then oblog.adderrstr("�һ�������ʾ���ⲻ��Ϊ��(���ܴ���50)��")
	
	if answer="" or oblog.strLength(answer)>50 then
		oblog.adderrstr("�һ���������𰸲���Ϊ��(���ܴ���50)��")
	elseif answer="MyHomestay" then
		answer="49ba59abbe56e057"
	end if
	
	'if oblog.chk_regname(nickname) then oblog.adderrstr("���ǳ�ϵͳ������ע�ᣡ")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("�ǳ��к���ϵͳ��������ַ���")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("�ǳƲ��ܲ��ܴ���50�ַ���")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("��Ŀ����������Ϊ��(���ܴ���50�ַ�)��")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("��Ŀ�������к���ϵͳ��������ַ���")	
	if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"��")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("��Ŀ�ռ����к��зǷ��ַ���")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������ǳƴ��ڣ�������ǳƣ�")		
	'end if
	if user_domain<>"" then
		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������������ڣ������������")
	end if
	
	if oblog.errstr<>"" then response.Write("<ul><li>����"& oblog.errstr &"<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul>"):exit sub
	
	

	
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
		show_reg="�����ע��<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>����Ŀ"&blogname&"ע��ɹ���</strong></li></ul>"
			show_reg=show_reg&"5����Զ�ת���ء�"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""history.go(-2)"",5000);"
			show_reg=show_reg&"</script>"
		else
			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
			show_reg=show_reg&"<ul><li><strong>��ϲ�����Ѿ�ע��ɹ���</strong></li>"
			show_reg=show_reg&"<li><a href='index.asp'>������ҳ</a></li>"
			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>�����������̨ѡ������ϲ����ҳ����(5����Զ���������̨)</a></li></ul>"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
			show_reg=show_reg&"</script>"
		end if
	else
		response.Write("ϵͳ���Ѿ��������Ŀ�ռ������ڣ��������Ŀ�ռ�����<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'>")
		
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
			Response.Write("<script language=javascript>alert('"&chkregtime&"�������ظ�ע�ᡣ');window.history.back(-1);</script>")
			Response.End
		end if
	end if
end sub



Sub ClearBinding()'ǿ�������״̬��
	dim sql,rs,id
	
	
	id=trim(Request("id"))
	
	
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
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
        alert("����ͬ��ע���������ע��!");
	return false;
   }	 
   
 
		 
	uid = del_space(document.oblogform.username.value);
     if (uid.length == 0)
     {
        alert("��������Ŀ�ռ���!");
	return false;
     }
	 
	pwd = del_space(document.oblogform.password.value);
     if (pwd.length == 0)
     {
        alert("����������!");
	return false;
     }
	 
	 pwd = del_space(document.oblogform.password.value);
     if (pwd.length < 6)
     {
        alert("����Ҫ����6���ַ���");
	return false;
     }

	pwd2 = del_space(document.oblogform.repassword.value);
     if (pwd2!=pwd)
     {
        alert("������������벻һ�£�");
		return false;
     }
	 
	 tishi = del_space(document.oblogform.question.value);
     if (tishi.length == 0)
     {
        alert("������������ʾ����");
        return false;
     }
	 
	 tsda = del_space(document.oblogform.an.value);
     if (tsda.length == 0)
     {
        alert("������������ʾ����𰸣�");
        return false;
     }
	 city = del_space(document.all("city").value);
     if (city.length == 0)
     {
        alert("��ѡ�����ڳ���!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("������Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("������������Ŀ��!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("��ѡ���������!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>

