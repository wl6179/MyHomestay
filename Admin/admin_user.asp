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
strFileName="admin_user.asp?UserSearch=" & UserSearch
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
<title>ע���й���ͥ�û�����</title>
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
    <td height="22" colspan=2 align=center><strong>�� �� �� ͥ �� �� �� ��</strong></td>
  </tr>
  <form name="form1" action="admin_user.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ����û���</strong></td>
      <td width="687" height="30"><select size=1 name="UserSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="0">���ע���500���û�</option>
    
		  <option value="3">VIP�û�</option>
		  <option value="4">ǰ̨����Ա</option>
		  <option value="5">�Ƽ��û�</option>
          <option value="6">�ȴ�����Ա��֤���û�</option>
          <option value="7">���б���ס���û�</option>
		  <option value="8">������See��ߵ��û�</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_user.asp">�й���ͥ�û�������ҳ</a>&nbsp;|&nbsp;<a href="/RegisterHome.asp" target="_blank">������û�</a></td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_user.asp">
  <tr class="tdbg">
    <td width="120"><strong>�û��߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">
	  <!--<option value="UserName" selected>�û���</option>-->
	  <option value="blogname" selected>����ʵ����</option>
      <option value="UserID" >�û�ID</option>
	  <option value="address" >����ַ</option>
      <option value="beizhu" >����ע</option>
	  <option value="xuwei_beizhu" >����ϸ</option>
	  <option value="lastlogintime" >��������δ��¼</option>

	  <option value="logintimes" >��¼����С��</option>
	  
	  <option value="province_city" >���ڵ���(ʡ/��)</option>
	  
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="UserSearch" type="hidden" id="UserSearch" value="10"> 
	  ��Ϊ�գ����ѯ�����û�</td>
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
elseif Action="ShowInfo" then
	call ShowInfo()

elseif Action="isValidate" then
	call change_isValidate()

elseif Action="isStop" then
	call change_isStop()
		
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_user.asp'>ע���й���ͥ�û�����</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 order by UserID desc"
			strGuide=strGuide & "���ע���500���û�"
		case 1
			sql="select top 100 * from oblog_user where user_level<=6 And user_level_team=0 order by log_count desc"
			strGuide=strGuide & "������������ǰ100���û�"
		case 2
			sql="select top 100 * from  oblog_user where user_level<=6 And user_level_team=0 order by log_count"
			strGuide=strGuide & "�����������ٵ�100���û�"
		case 3
			sql="select * from  oblog_user where user_level<=6 And user_level_team=0 and user_level=8 order by userid desc"
			strGuide=strGuide & "����VIP�û�"
		case 4
			sql="select * from  oblog_user where user_level<=6 And user_level_team=0 and user_level=9 order by userid desc"
			strGuide=strGuide & "ǰ̨����Ա"
		case 5
			sql="select * from  oblog_user where user_level<=6 And user_level_team=0 and user_isbest=1 order by userid desc"
			strGuide=strGuide & "�Ƽ��û�"
		case 6
			sql="select * from oblog_user where user_level<=6 And user_level_team=0 and User_Level=6 order by userID  desc"
			strGuide=strGuide & "�ȴ�������֤֤���û�"
		case 7
			sql="select * from oblog_user where user_level<=6 And user_level_team=0 and  LockUser =1 order by userID  desc"
			strGuide=strGuide & "���б���ס���û�"
		case 8
			sql="Select * From [oblog_user] Where user_level<=6 And user_level_team=0  Order By user_siterefu_num Desc,userID  Desc"
			strGuide=strGuide & "������See��ߵ��û�"
		case 10
			if Keyword="" then
				sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 order by userID desc"
				strGuide=strGuide & "�����û�"
			else
				select case strField
				case "UserID"
				if instr(1,Keyword,"MH00")>0 then Keyword=replace(Keyword,"MH00","")
				if instr(1,Keyword,"mh00")>0 then Keyword=replace(Keyword,"mh00","")
				'''response.Write Keyword
					if IsNumeric(Keyword)=false then
							FoundErr=true
						ErrMsg=ErrMsg & "<br><li>�û�ID������������</li>"
					else
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and userID =" & Clng(Keyword)
						strGuide=strGuide & "�û�ID����<font color=red> MH00" & Clng(Keyword) & " </font>���û�"
					end if
				case "blogname"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and blogname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��ʵ�����к��С� <font color=red>" & Keyword & "</font> �����û�"
					else
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and blogname= '" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��ʵ�������ڡ� <font color=red>" & Keyword & "</font> �����û�"
					end if
					
				case "address"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and address like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "ͨѶ��ַ�к��С� <font color=red>" & Keyword & "</font> �����û�"
					else
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and address='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "ͨѶ��ַ���ڡ� <font color=red>" & Keyword & "</font> �����û�"
					end if
				case "beizhu"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and beizhu like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��ע�к��С� <font color=red>" & Keyword & "</font> �����û�"
					else
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and beizhu='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��ע���ڡ� <font color=red>" & Keyword & "</font> �����û�"
					end if
				case "xuwei_beizhu"
					if is_sqldata=1 then
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and xuwei_beizhu like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��ϸ�к��С� <font color=red>" & Keyword & "</font> �����û�"
					else
						sql="select * from oblog_user where user_level<=6 And user_level_team=0 and xuwei_beizhu='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��ϸ���ڡ� <font color=red>" & Keyword & "</font> �����û�"
					end if
				case "logcount"
					sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 and log_count < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "������С�ڡ� <font color=red>" & Keyword & "</font> �����û�"
				case "logintimes"
					sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 and logintimes < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "��¼����С�ڡ� <font color=red>" & Keyword & "</font> �����û�"
				
				case "province_city"
					sql="Select Top 500 * From oblog_user Where user_level<=6 And user_level_team=0 "&_
									" And ( province Like '%" & Trim(Keyword) & "%' Or city Like '%" & Trim(Keyword) & "%' Or (LTRIM(RTRIM(province))='' And LTRIM(RTRIM(city))='') "&_
									" Or (province like '%����%' And LTRIM(RTRIM(city))='') ) "&_
								" Order By province Desc ,city Desc ,userID Desc"
					strGuide=strGuide & "���ڵ���(ʡ/��)Ϊ���� <font color=red>" & Keyword & "</font> �����û�"	
					
				case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 and datediff(d,lastlogintime,getdate())>"&clng(Keyword)&" order by userID  desc"
				   Else 
				   	sql="select top 500 * from oblog_user where user_level<=6 And user_level_team=0 and datediff('d',lastlogintime,now())>"&clng(Keyword)&" order by userID  desc"
				   End If 
					strGuide=strGuide & "�� <font color=red>" & Keyword & "</font> ������δ��¼���û�"
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
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ���û�</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ���û�</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���û�")
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
  <form name="myform" method="Post" action="admin_user.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> ��ͥ�ʺ�</font></td>
            <td width="50" height="22" align="center"><font color="#FFFFFF">��ʵ����</font></td>
            <td align="center"><font color="#FFFFFF">��ע</font></td>
			<td height="22" align="center"><font color="#FFFFFF">���ڵ���(ʡ/��)</font></td>
            
			<td align="center"> ��ַ</td>
			
			<td align="center"><font color="#cccccc">��ϸ</font></td>
            <td width="40" height="22" align="center"><font color="#FFFFFF"> ״̬</font></td>
            <td width="120" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='UserID' type='checkbox' onClick="unselectall()" id="UserID" value='<%=cstr(rs("userID"))%>'></td>
            <td width="30" align="center">MH00<%=rs("userID")%></td>
            <td width="80" align="center"><%
			response.write "<a href='?action=ShowInfo&userID="&rs("userID")&"'" 			
			response.write """  title='[ע������:"&rs("adddate")&"]..[��¼����:"&rs("LoginTimes")&"��]...[��¼IP:"&rs("LastLoginIP")&"]..[��һ�ε�¼ʱ��:"&rs("LastLoginTime")&"]'>" & rs("userName") & "</a>"
			%><br /><font color="#999999">See:</font><font color="#FF6600"><%=rs("user_siterefu_num")%></font> </td>
            <td align="center"> <%=rs("blogname")%> </td>
            
			<td align="center"> <%=rs("beizhu")%></td>
			
			<td align="center"> <%

		response.write rs("province")& "&nbsp;"
	
		response.write rs("city")

	%> </td>
            <td align="center"> <%=rs("address")%></td>
			
			<td align="center"> <%=rs("xuwei_beizhu")%></td>
			
            <td width="40" align="center"><%
	  if rs("LockUser")=1 then
	  	response.write "<font color=red>������</font>"
	  else
	  	response.write "����"
	  end if
	  %></td>
            <td width="120" align="center">
			

			
			<%
			
			response.write "<a href='admin_user.asp?Action=isValidate&isValidate="& rs("isValidate") &"&UserID=" & rs("userID") & "' title='����ʾ�ڽӴ���ͥ��Ŀ�У�'><font color=green>"
			If rs("isValidate")=1 Then
				Response.Write("<font color=green title='���û��Ѿ�ͨ����վ�����'>ͨ��</font>")
			ElseIf rs("isValidate")=0 Then
				Response.Write("<font color=red title='���û�δͨ��վ����ˣ����޷���վ����ʾ��'>�����</font>")
			End If
			Response.Write "</font></a>&nbsp;"
			
		response.write "<a href='admin_user.asp?Action=Modify&UserID=" & rs("userID") & "'>�޸�</a>&nbsp;"
		response.write "<a href='admin_user.asp?Action=gouser2&username=" & rs("username") & "' target='blank'>����̨</a>&nbsp;"
		if rs("LockUser")=0 then
			response.write "<br /><a href='admin_user.asp?Action=Lock&UserID=" & rs("userID") & "'>����</a>&nbsp;"
		else
            response.write "<br /><a href='admin_user.asp?Action=UnLock&UserID=" & rs("userID") & "'>����</a>&nbsp;"
		end if
        response.write "<a href='admin_user.asp?Action=Del&UserID=" & rs("userID") & "' onClick='return confirm(""ȷ��Ҫɾ�����û���"");'>ɾ��</a>&nbsp;"
		
		
		response.write "|<a href='admin_user.asp?Action=isStop&isStop="& rs("isStop") &"&UserID=" & rs("userID") & "' title='���ýӴ�״̬��'><font color=green>"
			If rs("isStop")=1 Then
				Response.Write("<font color=#666666 title='�ı�Ӵ�״̬'>�ѽӴ�</font>")
			ElseIf rs("isStop")=0 Then
				Response.Write("<font color=#038ad7 title='�ı�Ӵ�״̬��'>δ�Ӵ�</font>")
			End If
			Response.Write "</font></a>&nbsp;"
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
              ѡ�б�ҳ��ʾ�������û�</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=true">
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">�ƶ���
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">�й���ͥ�û�</option>
                <option value="7">ע����Ŀ����</option>
                <option value="8">VIP��Ŀ����</option>
                <option value="9">ǰ̨������</option>
              </select>-->
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" ִ �� "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_user.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�޸�ע���й���ͥ�û���Ϣ</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�û�����</TD>
      <TD width="60%"><%=rsUser("userName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>�û�������</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <select name="user_domainroot" ><%=oblog.type_domainroot(rsuser("user_domainroot"))%></select></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>��ʵ������</td>
      <td><input name=blogname   type=text id="blogname" value="<%=rsuser("blogname")%>" size=30 maxlength=20></td>
    </tr>
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�û��󶨵Ķ���������</td>
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
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD>ȷ������(����6λ)��<br>
        ������һ��ȷ��</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�������⣺<br>
        �����������ʾ����</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30>(�����û��뵽��̳�޸�) 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">����𰸣�<BR>
        �����������ʾ����𰸣�����ȡ������</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font></TD>
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
        <a href="mailto:<%=rsUser("userEmail")%>">�����û���һ������ʼ�</a> 
      </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">OICQ���룺</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">MSN��</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsUser("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�û�����</TD>
      <TD width="60%"><select name="User_Level" id="User_Level">
          <option value="6" <%if clng(rsUser("user_level"))=6 then response.Write("selected")%>>�й���ͥ�û�</option>
          <option value="7" <%if clng(rsUser("user_level"))=7 then response.Write("selected")%>>ע����Ŀ����</option>
          <option value="8" <%if clng(rsUser("user_level"))=8 then response.Write("selected")%>>vip��Ŀ����</option>
          <option value="9" <%if clng(rsUser("user_level"))=9 then response.Write("selected")%>>ǰ̨������</option>
        </select></TD>
    </TR>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ռ�(kb)��</td>
      <td><input name=user_upfiles_max   type=text value="<%=rsuser("user_upfiles_max")%>" size=20 maxlength=20>
        Ϊ��ʱΪϵͳĬ�����ã����赥�����ã�������kbֵ</td>
    </tr>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ֽ�(�ֽ�)��</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=20 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD>�Ƿ�Ϊ�Ƽ��û���</TD>
      <TD><input type="radio" name="isbest" value=1 <%if rsUser("user_isbest")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp; <input type="radio" name="isbest" value=0 <%if rsUser("user_isbest")<>1 then response.write "checked"%>>
        ��</TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�û�Ŀ¼��</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsUser("user_dir")%>" size=30 maxLength=50>
        ���ޱ�Ҫ�벻Ҫ�޸ģ���������û�Ŀ¼����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�û�״̬��</TD>
      <TD width="60%"><input type="radio" name="LockUser" value=0 <%if rsUser("LockUser")=0 then response.write "checked"%>>
        ����&nbsp;&nbsp; <input type="radio" name="LockUser" value=1 <%if rsUser("LockUser")=1 then response.write "checked"%>>
        ����</TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">��ע��(<strong>�벻Ҫ����10���֣�</strong>)</TD>
      <TD width="60%"><textarea name="beizhu" cols="38" ><%=rsUser("beizhu")%></textarea></TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ϸ��(<strong>255����֮��</strong>)</font></TD>
      <TD width="60%"><textarea name="xuwei_beizhu" cols="38" ><%=rsUser("xuwei_beizhu")%></textarea></TD>
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
<FORM name="Form1" action="admin_user.asp?action=DoUpdate" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">�����û���̬ҳ��</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          1�������������������û���̬ҳ�档<br>
          2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
      3 �������������û������¡� </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��ʼ�û�ID��</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      �û�ID��������д�������һ��ID�ſ�ʼ���и���</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">�����û�ID��</td>
    <td height="25"><input name="EndID" type="text" id="EndID" value="1000" size="10" maxlength="10">
      �����¿�ʼ������ID֮����û����ݣ�֮�����ֵ��ò�Ҫѡ�����</td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">&nbsp;</td>
    <td height="25"><input name="Submit" type="submit" id="Submit" value="���ɾ�̬ҳ��"></td>
  </tr>
</table>
</form>

<FORM name="Form1" action="admin_user.asp?action=DoUpdatelog" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">�������¾�̬ҳ��</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          1�������������������û���̬ҳ�档<br>
          2�����������ܽ��ǳ����ķ�������Դ�����Ҹ���ʱ��ܳ�������ϸȷ��ÿһ��������ִ�С�<br>
      3���������������£����¡�</p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">��ʼ����ID��</td>
    <td height="25"><input name="BeginID" type="text" id="BeginID" value="1" size="10" maxlength="10">
      �û�ID��������д�������һ��ID�ſ�ʼ���и���</td>
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
	response.Write("�����ⲿ���ݿ��û�������ʹ�ô˹���"):response.End()
end if
%>
<FORM name="Form1" action="admin_user.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼����Ŀ�����̨</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          ������������Ա��¼����Ŀ�Ĺ��������й���<br>
          ������Ŀ���������ϰ�ʱ���ɽ�����û���̨��<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">������Ŀ���ƣ�</td>
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
		response.Write("�����ⲿ���ݿ��û�������ʹ�ô˹���"):response.End()
	end if
	dim rs,username
	username=oblog.filt_badstr(trim(request("username")))
	if username="" then response.Write("�û�������Ϊ��"):response.End()
	if is_ot_user=0 then
		set rs=oblog.execute("select username,password from oblog_user where username='"&username&"'")
	else
		set rs=ot_conn.execute("select "&ot_username&","&ot_password&" from "&ot_usertable&" where "&ot_username&"='"&username&"'")
	end if
	if not rs.eof then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		set rs=nothing
		response.Redirect("../HomestayBackctrl-indexHome.asp")
	else
		set rs=nothing
		response.Write("�޴��û�"):response.End()
	end if
	
end sub
sub SaveModify()
	dim UserID,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,User_Level,LockUser,isbest
	dim rsUser,sqlUser
	dim blogname,usertype,user_upfiles_max,upfiles_size,user_domain,user_domainroot
	Dim beizhu,xuwei_beizhu
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
	
	beizhu= trim(request("beizhu"))
	xuwei_beizhu= trim(request("xuwei_beizhu"))
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
		ErrMsg=ErrMsg & "<br><li>��ָ���û�����</li>"
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
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
	
	rsuser("beizhu")=beizhu
	rsuser("xuwei_beizhu")=xuwei_beizhu
	rsUser.update
	rsUser.Close
	set rsUser=nothing
	oblog.showok "�޸ĳɹ�!",""
end sub


sub DelUser()
	dim rs,i
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ�����û�</li>"
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
	response.redirect "admin_user.asp"
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
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�������û�</li>"
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
		strtmp=strtmp&" window.location.replace('"&blogdir&"err.asp?message=���û��Ѿ�������������ϵ����Ա!');"&vbcrlf
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
	oblog.showok "�����û��ɹ�",""
end sub

sub UnLockUser()
	dim rs,udir
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�������û�</li>"
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
	oblog.showok "�����û��ɹ�",""
end sub

sub MoveUser()
	dim msg
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ƶ����û�</li>"
		exit sub
	end if
	dim User_Level
	User_Level=trim(request("User_Level"))
	if User_Level="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ���û���</li>"
		exit sub
	else
		User_Level=Clng(User_Level)
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>�й���ͥ�û�</font>����"
			sql="Update oblog_user set User_Level=6 where userID in (" & UserID & ")"
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update oblog_user set User_Level=7 where userID in (" & UserID & ")"
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update oblog_user set User_Level=8 where userID in (" & UserID & ")"
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update oblog_user set User_Level=9 where userID in (" & UserID & ")"
		end if
	else
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>�й���ͥ�û�</font>����"
			sql="Update oblog_user set User_Level=6 where userID=" & UserID 
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update oblog_user set User_Level=7 where userID="& UserID
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update oblog_user set User_Level=8 where userID=" & UserID 
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update oblog_user set User_Level=9 where userID=" & UserID
		end if
	end if
	Conn.Execute sql
	response.Redirect "admin_user.asp"
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
		Response.Write "<script>progress1.style.width ="""&int(i/p1*100)&"%"";progress1.innerHTML="""&int(i/p1*100)&"%"";pstr1.innerHTML=""ȫ�����ȣ���ǰ�û�ID:"&rsuser(0)&""";</script>" & VbCrLf
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




sub ShowInfo()
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
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	Dim strGuide
	strGuide=strGuide&"<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_user.asp'>ע���й���ͥ�û�����</a>&nbsp;&gt;&gt;&nbsp;" & "�й���ͥ<font color=red>"& rsuser("blogname") &"</font>����ϸ��Ϣ</td></tr></table>"
	
	Response.Write(strGuide)
%>

  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">(MH00<%=rsuser("userID")%>)�й���ͥ�û���ϸ��Ϣ</font></b>&nbsp;&nbsp;</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ͥ��¼�ʺţ�</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("userName")%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">Email��ַ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("userEmail")%> </font>&nbsp;&nbsp;
        <a href="mailto:<%=rsUser("userEmail")%>">�����û���һ������ʼ�</a> 
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ϵ��ʽ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("tel")%> &nbsp;&nbsp;&nbsp;<%=rsUser("mobile")%></font>
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">ͨѶ��ַ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("address")%> &nbsp;&nbsp;&nbsp;�ʱࣺ<%=rsUser("postnum")%></font>
      </TD>
    </TR>
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">���֪�����ǣ�</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("howknowsite")%></font> 
      </TD>
    </TR>
	
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''">
      <td>�û�������</td>
      <td><input name="user_domain" type="text" value="<%=oblog.filt_html(rsuser("user_domain"))%>" size=10 maxlength=20 /> <select name="user_domainroot" ><%=oblog.type_domainroot(rsuser("user_domainroot"))%></select></td>
    </tr>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ѡ��Ӵ���ͥ���ࣺ</td>
      <td><%
			Dim FamilyTypeRequest,FamilyType
				FamilyType = rsuser("FamilyType")
				
				Select case FamilyType
					Case "1"
					FamilyTypeRequest= "��ѽӴ���ͥ"
					Case "2"
					FamilyTypeRequest= "�շѽӴ���ͥ"
					Case "3"
					FamilyTypeRequest= "2008���˽Ӵ���ͥ"
					Case "1, 2"
					FamilyTypeRequest= "���/�շѽӴ���ͥ"
					Case "1, 3"
					FamilyTypeRequest= "��ѽӴ���ͥ/2008���˽Ӵ���ͥ"
					Case "1, 2, 3"
					FamilyTypeRequest= "���/�շ�/2008���˽Ӵ���ͥ"
					Case "2, 3"
					FamilyTypeRequest= "�շѽӴ���ͥ/2008���˽Ӵ���ͥ"
				End Select%>
	  <%=FamilyTypeRequest%></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td><strong>��ʵ������</strong></td>
      <td><strong><%=rsuser("blogname")%></strong></td>
    </tr>
	<TR class="tdbg" > 
      <TD width="40%">��ע��</TD>
      <TD width="60%"><%=rsUser("beizhu")%>
        </TD>
    </TR>
	
	
	<%if true_domain=1 then%>
	<tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>�û��󶨵Ķ���������</td>
      <td><input name=custom_domain   type=text id="custom_domain" value="<%=rsuser("custom_domain")%>" size=30 maxlength=20></td>
    </tr>
	<%end if%>
	
	<TR class="tdbg" > 
      <TD width="40%">���ڵ���(ʡ/��)��</TD>
      <TD width="60%"><%=rsUser("province")%>&nbsp;<%=rsUser("city")%> 
      </TD>
    </TR>
	
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>ѡ��ѧϰ��Χ��</td>
      <td><select name="usertype" id="usertype">
          <%if rsUser("user_classid")<>"" then
	  response.Write(oblog.show_class("user",rsUser("user_classid"),0))
	  else
	  response.Write(oblog.show_class("user",0,0))
	  end if
	  %>
        </select></td>
    </tr>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">����(����6λ)��<BR>
        ���������룬���ִ�Сд�� �벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ� </TD>
      <TD width="60%"> <INPUT   type=password maxLength=16 size=30 name=Password> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD>ȷ������(����6λ)��<br>
        ������һ��ȷ��</TD>
      <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=12> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font> </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�������⣺<br>
        �����������ʾ����</TD>
      <TD width="60%"> <INPUT name="Question"   type=text value="<%=rsUser("Question")%>" size=30>(�����û��뵽��̳�޸�) 
      </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">����𰸣�<BR>
        �����������ʾ����𰸣�����ȡ������</TD>
      <TD width="60%"> <INPUT   type=text size=30 name="Answer"> <font color="#FF0000">��������޸ģ�������(�����û��뵽��̳�޸�)</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�Ա�</TD>
      <TD width="60%"><%if rsUser("Sex")=1 then response.write "��"%>
        <%if rsUser("Sex")=0 then response.write "Ů"%>
        </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�������ڣ�</TD>
      <TD width="60%"><%=rsUser("birthday")%>
        </TD>
    </TR>
	
    
	
	
	<TR class="tdbg" > 
      <TD width="40%">����ְҵ��</TD>
      <TD width="60%"><%=rsUser("job")%> 
      </TD>
    </TR>
	
	
	
	<TR class="tdbg" > 
      <TD width="40%">��ͥ������</TD>
      <TD width="60%"><%=rsUser("familynumber")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">��ͥ�ṹ��</TD>
      <TD width="60%"><%=rsUser("familymember")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">���ݽ��ܣ�</TD>
      <TD width="60%">�ṹ��<%=rsUser("houseframe")%> &nbsp;&nbsp;&nbsp;�����<%=rsUser("housespace")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ��г��</TD>
      <TD width="60%"><%=rsUser("pet")%>&nbsp;&nbsp;&nbsp;<%=rsUser("PetType")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">����Ա�</TD>
      <TD width="60%"><%=rsUser("asksex")%>
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ��˵���ģ�</TD>
      <TD width="60%"><%=rsUser("issaychinese")%> 
      </TD>
    </TR>
	
	
	
	<TR class="tdbg" > 
      <TD width="40%">Ӣ��ˮƽ��</TD>
      <TD width="60%"><%=rsUser("englishlevel")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�Ƿ���˽�ҳ���</TD>
      <TD width="60%"><%=rsUser("car")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" style="display:none;" > 
      <TD width="40%">װ�ޱ�׼��</TD>
      <TD width="60%"><%=rsUser("fitment")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">����������ͣ�</TD>
      <TD width="60%"><%=rsUser("toilet")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�з���ԣ�</TD>
      <TD width="60%"><%=rsUser("computer")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�з���������</TD>
      <TD width="60%"><%=rsUser("internet")%> 
      </TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�ͷ�����(��ע)��</TD>
      <TD width="60%"><%=rsUser("remarkinfo")%> 
      </TD>
    </TR>
	
	
	
	
    
	
	
	<TR class="tdbg" > 
      <TD width="40%"><font color="#038ad7">��ϸ��</font></TD>
      <TD width="60%"><font color="#038ad7"><%=rsUser("xuwei_beizhu")%></font>
        </TD>
    </TR>
	
	
	
	
	<TR style="display:none;" class="tdbg" > 
      <TD width="40%">OICQ���룺</TD>
      <TD width="60%"> <INPUT name=OICQ value="<%=rsUser("qq")%>" size=30 maxLength=20></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">MSN��</TD>
      <TD width="60%"> <INPUT name=msn value="<%=rsUser("Msn")%>" size=30 maxLength=50></TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�û�����</TD>
      <TD width="60%"><select name="User_Level" id="User_Level">
          <option value="6" <%if clng(rsUser("user_level"))=6 then response.Write("selected")%>>�й���ͥ�û�</option>
          <option value="7" <%if clng(rsUser("user_level"))=7 then response.Write("selected")%>>ע����Ŀ����</option>
          <option value="8" <%if clng(rsUser("user_level"))=8 then response.Write("selected")%>>vip��Ŀ����</option>
          <option value="9" <%if clng(rsUser("user_level"))=9 then response.Write("selected")%>>ǰ̨������</option>
        </select></TD>
    </TR>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ռ�(kb)��</td>
      <td><input name=user_upfiles_max   type=text value="<%=rsuser("user_upfiles_max")%>" size=20 maxlength=20>
        Ϊ��ʱΪϵͳĬ�����ã����赥�����ã�������kbֵ</td>
    </tr>
    <tr style="display:none;" class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td>���ϴ��ֽ�(�ֽ�)��</td>
      <td><input name=upfiles_size   type=text id="upfiles_size" value="<%=rsuser("user_upfiles_size")%>" size=20 maxlength=20></td>
    </tr>
    <TR class="tdbg" > 
      <TD>�Ƿ�Ϊ�Ƽ��û���</TD>
      <TD><%if rsUser("user_isbest")=1 then response.write "��"%>
        <%if rsUser("user_isbest")<>1 then response.write "��"%>
        </TD>
    </TR>
    <TR style="display:none;" class="tdbg" > 
      <TD width="40%">�û�Ŀ¼��</TD>
      <TD width="60%"> <INPUT name=user_dir value="<%=rsUser("user_dir")%>" size=30 maxLength=50>
        ���ޱ�Ҫ�벻Ҫ�޸ģ���������û�Ŀ¼����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%">�û�״̬��</TD>
      <TD width="60%"><%if rsUser("LockUser")=0 then response.write "����"%>
        <%if rsUser("LockUser")=1 then response.write "����"%>
        </TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input type="button" onClick="history.go(-1);" value=" ������һҳ " style="cursor:hand;"></TD>
    </TR>
  </TABLE>

<%
	rsUser.close
	set rsUser=nothing
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


Sub change_isValidate()'�����й���ͥ�Ƿ�ͨ����վ�����...ͨ���������վ����ʾ����Ϊ������ͥ�û�.
	dim sql,rs,id
	Dim isValidate
	
	id=trim(Request("UserID"))
	isValidate= Cint(Request("isValidate"))
	
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

	If isValidate= 1 Then
		conn.execute("update oblog_user set isValidate=0 where userid=" & id )
	ElseIf isValidate= 0 Then
		conn.execute("update oblog_user set isValidate=1 where userid=" & id )
	End If

	oblog.showok "�������","admin_user.asp"
	
End Sub


Sub change_isStop()'�����й���ͥ�Ƿ�ͨ����վ�����...ͨ���������վ����ʾ����Ϊ������ͥ�û�.
	dim sql,rs,id
	Dim isStop
	
	id=trim(Request("UserID"))
	isStop= Cint(Request("isStop"))
	
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

	If isStop= 1 Then
		conn.execute("update oblog_user set isStop=0 where userid=" & id )
	ElseIf isStop= 0 Then
		conn.execute("update oblog_user set isStop=1 where userid=" & id )
	End If

	oblog.showok "�������","admin_user.asp"
	
End Sub
%>