<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
Dim tName,t
tName = "�������"

Const MaxPerPage = 20
Dim strFileName
Dim totalPut, TotalPages
Dim rs, sql
Dim id, usersearch, Keyword, strField, uid, action
Dim selectsub, substr, wsql, allsub

t = Trim(request("t"))
Keyword = Trim(request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
strField = Trim(request("Field"))
usersearch = Trim(request("usersearch"))
selectsub = Trim(request("selectsub"))
action = Trim(request("action"))
id = oblog.filt_badstr(Trim(request("id")))
uid = oblog.filt_badstr(Trim(request("uid")))
strFileName = "admin_userCooperate_PassTeacher.asp?t=" & t
If usersearch = "" Then
    usersearch = 0
Else
    usersearch = CLng(usersearch)
    strFileName = "admin_userCooperate_PassTeacher.asp?usersearch=" & usersearch & "&t=" & t
End If
If selectsub <> "" Then
    selectsub = CLng(selectsub)
    strFileName = "admin_userCooperate_PassTeacher.asp?usersearch=10&selectsub=" & selectsub & "&t=" & t
Else
    selectsub = 0
End If
If Keyword <> "" Then
    strFileName="admin_userCooperate_PassTeacher.asp?usersearch=10&keyword="&Keyword&"&Field="&strField & "&t="  & t
End If
If request("page") <> "" Then CurrentPage = CInt(request("page")) Else CurrentPage = 1
if oblog.logined_ulevel<>9 then wsql=" and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" )"




dim FoundErr,ErrMsg
dim tmpDays
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

%>
<html>
<head>
<title>ͨ����˵�������Ϲ���</title>
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
    <td height="22" colspan=2 align=center><strong>��ͨ���ҷ���˵� ������Ϲ���</strong></td>
  </tr>
  <form name="form1" action="admin_userCooperate_PassTeacher.asp" method="get">
    <tr class="tdbg"> 
      <td width="120" height="30"><strong>���ٲ���������ϣ�</strong></td>
      <td width="507" height="30"><input type="hidden" name="t" value="<%=t%>">
	  
	  <select size=1 name="usersearch" onChange="javascript:submit()">
          <option value="10">��ѡ������</option>
          <option value="0">�г�����<%=tName%></option>
          <option value="1" <% If usersearch=1 Then Response.Write("selected")%>>�г�û�й���������ַURL��<%=tName%></option>
		  <option value="2" <% If usersearch=2 Then Response.Write("selected")%>>�г��Ѿ���д����������ַURL��<%=tName%></option>
          <option value="3" <% If usersearch=3 Then Response.Write("selected")%>>�ڲݸ������<%=tName%></option>
			<option value="4" <% If usersearch=4 Then Response.Write("selected")%>>ǿ�Ʋ鿴�ѱ�ɾ����<%=tName%></option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_userCooperate_PassTeacher.asp">ͨ����˵�������Ϲ�����ҳ</a></td>
    </tr>
  </form>
  
  
  <form name="form2" method="post" action="admin_userCooperate_PassTeacher.asp">
  <tr class="tdbg">
    <td width="120"><strong>������ϸ߼���ѯ��</strong></td>
    <td >
      <select name="Field" id="Field">

            <option value="logid" <% If strField="logid" Then Response.Write("selected")%>>��<%=tName%>ID�Ų�ѯ</option>
            <option value="topic" <% If strField<>"logid" Then Response.Write("selected")%>>������ע�Ͳ�ѯ</option>
			<option value="author" <% If strField="author" Then Response.Write("selected")%>>ָ��ĳ�����ʺŵ�<%=tName%></option>
			<option value="logtext" <% If strField="logtext" Then Response.Write("selected")%>>���ύ���ݲ�ѯ</option>
			
         </select>	<input name="usersearch" type="hidden" value="10">
		 
         <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
		 
         <input type="submit" name="Submit2" value=" ���� "> 
	  ��Ϊ�գ����ѯ����ͨ����˵��������</td>
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
	
elseif Action="del" then
    	Call delblog
elseif Action="state_shenhe" then
    	Call change_state_shenhe
		
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
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_userCooperate_PassTeacher.asp'>ͨ����˵�������Ϲ���</a>&nbsp;&gt;&gt;&nbsp;"
	select case UserSearch
		case 0
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 order by addtime desc"
			strGuide=strGuide & "���500��ͨ����˵��������"
		case 1'�г�û�з���URL��������ϣ�
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 And logPublic_path='' order by addtime desc"
			strGuide=strGuide & "û����д������ַURL��ǰ500���������"
		case 2
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 And logPublic_path<>'' order by addtime desc"
			strGuide=strGuide & "�Ѿ���д�˷�����ַURL��ǰ500���������"
		case 3
			sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft=1 order by addtime desc"
			strGuide=strGuide & "�б����ڼ����û��ݸ������ǰ500���������"
		case 4
			sql="select * from  [oblog_logCooperateSubmit] where isDelete=1 And state_shenhe=1 And isdraft<>1 order by addtime desc"
			strGuide=strGuide & "ǿ�Ʋ鿴�ѱ�ɾ����ǰ500���������"
		case 5
			sql="select * from  [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and user_isbest=1 order by addtime desc"
			strGuide=strGuide & "�Ƽ��û�"
		case 6
			sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and User_Level=6 order by userID  desc"
			strGuide=strGuide & "�ȴ�������֤֤���û�"
		case 7
			sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and  LockUser =1 order by userID  desc"
			strGuide=strGuide & "���б���ס���û�"
		case 10
		
			if Keyword="" then
				sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 order by addtime desc"
				strGuide=strGuide & "�����������"
			else
				select case strField
				case "logid"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>�������ID������������</li>"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logid =" & Clng(Keyword)
						strGuide=strGuide & "�������ID����<font color=red> " & Clng(Keyword) & " </font>���������"
					end if
				case "topic"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and topic like '%" & Keyword & "%' order by addtime  desc"
						strGuide=strGuide & "������ϵı���ע���к��С� <font color=red>" & Keyword & "</font> �����������"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and topic= '" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "������ϵı���ע�͵��ڡ� <font color=red>" & Keyword & "</font> �����������"
					end if
					
				case "author"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and author like '%" & Keyword & "%' order by addtime  desc"
						strGuide=strGuide & "�����ʻ������У��� <font color=red>" & Keyword & "</font> �����ʺ� ����д������������ϣ�"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and author='" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "�����ʻ����ڡ� <font color=red>" & Keyword & "</font> �����ʺ� ����д�������������"
					end if
				case "logtext"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logtext like '%" & Keyword & "%' order by addtime desc"
						strGuide=strGuide & "�ύ�����к��С� <font color=red>" & Keyword & "</font> �����������"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>100 and logtext='" & Keyword & "' order by addtime  desc"
						strGuide=strGuide & "�ύ���ݵ��ڡ� <font color=red>" & Keyword & "</font> �����������"
					end if
					
					
					
					
				case "blogname"
					if is_sqldata=1 then
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and blogname like '%" & Keyword & "%' order by userID  desc"
						strGuide=strGuide & "��ʵ�����к��С� <font color=red>" & Keyword & "</font> �����û�"
					else
						sql="select * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and blogname='" & Keyword & "' order by userID  desc"
						strGuide=strGuide & "��ʵ�������ڡ� <font color=red>" & Keyword & "</font> �����û�"
					end if
				case "logcount"
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and log_count < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "������С�ڡ� <font color=red>" & Keyword & "</font> �����û�"
				case "logintimes"
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and logintimes < " & clng(Keyword) & " order by userID  desc"
					strGuide=strGuide & "��¼����С�ڡ� <font color=red>" & Keyword & "</font> �����û�"
				case "lastlogintime"
				   If is_sqldata=1 Then 
					sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and datediff(d,lastlogintime,getdate())>"&clng(Keyword)&" order by userID  desc"
				   Else 
				   	sql="select top 500 * from [oblog_logCooperateSubmit] where isDelete=0 And state_shenhe=1 And isdraft<>1 and datediff('d',lastlogintime,now())>"&clng(Keyword)&" order by userID  desc"
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
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ���������</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ���������</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���������")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���������")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���������")
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
  <form name="myform" method="Post" action="admin_userCooperate_PassTeacher.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td width="180" height="22" align="center"><font color="#FFFFFF"> ����ע��</font></td>
            <td height="22" align="center"><font color="#FFFFFF">�ύ��(������)</font></td>
            
			
            <td align="center"><font color="#FFFFFF">�ύʱ��</font></td>
            
            <td width="80" height="22" align="center"><font color="#FFFFFF"> ��˲���</font></td>
            <td width="150" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("logid"))%>'></td>
            <td width="30" align="center"><%=rs("logid")%></td>
            <td width="180" align="center"><%
			response.write "" & rs("topic") & ""
			If rs("isdraft")=1 Then response.write "<font color=red title='����������ѱ���������Ϊ�˲ݸ�'>*�ݸ�</font>"
			If rs("logPublic_path")<>"" And Len(rs("logPublic_path"))>6 Then
				response.write "<br />URL��<a href='"& rs("logPublic_path") &"'><font color=#038ad7>"& rs("logPublic_path") &"</font></a>&nbsp;"
			Else
				
			End If
			%> </td>
            <td align="center"> <%
			Response.Write rs("author")

			%> </td>
            
			
            <td align="center"> <%
	if rs("addtime")<>"" then
		response.write rs("addtime")
	else
		response.write "&nbsp;"
	end if
	%> </td>
            
            <td width="80" align="center"><input type="button" value=" ȡ �� " onClick="if (confirm('ȷ��Ҫȡ����������ϵ����״̬��') == false)return false;window.location.href='admin_userCooperate_PassTeacher.asp?action=state_shenhe&state_shenhe=<%=rs("state_shenhe")%>&logPublic_path=<%=rs("logPublic_path")%>&logid=<%=rs("logid")%>&userid=<%=rs("userid")%>'"></td>
			
            <td width="150" align="center"><%
		response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=gouser2&userid=" & rs("userid") & "&logid="& rs("logid") &"&actionTemp=preview' target='blank'>Ԥ������</a>&nbsp;"
		
		If rs("logPublic_path")<>"" And Len(rs("logPublic_path"))>6 Then
			response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=Modify&logid=" & rs("logid") & "' title='"& rs("logPublic_path") &"'><font color=#038ad7>�ѹ���URL��</font></a>&nbsp;"
		Else
			response.write "<a href='admin_userCooperate_PassTeacher.asp?Action=Modify&logid=" & rs("logid") & "' title='�������֮ǰ���ȹ����ô���̵�վ�㷢����ַURL��'><font color=red>�����������ַ</font></a>&nbsp;"
		End If
		
		response.write "<br /><a href='admin_userCooperate_PassTeacher.asp?Action=gouser2&userid=" & rs("userid") & "&logid="& rs("logid") &"' title='������ǿ���޸ļ����û��ύ�������������' target='blank'>ǿ���޸�</a>&nbsp;"
		
        response.write "<a href='admin_userCooperate_PassTeacher.asp?action=del&id=" & rs("logid") & "&t=" & t & "' onClick='return confirm(""ȷ��Ҫɾ�������������"");'>ɾ��</a>&nbsp;"
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
              ѡ�б�ҳ��ʾ���������</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="Del" checked onClick="document.myform.User_Level.disabled=true">
              ɾ��&nbsp;&nbsp;&nbsp;&nbsp; 
              <!--<input name="Action" type="radio" value="Move" onClick="document.myform.User_Level.disabled=false">�ƶ���
              <select name="User_Level" id="User_Level" disabled>
                <option value="6">�й���ͥor��������û�</option>
                <option value="7">ע����Ŀ����</option>
                <option value="8">VIP��Ŀ����</option>
                <option value="9">ǰ̨������</option>
              </select>
              &nbsp;&nbsp; -->
              <input type="submit" name="Submit" value=" ִ �� "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub


sub Modify()
	dim logid
	dim rslog,sqllog
	logid=trim(request("logid"))
	if logid="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		logid=Clng(logid)
	end if
	Set rslog=Server.CreateObject("Adodb.RecordSet")
	sqllog="select * from [oblog_logCooperateSubmit] where logid=" & logid
	if not IsObject(conn) then link_database
	rslog.Open sqllog,Conn,1,3
	if rslog.bof and rslog.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ����������ϣ�</li>"
		rslog.close
		set rslog=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">��������ϵ� վ�㷢����ַURL</font></b></TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">����ע�ͣ�</TD>
      <TD width="60%"><%=rslog("topic")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	
	<TR class="tdbg" > 
      <TD width="40%">�ύ��(������)��</TD>
      <TD width="60%"><%=rslog("author")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
	
    <TR class="tdbg" > 
      <TD width="40%">������ַURL��</TD>
      <TD width="60%"><input name="logPublic_path" type="text" id="logPublic_path" value="<%=rslog("logPublic_path")%>" size="50">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
    </TR>
   
   
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
	  <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'>
	  <input name="logid" type="hidden" id="logid" value="<%=rslog("logid")%>"></TD>
    </TR>
  </TABLE>
</form>
<%
	rslog.close
	set rslog=nothing
end sub


sub UpdateUser()
%>
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=DoUpdate" method="post">
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

<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=DoUpdatelog" method="post">
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
<FORM name="Form1" action="admin_userCooperate_PassTeacher.asp?action=gouser2" method="post" target="_blank">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr align="center" class="title"> 
      <td height="22" colspan="2"><strong><font color="#FFFFFF">��¼���û������̨</font></strong></td>
  </tr>
  <tr class="tdbg"> 
      <td colspan="2"><p>˵����<br>
          ������������Ա��¼���û��Ĺ��������й���<br>
          ���û����������ϰ�ʱ���ɽ�����û���̨��Э���û����в�����<br>
        </p>
      </td>
  </tr>
  <tr class="tdbg"> 
    <td height="25">�û��˺ţ�</td>
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
	dim rs,username,userid
	userid=oblog.filt_badstr(trim(request("userid")))
	if userid="" then response.Write("����Ϊ��"):response.End()
	userid=Cint(request("userid"))
	if is_ot_user=0 then
		set rs=oblog.execute("select username,password from [oblog_user] where userid='"&userid&"'")
	else
		'''set rs=ot_conn.execute("select "&ot_username&","&ot_password&" from "&ot_usertable&" where "&ot_username&"='"&username&"'")
	end if
	if not rs.eof then
		oblog.SaveCookie rs(0), rs(1), 0, ""
		set rs=nothing
		
		If Request("actionTemp")="preview" Then'������Ԥ��������
		response.Redirect("../HomestayBackctrl-logPreviewCooperate.asp?logid="& Request("logid") &"&t=0")
		Else
		response.Redirect("../HomestayBackctrl-PostSubmitCooperate.asp?logid="& Request("logid") &"&t=0")
		End If
	else
		set rs=nothing
		response.Write("�޴��û�"):response.End()
	end if
	
end sub


sub SaveModify()
	dim UserID,Password,PwdConfirm,Question,Answer,Sex,Email,Homepage,OICQ,MSN,User_Level,LockUser,isbest
	dim rsUser,sqlUser
	dim blogname,usertype,user_upfiles_max,upfiles_size,user_domain,user_domainroot
	
	Dim logid,logPublic_path
	
	logPublic_path=trim(request.Form("logPublic_path"))
	logid=trim(request("logid"))
	
	if logid="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		logid=Clng(logid)
	end if
	
	
'	if User_Level="" then
'		FoundErr=true
'		ErrMsg=ErrMsg & "<br><li>��ָ���û�����</li>"
'	else
'		User_Level=CLng(User_Level)
'	end if
	
	

	if founderr=true then
		''set rsuser=nothing
		exit sub
	end if
	
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [oblog_logCooperateSubmit] where logid=" & logid
	if not IsObject(conn) then link_database
	rsUser.Open sqlUser,Conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ����������ϣ�</li>"
		rsUser.close
		set rsUser=nothing
		exit sub
	end if
	

	rsUser("logPublic_path")=logPublic_path


	rsUser.update
	rsUser.Close
	set rsUser=nothing
	oblog.showok "�޸ĳɹ�!",""
end sub


sub DelUser()
	dim rs,i
	if UserID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ�����������</li>"
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
	response.redirect "admin_userCooperate_PassTeacher.asp"
end sub

sub deloneuser(logid)
	'''Dim logid
	logid=Clng(logid)
	


		oblog.execute("delete from oblog_logCooperateSubmit where logid="&logid)

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
	rs.open "select lockuser,user_dir from [oblog_logCooperateSubmit] where userid="&userid,conn,1,3
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
	rs.open "select lockuser,user_dir from [oblog_logCooperateSubmit] where userid="&userid,conn,1,3
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
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>�й���ͥor��������û�</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=6 where userID in (" & UserID & ")"
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=7 where userID in (" & UserID & ")"
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=8 where userID in (" & UserID & ")"
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=9 where userID in (" & UserID & ")"
		end if
	else
		if User_Level=6 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>�й���ͥor��������û�</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=6 where userID=" & UserID 
		elseif User_Level=7 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ע����Ŀ����</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=7 where userID="& UserID
		elseif User_Level=8 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>VIP��Ŀ����</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=8 where userID=" & UserID 
		elseif User_Level=9 then
			msg="&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>ǰ̨������</font>����"
			sql="Update [oblog_logCooperateSubmit] set User_Level=9 where userID=" & UserID
		end if
	end if
	Conn.Execute sql
	response.Redirect "admin_userCooperate_PassTeacher.asp"
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
	set rsuser=oblog.execute("select count(userid) from [oblog_logCooperateSubmit] where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID))
	p1=rsuser(0)
	set rsuser=oblog.execute("select userid from [oblog_logCooperateSubmit] where userID>=" & clng(BeginID) & " and userID<=" & clng(EndID)&" order by userid")
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




Sub delblog()'ɾ��������ϣ�
	id = Request("id")
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫɾ����" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id=FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            deloneblog (n(i))
        Next
    Else
        deloneblog (id)
    End If
    oblog.showok "ɾ��" & tName & "�ɹ���", ""
End Sub

Sub deloneblog(logid)
    logid = CLng(logid)
    Dim uid, delname,rst,fso, sid,rst1
    Set rst = server.CreateObject("adodb.recordset")
	If Not IsObject(conn) Then link_database
    rst.open "select userid,logfile,subjectid,logtype from oblog_logCooperateSubmit where logid="&logid,conn,1,3
    If Not rst.EOF Then
        uid = rst(0)
        delname = Trim(rst(1))
        sid = rst(2)
        '����ͼƬ��¼
        If rst("logtype")=1 Then
        	Call DeletePhotos(logid)
   	 	End If
		'��ʵ������Ҫ���������ļ�����
		if true_domain=1 and delname<>"" then
			if instr(delname,"archives") then
				delname=right(delname,len(delname)-InstrRev(delname,"archives")+1)
			else
				delname=right(delname,len(delname)-InstrRev(delname,"/"))
			end if
			if oblog.logined_ulevel=9 then
				set rst1=oblog.execute("select user_dir,user_folder from oblog_user where userid="&clng(uid))
				if not rst1.eof then
					delname=rst1(0)&"/"&rst1(1)&"/"&delname
				end if
				set rst1=nothing
			else
				delname=oblog.logined_udir&"/"&oblog.logined_ufolder&"/"&delname
			end if
			'response.write(delname)
			'response.end
		end if
'''        If delname <> "" Then
'''            Set fso = server.CreateObject("Scripting.FileSystemObject")
'''            If fso.FileExists(server.MapPath(delname)) Then fso.DeleteFile server.MapPath(delname)
'''        End If
        '''rst.Delete'''WL;��ɾ����¼����Ϊ���䡯���ء���
			oblog.execute ("update oblog_logCooperateSubmit set isDelete=1 where logid=" & CLng(logid))'�Ƿ��ѱ��û�ɾ������ʾ��������Ϲ��ڣ����������Ƕ��մ���...WL;
			oblog.execute ("update oblog_logCooperateSubmit set isAdminDeleted=1 where logid=" & CLng(logid))
        rst.Close
        '---------------------------------
'''		Call Tags_UserDelete(logid)
        '--------------------------------------------
        '�˴���Ҫ��һ���޸ģ�������
        oblog.execute ("update oblog_user set log_count=log_count-1 where userid=" & uid)
        oblog.execute ("delete from oblog_comment where mainid=" & CLng(logid))
        oblog.execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & CLng(sid))
        '--------------------------------------------
'''        Dim blog
'''        Set blog = New class_blog
'''        blog.userid = uid
'''        blog.Update_Subject uid
'''        blog.update_index 0
'''        blog.update_newblog (uid)
'''        Set blog = Nothing
        Set fso = Nothing
        Set rst = Nothing  
		
		
    Else
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
End Sub



Sub change_state_shenhe()'��˲���
	dim sql,rs,id,userid
	Dim isValidate
	Dim templogPublic_path
	
	id=trim(Request("logid"))
	userid= Trim(Request("userid"))
	isValidate= Cint(Request("state_shenhe"))
	
	templogPublic_path= Trim(Request("logPublic_path"))
	
	if templogPublic_path="" or Len(templogPublic_path)<6 then
		'FoundErr=True
		'ErrMsg=ErrMsg & "<br><li>���� ����������ַURL ,����ͨ����� ��</li>"
		'exit sub
	else
		id=CLng(id)
	end if
	
	if id="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		id=CLng(id)
	end if
	
	if userid="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>userid�������㣡</li>"
		exit sub
	else
		userid= CLng(userid)
	end if
	

	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if

Dim TempWord
	If isValidate= 1 Then
		conn.execute("update oblog_logCooperateSubmit set state_shenhe=0 where logid=" & id )
		
		conn.execute("Update oblog_user Set passTeacher_count = passTeacher_count-1 Where userid=" & userid )
		TempWord = "�ɹ�ͨ����ˣ�"
	ElseIf isValidate= 0 Then
		conn.execute("update oblog_logCooperateSubmit set state_shenhe=1 where logid=" & id )
		
		conn.execute("Update oblog_user Set passTeacher_count = passTeacher_count+1 Where userid=" & userid )
		TempWord = "��ȡ����ˣ�"
	End If

	oblog.showok "��ȡ������ˣ�","admin_userCooperate_PassTeacher.asp"
	
End Sub


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