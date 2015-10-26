<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%
const MaxPerPage=20
dim strFileName
dim totalPut,TotalPages
dim rs, sql
dim TagID,TagSearch,Keyword,strField
dim Action,FoundErr,ErrMsg
dim tmpDays
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
TagSearch=trim(request("TagSearch"))
Action=trim(request("Action"))
TagID=trim(Request("TagID"))
'ComeUrl=Request.ServerVariables("HTTP_REFERER")


if TagSearch="" then
	TagSearch=10
else
	TagSearch=Clng(TagSearch)
end if
strFileName="admin_tags.asp?TagSearch=" & TagSearch
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
<title>TAG����</title>
<meta http-equiv="Content-Type" content="text/html; charSet =gb2312">
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
    <td height="22" colspan=2 align=center><strong>ϵ ͳ TAG �� ��</strong></td>
  </tr>
  <form name="form1" action="admin_tags.asp" method="get">
    <tr class="tdbg"> 
      <td width="100" height="30"><strong>���ٲ���TAG��</strong></td>
      <td  height="30"><select size=1 name="TagSearch" onChange="javascript:submit()">
          <option value=>��ѡ���ѯ����</option>
		  <option value="1">ʹ��Ƶ����ߵ�100��TAG</option>          
          <option value="2">ʹ��Ƶ����͵�100��TAG</option>
          <option value="3">ʹ����Ϊ0��TAG</option>
          <option value="4">�����õ�TAG</option>
        </select>       </td>
    </tr>
  </form>
  <form name="form2" method="post" action="admin_tags.asp">
  <tr class="tdbg">
    <td width="100"><strong>�߼���ѯ��</strong></td>
    <td>
      <select name="Field" id="Field">
	  <option value="TagName" selected>TAG����</option>
      <option value="TagID" >TAG ID</option>      
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit" name="Submit2" value=" �� ѯ ">
      <input name="TagSearch" type="hidden" id="TagSearch" value="10"> 
	  ��Ϊ�գ����ѯ����TAG</td>
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
elseif Action="merge" then
	call MergeTags()
elseif Action="batchlock" then
	call batchlock()
else
	call main()
end if
if FoundErr=true then
	call WriteErrMsg()
end if

sub main()
	dim strGuide
	strGuide="<table width='100%'><tr><td align='left'>�����ڵ�λ�ã�<a href='admin_tags.asp'>ϵͳTAG����</a>&nbsp;&gt;&gt;&nbsp;"
	select case TagSearch
		case 1
			sql="Select Top 100 * From oblog_tags Where iState=1 And iNum>0 order by iNum Desc"
			strGuide=strGuide & "ʹ��Ƶ����ߵ�100��TAG"
		case 2
			sql="Select Top 100 * From oblog_tags Where iState=1 And iNum>0 order by iNum"
			strGuide=strGuide & "ʹ��Ƶ�����ٵ�100��TAG"
		case 3
			sql="Select  * From oblog_tags Where iState=1 And iNum=0"
			strGuide=strGuide & "ʹ����Ϊ0��TAG"
		case 4
			sql="Select  * From oblog_tags Where iState=0"
			strGuide=strGuide & "�����õ�TAG"		
		case 10
			if Keyword="" then
				sql="Select Top 100 * From oblog_tags Where iState=1 And iNum>0 order by iNum Desc"
				strGuide=strGuide & "����TAG"
			else
				select case UCASE(strField)
				case "TAGID"
					if IsNumeric(Keyword)=false then
						FoundErr=true
						ErrMsg=ErrMsg & "<br><li>TAG ID������������</li>"
					else
						sql="select * from oblog_tags where Tagid =" & Clng(Keyword)
						strGuide=strGuide & "TAG ID����<font color=red> " & Clng(Keyword) & " </font>"
					end if
				case "TAGNAME"
						sql="select * from oblog_tags where name like '%" & Keyword & "%' order by iNum Desc"
						strGuide=strGuide & "���С� <font color=red>" & Keyword & "</font> ����TAG"
				end select
			end if
		case else
			FoundErr=true
			ErrMsg=ErrMsg & "<br><li>����Ĳ�����</li>"
	end select
	strGuide=strGuide & "</td><td align='right'>"
	if FoundErr=true then exit sub
	if not IsObject(conn) then link_database
	Set  rs=Server.CreateObject("Adodb.RecordSet")
	'Response.Write sql
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & "���ҵ� <font color=red>0</font> ��TAG</td></tr></table>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & "���ҵ� <font color=red>" & totalPut & "</font> ��TAG</td></tr></table>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"�� TAG ")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"�� TAG ")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"�� TAG ")
	    	end if
		end if
	end if
	rs.Close
	Set  rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<table width='98%' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  <form name="myform" method="Post" action="admin_tags.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
          <tr class="title"> 
            <td width="30" align="center"><font color="#FFFFFF">ѡ��</font></td>
            <td width="30" align="center"><font color="#FFFFFF">ID</font></td>
            <td  height="22" align="center"><font color="#FFFFFF"> TAG����</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF">ʹ�ô���</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> ״̬</font></td>
            <td width="80" height="22" align="center"><font color="#FFFFFF"> 
              ����</font></td>
          </tr>
          <%do while not rs.EOF %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td align="center"><input name='TagID' type='checkbox' onClick="unselectall()" id="TagID" value='<%=cstr(rs("TagID"))%>'></td>
            <td  align="center"><%=rs("TagID")%></td>
            <td  align="Left" style="word-break: break-all; word-wrap:break-word;">&nbsp;&nbsp;<%
			response.write "<a href='admin_tags.asp?Action=Modify&TagID=" & rs("TagID") & "'" 			
			response.write """>" & rs("Name") & "</a>"
			%> </td>
            <td  align="center"> <%
	if rs("iNum")<>"" then
		response.write rs("iNum")
	else
		response.write "0"
	end if
	%> </td>
            <td  align="center"><%
	  if rs("iState")=1 then
	  	response.write "<font color=red>����ʹ��</font>"
	  else
	  	response.write "����ֹ"
	  end if
	  %></td>
            <td  align="center"><%
		response.write "<a href='admin_tags.asp?Action=Modify&TagID=" & rs("TagID") & "'>�޸�</a>&nbsp;"
		if rs("iState")=1 then
			response.write "<a href='admin_tags.asp?Action=Lock&TagID=" & rs("TagID") & "'>����</a>&nbsp;"
		else
            response.write "<a href='admin_tags.asp?Action=UnLock&TagID=" & rs("TagID") & "'>����</a>&nbsp;"
		end if        
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
              ѡ�б�ҳ��ʾ������TAG</td>
            <td> <strong>������</strong> 
              <input name="Action" type="radio" value="batchlock" checked onClick="document.myform.tagNames.disabled=true;document.myform.tagIds.disabled=true">
              ����&nbsp;&nbsp;&nbsp;&nbsp; 
              <input name="Action" type="radio" value="merge" onClick="document.myform.tagNames.disabled=false;document.myform.tagIds.disabled=false">�ϲ�Ϊ
              <input type="text" name="tagNames" id="tagNames" disabled>&nbsp;&nbsp;�ϲ����ID:<input type="text" name="tagIds" id="tagIds" size=10 disabled>
              &nbsp;&nbsp; 
              <input type="submit" name="Submit" value=" ִ �� "> </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub


sub Modify()
	dim TagID
	dim rst,sSql
	TagID=trim(request("TagID"))
	if TagID="" then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		TagID=Clng(TagID)
	end if
	Set  rst=Server.CreateObject("Adodb.RecordSet")
	sSql="select * from oblog_Tags where TagID=" & TagID
	if not IsObject(conn) then link_database
	rst.Open sSql,Conn,1,3
	if rst.bof and rst.eof then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ���� TAG ��</li>"
		rst.close
		Set  rst=nothing
		exit sub
	end if
%>
<FORM name="Form1" action="admin_tags.asp" method="post">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <TR class='title'> 
      <TD height=22 colSpan=2 align="center"><b><font color="#FFFFFF">�޸�ע�� TAG ��Ϣ</font></b></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="40%"> TAG ����</TD>
      <TD width="60%"><input type="text" name="name" value="<%=rst("Name")%>" size=50></TD>
      <input type="hidden" value="<%=rst("Tagid")%>"  name="TagID">
    </TR>
   
    <TR class="tdbg" > 
      <TD width="40%"> TAG ״̬��</TD>
      <TD width="60%"><input type="radio" name="iState" value=1 <%if rst("iState")=1 then response.write "checked"%>>
        ����&nbsp;&nbsp; <input type="radio" name="iState" value=0 <%if rst("iState")=0 then response.write "checked"%>>
        ����</TD>
    </TR>
    <TR class="tdbg" > 
      <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"></TD>
    </TR>
  </TABLE>
</form>
<%
	rst.close
	Set  rst=nothing
end sub

Sub BatchLock()
	Dim sID
	sId=Request.Form("TagId")
	conn.Execute("Update oblog_Tags Set iState=0 Where TagId In (" & sID & ")")
	oblog.showok "���������ɹ�!",""
End Sub

Sub MergeTags()
	Dim sIDs,sTargetId,sTargetName,aTags,i,sIDs0, rst,rst1,sSql,sTags,sTagsId,j
	sIDs=Trim(Request.Form("TagId"))
	sTargetName=Trim(Request.Form("tagNames"))
	sTargetId=Trim(Request.Form("tagIds"))
	If sIds="" Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	End If
	sIDs=Replace(sIDs," ","")
	aTags=Split(sIDs,",")
	sIDs0=sIDs
	If Right(sIDs,1)<>"," Then sIDs=sIDs & ","
	If Left(sIDs,1)<>"," Then sIDs= "," & sIDs 	
	If Instr(1,sIDs,"," & sTargetId & ",",1)<=0 Then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>Ŀ��ID����ȷ�����������" & Replace(sIDs,","," ")&"֮��</li>"
		exit sub
	End If		
	'���ȸ�oblog_Tags������
	Call conn.Execute("Update oblog_tags Set Name='" &sTargetName & "' Where TagId=" & sTargetId )
	'�滻����ID
	sIDs=Replace(sIDs,"," & sTargetId & ",","")	
	If Left(sIDs,1)="," Then sIDs=Right(sIDs,Len(sIDs)-1)
	If Right(sIDs,1)="," Then sIDs=Left(sIDs,Len(sIDs)-1)
	'��ɾ����������,ע��SQL SERVER	
	Call conn.Execute("Delete * From  oblog_tags  Where TagId IN (" & sIDs & ")" )
	'�����Ѿ�ʹ�õ��û�TAG����¼��Щ���µ�ID��Ȼ���������һ���µ�TAG ID����
	'��ȡΨһ����
	Set rst=Server.CreateObject("Adodb.Recordset")
	sSql="Select a.logId,a.UserId From (Select Userid,logId From oblog_Usertags Where TagId In (" & sIDs0 & ")) a Group by a.logId,a.UserId" 
	Response.Write sSql
	rst.Open sSql,conn,1,3
	'���½���ϵͳ����
	Call conn.Execute("Update oblog_tags Set iNum=" & rst.Recordcount & " Where TagId=" & sTargetId)
	'�����û�TAG������������
	Call conn.Execute("Delete * From oblog_Usertags Where TagId In (" & sIDs0 & ")")
	'�������ݲ���
	Do While Not rst.Eof
		Call conn.Execute("Insert Into oblog_UserTags(tagid,userid,logid) Values(" & sTargetId &"," & rst("userid")& "," & rst("logid") & ")")
		'���������������Tag
		Set rst1=conn.Execute("Select b.* From oblog_UserTags a ,oblog_Tags b Where a.tagId=b.tagId And logid=" & rst("logid"))
		j=0
		sTags=""
		sTagsId=""
		'���TAG�ִ���ID�ִ�
		Do While Not rst1.Eof
			j=j+1
			If j=1 Then 
				sTags=rst1("Name")
				sTagsId=rst1("TagId")
			Else
				sTags= sTags & P_TAGS_SPLIT & rst1("Name")
				sTagsId= sTagsId & P_TAGS_SPLIT & rst1("tagId")
			End if
			rst1.MoveNext
		Loop	
		'���¹ؼ����ִ�	
		Call conn.Execute("Update oblog_log Set logtags='" & sTags &"',logtagsid='" & sTagsId & "' Where logId=" & rst("logid"))
		'�������ɾ�̬ҳ�棿
		rst.Movenext
	Loop
	rst.close
	Set rst=Nothing
	Set rst1=Nothing
	oblog.showok "TAG�ϲ��ɹ�!",""
End Sub

sub SaveModify()
	dim sID,sName,sState,rst,sSql	
	sName=trim(request.Form("Name"))
	sID=trim(request.Form("TagID"))
	sState=trim(request.Form("iState"))
	if sID="" Or Not IsNumeric(sId) then
		FoundErr=true
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		sID=Clng(sID)
	end if

	if founderr=true then
		exit sub
	end if
	
	Set  rst=Server.CreateObject("Adodb.RecordSet")
	sSql="select * from oblog_Tags where TagID=" & sID
	if not IsObject(conn) then link_database
	rst.Open sSql,Conn,1,3

	rst("Name")=sName
	rst("iState")=sState
	rst.update
	rst.Close
	Set  rst=nothing
	oblog.showok "�޸ĳɹ�!",""
end sub

sub MoveUser()
	
	Conn.Execute sql
	response.Redirect "admin_tags.asp"
	'call WriteSuccessMsg(msg)
end sub

sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charSet =gb2312'>" & vbcrlf
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