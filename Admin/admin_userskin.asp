<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<title>�û�ģ�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.form2.chkAll.checked){
	document.form2.chkAll.checked = document.form2.chkAll.checked&0;
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
<%
dim action
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages,strGuide,ispass
Action=trim(request("Action"))

select case Action
	case "saveconfig" 
		call saveconfig()
	case "showskin"
		call showskin()
	case "modiskin"
		call modiskin()
	case "savedefault"
		call savedefault()
	case "delconfig"
		call delconfig()
	case "addskin"
		call addskin()
	case "saveaddskin"
		call saveaddskin()
	case "passskin"
		call passskin
	case "unpassskin"
		call unpassskin	
end select

sub showskin()
	dim rs,psql,sql,strFileName
	ispass=clng(request("ispass"))
	if ispass=1 then
	    strFileName="admin_userskin.asp?action=showskin&ispass=1"
		psql=" where ispass=1 "
	else
	    strFileName="admin_userskin.asp?action=showskin&ispass=0"
		psql=" where ispass=0 "
	end if
	if request("page")<>"" then
	    currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select id,userskinname,skinauthor,skinauthorurl,isdefault,ispass,skinpic from oblog_userskin "&psql&" order by id desc "
		rs.Open sql,Conn,1,1
	  	if rs.eof and rs.bof then
            showContent(rs)
			strGuide=strGuide & " (����0��ģ��)</h1>"
			response.write "<div align='right'>"&strGuide&"</div>"
		else
	    	totalPut=rs.recordcount
			strGuide=strGuide & " (����" & totalPut & "��ģ��)</h1>"
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
	        	Call showContent(rs)
	        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"��ģ��")
	   	 	else
	   	     	if (currentPage-1)*MaxPerPage<totalPut then
	         	   	rs.move  (currentPage-1)*MaxPerPage
	         		dim bookmark
	           		bookmark=rs.bookmark            	
	        	else
		        	currentPage=1           		           		
		    	end if
		    	Call showContent(rs)
		    	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"��ģ��")
			end if
		end if
		rs.Close
		set rs=Nothing
end sub

sub showContent(rs)
	dim i 
	i=0
	
%>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
  <tr align="center">
    <td  height=25 class="topbg" align="left"><strong>�û�ģ���������<a href="admin_userskin.asp?action=showskin&ispass=1">&gt;&gt;ͨ����˵�ģ��</a>����<a href="admin_userskin.asp?action=showskin&ispass=0">&gt;&gt;δͨ����˵�ģ��</a></strong>
  </tr>
</table>
<form name="form2" method="post" action="admin_userskin.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr class="topbg"> 
      <td height="25" colspan="6" ><strong><%if ispass=1 then response.Write "ͨ����˵�ģ��" else response.write "δͨ����˵�ģ��"%></strong></td>
    </tr>
    <tr class="topbg"> 
      <td width="10%" height="25" > <div align="center">ID</div></td>
      <td width="20%" > <div align="center">����<b>����ɫΪĬ��ģ�壩</b></div></td>
      <td width="20%" ><div align="center">����</div></td>
	  <td width="10%" > <div align="center">���</div></td>
      <td width="10%" > <div align="center">ѡ��</div></td>
      <td width="40%" > <div align="center">ģ�����</div></td>
    </tr>
    <% 
do while not rs.eof	  
dim userskinname
    userskinname=rs("userskinname")
%>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="10%"> <div align="center"><%= rs("id") %>&nbsp;</div></td>
      <td width="20%" > <div align="center">
	  <%if rs("isdefault")=1 then
	  response.Write "<font color='red'>"&userskinname&"</red>"
	  else 
	  response.Write userskinname
	  end if
	  %>
	  </div></td>
      <td width="20%" ><div align="center">
	  <%if rs("skinauthorurl")="" or isnull(rs("skinauthorurl")) then
	  response.Write rs("skinauthor") 
	  else
	  response.Write "<a href="""&oblog.filt_html(rs("skinauthorurl"))&""" target='_blank'>"&rs("skinauthor")&"</a>" 
	  end if%>
	  </div></td>
	  <td width="10%" > <div align="center"><%if rs("ispass")=1 then response.Write("�����") else response.Write("δ���")%></div></td>
      <td width="10%"> <div align="center"> 
          <input name="checkbox" type="checkbox" onClick="unselectall()" id= "checkbox" class="tdbg" value='<%=rs("id")%>'>
        </div></td>
      <td width="40%"> <div align="left">
	<a href="../showskin.asp?id=<%=rs("id")%>" target="_blank">Ԥ��</a> 
	<%if ispass=0 then%>
	<a href="admin_userskin.asp?action=passskin&id=<%=rs("id")%>">ͨ�����</a> 
	<%else%>
	<a href="admin_userskin.asp?action=unpassskin&id=<%=rs("id")%>">ȡ�����</a> 
	<%end if%>
  <a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=0&id=<%=rs("id")%>"  target="_blank">�޸���ģ��</a>
  <a href="../admin_edit.asp?action=modiskin&skintype=user&t=0&editm=1&skinorder=1&id=<%=rs("id")%>"  target="_blank">�޸ĸ�ģ��</a><br>
	<a href="admin_userskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��(�ı���ʽ)</a> 
	��<a href="admin_userskin.asp?action=delconfig&id=<%=rs("id")%>" onclick=return(confirm("ȷ��Ҫɾ�����ģ����"))>ɾ��ģ��</a></div></td>
    </tr>
    <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>

    <tr> 
      <td height="40" colspan="6" align="center" class="tdbg"> <div align="center"> 
	  <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
	  ȫѡ
	 <input type="radio" value="savedefault" name="action" checked>Ĭ��ģ��</option>
	 <%if ispass=0 then%>
	  <input type="radio" value="passskin" name="action" >ͨ�����</option>
	  <%else%>
	  <input type="radio" value="unpassskin" name="action">ȡ�����</option>
	  <%end if%>
	   <input type="radio" value="delconfig" name="action" >ɾ��</option>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
		  <input type="submit" name="Submit" value="��������">
        </div></td>
    </tr>
  </table>
</form>
<%
end sub

sub savedefault()
	dim isdefaultID
	isdefaultID=trim(request("checkbox"))
		if instr(isdefaultID,",")>0 then
		Response.Write("<script language=javascript>alert('�û�Ĭ��ģ��ֻ����ѡ��һ����');history.back();</script>")
		Response.End()
	elseif isdefaultID="" then
		Response.Write("<script language=javascript>alert('��ָ��Ҫ�趨ΪĬ�ϵ�ģ�壡');history.back();</script>")
		Response.End()
		exit sub
		end if
	oblog.execute("update oblog_userskin set isdefault=0")
	oblog.execute("update oblog_userskin set isdefault=1 where id="&isdefaultID)
	Response.Write"<script language=JavaScript>"
	Response.Write"alert(""�޸ĳɹ���"");"
	Response.Write"window.history.go(-1);"
	Response.Write"</script>"
end sub

sub passskin()
	dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=1 where id in ("&id&")")
	elseif id="" then
	id=clng(request("id"))
	oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	else
    oblog.execute("update oblog_userskin set ispass=1 where id="&id)
	end if
	oblog.showok "ͨ����˳ɹ�",""
end sub

sub unpassskin()
	dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("update oblog_userskin set ispass=0 where id in ("&id&")")
	elseif id="" then
	id=clng(request("id"))
	oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	else
	oblog.execute("update oblog_userskin set ispass=0 where id="&id)
	end if
	oblog.showok "ȡ����˳ɹ�",""
end sub

sub saveconfig()
	dim rs,sql
	if trim(request("userskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("��ģ�治��Ϊ��"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("��ģ�治��Ϊ��"):response.End()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_userskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs("userskinname")=trim(request("userskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=request("skinmain")
	rs("skinshowlog")=request("skinshowlog")
	rs("skinpic")=trim(request("skinpic"))
	rs("skinauthorurl")=trim(request("skinauthorurl"))
	
	rs("skinmain_Nav")=request("skinmain_Nav")
	
	
	rs.update
	rs.close
	set rs=nothing
	oblog.showok "����ɹ�",""
end Sub
sub delconfig()
    dim id
	id=trim(request("checkbox"))
	if instr(id,",")>0 then
	id=replace(id," ","")
	oblog.execute("delete from oblog_userskin where id in ("&id&")")
	elseif id="" then
	id=clng(request.QueryString("id"))
	oblog.execute("delete from oblog_userskin where id="&id)
	else
	oblog.execute("delete from oblog_userskin where id="&id)
	end if
	oblog.showok "ɾ���ɹ�",""
end sub
sub modiconfig()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&clng(request.QueryString("id")))
	End Sub 
sub saveaddskin()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	if trim(request("userskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("��ģ�治��Ϊ��"):response.End()
	if trim(request("skinshowlog"))="" then oblog.sys_err("��ģ�治��Ϊ��"):response.End()
	sql="select * from oblog_userskin where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	rs.addnew
	rs("userskinname")=trim(request("userskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=trim(request("skinmain"))
	rs("skinshowlog")=trim(request("skinshowlog"))
	rs("skinpic")=trim(request("skinpic"))
	rs("skinauthorurl")=trim(request("skinauthorurl"))
	rs.update
	rs.close
	set rs=nothing
	response.Redirect "admin_userskin.asp?action=showskin"
end sub

sub modiskin()
	dim rs
	set rs=oblog.execute("select * from oblog_userskin where id="&clng(request.QueryString("id")))
%>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
    <tr class="topbg"> 
      
    <td height="22" colspan=2 align=center><strong>�޸��û�ģ��</strong></td>
    </tr>
    <tr class="tdbg"> 
      
    <td width="253" height="30"><strong>�����޸ĵ�ģ���ǣ�<%=rs("userskinname")%></strong></td>
      
    <td width="516" height="30">
	<a href="admin_userskin.asp?action=modiskin&id=<%=rs("id")%>">�޸�ģ��</a>����<a href="admin_userskin.asp?action=showskin&ispass=1">���ع���˵�</a>���� 
      <a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a></td>
    </tr>
</table>

<form method="POST" action="admin_userskin.asp?id=<%=clng(request.QueryString("id"))%>" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>�޸�ģ��</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ� 
        <input name="userskinname" type="text" id="userskinname" value=<%=rs("userskinname")%>>
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
        <br>
        ģ���������ӣ� 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        ģ��Ԥ��ͼƬ<strong>�� 
        <input name="skinpic" type="text" id="skinpic" size="50" value="<%=rs("skinpic")%>">
        </td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <strong>��ģ�棺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"><%if rs("skinmain")<>"" then response.Write Server.HtmlEncode(rs("skinmain")) else response.Write("")%></textarea>
        <br>
        <br>
        <strong>��ģ�棺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"><%if rs("skinshowlog")<>"" then response.Write Server.HtmlEncode(rs("skinshowlog")) else response.Write("")%></textarea>
        </strong>
		<br>
        <br>
        <strong>��Ŀ��ģ�棺 <br>
        <textarea cols="60" rows="20" name="skinmain_Nav" id="edit"><%=Server.HtmlEncode(rs("skinmain_Nav"))%></textarea>
        </strong>
		</td>
    </tr>
		
	 
    <tr> 
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveconfig"> <br />
        <input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " > &nbsp;<br />&nbsp;
      </div>
	  </td>
    </tr>
	
  </table>

</form>

<%
set rs=nothing
end sub

sub addskin()
%>
  
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border" >
  <tr class="topbg"> 
    <td height="22" align=center><strong>����û�ģ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td height="30"><div align="center"><a href="admin_userskin.asp?action=showskin"><strong>���ع���˵�</strong></a>���� <a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	 <a href="admin_adduserskin.asp?add=first"><strong>2.52ģʽ</strong></a>
	</div></td>
  </tr>
</table>

<form method="POST" action="admin_userskin.asp" id="form1" name="form1" >
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr> 
      <td width="769" height="22" class="topbg"><strong>ģ�����</strong></td>
    </tr>
    <tr>
      <td height="25" class="tdbg">ģ�����ƣ� 
        <input name="userskinname" type="text" id="userskinname">
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor">
        <br>
        ģ����������<strong>�� 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="">
        </strong> <br>
        ģ��Ԥ��ͼƬ<strong>�� 
        <input name="skinpic" type="text" id="skinpic" size="50">
        </strong> </td>
    </tr>
    <tr> 
      <td height="25" class="tdbg"> <strong>��ģ�棺</strong><br>
        <textarea name="skinmain" cols="100" rows="12" id="edit"></textarea>
        <br>
        <br>
        <strong>��ģ�棺 <br>
        <textarea name="skinshowlog" cols="100" rows="12" id="skinshowlog"></textarea>
        </strong></td>
    </tr>
    <tr> 
      <td class="tdbg"> <div align="center">
        <input name="Action" type="hidden" id="Action" value="saveaddskin"> 
          <input name="cmdadd" type="submit" id="cmdadd" value=" ��� " > 
      </div></td>
    </tr>
  </table>

</form>

<%
end sub
%>