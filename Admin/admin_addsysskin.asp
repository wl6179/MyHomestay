<!--#include file="inc/inc_sys.asp"-->
<html>
<head>
<title>ϵͳģ�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="tdbg">
<SCRIPT language=javascript>

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

function VerifySubmit1()
{
  	topic = del_space(document.form2.sysskinname.value);
     if (topic.length == 0)
     {
        alert("ģ�����Ʋ���Ϊ��!");
	return false;
     }
	 submits();
 	if (document.form2.edit.value == "")
     {
        alert("ģ�治��Ϊ��!");
	return false;
     }		  	 
	return true;
}

function VerifySubmit2()
{
  	  submits();
 	if (document.form2.edit.value == "")
     {
        alert("ģ�治��Ϊ��!");
	return false;
     }	  	 
	return true;
}

</SCRIPT>
<%
dim action,add
action=trim(request("action"))
add=trim(request.QueryString("add"))
if add="next" then
	call addskin2()
elseif add="first" then
	call addskin1()
end if

if action="saveconfig1" then
	call saveconfig1()
elseif action="saveconfig2" then
	call saveconfig2()
end if	

sub saveconfig1()
	dim rs,sql
	dim temp,addid
	if trim(request("sysskinname"))="" then oblog.sys_err("ģ��������Ϊ��"):response.End()
	if trim(request("skinmain"))="" then oblog.sys_err("��ģ�岻��Ϊ��"):response.End()
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_sysskin"
	rs.open sql,conn,1,3
	rs.addnew
    rs("sysskinname")=trim(request("sysskinname"))
	rs("skinauthor")=trim(request("skinauthor"))
	rs("skinmain")=request("skinmain")
	'rs("skinshowlog")=request("skinshowlog")
	rs.update
	set rs=oblog.execute("select max(id) from oblog_sysskin")
	addid=rs(0)  
	rs.close
	set rs=nothing
	'call closeconn
	response.Redirect "admin_addsysskin.asp?add=next&id="&addid
end sub
sub saveconfig2()
	dim rs,sql,id
	id=request("id")
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from oblog_sysskin where id="&id
	rs.open sql,conn,1,3
	rs("skinshowlog")=request("skinshowlog")
	rs.update
	rs.close
	set rs=nothing
	'call closeconn
	response.Write "���ģ��ɹ���"
end sub

sub addskin1()%>
 <form method="POST" action="admin_addsysskin.asp" id="form2" name="form2" onSubmit="return VerifySubmit1()" >
  <p>&nbsp;</p>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border" >
    <tr class="topbg"> 
      
    <td height="25" align=center><strong>���ϵͳģ��(��һ��)</strong></td>
    </tr>
    <tr class="tdbg">
      <td height="30">ģ�����ƣ�
            <input name="sysskinname" type="text" id="sysskinname">
�������ߣ�
<input name="skinauthor" type="text" id="skinauthor">��
<a href="admin_skin_help.asp" target="_blank"></a></td>
    </tr>
    <tr class="tdbg">
        <td height="30"><strong>�����ģ��(<a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>)</strong></td>
    </tr>
    <tr class="tdbg">
        <td height="30"><input type="hidden" name="skinmain" id='edit' value="")>                
           <!--#include file="edit.asp"--></td>
    </tr>
    <tr class="tdbg">
      <td height="30"><div align="center">
        <INPUT type="hidden" name="action" value="saveconfig1">
        <input name="cmdSave" type="submit" id="cmdSave2" value=" ������ģ��(��һ��) " >
      </div></td>
    </tr>
   </table>
   </form>
   <%end sub
   
   sub addskin2()%>
	<br>
	 <form method="POST" action="admin_addsysskin.asp" id="form2" name="form2" onSubmit="return VerifySubmit2()">
	<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
    <tr >
      <td height="25" class="topbg"><div align="center" ><strong>���ϵͳģ��(�ڶ���)</strong></div></td>
    </tr>
    <tr class="tdbg">
        <td height="30"><strong>��Ӹ�ģ��(<a href="admin_skin_help.asp" target="_blank"><strong>ģ���ǰ���</strong></a>)</strong></td>
    </tr>
    <tr class="tdbg">
        <td height="30"><INPUT type="hidden" name="skinshowlog" id='edit' value="">
                 <!--#include file="edit.asp"--></td>
    </tr>
    <tr class="tdbg"> 
      
    <td height="30"><div align="center">
          <INPUT type="hidden" name="action" value="saveconfig2">
		  <INPUT type="hidden" name="id" value="<%=request.QueryString("id")%>">��
          <input name="cmdSave" type="submit" id="cmdSave2" value=" ����ģ��(���) " >
    </div></td>
    </tr>
  </table>
</form>
<%end sub%>
</html>
