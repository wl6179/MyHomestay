<!--#include file="inc/inc_syssite.asp"-->
<html>
<head>
<title>MyHomestay��̨����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
BODY
{
    BACKGROUND: #ffffff;
	FONT: 14px "Arial", "Helvetica", "sans-serif";
}
table{
	FONT: 14px "Arial", "Helvetica", "sans-serif";
}
</style>
</head>
<%
if session("adminname")="" then 
	response.Write("��Ȩ��")
	response.end
end if
dim action,skintype,skinorder,t,ActionText,ActionField,actionid
t=0
Action=trim(request("Action"))
skintype=trim(request("skintype"))
skinorder=trim(request("skinorder"))
actionid=trim(request("do"))
select case actionid
	Case "1"
		ActionText="<font color=#038ad7>�޸���������(htm����)</font>"
		ActionField="site_friends"
	Case "2"
		ActionText="<font color=#038ad7>�޸���վ���桪�����Ƽ�Homestay��ͥ</font>"
		ActionField="site_placard"
	Case "3"
		ActionText="<font color=#038ad7>�޸� �й���ͥע������(htm����)</font>"
		ActionField="reg_text"
	Case "4"
		ActionText="<font color=#038ad7>�޸��û������̨֪ͨ(htm����)</font>"
		ActionField="user_placard"
	
	Case "5"'WL;------------------------------------------
		ActionText="<font color=#038ad7>�޸� ���˺���ע������(htm����)</font>"'WL;
		ActionField="reg_text2"
	Case "6"'WL;------------------------------------------
		ActionText="<font color=#038ad7>�޸���վ���桪�������Ƽ����</font>"'WL;
		ActionField="user_placard2"
		
	Case "86"'WL;----------------
		ActionText="<font color=#038ad7>�û�ע��ɹ���Ϣ</font>"'WL;
		ActionField="reg_success_text"
	Case "87"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ���������������</font>"'WL;
		ActionField="myhomestay_about_us"
	Case "88"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ�����ϵ��������</font>"'WL;
		ActionField="myhomestay_contact_us"
	Case "89"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ���������������</font>"'WL;
		ActionField="myhomestay_cooperate_info"
	Case "90"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ���������������</font>"'WL;
		ActionField="myhomestay_help"
	Case "91"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ����ر���ʾ���ݣ����� IE7�����⣩</font>"'WL;
		ActionField="myhomestay_especially"
		
	Case "92"'WL;----------------
		ActionText="<font color=#038ad7>�޸ġ����Ƽ�������������� ����֮�������+��ȫ��ʿ360�ȵ�...��</font>"'WL;
		ActionField="myhomestay_recommend"
end select
select case request("Actionsave")
	case "saveskin" 
		call saveskin()
	case "savemodi"
		call savemodi()'wl��
end select

select case Action
	case "modiskin"
		call modiskin()
	Case Else
		call modi()
end select

sub savemodi()
	dim rs,strNote
	strNote=request("note")
	if not IsObject(conn) then link_database
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select "& ActionField &" from oblog_setup",conn,1,3
	rs(0)=strNote
	rs.update
	rs.close
	
	Dim strNote2
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select reg_text,reg_text2 from oblog_setup",conn,1,3'WL;�޸�ע������(htm����)
	strNote = rs(0)'WL;ֻ����ע������(htm����)
	strNote2 = rs(1)'WL;���˺�����ע������(htm����)
	rs.close
	set rs=nothing
	oblog.reloadsetup
	Dim dirstr
	dirstr= Server.MapPath("ad/")
	RESPONSE.Write(dirstr&"\item.htm(W.L)")'WL;
	Call oblog.BuildFile(dirstr&"\item.htm",strNote)'WL;ֻ����ע������(htm����)
	Call oblog.BuildFile(dirstr&"\item2.htm",strNote2)'WL;
	oblog.showok "����ɹ�",""
end sub

sub saveskin()
	dim rs,sql,table
	if trim(request("skinname"))="" then response.Write("ģ��������Ϊ��"):response.End()
	if trim(request("skintext"))="" then response.Write("ģ�����ݲ���Ϊ��"):response.End()
	if skintype="user" then 
		table="oblog_userskin"
	elseif skintype="sys" then 
		table="oblog_sysskin"
	else
		response.Write("��������")
		response.end
	end if	
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from "&table&" where id="&clng(request.QueryString("id"))
	if not IsObject(conn) then link_database
	rs.open sql,conn,1,3
	if skintype="sys" then
		rs("sysskinname")=trim(request("skinname"))
	else
		rs("userskinname")=trim(request("skinname"))
		rs("skinpic")=trim(request("skinpic"))
		rs("skinauthorurl")=trim(request("skinauthorurl"))
	end if
	rs("skinauthor")=trim(request("skinauthor"))
	if skinorder="0" then
		rs("skinmain")=request("skintext")
	else
		rs("skinshowlog")=request("skintext")
	end if
	rs.update
	rs.close
	set rs=nothing
	oblog.reloadsetup
	Application(oblog.cache_name&"index_update")=true
	oblog.showok "����ɹ�",""
end Sub
sub modiskin()
	dim rs,table
	if skintype="user" then 
		table="oblog_userskin"
	elseif skintype="sys" then 
		table="oblog_sysskin"
	else
		response.Write("��������")
		response.end
	end if	
	set rs=oblog.execute("select * from "&table&" where id="&clng(request.QueryString("id")))
	if rs.eof then
		response.write("�޼�¼")
		response.End()
	end if
%>
<body class="bgcolor" onLoad="oblog_Size(500);" style="background:#ffffff;scrollbar-face-color:#ffffff;scrollbar-highlight-color: #cccccc;scrollbar-shadow-color: #cccccc;scrollbar-3d-light-color: #ffffff;scrollbar-arrow-color : #cccccc;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;">
<form method="POST" action="" id="oblogform" name="oblogform" onSubmit="submits();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" >
    <tr> 
      <td width="769" height="22"><strong>�޸�ģ��</strong></td>
    </tr>
    <tr>
      <td height="25">ģ�����ƣ� 
        <input name="skinname" type="text" id="skinname" value=<%if skintype="sys" then response.Write rs("sysskinname") else response.Write rs("userskinname")%>>
        �������ߣ�
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
<%if skintype="user" then%>
<br>
        ģ���������ӣ� 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        ģ��Ԥ��ͼƬ�� 
        <input name="skinpic" type="text" id="skinpic" size="50" value="<%=rs("skinpic")%>">
<%end if%>
		</td>
    </tr>
    <tr> 
      <td ><INPUT type="hidden" id='edit' name="skintext" value="<%
if skinorder="0" then
	if rs("skinmain")<>"" then response.write Server.HtmlEncode(rs("skinmain"))
else
	if rs("skinshowlog")<>"" then response.write Server.HtmlEncode(rs("skinshowlog"))
end if
%>">
<!--#include file="edit.asp"--></td>
    </tr>
    <tr> 
      <td> <div align="center">
        <input name="Actionsave" type="hidden" id="Action" value="saveskin"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " > 
      </div></td>
    </tr>
  </table>

</form>
<%
set rs=nothing
end sub

sub modi()
	dim rs
	if ActionField="" then
		response.write("����Ĳ���")
		response.end
	end if
	set rs=oblog.execute("select "& ActionField &" from oblog_setup")
%>
<body class="tdbg" style="background:#ffffff;scrollbar-face-color:#ffffff;scrollbar-highlight-color: #cccccc;scrollbar-shadow-color: #cccccc;scrollbar-3d-light-color: #ffffff;scrollbar-arrow-color : #cccccc;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr> 
    <td height="25" class="title"><div align="center"><strong><font color="#FFFFFF"><%=ActionText%></font></strong></div></td>
  </tr>
  <tr> 
    <td><form name="oblogform" method="post" onSubmit="submits();">
	        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><div align="center"> 
                <input name="note" id="edit" type="hidden" value="<%if rs(0)<>"" then response.Write(Server.HtmlEncode(rs(0)))%>"><!--#include file="edit.asp"-->
				<br>
                <br>
				<input name="Actionsave" type="hidden" id="Action" value="savemodi"> 
                <input type="submit" name="Submit" value="�ύ�޸�">
              </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
<%
set rs=nothing
end sub
%>
</body>
<script language="javascript">oblog_Size(620);</script>
</html>