<!--#include file="inc/inc_sys.asp"-->
<%
'----------------------------------------------
'�ϲ�ԭ����
'admin_regtext.asp/admin_userplacard.asp/admin_placard.asp/admin_friendsite.asp
'�ĸ�ֻ����oblog_setup���е�һ�ֶε��ļ�
'do-1:regtext;2:userolacard;3:placard;4:friendsite
'save-1/2/3/4
'URL�в�����ֱ�Ӵ����ֶ����Ƶ���������
'----------------------------------------------
Dim Action,ActionId,ActionText,ActionField
Dim rs,strNote,strField
Action=LCase(Request.QueryString("Action"))
ActionId=Right(Action,1)
Action=Left(Action,Len(Action)-1)
Select Case ActionId
		Case "1"
			ActionText="�޸���������(htm����)"
			ActionField="site_friends"
		Case "2"
			ActionText="�޸���վ����(htm����)"
			ActionField="site_placard"
		Case "3"
			ActionText="�޸�ע������(htm����)"
			ActionField="reg_text"
		Case "4"
			ActionText="�޸��û������̨֪ͨ(htm����)"
			ActionField="user_placard"
		Case Else
			Response.Write "����Ĳ�����ʽ��"
			Response.End
End Select

if Action="saveconfig" then
	strNote=request("note")
	if not IsObject(conn) then link_database
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select "& ActionField &" from oblog_setup",conn,1,3
	rs(0)=strNote
	rs.update
	rs.close
	oblog.reloadsetup
	Set oBlog=Nothing
	response.Redirect("admin_note.asp?action=do" & ActionId)
else
	set rs=oblog.execute("select "& ActionField &" from oblog_setup")
%>
<html>
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css"> 
<body class="tdbg" style="background:#ffffff;">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr> 
    <td height="25" class="title"><div align="center"><strong><font color="#FFFFFF"><%=ActionText%></font></strong></div></td>
  </tr>
  <tr> 
    <td><form name="form1" method="post" action="admin_note.asp?Action=saveconfig<%=ActionId%>">
	        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><div align="center"> 
                <textarea name="note" cols="100" rows="25" id="edit"><%=rs(0)%></textarea>
				<br>
                <br>
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
end if
%>

