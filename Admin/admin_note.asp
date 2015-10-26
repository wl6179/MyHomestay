<!--#include file="inc/inc_sys.asp"-->
<%
'----------------------------------------------
'合并原来的
'admin_regtext.asp/admin_userplacard.asp/admin_placard.asp/admin_friendsite.asp
'四个只操作oblog_setup表中单一字段的文件
'do-1:regtext;2:userolacard;3:placard;4:friendsite
'save-1/2/3/4
'URL中不允许直接传递字段名称等敏感数据
'----------------------------------------------
Dim Action,ActionId,ActionText,ActionField
Dim rs,strNote,strField
Action=LCase(Request.QueryString("Action"))
ActionId=Right(Action,1)
Action=Left(Action,Len(Action)-1)
Select Case ActionId
		Case "1"
			ActionText="修改友情连接(htm代码)"
			ActionField="site_friends"
		Case "2"
			ActionText="修改网站公告(htm代码)"
			ActionField="site_placard"
		Case "3"
			ActionText="修改注册条款(htm代码)"
			ActionField="reg_text"
		Case "4"
			ActionText="修改用户管理后台通知(htm代码)"
			ActionField="user_placard"
		Case Else
			Response.Write "错误的操作方式！"
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
                <input type="submit" name="Submit" value="提交修改">
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

