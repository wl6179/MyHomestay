<!--#include file="inc/inc_syssite.asp"-->
<html>
<head>
<title>MyHomestay后台管理</title>
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
	response.Write("无权限")
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
		ActionText="<font color=#038ad7>修改友情连接(htm代码)</font>"
		ActionField="site_friends"
	Case "2"
		ActionText="<font color=#038ad7>修改网站公告——日推荐Homestay家庭</font>"
		ActionField="site_placard"
	Case "3"
		ActionText="<font color=#038ad7>修改 中国家庭注册条款(htm代码)</font>"
		ActionField="reg_text"
	Case "4"
		ActionText="<font color=#038ad7>修改用户管理后台通知(htm代码)</font>"
		ActionField="user_placard"
	
	Case "5"'WL;------------------------------------------
		ActionText="<font color=#038ad7>修改 加盟合作注册条款(htm代码)</font>"'WL;
		ActionField="reg_text2"
	Case "6"'WL;------------------------------------------
		ActionText="<font color=#038ad7>修改网站公告——今日推荐外教</font>"'WL;
		ActionField="user_placard2"
		
	Case "86"'WL;----------------
		ActionText="<font color=#038ad7>用户注册成功消息</font>"'WL;
		ActionField="reg_success_text"
	Case "87"'WL;----------------
		ActionText="<font color=#038ad7>修改——关于我们内容</font>"'WL;
		ActionField="myhomestay_about_us"
	Case "88"'WL;----------------
		ActionText="<font color=#038ad7>修改——联系我们内容</font>"'WL;
		ActionField="myhomestay_contact_us"
	Case "89"'WL;----------------
		ActionText="<font color=#038ad7>修改——合作代理内容</font>"'WL;
		ActionField="myhomestay_cooperate_info"
	Case "90"'WL;----------------
		ActionText="<font color=#038ad7>修改——帮助中心内容</font>"'WL;
		ActionField="myhomestay_help"
	Case "91"'WL;----------------
		ActionText="<font color=#038ad7>修改——特别提示内容（比如 IE7的问题）</font>"'WL;
		ActionField="myhomestay_especially"
		
	Case "92"'WL;----------------
		ActionText="<font color=#038ad7>修改——推荐保护软件（比如 世界之窗浏览器+安全卫士360等等...）</font>"'WL;
		ActionField="myhomestay_recommend"
end select
select case request("Actionsave")
	case "saveskin" 
		call saveskin()
	case "savemodi"
		call savemodi()'wl；
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
	rs.open "select reg_text,reg_text2 from oblog_setup",conn,1,3'WL;修改注册条款(htm代码)
	strNote = rs(0)'WL;只保存注册条款(htm代码)
	strNote2 = rs(1)'WL;加盟合作的注册条款(htm代码)
	rs.close
	set rs=nothing
	oblog.reloadsetup
	Dim dirstr
	dirstr= Server.MapPath("ad/")
	RESPONSE.Write(dirstr&"\item.htm(W.L)")'WL;
	Call oblog.BuildFile(dirstr&"\item.htm",strNote)'WL;只保存注册条款(htm代码)
	Call oblog.BuildFile(dirstr&"\item2.htm",strNote2)'WL;
	oblog.showok "保存成功",""
end sub

sub saveskin()
	dim rs,sql,table
	if trim(request("skinname"))="" then response.Write("模版名不能为空"):response.End()
	if trim(request("skintext"))="" then response.Write("模版内容不能为空"):response.End()
	if skintype="user" then 
		table="oblog_userskin"
	elseif skintype="sys" then 
		table="oblog_sysskin"
	else
		response.Write("参数错误")
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
	oblog.showok "保存成功",""
end Sub
sub modiskin()
	dim rs,table
	if skintype="user" then 
		table="oblog_userskin"
	elseif skintype="sys" then 
		table="oblog_sysskin"
	else
		response.Write("参数错误")
		response.end
	end if	
	set rs=oblog.execute("select * from "&table&" where id="&clng(request.QueryString("id")))
	if rs.eof then
		response.write("无记录")
		response.End()
	end if
%>
<body class="bgcolor" onLoad="oblog_Size(500);" style="background:#ffffff;scrollbar-face-color:#ffffff;scrollbar-highlight-color: #cccccc;scrollbar-shadow-color: #cccccc;scrollbar-3d-light-color: #ffffff;scrollbar-arrow-color : #cccccc;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;">
<form method="POST" action="" id="oblogform" name="oblogform" onSubmit="submits();">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" >
    <tr> 
      <td width="769" height="22"><strong>修改模版</strong></td>
    </tr>
    <tr>
      <td height="25">模版名称： 
        <input name="skinname" type="text" id="skinname" value=<%if skintype="sys" then response.Write rs("sysskinname") else response.Write rs("userskinname")%>>
        　　作者：
        <input name="skinauthor" type="text" id="skinauthor" value=<%=rs("skinauthor")%>>
<%if skintype="user" then%>
<br>
        模板作者连接： 
        <input name="skinauthorurl" type="text" id="skinauthorurl" size="50" value="<%=rs("skinauthorurl")%>">
         <br>
        模版预览图片： 
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
        <input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " > 
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
		response.write("错误的参数")
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
end sub
%>
</body>
<script language="javascript">oblog_Size(620);</script>
</html>