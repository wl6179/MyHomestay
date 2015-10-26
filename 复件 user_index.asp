<!--#include file="user_top.asp"-->
<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
		&#8226;<a href="user_help.asp">用户管理后台帮助</a>
		&#8226;<a href="user_setting.asp?action=placard">修改公告</a>
		&#8226;<a href="user_setting.asp?action=links">友情连接管理</a>
		&#8226;<a href="user_setting.asp?action=blogstar">申请用户明星</a>
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			管理首页
	</div>
	
    <div class="content_body"><%main()%>
<!-- 在FIREFOX下背景正常显示 -->
	<br style="clear: both;" />
<!-- 结束 -->
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>  
  <div id="bottom"><%=oblog.user_copyright%></div>
</body>
</html>
<%
sub main
dim rs
%>
<div class="msg_left">
<div class="msg">
	<h1>我的未读短信</h1>
	<ul>
<%
set rs=oblog.execute("select id,sender,topic,addtime,content from oblog_pm where incept='"&oblog.logined_uname&"' and delr=0 and isreaded=0 order by id desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&oblog.filt_html(rs("sender"))&"：</span><a href=""javascript:openScript('user_pm.asp?action=read&id="&rs("id")&"',450,380)"" title="""&oblog.filt_html(rs("content"))&""">"&oblog.filt_html(rs("topic"))&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
%>
	</ul>
</div>
<div class="msg">
	<h1>我的最近10条评论</h1>
	<ul>
<%
set rs=oblog.execute("select top 10 commenttopic,addtime,commentid,comment_user,mainid,comment from oblog_comment where userid="&oblog.logined_uid&" order by commentid desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&rs("comment_user")&"：</span><a href='go.asp?logid="&rs("mainid")&"&commentid="&rs("commentid")&"' target='_blank' title="""&oblog.filt_html(left(rs("comment"),150))&""">"&rs("commenttopic")&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
%>
	</ul>
</div>
<div class="msg">
	<h1>我的最近10条留言</h1>
	<ul>
<%
set rs=oblog.execute("select top 10 messagetopic,addtime,messageid,message_user,message from oblog_message where userid="&oblog.logined_uid&" order by messageid desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&rs("message_user")&"：</span><a href='go.asp?messageid="&rs("messageid")&"' target='_blank' title="""&oblog.filt_html(left(rs("message"),150))&""">"&rs("messagetopic")&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
set rs=nothing
%>
	</ul>
</div>
</div>
<div class="msg_right">
<div class="msg">
	<h1>系统通知</h1>
	<ul class="list_edit"><%=oblog.setup(20,0)%></ul>
</div>
</div>

<%end sub%>