
<script>
function menu(oblog){
if(oblog.style.display=='')oblog.style.display='none';
else oblog.style.display='';
}
</script>

<style>
	img {
	background:#ffffff;
	border:0px;
	border-collapse: collapse;
	font-size:13px;
	}
	.menu-t1 {
	/*background:#eeeeee;*/
	font-size:13px;
	cursor:hand;
margin:6px,3px,6px,0px;
	padding-top:6px;
	}
</style>
	<div id="toplogo"></div>
	<div id="sub">
	<a href="HomestayBackctrl-index.asp" target="_top">管理首页</a> | <a href="user_update.asp">更新站点</a>
	<br /><%=uurl%>| <a href="user_pmmanage.asp">短消息(<%=chkpm()%>)</a>
	</div>
	<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 文章管理</div>
	<ul id="oblog_1">
	<li><a href="user_post.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> 添加文章</a></li>
	<li><a href="user_blogmanage.asp?usersearch=5"><img src="oBlogStyle/OblogStyleAdminImages/li/12.gif" align="absmiddle" /> 草稿箱</a></li>
	<li><a href="user_blogmanage.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> 所有文章</a></li>
	<li><a href="user_blogmanage.asp?action=downlog"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 备份文章</a></li>
	<!--<li><a href="user_subject.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/16.gif" align="absmiddle" /> 文章分类管理</a></li>-->
	</ul>
	
	<div class="menu-t1" onClick="menu(oblog_3)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 评论留言</div>
	<ul id="oblog_3">
	<li><a href="user_comments.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/31.gif" align="absmiddle" /> 评论管理</a></li>
	<li><a href="user_messages.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/32.gif" align="absmiddle" /> 留言管理</a></li>
	</ul>
	
	<div class="menu-t1" onClick="menu(oblog_4)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 模版设置</div>
	<ul id="oblog_4" style="display: ;">
	<li><a href="user_template.asp?action=showconfig"><img src="oBlogStyle/OblogStyleAdminImages/li/41.gif" align="absmiddle" /> 选择模板</a></li>
	<li><a href="user_template.asp?action=modiconfig&amp;editm=1"><img src="oBlogStyle/OblogStyleAdminImages/li/42.gif" align="absmiddle" /> 修改主模板</a></li>
	<li><a href="user_template.asp?action=modiviceconfig&amp;editm=1"><img src="oBlogStyle/OblogStyleAdminImages/li/43.gif" align="absmiddle" /> 修改内容模板</a></li>
	<li><a href="user_template.asp?action=bakskin"><img src="oBlogStyle/OblogStyleAdminImages/li/44.gif" align="absmiddle" /> 备份模板</a></li>
	<!--<li><a href="user_template.asp?action=good"><img src="oBlogStyle/OblogStyleAdminImages/li/45.gif" align="absmiddle" /> 推荐我的模板</a></li>-->
	</ul>
	<%If EN_PHOTO=1 Then%>
	<div class="menu-t1" onClick="menu(oblog_2)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 相册管理</div>
	<ul id="oblog_2" style="display: ;">
	<li><a href="user_post.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/21.gif" align="absmiddle" /> 添加图片</a></li>
	<li><a href="user_blogmanage.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/22.gif" align="absmiddle" /> 所有图片</a></li>
	<!--<li><a href="user_subject.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/24.gif" align="absmiddle" /> 相册分类管理</a></li>-->
	</li>
	</ul>
	<%end if%>
	<div class="menu-t1" onClick="menu(oblog_7)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 文件管理</div>
	<ul id="oblog_7" style="display: ;">
	<li><a href="user_files.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/71.gif" align="absmiddle" /> 所有文件</a></li>
	<li><a href="user_files.asp?usersearch=1"><img src="oBlogStyle/OblogStyleAdminImages/li/72.gif" align="absmiddle" /> 图片文件</a></li>
	<li><a href="user_files.asp?usersearch=2"><img src="oBlogStyle/OblogStyleAdminImages/li/73.gif" align="absmiddle" /> 压缩文件</a></li>
	<li><a href="user_files.asp?usersearch=3"><img src="oBlogStyle/OblogStyleAdminImages/li/74.gif" align="absmiddle" /> 文档文件</a></li>
	<li style="display:none"><a href="user_files.asp?usersearch=4"><img src="oBlogStyle/OblogStyleAdminImages/li/75.gif" align="absmiddle" /> 媒体文件</a></li>
	</ul>
	
	<!--<div class="menu-t1" onClick="menu(oblog_6)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 团队好友</div>
	<ul id="oblog_6" style="display: none;">
	<li><a href="user_friends.asp?action=add"><img src="oBlogStyle/OblogStyleAdminImages/li/61.gif" align="absmiddle" /> 添加好友</a></li>
	<li><a href="user_friends.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/62.gif" align="absmiddle" /> 好友管理</a></li>
	<li><a href="user_friends.asp?action=add&amp;type=black"><img src="oBlogStyle/OblogStyleAdminImages/li/63.gif" align="absmiddle" /> 添加黑名单</a></li>
	<li><a href="user_friends.asp?usersearch=1"><img src="oBlogStyle/OblogStyleAdminImages/li/64.gif" align="absmiddle" /> 管理黑名单</a></li>
	<li><a href="user_team.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/65.gif" align="absmiddle" /> 用户团队</a></li>
	</ul>-->
	
	
	<div class="menu-t1" onClick="menu(oblog_5)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 参数设置</div>
	<ul id="oblog_5" style="display: ;">
	<li><a href="user_setting.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/51.gif" align="absmiddle" /> 站点设置</a></li>
	<li><a href="user_setting.asp?action=placard"><img src="oBlogStyle/OblogStyleAdminImages/li/52.gif" align="absmiddle" /> 修改公告</a></li>
	<li><a href="user_setting.asp?action=links"><img src="oBlogStyle/OblogStyleAdminImages/li/53.gif" align="absmiddle" /> 友情连接</a></li>
	<li><a href="user_setting.asp?action=userinfo"><img src="oBlogStyle/OblogStyleAdminImages/li/54.gif" align="absmiddle" /> 修改个人资料</a></li>
	<li><a href="user_setting.asp?action=userpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/55.gif" align="absmiddle" /> 修改登录密码</a></li>
	<li><a href="user_setting.asp?action=blogpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/57.gif" align="absmiddle" /> 加密整站</a></li>
	<!--<li><a href="user_setting.asp?action=blogstar"><img src="oBlogStyle/OblogStyleAdminImages/li/58.gif" align="absmiddle" /> 申请用户明星</a></li>-->
	</ul>
