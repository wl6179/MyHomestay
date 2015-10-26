
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
<div style="position:absolute; top:250px; left:830px; width:120px;">
	<img src="/images/mailmanage.gif"><a href="HomestayBackctrl-personalMessageCooperate.asp">企业消息(<%=chkpm()%>)</a>
</div>

<p class="public-title011">
	<%=oblog.logined_uname%>—加盟用户中心
</p>
	<div id="toplogo"></div>
	<div id="sub">
	<img src='/images/[MyHomestay]Manager.gif' /><a href="HomestayBackctrl-indexCooperate.asp" target="_top">管理首页</a>| <img src='/images/logout.gif' /><a href="HomestayBackctrl-indexCooperate.asp?t=logout&re=1">退出后台</a>
	 <!--<br />
	 <img src="/images/mailmanage.gif"><a href="HomestayBackctrl-personalMessageCooperate.asp">企业消息(<%=chkpm()%>)</a>--><br />&nbsp;
	</div>
	<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 提交外教资料</div>
	<ul id="oblog_1">
	
	<li><a href="HomestayBackctrl-PostSubmitManageCooperate.asp" title="查看所有提交过的外教资料"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> 外教资料管理</a></li>
	<li><a href="HomestayBackctrl-PostSubmitManageCooperate.asp?usersearch=5" title="暂存在草稿箱中，不提交。"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> 暂存草稿箱</a></li>
	
	<% If oblog.logined_uisValidate=0 Then %>
	<li><a  title="请您先等待通过审核..." onclick="alert('请您先等待通过审核...');" style="cursor:hand;"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> <font color="#038ad7">提交外教信息</font></a></li>
	<% Else %>
	<li><a href="HomestayBackctrl-PostSubmitCooperate.asp" title="我们将会在24小时内，完成您提交资料的审核工作，令其尽快发布在住家网上。"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> <font color="#038ad7">提交外教信息</font></a></li>
	<% End If %>
	
	</ul>
	
	
	
	
	
	<div class="menu-t1" onClick="menu(oblog_2)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 加盟用户资料管理</div>
	<ul id="oblog_2">
	
	
	<li><a href="HomestayBackctrl-InfoCooperate.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> 修改我的加盟资料</a></li>
	
	<li><a href="HomestayBackctrl-PassSettingCooperate.asp?action=userpassword&pass=1"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 修改我的密码</a></li>
	<li><a href="HomestayBackctrl-PassSettingCooperate.asp?action=userpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 设置密码提示</a></li>
	
	</ul>

	
	
	