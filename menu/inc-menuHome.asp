
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
	<a href="HomestayBackctrl-personalMessageHome.asp"><img src="/images/mailmanage.gif">住家网消息(<%=chkpm()%>)</a>
</div>

<p class="public-title011">
	<%=oblog.logined_uname%>—家庭用户中心
</p>
	<div id="toplogo"></div>
	<div id="sub" style="margin-left:10px;">
	<img src='/images/[MyHomestay]Manager.gif' /><a href="HomestayBackctrl-indexHome.asp" target="_top">管理首页</a>  |&nbsp;&nbsp;<img src='/images/logout.gif' /> <a href="HomestayBackctrl-indexHome.asp?t=logout&re=1">退出后台</a>
	 <!--<br />-->
	 <!--<a href="HomestayBackctrl-personalMessageHome.asp"><img src="/images/mailmanage.gif">住家网消息(<%=chkpm()%>)</a>--><br />&nbsp;
	</div>
	<div class="menu-t1" onClick="menu(oblog_1)" style="margin-left:30px;"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> 家庭资料管理</div>
	<ul id="oblog_1" style="margin-left:30px;">
	<li><a href="HomestayBackctrl-InfoHome.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> 修改家庭资料</a></li>
	
	
	<li><a href="HomestayBackctrl-PhotofilesHome.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> 上传家庭照片</a></li>
	<li><a href="HomestayBackctrl-PassSettingHome.asp?action=userpassword&pass=1"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 修改我的密码</a></li>
	<li><a href="HomestayBackctrl-PassSettingHome.asp?action=userpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> 设置密码提示</a></li>
	
	</ul>

	
	
	