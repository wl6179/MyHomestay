
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<title>tests-用户管理后台</title>
<link title="Oblog_Green" href="OblogStyle/OblogStyleAdminMainFrame.css" rel="stylesheet" media="all" type="text/css" rev="alternate" />
<script src="inc/function.js" type="text/javascript"></script>
</head>

<body onUnload="chkcopy();">

<div class="none"><p>Oblog用户管理界面（无样式纯文字版，适合手机及PDA，您可以点样式切换回其他界面）</p></div>

<div id="menu">
	<div class="menu_content">
		<div class="menu_navimg"></div>
		<div id="nav">
			<ul id="div1">
				<li class="w100p"><a href="HomestayBackctrl-index.asp" >管理首页</a>
					
          <ul>
            <li><a href="HomestayBackctrl-index.asp">管理首页</a></li>
            <li><a href="user_help.asp">用户管理后台帮助</a></li>
          </ul>
				</li>
				<li><a href="user_blogmanage.asp">文章管理</a>
					
          <ul>
            <li><a href="user_post.asp">添加文章</a></li>
            <li><a href="user_blogmanage.asp?usersearch=5">草稿箱</a></li>
            <li><a href="user_blogmanage.asp">所有文章</a></li>
            <li><a href="user_blogmanage.asp?action=downlog">备份文章</a></li>
            <li><a href="user_subject.asp">文章分类管理</a></li>
          </ul>
				</li>
				
				<li><a href="user_photomanage.asp">相册管理</a>
					
          <ul>
            <li><a href="user_post.asp?t=1">添加图片</a></li>
            <li><a href="user_blogmanage.asp?t=1">所有图片</a></li>
            <li><a href="user_subject.asp?t=1">相册分类管理</a></li>
          </ul>
				</li>
				
				<li  class="w100p"><a href="user_comments.asp">评论留言</a>
					
          <ul>
            <li><a href="user_comments.asp">评论管理</a></li>
            <li><a href="user_messages.asp">留言管理</a></li>
          </ul>
				</li>
				<li><a href="user_template.asp?action=showconfig">模版设置</a>
					
          <ul>
            <li><a href="user_template.asp?action=showconfig">选择模板</a></li>
            <li><a href="user_template.asp?action=modiconfig&amp;editm=1">修改主模板</a></li>
            <li><a href="user_template.asp?action=modiviceconfig&amp;editm=1">修改Blog内容模板</a></li>
            <li><a href="user_template.asp?action=bakskin">备份我当前的模板</a></li>
            <li><a href="user_template.asp?action=good">推荐共享我的模板</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_setting.asp">控制面板</a>
					
          <ul>
            <li><a href="user_setting.asp">站点设置</a></li>
            <li><a href="user_setting.asp?action=placard">修改公告</a></li>
            <li><a href="user_setting.asp?action=links">我的友情连接</a></li>
            <li><a href="user_setting.asp?action=userinfo">我的个人资料</a></li>
            <li><a href="user_setting.asp?action=userpassword">修改我的密码</a></li>
            <li><a href="user_setting.asp?action=blogpassword">加密整站</a></li>
            <li><a href="user_setting.asp?action=blogstar">申请用户明星</a></li>
          </ul>
				</li>
				<li><a href="user_team.asp">团队好友</a>
					
          <ul>
            <li><a href="user_friends.asp?action=add">添加好友</a></li>
            <li><a href="user_friends.asp">好友管理</a></li>
            <li><a href="user_friends.asp?action=add&amp;type=black">添加黑名单</a></li>
            <li><a href="user_friends.asp?usersearch=1">管理黑名单</a></li>
            <li><a href="user_team.asp">用户团队</a></li>
            <li>&nbsp;</li>
          </ul>
				</li>
				<li><a href="user_files.asp">文件管理</a>
					
          <ul>
            <li><a href="user_files.asp">所有文件</a></li>
            <li><a href="user_files.asp?usersearch=1">图片文件</a></li>
            <li><a href="user_files.asp?usersearch=2">压缩文件</a></li>
            <li><a href="user_files.asp?usersearch=3">文档文件</a></li>
            <li style="display:none"><a href="user_files.asp?usersearch=4">媒体文件</a></li>
          </ul>
				</li>
			</ul>
		</div>
	</div>
</div>

<div id="top">

<div class="logo"></div>
  <div class="top_button">
	<a href="user_post.asp"><div class="none">发布Blog </div><img src="OblogStyle/OblogStyleAdminImages/space.gif" alt="发布Blog" border="0" class="ico_5" /></a>
		<a href="user_blogmanage.asp">
	<div class="none">管理Blog </div><img class="ico_4" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="管理Blog" /></a>
		
		<a href="user_photomanage.asp"><div class="none">我的相册</div><img class="ico_7" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="我的相册" /></a>
		
		<a href="user_messages.asp"><div class="none">留言管理</div><img class="ico_3" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="留言管理"/></a>
		<a href="user_comments.asp"><div class="none">评论管理</div><img  class="ico_2" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="评论管理" /></a>
		<a href="user_update.asp"><div class="none">重建首页</div><img class="ico_1" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="重建首页" /></a>
		<a  href="HomestayBackctrl-index.asp?t=logout"><div class="none">退出登陆</div><img class="ico_6" src="OblogStyle/OblogStyleAdminImages/space.gif" alt="退出登陆" /></a>
  </div>
  <div class="top_menu">
  <a href="homestay/tests/index.html" target="_blank">我的首页</a> | <a href="user_pmmanage.asp">短消息(0)</a> | <a href="index.asp" target="_blank">站点首页</a> | <a href="HomestayBackctrl-index.asp?t=change0">切换全屏</a></div>
</div>

<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			发布更新
	</div>
	
    
	
	<div class="content_body">
<div id="prompt"><ul>系统消息：<span id="pstr"></span>

<div class="progress1">
<div class="progress2" id="progress">
</div>
</div>
<script>progress.style.width="5%";progress.innerHTML="5%";pstr.innerHTML="更新ID为4的文章...";</script>
<script>progress.style.width="10%";progress.innerHTML="10%";pstr.innerHTML="更新ID为5的文章...";</script>
<script>progress.style.width="15%";progress.innerHTML="15%";pstr.innerHTML="更新ID为12的文章...";</script>
<script>progress.style.width="20%";progress.innerHTML="20%";pstr.innerHTML="更新ID为13的文章...";</script>
<script>progress.style.width="25%";progress.innerHTML="25%";pstr.innerHTML="更新ID为14的文章...";</script>
<script>progress.style.width="30%";progress.innerHTML="30%";pstr.innerHTML="更新ID为17的文章...";</script>
<script>progress.style.width="35%";progress.innerHTML="35%";pstr.innerHTML="更新ID为19的文章...";</script>
<script>progress.style.width="40%";progress.innerHTML="40%";pstr.innerHTML="更新ID为20的文章...";</script>
<script>progress.style.width="45%";progress.innerHTML="45%";pstr.innerHTML="更新ID为28的文章...";</script>
<script>progress.style.width="50%";progress.innerHTML="50%";pstr.innerHTML="更新ID为32的文章...";</script>
<script>progress.style.width="55%";progress.innerHTML="55%";pstr.innerHTML="更新ID为33的文章...";</script>
<script>progress.style.width="60%";progress.innerHTML="60%";pstr.innerHTML="更新ID为34的文章...";</script>
<script>progress.style.width="65%";progress.innerHTML="65%";pstr.innerHTML="更新ID为35的文章...";</script>
<script>progress.style.width="70%";progress.innerHTML="70%";pstr.innerHTML="更新ID为36的文章...";</script>
<script>progress.style.width="75%";progress.innerHTML="75%";pstr.innerHTML="更新ID为37的文章...";</script>
<script>progress.style.width="80%";progress.innerHTML="80%";pstr.innerHTML="更新ID为38的文章...";</script>
<script>progress.style.width="85%";progress.innerHTML="85%";pstr.innerHTML="更新ID为39的文章...";</script>
<script>progress.style.width="90%";progress.innerHTML="90%";pstr.innerHTML="更新ID为40的文章...";</script>
<script>progress.style.width="95%";progress.innerHTML="95%";pstr.innerHTML="更新ID为41的文章...";</script>
<script>progress.style.width="8%";progress.innerHTML="8%";pstr.innerHTML="更新首页...";</script>
<script>progress.style.width="16%";progress.innerHTML="16%";pstr.innerHTML="更新站点信息文件...";</script>
<script>progress.style.width="25%";progress.innerHTML="25%";pstr.innerHTML="生成新文章列表文件...";</script>
<script>progress.style.width="33%";progress.innerHTML="33%";pstr.innerHTML="更新最新留言...";</script>
<script>progress.style.width="41%";progress.innerHTML="41%";pstr.innerHTML="生成首页文章分类文件...";</script>
<script>progress.style.width="50%";progress.innerHTML="50%";pstr.innerHTML="更新功能页...";</script>
<script>progress.style.width="58%";progress.innerHTML="58%";pstr.innerHTML="更新留言板...";</script>
<script>progress.style.width="66%";progress.innerHTML="66%";pstr.innerHTML="更新最新回复...";</script>
<script>progress.style.width="75%";progress.innerHTML="75%";pstr.innerHTML="更新公告...";</script>
<script>progress.style.width="83%";progress.innerHTML="83%";pstr.innerHTML="更新友情连接...";</script>
<script>progress.style.width="91%";progress.innerHTML="91%";pstr.innerHTML="更新BlogName...";</script>
<script>progress.style.width="100%";progress.innerHTML="100%";pstr.innerHTML="整站重新发布完成!";</script>
<li><a href='user_update.asp'>&lt;&lt; 返回上一页</a></li>

</ul>
</div>
	</div>
	
	
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>   
  
  <div id="bottom"><center>  <a href="http://www.myhomestay.com.cn">www.myhomestay.com.cn</a> &copy;&nbsp;Copyright&nbsp;2007-2008.&nbsp;All&nbsp;rights&nbsp;reserved.&nbsp;</center></div><div id="powered"><a href="http://www.myhomestay.com.cn" target="_blank"><img src="images\meiqee.com.gif" border="0" alt="Powered by meiqee.com" /></a></div>

</body>
</html>
