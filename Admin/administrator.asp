<!--#include file="inc/inc_sys.asp"-->
<%
select case request("action")
case "top"
	call admin_top()
case "left"
	call admin_left()
case "main"
	call admin_main()
case "state"
	If Application(cache_name_user&"_systemstate")<>"stop" Then
		Application(cache_name_user&"_systemstate")="stop"
	Else
		Application(cache_name_user&"_systemstate")="run"
	End If
	Application(cache_name_user&"_systemnote")=Request.Form("systemnote")
	Response.Write "<script language=javascript>parent.location.href=""administrator_base.asp"";</script>"
case else
	call main()
end select

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
sub main()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>我的住家网 MyHomestay--后台管理</title>
</head>
<frameset rows="*" cols="251,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="administrator_base.asp?action=left">
  <frameset framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="administrator_base.asp?action=top">
    <frame name="main" scrolling="auto" src="administrator_base.asp?action=main">
  </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>你的浏览器版本过低！本系统要求IE5及以上版本才能使用本系统。</p>
  </body>
</noframes>
</html>
<%
end sub

sub admin_top()
%>
<html>
<head>
<title>我的住家网 MyHomestay 后台管理页面</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
a:link {
	color:#333333;
	text-decoration:none;
	font-size: 12px;
}
a:hover {color:#2786CB;}
a:visited {color:#333333;text-decoration:none}

td {FONT-SIZE: 9pt;COLOR: #000000; FONT-FAMILY: "宋体"}
img {filter:Alpha(opacity:80); chroma(color=#FFFFFF)}
body {background:#ffffff;}
</style>
<base target="main">
<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("images/admin/admin_top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/admin/admin_top_open.gif";
		obj.title="打开左边管理菜单";
	}
	else{
		parent.frame.cols="250,*";
		displayBar=true;
		obj.src="images/admin/admin_top_close.gif";
		obj.title="关闭左边管理菜单";
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0 bgcolor="#CCCCCC">
  <tr valign=middle> 
    <td width=10> <img onClick="switchBar(this)" src="images/admin/admin_top_close.gif" title="关闭左边管理菜单" style="cursor:hand"> </td>
	<td width=80><a href="admin_adminmodifypwd.asp"><font style="font-size:14px;color:red;"><b>修改密码</b></font></a></td>
    <td align="left" width="500">admin_index.广告位1</td>
    <td width="50" align="left"><a href="../index.asp" target="_blank">站点首页</a></td>
  </tr>
</table>
</body>
</html>
<%end sub
sub admin_left()
%>
<html>
<head>
<title>MyHomestay--管理导航</title>
<STYLE type=text/css>
BODY {
 BACKGROUND: #ffffff url(images/admin/topbg3.gif) repeat-x fixed top; MARGIN: 0px; FONT: 12px 宋体;
 	scrollbar-face-color:#ffffff;
	scrollbar-highlight-color: #cccccc;
	scrollbar-shadow-color: #cccccc;
	scrollbar-3d-light-color: #ffffff;
	scrollbar-arrow-color : #cccccc;

	SCROLLBAR-TRACK-COLOR: #ffffff;
	SCROLLBAR-DARKSHADOW-COLOR: #ffffff;
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px 宋体
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px 宋体; COLOR: #333333; TEXT-DECORATION: none
}
A:hover {
	COLOR: #2786CB; TEXT-DECORATION: underline
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #cccccc; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}
.menu_title {
	
}
.menu_title SPAN {
	FONT-WEIGHT: normal;
	LEFT: 8px;
	COLOR: #333333;
	POSITION: relative;
	TOP: 2px;
	font-size: 14px;
}
.menu_title2 {
	
}
.menu_title2 SPAN {
	FONT-WEIGHT: normal;
	LEFT: 8px;
	COLOR: #333333;
	POSITION: relative;
	TOP: 2px;
	font-size: 14px;
}
</STYLE>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr>
  <td valign=top class="menu_title2"> 
    <table width=158 border="0" align=center cellpadding=0 cellspacing=0>
  <tr>
    <td height=42 valign=bottom>
	  <img src="images/admin/title.gif" width=158 height=38>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin/bg.gif id=menuTitle0> 
            <span><a href="administrator_base.asp?action=main" target=main><b>管理首页</b></a> 
            | <a href="administratorLogin.asp?action=logout" target="_top" ><b>退出</b></a></span> </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
	<tr>
                <td height=20><font color=blue><!--#include file="../ver.asp"--></font></td>
              </tr>
    <tr>
                <td height=20>
                	<div id=newver><font color=red>admin_index.广告位2</font></div>
                	</td>
              </tr>
<tr>
                <td height=20>用户名：<%= session("adminname") %></td>
              </tr>
<tr>
                <td height=20>权&nbsp;&nbsp;限：后台管理员</td>
              </tr>
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>



<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title9'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle9 onClick="showsubmenu(9)" style="cursor:hand;"> 
            <span>外教资料管理</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu9'>
<div class=sec_menu style="width:158">
            <table width=130 align=center cellpadding=0 cellspacing=0 >
       
    
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherNo.asp target=main>查看待审核外教资料</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacher.asp target=main>查看已审核外教资料</a></td>
              </tr>
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherDeleted.asp target=main>被删除的外教资料</a></td>
              </tr>
              

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle4 onClick="showsubmenu(4)" style="cursor:hand;"> 
            <span>用户管理</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu4'>
<div class=sec_menu style="width:158">
            <table width=130 align=center cellpadding=0 cellspacing=0 >
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_UserLogclass_Nav.asp target=main>添加新栏目空间</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_user.asp target=main>中国家庭用户管理</a></td>
              </tr>
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate.asp target=main>加盟用户管理</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherNo.asp target=main>查看待审核外教资料</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacher.asp target=main>查看已审核外教资料</a></td>
              </tr>
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherDeleted.asp target=main>被删除的外教资料</a></td>
              </tr>
              
			  <!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_user.asp?action=gouser1" target="main">直接进入栏目管理</a></td>
              </tr>-->
				<!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_blogstar.asp" target="main">用户明星</a></td>
              </tr>-->
              

              <!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_rename.asp" target="main">用户改名</a></td>
              </tr>-->
              
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_admin.asp?Action=Add target="main">添加管理员</a> | <a href=admin_admin.asp target=main>管理</a></td>
              </tr>
			   <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_UserLogclass_Nav.asp?Action=gouser2&username=myhomestay"  target=main >接收站内消息</a></td>
              </tr>
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_pmall.asp  target=main >发送站内消息</a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle2 onClick="showsubmenu(2)" style="cursor:hand;"> 
            <span>常规设置</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu2'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>

	
	
	      
              <!--<tr> 
                <td width="15" height=20 align="right"><img src="images/admin/bullet.gif"></td>
                <td width="115"><a href=admin_setup.asp target=main>网站信息配置</a></td>
              </tr>-->
              
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_logclass_IndexAdMapsControl.asp" target="main">首页变换广告管理</a></td>
              </tr>
              
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userclass.asp" target="main">用户分类管理</a></td>
              </tr>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_logclass.asp" target="main">文章分类管理</a></td>
              </tr>
              <!--<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_logclass.asp?t=1" target="main">相册分类管理</a></td>
              </tr>-->
              <!--<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_tags.asp" target="main">系统TAG管理</a></td>
              </tr>-->
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_filtrate.asp target=main>文章敏感字过滤管理</a></td>
              </tr>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_lockip.asp target=main>限制IP管理</a></td>
              </tr>
              
              
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=3 target=main>家庭注册条款</a> </td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=5 target=main>加盟注册条款</a> </td>
              </tr>
              
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=2 target=main>修改今日推荐家庭☆</a> </td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=6 target=main>修改今日推荐外教★</a> </td>
              </tr>
              
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=1 target=main>修改友情链接</a> | <a href=admin_note.asp?action=do1 target=main>文本</a></td>
              </tr>
              
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_ad.asp target=main>用户页面广告管理</a></td>
              </tr>
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=4 target=main>用户后台通知</a> | <a href=admin_note.asp?action=do4 target=main>文本</a></td>
			  </tr>
              <!--<tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userdir.asp" target="main">用户目录管理</a></td>
              </tr>-->
              <tr>
                <td height=20>+</td>
                <td height=20>-----------------</td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=86 target=main>用户注册成功消息</a> </td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=87 target=main>修改—关于我们内容</a> </td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=88 target=main>修改—联系我们内容</a> </td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=89 target=main>修改—合作代理内容</a> </td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=90 target=main>修改—帮助中心内容</a> </td>
			  </tr>
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=91 target=main>修改—特别问题提示</a> </td>
			  </tr>
			  
			  <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=92 target=main>修改—推荐保护软件</a> </td>
			  </tr>
              
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<!--<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(5)" style="cursor:hand;"><span>模板管理</span> 
          </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu5'><div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_sysskin.asp?action=addskin target=main>添加系统模版</a> 
                  | <a href=admin_sysskin.asp?action=showskin target=main>管理</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_skin.asp?action=insys1" target="main">系统模版导入</a> 
                  | <a href="admin_skin.asp?action=outsys" target="main">导出</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userskin.asp?action=addskin" target="main">添加用户模版</a> 
                  | <a href="admin_userskin.asp?action=showskin&ispass=1" target="main">管理</a></td>
              </tr>
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_skin.asp?action=inuser1" target="main">用户模版导入</a> 
                  | <a href="admin_skin.asp?action=outuser" target="main">导出</a></td>
              </tr>
            </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20 class="menu_title"></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>-->
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(6)" style="cursor:hand;"> 
            <span>上传文件管理</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu6'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_uploadfile_user.asp target=main>上传管理用户清单</a><a href=admin_sysskin.asp?action=showconfig target=main></a></td>
              </tr>
				<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_uploadfile.asp target=main>上传管理文件清单</a><a href=admin_sysskin.asp?action=showconfig target=main></a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 class="menu_title"></td>
              </tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%
'如果是SQL Server数据库则不再显示该节
If IS_SQLDATA=0 Then
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle7 onClick="showsubmenu(7)" style="cursor:hand;"> 
            <span>数据库管理</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu7'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td width="15" height=20><img src="images/admin/bullet.gif"></td>
  <td><a href=admin_database.asp?Action=Compact target=main>压缩数据库</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=Backup target=main>备份数据库</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=Restore target=main>恢复数据库</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=SpaceSize target=main>系统空间占用</a></td>
</tr>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%End If%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(8)" style="cursor:hand;"> 
            <span>系统分析</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu8'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=year target=main>数据年度分析</a></td>
              </tr>
				<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=month target=main>数据月度分析</a></td>
              </tr>
			  <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=hour target=main>数据时段分析</a></td>
              </tr>
              
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_count.asp" target="main">更新系统数据</a></td>
              </tr>
              <tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_user.asp?Action=Update" target="main">生成用户静态页</a></td>
              </tr>
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_user_OldVersionWebSite.asp target=main><font color="#Fcc000"><b>旧版网站管理后台</b></font></a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 class="menu_title"></td>
              </tr>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 class="menu_title"></td>
              </tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle9> 
            <span>MyHomestay信息</span> </td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20><br>
                  版权所有：&nbsp;<a href="http://www.myhomestay.com.cn" target=_blank>myhomestay.com.cn</a><br>
                  设计制作：&nbsp;<a href='mailto:wangliang6179@163.com'>王亮</a><BR>
                  技术支持：&nbsp;<a href="http://www.myhomestay.com.cn" target=_blank>我的住家网</a><br>
                  
                  <br>
</td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
</table>
</body>
</html>
<%
end sub

sub admin_main()
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
    
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
%>
<html>
<head>
<title>MyHomestay后台管理首页</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/admin/admin_STYLE.CSS">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
<tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>MyHomestay后台管理首页</strong></div>
</tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
<td height=25 colspan=2 class="topbg"><strong>系统启动与关闭</strong></th></tr> 
     <tr> 
    <td width="20%" class="tdbg" height=23>当前系统状态</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=green>
    <%
    	Dim strSubmit
    	If Application(cache_name_user&"_systemstate")="stop" Then
    		Response.Write "<font color=red>关闭中</font>"
    		strSubmit="重新启动"
    	Else
    		Response.Write "正常运行中......"
    		strSubmit="关闭系统"
    	End If
    %></font>
    </td>
  </tr>
   <form name="systemcontrol" method="post" action="administrator_base.asp?action=state">
	  <tr> 
	    <td width="20%" class="tdbg" height=23>关闭时提示：</td>
	    <td width="80%" class="tdbg"> 
	    	<table width="100%" border="0" cellspacing="0" cellpadding="2">
	        <tr> 
	          <td>
	          	<textarea cols=60 rows=5 name="systemnote"><%
	          		If Application(cache_name_user&"_systemnote")="" Then
	          			Response.Write "系统正在维护，稍后开放"
	          		Else
	          			Response.Write Application(cache_name_user&"_systemnote")
	          		End If
	          		%>
	          		</textarea>	          	
	          	</td>
	        </tr>
	      </table></td>
	  </tr>
	  <tr>
	    <td class="tdbg" height=23></td>
	    <td class="tdbg"><input type="submit" value="<%=strSubmit%>">(系统处于关闭状态时,以管理员身份仍能完整访问)</td>
	  </tr>
	</form>

</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 class="topbg"><strong>MyHomestay 帮 助</strong>
  <tr>
    <td height=23 class="tdbg">1、<a href="../jsreadme.asp" target="_blank"><strong>首页调用文件的测试及帮助说明请点击这里</strong>。</a></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">2、<a href="admin_skin_help.asp" target="_blank"><strong>系统模版及用户模版的标记说明请点击这里</strong>。</a></td>
  </tr>
  <tr>
    <td height=23 class="tdbg"><p>3、用户权限<br> 
        　①已注册待审核用户：不能发布文章，但可以在关闭匿名评论的情况下发布留言和评论。<br>      
        　②普通注册用户：可以发布文章、评论、回复。<br>
      　③VIP用户：有普通注册用户的所有权限，可单独设置上传空间。<br>
      　④前台管理员：有VIP用户的所有权限，另可以修改，删除所有用户的文章、评论、留言。<br>
      　⑤各用户组的上传权限，请在网站信息配置里设定。</p>
    </td>
  </tr>
  <tr>
    <td class="tdbg">4、将用户锁定以后，此用户的blog页面也将被屏蔽。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">5、将IP屏蔽以后，此IP用户将不能登陆，且不能发布评论及留言。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">6、将用户设置为推荐，必须在后台修改用户资料才能实现。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">7、若上传文件不正常，请检查是否文件尺寸过大及服务器是否支持fso。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">8、有任何问题，请咨询MyHomestay官方网站<a href="http://www.myhomestay.com.cn" target="_blank">http://www.myhomestay.com.cn</a>。</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>服 务 器 信 息</strong><tr> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%" class="tdbg">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%" class="tdbg">数据库地址：</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>FSO文本读写：
      <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>

      <%end if%></td>
    <td width="50%" class="tdbg">数据库使用：
      <%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>Jmail组件支持：
      <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
    <td width="50%" class="tdbg">CDONTS组件支持：
      <%If Not IsObjInstalled(theInstalledObjects(14)) Then%>
      <font color="red"><b>×</b></font>
      <%else%>
      <b>√</b>
      <%end if%></td>
  </tr>
  <tr>
    <td class="tdbg" height=23>&nbsp;</td>
    <td align="right" class="tdbg">&nbsp;</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>MyHomestay开发</th></strong> 
  <tr> 
    <td width="20%" class="tdbg" height=23>核心程序设计：</td>
    <td width="80%" class="tdbg"> <table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td width="60">王亮</td>
          <td>（&nbsp;Email：wangliang6179@163.com;&nbsp;&nbsp;&nbsp;&nbsp;QQ：4985612）</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="20%" class="tdbg" height=23>&nbsp;</td>
    <td width="80%" class="tdbg">注：模版部分仅供测试使用，版权归原作者所有。</td>
  </tr>
  <tr>
    <td class="tdbg" height=23>参与界面设计：</td>
    <td class="tdbg">王亮 Email:wangliang6179@163.com </td>
  </tr>
</table>
</body>
</html>
<%

end sub
%>

