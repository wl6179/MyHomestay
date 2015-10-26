<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
'退出系统
'If request("logout")="logout" Then
'	If cookies_domain <> "" Then
'        response.Cookies(cookies_name).domain = cookies_domain
'    End If
'	Response.Cookies(cookies_name)("username")=oblog.CodeCookie("")
'	Response.Cookies(cookies_name)("password")=oblog.CodeCookie("")
'	Response.Cookies(cookies_name)("userurl")=oblog.CodeCookie("")
'	if request("re")="1" then
'		'response.Redirect(oblog.comeurl)
'		response.Write("<script language=JavaScript>top.location='"&oblog.comeurl&"';< /script>")
'	else'退出后回到主页；
'		response.Write("< script language=JavaScript>top.location='login.asp';</ script>")
'	end if
'	Set oBlog=Nothing
'	response.End()
'End If
%>
<%
dim username,password,show_login,CookieDate,fromurl,action
action=request("action")
if action<>"showindexlogin" And action<>"showindexlogin_team" and action<>"showjs" then
	if  oblog.checkuserlogined() And session("indexHomeLogined")<>"indexHomeLogined" then response.Redirect("HomestayBackctrl-indexHome.asp")'如果用户已经跳转到 管理页面后，就不需要再跳转了，直接访问当前login页面；
	if  oblog.checkuserlogined_team() And session("indexCooperateLogined")<>"indexCooperateLogined" then response.Redirect("HomestayBackctrl-indexCooperate.asp")'如果用户已经跳转到 管理页面后，就不需要再跳转了，直接访问当前login页面；
end if

'防注入功能WL[2007-09-23];参数处理___oblog.filt_badstr()函数；
username=oblog.filt_badstr(trim(request.Form("username")))
password=trim(request.form("password"))
CookieDate=trim(request("CookieDate"))
fromurl=trim(request("fromurl"))
if username<>"" or request("chk")="1" then '需要check的时候，执行检测；	
	call sub_chklogin
else
	if action="showindexlogin" then
		call sub_showindexlogin()
	elseif action="showindexlogin_team" then
		call sub_showindexlogin_team()
		
	elseif action="showjs" then
		blogurl=oblog.setup(3,0)
		call sub_showindexlogin()
	elseif action="showjs_team" then
		blogurl=oblog.setup(3,0)
		call sub_showindexlogin_team()
		
	else
		call sub_showlogin()
	end if
end if


sub sub_showlogin()

	call sysshow_4()  '首先获得模版~...在往下，就是判断是处理注册 或者 显示出注册表单了；
	select case action
		case "chkreg"
		call aaa()'处理注册;
		case else 
		'''call aaa()'显示出注册表单了;
	end select
	'show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'将标签替$show_list$换为 该有的内容 show_reg 后 + 底部版权信息代码；
	show=replace(show,"$show_list$",fromurl)&oblog.site_bottom'将标签替$show_list$换为 该有的内容 show_reg 后 + 底部版权信息代码；
	response.Write show'全面输出。

end sub
%>


<% sub aaa()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <meta name="Robots" content="index,follow" />
	<meta name="author" content="王亮, wangliang6179@163.com" />
	<meta name="copyright" content="www.myhomestay.com.cn,我的住家网" />
	<meta name="description" content="homestay,myhomestay,china,beijing,中国,北京,住家,外教住家,老外住家,住家外教,外教,外国人,老外,中国家庭,外语,rent" />
	<meta name="keywords" content="homestay,myhomestay,china,beijing,learn,english,中国,北京,住家,老外,家庭,外语" />
    <title>
      我的住家网- 登录区 homestay in beiijng|guest house in beijing|Learn Chinese in China
    </title>
    <link rel="stylesheet" type="text/css" href="css/StyleIndex.css" />

	<link rel="stylesheet" type="text/css" href="inc/divCN-EN.css" />
<script type="text/javascript" language="javascript" src="images/blog/tab2.js"></script>
  </head>
  <body>



	<div id="head">
	  	<div id="head2">
		
			<div id="head2-logo">
				
				<ul>
					<li class="active"><a href="http://www.myhomestay.com.cn">简体中文</a></li>
					<li><a href="http://www.myhomestay.com.cn">English</a></li>
					<!--<li><a href="http://www.myhomestay.com.cn">Janpan</a></li>-->
				</ul>
				
			</div>
			
			
			<div id="head2-menu">
				<div id="head2-search">
					<span id="joincompany_login"><a href="/login.asp">企业登录</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">企业加盟</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>客户咨询热线</span>：<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">设为首页</a>&nbsp;&nbsp;&nbsp;-->
					<a href="#" onClick="javascript:AddFavoriteOnClick();">按空格键加入收藏夹</a>&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">帮助中心</a>&nbsp;&nbsp;&nbsp;-->
				</div>
				
				<div id="divCN-EN">
				<ul>
					<li><a href="/index.asp">首页<br />Home</a></li><!--#include file="menu/incIndexNav.htm"-->

					
				</ul>
				</div>
			</div>
		
		</div>
		
	  </div>
	
	
	<div id="body"><!--#body-->
	<div id="gbody"><!--#gbody-->
	
		<div id="left-margin"></div>
		
		<div id="main"><!--the main content area-->
			
			
			
			
			<div id="mainContentList">
				<div id="main-top-round">
					<div id="main-top-round-left"></div>
					<div id="main-top-round-right"></div>
				</div>
				
				<div id="main-main-round">
					<style>
						table {
						background:#ffffff;
						border:1px solid #999999;
						border-collapse: collapse;
						font-size:13px;
						}
						td {
						padding:0px 3px 0px 12px;
						width:100px;
						height:18px;
						border:1px solid #999999;
						}
						.tr01 {
						/*font:bold;*/
						background:#A1E5FA;
						font-boil:2px;
						color:#4f6b72;
						}
						.tr02 {
						background:#ffffff;
						color:#797268;
						}
						.tr03 {
						background:#F7FDFB;
						color:#4f6b72;
						}
					</style>
					<div id="main-TextArea01" style="text-align:right">
					
					
					
						<center><h2 style="font-size:15px">登录区</h2></center>
						
						
						
<form name="reg" method="post" action="login.asp?fromurl=<%=fromurl%>">
<p class="p1" style="margin:13px;">帐户名：<input name="username" type="text" class="put1" size="25" />
</p>
<p class="p2" style="margin:13px;">密码名：<input name="password" type="password" class="put1" size="25" />
</p>
<p class="p3" style="margin:13px;">
	<input name="CookieDate" type="radio" value="0" checked="checked" title="若您不在自己的电脑登陆，请不要保存密码" />不保存
	<input type="radio" name="CookieDate" value="1" />一天
	<input type="radio" name="CookieDate" value="2" />一月
	<input type="radio" name="CookieDate" value="3" />一年</p>
<p class="p4" style="margin:13px;"></p>
<div class="but" style="margin:13px;">
	<input name="Submit" type="image" src="OblogStyle/OblogStyleAdminImages/login.gif" value="提交" onclick="submit()" /></div>
	<div class="code" style="margin:13px;">
	<%if oblog.setup(58,0)=1 then
	%>验证码：<input name="codestr" type="text" class="put2" size="6" maxlength="4" />
	<%response.Write(oblog.getcode)
	end if
	%>
	</div></p>
	</form>



</div>
  
  					<!--wl-->
					</div>
					
					
					<div id="main-main-round2">
						<center>
					  
					    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="585" height="90">
                          <param name="movie" value="images/90620_hp585_90.swf" />
                          <param name="quality" value="high" />
                          <embed src="images/90620_hp585_90.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="585" height="90"></embed>
					      </object>
					  	</center>
					</div>
					
					
		
					
					
				
					<div id="main-banner01">
						<div id="main-bottom-round-left"></div>
						<div id="main-bottom-round-right"></div>
					</div>
				</div>
				
				
				<div id="publicList">
					<div id="public-top-round">
						<div id="public-top-round-left"></div>
						<div id="public-top-round-right"></div>
					</div>
					
					<div id="public-main-round">
					<%'main()%>
						<div class="public-TextArea01">
							<p class="public-title01">
							
							</p>
							
					
							
						</div>
						
					
					</div>
				
					<div id="public-bottom-round">
						<div id="public-bottom-round-left"></div>
						<div id="public-bottom-round-right"></div>
					</div>
				</div>				
			</div>
			

</div>
		</div><!--#gbody-->	
	  </div><!--#body-->


  

	  		
<%=oblog.user_copyright%>
			
			
	  	


  <script src="inc/divCN-EN.js" type="text/javascript"></script>
  </body>
</html>

<%

end sub



sub sub_chklogin()
	if oblog.chkiplock() then
		oblog.adderrstr("对不起！你的IP已被锁定,不允许登陆，！")
		oblog.showerr
		exit sub
	end if
	if oblog.setup(58,0)=1 then
		if not oblog.codepass then oblog.adderrstr("验证码错误，请刷新后重新输入！"):oblog.showerr
	end if
	if UserName="" then oblog.adderrstr("登陆用户名不能为空！")
	if Password="" then oblog.adderrstr("登陆密码不能为空！")
	if oblog.errstr<>"" then oblog.showerr:exit sub
	if CookieDate="" then CookieDate=0	else CookieDate=Clng(CookieDate)
	password=md5(password)
	if is_ot_user=1 then
		call ot_chklogin()
	else
		If Request("teamSubmit")<>"teamSubmit" Then
		oblog.ob_chklogin UserName,password,CookieDate
		Else
		oblog.ob_chklogin_team UserName,password,CookieDate
		End If
	end if
	if fromurl<>"" then
		response.Redirect(replace(fromurl,"$","&"))
	else
		if action="showindexlogin" Or action="showindexlogin_team" then
			if instr(oblog.comeurl,"/err.asp")>0 then
				response.Redirect("index.asp")
			else
				response.Redirect(oblog.comeurl)
		
			end if
		else
			response.Redirect("HomestayBackctrl-indexHome.asp")
		end if
	end if
end sub



sub ot_chklogin()
	dim sql,rs,rsreg
	if not IsObject(ot_conn) then link_database
	sql="select * from "&ot_usertable&" where "&ot_username&"='"& username & "' and "&ot_password&" ='" & password &"'"
	set rs=ot_conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.adderrstr("用户名或密码错误，请重新输入！"):oblog.showerr
		exit sub
	else
		set rsreg=server.CreateObject("adodb.recordset")
		rsreg.open "select * from [oblog_user] where username='"& username &"'",conn,1,3
		if rsreg.eof then
			dim reguserlevel
			if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
			set rsreg=server.CreateObject("adodb.recordset")
			rsreg.open "select top 1 * from [oblog_user]",conn,1,3
			rsreg.addnew
			rsreg("username")=username
			rsreg("password")="othertable"
			rsreg("user_dir")=oblog.setup(30,0)
			rsreg("user_level")=reguserlevel
			rsreg("lockuser")=0
			rsreg("en_blogteam")=1
			rsreg("adddate")=now()
			rsreg.update
			oblog.execute("update oblog_user set user_folder=userid where username='"&username&"'")
			oblog.execute("update oblog_setup set user_count=user_count+1")
			rsreg.close
			set rsreg=nothing
			oblog.SaveCookie username,password,0," "
			oblog.CreateUserDir username,1
			set rs=nothing
			oblog.showok "您是第一次激活blog系统，请完善blog资料!","user_setting.asp"
			response.End()
		else
			rsreg("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
			rsreg("LastLoginTime")=now()
			rsreg("LoginTimes")=rsreg("LoginTimes")+1
			rsreg.update			
		end if
		rsreg.close
		set rsreg=nothing
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.SaveCookie username,password,CookieDate,""
		if action="showindexlogin" then
			'response.Redirect(oblog.comeurl)
			response.Redirect("HomestayBackctrl-index.asp")
		else
			response.Redirect("HomestayBackctrl-index.asp")
		end if	
	end if
end sub

sub sub_showindexlogin()
	dim show_userlogin
	'if oblog.CheckUserLogined()=False Or oblog.logined_ulevel_team=1 then
	if oblog.CheckUserLogined()=False  then
			

			show_userlogin="<form style='color:#333333;' action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method='post' name='UserLogin'>" 
			  show_userlogin=show_userlogin&"<ul style='padding:0px;margin:8px 0px 0px 15px; font-size:13px;'><li style='padding:0px;margin:0px'><span>家&nbsp;庭&nbsp;帐&nbsp;号&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input id='UserName' name='UserName' type='text' size='13'></li>"
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>家&nbsp;庭&nbsp;密&nbsp;码&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input name='Password' type='password' id='Password' size='13' /></li>"
			  if oblog.setup(58,0)=1 then
				  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>验&nbsp;证&nbsp;码：</span><input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />&nbsp;"&oblog.getcode&"</li>"
			  end if 
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><input name='Login' type='submit' id='Login' value='登 录' style='font-size:12px'>&nbsp;<input name='regHomestay' type='button' value=' 申 请 接 待 ' onClick=""window.location.href='"&blogurl&"RegisterHome.asp'""><br /><a href='"&blogurl&"lostpassword.asp'>忘记密码？</a></li>"


				show_userlogin=show_userlogin&"</ul>"
				show_userlogin=show_userlogin&"</form>"


				

			
		
		
'		show_userlogin="<form action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method='post' name='UserLogin'>" 
'    	show_userlogin=show_userlogin&"<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0' style=""font-size:12px"">" 
'        show_userlogin=show_userlogin & "<tr><td height='25' class=login>用户名：<input name='UserName' type='text' id='UserName' size='12' maxlength='20'></td></tr>" 
'        show_userlogin=show_userlogin & "<tr><td height='25' class=login >密　码：<input name='Password' type='password' id='Password' size='12' maxlength='20'></td></tr>" 
'		if oblog.setup(58,0)=1 then
'			show_userlogin=show_userlogin&"<tr><td height='25' >验证码：<input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />"&oblog.getcode&"</td></tr>"
'		end if
'        show_userlogin=show_userlogin & "<tr><td height='25'>　　　　<input type=""checkbox"" name=""CookieDate"" value=""3"">记住密码</td></tr>" 
'		
'		show_userlogin=show_userlogin & "<tr><td height='25'><input name='Login' type='submit' id='Login' value='登录'> " 
'        show_userlogin=show_userlogin & "<a href='"&blogurl&"reg.asp'>用户注册</a>&nbsp;<a href='"&blogurl&"lostpassword.asp'>忘记密码</a><br></td>"       
'        show_userlogin=show_userlogin & "</tr></table></form>" 
	else
	
		
		show_userlogin="<table style='color:#333333;' border='0' align='center' style=' font-size:13px;'>"
			
                show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>欢迎您 <font color='#DE728A'>" & oblog.logined_uname & "</font><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				
				show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>您的身份："
						  if oblog.logined_ulevel=6 then
							show_userlogin= show_userlogin&"<font color='#038ad7'>注册家庭用户</font>"
						  elseif oblog.logined_ulevel=7 then
							show_userlogin= show_userlogin&"注册栏目发布"
						  elseif oblog.logined_ulevel=8 then
							show_userlogin= show_userlogin&"VIP栏目发布"
						  elseif oblog.logined_ulevel=9 then
							show_userlogin= show_userlogin& "前台管理发布"
						  end if 
						  'show_userlogin= show_userlogin& "<br><b>用户控制面板：</b><br>" & vbcrlf
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<br /><a style='color:#ffffff;' href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>栏目首页</a>"
						  else
						  show_userlogin= show_userlogin& "<br />"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<br /><img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-index.asp' ><strong>进入栏目空间管理</strong></a><br />"
						  elseif oblog.logined_ulevel_team=1 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-indexCooperate.asp' >合作伙伴管理中心</a>&nbsp;&nbsp;&nbsp;"
						  else
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-indexHome.asp' ><strong>进入我的家庭管理</strong></a>&nbsp;&nbsp;&nbsp;"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  show_userlogin= show_userlogin& "<a href='"&blogurl&"user_post.asp' >发布栏目文章</a>&nbsp;&nbsp;"
						  else
						  
						  end if

						  
						  show_userlogin= show_userlogin& "<img src='/images/logout.gif' /><a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>退出</a><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				

				


		
		
'		show_userlogin="--欢迎您," & oblog.logined_uname & "--<br />"
'		show_userlogin= show_userlogin&"您的身份："
'		if oblog.logined_ulevel=7 then
'			show_userlogin= show_userlogin&"注册栏目发布"
'		elseif oblog.logined_ulevel=8 then
'			show_userlogin= show_userlogin&"VIP栏目发布"
'		elseif oblog.logined_ulevel=9 then
'			show_userlogin= show_userlogin& "前台管理发布"
'		end if
'		'show_userlogin= show_userlogin& "<br><b>用户控制面板：</b><br>" & vbcrlf
'		show_userlogin= show_userlogin& "<br /><a href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>我的首页</a>"
'		show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-index.asp' target='_blank'>管理中心</a><br />" 
'		show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?url=user_post.asp' target='_blank'>发布文章</a>&nbsp;&nbsp;"
'		show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>注销登录</a><br />" 
	end if
	Response.Write oblog.htm2js_div(show_userlogin,"ob_login")
end sub

sub sub_showindexlogin_team()
	dim show_userlogin
	
	'if oblog.CheckUserLogined_team()=False Or oblog.logined_ulevel_team=0 then'并且不是合作伙伴的登录；
	if oblog.CheckUserLogined_team()=False then'并且不是合作伙伴的登录；
			

			show_userlogin="<form style='color:#333333;' action="""&blogurl&"login.asp?action=showindexlogin_team&chk=1"" method='post' name='UserLogin_team'>" 
			  show_userlogin=show_userlogin&"<ul style='padding:0px;margin:8px 0px 0px 15px; font-size:13px;'><li style='padding:0px;margin:0px'><span>LoginName</span>" 
			    show_userlogin=show_userlogin&"<input id='UserName' name='UserName' type='text' size='13'></li>"
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>Password&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input name='Password' type='password' id='Password' size='13' /></li>"
			  if oblog.setup(58,0)=1 then
				  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>验&nbsp;证&nbsp;码：</span><input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />&nbsp;"&oblog.getcode&"</li>"
			  end if 
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><input name='Login' type='submit' id='Login' value='登 录' style='font-size:12px'>&nbsp;<input name='regHomestay' type='button' value='申 请 加 盟' onClick=""window.location.href='"&blogurl&"RegisterCooperate.asp'""><br /><a href='"&blogurl&"lostpassword.asp'>忘记密码？</a></li>"

			  show_userlogin=show_userlogin&"<input name='teamSubmit' type='hidden' id='teamSubmit' value='teamSubmit'>"
				
				show_userlogin=show_userlogin&"</ul>"
				show_userlogin=show_userlogin&"</form>"

	else
	
		
		show_userlogin="<table style='color:#333333; font-size:13px;' border='0' align='center'>"
			
                show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>欢迎您," & oblog.logined_uname & "<br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				
				show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>您的身份："
						  if oblog.logined_ulevel=6 then
							show_userlogin= show_userlogin&"<font color='#038ad7'>合作伙伴</font>"
						  elseif oblog.logined_ulevel=7 then
							show_userlogin= show_userlogin&"注册栏目发布"
						  elseif oblog.logined_ulevel=8 then
							show_userlogin= show_userlogin&"VIP栏目发布"
						  elseif oblog.logined_ulevel=9 then
							show_userlogin= show_userlogin& "前台管理发布"
						  end if 
						  'show_userlogin= show_userlogin& "<br><b>用户控制面板：</b><br>" & vbcrlf
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<br /><a style='color:#ffffff;' href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>栏目首页</a>"
						  else
						  show_userlogin= show_userlogin& "<br />"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-index.asp' >栏目管理</a><br />"
						  elseif oblog.logined_ulevel_team=1 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-indexCooperate.asp' ><strong>加盟帐号管理中心</strong></a>&nbsp;&nbsp;&nbsp;"
						  else
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-indexHome.asp' ><strong>进入我的家庭管理</strong></a>&nbsp;&nbsp;&nbsp;"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<a href='"&blogurl&"user_post.asp' target='_blank'>发布栏目文章</a>&nbsp;&nbsp;"
						  else
						  
						  end if

						  
						  show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>退出</a><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				

	end if
	Response.Write oblog.htm2js_div(show_userlogin,"ob_login_team")
end sub
%>

