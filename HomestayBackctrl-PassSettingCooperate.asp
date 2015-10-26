<!--#include file="user_topCooperate.asp"-->
<!--#include file="inc/class_blog.asp"-->


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
				<span id="joincompany_login"><a href="/login.asp">企业登录</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">企业加盟</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>客户咨询热线</span>：<span>010-85493388</span>/<span>13146398085</span> &nbsp;&nbsp;&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">设为首页</a>&nbsp;&nbsp;&nbsp;-->
				<!--<a href="#" onClick="javascript:AddFavoriteOnClick();">按空格键加入收藏夹</a>-->&nbsp;
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
					<div id="main-TextArea01">
					
					
					
						<center><h2>密码资料</h2></center>
			

<%
dim action
action=request("action")
select case action
	case "links"
	call links
	case "savelinks"
	call savelinks
	case "placard"
	call placard
	case "saveplacard"
	call saveplacard
	case "blogpassword"
	call blogpassword
	case "addblogpassword"
	call addblogpassword
	case "unblogpassword"
	call unblogpassword
	case "savesitesetup"
	call savesitesetup
	case "blogstar"
	call blogstar
	case "saveblogstar"
	call saveblogstar
	case "saveuserlostpass"
	call saveuserlostpass
	case "userpassword"
	call userpassword
	case "saveuserpassword"
	call saveuserpassword
	case "saveuserinfo"
	call saveuserinfo
	case "userinfo"
	call userinfo
	case else
	call sitesetup	
end select

%>


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
								加盟管理中心
							</p>
							
							<!--#include file="menu/inc-menuCooperate.asp"-->
							
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
sub saveplacard
	dim rs,userplacard
	userplacard=FilterJS(oblog.filt_astr(request.form("userplacard"),20000))
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_placard from [oblog_user] where userid="&oblog.logined_uid,conn,1,3
	rs(0)=oblog.filtpath(userplacard)
	rs.update
	rs.close
	dim blog
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_placard oblog.logined_uid
	set rs=nothing
	oblog.showok "修改公告成功",""
end sub

sub placard
dim rs
set rs=oblog.execute("select user_placard from [oblog_user] where userid="&oblog.logined_uid)
%>
<h1 style='font-size:13px;'>修改站点公告</h1>
<div id="list">
	<ul class="list_edit">
	你可以在这里放置你的照片，关于你的介绍，或者你愿意放上去的任何信息。<hr>
	</ul>
	<form name="oblogform" method="post" action="HomestayBackctrl-PassSettingCooperate.asp" onSubmit="submits(); ">
	<input type="hidden" name="userplacard" value="<%
if rs(0)<>"" then
	response.Write  Server.HtmlEncode(rs(0))
else
	response.Write ""
end if
%>" id="edit" />
	<!--#include file="edit.asp"-->
	<ul class="list_edit">
	<input name="Action" type="hidden" id="Action" value="saveplacard" /> 
	<input type="submit" name="Submit" value="提交修改" />
	</ul>
	</form>
</div>
<%
set rs=nothing
end sub

sub savelinks
	dim rs,links
	links=oblog.filt_astr(request.form("links"),20000)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_links from [oblog_user] where userid="&oblog.logined_uid,conn,1,3
	rs(0)=oblog.filtpath(links)
	rs.update
	rs.close
	dim blog
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_links oblog.logined_uid
	set rs=nothing
	oblog.showok "修改友情连接成功",""
end sub

sub links
dim rs
set rs=oblog.execute("select user_links from [oblog_user] where userid="&oblog.logined_uid)
%>
<h1 style='font-size:13px;'>修改友情连接</h1>
<div id="list">
	<ul class="list_edit">
	你可以先输入文字或者图片，然后用 <image src="images/wlink.gif" /> 按钮插入超级连接。<hr>
	</ul>
	<form name="oblogform" method="post" action="HomestayBackctrl-PassSettingCooperate.asp" onSubmit="submits(); ">
	<input type="hidden" name="links" value="<%
if rs(0)<>"" then
	response.Write  Server.HtmlEncode(rs(0))
else
	response.Write ""
end if
%>" id="edit" />
	<!--#include file="edit.asp"-->
	<ul class="list_edit">
	<input name="Action" type="hidden" id="Action" value="savelinks" /> 
	<input type="submit" name="Submit" value="提交修改" />
	</ul>
	</form>
</div>
<%
set rs=nothing
end sub



sub blogpassword()
%>
<h1 style='font-size:13px;'>整个blog站点的密码设定</h1>
<div id="list">	
	<ul class="list_edit">
	全站加密后，你所有文章都需要通过密码验证后才能访问。</br>
	注意：设置完密码以后，需要<a href="user_update.asp">重新发布全站</a>！
	<hr>
	<form name="form1" method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=addblogpassword" onSubmit='return Verify()'>
	密码：
	<input type="password" name="password" /> 
	<input type="submit" name="Submit" value="全站加密" />        
	</form>
	<br />
	<form name="form2" method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=unblogpassword" />
	<input type="submit" name="Submit" value="解除我站点的密码保护" />
	</form>
	</ul>
</div>
<SCRIPT language="javascript">
function Verify()
{
	v = del_space(document.form1.password.value);
     if (v.length == 0)
     {
        alert("密码不能为空!");
	return false;
     }
	
	return true;
}
</SCRIPT>

<%end sub

sub addblogpassword
	dim password,strtmp,upath,blog
	password=md5(trim(request("password")))
	oblog.execute("update [oblog_user] set blog_password='"&password&"' where userid="&oblog.logined_uid)
	oblog.execute("update [oblog_log] set blog_password=1 where userid="&oblog.logined_uid)
	'strtmp="<script language=javascript>"&vbcrlf
	'strtmp=strtmp&"var blogpassword=GetCookie('blogpassword')"&vbcrlf
	'strtmp=strtmp&"if (blogpassword != '"&password&"')"&vbcrlf
	'strtmp=strtmp&"{ window.location.replace('"&blogdir&"chkblogpassword.asp?userid="&oblog.logined_uid&"');}"&vbcrlf
	'strtmp=strtmp&"<script>"&vbcrlf
	upath=server.MapPath(oblog.logined_udir)
	Call oblog.BuildFile(upath&"\"&oblog.logined_ufolder&"\inc\chkblogpassword.htm","")	
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_index 0
	blog.update_message 0
	set blog=nothing
	oblog.showok "设置整站密码成功,请重新更新全站才可获得安全的加密保护！",""
end sub

sub unblogpassword
	dim upath,blog
	oblog.execute("update [oblog_user] set blog_password='' where userid="&oblog.logined_uid)
	oblog.execute("update [oblog_log] set blog_password=0 where userid="&oblog.logined_uid)
	upath=server.MapPath(oblog.logined_udir)
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_index 0
	blog.update_message 0
	set blog=nothing
	oblog.showok "取消密码成功,请重新更新全站才可全部解密！",""
end sub

sub sitesetup()
dim rs,sstr,sublist,us,i,u_frame
if oblog.setup(54,0)=0 then sstr="disabled"
set rs=oblog.execute("select * from oblog_user where userid="&oblog.logined_uid)
us=rs("user_info")
if us="" or isnull(us) then
	sublist=0
	u_frame=0
else
	us=split(us,"$")
	if us(0)<>"" then sublist=cint(us(0)) else sublist=0
	if us(1)<>"" then u_frame=cint(us(1)) else u_frame=1
end if
%>
<style>
	#list ul{ width: 90%;margin-bottom: 0;text-align:left;
	list-style-type:none;
	margin:0px;
	}
	#list ul li.t1 { width: 180px;text-align: left;float:left;} 
	#list ul li.t2 { width: 380px;text-align: left;} 
</style>
<h1 style='font-size:13px;'>修改我的站点信息</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
	<%if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then%>
    <ul>
	<li class="t1">域名：</li>
	<li class="t2"><input name="user_domain" type="text" value="<%=oblog.filt_html(trim(rs("user_domain")))%>" size=10 maxlength=20 <%=sstr%> /> <select name="user_domainroot" <%=sstr%>><%=oblog.type_domainroot(rs("user_domainroot"))%></select><input type="hidden" name="old_userdomain" value="<%=oblog.filt_html(rs("user_domain"))%>"></li>
    </ul>
	<%end if%>
	<%if true_domain=1 and oblog.logined_ulevel=8 then%>
    <ul>
	<li class="t1">绑定我的顶级域名：</li>
	<li class="t2"><input name="custom_domain" type="text" value="<%=oblog.filt_html(trim(rs("custom_domain")))%>" size=20 maxlength=50 <%=sstr%> /> </li>
    </ul>
	<%end if%>
    <ul>
	<li class="t1">站点名称：</li>
	<li class="t2"><input name="blogname" type="text" value="<%=oblog.filt_html(rs("blogname"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul> 
	<li class="t1">站点类别：</li>
	<li class="t2"><select name="user_classid" ><%=oblog.show_class("user",rs("user_classid"),0)%></select></li>
    </ul>
	<ul> 
	<li class="t1">默认编辑器：</li>
	<li class="t2"><input type="radio" value="0" name="isubbedit" <%if rs("isubbedit")=0 then response.write "checked"%> />
	html编辑器&nbsp;&nbsp; <input type=radio value="1" name="isubbedit" <%if rs("isubbedit")=1 then response.write "checked"%> />ubb编辑器</li>
    </ul>

     <ul> 
	<li class="t1">允许将我加入用户团队：</li>
	<li class="t2"><input type="radio" value="1" name="en_blogteam" <%if rs("en_blogteam")<>0 then response.write "checked"%> />
        是 &nbsp;&nbsp; <input type=radio value="0" name="en_blogteam" <%if rs("en_blogteam")=0 then response.write "checked"%> />否</li>
    </ul>
	 <ul> 
	<li class="t1">是否隐藏转向url：</li>
	<li class="t2"><input type="radio" value="1" name="hideurl" <%if rs("hideurl")=1 then response.write "checked"%> />
        是 &nbsp;&nbsp; <input type="radio" value="0" name="hideurl" <%if rs("hideurl")=0 then response.write "checked"%> />否</li>
    </ul>
	<ul> 
	<li class="t1">分类文章是否以列表显示：</li>
	<li class="t2"><input type="radio" value="1" name="sublist" <%if sublist=1 then response.write "checked"%> />
        是 &nbsp;&nbsp; <input type="radio" value="0" name="sublist" <%if sublist=0 then response.write "checked"%> />否</li>
    </ul>
	<ul> 
	<li class="t1">管理界面是否分栏显示：</li>
	<li class="t2"><input type="radio" value="1" name="u_frame" <%if u_frame=1 then response.write "checked"%> />
        是 &nbsp;&nbsp; <input type="radio" value="0" name="u_frame" <%if u_frame=0 then response.write "checked"%> />否</li>
    </ul>

	 <ul> 
	<li class="t1">文章默认部分显示字数：</li>
	<li class="t2"><input name="user_showlogword_num" type="text"  value="<%=rs("user_showlogword_num")%>" size="5" /></li>
    </ul>
	 <ul> 
	<li class="t1">每页显示文章篇数：</li>
	<li class="t2"><input name="user_showlog_num" type="text" id="user_showlog_num" value="<%=rs("user_showlog_num")%>" size="5" /></li>
    </ul>
	<ul> 
	<li class="t1">每行显示图片数：</li>
	<li class="t2"><input name="user_photorow_num" type="text" id="user_photorow_num" value="<%=rs("user_photorow_num")%>" size="5" /></li>
    </ul>
	 <ul> 
	<li class="t1">显示最新回复条数：</li>
      <li class="t2"><input name="user_shownewcomment_num" type="text" value="<%=rs("user_shownewcomment_num")%>" size="5" /></li>
    </ul>
    <ul> 
	<li class="t1">显示最新文章条数：</li>
	<li class="t2"><input name="user_shownewlog_num" type="text" value="<%=rs("user_shownewlog_num")%>" size="5" /></li>
    </ul>
    <ul> 
	<li class="t1">显示最新留言条数：</li>
      <li class="t2"><input name="user_shownewmessage_num" type="text" value="<%=rs("user_shownewmessage_num")%>" size="5" /></li>
	  </ul>
	 <ul>
	<li class="t1">文章评论排列顺序：</li>
	<li class="t2"> <input type="radio" value="1" name="comment_isasc" <%if rs("comment_isasc")=1 then response.write "checked"%> />时间顺序 &nbsp;&nbsp; <input type="radio" value="0" name="comment_isasc" <%if rs("comment_isasc")=0 then response.write "checked"%> />时间倒序</li>
    </ul>
	<ul> 
	<li class="t1">站点简介：</li>
	<li class="t2"><textarea name="siteinfo" cols="40" rows="5"><%=oblog.filt_html(rs("siteinfo"))%></textarea></li>
    </li>
	</ul>

    <ul>
	<li class="t1"></li>
	<li class="">
	<input name="action" type="hidden" value="savesitesetup" /> 
	<input type="submit" value=" 更新 " />
	</li>
    </ul>

  </form>
</div>
<%
set rs=nothing
end sub

sub savesitesetup()
	dim user_domain,user_domainroot,blogname,user_classid,en_blogteam,user_showlogword_num,user_showlog_num,user_shownewcomment_num,user_shownewlog_num,user_shownewmessage_num,hideurl,comment_isasc,siteinfo,user_photorow_num,custom_domain
	dim rs,blog
	user_domain=Lcase(trim(request("user_domain")))
	user_domainroot=trim(request("user_domainroot"))
	blogname=EncodeJP(trim(request("blogname")))
	user_classid=trim(request("user_classid"))
	en_blogteam=trim(request("en_blogteam"))
	user_showlogword_num=trim(request("user_showlogword_num"))
	user_showlog_num=trim(request("user_showlog_num"))
	user_photorow_num=trim(request("user_photorow_num"))
	user_shownewcomment_num=trim(request("user_shownewcomment_num"))
	user_shownewlog_num=trim(request("user_shownewlog_num"))
	user_shownewmessage_num=trim(request("user_shownewmessage_num"))
	hideurl=trim(request("hideurl"))
	comment_isasc=trim(request("comment_isasc"))
	siteinfo=EncodeJP(trim(request("siteinfo")))
	custom_domain=trim(request("custom_domain"))
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 and oblog.setup(54,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>20  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字！")
		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
	end if
	if oblog.strLength(siteinfo)>255 then oblog.adderrstr("站点简介不能大于255个字符！")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("真实姓名中含有系统不允许的字符！")
	if not IsNumeric(user_showlogword_num) then
		oblog.adderrstr("文章默认部分显示字数必须为数字！")
	end if	
	if not IsNumeric(user_showlog_num) then
		oblog.adderrstr("每页显示文章数量必须为数字！")
	else
		user_showlog_num=clng(user_showlog_num)
		if user_showlog_num>50 then oblog.adderrstr("每页显示文章数量必须小于50！")
	end if
	if not IsNumeric(user_photorow_num) then
		oblog.adderrstr("每行显示图片数量必须为数字！")
	else
		user_photorow_num=clng(user_photorow_num)
		if user_photorow_num>50 then oblog.adderrstr("每行显示图片数量必须小于50！")
	end if
	
	if not IsNumeric(user_shownewcomment_num) then
		oblog.adderrstr("显示最新回复条数必须为数字！")
	else
		user_shownewcomment_num=clng(user_shownewcomment_num)
		if user_shownewcomment_num>50 then oblog.adderrstr("显示最新回复条数不能大于50！")
	end if
	
	if not IsNumeric(user_shownewlog_num) then
		oblog.adderrstr("显示最新文章条数必须为数字！")
	else
		user_shownewlog_num=clng(user_shownewlog_num)
		if user_shownewlog_num>50 then oblog.adderrstr("显示最新文章条数不能大于50！")
	end if
	
	if not IsNumeric(user_shownewmessage_num) then
		oblog.adderrstr("显示最新留言条数必须为数字！")
	else
		user_shownewmessage_num=clng(user_shownewmessage_num)
		if user_shownewmessage_num>50 then oblog.adderrstr("显示最新留言条数不能大于50！")
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 and oblog.setup(54,0)=1 then
		set rs=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"' and userid<>"&oblog.logined_uid)
		if not rs.eof or not rs.bof then oblog.adderrstr("系统中已经有这个域名存在，请更改域名！")
	end if
	if true_domain=1 and oblog.logined_ulevel=8 and custom_domain<>"" then
		if oblog.chk_badword(custom_domain)>0 then oblog.adderrstr("绑定的顶级域名中含有系统不允许的字符！")
		set rs=oblog.execute("select userid from oblog_user where custom_domain='"&oblog.filt_badstr(custom_domain)&"'"&" and userid<>"&oblog.logined_uid)
		if not rs.eof or not rs.bof then oblog.adderrstr("系统中已经有其他人绑定了这个顶级域名，请更改域名或者联系管理员！")		
	end if
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	if hideurl="" or isnull(hideurl) then hideurl=0
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from oblog_user where userid="&oblog.logined_uid,conn,1,3
	if not rs.eof then
		rs("blogname")=oblog.filt_astr(blogname,50)
		if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 and oblog.setup(54,0)=1 then
			rs("user_domain")=user_domain
			rs("user_domainroot")=user_domainroot
		end if
		if oblog.setup(54,0)=1 then
			if true_domain=1 and oblog.logined_ulevel=8 then
				rs("custom_domain")=custom_domain
			end if
		end if
		rs("user_classid")=user_classid
		rs("en_blogteam")=en_blogteam
		rs("user_showlogword_num")=user_showlogword_num
		rs("user_showlog_num")=user_showlog_num
		rs("user_photorow_num")=user_photorow_num
		rs("user_shownewcomment_num")=user_shownewcomment_num
		rs("user_shownewlog_num")=user_shownewlog_num
		rs("user_shownewmessage_num")=user_shownewmessage_num
		rs("hideurl")=hideurl
		rs("comment_isasc")=comment_isasc
		rs("siteinfo")=siteinfo
		rs("isubbedit")=request.Form("isubbedit")
		rs("user_info")=cint(request.Form("sublist"))&"$"&cint(request.Form("u_frame"))
		rs.update
		rs.close
	end if
	set rs=nothing
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_blogname
	set blog=nothing
	oblog.showok "保存设置成功!",""		
end sub

sub saveblogstar()
	dim rs,picurl,bloginfo,blogname
	picurl=trim(request("picurl"))
	bloginfo=trim(request("bloginfo"))
	blogname=trim(request("blogname"))
	if picurl="" or oblog.strLength(picurl)>250  then oblog.adderrstr("图片连接地址不能为空,且不能大于250个字符！")
	if bloginfo="" or oblog.strLength(bloginfo)>250  then oblog.adderrstr("站点介绍不能为空,且不能大于250个字符！")
	if blogname="" or oblog.strLength(blogname)>250  then oblog.adderrstr("用户名不能为空,且不能大于50个字符！")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from oblog_blogstar Where userid=" & oblog.logined_uid,conn,1,3
	If rs.Eof Then
		rs.addnew	
		rs("userid")=oblog.logined_uid
	End If
	rs("picurl")=picurl
	rs("info")=bloginfo
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		rs("userurl")="http://"&oblog.logined_udomain
	else
		rs("userurl")=oblog.setup(3,0)&oblog.logined_udir&"/"&oblog.logined_ufolder&"/index."&f_ext
	end if
	rs("blogname")=blogname
	rs.update
	rs.close
	set rs=nothing
	oblog.showok "提交完成，请等待管理员审核通过。",""
end sub

sub blogstar()
	Dim rs,strTitle,strBlogName,strPicUrl,strBlogInfo,intState
	Set rs=Server.CreateObject("ADODB.RecordSet")	
	rs.open "select * from oblog_blogstar Where userid=" & oblog.logined_uid,conn,1,1
	If rs.Eof Then
		strTitle="你目前还没有申请"
		intState=-1
	Else
		strPicUrl=rs("picurl")
		strBlogName=rs("blogname")
		strBlogInfo=rs("info")
		intState=rs("ispass")
		If intState=1 Then
			strTitle="你目前已经是用户明星，资料不可更改"
		Else	
			strTitle="你目前正在等待审核中，可以修改之前提交的资料"
		End If
	End If
	rs.Close
	Set rs = Nothing
		
%>
<style>
	#list ul{ width: 100%;margin-bottom: 0;}
	#list ul li.t1 { width: 100px;text-align: right;} 
	#list ul li.t2 { width: 400px;text-align: left;} 
</style>
<h1 style='font-size:13px;'>申请用户明星: <%=strTitle%></h1>
<div id="list">
	
	<ul class="list_edit">
	图片地址可以放上你的照片，站点logo或者站点缩略图；blog介绍请不要超过50字。<br />
	注：图片尺寸最好缩小到130*100左右，以便管理员操作。
	<hr /></ul>
	<form  method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=saveblogstar">
	<ul>
	<li class="t1">用户名字：</li>
	<li class="t2"><input type="text" maxlength="50" size="60" name="blogname" value="<%=strBlogname%>" /></li>
	</ul>
	<ul>
	<li class="t1">图片地址：</li>
	<li class="t2"><input type="text" maxlength="250" size="60" name="picurl" value="<%=strPicUrl%>" /></li>
	</ul>
	<ul>
	<li class="t1">blog介绍：</li>
	<li class="t2"><textarea name="bloginfo" cols="60" rows="5"><%=strBlogInfo%></textarea></li></ul>
	<ul class="list_bot">
	<%
	Select Case intState
		Case -1
	%>
		<input type="submit" value="提交申请资料"></ul>
		<%
		Case 0
		%>
		<input type="submit" value="修改申请资料"></ul>
		<%
		Case 1
		%>
		已经被确认为用户明星，资料不可更改，如果需要修改或撤销请与管理员联系。</BR>
		
	<%End Select%>
	</ul>
	</form>
</div>
<%
end sub

sub userpassword
if is_ot_user=1 then
	response.Redirect(ot_modifypass1)
end if
dim rs
set rs=oblog.execute("select question from oblog_user where userid="&oblog.logined_uid)
%>
<style>
#list ul{ width: 90%;margin-bottom: 0;text-align:left;
	list-style-type:none;
	margin:0px;
	}
	#list ul li.t1 { width: 150px;text-align: left;float:left;} 
	#list ul li.t2 { width: 380px;text-align: left;}  
</style>

<% If Request("pass")=1 Then%>
<h1 style='font-size:13px;'>修改我的密码</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul>
      <li class="t1">原密码：</li>
      <li class="t2"><input name="oldpassword" type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
      <li class="t1">新密码：</li>
      <li class="t2"><input name="newpassword"  type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	 <ul> 
      <li class="t1">重复密码：</li>
      <li class="t2"><input name="newpassword1" type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
	<li class="t1"></li>
	<li class="t2">
	<input name="action" type="hidden" value="saveuserpassword" /> 
	<input type=submit value=" 修改密码 " />
	</li>
    </ul>
	</form>
</div>
<% End If%>
<% If Request("pass")<>1 Then%>

<h1 style='font-size:13px;'>更新我的找回密码设置</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul> 
      <li class="t1">请输入登陆密码：</li>
      <li class="t2"><input name="password"  type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
    <ul> 
      <li class="t1">密码提示问题：</li>
      <li class="t2"><input name=question  onFocus="this.select(); " type="text" size="30" maxlength="20" value="<%=oblog.filt_html(rs(0))%>"><font color="#FF0000"> *</font>&nbsp;&nbsp;&nbsp;(如需修改，请填写)</li>
    </ul>
	 <ul> 
      <li class="t1">找回密码答案：</li>
      <li class="t2"><input name=answer  type="text" size="30" maxlength="20"><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
		<li class="t1"></li>
		<li class="t2">
	  	<input name="action" type="hidden" value="saveuserlostpass"> 
        <input type=submit value=" 更新 ">
		</li>
    </ul>
	</form>
</div>
<% End If%>
<%
end sub

sub saveuserlostpass()
	dim password,question,answer,rs
	password=trim(request("password"))
	question=trim(request("question"))
	answer=trim(request("answer"))
	if password="" then oblog.adderrstr("错误：登陆密码不能为空!")
	if question="" then oblog.adderrstr("错误：提示问题不能为空！")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select question,answer from oblog_user where userid="&oblog.logined_uid&" and password='"&md5(password)&"'",conn,1,3
	if rs.eof then
		rs.close
		set rs=nothing
		oblog.adderrstr("错误：登陆密码输入错误!")
		oblog.showusererrCooperate
		exit sub
	else
		rs("question")=question
		if answer<>"" then rs("answer")=md5(answer)
		rs.update
		rs.close
		set rs=nothing
		oblog.showok "修改找回密码资料成功！",""
	end if

end sub

sub saveuserpassword
	dim oldpassword,newpassword,rs
	oldpassword=trim(request("oldpassword"))
	newpassword=trim(request("newpassword"))
	if oldpassword="" then oblog.adderrstr("错误：原密码不能为空!")
	if newpassword="" or oblog.strLength(newpassword)>14 or oblog.strLength(newpassword)<4 then oblog.adderrstr("错误：新密码不能为空(不能大于14小于4)！")
	if newpassword<>trim(request("newpassword1")) then oblog.adderrstr("错误：重复密码输入错误!")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select password from oblog_user where userid="&oblog.logined_uid&" and password='"&md5(oldpassword)&"'",conn,1,3
	if rs.eof then
		rs.close
		set rs=nothing
		oblog.adderrstr("错误：原密码输入错误!")
		oblog.showusererrCooperate
		exit sub
	else
		rs("password")=md5(newpassword)
		rs.update
		rs.close
		set rs=nothing
		oblog.savecookie oblog.logined_uname,newpassword,0,""
		oblog.showok "修改密码成功,下次需要重新登录！",""
	end if
end sub

sub userinfo
dim rs
set rs=oblog.execute("select * from oblog_user where userid="&oblog.logined_uid)
%>
<style>

#list ul{ width: 90%;margin-bottom: 0;text-align:left;
	list-style-type:none;
	margin:0px;
	}
	#list ul li.t1 { width: 150px;text-align: left;float:left;} 
	#list ul li.t2 { width: 380px;text-align: left;} 
</style>
<h1 style='font-size:13px;'>修改我的个人信息</h1>
<div id="list">
<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul>
	<li class="t1">用户id：</li>
	<li class="t2"><%=rs("userName")%></li>
    </ul>
    <ul>
	<li class="t1">昵称：</li>
	<li class="t2"><input name="nickname" type="text" value="<%=oblog.filt_html(rs("nickname"))%>" size="30" maxlength="20" /><input type="hidden" name="o_nickname" value="<%=oblog.filt_html(rs("nickname"))%>" /></li>
    </ul>
    <ul>
	<li class="t1">真实姓名：</li>
	<li class="t2"><input name="truename"   type="text" value="<%=oblog.filt_html(rs("truename"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul>
	<li class="t1">性别：</li>
	<li class="t2"><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then response.write "checked"%> />
        男 &nbsp;&nbsp; <input type="radio" value="0" name="sex" <%if rs("Sex")=0 then response.write "checked"%> />女</li>
    </ul>
	 <ul>
	<li class="t1">省/市：</li>
	<li class="t2"> <%=oblog.type_city(rs("province"),rs("city"))%> </li>
    </ul>
	 <ul>
	<li class="t1">出生日期：</li>
	<li class="t2">
		<input value="<%=year(rs("birthday"))%>" name="y" size="4" maxlength="4" />年
		<input value="<%=month(rs("birthday"))%>" name="m" size="2" maxlength="2" />月
		<input value="<%=day(rs("birthday"))%>"  name="d" size="2" maxlength="2" />日
	  </li>
    </ul>
	 <ul>
	<li class="t1">职业：</li>
	<li class="t2"> <%oblog.type_job(rs("job"))%></li>
    </ul>
  <ul>
	<li class="t1">Email：</li>
	<li class="t2"><input name="Email" value="<%=oblog.filt_html(rs("userEmail"))%>" size="30" maxlength="50" /> 
      </li>
    </ul>
    <ul>
	<li class="t1">主页：</li>
	<li class="t2"><input maxlength="100" size="30" name="homepage" value="<%=oblog.filt_html(rs("Homepage"))%>" /></li>
    </ul>
    <ul>
	<li class="t1">QQ号码：</li>
	<li class="t2"><input name="qq" value="<%=oblog.filt_html(rs("qq"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul>
	<li class="t1">MSN：</li>
	<li class="t2"><input name="msn" value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></li>
    </ul>
	 <ul>
	<li class="t1">电话：</td>
	<li class="t2"><input name="tel" value="<%=oblog.filt_html(rs("tel"))%>" size="30" maxlength="50" /></li>
    </ul>
	<ul>
	<li class="t1">通信地址：</li>
	<li class="t2"><input name="address" value="<%=oblog.filt_html(rs("address"))%>" size="40" maxlength="250" /></li>
    </ul>
	  <ul>
	  <li class="t1"></li>
	  <li class="t2"><input name="action" type="hidden" value="saveuserinfo" /> 
        <input type=submit value=" 更新 " /></li>
		</ul>
  </form>
</div>
<%
set rs=nothing
end sub

sub saveuserinfo
	dim rs,nickname,email,birthday
	nickname=trim(request("nickname"))
	email=trim(request("email"))
	birthday=trim(request("y"))&"-"&trim(request("m"))&"-"&trim(request("d"))
	if oblog.chk_regname(nickname) then oblog.adderrstr("此昵称系统不允许注册！")
	if oblog.chk_badword(nickname)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
	if oblog.strLength(nickname)>50 then oblog.adderrstr("昵称不能不能大于50字符！")
	if oblog.setup(25,0)=0 and nickname<>"" and nickname<>trim(request("o_nickname")) then
		set rs=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
		if not rs.eof or not rs.bof then oblog.adderrstr("系统中已经有这个昵称存在，请更改昵称！")		
	end if
	if birthday = "--" then
		birthday=""
	else
		if not isdate(birthday) then oblog.adderrstr("生日日期格式错误！")
		if clng(trim(request("y")))>2060 then oblog.adderrstr("生日年份过大！")
		if clng(trim(request("y")))<1900 then oblog.adderrstr("生日年份过小！")
	end if
	if email="" then oblog.adderrstr("电子邮件地址不能为空！")
	if not oblog.IsValidEmail(email) then oblog.adderrstr("电子邮件地址格式错误！")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from oblog_user where userid="&oblog.logined_uid,conn,1,3
	if not rs.eof then
		rs("nickname")=oblog.filt_astr(nickname,50)
		rs("truename")=oblog.filt_astr(request("truename"),20)
		rs("sex")=clng(request("sex"))
		rs("province")=request("province")
		rs("city")=request("city")
		if birthday<>"" then rs("birthday")=birthday
		rs("job")=request("job")
		rs("useremail")=oblog.filt_astr(email,50)
		rs("homepage")=oblog.filt_astr(request("homepage"),100)
		rs("qq")=oblog.filt_astr(request("qq"),20)
		rs("msn")=oblog.filt_astr(request("msn"),50)
		rs("tel")=oblog.filt_astr(request("tel"),50)
		rs("address")=oblog.filt_astr(request("address"),250)
		rs.update
		rs.close
	end if
	set rs=nothing
	oblog.showok "保存资料成功!",""
end sub

%>