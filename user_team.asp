<!--#include file="user_top.asp"-->
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
				<span>客户咨询热线</span>：<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
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
					
					
					
						<center><h2>用户团队</h2></center>
			

<%
dim action,id
action=request.QueryString("action")
id=trim(request("id"))
select case action
	case "addotheruser"
		call addotheruser()
	case "del"
		call delotheruser()
	case else
		call main
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
								栏目发布管理
							</p>
							
							<!--#include file="menu/inc-menu.asp"-->
							
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

sub addotheruser()
	dim otheruser,rs
	otheruser=oblog.filt_badstr(Trim(Request.Form("otheruser")))
	If otheruser="" Then
		oblog.adderrstr("用户名称不能为空！")
		oblog.showusererr
		exit sub
	End If	 
	set rs=oblog.execute("select en_blogteam,userid from [oblog_user] where username='"&otheruser&"'")
	if rs.eof then
		set rs=nothing
		oblog.adderrstr("无此用户！")
		oblog.showusererr
		exit sub
	else
		if rs(0)<>1 then
			set rs=nothing
			oblog.adderrstr("该用户设置为不允许被加入团队！")
			oblog.showusererr
			exit sub
		else
			oblog.execute("insert into [oblog_blogteam] (mainuserid,otheruserid) values ("&oblog.logined_uid&","&rs(1)&")")
			set rs=nothing
			oblog.showok "加入团队成员成功！","user_team.asp"		
		end if		
	end if
end sub

sub delotheruser
	if id="" then
		oblog.adderrstr("请指定删除参数！")
		oblog.showusererr
		exit sub
	end if
	id=clng(id)
	oblog.execute("delete  from [oblog_blogteam] where id="&id&" and ( mainuserid="&oblog.logined_uid&" or otheruserid="&oblog.logined_uid&" )")
	oblog.showok "操作成功！",""
end sub
sub main()
%>
<style type="text/css">
<!--
	#list ul{ width: 85%;margin-left:-0px;margin-bottom:10px;} 
	#list ul li.t1 { width: 50px} 
	#list ul li.t2 { width: 300px} 
	#list ul li.t3 { width: 150px} 
-->
</style>
<h1 style='font-size:13px;'>加入朋友到我的用户团队</h1>
<div id="list">	
	<ul class="list_edit" style="margin-left:0">
	可以让好友把文章发布到我的blog，现在就邀请把!<hr>
	<form name="form1" method="post" action="user_team.asp?action=addotheruser">
	用户名： 
	<input name="otheruser" type="text" id="otheruser" maxlength="20" />            
	<input type="submit" value="邀请" />			             
	</form>
	</ul>
</div>
<h1 style='font-size:13px;'>我的团队管理</h1>
<div id="list">
	<ul class="list_top tr01" style="padding-left:3px;">
	<li class="t1" style="width:60px;float:left;">排序</li>
	<li class="t2" style="width:130px;float:left;">团队成员</li>
	<li class="t3" style="width:130px;float:left;">管理操作</li>
	</ul>
<%
dim rs,i
set rs=oblog.execute("select username,oblog_blogteam.id from oblog_user,oblog_blogteam where userid=otheruserid and mainuserid="&oblog.logined_uid)
while not rs.eof
	i=i+1 
%>
	<ul class="list_content tr01" style="padding-left:3px;"> 
	<li class="t1" style="width:60px;float:left;"><%=i%></li>
	<li class="t2" style="width:130px;float:left;"><%="<a href=""blog.asp?name="&rs(0)&""" target=""_blank"">"&rs(0)&"</a>"%></li>
	<li class="t3" style="width:130px;float:left;"><a href="user_team.asp?action=del&id=<%=rs(1)%>" <%="onClick='return confirm(""删除后，此成员所有文章将不会在你的blog中显示,确定要删除吗？"");'"%>>删除此用户成员</a></li></ul>
<%
	rs.movenext
wend 
set rs=nothing
%>
</div>

<h1 style='font-size:13px;'>我加入的团队</h1>
<div id="list">
	<ul class="list_top tr01" style="padding-left:3px;">
	<li class="t1" style="width:60px;float:left;">排序</li>
	<li class="t2" style="width:110px;float:left;">团队网站名</li>
	<li class="t3" style="width:120px;float:left;">团队管理员</li>
	<li class="t4" style="width:130px;float:left;">管理操作</li>
	</ul>
  <%
set rs=oblog.execute("select oblog_blogteam.id,oblog_blogteam.mainuserid,oblog_user.blogname,oblog_user.user_dir,oblog_user.username from oblog_blogteam,oblog_user where otheruserid="&oblog.logined_uid&" and [oblog_user].userid=oblog_blogteam.mainuserid")
i=0
while not rs.eof
	i=i+1 
%>
	<ul class="list_content tr03" style="padding-left:3px;">
	<li class="t1" style="width:60px;float:left;"><%=i%></li>
	<li class="t2" style="width:110px;float:left;"><a href="<%=rs(3)&"/"&rs(1)&"/index."&f_ext%>" target="_blank"><%=rs(2)%></a></li>
	<li class="t3" style="width:120px;float:left;"><a href="<%=rs(3)&"/"&rs(1)&"/index."&f_ext%>" target="_blank"><%=rs(4)%></a></li>
	<li class="t4" style="width:130px;float:left;"><a href="user_team.asp?action=del&id=<%=rs("id")%>" <%="onClick='return confirm(""确认退出吗？"");'"%>>退出此用户团队</a></li>
	</ul>  
<%
rs.movenext
wend 
%>
</div>
<%end sub%>

