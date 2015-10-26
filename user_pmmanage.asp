<!--#include file="user_top.asp"-->
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
				<span id="joincompany_login"><a href="/login.asp">企业登录</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">企业加盟</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>客户咨询热线</span>：<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
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
						padding:0px 3px 0px 6px;
						text-align:center;
						height:18px;
						border:1px solid #999999;
						
						}
							td.tr02-1 {
							width:30px;
							}
						 	td.tr02-2 {
							width:40px;
							}
							td.tr02-3 {
							width:210px;
							text-align:left;
							}
							td.tr02-4 {
							width:100px;
							}
							td.tr02-5 {
							width:110px;
							}
							td.tr02-6 {
							width:80px;
							}
							td.tr02-7 {
							width:50px;
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
					
					
						<center><h2>MyHomestay 消息管理</h2></center>
						


    <div class="content_top">
            <div class="side_d1 side11"></div>
            <div class="side_d2 side21"></div>
           
    </div>
	<div class="submenu_content" style="display:block;">
		  &nbsp;<img src="/images/mailmanage.gif" /><a href="user_pmmanage.asp" title="接收住家网写给我的所有消息">收件箱</a>
		| &nbsp;<img src="/images/mailsend.gif" /><a href="user_pmmanage.asp?usersearch=1" title="我写给住家网的所有消息保存">发件箱</a>
		| &nbsp;<img src="/images/mail_write.gif" /><a href="javascript:openScript('user_pm.asp?action=send',450,400)" title="写信与住家网进行沟通">我要写信</a>
	</div>
	
	
<%const MaxPerPage=20
dim strFileName,totalPut,TotalPages
dim rs,sql,blog
dim id,usersearch,action
usersearch=trim(request("usersearch"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
if usersearch="" then
	usersearch=0
else
	usersearch=clng(usersearch)
end if
strFileName="user_pmmanage.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
select case action
	case "add"
	call add()
	case "saveadd" 
	call saveadd()
	case "del" 
	call del()
	case else	
	call main()
end select
set rs=nothing
set blog=nothing
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
sub main()
%>
<style type="text/css">
<!--
	#list ul{ width: 100%} 
	#list ul li.t1 { width: 50px} 
	#list ul li.t2 { width: 50px} 
	#list ul li.t3 { width: 160px} 
	#list ul li.t4 { width: 80px} 
	#list ul li.t5 { width: 80px} 
	#list ul li.t6 { width: 100px} 
	#list ul li.t7 { width: 50px} 
-->
</style>

<%
	dim strGuide,ssql
	ssql="id,sender,incept,topic,addtime,isguest,isreaded"
	strGuide="<h1 style='font-size:13px;'>当前选择&nbsp;&gt;&gt;&nbsp;"
	select case usersearch
		case 0
			sql="select "&ssql&" from oblog_pm where incept='"&oblog.logined_uname&"' and delr=0 order by id desc"
			strGuide=strGuide & "收件箱"
		case 1
			sql="select "&ssql&" from oblog_pm where sender='"&oblog.logined_uname&"' and dels=0 order by id desc"
			strGuide=strGuide & "发件箱"
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (共有0条消息)</h1><div id=""list""></div>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & " (共有" & totalPut & "条消息)</h1><div id=""list"">"
		response.write strGuide
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if
    	end if
	    if currentPage=1 then
        	showContent
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"条消息")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"条消息")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"条消息")
	    	end if
		end if
	end if
	rs.Close
	if totalPut>oblog.setup(53,0) then
			'oblog.execute("update oblog_pm set delr=1 where id not in (select top "&oblog.setup(53,0)&" id from  oblog_pm where incept='"&oblog.logined_uname&"' order by id desc ) and incept='"&oblog.logined_uname&"'")
		oblog.execute("delete from oblog_pm where delr=1 and dels=1")
		Response.Write"<script language=JavaScript>alert(""您的信箱已满，请及时清理！"");</script>"
	end if
	set rs=Nothing
end sub

sub showContent()
   	dim i,freen
    i=0
	freen=oblog.setup(53,0)-totalPut
%>
	
	
	<div class="list_right">
	<form name="myform" method="Post" action="user_pmmanage.asp?usersearch=<%=usersearch%>" onSubmit="return confirm('确定要执行选定的操作吗？');">
<table>
	<tr class="tr01">
		<td class="tr02-1">选中</td><%if usersearch=0 then%><td class="tr02-2">状态</td><%end if%><td class="tr02-3" align="center" style="text-align:center;">消息标题 </td><td class="tr02-4">收件人</td><td class="tr02-5">发件人</td><td class="tr02-6">发送时间</td><td class="tr02-7">操作</td>
	</tr>
<style>
	.list_content_mouserover {
	background-color:none;
	color:#666666;
	}
	.list_content {
	background:none;
	color:#666666;
	}

</style>
	<%do while not rs.eof %>
	<tr class="tr02" onMouseOver="fSetBg(this)" onMouseOut="fReBg(this)">
    	<td class="tr02-1"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("id"))%>' /></td>
		
	<%if usersearch=0 then%>
		<td class="tr02-2"><%if rs("isreaded")=1 then response.Write("<img src='/images/mail_open.gif' title='已读' />") else response.Write("<img src='/images/mailmanage.gif' title='未读' />")%></td>
	<%end if%>
	
	
	
		<td class="tr02-3"><a href="javascript:openScript('user_pm.asp?action=read&id=<%=rs("id")%>',650,468)"><%=oblog.filt_html(rs("topic"))%></a></td>
		
		<td class="tr02-4">
		<% If rs("incept")="我的住家网" Or rs("incept")="管理员" Or rs("incept")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />我的住家网
		<% Else 			
			Response.Write oblog.filt_html(rs("incept"))
		   End If%>
		   
		<%'oblog.filt_html(rs("incept"))%></td>
		
		<td class="tr02-5">
		<% If rs("sender")="我的住家网" Or rs("sender")="管理员" Or rs("sender")="myhomestay" Then %>
			<img src="/images/sitemanagers.gif" />我的住家网
		<% Else 			
			Response.Write oblog.filt_html(rs("sender"))
		   End If%>
		</td>
		
		<td class="tr02-6" title="<%=rs("addtime")%>"> <%if rs("addtime")<>"" then response.write formatdatetime(rs("addtime"),2) else response.write "&nbsp;" %></td>
		
		<td class="tr02-7">
	<%response.write "<a href='user_pmmanage.asp?action=del&id=" & rs("id") &"&usersearch="&usersearch&"' onClick='return confirm(""确定要删除吗？"");'>删除</a>"%>
		</td>
	</tr>
<%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
</table>

	<ul class="list_bot"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />全选
	<input name="action"  type="hidden" value="del" />
	<input type="submit" name="Submit" value="删除" />
	</ul>
	</form>
	</div>
</div>
<%
end sub


sub del()
	if id="" then
		oblog.adderrstr( "错误：请指定要删除的对象！")
		oblog.showusererr
		exit sub
	end if
	if instr(id,",")>0 then
		dim n,i
		id=FilterIDs(id)
		n=split(id,",")
		for i=0 to ubound(n)
			delone(n(i))
		next
	else
		delone(id)
	end if
	set rs=nothing
	oblog.showok "删除消息成功！",""
end sub

sub delone(id)
	id=clng(id)
	select case usersearch
		case 0
		sql="update oblog_pm set delr=1 where id=" & clng(id)&" and incept='"&oblog.logined_uname&"'"
		case 1
		sql="update oblog_pm set dels=1 where id=" & clng(id)&" and sender='"&oblog.logined_uname&"'"
	end select	
	oblog.execute(sql)
end sub
%>