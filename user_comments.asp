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
					
					
					
						<center><h2>评论管理</h2></center>

			
	
<%const MaxPerPage=20
dim strFileName
dim totalPut,TotalPages
dim rs, sql
dim id,usersearch,Keyword,strField
dim action,mainid
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
usersearch=trim(request("usersearch"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
mainid=oblog.filt_badstr(trim(Request("mainid")))
if usersearch="" then
	usersearch=0
else
	usersearch=Clng(usersearch)
end if
strFileName="user_comments.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
if keyword<>"" then
	strFileName="user_comments.asp?usersearch=10&keyword="&Keyword&"&Field="&strField
end if
if action="modify" then
	call modify()
elseif action="Savemodify" then
	call Savemodify()
elseif action="del" then
	call delcomment()
else
	call top()
	call main()
end if
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



<SCRIPT language=javascript>

function VerifySubmit()
{
	topic = del_space(document.oblogform.topic.value);
     if (topic.length == 0)
     {
        alert("您忘了填写题目!");
	return false;
     }
	 
	submits(); 
	if (document.oblogform.edit.value == "")
     {
        alert("请输入内容!");
	return false;
     }	
	return true;
}
</SCRIPT>
<%
sub top()
%>
<style type="text/css">
<!--
	#list ul{ width: 95%} 
	#list ul li.t1 { width: 50px} 
	#list ul li.t2 { width: 200px} 
	#list ul li.t3 { width: 120px} 
	#list ul li.t4 { width: 160px} 
	#list ul li.t5 { width: 160px} 
-->
</style>

<style type="text/css">
<!-- 
	ul.list_top{ width: 100%;padding:0; font-size:12px;}
	ul.list_top li.t1 { width: 50px;float:left;}
	ul.list_top li.t2 { width: 90px;float:left;}
	ul.list_top li.t3 { width: 50px;float:left;}
	ul.list_top li.t4 { width: 90px;float:left;}
	ul.list_top li.t5 { width: 100px;float:left;}

	 
     ul.list_content{ width: 100%;padding:0; font-size:12px;}
     ul.list_content li.t1 { width: 50px;float:left;}
     ul.list_content li.t2 { width: 90px;float:left;}
     ul.list_content li.t3 { width: 50px;float:left;}
     ul.list_content li.t4 { width: 90px;float:left;}
     ul.list_content li.t5 { width: 100px;float:left;}

-->
</style>

<div id="list" style="width:99%;text-align:left;">
	<form action="user_comments.asp" method="get">
      <h2 style="width:100%;font-size:13px; text-align:left;padding:8px;">快速查找评论：
	<select size=1 name="usersearch" onChange="javascript:submit()">
          <option value="0">列出所有评论</option>
       	  <%if oblog.logined_ulevel=9 then%><option value="1">我文章里的评论</option><%end if%>
          <option value="10" selected>请选择查询类型</option>
         
        </select>
		 　<br />搜索：
		 <select name="Field" id="Field">
        <option value="id">作者</option>
		<option value="ip">作者ip</option>
        <option value="topic" selected>评论标题</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value="搜索" />
	</h2>
	</form>
<%
end sub

sub main()
	dim strGuide,ssql
	ssql="userid,mainid,commenttopic,addtime,commentid,comment_user,addip"
	strGuide="<h1 style='font-size:13px;'>当前选择&nbsp;&gt;&gt;&nbsp;"
	select case usersearch
		case 0
			if oblog.logined_ulevel=9 then
				sql="select top 1000 "&ssql&" from [oblog_comment] order by commentid desc"
			else
				sql="select "&ssql&" from [oblog_comment] where userid="&oblog.logined_uid&" order by commentid desc"
			end if
			strGuide=strGuide & "所有评论"
		case 1
			sql="select "&ssql&" from [oblog_comment] where userid="&oblog.logined_uid&" order by commentid desc"
			strGuide=strGuide & "我文章里的评论"
		
		case 10
			if Keyword="" then
				oblog.adderrstr( "错误：关键字不能为空！")
				oblog.showusererr
				exit sub
			else
				select case strField
				case "id"
					if oblog.logined_ulevel=9 then 
						sql="select "&ssql&" from [oblog_comment] where comment_user like '%" & Keyword&"%' order by commentid desc"
					else
						sql="select "&ssql&" from [oblog_comment] where comment_user like '%" & Keyword&" %' and userid="&oblog.logined_uid&" order by commentid desc"
					end if
					strGuide=strGuide & "作者名称含有<font color=red> " & Keyword & " </font>的评论"
				case "topic"
					if oblog.logined_ulevel=9 then
						sql="select "&ssql&" from [oblog_comment] where commenttopic like '%" & Keyword & "%' order by commentid desc"
					else
						sql="select "&ssql&" from [oblog_comment] where commenttopic like '%" & Keyword & "%' and userid="&oblog.logined_uid&" order by commentid desc"
					end if
					strGuide=strGuide & "标题中含有“ <font color=red>" & Keyword & "</font> ”的评论"
				case "ip"
					if oblog.logined_ulevel=9 then 
						sql="select "&ssql&" from [oblog_comment] where addip='" & Keyword&"' order by commentid desc"
					else
						sql="select "&ssql&" from [oblog_comment] where addip='" & Keyword&"' and userid="&oblog.logined_uid&" order by commentid desc"
					end if
					strGuide=strGuide & "作者ip含有<font color=red> " & Keyword & " </font>的评论"
				end select
			end if
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (共有0个评论)</h1></div>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & " (共有" & totalPut & "个评论)</h1>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"篇评论")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"篇评论")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"篇评论")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i
    i=0
%>
<form name="myform" method="Post" action="user_comments.asp" onSubmit="return confirm('确定要执行选定的操作吗？');">
	<ul class="list_top tr01" style="line-height:15px;padding:6px;border:1px #999999 solid;">
	<li class="t1" style="float:left;">选中</li>
	<li class="t2" style="width:170px;float:left;">评论标题</li>
	<li class="t3" style="width:70px;float:left;">发布人</li>	
	<li class="t4" style="width:150px;float:left;">发布时间</li>
	<li class="t5" style="width:150px;float:left;">操作</li>
	</ul>
          <%do while not rs.EOF %>
	<ul class="list_content tr03"> 
	<li class="t1"style="float:left;"><input name='id' type='checkbox' id="id" value='<%=cstr(rs("commentid"))%>'></li>
	<li class="t2" style="width:170px;float:left;"> <a href=go.asp?logid=<%=rs("mainid")%> target=_blank><%=oblog.filt_html(rs("commenttopic"))%>  </a> </li>
	<li class="t3" style="width:70px;float:left;" title="<%="来自:"&rs("addip")%>"><%=rs("comment_user")%></li>
	<li class="t4" style="width:150px;float:left;"> <%
	if rs("addtime")<>"" then
		response.write rs("addtime")
	else
		response.write "&nbsp;"
	end if
	%> </li>
	<li class="t5" style="width:150px;float:left;"><%
		response.write "<a href='user_comments.asp?action=modify&id=" & rs("commentid") & "&re=true'>回复</a>&nbsp;"
		response.write "<a href='user_comments.asp?action=modify&id=" & rs("commentid") & "'>修改</a>&nbsp;"
		response.write "<a href='user_comments.asp?action=del&id=" & rs("commentid") &"' onClick='return confirm(""确定要删除此评论吗？"");'>删除</a>"
%> </li>
	</ul>
<%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
	<ul class="list_bot"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
              全选
	<input name="action"  type="hidden" value="del" />
	<input type="submit" name="Submit" value="删除" />
	</ul>
</form>
</div>
<%
end sub

sub modify()
	dim id
	dim rsblog,sql
	dim restr
	id=trim(request("id"))
	if id="" then
		oblog.adderrstr( "错误：参数不足！")
		oblog.showusererr
		exit sub
	else
		id=Clng(id)
	end if
	Set rsblog=Server.CreateObject("Adodb.RecordSet")
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_comment] where commentid=" & id
	else
		sql="select * from [oblog_comment] where commentid=" & id&" and userid="&oblog.logined_uid
	end if
	rsblog.Open sql,Conn,1,1
	if rsblog.bof and rsblog.eof then
		rsblog.close
		set rsblog=nothing
		oblog.adderrstr( "错误：无权限，只有blog主人才能操作！")
		oblog.showusererr
		exit sub
	end if
	if request("re")="true" then
		restr="<br><table align=center bgcolor=#f3f3f3 border=1 bordercolor=#cccccc cellPadding=2 cellSpacing=0 style='BORDER-COLLAPSE: collapse' width='90%'>"
		restr=restr&"<tbody><tr><td><p><strong>以下为blog主人的回复：</strong></p><p>&nbsp;</p></li></tr></tbody></table>"
	end if
%>
<div id="list">
<form action="user_comments.asp?action=Savemodify" method="post" name="oblogform" onsubmit="submits();">
  <h1 style='font-size:13px;'>修改评论</h1>
  <ul class="list_edit">
  标题：<input name="topic" type="text" id="topic" value="<%=oblog.filt_html(rsblog("commenttopic"))%>" size="50" maxlength="30">
  </ul>
	<input type="hidden" name="edit" id="edit"value="<%
				if rsblog("comment")<>"" then response.Write Server.HtmlEncode(rsblog("comment")+restr)%>"> 
    <!--#include file="edit.asp"-->
  <ul class="list_edit"> 
	<input type="hidden" name="id" value="<%=rsblog("commentid")%>">
	<input type="submit" name="Submit2" value="提交修改">
  </ul>
</form>
</div>
<%
	rsblog.close
	set rsblog=nothing
end sub

sub Savemodify()
	dim id,rsblogchk,blog,logid,uid
	id=clng(trim(request("id")))
	if oblog.logined_ulevel=9 then
		sql="select * from oblog_comment where commentid="&id
	else
		sql="select * from oblog_comment where commentid="&id&" and userid="&oblog.logined_uid
	end if	
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	uid=rs("userid")
	logid=rs("mainid")
	rs("commenttopic")=oblog.InterceptStr(oblog.filt_badword(trim(request("topic"))),250)
	rs("comment")=oblog.filt_badword(request("edit"))
	rs.update
	rs.close
	set rs=nothing
	set blog= new class_blog
	blog.userid=uid
	blog.update_log logid,0
	set blog=nothing
	oblog.showok "修改评论成功！","user_comments.asp"
end sub


sub delcomment()
	dim blog,rstComment
	if id="" then
		oblog.adderrstr( "错误：请指定要删除的评论！")
		oblog.showusererr
		exit sub
	end if
	if instr(id,",")>0 then
		id=FilterIDs(id)
		dim n,i
		n=split(id,",")
		for i=0 to ubound(n)
			delonecomment(n(i))
		next
	else
		delonecomment(id)		
	end if	
	oblog.showok "删除评论成功!",""	
end sub

sub delonecomment(id)
	dim blog,rstComment,CommentNum
	id=clng(id)
	dim uid,mainid
	if oblog.logined_ulevel=9 then
		sql="select userid,mainid from [oblog_comment] where commentid=" & clng(id)
	else
		sql="select userid,mainid from [oblog_comment] where commentid=" & clng(id)&" and userid="&oblog.logined_uid
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		uid=rs(0)
		mainid=rs(1)
		rs.delete
		rs.close
		set blog=new class_blog
		blog.userid=uid
		'重新计算评论数目
		set rstComment=server.createobject("adodb.recordset")
		rstComment.Open "Select Count(commentid) From [oblog_comment] Where mainid=" & CLNG(mainid),conn,1,1
		If rstComment.Eof Then
			CommentNum=0
		Else
			If IsNull(rstComment(0)) Or Not IsNumeric(rstComment(0))Then
				CommentNum=0
			Else
				CommentNum=rstComment(0)
			End If
		End If
		Set rstComment = Nothing
		oblog.execute("update [oblog_log] set commentnum=" & CommentNum &" where logid="&mainid)
		blog.update_log mainid,0
		oblog.execute("update [oblog_user] set comment_count=comment_count-1 where userid="&uid)
		set blog=nothing
	else
		rs.close
		set rs=nothing
		oblog.adderrstr( "错误：无删除权限！")
		oblog.showusererr
		exit sub			
	end if
end sub


%>