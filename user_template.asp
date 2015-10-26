<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->
<%
'设置每页模板显示显示数目
Const P_USER_TEMPLATE_PERPAGE=12
'模板显示顺序,1:倒序,最新的在最前面;2:顺序,最老的在最前面
Const P_USER_TEMPLATE_ORDERBY=1
%>

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
					
					
					
						<center><h2>模板管理</h2></center>

			


<%
dim action
Action=trim(request("Action"))
select case Action
	case "saveconfig" 
		call saveconfig()
	case "modiconfig"
		call modiconfig()
	case "savedefault"
		call savedefault()
	case "modiviceconfig"
		call modiviceconfig()
	case "saveviceconfig"
		call saveviceconfig()
	case "bakskin"
		call bakskin()
	case "good"
		call good()
		
	case "Add_user_skin_main_Nav"
		call Add_user_skin_main_Nav()
	case "SaveAdd_user_skin_main_Nav"
		call SaveAdd_user_skin_main_Nav()
	case Else
		call showconfig()
End select
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
function VerIfySubmit()
{
	submits(); 
	If (document.oblogform.edit.value == "")
     {
        alert("请输入模板内容!");
	return false;
     }	
	return true;
}
function setbak()
{
	document.bakform.bak.value='bak';
}
function setrestore()
{
	document.bakform.bak.value='restore';
}

</SCRIPT>
<%
sub showconfig()
	Dim rs,totalSkins,SkinStrings,defaultskin
	Dim bookmark,sOrderby
	Set rs=Server.CreateObject("Adodb.Recordset")
	rs.Open "select defaultskin from [oblog_user] where userid="&oblog.logined_uid,conn,1,3		
	defaultskin=rs(0)
	rs.Close
	If P_USER_TEMPLATE_ORDERBY=1 Then sOrderby=" Desc"
	rs.Open "select * from oblog_userskin where ispass=1 Order By Id" & sOrderby,conn,1,3
	If rs.Eof Then
		%>
		<h1 style='font-size:13px;'><font color=red>当前没有任何模板可以使用</font></h1>
		<%
		Exit Sub
	End If
%>
<h1 style='font-size:13px;'>选择默认模板&nbsp;&nbsp;>>&nbsp;&nbsp;
<%
If defaultskin=0 Then 
	Response.Write "你当前还没有选择任何模板"
Else
	rs.Filter="id=" & defaultskin
	If rs.RecordCount>0 Then
		Response.Write "你当前选择的模板是<a href='showskin.asp?id=" & rs("id") & "' target=_blank><font color=red>" & rs("id") & " : " & rs("userskinname") &"</font></a>"
	End If
	rs.Filter=""
	rs.Movefirst
End If
%>	
	</h1>
<div id="list">
	<form name="form2" method="post" action="user_template.asp?action=savedefault">
		<center><input type="submit" name="Submit1" value="保存选择(将覆盖使用中的模板!)" /></center>	
<% 
		SkinStrings="<ul>" & VBCRLF
		totalSkins=rs.recordcount	
		If request("page")<>"" Then
    		currentPage=cint(request("page"))
		else
			currentPage=1
		End If	
		If currentPage<1 Then
       		currentPage=1
    	End If
    	If (currentPage-1)* P_USER_TEMPLATE_PERPAGE>totalSkins Then
	   		If (totalSkins mod P_USER_TEMPLATE_PERPAGE)=0 Then
	     		currentPage= totalSkins \ P_USER_TEMPLATE_PERPAGE
		  	Else
		      	currentPage= totalSkins \ P_USER_TEMPLATE_PERPAGE + 1
	   		End If
    	End If
    	
  	    If (currentPage-1)*P_USER_TEMPLATE_PERPAGE<totalSkins Then
           	rs.move  (currentPage-1)*P_USER_TEMPLATE_PERPAGE
         	
           	bookmark=rs.bookmark
        Else
	       	currentPage=1
	    End If		
		SkinStrings = SkinStrings & GetSkinList(rs,defaultskin)    
	rs.Close
	set rs=Nothing	
	SkinStrings = SkinStrings & "</ul>" & VBCRLF	
	Response.Write SkinStrings
	SkinStrings=""
	%>	
	<ul class="list_bot"> 
	<input type="hidden" value="<%=request("u")%>">
	<center><input type="submit" name="Submit1" value="保存选择(将覆盖使用中的模板!)" /></center>
	</ul>
	<%
	Response.Write "<ul class=""list_bot""> " & oblog.showpage("user_template.asp",totalSkins,P_USER_TEMPLATE_PERPAGE,false,true,"个模板") & "</ul>" & VBCRLF		
	%>
	</form>
</div>
<%
	set rs=nothing
End sub

sub savedefault()
	dim rs,rsskin,isdefaultID
	isdefaultID=clng(trim(request("radiobutton")))
	set rsskin=oblog.execute("select skinmain,skinshowlog,skinmain_Nav from oblog_userskin where id="&isdefaultID)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select user_skin_main,user_skin_showlog,defaultskin,user_skin_main_Nav from [oblog_user] where userid="&oblog.logined_uid,conn,1,3
	rs(0)=rsskin(0)
	rs(1)=rsskin(1)
	rs("user_skin_main_Nav")=rsskin("skinmain_Nav")
	rs(2)=isdefaultID
	set rsskin=nothing
	rs.update
	rs.close
	set rs=nothing
	updateindex()
	oblog.showok "修改成功,首页已经更新，其他页面请手动更新！",""
End sub

sub saveconfig()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_main from [oblog_user] where userid="&oblog.logined_uid
	rs.open sql,conn,1,3
	rs(0)=oblog.filtpath(oblog.filt_badword(trim(request("skinmain"))))
	rs.update
	rs.close
	set rs=nothing
	updateindex()
	oblog.showok "修改成功,首页已经更新，其他页面请手动更新！",""
End sub

sub saveviceconfig()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_showlog from [oblog_user] where userid="&oblog.logined_uid
	rs.open sql,conn,1,3
	rs(0)=oblog.filtpath(oblog.filt_badword(trim(request("skinshowlog"))))
	rs.update
	rs.close
	set rs=nothing
	updateindex()
	oblog.showok "修改成功,首页已经更新，其他页面请手动更新！",""
End sub

sub modiconfig()
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		oblog.adderrstr("请先选择一个默认模板！")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 

</style>
<h1 style='font-size:13px;'>现在修改的是我的主模板</h1>
<div id="list">
	<ul class="list_edit">
	主模板决定页面整体风格，<a href="#help1">点击这里可以查看模板标签说明</a>，建议修改前先<a href="user_template.asp?action=bakskin">备份模板</a>。<hr>
	</ul>
	<form method="POST" action="user_template.asp" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
	<input type="hidden" name="skinmain" id="edit" value="<%=Server.HtmlEncode(rs(0))%>" />
	<!--#include file="edit.asp"-->
    <ul class="list_edit"> 
	<input name="Action" type="hidden" id="Action" value="saveconfig" /> 
	<input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " /> 
	</ul>
	</form> 
<style>
	.list_top ul { 
		width: 630px;
		margin:0px;
		padding:0px;
		}
	.list_top ul li.1 { 
		width: 200px;
		text-align: left;
		float:left;
		} 
	.list_top ul li.2 { 
		width: 400px;
		text-align: left;
		float:left;
	} 
	.list_content ul { 
		width: 630px;
		margin:0px;
		padding:0px;
		}
</style> 
<!--	<h1 style='font-size:13px;'><a name="help1">主模板标记说明</a></h1> 
	
	<ul class="list_top" style="width:610px;line-height:18px">
		<li class="1" style="float:left;width: 150px;line-height:18px">标记名</li>
		<li class="2" style="float:left;width: 400px;line-height:18px">标记说明</li>
	</ul>
	<ul class="list_content" style="width: 610px;line-height:18px">
		<li class="1" style="float:left;width: 150px;line-height:18px">$show_log$</li>
		<li class="2" style="float:left;width: 400px;line-height:18px">重要，此标记显示文章主体部分，包括评论等信息。</li>
	</ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_placard$ </li>
		<li class="2" style="float:left;width: 400px;">此标记显示用户公告。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_calendar$</li>
		<li class="2" style="float:left;width: 400px;"> 此标记显示日历。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newblog$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示最新文章列表。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_comment$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示最新回复列表。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_subject$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示专题分类。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_subject_l$</li>
		<li class="2" style="float:left;width: 400px;">此标记横向显示专题分类。 </li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newblog$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示最新文章列表。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newmessage$ </li>
		<li class="2" style="float:left;width: 400px;">此标记显示最新留言列表。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_info$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示真实姓名称，统计信息等。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_login$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示登陆窗口。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_links$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示链接信息。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_blogname$ </li>
		<li class="2" style="float:left;width: 400px;">此标记显示用户真实姓名称，若名称为空则显示用户id。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_search$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示搜索表单。</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_xml$</li>
		<li class="2" style="float:left;width: 400px;">此标记显示rss连接标志。</li></ul>
	-->
</div>
<SCRIPT language=javascript>
		var IsIE5=document.all;
		if (IsIE5){
			var IframeID=frames["oblog_Composition"];
		}
		else{
			var IframeID=document.getElementById("oblog_Composition").contentWindow;
		}
		if(IframeID !=null){ 
		oblog_Size(500);
		}
</SCRIPT>
<%
set rs=nothing
End sub


sub Add_user_skin_main_Nav()'修改栏目级模板（字段user_skin_main_Nav）；
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main_Nav,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		oblog.adderrstr("请先选择一个默认模板！")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 

</style>
<h1 style='font-size:13px;'>现在修改的是 栏目级的主模板</h1>
<div id="list">
	<ul class="list_edit">
	主模板决定页面整体风格，<a href="#help1">点击这里可以查看模板标签说明</a>，建议修改前先<a href="user_template.asp?action=bakskin">备份模板</a>。<hr>
	</ul>
	<form method="POST" action="user_template.asp" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
		<textarea cols="60" rows="20" name="skinmain" id="edit"><%=Server.HtmlEncode(rs(0))%></textarea>
		
		<ul class="list_edit"> 
			<input name="Action" type="hidden" id="Action" value="SaveAdd_user_skin_main_Nav" /> 
			<input name="cmdSave" type="submit" id="cmdSave" value=" 保存修改 " /> 
		</ul>
	</form> 
<style>
	.list_top ul { 
		width: 630px;
		margin:0px;
		padding:0px;
		}
	.list_top ul li.1 { 
		width: 200px;
		text-align: left;
		float:left;
		} 
	.list_top ul li.2 { 
		width: 400px;
		text-align: left;
		float:left;
	} 
	.list_content ul { 
		width: 630px;
		margin:0px;
		padding:0px;
		}
</style> 
</div>
<SCRIPT language=javascript>
		var IsIE5=document.all;
		if (IsIE5){
			var IframeID=frames["oblog_Composition"];
		}
		else{
			var IframeID=document.getElementById("oblog_Composition").contentWindow;
		}
		if(IframeID !=null){ 
		oblog_Size(500);
		}
</SCRIPT>
<%
set rs=nothing
End sub

Sub SaveAdd_user_skin_main_Nav()
	dim rs,sql
	set rs=server.CreateObject("adodb.recordset")
	sql="select user_skin_main_Nav from [oblog_user] where userid="&oblog.logined_uid
	rs.open sql,conn,1,3
	rs(0)=oblog.filtpath(oblog.filt_badword(trim(request("skinmain"))))
	rs.update
	rs.close
	set rs=nothing
	updateindex()
	oblog.showok "栏目级模板修改成功,首页已经更新，其他页面请手动更新！",""
End sub


sub modiviceconfig()
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_showlog,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		set rs=nothing
		oblog.adderrstr("请先选择一个默认模板！")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 
</style>
<h1 style='font-size:13px;'>现在修改的是我的副模板</h1>
<div id="list">
	<ul class="list_edit">
	副模板决定文章部分显示风格，<a href="#help2">点击这里可以查看模板标签说明</a>，建议修改前先<a href="user_template.asp?action=bakskin">备份模板</a>。<hr>
	</ul>
	<form method="POST" action="user_template.asp?id=<%=clng(request.QueryString("id"))%>" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
	<input type="hidden" name="skinshowlog" id="edit" value="<%=Server.HtmlEncode(rs(0))%>" />
	<!--#include file="edit.asp"-->
	<ul class="list_edit">
	<input name="action" type="hidden" id="action" value="saveviceconfig" /> 
	<input type="submit"  value="保存修改" /> 
	</ul>
	</form>
	<!--<h1 style='font-size:13px;'><a name="help2">副模板标记说明</a></h1> 
	<ul class="list_top">
	<li class="1">标记名</li>
	<li class="2">标记说明</li>
	</ul>
	<ul class="list_content">
	<li class="1">$show_topic$</li>
	<li class="2">此标记显示表情图标，专题名，文章标题。</li></ul>
	<ul class="list_content">
    <li class="1">$show_loginfo$</li>
	<li class="2">此标记显示文章作者，发布时间等信息。</li></ul>
	<ul class="list_content">
    <li class="1">$show_logtext$</li>
	<li class="2">此标记显示文章正文。</li></ul>
	<ul class="list_content">
    <li class="1">$show_more$</li>
	<li class="2">此标记显示阅读全文(次数)，回复(次数)，引用链接。</li></ul>
	<ul class="list_content">
    <li class="1">$show_emot$</li>
	<li class="2">此标记仅显示显示表情图标。</li></ul>
	<ul class="list_content">
    <li class="1">$show_author$</li>
	<li class="2">此标记仅显示作者名。</li></ul>
	<ul class="list_content">
    <li class="1">$show_addtime$</li>
	<li class="2">此标记仅显示发布时间。</li></ul>
	<ul class="list_content">
    <li class="1">$show_topictxt$</li>
	<li class="2">此标记仅显示文章标题。</li></ul>
	<ul class="list_content">
	<li class="1">$show_blogzhai$</li>
	<li class="2">此标记显示加入到网摘的连接。</li></ul>
	<ul class="list_content">
	<li class="1">$show_blogtag$</li>
	<li class="2">此标记显示文章标签。</li></ul>-->
</div>
<SCRIPT language=javascript>
	var IsIE5=document.all;
	if (IsIE5){
		var IframeID=frames["oblog_Composition"];
	}
	else{
		var IframeID=document.getElementById("oblog_Composition").contentWindow;
	}
	if(IframeID !=null){ 

	oblog_Size(500);
	}
</SCRIPT>
<%
	set rs=nothing
End sub

sub bakskin()
	dim bak,rs
	bak=request("bak")
	If bak="bak" Then
		oblog.execute("update oblog_user set bak_skin1=user_skin_main,bak_skin2=user_skin_showlog where userid="&oblog.logined_uid)
		oblog.showok "备份模板成功！",""
	ElseIf bak="restore" Then
		set rs=oblog.execute("select bak_skin1,bak_skin2 from oblog_user where userid="&oblog.logined_uid)
		If rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) Then
			set rs=nothing
			oblog.adderrstr("所备份的模板为空，不允许恢复！")
			oblog.showusererr
		End If
		oblog.execute("update oblog_user set user_skin_main=bak_skin1,user_skin_showlog=bak_skin2 where userid="&oblog.logined_uid)
		set rs=nothing
		oblog.showok "恢复模板成功！",""
	End If
%>
	<h1 style='font-size:13px;'>备份及恢复模板</h1> 
<div id="list">
	<ul class="list_edit">
	建议修改或者更换模板前先备份模板，误操作以后可以随时恢复!<hr>
	<form name="bakform" method="post" action="user_template.asp?action=bakskin">
	<input type="hidden" name="bak" id="bak" value="" /> 
	<input type="submit" name="Submit" value="备份现在使用的模板" onClick="setbak();" /><br /><br />
	<input type="submit" name="Submit2" value="恢复模板(将覆盖使用中的模板)" onClick="setrestore();" /><br /><br />
	<input type="button" name="Submit1" value="查看已经备份的模板" onClick="window.open('showskin.asp?userid=<%=oblog.logined_uid%>')" />
	</form>
	</ul>
</div>
<%End sub


sub good()
	dim good,rs,rsu,skinname
	good=request("good")
	If good="save" Then
		skinname=request("skinname")
		If skinname="" or oblog.strLength(skinname)>50  Then oblog.adderrstr("用户名不能为空(不能大于50)！")
		If oblog.errstr<>"" Then oblog.showusererr:exit sub
		set rsu=oblog.execute("select user_skin_main,user_skin_showlog from oblog_user where userid="&oblog.logined_uid)
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select top 1 * from [oblog_userskin]",conn,1,3
		rs.addnew
		rs("userskinname")=skinname
		rs("skinmain")=rsu(0)
		rs("skinshowlog")=rsu(1)
		rs("skinauthor")=oblog.logined_uname
		rs("skinauthorurl")=oblog.logined_udir&"/"&oblog.logined_uid&"/index."&f_ext
		rs.update
		rs.close
		set rsu=nothing
		set rs=nothing		
		oblog.showok "推荐成功，请等待管理员审核！",""
	End If
%>
<h1 style='font-size:13px;'>推荐我的模板</h1> 
<div id="list">
    <ul class="list_edit">
	如果你的模板很漂亮，推荐给管理员，可以放在系统模板里，让更多人使用哦!<br />
	模板会注明作者和你的blog连接。请不要提交已经存在的用户模板！<hr>
	给我的模板取个名字：
	<form name="good" action="user_template.asp?action=good&good=save" method="post">
	<input name="skinname"   type="text" value="" size="30" maxlength="20" />
	<input type="submit" value="推荐" />
	</form>
    </ul>
</div>
<%End sub


sub updateindex()
	dim blog
	set blog=new class_blog
	blog.userid=oblog.logined_uid
	blog.update_index 0
	blog.update_message 0
	blog.CreateFunctionPage
	set blog=nothing
End sub

Function GetSkinList(byref rst,byval defaultskin)
	Dim strSkins,ustr,iCount
	strSkins=""
	Do While Not rst.eof 
		strSkins = strSkins & "<div class=""skinpic"">" & VBCRLF
		strSkins = strSkins & "<a href='showskin.asp?id=" & rst("id") & "' target=_blank>" & VBCRLF

		If rst("skinpic")="" or isnull(rst("skinpic")) Then 
			strSkins = strSkins & "<img src=""images/nopic.gIf"" alt=""无预览图片"" width=200 height=122 border=""0"" />" & VBCRLF
		Else
			strSkins = strSkins & "<img src="""&oblog.filt_html(rst("skinpic"))&""" alt=""预览图片"" width=200 height=122 border=""0"" />" & VBCRLF
		End If

		strSkins = strSkins & "</a><br />" & VBCRLF
		strSkins = strSkins & "<input name=""radiobutton"" type=""radio"" value=""" & rst("id") & """"

		If defaultskin=0 and rst("isdefault")=1 Then
			strSkins = strSkins & " checked"
		Else
			If defaultskin=rst("id") Then strSkins = strSkins & "checked"						
		End If
		If rst("skinauthorurl")<>"" Then 
			ustr="<a href="""&rst("skinauthorurl")&""">"&rst("skinauthor")&"</a>"
		Else
			ustr=rst("skinauthor")
		End If				
		
		strSkins = strSkins & " />"
		strSkins = strSkins & rst("userskinname")
		strSkins = strSkins & "<br />by:" & ustr & VBCRLF
		strSkins = strSkins & "</div>	" & VBCRLF & VBCRLF
		
		iCount = iCount+1
		if iCount >=P_USER_TEMPLATE_PERPAGE then exit do
		rst.movenext
	Loop
	GetSkinList = strSkins
End Function

%>