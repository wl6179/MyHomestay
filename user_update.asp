<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->


<div id="head">
	<div id="head2">
	
		<div id="head2-logo">
			
			<ul>
				<li class="active"><a href="http://www.myhomestay.com.cn">��������</a></li>
				<li><a href="http://www.myhomestay.com.cn">English</a></li>
				<!--<li><a href="http://www.myhomestay.com.cn">Janpan</a></li>-->
			</ul>
			
		</div>
		
		
		<div id="head2-menu">
			<div id="head2-search">
				<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
				<!--<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>-->&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��������</a>&nbsp;&nbsp;&nbsp;-->
			</div>
			
			<div id="divCN-EN">
			<ul>
				<li><a href="/index.asp">��ҳ<br />Home</a></li><!--#include file="menu/incIndexNav.htm"-->
	
				
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
					
					
					
						<center><h2>����<%=tName%></h2></center>






<div id="main2">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
	</div>
  </div>
  <div class="content">
  	<div class="content_top2">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			
	</div>
	
    <div class="content_body">
	<style>
		#progress {
		background:#blue;
		color:#FF6600;
		}
		#prompt ul {
		list-style-type:none;
		}
		#prompt ul li {
		display:block;
		background:#eeeeee;
		color:#FFFFFF;
		width:130px;
		}

	</style>
<%
dim action,blog
Dim totalPut,TotalPages,MaxPerPage,strFileName
action=request("action")
set blog=new class_blog
Server.ScriptTimeOut=999999999
select case action
	case "update_message"
		Call update_message
	case "update_blog"
		Call update_blog
	case "update_alllog"
		Call update_alllog 
	case else
	call main()
end select
set blog=nothing
%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>   
  
  
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
								��Ŀ��������
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
sub main
%>
<h1 style="font-size:16px;">�������¾�̬ҳ��</h1>
<div id="list">
	<ul class="list_edit">
	�����ݺ;�̬ҳ�治ͬ��ʱ�������ֶ�������ҳ�ļ���ͳ�����ݡ�<hr />
	<form>
	<input type="button" value="������վ��ҳ" onClick="window.location='user_update.asp?action=update_blog'" />
	��<input type="button" value="�������԰�"  onclick="window.location='user_update.asp?action=update_message'" />
	��<input type="button" value="����ͳ������" onClick="window.location='user_update.asp?action=update_blog'" /></form>
	<br /><br />
	���·���ȫվ�Ϻķ�ϵͳ��Դ�������ʹ��!<br />
	<%If CLNG(oblog.setup(26,0))>0 Then %>
	(ÿ��ֻ�ܽ���<%=oblog.setup(26,0)%>��!)<hr />
	<%End If%>
	<form><input type="button" value="���·���ȫվ" onClick="window.location='user_update.asp?action=update_alllog'" /></form>
	</ul>
</div>
<%end sub

sub update_message()
	const p=4
	response.Write("<div id=""prompt""><ul>")
	blog.progress_init
	blog.progress int(1/p*100),"��ȡ�û�����..."
	blog.userid=oblog.logined_uid
	blog.progress int(2/p*100),"�������԰�..."
	'blog.update_message 1,0,0,""
	blog.update_message 1
	blog.progress int(3/p*100),"������������..."
	blog.update_newmessage oblog.logined_uid
	blog.progress int(4/p*100),"�����������!"
	response.Write("<li><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul></div>")
end sub

sub update_blog()
	dim rsu,rst,n_log,n_comment,n_message	
	const p=10
	response.Write("<div id=""prompt""><ul>")
	blog.progress_init
	blog.progress int(1/p*100),"����ȫվͳ������..."
	set rsu=oblog.execute("select count(logid) from oblog_log where userid="&oblog.logined_uid)
	if not rsu.eof then n_log=rsu(0) else n_log=0
	set rsu=oblog.execute("select count(commentid) from oblog_comment where userid="&oblog.logined_uid)
	if not rsu.eof then n_comment=rsu(0) else n_comment=0
	set rsu=oblog.execute("select count(messageid) from oblog_message where userid="&oblog.logined_uid)
	if not rsu.eof then n_message=rsu(0) else n_message=0
	oblog.execute("update oblog_user set log_count="&n_log&",comment_count="&n_comment&",message_count="&n_message&" where userid="&oblog.logined_uid)
	
	blog.progress int(2/p*100),"���·���ͳ������..."
	set rst=server.createobject("adodb.recordset")
	rst.open "select subjectid,subjectlognum from oblog_subject where userid="&oblog.logined_uid,conn,2,2
	while not rst.eof
		set rsu=oblog.execute("select count(logid) from oblog_log where subjectid="&rst("subjectid"))
		if not rsu.eof then rst("subjectlognum")=rsu(0) else rst("subjectlognum")=0
		rst.update
		rst.movenext		
	wend
	rst.close
	set rst=nothing
	set rsu=nothing	
	
	blog.progress int(3/p*100),"��ȡ�û�����..."
	blog.userid=oblog.logined_uid
	blog.progress int(4/p*100),"������ҳ..."
	blog.update_index 1
	blog.progress int(5/p*100),"����վ����Ϣ�ļ�..."
	blog.update_info oblog.logined_uid
	blog.progress int(6/p*100),"�����������б��ļ�..."
	blog.update_newblog(oblog.logined_uid)
	blog.progress int(7/p*100),"������������..."
	blog.update_newmessage oblog.logined_uid
	blog.progress int(8/p*100),"������ҳ���·����ļ�..."
	blog.update_subject(oblog.logined_uid)
	blog.progress int(9/p*100),"���ɹ���ҳ��..."
	blog.CreateFunctionPage
	blog.progress int(10/p*100),"������վ��ҳ���!"
	response.Write("<li><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul></div>")
end sub

sub update_alllog()
	dim updateblognum,lastlogid
	lastlogid=trim(request("lastlogid"))
	if lastlogid<>"" then lastlogid=clng(lastlogid) else lastlogid=0
	updateblognum=request.Cookies(cookies_name&"uplognum")("updateblognum")
	if updateblognum="" or isnull(updateblognum) then updateblognum=0
	'updateblognum=0
	If clng(oblog.setup(26,0))=0 Then 
		updateblognum=0
	Else
		if clng(updateblognum)>=clng(oblog.setup(26,0)) and lastlogid=0 then
			oblog.adderrstr("��վ����ÿ��ֻ�ܽ���"&oblog.setup(26,0)&"��!")
			oblog.showusererr
			exit sub
		end if
	End if
	if lastlogid=0 then
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.Cookies(cookies_name&"uplognum")("updateblognum")=updateblognum+1
		Response.Cookies(cookies_name&"uplognum").Expires=Date+1
	end if
	response.Write("<div id=""prompt""><ul>")
	blog.progress_init
	blog.update_alllog oblog.logined_uid,lastlogid
	response.Write("<li>&lt;&lt; <input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></ul></div>")
end sub

%>