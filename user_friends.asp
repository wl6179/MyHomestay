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
					
					
					
						<center><h2>���Ѽ�����������</h2></center>
			

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
strFileName="user_friends.asp?usersearch=" & usersearch
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
sub main()
%>
<style type="text/css">
<!--
	#list ul{ width: 90%} 
	#list ul li.t1 { width: 60px} 
	#list ul li.t2 { width: 120px} 
	#list ul li.t3 { width: 120px} 
	#list ul li.t4 { width: 200px} 
	#list ul li.t5 { width: 160px} 
-->
</style>

	<h1 style='font-size:13px;'>

<%
	dim strGuide,ssql
	ssql="id,username,nickname,user_dir,oblog_user.userid,oblog_friend.addtime"
	strGuide="��ǰѡ��&nbsp;&gt;&gt;&nbsp;"
	select case usersearch
		case 0
			sql="select "&ssql&" from oblog_friend,oblog_user where isblack=0 and oblog_friend.userid="&oblog.logined_uid&" and oblog_friend.friendid=oblog_user.userid order by id desc"
			strGuide=strGuide & "�����б�"
		case 1
			sql="select "&ssql&" from oblog_friend,oblog_user where isblack=1 and oblog_friend.userid="&oblog.logined_uid&" and oblog_friend.friendid=oblog_user.userid order by id desc"
			strGuide=strGuide & "������"
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (����0������)</h1>"
		response.write strGuide
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & " (����" & totalPut & "������)</h1>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����¼")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����¼")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"����¼")
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
<div id="list">
  <form name="myform" method="Post" action="user_friends.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<ul class="list_top tr01">
	<li class="t1" style="width:73px;float:left;">ѡ��</li>
	<li class="t2" style="width:113px;float:left;">����ID</li>
	<li class="t3" style="width:83px;float:left;">�����ǳ�</li>
	<li class="t4" style="width:173px;float:left;">���ʱ��</li>
	<li class="t5" style="width:120px;float:left;">����</li>
	</ul>
	<%do while not rs.eof %>
	<ul class="list_content tr03" > 
	<li class="t1"  style="width:73px;float:left;"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("id"))%>'></li>
	<li class="t2" style="width:113px;float:left;"><a href="<%=rs("user_dir")&"/"&rs("userid")&"/index."&f_ext%>" target=_blank><%=oblog.filt_html(rs("username"))%>  </a> </li>
	<li class="t3" style="width:83px;float:left;"><%oblog.filt_html(rs("nickname"))%></li>
	<li class="t4" style="width:173px;float:left;"> <%if rs("addtime")<>"" then	response.write rs("addtime") else response.write "&nbsp;" %> </li>
	<li class="t5" style="width:120px;float:left;">
<%
		response.write "<a href='user_friends.asp?action=del&id=" & rs("id") &"' onClick='return confirm(""ȷ��Ҫɾ����"");'>ɾ��</a>"
		response.write " <a href=""javascript:openScript('user_pm.asp?action=send&incept="&oblog.filt_html(rs("username"))&"',450,400)"">���Ͷ���</a>"
%> </li>
	</ul>
<%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
	<ul class="list_bot">
	<input name="chkAll" type="checkbox" id="chkAll" onClick="CheckAll(this.form)" value="checkbox" />
	 ȫѡ
	<input name="action"  type="hidden" value="del" />
	<input type="submit" name="Submit" value="����ɾ��" />
	</ul>
	</form>
</div>
<%
end sub

sub add()
dim str1,str2
if request("type")="black" then
	str1="��Ӻ�����"
	str2="�������û�����"
else
	str1="��Ӻ���"
	str2="�����û�����"
end if	
%>
<h1 style='font-size:13px;'><%=str1%></h1>
<div id="list">
	<ul class="list_edit"> 
	<%if request("type")="black" then%>
	����������Ժ󣬾Ͳ����ܵ�����û��Ķ���ɧ���ˡ�
	<%else%>
	���û���Ϊ���ѣ����Է���ķ���վ�ڶ��ţ������Ժͺ��ѹ���˽������!
	<%end if%>
	<hr>
	<form action="user_friends.asp?action=saveadd&type=<%=request("type")%>" method="post" name="oblogform">
	<%=str2%>
	<input name="friendname" type=text size="20" maxlength="30" />
	<input type="submit" value="���" />
	</form>
	</ul>
</div>
<%
end sub

sub saveadd()
	dim friendname,uid,isblack
	friendname=oblog.filt_badstr(trim(request("friendname")))
	if friendname="" then
		oblog.adderrstr("��������û�������Ϊ��")
		oblog.showusererr
		exit sub
	end if
	if request("type")="black" then isblack=1 else isblack=0
	sql="select userid from oblog_user where username='"&friendname&"'"
	set rs=oblog.execute(sql)
	if rs.eof then
		oblog.adderrstr("�����޴��û�")
		oblog.showusererr
		exit sub
	end if
	uid=rs(0)
	set rs=oblog.execute("select id from oblog_friend where userid="&oblog.logined_uid&" and friendid="&uid&" and isblack="&isblack)
	if rs.eof then
		oblog.execute("insert into [oblog_friend] (userid,friendid,isblack) values ("&oblog.logined_uid&","&uid&","&isblack&")")
		set rs=nothing
		oblog.showok "��ӳɹ�","user_friends.asp"
	else
		set rs=nothing
		oblog.adderrstr("���󣺴��û��Ѿ����б���")
		oblog.showusererr
	end if
end sub


sub del()
	if id="" then
		oblog.adderrstr( "������ָ��Ҫɾ���Ķ���")
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
	oblog.showok "ɾ���ɹ���",""
end sub

sub delone(id)
	id=clng(id)
	sql="delete from [oblog_friend] where id=" & clng(id)&" and userid="&oblog.logined_uid
	oblog.execute(sql)
end sub
%>