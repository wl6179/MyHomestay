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
					
					
					
						<center><h2><%=tName%>������</h2></center>


			


<%
const MaxPerPage=20
dim strFileName
dim totalPut,TotalPages
dim rs,sql,blog
dim id,action
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
select case action
	case "addclass"
	call addclass()
	case "del"
	call delclass()
	case "modify"
	call modifyclass()
	case "savemodi"
	call savemodify()
	case "order"
	call order()
	case else
	call main
end select
set rs=nothing
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


sub addclass()
	call uporder()
	dim subjectname,rs,ordernum
	subjectname=Trim(Request.Form("subjectname"))
	If subjectname="" or oblog.strLength(subjectname)>50 then oblog.adderrstr("����������Ϊ���Ҳ��ܴ���50�ַ�)��")
	if oblog.chk_badword(subjectname)>0 then oblog.adderrstr("�������к���ϵͳ��������ַ���")
	oblog.showusererr
	set rs=oblog.execute("select max(ordernum) from oblog_subject where userid="&oblog.logined_uid & " And SubjectType=" & t)
	if not isnull(rs(0)) then
		ordernum=rs(0)+1
	else
		ordernum=1
	end if
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from [oblog_subject] Where SubjectType=" & t,conn,1,3
	rs.addnew
	rs("subjectname")=subjectname
	rs("userid")=oblog.logined_uid
	rs("ordernum")=ordernum
	rs("subjectType")=t
	rs.update
	rs.close
	set rs=nothing
	oblog.showok "��ӷ���ɹ�!","user_subject.asp?t=" & t
end sub

sub delclass
	dim id
	id=clng(request.QueryString("id"))
	oblog.execute("delete  from [oblog_subject] where subjectid="&id&" and userid="&oblog.logined_uid)
	oblog.execute("update [oblog_log] set subjectid=0 where subjectid="&id&" and userid="&oblog.logined_uid)
	call uporder()
    oblog.showok "ɾ������ɹ�!",""
end sub

sub savemodify()
	dim subjectname,rs
	id=clng(id)
	subjectname=Trim(Request.Form("subjectname"))
	If subjectname="" or oblog.strLength(subjectname)>50 then oblog.adderrstr("����������Ϊ���Ҳ��ܴ���50�ַ�)��")
	if oblog.chk_badword(subjectname)>0 then oblog.adderrstr("�������к���ϵͳ��������ַ���")
	if oblog.errstr<>"" then oblog.showusererr:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select subjectname from [oblog_subject] where subjectid="&id&" and userid="&oblog.logined_uid,conn,1,3
	if not rs.eof then
		rs("subjectname")=subjectname
		rs.update
	end if
	rs.close
	set rs=nothing
	oblog.showok "�޸ķ������Ƴɹ�!","user_subject.asp?t=" & t
end sub

sub order()
	dim ordernum,modi,rs
	ordernum=clng(request.QueryString("ordernum"))
	modi=request.QueryString("modi")
	select case modi
		case "up"
			if ordernum-1>0 then
				oblog.execute("update [oblog_subject] set ordernum=9999999 where ordernum="&ordernum-1&" and userid="&oblog.logined_uid & " And SubjectType=" & t)
				oblog.execute("update [oblog_subject] set ordernum=ordernum-1 where ordernum="&ordernum&" and userid="&oblog.logined_uid & " And SubjectType=" & t)
				oblog.execute("update [oblog_subject] set ordernum="&ordernum&" where ordernum=9999999"&" and userid="&oblog.logined_uid & " And SubjectType=" & t)
			end if			
		case "down"
			set rs=oblog.execute("select max(ordernum) from oblog_subject where userid="&oblog.logined_uid)
			if ordernum<rs(0) then
				oblog.execute("update [oblog_subject] set ordernum=9999999 where ordernum="&ordernum+1&" and userid="&oblog.logined_uid & " And SubjectType=" & t)
				oblog.execute("update [oblog_subject] set ordernum=ordernum+1 where ordernum="&ordernum&" and userid="&oblog.logined_uid & " And SubjectType=" & t)			
				oblog.execute("update [oblog_subject] set ordernum="&ordernum&" where ordernum=9999999 and userid=" & oblog.logined_uid & " And SubjectType=" & t)
			end if
			set rs=nothing
	end select
	'uporder()
	response.Redirect "user_subject.asp?t=" & t
end sub

sub uporder()
	dim rs,i,n
	n=0
	set rs=oblog.execute("select count(subjectid) from [oblog_subject] where userid="&oblog.logined_uid & " And SubjectType=" & t)
	redim ordernum(rs(0))
	set rs=oblog.execute("select subjectid from [oblog_subject] where userid=" & oblog.logined_uid & " And SubjectType=" & t & " order by ordernum")
	while not rs.eof
		ordernum(n)=rs(0)
		n=n+1
		rs.movenext		
	wend
	i=1
	for n=0 to ubound(ordernum)
		oblog.execute("update oblog_subject set ordernum="&i&" where subjectid="&clng(ordernum(n)))
		i=i+1
		'response.Write(i)		
	next
	set rs=nothing	
end sub

sub main()
%>
<h1 style='font-size:13px;'>���<%=tName%>����</h1>
<div id="list">
	<ul class="list_edit" style="margin-left:0px">
	���<%=tName%>�����ֻ���ڴ˷��෢��<%=tName%>�Ż�����ҳ��ʾ����!<hr />
   <form name="form1" method="post" action="user_subject.asp?action=addclass&t=<%=t%>">
		<%=tName%>�������ƣ�
		<input name="subjectname" type="text" id="subjectname" maxlength="50" />
		<input type="submit" value="���" />			             
		</form>
  </ul>
</div>
<style type="text/css">
<!--
	#list ul{ width: 85%;display:block;clear:both;margin:0px;} 
	li.t1 { width: 50px; float:left;} 
	li.t2 { width: 220px; float:left;} 
	li.t3 { width: 300px; float:left;}
-->
</style> 

<h1 style='font-size:13px;'><%=tName%>����������</h1>
<div id="list">
	<ul class="list_edit" style="margin-left:0px">
  ���ķ�������,����ɾ���������Ҫ�ֶ�������ҳ�Ż���Ч!<hr />
  </ul>
	<ul class="list_top tr01" style="line-height:15px;padding:6px;border:1px #999999 solid;width:98%;">
	<li class="t1" style="float:left;width:10%;">����</li>
	<li class="t2" style="float:left;width:30%;"><%=tName%>������</li>
	<li class="t3" style="float:left;width:55%;">����</li>
	</ul>
<%
dim rs
set rs=oblog.execute("select * from oblog_subject where userid=" & oblog.logined_uid & " And SubjectType=" & t & " order by ordernum")
while not rs.eof 
%>
  <ul class="list_content tr03" style="float:none;width:98%;">
  <li class="t1" style="float:left;width:10%;"><%=rs("ordernum")%></li>
    <li class="t2" style="float:left;width:30%;"><%="<a href='"&blogdir&oblog.logined_udir&"/"&oblog.logined_ufolder&"/MyHomestay."&f_ext&"?uid="&oblog.logined_uid&"&do=blogs&id="&rs("subjectid")&"' target='_blank'>"&oblog.filt_html(rs("subjectname"))&"</a>"%></li>
    <li class="t3" style="float:left;width:55%;"><a href="user_subject.asp?action=modify&id=<%=rs("subjectid")%>&oldname=<%=rs("subjectname")%>&t=<%=t%>">�޸�</a> <a href="user_subject.asp?action=del&id=<%=rs("subjectid")%>&t=<%=t%>" <%="onClick='return confirm(""ȷ��Ҫɾ���˷�����(���ɻָ�)��"");'"%>>ɾ��</a>
	 <a href="user_subject.asp?action=order&modi=up&ordernum=<%=rs("ordernum")%>&t=<%=t%>">�����ƶ�</a> <a href="user_subject.asp?action=order&modi=down&ordernum=<%=rs("ordernum")%>&t=<%=t%>">�����ƶ�</a></li></ul>
<%
rs.movenext
wend 
set rs=nothing
%>
</div>
<%
end sub

sub modifyclass()
	dim oldname,rs
	id=clng(id)
	set rs=oblog.execute("select subjectname from oblog_subject where subjectid="&id&" and userid="&oblog.logined_uid)
	if not rs.eof then
	oldname=oblog.filt_html(rs(0))
%>

<h1 style='font-size:13px;'>�޸�<%=tName%>����</h1>
<div id="list">
	<ul class="list_edit">
	����<%=tName%>����������Ҫ������ҳ�Ż�ʹ�޸���Ч!<hr />	
	<form name="form1" method="post" action="user_subject.asp?action=savemodi&id=<%=id%>&t=<%=t%>">
		<%=tName%>�������ƣ�
		<input name="subjectname" type="text" id="subjectname" maxlength="20" value="<%=oldname%>" />
        <input type="submit" value="�޸�" />
    	</form>
	</ul>
</div>
<%
	set rs=nothing
	end if
end sub
%>