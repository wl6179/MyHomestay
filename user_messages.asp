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
					
					
					
						<center><h2>���Թ���</h2></center>


			


<%const MaxPerPage=20
dim strFileName
dim totalPut,TotalPages
dim rs,sql,blog
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
strFileName="user_messages.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
if keyword<>"" then
	strFileName="user_messages.asp?usersearch=10&keyword="&Keyword&"&Field="&strField
end if
select case action
	case "modify"
	call modify()
	case "savemodify" 
	call savemodify()
	case "del" 
	call delmessage()
	case else
	call top()	
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
sub top()
%>
<style type="text/css">
<!--
	#list ul{ width: 95%} 
	#list ul li.t1 { width: 50px} 
	#list ul li.t2 { width: 200px} 
	#list ul li.t3 { width: 120px} 
	#list ul li.t4 { width: 160px} 
	#list ul li.t5 { width: 180px} 
-->
</style>
<style type="text/css">
<!--
	#list_top ul{ width: 95%;float:left;} 
	#list_top ul li.t1 { width: 50px;float:left;} 
	#list_top ul li.t2 { width: 200px;float:left;} 
	#list_top ul li.t3 { width: 120px;float:left;} 
	#list_top ul li.t4 { width: 160px;float:left;} 
	#list_top ul li.t5 { width: 180px;float:left;} 
-->
</style>

<div id="list" style="width:99%;text-align:left;">
	<h2 style="width:100%;font-size:13px; text-align:left;padding:8px;">
  <form name="form1" action="user_messages.asp" method="get">
      ���ٲ������ԣ�
	    <select size=1 name="usersearch" onChange="javascript:submit()">
          <option value="0">�г���������</option>
       	  <%if oblog.logined_ulevel=9 then%><option value="1">�ҵ�����</option><%end if%>
          <option value="10" selected>��ѡ���ѯ����</option>         
        </select>
		 ��<br />����������
		 <select name="Field" id="Field">
        <option value="id">����</option>
		<option value="ip">����ip</option>
        <option value="topic" selected>���Ա���</option>
      </select>
      <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
      <input type="submit"  value=" ���� ">
  </form>
	</h2>
<%
end sub

sub main()
	dim strGuide,ssql
	ssql="userid,messagetopic,addtime,messageid,message_user,issave,addip"
	strGuide="<h1 style='font-size:13px;'>��ǰѡ��&nbsp;&gt;&gt;&nbsp;"
	select case usersearch
		case 0
			if oblog.logined_ulevel=9 then
				sql="select top 1000 "&ssql&" from [oblog_message] order by messageid desc"
			else
				sql="select "&ssql&" from [oblog_message] where userid="&oblog.logined_uid&" order by messageid desc"
			end if
			strGuide=strGuide & "��������"
		case 1
			sql="select "&ssql&" from [oblog_message] where userid="&oblog.logined_uid&" order by messageid desc"
			strGuide=strGuide & "�ҵ�����"
		case 10
			if Keyword="" then
				oblog.adderrstr( "���󣺹ؼ��ֲ���Ϊ�գ�")
				oblog.showusererr
				exit sub
			else
				select case strField
				case "id"
					if oblog.logined_ulevel=9 then 
						sql="select "&ssql&" from [oblog_message] where message_user like '%" & Keyword&"%' order by messageid desc"
					else
						sql="select "&ssql&" from [oblog_message] where message_user like '%" & Keyword&"%' and userid="&oblog.logined_uid&" order by messageid desc"
					end if
					strGuide=strGuide & "���������л��к���<font color=red> " & Keyword & " </font>������"
				case "topic"
					if oblog.logined_ulevel=9 then
						sql="select "&ssql&" from [oblog_message] where messagetopic like '%" & Keyword & "%' order by messageid desc"
					else
						sql="select "&ssql&" from [oblog_message] where messagetopic like '%" & Keyword & "%' and userid="&oblog.logined_uid&" order by messageid desc"
					end if
					strGuide=strGuide & "�����к��С� <font color=red>" & Keyword & "</font> ��������"
				case "ip"
					if oblog.logined_ulevel=9 then 
						sql="select "&ssql&" from [oblog_message] where addip='" & Keyword&"' order by messageid desc"
					else
						sql="select "&ssql&" from [oblog_message] where addip='" & Keyword&"' and userid="&oblog.logined_uid&" order by messageid desc"
					end if
					strGuide=strGuide & "����ipΪ<font color=red> " & Keyword & " </font>������"
				end select
			end if
		case else
	end select
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	'response.End()
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (����0������)</h1></div>"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"ƪ����")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"ƪ����")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"ƪ����")
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
	<form name="myform" method="Post" action="user_messages.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
	<ul class="list_top tr01" style="line-height:15px;padding:6px;border:1px #999999 solid;">
	<li class="t1" style="width: 50px;float:left;">ѡ��</li>
	<li class="t2" style="width: 160px;float:left;">���Ա���</li>
	<li class="t3" style="width: 120px;float:left;">������</li>
	<li class="t4" style="width: 130px;float:left;">����ʱ��</li>
	<li class="t5" style="width: 110px;float:left;">����</li>
	</ul>
	<%do while not rs.EOF %>
	<ul class="list_content tr03" style="width: 95%;"> 
	<li class="t1" style="width: 50px;float:left;"><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("messageid"))%>'></li>
	<li class="t2" style="width: 160px;float:left;"><a href=go.asp?messageid=<%=rs("messageid")%> target=_blank><%=oblog.filt_html(rs("messagetopic"))%>  </a> </li>
	<li class="t3" style="width: 120px;float:left;" title="<%="����:"&rs("addip")%>"><%=oblog.filt_html(rs("message_user"))%></li>
	<li class="t4" style="width: 130px;float:left;"><%if rs("addtime")<>"" then	response.write rs("addtime") else response.write "&nbsp;" %> </li>
	<li class="t5" style="width: 110px;float:left;">
<%
		response.write "<a href='user_messages.asp?action=modify&id=" & rs("messageid") & "&re=true'>�ظ�</a>&nbsp;"
		response.write "<a href='user_messages.asp?action=modify&id=" & rs("messageid") & "'>�޸�</a>&nbsp;"
		response.write "<a href='user_messages.asp?action=del&id=" & rs("messageid") &"' onClick='return confirm(""ȷ��Ҫɾ����������"");'>ɾ��</a>"
%> </li>
	</ul>
          <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
	<ul class="list_bot"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ȫѡ</li>
	<input name="action"  type="hidden" value="del" >
	<input type="submit" name="Submit" value=" ɾ�� ">
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
		oblog.adderrstr( "���󣺲������㣡")
		oblog.showusererr
		exit sub
	else
		id=Clng(id)
	end if
	Set rsblog=Server.CreateObject("Adodb.RecordSet")
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_message] where messageid=" & id
	else
		sql="select * from [oblog_message] where messageid=" & id&" and userid="&oblog.logined_uid
	end if
	rsblog.Open sql,Conn,1,1
	if rsblog.bof or rsblog.eof then
		rsblog.close
		set rsblog=nothing
		oblog.adderrstr( "�����Ҳ���ָ�������ԣ�")
		oblog.showusererr
		exit sub
	end if
	if request("re")="true" then
		restr="<table align=""center"" bgcolor=""#f3f3f3"" border=""1"" bordercolor=""#cccccc"" cellPadding=""2"" cellSpacing=""0"" style=""BORDER-COLLAPSE: collapse"" width=""90%"">"
		restr=restr&"<tbody><tr><td><p><strong>����Ϊblog���˵Ļظ���</strong></p><p>&nbsp;</p></td></tr></tbody></table>"
	end if
%>
<SCRIPT language=javascript>
function VerifySubmit()
{
	topic = del_space(document.oblogform.topic.value);
     if (topic.length == 0)
     {
        alert("��������д��Ŀ!");
	return false;
     }
	 
	submits(); 
	if (document.oblogform.edit.value == "")
     {
        alert("����������!");
	return false;
     }	
	return true;
}
</SCRIPT>

<div id="list">
<form action="user_messages.asp?action=savemodify" method="post" name="oblogform" onSubmit="return VerifySubmit()">
	<h1 style='font-size:13px;'>�޸�����</h1>
	<ul class="list_edit"> 
   ���Ա��⣺
	<input name="topic" type=text class="cont" id="topic" value="<%=rsblog("messagetopic")%>" size="50" maxlength="30" />
	</ul>
    <input type="hidden" name="edit" id="edit" value="<%if rsblog("message")<>"" then response.Write Server.HtmlEncode(rsblog("message")+restr)%>" /> 
    <!--#include file="edit.asp"-->
	<ul class="list_edit">
	<input type="hidden" name="id" value="<%=rsblog("messageid")%>" />
        <input type="submit" name="Submit2" value="�ύ�޸�" />
	</ul>
</form>
</div>

<%
	rsblog.close
	set rsblog=nothing
end sub

sub savemodify()
	dim id,blog,issave,messagefile,userid
	id=clng(trim(request("id")))
	if oblog.logined_ulevel=9 then
		sql="select * from oblog_message where messageid="&id
	else
		sql="select * from oblog_message where messageid="&id&" and userid="&oblog.logined_uid
	end if	
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		userid=rs("userid")
		messagefile=rs("messagefile")
		rs("messagetopic")=oblog.InterceptStr(oblog.filt_badword(trim(request("topic"))),250)
		rs("message")=oblog.filt_badword(request("edit"))
		rs.update
		rs.close	
		set blog= new class_blog
		blog.userid=userid	
		blog.update_message 0
		set rs=nothing
		set blog=nothing
	end if
	oblog.showok "�޸����Գɹ���","user_messages.asp"
end sub

sub delmessage()
	if id="" then
		oblog.adderrstr( "������ָ��Ҫɾ�������ԣ�")
		oblog.showusererr
		exit sub
	end if
	if instr(id,",")>0 then
		dim n,i
		id=FilterIDs(id)
		n=split(id,",")
		for i=0 to ubound(n)
			delonemessage(n(i))
		next
	else
		delonemessage(id)
	end if
	set rs=nothing
	oblog.showok "ɾ�����Գɹ�!","" 
end sub

sub delonemessage(id)
	id=clng(id)
	dim userid,issave,messagefile
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_message] where messageid=" & clng(id)
	else
		sql="select * from [oblog_message] where messageid=" & clng(id)&" and userid="&oblog.logined_uid
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		userid=rs("userid")
		messagefile=rs("messagefile")
		rs.delete
		rs.close
		set blog=new class_blog
		blog.userid=userid	
		'blog.update_message 0,0,0,""
		blog.update_message 0
		blog.update_newmessage userid
		oblog.execute("update [oblog_user] set message_count=message_count-1 where userid="&userid)
		set blog=nothing
	else
		rs.close
		set rs=nothing
		oblog.adderrstr( "������ɾ��Ȩ�ޣ�")
		oblog.showusererr
		exit sub			
	end if
end sub
%>