<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->
<%
'����ÿҳģ����ʾ��ʾ��Ŀ
Const P_USER_TEMPLATE_PERPAGE=12
'ģ����ʾ˳��,1:����,���µ�����ǰ��;2:˳��,���ϵ�����ǰ��
Const P_USER_TEMPLATE_ORDERBY=1
%>

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
					
					
					
						<center><h2>ģ�����</h2></center>

			


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


<SCRIPT language=javascript>
function VerIfySubmit()
{
	submits(); 
	If (document.oblogform.edit.value == "")
     {
        alert("������ģ������!");
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
		<h1 style='font-size:13px;'><font color=red>��ǰû���κ�ģ�����ʹ��</font></h1>
		<%
		Exit Sub
	End If
%>
<h1 style='font-size:13px;'>ѡ��Ĭ��ģ��&nbsp;&nbsp;>>&nbsp;&nbsp;
<%
If defaultskin=0 Then 
	Response.Write "�㵱ǰ��û��ѡ���κ�ģ��"
Else
	rs.Filter="id=" & defaultskin
	If rs.RecordCount>0 Then
		Response.Write "�㵱ǰѡ���ģ����<a href='showskin.asp?id=" & rs("id") & "' target=_blank><font color=red>" & rs("id") & " : " & rs("userskinname") &"</font></a>"
	End If
	rs.Filter=""
	rs.Movefirst
End If
%>	
	</h1>
<div id="list">
	<form name="form2" method="post" action="user_template.asp?action=savedefault">
		<center><input type="submit" name="Submit1" value="����ѡ��(������ʹ���е�ģ��!)" /></center>	
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
	<center><input type="submit" name="Submit1" value="����ѡ��(������ʹ���е�ģ��!)" /></center>
	</ul>
	<%
	Response.Write "<ul class=""list_bot""> " & oblog.showpage("user_template.asp",totalSkins,P_USER_TEMPLATE_PERPAGE,false,true,"��ģ��") & "</ul>" & VBCRLF		
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
	oblog.showok "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
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
	oblog.showok "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
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
	oblog.showok "�޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
End sub

sub modiconfig()
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		oblog.adderrstr("����ѡ��һ��Ĭ��ģ�壡")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 

</style>
<h1 style='font-size:13px;'>�����޸ĵ����ҵ���ģ��</h1>
<div id="list">
	<ul class="list_edit">
	��ģ�����ҳ��������<a href="#help1">���������Բ鿴ģ���ǩ˵��</a>�������޸�ǰ��<a href="user_template.asp?action=bakskin">����ģ��</a>��<hr>
	</ul>
	<form method="POST" action="user_template.asp" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
	<input type="hidden" name="skinmain" id="edit" value="<%=Server.HtmlEncode(rs(0))%>" />
	<!--#include file="edit.asp"-->
    <ul class="list_edit"> 
	<input name="Action" type="hidden" id="Action" value="saveconfig" /> 
	<input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " /> 
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
<!--	<h1 style='font-size:13px;'><a name="help1">��ģ����˵��</a></h1> 
	
	<ul class="list_top" style="width:610px;line-height:18px">
		<li class="1" style="float:left;width: 150px;line-height:18px">�����</li>
		<li class="2" style="float:left;width: 400px;line-height:18px">���˵��</li>
	</ul>
	<ul class="list_content" style="width: 610px;line-height:18px">
		<li class="1" style="float:left;width: 150px;line-height:18px">$show_log$</li>
		<li class="2" style="float:left;width: 400px;line-height:18px">��Ҫ���˱����ʾ�������岿�֣��������۵���Ϣ��</li>
	</ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_placard$ </li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ�û����档</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_calendar$</li>
		<li class="2" style="float:left;width: 400px;"> �˱����ʾ������</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newblog$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ���������б�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_comment$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ���»ظ��б�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_subject$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾר����ࡣ</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_subject_l$</li>
		<li class="2" style="float:left;width: 400px;">�˱�Ǻ�����ʾר����ࡣ </li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newblog$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ���������б�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_newmessage$ </li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ���������б�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_info$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ��ʵ�����ƣ�ͳ����Ϣ�ȡ�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_login$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ��½���ڡ�</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_links$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ������Ϣ��</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_blogname$ </li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ�û���ʵ�����ƣ�������Ϊ������ʾ�û�id��</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_search$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾ��������</li></ul>
	<ul class="list_content" style="width: 610px;">
		<li class="1" style="float:left;width: 150px;">$show_xml$</li>
		<li class="2" style="float:left;width: 400px;">�˱����ʾrss���ӱ�־��</li></ul>
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


sub Add_user_skin_main_Nav()'�޸���Ŀ��ģ�壨�ֶ�user_skin_main_Nav����
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_main_Nav,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		oblog.adderrstr("����ѡ��һ��Ĭ��ģ�壡")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 

</style>
<h1 style='font-size:13px;'>�����޸ĵ��� ��Ŀ������ģ��</h1>
<div id="list">
	<ul class="list_edit">
	��ģ�����ҳ��������<a href="#help1">���������Բ鿴ģ���ǩ˵��</a>�������޸�ǰ��<a href="user_template.asp?action=bakskin">����ģ��</a>��<hr>
	</ul>
	<form method="POST" action="user_template.asp" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
		<textarea cols="60" rows="20" name="skinmain" id="edit"><%=Server.HtmlEncode(rs(0))%></textarea>
		
		<ul class="list_edit"> 
			<input name="Action" type="hidden" id="Action" value="SaveAdd_user_skin_main_Nav" /> 
			<input name="cmdSave" type="submit" id="cmdSave" value=" �����޸� " /> 
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
	oblog.showok "��Ŀ��ģ���޸ĳɹ�,��ҳ�Ѿ����£�����ҳ�����ֶ����£�",""
End sub


sub modiviceconfig()
	dim rs,rsshowlog
	set rs=oblog.execute("select user_skin_showlog,defaultskin from [oblog_user] where userid="&oblog.logined_uid)
	If rs("defaultskin")=0  or rs("defaultskin")="" Then 
		set rs=nothing
		set rs=nothing
		oblog.adderrstr("����ѡ��һ��Ĭ��ģ�壡")
		oblog.showusererr
		exit sub		
	End If
%>
<style>
#list ul li.1 { width: 200px;text-align: left;} 
#list ul li.2 { width: 400px;text-align: left;} 
</style>
<h1 style='font-size:13px;'>�����޸ĵ����ҵĸ�ģ��</h1>
<div id="list">
	<ul class="list_edit">
	��ģ��������²�����ʾ���<a href="#help2">���������Բ鿴ģ���ǩ˵��</a>�������޸�ǰ��<a href="user_template.asp?action=bakskin">����ģ��</a>��<hr>
	</ul>
	<form method="POST" action="user_template.asp?id=<%=clng(request.QueryString("id"))%>" id="oblogform" name="oblogform" onSubmit="return VerIfySubmit()">
	<input type="hidden" name="skinshowlog" id="edit" value="<%=Server.HtmlEncode(rs(0))%>" />
	<!--#include file="edit.asp"-->
	<ul class="list_edit">
	<input name="action" type="hidden" id="action" value="saveviceconfig" /> 
	<input type="submit"  value="�����޸�" /> 
	</ul>
	</form>
	<!--<h1 style='font-size:13px;'><a name="help2">��ģ����˵��</a></h1> 
	<ul class="list_top">
	<li class="1">�����</li>
	<li class="2">���˵��</li>
	</ul>
	<ul class="list_content">
	<li class="1">$show_topic$</li>
	<li class="2">�˱����ʾ����ͼ�꣬ר���������±��⡣</li></ul>
	<ul class="list_content">
    <li class="1">$show_loginfo$</li>
	<li class="2">�˱����ʾ�������ߣ�����ʱ�����Ϣ��</li></ul>
	<ul class="list_content">
    <li class="1">$show_logtext$</li>
	<li class="2">�˱����ʾ�������ġ�</li></ul>
	<ul class="list_content">
    <li class="1">$show_more$</li>
	<li class="2">�˱����ʾ�Ķ�ȫ��(����)���ظ�(����)���������ӡ�</li></ul>
	<ul class="list_content">
    <li class="1">$show_emot$</li>
	<li class="2">�˱�ǽ���ʾ��ʾ����ͼ�ꡣ</li></ul>
	<ul class="list_content">
    <li class="1">$show_author$</li>
	<li class="2">�˱�ǽ���ʾ��������</li></ul>
	<ul class="list_content">
    <li class="1">$show_addtime$</li>
	<li class="2">�˱�ǽ���ʾ����ʱ�䡣</li></ul>
	<ul class="list_content">
    <li class="1">$show_topictxt$</li>
	<li class="2">�˱�ǽ���ʾ���±��⡣</li></ul>
	<ul class="list_content">
	<li class="1">$show_blogzhai$</li>
	<li class="2">�˱����ʾ���뵽��ժ�����ӡ�</li></ul>
	<ul class="list_content">
	<li class="1">$show_blogtag$</li>
	<li class="2">�˱����ʾ���±�ǩ��</li></ul>-->
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
		oblog.showok "����ģ��ɹ���",""
	ElseIf bak="restore" Then
		set rs=oblog.execute("select bak_skin1,bak_skin2 from oblog_user where userid="&oblog.logined_uid)
		If rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) Then
			set rs=nothing
			oblog.adderrstr("�����ݵ�ģ��Ϊ�գ�������ָ���")
			oblog.showusererr
		End If
		oblog.execute("update oblog_user set user_skin_main=bak_skin1,user_skin_showlog=bak_skin2 where userid="&oblog.logined_uid)
		set rs=nothing
		oblog.showok "�ָ�ģ��ɹ���",""
	End If
%>
	<h1 style='font-size:13px;'>���ݼ��ָ�ģ��</h1> 
<div id="list">
	<ul class="list_edit">
	�����޸Ļ��߸���ģ��ǰ�ȱ���ģ�壬������Ժ������ʱ�ָ�!<hr>
	<form name="bakform" method="post" action="user_template.asp?action=bakskin">
	<input type="hidden" name="bak" id="bak" value="" /> 
	<input type="submit" name="Submit" value="��������ʹ�õ�ģ��" onClick="setbak();" /><br /><br />
	<input type="submit" name="Submit2" value="�ָ�ģ��(������ʹ���е�ģ��)" onClick="setrestore();" /><br /><br />
	<input type="button" name="Submit1" value="�鿴�Ѿ����ݵ�ģ��" onClick="window.open('showskin.asp?userid=<%=oblog.logined_uid%>')" />
	</form>
	</ul>
</div>
<%End sub


sub good()
	dim good,rs,rsu,skinname
	good=request("good")
	If good="save" Then
		skinname=request("skinname")
		If skinname="" or oblog.strLength(skinname)>50  Then oblog.adderrstr("�û�������Ϊ��(���ܴ���50)��")
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
		oblog.showok "�Ƽ��ɹ�����ȴ�����Ա��ˣ�",""
	End If
%>
<h1 style='font-size:13px;'>�Ƽ��ҵ�ģ��</h1> 
<div id="list">
    <ul class="list_edit">
	������ģ���Ư�����Ƽ�������Ա�����Է���ϵͳģ����ø�����ʹ��Ŷ!<br />
	ģ���ע�����ߺ����blog���ӡ��벻Ҫ�ύ�Ѿ����ڵ��û�ģ�壡<hr>
	���ҵ�ģ��ȡ�����֣�
	<form name="good" action="user_template.asp?action=good&good=save" method="post">
	<input name="skinname"   type="text" value="" size="30" maxlength="20" />
	<input type="submit" value="�Ƽ�" />
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
			strSkins = strSkins & "<img src=""images/nopic.gIf"" alt=""��Ԥ��ͼƬ"" width=200 height=122 border=""0"" />" & VBCRLF
		Else
			strSkins = strSkins & "<img src="""&oblog.filt_html(rst("skinpic"))&""" alt=""Ԥ��ͼƬ"" width=200 height=122 border=""0"" />" & VBCRLF
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