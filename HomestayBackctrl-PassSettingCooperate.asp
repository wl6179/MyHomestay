<!--#include file="user_topCooperate.asp"-->
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
				<span id="joincompany_login"><a href="/login.asp">��ҵ��¼</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">��ҵ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13146398085</span> &nbsp;&nbsp;&nbsp;
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
					
					
					
						<center><h2>��������</h2></center>
			

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
								���˹�������
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
	oblog.showok "�޸Ĺ���ɹ�",""
end sub

sub placard
dim rs
set rs=oblog.execute("select user_placard from [oblog_user] where userid="&oblog.logined_uid)
%>
<h1 style='font-size:13px;'>�޸�վ�㹫��</h1>
<div id="list">
	<ul class="list_edit">
	�������������������Ƭ��������Ľ��ܣ�������Ը�����ȥ���κ���Ϣ��<hr>
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
	<input type="submit" name="Submit" value="�ύ�޸�" />
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
	oblog.showok "�޸��������ӳɹ�",""
end sub

sub links
dim rs
set rs=oblog.execute("select user_links from [oblog_user] where userid="&oblog.logined_uid)
%>
<h1 style='font-size:13px;'>�޸���������</h1>
<div id="list">
	<ul class="list_edit">
	��������������ֻ���ͼƬ��Ȼ���� <image src="images/wlink.gif" /> ��ť���볬�����ӡ�<hr>
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
	<input type="submit" name="Submit" value="�ύ�޸�" />
	</ul>
	</form>
</div>
<%
set rs=nothing
end sub



sub blogpassword()
%>
<h1 style='font-size:13px;'>����blogվ��������趨</h1>
<div id="list">	
	<ul class="list_edit">
	ȫվ���ܺ����������¶���Ҫͨ��������֤����ܷ��ʡ�</br>
	ע�⣺�����������Ժ���Ҫ<a href="user_update.asp">���·���ȫվ</a>��
	<hr>
	<form name="form1" method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=addblogpassword" onSubmit='return Verify()'>
	���룺
	<input type="password" name="password" /> 
	<input type="submit" name="Submit" value="ȫվ����" />        
	</form>
	<br />
	<form name="form2" method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=unblogpassword" />
	<input type="submit" name="Submit" value="�����վ������뱣��" />
	</form>
	</ul>
</div>
<SCRIPT language="javascript">
function Verify()
{
	v = del_space(document.form1.password.value);
     if (v.length == 0)
     {
        alert("���벻��Ϊ��!");
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
	oblog.showok "������վ����ɹ�,�����¸���ȫվ�ſɻ�ð�ȫ�ļ��ܱ�����",""
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
	oblog.showok "ȡ������ɹ�,�����¸���ȫվ�ſ�ȫ�����ܣ�",""
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
<h1 style='font-size:13px;'>�޸��ҵ�վ����Ϣ</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
	<%if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then%>
    <ul>
	<li class="t1">������</li>
	<li class="t2"><input name="user_domain" type="text" value="<%=oblog.filt_html(trim(rs("user_domain")))%>" size=10 maxlength=20 <%=sstr%> /> <select name="user_domainroot" <%=sstr%>><%=oblog.type_domainroot(rs("user_domainroot"))%></select><input type="hidden" name="old_userdomain" value="<%=oblog.filt_html(rs("user_domain"))%>"></li>
    </ul>
	<%end if%>
	<%if true_domain=1 and oblog.logined_ulevel=8 then%>
    <ul>
	<li class="t1">���ҵĶ���������</li>
	<li class="t2"><input name="custom_domain" type="text" value="<%=oblog.filt_html(trim(rs("custom_domain")))%>" size=20 maxlength=50 <%=sstr%> /> </li>
    </ul>
	<%end if%>
    <ul>
	<li class="t1">վ�����ƣ�</li>
	<li class="t2"><input name="blogname" type="text" value="<%=oblog.filt_html(rs("blogname"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul> 
	<li class="t1">վ�����</li>
	<li class="t2"><select name="user_classid" ><%=oblog.show_class("user",rs("user_classid"),0)%></select></li>
    </ul>
	<ul> 
	<li class="t1">Ĭ�ϱ༭����</li>
	<li class="t2"><input type="radio" value="0" name="isubbedit" <%if rs("isubbedit")=0 then response.write "checked"%> />
	html�༭��&nbsp;&nbsp; <input type=radio value="1" name="isubbedit" <%if rs("isubbedit")=1 then response.write "checked"%> />ubb�༭��</li>
    </ul>

     <ul> 
	<li class="t1">�����Ҽ����û��Ŷӣ�</li>
	<li class="t2"><input type="radio" value="1" name="en_blogteam" <%if rs("en_blogteam")<>0 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type=radio value="0" name="en_blogteam" <%if rs("en_blogteam")=0 then response.write "checked"%> />��</li>
    </ul>
	 <ul> 
	<li class="t1">�Ƿ�����ת��url��</li>
	<li class="t2"><input type="radio" value="1" name="hideurl" <%if rs("hideurl")=1 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type="radio" value="0" name="hideurl" <%if rs("hideurl")=0 then response.write "checked"%> />��</li>
    </ul>
	<ul> 
	<li class="t1">���������Ƿ����б���ʾ��</li>
	<li class="t2"><input type="radio" value="1" name="sublist" <%if sublist=1 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type="radio" value="0" name="sublist" <%if sublist=0 then response.write "checked"%> />��</li>
    </ul>
	<ul> 
	<li class="t1">��������Ƿ������ʾ��</li>
	<li class="t2"><input type="radio" value="1" name="u_frame" <%if u_frame=1 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type="radio" value="0" name="u_frame" <%if u_frame=0 then response.write "checked"%> />��</li>
    </ul>

	 <ul> 
	<li class="t1">����Ĭ�ϲ�����ʾ������</li>
	<li class="t2"><input name="user_showlogword_num" type="text"  value="<%=rs("user_showlogword_num")%>" size="5" /></li>
    </ul>
	 <ul> 
	<li class="t1">ÿҳ��ʾ����ƪ����</li>
	<li class="t2"><input name="user_showlog_num" type="text" id="user_showlog_num" value="<%=rs("user_showlog_num")%>" size="5" /></li>
    </ul>
	<ul> 
	<li class="t1">ÿ����ʾͼƬ����</li>
	<li class="t2"><input name="user_photorow_num" type="text" id="user_photorow_num" value="<%=rs("user_photorow_num")%>" size="5" /></li>
    </ul>
	 <ul> 
	<li class="t1">��ʾ���»ظ�������</li>
      <li class="t2"><input name="user_shownewcomment_num" type="text" value="<%=rs("user_shownewcomment_num")%>" size="5" /></li>
    </ul>
    <ul> 
	<li class="t1">��ʾ��������������</li>
	<li class="t2"><input name="user_shownewlog_num" type="text" value="<%=rs("user_shownewlog_num")%>" size="5" /></li>
    </ul>
    <ul> 
	<li class="t1">��ʾ��������������</li>
      <li class="t2"><input name="user_shownewmessage_num" type="text" value="<%=rs("user_shownewmessage_num")%>" size="5" /></li>
	  </ul>
	 <ul>
	<li class="t1">������������˳��</li>
	<li class="t2"> <input type="radio" value="1" name="comment_isasc" <%if rs("comment_isasc")=1 then response.write "checked"%> />ʱ��˳�� &nbsp;&nbsp; <input type="radio" value="0" name="comment_isasc" <%if rs("comment_isasc")=0 then response.write "checked"%> />ʱ�䵹��</li>
    </ul>
	<ul> 
	<li class="t1">վ���飺</li>
	<li class="t2"><textarea name="siteinfo" cols="40" rows="5"><%=oblog.filt_html(rs("siteinfo"))%></textarea></li>
    </li>
	</ul>

    <ul>
	<li class="t1"></li>
	<li class="">
	<input name="action" type="hidden" value="savesitesetup" /> 
	<input type="submit" value=" ���� " />
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
		if user_domain="" or oblog.strLength(user_domain)>20  then oblog.adderrstr("��������Ϊ��(���ܴ���14���ַ�)��")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("��������С��4���ַ���")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("������ϵͳ������ע�ᣡ")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
		if user_domainroot="" then oblog.adderrstr("����������Ϊ�գ�")
	end if
	if oblog.strLength(siteinfo)>255 then oblog.adderrstr("վ���鲻�ܴ���255���ַ���")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("��ʵ�����к���ϵͳ��������ַ���")
	if not IsNumeric(user_showlogword_num) then
		oblog.adderrstr("����Ĭ�ϲ�����ʾ��������Ϊ���֣�")
	end if	
	if not IsNumeric(user_showlog_num) then
		oblog.adderrstr("ÿҳ��ʾ������������Ϊ���֣�")
	else
		user_showlog_num=clng(user_showlog_num)
		if user_showlog_num>50 then oblog.adderrstr("ÿҳ��ʾ������������С��50��")
	end if
	if not IsNumeric(user_photorow_num) then
		oblog.adderrstr("ÿ����ʾͼƬ��������Ϊ���֣�")
	else
		user_photorow_num=clng(user_photorow_num)
		if user_photorow_num>50 then oblog.adderrstr("ÿ����ʾͼƬ��������С��50��")
	end if
	
	if not IsNumeric(user_shownewcomment_num) then
		oblog.adderrstr("��ʾ���»ظ���������Ϊ���֣�")
	else
		user_shownewcomment_num=clng(user_shownewcomment_num)
		if user_shownewcomment_num>50 then oblog.adderrstr("��ʾ���»ظ��������ܴ���50��")
	end if
	
	if not IsNumeric(user_shownewlog_num) then
		oblog.adderrstr("��ʾ����������������Ϊ���֣�")
	else
		user_shownewlog_num=clng(user_shownewlog_num)
		if user_shownewlog_num>50 then oblog.adderrstr("��ʾ���������������ܴ���50��")
	end if
	
	if not IsNumeric(user_shownewmessage_num) then
		oblog.adderrstr("��ʾ����������������Ϊ���֣�")
	else
		user_shownewmessage_num=clng(user_shownewmessage_num)
		if user_shownewmessage_num>50 then oblog.adderrstr("��ʾ���������������ܴ���50��")
	end if
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 and oblog.setup(54,0)=1 then
		set rs=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"' and userid<>"&oblog.logined_uid)
		if not rs.eof or not rs.bof then oblog.adderrstr("ϵͳ���Ѿ�������������ڣ������������")
	end if
	if true_domain=1 and oblog.logined_ulevel=8 and custom_domain<>"" then
		if oblog.chk_badword(custom_domain)>0 then oblog.adderrstr("�󶨵Ķ��������к���ϵͳ��������ַ���")
		set rs=oblog.execute("select userid from oblog_user where custom_domain='"&oblog.filt_badstr(custom_domain)&"'"&" and userid<>"&oblog.logined_uid)
		if not rs.eof or not rs.bof then oblog.adderrstr("ϵͳ���Ѿ��������˰�������������������������������ϵ����Ա��")		
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
	oblog.showok "�������óɹ�!",""		
end sub

sub saveblogstar()
	dim rs,picurl,bloginfo,blogname
	picurl=trim(request("picurl"))
	bloginfo=trim(request("bloginfo"))
	blogname=trim(request("blogname"))
	if picurl="" or oblog.strLength(picurl)>250  then oblog.adderrstr("ͼƬ���ӵ�ַ����Ϊ��,�Ҳ��ܴ���250���ַ���")
	if bloginfo="" or oblog.strLength(bloginfo)>250  then oblog.adderrstr("վ����ܲ���Ϊ��,�Ҳ��ܴ���250���ַ���")
	if blogname="" or oblog.strLength(blogname)>250  then oblog.adderrstr("�û�������Ϊ��,�Ҳ��ܴ���50���ַ���")
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
	oblog.showok "�ύ��ɣ���ȴ�����Ա���ͨ����",""
end sub

sub blogstar()
	Dim rs,strTitle,strBlogName,strPicUrl,strBlogInfo,intState
	Set rs=Server.CreateObject("ADODB.RecordSet")	
	rs.open "select * from oblog_blogstar Where userid=" & oblog.logined_uid,conn,1,1
	If rs.Eof Then
		strTitle="��Ŀǰ��û������"
		intState=-1
	Else
		strPicUrl=rs("picurl")
		strBlogName=rs("blogname")
		strBlogInfo=rs("info")
		intState=rs("ispass")
		If intState=1 Then
			strTitle="��Ŀǰ�Ѿ����û����ǣ����ϲ��ɸ���"
		Else	
			strTitle="��Ŀǰ���ڵȴ�����У������޸�֮ǰ�ύ������"
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
<h1 style='font-size:13px;'>�����û�����: <%=strTitle%></h1>
<div id="list">
	
	<ul class="list_edit">
	ͼƬ��ַ���Է��������Ƭ��վ��logo����վ������ͼ��blog�����벻Ҫ����50�֡�<br />
	ע��ͼƬ�ߴ������С��130*100���ң��Ա����Ա������
	<hr /></ul>
	<form  method="post" action="HomestayBackctrl-PassSettingCooperate.asp?action=saveblogstar">
	<ul>
	<li class="t1">�û����֣�</li>
	<li class="t2"><input type="text" maxlength="50" size="60" name="blogname" value="<%=strBlogname%>" /></li>
	</ul>
	<ul>
	<li class="t1">ͼƬ��ַ��</li>
	<li class="t2"><input type="text" maxlength="250" size="60" name="picurl" value="<%=strPicUrl%>" /></li>
	</ul>
	<ul>
	<li class="t1">blog���ܣ�</li>
	<li class="t2"><textarea name="bloginfo" cols="60" rows="5"><%=strBlogInfo%></textarea></li></ul>
	<ul class="list_bot">
	<%
	Select Case intState
		Case -1
	%>
		<input type="submit" value="�ύ��������"></ul>
		<%
		Case 0
		%>
		<input type="submit" value="�޸���������"></ul>
		<%
		Case 1
		%>
		�Ѿ���ȷ��Ϊ�û����ǣ����ϲ��ɸ��ģ������Ҫ�޸Ļ����������Ա��ϵ��</BR>
		
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
<h1 style='font-size:13px;'>�޸��ҵ�����</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul>
      <li class="t1">ԭ���룺</li>
      <li class="t2"><input name="oldpassword" type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
      <li class="t1">�����룺</li>
      <li class="t2"><input name="newpassword"  type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	 <ul> 
      <li class="t1">�ظ����룺</li>
      <li class="t2"><input name="newpassword1" type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
	<li class="t1"></li>
	<li class="t2">
	<input name="action" type="hidden" value="saveuserpassword" /> 
	<input type=submit value=" �޸����� " />
	</li>
    </ul>
	</form>
</div>
<% End If%>
<% If Request("pass")<>1 Then%>

<h1 style='font-size:13px;'>�����ҵ��һ���������</h1>
<div id="list">
	<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul> 
      <li class="t1">�������½���룺</li>
      <li class="t2"><input name="password"  type="password" size="30" maxlength="20" /><font color="#FF0000"> *</font></li>
    </ul>
    <ul> 
      <li class="t1">������ʾ���⣺</li>
      <li class="t2"><input name=question  onFocus="this.select(); " type="text" size="30" maxlength="20" value="<%=oblog.filt_html(rs(0))%>"><font color="#FF0000"> *</font>&nbsp;&nbsp;&nbsp;(�����޸ģ�����д)</li>
    </ul>
	 <ul> 
      <li class="t1">�һ�����𰸣�</li>
      <li class="t2"><input name=answer  type="text" size="30" maxlength="20"><font color="#FF0000"> *</font></li>
    </ul>
	<ul> 
		<li class="t1"></li>
		<li class="t2">
	  	<input name="action" type="hidden" value="saveuserlostpass"> 
        <input type=submit value=" ���� ">
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
	if password="" then oblog.adderrstr("���󣺵�½���벻��Ϊ��!")
	if question="" then oblog.adderrstr("������ʾ���ⲻ��Ϊ�գ�")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select question,answer from oblog_user where userid="&oblog.logined_uid&" and password='"&md5(password)&"'",conn,1,3
	if rs.eof then
		rs.close
		set rs=nothing
		oblog.adderrstr("���󣺵�½�����������!")
		oblog.showusererrCooperate
		exit sub
	else
		rs("question")=question
		if answer<>"" then rs("answer")=md5(answer)
		rs.update
		rs.close
		set rs=nothing
		oblog.showok "�޸��һ��������ϳɹ���",""
	end if

end sub

sub saveuserpassword
	dim oldpassword,newpassword,rs
	oldpassword=trim(request("oldpassword"))
	newpassword=trim(request("newpassword"))
	if oldpassword="" then oblog.adderrstr("����ԭ���벻��Ϊ��!")
	if newpassword="" or oblog.strLength(newpassword)>14 or oblog.strLength(newpassword)<4 then oblog.adderrstr("���������벻��Ϊ��(���ܴ���14С��4)��")
	if newpassword<>trim(request("newpassword1")) then oblog.adderrstr("�����ظ������������!")
	if oblog.errstr<>"" then oblog.showusererrCooperate:exit sub
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select password from oblog_user where userid="&oblog.logined_uid&" and password='"&md5(oldpassword)&"'",conn,1,3
	if rs.eof then
		rs.close
		set rs=nothing
		oblog.adderrstr("����ԭ�����������!")
		oblog.showusererrCooperate
		exit sub
	else
		rs("password")=md5(newpassword)
		rs.update
		rs.close
		set rs=nothing
		oblog.savecookie oblog.logined_uname,newpassword,0,""
		oblog.showok "�޸�����ɹ�,�´���Ҫ���µ�¼��",""
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
<h1 style='font-size:13px;'>�޸��ҵĸ�����Ϣ</h1>
<div id="list">
<form name="oblogform" action="HomestayBackctrl-PassSettingCooperate.asp" method="post">
    <ul>
	<li class="t1">�û�id��</li>
	<li class="t2"><%=rs("userName")%></li>
    </ul>
    <ul>
	<li class="t1">�ǳƣ�</li>
	<li class="t2"><input name="nickname" type="text" value="<%=oblog.filt_html(rs("nickname"))%>" size="30" maxlength="20" /><input type="hidden" name="o_nickname" value="<%=oblog.filt_html(rs("nickname"))%>" /></li>
    </ul>
    <ul>
	<li class="t1">��ʵ������</li>
	<li class="t2"><input name="truename"   type="text" value="<%=oblog.filt_html(rs("truename"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul>
	<li class="t1">�Ա�</li>
	<li class="t2"><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type="radio" value="0" name="sex" <%if rs("Sex")=0 then response.write "checked"%> />Ů</li>
    </ul>
	 <ul>
	<li class="t1">ʡ/�У�</li>
	<li class="t2"> <%=oblog.type_city(rs("province"),rs("city"))%> </li>
    </ul>
	 <ul>
	<li class="t1">�������ڣ�</li>
	<li class="t2">
		<input value="<%=year(rs("birthday"))%>" name="y" size="4" maxlength="4" />��
		<input value="<%=month(rs("birthday"))%>" name="m" size="2" maxlength="2" />��
		<input value="<%=day(rs("birthday"))%>"  name="d" size="2" maxlength="2" />��
	  </li>
    </ul>
	 <ul>
	<li class="t1">ְҵ��</li>
	<li class="t2"> <%oblog.type_job(rs("job"))%></li>
    </ul>
  <ul>
	<li class="t1">Email��</li>
	<li class="t2"><input name="Email" value="<%=oblog.filt_html(rs("userEmail"))%>" size="30" maxlength="50" /> 
      </li>
    </ul>
    <ul>
	<li class="t1">��ҳ��</li>
	<li class="t2"><input maxlength="100" size="30" name="homepage" value="<%=oblog.filt_html(rs("Homepage"))%>" /></li>
    </ul>
    <ul>
	<li class="t1">QQ���룺</li>
	<li class="t2"><input name="qq" value="<%=oblog.filt_html(rs("qq"))%>" size="30" maxlength="20" /></li>
    </ul>
    <ul>
	<li class="t1">MSN��</li>
	<li class="t2"><input name="msn" value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></li>
    </ul>
	 <ul>
	<li class="t1">�绰��</td>
	<li class="t2"><input name="tel" value="<%=oblog.filt_html(rs("tel"))%>" size="30" maxlength="50" /></li>
    </ul>
	<ul>
	<li class="t1">ͨ�ŵ�ַ��</li>
	<li class="t2"><input name="address" value="<%=oblog.filt_html(rs("address"))%>" size="40" maxlength="250" /></li>
    </ul>
	  <ul>
	  <li class="t1"></li>
	  <li class="t2"><input name="action" type="hidden" value="saveuserinfo" /> 
        <input type=submit value=" ���� " /></li>
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
	if oblog.chk_regname(nickname) then oblog.adderrstr("���ǳ�ϵͳ������ע�ᣡ")
	if oblog.chk_badword(nickname)>0 then oblog.adderrstr("�ǳ��к���ϵͳ��������ַ���")
	if oblog.strLength(nickname)>50 then oblog.adderrstr("�ǳƲ��ܲ��ܴ���50�ַ���")
	if oblog.setup(25,0)=0 and nickname<>"" and nickname<>trim(request("o_nickname")) then
		set rs=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
		if not rs.eof or not rs.bof then oblog.adderrstr("ϵͳ���Ѿ�������ǳƴ��ڣ�������ǳƣ�")		
	end if
	if birthday = "--" then
		birthday=""
	else
		if not isdate(birthday) then oblog.adderrstr("�������ڸ�ʽ����")
		if clng(trim(request("y")))>2060 then oblog.adderrstr("������ݹ���")
		if clng(trim(request("y")))<1900 then oblog.adderrstr("������ݹ�С��")
	end if
	if email="" then oblog.adderrstr("�����ʼ���ַ����Ϊ�գ�")
	if not oblog.IsValidEmail(email) then oblog.adderrstr("�����ʼ���ַ��ʽ����")
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
	oblog.showok "�������ϳɹ�!",""
end sub

%>