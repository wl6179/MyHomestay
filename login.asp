<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
'�˳�ϵͳ
'If request("logout")="logout" Then
'	If cookies_domain <> "" Then
'        response.Cookies(cookies_name).domain = cookies_domain
'    End If
'	Response.Cookies(cookies_name)("username")=oblog.CodeCookie("")
'	Response.Cookies(cookies_name)("password")=oblog.CodeCookie("")
'	Response.Cookies(cookies_name)("userurl")=oblog.CodeCookie("")
'	if request("re")="1" then
'		'response.Redirect(oblog.comeurl)
'		response.Write("<script language=JavaScript>top.location='"&oblog.comeurl&"';< /script>")
'	else'�˳���ص���ҳ��
'		response.Write("< script language=JavaScript>top.location='login.asp';</ script>")
'	end if
'	Set oBlog=Nothing
'	response.End()
'End If
%>
<%
dim username,password,show_login,CookieDate,fromurl,action
action=request("action")
if action<>"showindexlogin" And action<>"showindexlogin_team" and action<>"showjs" then
	if  oblog.checkuserlogined() And session("indexHomeLogined")<>"indexHomeLogined" then response.Redirect("HomestayBackctrl-indexHome.asp")'����û��Ѿ���ת�� ����ҳ��󣬾Ͳ���Ҫ����ת�ˣ�ֱ�ӷ��ʵ�ǰloginҳ�棻
	if  oblog.checkuserlogined_team() And session("indexCooperateLogined")<>"indexCooperateLogined" then response.Redirect("HomestayBackctrl-indexCooperate.asp")'����û��Ѿ���ת�� ����ҳ��󣬾Ͳ���Ҫ����ת�ˣ�ֱ�ӷ��ʵ�ǰloginҳ�棻
end if

'��ע�빦��WL[2007-09-23];��������___oblog.filt_badstr()������
username=oblog.filt_badstr(trim(request.Form("username")))
password=trim(request.form("password"))
CookieDate=trim(request("CookieDate"))
fromurl=trim(request("fromurl"))
if username<>"" or request("chk")="1" then '��Ҫcheck��ʱ��ִ�м�⣻	
	call sub_chklogin
else
	if action="showindexlogin" then
		call sub_showindexlogin()
	elseif action="showindexlogin_team" then
		call sub_showindexlogin_team()
		
	elseif action="showjs" then
		blogurl=oblog.setup(3,0)
		call sub_showindexlogin()
	elseif action="showjs_team" then
		blogurl=oblog.setup(3,0)
		call sub_showindexlogin_team()
		
	else
		call sub_showlogin()
	end if
end if


sub sub_showlogin()

	call sysshow_4()  '���Ȼ��ģ��~...�����£������ж��Ǵ���ע�� ���� ��ʾ��ע����ˣ�
	select case action
		case "chkreg"
		call aaa()'����ע��;
		case else 
		'''call aaa()'��ʾ��ע�����;
	end select
	'show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'����ǩ��$show_list$��Ϊ ���е����� show_reg �� + �ײ���Ȩ��Ϣ���룻
	show=replace(show,"$show_list$",fromurl)&oblog.site_bottom'����ǩ��$show_list$��Ϊ ���е����� show_reg �� + �ײ���Ȩ��Ϣ���룻
	response.Write show'ȫ�������

end sub
%>


<% sub aaa()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <meta name="Robots" content="index,follow" />
	<meta name="author" content="����, wangliang6179@163.com" />
	<meta name="copyright" content="www.myhomestay.com.cn,�ҵ�ס����" />
	<meta name="description" content="homestay,myhomestay,china,beijing,�й�,����,ס��,���ס��,����ס��,ס�����,���,�����,����,�й���ͥ,����,rent" />
	<meta name="keywords" content="homestay,myhomestay,china,beijing,learn,english,�й�,����,ס��,����,��ͥ,����" />
    <title>
      �ҵ�ס����- ��¼�� homestay in beiijng|guest house in beijing|Learn Chinese in China
    </title>
    <link rel="stylesheet" type="text/css" href="css/StyleIndex.css" />

	<link rel="stylesheet" type="text/css" href="inc/divCN-EN.css" />
<script type="text/javascript" language="javascript" src="images/blog/tab2.js"></script>
  </head>
  <body>



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
					<span id="joincompany_login"><a href="/login.asp">��ҵ��¼</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">��ҵ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
					<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>&nbsp;
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
					<div id="main-TextArea01" style="text-align:right">
					
					
					
						<center><h2 style="font-size:15px">��¼��</h2></center>
						
						
						
<form name="reg" method="post" action="login.asp?fromurl=<%=fromurl%>">
<p class="p1" style="margin:13px;">�ʻ�����<input name="username" type="text" class="put1" size="25" />
</p>
<p class="p2" style="margin:13px;">��������<input name="password" type="password" class="put1" size="25" />
</p>
<p class="p3" style="margin:13px;">
	<input name="CookieDate" type="radio" value="0" checked="checked" title="���������Լ��ĵ��Ե�½���벻Ҫ��������" />������
	<input type="radio" name="CookieDate" value="1" />һ��
	<input type="radio" name="CookieDate" value="2" />һ��
	<input type="radio" name="CookieDate" value="3" />һ��</p>
<p class="p4" style="margin:13px;"></p>
<div class="but" style="margin:13px;">
	<input name="Submit" type="image" src="OblogStyle/OblogStyleAdminImages/login.gif" value="�ύ" onclick="submit()" /></div>
	<div class="code" style="margin:13px;">
	<%if oblog.setup(58,0)=1 then
	%>��֤�룺<input name="codestr" type="text" class="put2" size="6" maxlength="4" />
	<%response.Write(oblog.getcode)
	end if
	%>
	</div></p>
	</form>



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
							
							</p>
							
					
							
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

end sub



sub sub_chklogin()
	if oblog.chkiplock() then
		oblog.adderrstr("�Բ������IP�ѱ�����,�������½����")
		oblog.showerr
		exit sub
	end if
	if oblog.setup(58,0)=1 then
		if not oblog.codepass then oblog.adderrstr("��֤�������ˢ�º��������룡"):oblog.showerr
	end if
	if UserName="" then oblog.adderrstr("��½�û�������Ϊ�գ�")
	if Password="" then oblog.adderrstr("��½���벻��Ϊ�գ�")
	if oblog.errstr<>"" then oblog.showerr:exit sub
	if CookieDate="" then CookieDate=0	else CookieDate=Clng(CookieDate)
	password=md5(password)
	if is_ot_user=1 then
		call ot_chklogin()
	else
		If Request("teamSubmit")<>"teamSubmit" Then
		oblog.ob_chklogin UserName,password,CookieDate
		Else
		oblog.ob_chklogin_team UserName,password,CookieDate
		End If
	end if
	if fromurl<>"" then
		response.Redirect(replace(fromurl,"$","&"))
	else
		if action="showindexlogin" Or action="showindexlogin_team" then
			if instr(oblog.comeurl,"/err.asp")>0 then
				response.Redirect("index.asp")
			else
				response.Redirect(oblog.comeurl)
		
			end if
		else
			response.Redirect("HomestayBackctrl-indexHome.asp")
		end if
	end if
end sub



sub ot_chklogin()
	dim sql,rs,rsreg
	if not IsObject(ot_conn) then link_database
	sql="select * from "&ot_usertable&" where "&ot_username&"='"& username & "' and "&ot_password&" ='" & password &"'"
	set rs=ot_conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.adderrstr("�û���������������������룡"):oblog.showerr
		exit sub
	else
		set rsreg=server.CreateObject("adodb.recordset")
		rsreg.open "select * from [oblog_user] where username='"& username &"'",conn,1,3
		if rsreg.eof then
			dim reguserlevel
			if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
			set rsreg=server.CreateObject("adodb.recordset")
			rsreg.open "select top 1 * from [oblog_user]",conn,1,3
			rsreg.addnew
			rsreg("username")=username
			rsreg("password")="othertable"
			rsreg("user_dir")=oblog.setup(30,0)
			rsreg("user_level")=reguserlevel
			rsreg("lockuser")=0
			rsreg("en_blogteam")=1
			rsreg("adddate")=now()
			rsreg.update
			oblog.execute("update oblog_user set user_folder=userid where username='"&username&"'")
			oblog.execute("update oblog_setup set user_count=user_count+1")
			rsreg.close
			set rsreg=nothing
			oblog.SaveCookie username,password,0," "
			oblog.CreateUserDir username,1
			set rs=nothing
			oblog.showok "���ǵ�һ�μ���blogϵͳ��������blog����!","user_setting.asp"
			response.End()
		else
			rsreg("LastLoginIP")=Request.ServerVariables("REMOTE_ADDR")
			rsreg("LastLoginTime")=now()
			rsreg("LoginTimes")=rsreg("LoginTimes")+1
			rsreg.update			
		end if
		rsreg.close
		set rsreg=nothing
		set rs=nothing
		if isobject(ot_conn) then ot_conn.close:set ot_conn=nothing
		oblog.SaveCookie username,password,CookieDate,""
		if action="showindexlogin" then
			'response.Redirect(oblog.comeurl)
			response.Redirect("HomestayBackctrl-index.asp")
		else
			response.Redirect("HomestayBackctrl-index.asp")
		end if	
	end if
end sub

sub sub_showindexlogin()
	dim show_userlogin
	'if oblog.CheckUserLogined()=False Or oblog.logined_ulevel_team=1 then
	if oblog.CheckUserLogined()=False  then
			

			show_userlogin="<form style='color:#333333;' action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method='post' name='UserLogin'>" 
			  show_userlogin=show_userlogin&"<ul style='padding:0px;margin:8px 0px 0px 15px; font-size:13px;'><li style='padding:0px;margin:0px'><span>��&nbsp;ͥ&nbsp;��&nbsp;��&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input id='UserName' name='UserName' type='text' size='13'></li>"
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>��&nbsp;ͥ&nbsp;��&nbsp;��&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input name='Password' type='password' id='Password' size='13' /></li>"
			  if oblog.setup(58,0)=1 then
				  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>��&nbsp;֤&nbsp;�룺</span><input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />&nbsp;"&oblog.getcode&"</li>"
			  end if 
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><input name='Login' type='submit' id='Login' value='�� ¼' style='font-size:12px'>&nbsp;<input name='regHomestay' type='button' value=' �� �� �� �� ' onClick=""window.location.href='"&blogurl&"RegisterHome.asp'""><br /><a href='"&blogurl&"lostpassword.asp'>�������룿</a></li>"


				show_userlogin=show_userlogin&"</ul>"
				show_userlogin=show_userlogin&"</form>"


				

			
		
		
'		show_userlogin="<form action="""&blogurl&"login.asp?action=showindexlogin&chk=1"" method='post' name='UserLogin'>" 
'    	show_userlogin=show_userlogin&"<table align='center' width='100%' border='0' cellspacing='0' cellpadding='0' style=""font-size:12px"">" 
'        show_userlogin=show_userlogin & "<tr><td height='25' class=login>�û�����<input name='UserName' type='text' id='UserName' size='12' maxlength='20'></td></tr>" 
'        show_userlogin=show_userlogin & "<tr><td height='25' class=login >�ܡ��룺<input name='Password' type='password' id='Password' size='12' maxlength='20'></td></tr>" 
'		if oblog.setup(58,0)=1 then
'			show_userlogin=show_userlogin&"<tr><td height='25' >��֤�룺<input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />"&oblog.getcode&"</td></tr>"
'		end if
'        show_userlogin=show_userlogin & "<tr><td height='25'>��������<input type=""checkbox"" name=""CookieDate"" value=""3"">��ס����</td></tr>" 
'		
'		show_userlogin=show_userlogin & "<tr><td height='25'><input name='Login' type='submit' id='Login' value='��¼'> " 
'        show_userlogin=show_userlogin & "<a href='"&blogurl&"reg.asp'>�û�ע��</a>&nbsp;<a href='"&blogurl&"lostpassword.asp'>��������</a><br></td>"       
'        show_userlogin=show_userlogin & "</tr></table></form>" 
	else
	
		
		show_userlogin="<table style='color:#333333;' border='0' align='center' style=' font-size:13px;'>"
			
                show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>��ӭ�� <font color='#DE728A'>" & oblog.logined_uname & "</font><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				
				show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>������ݣ�"
						  if oblog.logined_ulevel=6 then
							show_userlogin= show_userlogin&"<font color='#038ad7'>ע���ͥ�û�</font>"
						  elseif oblog.logined_ulevel=7 then
							show_userlogin= show_userlogin&"ע����Ŀ����"
						  elseif oblog.logined_ulevel=8 then
							show_userlogin= show_userlogin&"VIP��Ŀ����"
						  elseif oblog.logined_ulevel=9 then
							show_userlogin= show_userlogin& "ǰ̨������"
						  end if 
						  'show_userlogin= show_userlogin& "<br><b>�û�������壺</b><br>" & vbcrlf
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<br /><a style='color:#ffffff;' href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>��Ŀ��ҳ</a>"
						  else
						  show_userlogin= show_userlogin& "<br />"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<br /><img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-index.asp' ><strong>������Ŀ�ռ����</strong></a><br />"
						  elseif oblog.logined_ulevel_team=1 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-indexCooperate.asp' >��������������</a>&nbsp;&nbsp;&nbsp;"
						  else
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-indexHome.asp' ><strong>�����ҵļ�ͥ����</strong></a>&nbsp;&nbsp;&nbsp;"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  show_userlogin= show_userlogin& "<a href='"&blogurl&"user_post.asp' >������Ŀ����</a>&nbsp;&nbsp;"
						  else
						  
						  end if

						  
						  show_userlogin= show_userlogin& "<img src='/images/logout.gif' /><a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>�˳�</a><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				

				


		
		
'		show_userlogin="--��ӭ��," & oblog.logined_uname & "--<br />"
'		show_userlogin= show_userlogin&"������ݣ�"
'		if oblog.logined_ulevel=7 then
'			show_userlogin= show_userlogin&"ע����Ŀ����"
'		elseif oblog.logined_ulevel=8 then
'			show_userlogin= show_userlogin&"VIP��Ŀ����"
'		elseif oblog.logined_ulevel=9 then
'			show_userlogin= show_userlogin& "ǰ̨������"
'		end if
'		'show_userlogin= show_userlogin& "<br><b>�û�������壺</b><br>" & vbcrlf
'		show_userlogin= show_userlogin& "<br /><a href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>�ҵ���ҳ</a>"
'		show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-index.asp' target='_blank'>��������</a><br />" 
'		show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?url=user_post.asp' target='_blank'>��������</a>&nbsp;&nbsp;"
'		show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>ע����¼</a><br />" 
	end if
	Response.Write oblog.htm2js_div(show_userlogin,"ob_login")
end sub

sub sub_showindexlogin_team()
	dim show_userlogin
	
	'if oblog.CheckUserLogined_team()=False Or oblog.logined_ulevel_team=0 then'���Ҳ��Ǻ������ĵ�¼��
	if oblog.CheckUserLogined_team()=False then'���Ҳ��Ǻ������ĵ�¼��
			

			show_userlogin="<form style='color:#333333;' action="""&blogurl&"login.asp?action=showindexlogin_team&chk=1"" method='post' name='UserLogin_team'>" 
			  show_userlogin=show_userlogin&"<ul style='padding:0px;margin:8px 0px 0px 15px; font-size:13px;'><li style='padding:0px;margin:0px'><span>LoginName</span>" 
			    show_userlogin=show_userlogin&"<input id='UserName' name='UserName' type='text' size='13'></li>"
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>Password&nbsp;</span>" 
			    show_userlogin=show_userlogin&"<input name='Password' type='password' id='Password' size='13' /></li>"
			  if oblog.setup(58,0)=1 then
				  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><span>��&nbsp;֤&nbsp;�룺</span><input name=""codestr"" type=""text"" size=""4"" maxlength=""4"" />&nbsp;"&oblog.getcode&"</li>"
			  end if 
			  show_userlogin=show_userlogin&"<li style='padding:0px;margin:0px;line-height:18px;'><input name='Login' type='submit' id='Login' value='�� ¼' style='font-size:12px'>&nbsp;<input name='regHomestay' type='button' value='�� �� �� ��' onClick=""window.location.href='"&blogurl&"RegisterCooperate.asp'""><br /><a href='"&blogurl&"lostpassword.asp'>�������룿</a></li>"

			  show_userlogin=show_userlogin&"<input name='teamSubmit' type='hidden' id='teamSubmit' value='teamSubmit'>"
				
				show_userlogin=show_userlogin&"</ul>"
				show_userlogin=show_userlogin&"</form>"

	else
	
		
		show_userlogin="<table style='color:#333333; font-size:13px;' border='0' align='center'>"
			
                show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>��ӭ��," & oblog.logined_uname & "<br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				
				show_userlogin=show_userlogin&"<tr>" 
                  show_userlogin=show_userlogin&"<td align='center'>������ݣ�"
						  if oblog.logined_ulevel=6 then
							show_userlogin= show_userlogin&"<font color='#038ad7'>�������</font>"
						  elseif oblog.logined_ulevel=7 then
							show_userlogin= show_userlogin&"ע����Ŀ����"
						  elseif oblog.logined_ulevel=8 then
							show_userlogin= show_userlogin&"VIP��Ŀ����"
						  elseif oblog.logined_ulevel=9 then
							show_userlogin= show_userlogin& "ǰ̨������"
						  end if 
						  'show_userlogin= show_userlogin& "<br><b>�û�������壺</b><br>" & vbcrlf
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<br /><a style='color:#ffffff;' href="&blogurl&"blog.asp?name="&oblog.logined_uname&" target='_blank'>��Ŀ��ҳ</a>"
						  else
						  show_userlogin= show_userlogin& "<br />"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-index.asp' >��Ŀ����</a><br />"
						  elseif oblog.logined_ulevel_team=1 then
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<img src='/images/[MyHomestay]Manager.gif' /><a href='"&blogurl&"HomestayBackctrl-indexCooperate.asp' ><strong>�����ʺŹ�������</strong></a>&nbsp;&nbsp;&nbsp;"
						  else
						  show_userlogin= show_userlogin& "&nbsp;&nbsp;<a href='"&blogurl&"HomestayBackctrl-indexHome.asp' ><strong>�����ҵļ�ͥ����</strong></a>&nbsp;&nbsp;&nbsp;"
						  end if
						  
						  if oblog.logined_ulevel>=7 then
						  'show_userlogin= show_userlogin& "<a href='"&blogurl&"user_post.asp' target='_blank'>������Ŀ����</a>&nbsp;&nbsp;"
						  else
						  
						  end if

						  
						  show_userlogin= show_userlogin& "<a href='"&blogurl&"HomestayBackctrl-index.asp?t=logout&re=1'>�˳�</a><br />" 
                    show_userlogin=show_userlogin&"</td>" 
                show_userlogin=show_userlogin&"</tr>" 
				

	end if
	Response.Write oblog.htm2js_div(show_userlogin,"ob_login_team")
end sub
%>

