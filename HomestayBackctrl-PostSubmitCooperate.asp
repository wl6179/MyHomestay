<!--#include file="user_topCooperate.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/class_Trackback.asp"-->
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
					
					
					
						<center><h2>�ύ������� </h2></center>


<div id="main">

  <!--<div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
		<%
		If Request("ubb")="1" Then
			%>
		&nbsp; <a href="HomestayBackctrl-PostSubmitCooperate.asp?ubb=0&t=<%=t%>">�л���html��ʽ���<%=tName%></a>
		<%
		Else
		%>
		&nbsp; <a href="HomestayBackctrl-PostSubmitCooperate.asp?ubb=1&t=<%=t%>">�л���ubb��ʽ���<%=tName%>(firefox�Ƽ�)</a>
		<%End If%>
		&nbsp; <a href="user_blogmanage.asp?t=<%=t%>">����<%=tName%>����</a>
		&nbsp; <a href="user_blogmanage.asp?usersearch=5&t=<%=t%>">�ݸ���</a>
		&nbsp; <a href="user_subject.asp?t=<%=t%>"><%=tName%>�������</a>
		
	</div>
  </div>-->
  <div class="content">  
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			
	</div>
	
    <div class="content_body">
<%
If EN_PHOTO=0 And t="1" Then
	Response.Write ("<center>ϵͳδ������Ṧ��</center>")
Else
	dim action
	action=trim(request("action"))
	select case action
		case "savelog"
			call savelog
		case else
			call main()
	end select
End If
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
							<p class="public-title01" style="position:relative;float:left;width:140px;">
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
sub main()
'--------------------------
If oblog.logined_uisValidate=0 Then
	oblog.adderrstr("��ʱû��Ϊ����ͨ�˹��ܣ������ȴ�վ����˲���ͨ��лл����")
	oblog.showusererr
	exit sub
End If
'--------------------------
	dim rs,logid,photofile
	logid=clng(request.QueryString("logid"))
	photofile=trim(request.QueryString("photofile"))
	'ATFLAG 
	Dim log_tags,tags,filename,log_pics,log_type
	dim face,topic,classid,subjectid,logtext,istop,ishide,isencomment,showword,addtime,userid,ispassword,tburl,postchknum,oldisdraft
	oldisdraft=0
	if logid>0 then
		if oblog.logined_ulevel=9 then
			set rs=oblog.execute("select * from oblog_logCooperateSubmit where logid="&logid)
		else
			set rs=oblog.execute("select * from oblog_logCooperateSubmit where logid="&logid&" and ( userid="&oblog.logined_uid&"or authorid="&oblog.logined_uid&" )")
		end if
		if rs.eof then 
			set rs=nothing
			oblog.adderrstr("�޴�Ȩ�޲�����"& tName &"��")
			oblog.showusererr
		end if
		topic=oblog.filt_html(rs("topic"))
		face=rs("face")
		classid=rs("classid")
		subjectid=rs("subjectid")
		logtext=replace(rs("logtext"),"#isubb#","")
		istop=rs("istop")
		ishide=rs("ishide")
		isencomment=rs("isencomment")
		showword=rs("showword")
		addtime=rs("addtime")
		userid=rs("userid")
		ispassword=rs("ispassword")
		tburl=rs("tburl")
		oldisdraft=rs("isdraft")
		filename=rs("filename")	
		'----------------------------------
		'ATFLAG
		log_pics=rs("logpics")
		log_type=rs("logtype")
		If IsNull(rs("logtags")) Then	
			tags=""
		Else
			tags=rs("logtags")
		End If
		log_pics=rs("logpics")
		'----------------------------------
	end if
	if photofile<>"" then 
		logtext="<img src="""&photofile&""">"
		log_pics=photofile
	end if
	Randomize
	postchknum = Int(900*rnd)+1000
	if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
	Response.Cookies(cookies_name)("postchknum") = Cstr(postchknum)
	'Session("postchknum") = Cstr(postchknum)
	if isencomment="" then isencomment=1
	if userid="" then userid=oblog.logined_uid
	if oblog.setup(83,0)=1 then
		if filename="" then filename="Homestay"&"-"&Year(now) & Month(now) & Day(now)&hour(now())&minute(now())&second(now())
	else
		if filename="" then filename="�Զ����"
	end if
	call getteam()
%>
<SCRIPT language=javascript>
function hord(){
	if(divAdvance.style.display == "none"){
		divAdvance.style.display = "";
	}
	else{
		divAdvance.style.display = "none";
	}
}
function chkfilename()
{
	var filename=del_space(document.oblogform.filename.value);
	if (filename=="�Զ����"){document.oblogform.filename.value=""}
	if (filename==""){document.oblogform.filename.value="�Զ����"}
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i) < 'A' || string.charAt(i) > 'Z')&& (string.charAt(i)!='-')&& (string.charAt(i)!='_')) 
{
return 1;
}
}
return 0;//pass
}
function setdraft()
{
	document.oblogform.isdraft.value='1';
}
function VerifySubmit()
{
	submits(); 
	topic = del_space(document.oblogform.topic.value);
     if (topic.length == 0)
     {
        alert("��������д��Ŀ!");
	return false;
     }
	 
	 var needclassid=<%=oblog.setup(48,0)%>
	 if (needclassid==1) {
	 if (document.oblogform.classid.value == 0)
     {
        alert("��ѡ��<%=tName%>�����!");
	return false;
     }
	 }
	 
	 var filename=del_space(document.oblogform.filename.value);
	if ((checkerr(filename) == "1")&&(filename!="")&&(filename!="�Զ����")){
		alert("�ļ���������0-9�����ֺ�a-z�İ����ĸ���»���,���������ĺ͹��ַ�����!@#$%^�ȣ���")
	return false
	}
	
	 showword = del_space(document.oblogform.showword.value);
      for(i = 0; i < showword.length; i++){
	  if(showword.charAt(i) < "0" || showword.charAt(i) > "9"){
	  alert("��������ȷ����ʾ������");
	  return false;  
		 }
     }
	
 	if (document.oblogform.edit.value == "")
     {
        alert("������<%=tName%>����!");
	return false;
     }
	var date=document.oblogform.selecty.value+"-"+document.oblogform.selectm.value+"-"+document.oblogform.selectd.value
	var datereg=/^(\d{4})-(\d{1,2})-(\d{1,2})$/
	var datareg=/^(\d){1,2}$/
	if (!datereg.test(date)){
	  alert("<%=tName%>ʱ�������ʽ����");
	   return false;
	 }
	var r=date.match(datereg)
	var d=new Date(r[1],r[2]-1,r[3])
	if (!(d.getFullYear()==r[1]&&d.getMonth()==r[2]-1&&d.getDate()==r[3])){
	  alert("<%=tName%>ʱ�������ʽ����");
	   return false;
	 }
	 
  var FormLimit = 51100
  var TempVar = new String
  TempVar = document.oblogform.edit.value
  if (TempVar.length > FormLimit)
	{
	if (confirm("��Ҫ������<%=tName%>̫�������������Ϊ�����ֱַ𷢲���\n���������ύ��ע����Ҫ�ϳ�ʱ������ύ�ɹ���\n\n�Ƿ����ύ��") == false)
	return false;
	}
	
	  if (TempVar.length > FormLimit)
  {
    document.oblogform.edit.value = TempVar.substr(0, FormLimit)
    TempVar = TempVar.substr(FormLimit)
    while (TempVar.length > 0)
    {
      var objTEXTAREA = document.createElement("hidden")
      objTEXTAREA.name = "logtext"
      objTEXTAREA.value = TempVar.substr(0, FormLimit)
      document.oblogform.appendChild(objTEXTAREA)
      
      TempVar = TempVar.substr(FormLimit)
    }
  }	
 return true;
}


</script> 
<div id="list">
	<!--<form action="HomestayBackctrl-PostSubmitCooperate.asp?action=savelog&t=<%=t%>" method="post" name="oblogform" id="oblogform" onSubmit="return VerifySubmit()">-->
	<form action="HomestayBackctrl-PostSubmitCooperate.asp?action=savelog&t=<%=t%>" method="post" name="oblogform" id="oblogform" onSubmit="return VerifySubmit();">
    <ul class="list_edit">����ע�ͣ� 
	<input name="topic" type=text id="topic" size="50" maxlength="250" value="<%=topic%>"> <font color="#FF0000"> *</font></ul>		
	<ul style="display:none;" class="list_edit"> 
	���·��ࣺ

	<select name="classid" id="classid">	
<option value=10001></option>
	
	</select>
	<%if oblog.setup(48,0)=1 then response.Write("<font color=""#FF0000""> *</font>")%>
	���� <!--�ҵ�<%=tName%>����(ר��)��--> 
	<!--<select name="subjectid" id="subjectid">
	<option value="0">��ѡ��</option>
<%
set rs=oblog.execute("select subjectid,subjectname from oblog_subject where userid="&userid & " And subjectType=" & t)
while not rs.eof
	if rs(0)=subjectid then
		response.Write("<option value="&rs(0)&" selected>"&oblog.filt_html(rs(1))&"</option>")
	else
		response.Write("<option value="&rs(0)&" >"&oblog.filt_html(rs(1))&"</option>")
	end if
	rs.movenext
wend
%>				
	</select>-->
	<input type="hidden" name="postchknum" value="<%=postchknum%>" />
	<%if oblog.setup(60,0)=1 then%>
	���� ��֤�룺<input name="codestr" type="text" size="6" maxlength="4" /> <%=oblog.getcode%>
	<%end if%>
	</ul>
	<%
	'-------------------------------------------
	'ATFLAG
	If EN_TAGS=1 Then%>
	<!--<ul class="list_edit"> 
		<%=P_TAGS_DESC%>��
		<input name="logtags" type="text" size="40" maxlength="255" Value="<%=tags%>"/> 
		��<%
		Select Case P_TAGS_SPLIT
			Case " "
				Response.Write "�ո�"
			Case ","
				Response.Write "����"
			Case Else		
				Response.Write P_TAGS_SPLIT
		End Select
		%>�ָ�,���<%=P_TAGS_MAX%>��(<a href="user_help.asp#tag" target="_blank">ʲô��<%=P_TAGS_DESC%>?</a>)
	</ul>-->		
	<%	
	'-------------------------------------------
	End If%>
	<input type="hidden" id="edit" name="edit" value="<%if logtext<>"" then response.Write(Server.HtmlEncode(logtext))%>" />			
	<!--#include file="HomestayBackctrl-editCooperate.asp"--> 		
	
	
	<ul class="list_edit"> 
	<input type="hidden" name="isdraft" id="isdraft" value="0" /> 
	<input type="submit" value=" �ύ������� " />
	<%if t="0" then%>
	����<input type="submit" value="��ʱ����Ϊ�ݸ�" onClick="setdraft();" />
	����<!--<input type="button" onClick="part();" value="������ʾ���" />-->
	����<!--<input type="button" onClick="pastestr();" value="���ѻָ�" />-->			
	����<!--<input type="reset" value="�����д" />-->
	<%end if%>
	</ul>	
	<!--<h1 style="margin-left:40px; font-size:14px; margin-bottom:5px"> 
	<%=tName%>�߼�ѡ��(��ѡ)��<input type="checkbox" onClick="hord();"> 
	</h1>-->
	<div id="divAdvance" style="display:none">
		<ul class="list_edit"> 
			����ظ���<input name="isencomment" type="checkbox" value="1" <%'if isencomment=0 then response.Write("") else response.Write("checked") %>>
		���������ѿɼ�(����)��<input name="ishide" type="checkbox" value="1" <%if ishide=1 then response.Write("checked")%>>
		����<span <%If t="1" Then response.Write("style='display:none'")%>>��ҳ�̶���<input name="istop" type="checkbox" value="1" <%if istop=1 then response.Write("checked")%>></span>
		</ul>
		<ul class="list_edit"> 
		<%=tName%>���룺 
		<input name="ispassword" type="password" id="ispassword" size="15" maxlength="15" value="<%=ispassword%>">
		(������������ )����������ʾ������ 
		<input name="showword" type="text" id="showword" value="<%if showword<>"" then response.Write(showword) else response.Write(oblog.logined_ushowlogword)%>" size="6" maxlength="10">
		(ȫ����ʾ������) 
		</ul>					
		
		<ul class="list_edit" <%If t="1" Then response.Write("style='display:none'")%>> 
			<%=tName%>���飺 
			<input type="radio" name="face" value="0" <%if face=0 then response.Write("checked")%>>
			�� 
			<input name="face" type="radio" value="1" <%if face=1 then response.Write("checked")%>> <img src="images/face/1.gif" width="20" height="20"> 
			<input type="radio" name="face" value="2" <%if face=2 then response.Write("checked")%>> <img src="images/face/2.gif" width="20" height="20"> 
			<input type="radio" name="face" value="3" <%if face=3 then response.Write("checked")%>> <img src="images/face/3.gif" width="20" height="20"> 
			<input type="radio" name="face" value="4" <%if face=4 then response.Write("checked")%>> <img src="images/face/4.gif" width="20" height="20"> 
			<input type="radio" name="face" value="5" <%if face=5 then response.Write("checked")%>> <img src="images/face/5.gif" width="20" height="20"> 
			<input type="radio" name="face" value="6" <%if face=6 then response.Write("checked")%>> <img src="images/face/6.gif" width="18" height="20"> 
			<input type="radio" name="face" value="7" <%if face=7 then response.Write("checked")%>> <img src="images/face/7.gif" width="20" height="20"> 
			<input type="radio" name="face" value="8" <%if face=8 then response.Write("checked")%>> <img src="images/face/8.gif" width="20" height="20"> 
			<input type="radio" name="face" value="9" <%if face=9 then response.Write("checked")%>> <img src="images/face/9.gif" width="20" height="20">
		</ul>
		<ul class="list_edit">�ļ����ƣ� 
		<input name="filename" type=text id="filename" size="20" maxlength="30" value="<%=filename%>" onClick="chkfilename();"  onBlur="chkfilename();">.<%=f_ext%> (ֻ��ΪӢ�ġ����ּ��»��ߣ��30���ַ���)
</ul>	
		<ul class="list_edit" <%If t="1" Then response.Write("style='display:none'")%>> 
			�Ŷ��û��� 
			<select name="blogteam" id="s1" onChange="setSort(this,this.form.s2);">
			<option value="<%=userid%>" >�ҵ�BLOG</option>
		<%
		set rs=oblog.execute("select oblog_blogteam.mainuserid,oblog_user.blogname,oblog_blogteam.id from oblog_blogteam,[oblog_user] where oblog_blogteam.otheruserid="&userid&" and oblog_blogteam.mainuserid=[oblog_user].userid")
		while not rs.eof
			if clng(rs(0))=clng(userid) then 
				response.Write "<option value="&rs(0)&" selected>"&oblog.filt_html(rs(1))&"</option>"
			else
				response.Write "<option value="&rs(0)&">"&oblog.filt_html(rs(1))&"</option>"
			end if
			rs.movenext
		wend
		%>
			</select>
			(ѡ�����·������ĸ�BLOG)�����Ŷ��û�ר�⣺ 
			<select name="blogteamsubject" id="s2">
		<%
		if oblog.logined_uid<>userid and userid>0 then
			set rs=oblog.execute("select subjectid,subjectname from oblog_subject where userid="&userid)
			response.Write("<option value=0>�ҵ�ר��</option>")
			while not rs.eof
				if subjectid=rs(0) then
					response.Write("<option value="&rs(0)&" selected>"&rs(1)&"</option>")
				else
					response.Write("<option value="&rs(0)&" >"&rs(1)&"</option>")
				end if
				rs.movenext
			wend
		else
			response.Write("<option value=0 selected>�ҵ�ר��</option>")
		end if
		set rs=nothing
		%>                
			</select> 
			</ul>	
		<%If t<>"1" And log_type<>"1" Then%>		
			<ul class="list_edit"> 
			����ͨ��(֧�ֶ������ͨ�棬ÿһ��Ϊһ��)�� (<a href="user_help.asp#tb" target="_blank">ʲô������ͨ��?</a>)</ul>
			<ul class="list_edit" style="margin-left:15px">
			<textarea name="tb" type="text" id="tb" size="60" rows=3 cols=90><%=oblog.filt_html(tburl)%></textarea>		
			</ul>
		<%Else%>
			<ul class="list_edit" style="display:none;"> 		
			ͼƬ��ַ(ÿ��Ϊһ��ͼƬ,ͼƬ���ڲ��������Զ���ʾ)��<BR/>
			<!--<input type="hidden" id="picedit" name="picedit" value=""/>			-->
			<textarea name="log_pics" type="text" id="log_pics" size="60" rows=5 cols=90><%=log_pics%></textarea>		
		</ul>
		<%End If%>
	<ul class="list_edit"> 
		���÷���ʱ�䣺<%show_selectdate(addtime)%> 
		<input name="logid"  type="hidden" id="logid" value=<%=logid%> >
		<input name="oldisdraft"  type="hidden" id="oldisdraft" value=<%=oldisdraft%> >
		<BR/><BR/>
	</ul>
	</div>
	</form>
</div>
<%
end sub

sub savelog()
	dim blog,logtext,i,rs,logid,isdraft,p,tid,log_tags,filename,log_pics,log_Abstract
	dim log_topic,log_text,log_face,log_time,log_classid,log_showword,log_blogteam,log_teamsubject,log_subjectid,log_password,log_ishide,log_istop,log_isencomment,log_isdraft,log_modiid,log_tb,log_filename,todraft,log_str,log_oldtb
	if oblog.ChkPost()=false then
		oblog.adderrstr("ϵͳ��������ⲿ�ύ��")
		oblog.showusererr
		exit sub
	end if
	if oblog.setup(60,0)=1 then
		if not oblog.codepass then oblog.adderrstr("��֤�������ˢ�º��������룡"):oblog.showusererr
	end if
	if Cint(oblog.setup(48,0))=1 and Cint(Request("classid"))=0 then
		oblog.adderrstr("��ѡ��"& tName &"�����"):oblog.showusererr
	end if
	
	'if Session("postchknum") <> cstr(request("postchknum")) or Session("postchknum") = Empty Then
	if request.Cookies(cookies_name)("postchknum") <> cstr(request("postchknum")) or request.Cookies(cookies_name)("postchknum") = Empty Then
		oblog.adderrstr("�������ظ��ύ����")
		oblog.showusererr
		exit sub
	End If
	'session("postchknum")=cstr(Int(900*rnd)+1000)
	Randomize
	if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
	'Response.Cookies(cookies_name)("postchknum") = cstr(Int(900*rnd)+1000)
	response.Write("<div id=""prompt""><ul>")
	logid=clng(request.Form("logid"))
	isdraft=clng(request.Form("isdraft"))
	if isdraft=1 then p=1 else p=10
	log_topic= Request.Form("topic")
	log_face= Request.Form("face")
	for i = 1 to Request.Form("logtext").Count 
  		log_text = log_text & Request.Form("logtext")(i)
	Next
	if log_text="" then log_text=trim(request.Form("edit"))
	log_time=request("selecty")&"-"&request("selectm")&"-"&request("selectd")&" "&request("selecth")&":"&request("selectmi")&":"&request("selects")
	log_classid=trim(request("classid"))
	log_showword=trim(request("showword"))
	log_blogteam=clng(trim(request("blogteam")))
	log_teamsubject=trim(request("blogteamsubject"))
	log_subjectid=trim(request("subjectid"))
	log_password=trim(request("ispassword"))
	log_ishide=trim(request("ishide"))
	log_isencomment=trim(request("isencomment"))
	log_istop=trim(request("istop"))
	log_tb=trim(request("tb"))
	log_filename=trim(request("filename"))
	log_isdraft=isdraft
	log_pics=trim(request("log_pics"))
	'log_str=oblog.logined_uname&"|||"&log_topic&"|||"&oblog.InterceptStr(RemoveHTML(log_text),500)
	if logid>0 then 
		log_modiid=logid
	end if
	if log_topic="" or oblog.strLength(log_topic)>100 then oblog.adderrstr( tName & "���ⲻ��Ϊ��(���ܴ���100)��")
	if trim(log_filename)="�Զ����" then log_filename=""
	if oblog.chkdomain(log_filename)=false and log_filename<>"" then oblog.adderrstr("�ļ����Ʋ��Ϲ淶��ֻ��ʹ��Сд��ĸ�����ּ��»��ߣ�")
	if log_text="" or  oblog.strLength(log_topic)>oblog.setup(75,0) then oblog.adderrstr( tName & "���ݲ���Ϊ���Ҳ��ܴ���"&oblog.setup(75,0)&"�ַ���")
	if oblog.chk_badword(log_text)>0 then oblog.adderrstr("�����к���ϵͳ���������Ĺؼ��֣�")
	if not isdate(log_time) then oblog.adderrstr( tName & "ʱ���ʽ����")
	if log_showword="" then log_showword=0
	if not isnumeric(log_showword) then oblog.adderrstr( tName & "������ʾ��������Ϊ���֣�")
	if log_teamsubject="" then log_teamsubject=0
	if log_subjectid="" then log_subjectid=0
	if log_classid="" then log_classid=0
	if log_istop="" then log_istop=0
	if log_isencomment="" then log_isencomment=0
	if log_ishide="" then log_ishide=0
	if log_isdraft="" then log_isdraft=0
	if clng(request("oldisdraft"))=0 and log_isdraft=1 and log_modiid>0 then todraft=1
	if clng(request("oldisdraft"))=1 and log_isdraft=0 and log_modiid>0 then todraft=-1
	If EN_TAGS=1 Then
		log_tags=Trim(Request.Form("logtags"))
		If log_tags<>"" Then
			log_tags=Replace(log_tags,"'","")
			If Len(log_tags)>255 Then
				oblog.adderrstr("TAG�ܳ��Ȳ��ܴ���255���ַ�")
			End If
			If Ubound(Split(log_tags,P_TAGS_SPLIT))>(P_TAGS_MAX-1) Then
				oblog.adderrstr("ÿƪ�������֧��" & P_TAGS_MAX & "��TAG")
			End If			
		End If
	End If
	if oblog.errstr<>"" then oblog.showusererr
	'ȡ������ժҪ����
'	If log_password<>"" Then
'		if log_password_old=log_password then
'			log_Abstract="<form method=""post"" action="""&blogdir&"more.asp?id="&tid&""" target=""_blank"">������" & tName & "�������룺<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""�ύ""></form>"
'		else
'			if log_password<>"" then
'				log_Abstract="<form method=""post"" action="""&blogdir&"more.asp?id="&tid&""" target=""_blank"">������" & tName & "�������룺<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""�ύ""></form>"
'			else
'				log_Abstract=""
'			end if
'		end if
'	End If	
	'If log_ishide=1 then log_Abstract="��" & tName & "Ϊ����" & tName & "�������ѿɼ���<a href=''"&blogdir&"more.asp?id="&tid&"''>���������֤ҳ��</a>��"
	if log_Abstract="" then	
		log_Abstract=trimlog(log_text,clng(trim(log_showword)))
		If request("isubb")="1" Then
			log_Abstract=UBBCode(log_Abstract,1)
			log_Abstract=replace(log_Abstract,CHR(10),"<br /> ")
		End If
		log_Abstract=replace(log_Abstract,"#isubb#","")
		log_Abstract=filtimg(log_Abstract)
		If oblog.setup(29,0)=1 Then	log_Abstract=profilthtm(log_Abstract)
	End if
	'����ժҪ
	set blog=new class_blog
	set rs=server.createobject("adodb.recordset")
	if log_modiid>0 then
		rs.open "select * from oblog_logCooperateSubmit where logid="&log_modiid,conn,2,2
	else
		rs.open "select top 1 * from oblog_logCooperateSubmit",conn,2,2
		rs.addnew
	end if
	'��ʼд�����
	rs("topic")=EncodeJP(oblog.filt_astr(log_topic,250))
	log_text=replace(log_text,"#isubb#","")
	if request("isubb")="1" then 
		log_text="#isubb#"&log_text
		rs("EditorType")=1
	Else
		rs("EditorType")=0
	End If
	log_text = EncodeJP(oblog.filtpath(oblog.filt_badword(log_text)))
	rs("logtext")=log_text
	rs("face")=log_face
	rs("addtime")=log_time
	rs("classid")=log_classid
	if log_teamsubject>0 then log_subjectid=clng(log_teamsubject)
	if rs("subjectid")<>clng(log_subjectid) and log_modiid>0 then
		oblog.execute("update oblog_subject set subjectlognum=subjectlognum+1 where subjectid="&clng(log_subjectid))
		oblog.execute("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid="&clng(rs("subjectid")))
	end if
	rs("subjectid")=clng(log_subjectid)
	rs("showword")=clng(trim(log_showword))
	if log_modiid=0 then
		rs("authorid")=oblog.logined_uid
		rs("author")=EncodeJP(oblog.logined_uname)
	end if
	rs("userid")=log_blogteam
	rs("ishide")=log_ishide
	rs("istop")=log_istop
	log_oldtb=rs("tburl")
	rs("tburl")=log_tb
	rs("logpics")=log_pics
	rs("logtype")=t
	rs("isencomment")=log_isencomment
	rs("Abstract")=log_Abstract
	if rs("ispassword")=log_password then
	else
		if log_password<>"" then
			rs("ispassword")=md5(trim(log_password))
		else
			log_password=""
			rs("ispassword")=""
		end if
	end if
	if oblog.setup(9,0)=1 then
		rs("passcheck")=0
		log_Abstract="��" & tName & "��Ҫ����Ա��˺�ſɼ���"
	else
		rs("passcheck")=1			
	end if
	rs("isdraft")=log_isdraft
	rs("filename")=log_filename
	if log_modiid=0 then
		rs("iis")=0
		rs("commentnum")=0
		rs("trackbacknum")=0
		rs("blog_password")=0
		rs("truetime")=now()
	end if
	rs.update
	rs.close
	Response.Cookies(cookies_name)("postchknum") = cstr(Int(900*rnd)+1000)
	'''blog.progress_init
	'''blog.progress int(1/p*100),"���浽���ݿ����"
	if (log_modiid=0 and log_isdraft=0) or todraft=-1 then
		oblog.execute("update [oblog_user] set log_count=log_count+1 where userid="&log_blogteam)
		if log_classid>0 then
			'''oblog.execute("update [oblog_logclass] set classlognum=classlognum+1 where id="&log_classid)
		end if
		if log_subjectid>0 then
			oblog.execute("update [oblog_subject] set subjectlognum=subjectlognum+1 where subjectid="&log_subjectid)
		end if
		oblog.execute("update [oblog_setup] set log_count=log_count+1")
		
	end if
	if log_modiid=0 then
		set rs=oblog.execute("select max(logid) from oblog_logCooperateSubmit where userid="&log_blogteam)
		tid=rs(0)
	else
		tid=log_modiid
	end if
	'-------------------------------------------
	'ATFLAG ����Tag����
	if log_modiid=0 then
		set rs=oblog.execute("select max(logid) from oblog_logCooperateSubmit where userid="&log_blogteam)
		tid=rs(0)
		If EN_TAGS=1 Then Call Tags_UserAdd(log_tags, oblog.logined_uid, tid)
	else
		tid=log_modiid
		If EN_TAGS=1 Then Call Tags_UserEdit(log_tags, oblog.logined_uid, tid)
	end if
	'-------------------------------------------
	'��Ƭֱ�ӷ�����������ݸ��书�ܣ�������Ҫ���������ҳ
	if isdraft=0  then
		blog.userid=log_blogteam
'''		blog.progress int(2/p*100),"���ɾ�̬�����ļ�..."
'''		blog.CreateFunctionPage
'''		blog.update_log tid,1
''//		if log_modiid=0 then
''//			set rs=oblog.execute("select top 1 logid from oblog_logCooperateSubmit where logid<"&tid&" and userid="&log_blogteam&" and logtype="&cint(t)&" order by addtime desc")
''//			if not rs.eof then blog.update_log rs(0),0
''//		end if
''//		blog.progress int(3/p*100),"���ɾ�̬�����ļ�..."
''//		blog.update_calendar(tid)
''//		blog.progress int(4/p*100),"���ɾ�̬�������б��ļ�..."
''//		blog.update_newblog(log_blogteam)
''//		blog.progress int(5/p*100),"������ҳ���·����ļ�..."
''//		blog.update_subject(log_blogteam)
''//		blog.progress int(6/p*100),"���ɾ�̬��ҳ�ļ�..."
''//		if t<>"1" then	blog.update_index 1
''//		blog.progress int(7/p*100),"����ͼƬ����..."
''//		if t="1" then Call UpdatePicLinks(log_pics,tid,log_blogteam,log_classid,log_subjectid,log_password,log_ishide)
''//		blog.progress int(8/p*100),"����վ����Ϣ..."
''//		blog.update_info log_blogteam
''//		blog.progress int(9/p*100),"�ύ����ͨ��..."
''//		'��Ŀ�����ӷ���Pingָ��
		if log_tb<>"" and log_tb<>log_oldtb then
			Dim objTrackBack
			Set objTrackBack=New Class_TrackBack
			objTrackBack.Logid=tid
			objTrackBack.Blog_Name=blog.BlogName
			objTrackBack.Title=log_topic
			objTrackBack.URL=oblog.setup(3,0)  & "go.asp?logid=" &tid
			objTrackBack.Excerpt=log_topic & "<br />oBlog Created"
			Call objTrackBack.ProcessMultiPing(log_tb)
			Set objTrackBack =Nothing
			'response.Write("<script src="""&log_tb&"&url="&trim(oblog.setup(3,0))&blog.gourl&"&topic="&oblog.filt_astr(unHtml(log_topic),250)&"&tbuser="&oblog.logined_uname&"""><script>")
		end if
'''		blog.progress int(10/p*100),"�ύ���!"
	else
		if todraft=1 then
			logtodraft(tid)
		end if
		response.Write("<li><a href=""HomestayBackctrl-PostSubmitCooperate.asp?logid="&tid&""">�Ѿ����浽�ݸ��䣬��������޸�����</a></li>")
	end if
	'�������޸����º��������ɹ���ҳ��
	set rs=nothing
	set blog=nothing
	response.Write("</ul></div>")
	oblog.showok "�ύ�ɹ������ǽ���24Сʱ��Ϊ��������","HomestayBackctrl-PostSubmitManageCooperate.asp"
end sub

sub getteam()
	dim s,i,s1,rs,rs1
	set rs=server.createobject("adodb.recordset")
	rs.open "select oblog_blogteam.mainuserid,[oblog_user].blogname from oblog_blogteam,[oblog_user] where oblog_blogteam.otheruserid="&oblog.logined_uid&" and oblog_blogteam.mainuserid=[oblog_user].userid",conn,1,1
	if not rs.eof then
		response.write "<script language=""JavaScript"">"&vbcrlf
		s = "var p_array = new Array(" + cstr(rs.recordcount-1) + ");"&vbcrlf
		response.write s
		s = "var p_array_id = new Array(" + cstr(rs.recordcount-1) + ");"&vbcrlf
		response.write s
		i = 0
		while not rs.eof
			set rs1=server.createobject("adodb.recordset")
			rs1.open "select subjectid,subjectname from oblog_subject where userid="&rs("mainuserid"),conn,1,1
			s = "var p"+cstr(rs("mainuserid"))+"_array = Array("
			s1 = "var p"+cstr(rs("mainuserid"))+"_array_id = Array("
			if rs1.recordcount > 0 then
			while not rs1.eof
				if trim(rs1("subjectname"))<>"" then
					s = s + """" + oblog.filt_html(rs1("subjectname")) + """" 
					s1 = s1 + """" + cstr(rs1("subjectid")) + """"
					s = s + ","
					s1 = s1 + ","
				end if
				rs1.movenext
			wend
			s = s + """" + "��ѡ��ר��" + """" 
			s1 = s1 + """" + cstr(0) + """"
			else
				s = s + """" + "�޿���ר��" + """" 
				s1 = s1 + """" + cstr(0) + """"
			end if
			s = s+ ");"&vbcrlf
			s1 = s1+ ");"&vbcrlf
			response.write s
			response.write s1 
			response.write "p_array["+cstr(i)+"] = p"+cstr(rs("mainuserid"))+"_array;"&vbcrlf
			response.write "p_array_id["+cstr(i)+"] = p"+cstr(rs("mainuserid"))+"_array_id;"&vbcrlf
			i = i + 1
			rs.movenext
		wend
		response.write  "</script>"&vbcrlf
		rs.close
		set rs=nothing
		rs1.close
		set rs1=nothing
	end if
end sub

sub logtodraft(logid)
	logid=clng(logid)
	dim uid,delname,subjectfile,sdate,edate,fso,sid,rs,readme
	set rs=server.createobject("adodb.recordset")
	rs.open "select userid,logfile,issave,subjectfile,subjectid,isdraft from oblog_logCooperateSubmit where logid="&logid,conn,1,3
	if not rs.eof then
		uid=rs(0)
		delname=trim(rs(1))
		subjectfile=rs(3)
		sid=rs(4)
		if true_domain=1 then
			if instr(delname,"archives") then
				delname=right(delname,len(delname)-InstrRev(delname,"archives")+1)
			else
				delname=right(delname,len(delname)-InstrRev(delname,"/"))
			end if
			if oblog.logined_ulevel=9 then
				set rst1=oblog.execute("select user_dir,user_folder from oblog_user where userid="&clng(uid))
				if not rst1.eof then
					delname=rst1(0)&"/"&rst1(1)&"/"&delname
				end if
				set rst1=nothing
			else
				delname=oblog.logined_udir&"/"&oblog.logined_ufolder&"/"&delname
			end if
		end if
		if delname<>"" then 
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			if fso.FileExists(Server.MapPath(delname)) then	fso.deleteFile Server.MapPath(delname)
		end if
		rs("logfile")=""
		rs("isdraft")=1
		rs.update
		rs.close
		oblog.execute("update oblog_user set log_count=log_count-1 where userid="&uid)
		oblog.execute("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid="&clng(sid))
		dim blog
		set blog=new class_blog
		blog.userid=uid
		'blog.update_index_subject 0,0,0,""			
		blog.update_index 0
		blog.update_newblog(uid)
		set blog=nothing
		set fso=nothing	
		set rs=nothing
	else
		rs.close
		set rs=nothing
		exit sub
	end if	
end sub

'����ͼƬ���������µĹ���
Private Function UpdatePicLinks(byval sPics,byval sLogid,byval sUserId,byval SysClassId,byval UserClassId,byval log_password,byval log_ishide)
	Dim aPics,i,sPic,ispw
	sPics=Trim(sPics)
	log_password=trim(log_password)
	if (log_password="" or isnull(log_password)) and cint(log_ishide)=0 then
		ispw=0
	else
		ispw=1
	end if
	aPics=Split(sPics,VBCRLF)
	For i=0 To UBound(aPics)
		sPic = Trim(aPics(i))
		If sPic<>"" Then		
			Select Case LCase(Right(sPic,3))
				Case "jpg","gif","png","bmp","tif"
					Call oblog.Execute("Update oblog_upfile Set logid=" & sLogid & ",sysClassId=" &SysClassId &", UserClassId='" & UserClassId &"',isphoto=1,isPower="&ispw&" Where file_path='" & sPic & "' And userid=" & sUserId)
				Case Else
			End Select
		End If			
	Next
End Function

Private Function SendTrack(byval sTrackBacks,byval sTopics)
	Dim aTBs,i,sTB
	sTrackBacks=Trim("sTrackBacks")
	aTBs=Split(sTrackBacks,VBCRLF)
	For i=0 To UBound(aTBs)
		sTB = aTBs(i)
		If sTB<>"" Then
			'Send TrackBack			
		End If			
	Next
End Function

sub show_selectdate(addtime)	              
	dim y,m,d,h,mi,s,ttime
	if addtime="" then ttime=ServerDate(now()) else ttime=addtime
	response.Write("<select name=selecty id=selecty>")
	for  y=2000 to 2010
		if year(ttime)=y then
			response.Write "<option value="&y&" selected>"&y&"��</option>"
		else
			response.Write "<option value="&y&">"&y&"��</option>"
		end if
	next
	response.Write "</select><select name=selectm id=selectm >"
	for m=1 to 12
	  	if month(ttime)=m then
		  	response.Write "<option value="&m&" selected>"&m&"��</option>"
		else
		  	response.Write "<option value="&m&">"&m&"��</option>"
		end if
	next
	response.Write("</select><select name=selectd id=selectd >")
	for d=1 to 31
		if day(ttime)=d then
		  	response.Write "<option value="&d&" selected>"&d&"��</option>"
		else
		  	response.Write "<option value="&d&">"&d&"��</option>"
		end if
	next
	response.Write("</select><select name=selecth id=selecth>")             
	for h=0 to 23
	  	if hour(ttime)=h then
		  	response.Write "<option value="&h&" selected>"&h&"ʱ</option>"
		else
		  	response.Write "<option value="&h&">"&h&"ʱ</option>"
		end if
	next
	response.Write("</select><select name=selectmi id=selectmi>")       
	for mi=0 to 59
	  	if minute(ttime)=mi then
		  	response.Write "<option value="&mi&" selected>"&mi&"��</option>"
		else
		  	response.Write "<option value="&mi&">"&mi&"��</option>"
		end if
	next
	response.Write("</select><select name=selects id=selects>")
	for s=0 to 59
	  	if second(ttime)=s then
		  	response.Write "<option value="&s&" selected>"&s&"��</option>"
		else
		  	response.Write "<option value="&s&">"&s&"��</option>"
		end if
	next
	response.Write("</select>")   
end sub
%>