<!--#include file="user_topHome.asp"-->
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
				<span>客户咨询热线</span>：<span>010-85493388</span>/<span>13146398085</span> &nbsp;&nbsp;&nbsp;
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
					
					
					
						<center><h2>修改我的申请接待家庭资料</h2></center>
						
			


<%const MaxPerPage=24
dim strFileName,totalPut,TotalPages
dim rs, sql,action
dim id,usersearch,Keyword,strField
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
usersearch=trim(request("usersearch"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
if usersearch="" then
	usersearch=0
else
	usersearch=Clng(usersearch)
end if
strFileName="HomestayBackctrl-InfoHome.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
select case action
	case "modifyphoto"
	call modify()
	case "savemodify"
	call Savemodify()
	case "delfile" 
	call delfile()
	
	case "chkreg"
	call sub_chkreg()'处理修改家庭个人资料;
	case else
	call main()
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
								家庭用户中心
							</p>
							
							<!--#include file="menu/inc-menuHome.asp"-->
							
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
	dim strGuide,ssql
	ssql="*"
	strGuide="<h1 style='font-size:13px;'>"
	select case usersearch
		case 0
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_user] where userid="& oblog.logined_uid &" "
			else
				sql="select "&ssql&" from [oblog_user] where userid="&oblog.logined_uid&" "
			end if
			strGuide=strGuide & "您当前的申请类型是："
		case 1
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' ) order by fileid desc"
			end if
			strGuide=strGuide & "图片文件"
		case 2
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit') order by fileid desc"
			end if
			strGuide=strGuide & "压缩文件"
		case 3
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='doc' or file_ext='xsl' or file_ext='txt' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='doc' or file_ext='xsl' or file_ext='txt') order by fileid desc"
			end if
			strGuide=strGuide & "文档文件"
		case 4
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where  file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf') order by fileid desc"
			end if
			strGuide=strGuide & "媒体文件"
		case 5
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where isphoto=1 order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and isphoto=1 order by fileid desc"
			end if
			strGuide=strGuide & "相册"
		case else
	end select
	'strGuide=strGuide & "</td><td align='right'>"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (共有0个文件)</h1><div id=""list"">"
		response.write strGuide
		showContent
	else
	    'response.write sql
        showContent
        
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()


	
	dim str_usertype
	Dim show_reg
	
	'show_reg=show_reg&"<div  style='font-size:12px;' CLASS='Register'>"
	Dim FamilyTypeRequest,FamilyType,FamilyType_n,i
	FamilyType = rs("FamilyType")
	
	Select case rs("FamilyType")
	Case "1"
	FamilyTypeRequest= "免费接待家庭"
	Case "2"
	FamilyTypeRequest= "收费接待家庭"
	Case "3"
	FamilyTypeRequest= "2008奥运接待家庭"
	Case "1, 2"
	FamilyTypeRequest= "免费/收费接待家庭"
	Case "1, 3"
	FamilyTypeRequest= "免费接待家庭/2008奥运接待家庭"
	Case "1, 2, 3"
	FamilyTypeRequest= "免费/收费/2008奥运接待家庭"
	Case "2, 3"
	FamilyTypeRequest= "收费接待家庭/2008奥运接待家庭"
    End Select
	'FamilyType_n = Split(FamilyType, ",")
        'For i = 0 To UBound(FamilyType_n)
            'response.Write FamilyType_n(i)
        'Next
	'Response.Write(rs("FamilyType"))

	show_reg=show_reg&"<font color=#038ad7>您当前申请的是 :</font></center>"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"当前系统已关闭注册。"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",rs("user_classid"),0)
    str_usertype=str_usertype&"</select><font color=#038ad7> *</font>"
	show_reg=show_reg&"<form name=oblogform id=oblogform method=post action=HomestayBackctrl-InfoHome.asp onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    
	
	show_reg=show_reg&"<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	  If instr(FamilyType,"1")>0 Then
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='1' checked='checked'>"
	  Else
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='1'>"
	  End If	  
	  show_reg=show_reg&"申请免费接待家庭 "
	  
	  If instr(FamilyType,"2")>0 Then
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2' checked='checked'>"
	  Else
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2'>"
	  End If
	  show_reg=show_reg&"申请收费接待家庭 "
	  
	  If instr(FamilyType,"3")>0 Then
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3' checked='checked'>"
	  Else
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3'>"
	  End If
	  show_reg=show_reg&"<font color='#038ad7'>申请2008奥运接待家庭</font>"
	  
	  
	  show_reg=show_reg&"<br />"
	  
	  
	  show_reg=show_reg&"<br />"
	

	
	
	
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>家庭帐号信息</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>家庭登录帐号：</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=63% colspan=2>"&  rs("username")& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=20% height=25><div align=left >域名：</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
'    show_reg=show_reg&"<td width=63% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>查看以上帐号是否被占用</a>"& vbcrlf 
'    show_reg=show_reg&"</td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
''    show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>输入密码：</div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=password type=password id=password size=15 maxlength=30><font color=#038ad7>如不修改，请保留为空</font></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
''	
''    show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>登录密码确认：</div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=repassword type=password id=repassword size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>密码提示问题：</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>密码提示答案：</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>家庭基本信息</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>所在地区(省/市)：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_city(rs("province"),rs("city")) &"<font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>Email地址：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=35 maxlength=30 value='"& rs("useremail")& "'><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>真实姓名：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30 value='"& rs("blogname") &"'><font color=#038ad7> *</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您的称呼：<select name=sex>"
		If Cint(rs("sex"))=3 Then
		show_reg=show_reg&"<option value=3 selected='selected'>请选择</option>"
		Else
		show_reg=show_reg&"<option value=3>请选择</option>"
		End If
		If Cint(rs("sex"))=0 Then
		show_reg=show_reg&"<option value=0 selected='selected'>女士</option>"
		Else
		show_reg=show_reg&"<option value=0>女士</option>"
		End If
		If Cint(rs("sex"))=1 Then
		show_reg=show_reg&"<option value=1 selected='selected'>先生</option>"
		Else
		show_reg=show_reg&"<option value=1>先生</option>"
		End If
		
	show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>请选择学习范围：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>出生日期：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"

	show_reg=show_reg& oblog.type_dateselect(rs("birthday"),1) &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>您的职业：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_job2(rs("job")) &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>联系方式：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>电话：<input name=tel value='"& rs("tel") &"' size='15' maxlength='50' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;手机：<input name=mobile  value='"& rs("mobile") &"' size='15' maxlength='50' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>地址：<input name='address' value='"& rs("address") &"' size='30' maxlength='250' />&nbsp;&nbsp;&nbsp;邮编：<input name=postnum value='"& rs("postnum") &"' size='8' maxlength='6' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>家庭成员：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>家庭人数：<input id='familynumber_0' type='radio' name='familynumber' value='1' "
	If Instr(rs("familynumber"),"1")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_0'>1</label>&nbsp;&nbsp;<input id='familynumber_1' type='radio' name='familynumber' value='2' "
	If Instr(rs("familynumber"),"2")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_1'>2</label>&nbsp;&nbsp;<input id='familynumber_2' type='radio' name='familynumber' value='3' "
	If Instr(rs("familynumber"),"3")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_2'>3</label>&nbsp;&nbsp;<input id='familynumber_3' type='radio' name='familynumber' value='4' "
	If Instr(rs("familynumber"),"4")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_3'>4</label>&nbsp;&nbsp;<input id='familynumber_4' type='radio' name='familynumber' value='更多' "
	If Instr(rs("familynumber"),"更多")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_4'>更多</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>家庭结构：<input name='familymember' type='checkbox' id='Father' value='父亲' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"父亲")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Father'>父亲</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Mother' value='母亲' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"母亲")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Mother'>母亲</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Husband' value='丈夫' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"丈夫")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Husband'>丈夫</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Wife' value='妻子' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"妻子")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Wife'>妻子</label><br>&nbsp;&nbsp;<font color=#ffffff>Homestay</font><input name='familymember' type='checkbox' id='Son' value='儿子' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"儿子")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Son'>儿子</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Daughter' value='女儿' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"女儿")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Daughter'>女儿</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Brother' value='兄弟' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"兄弟")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Brother'>兄弟</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Sister' value='姐妹' title='请选择您的家庭结构，可多选' "
	If Instr(rs("familymember"),"姐妹")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Sister'>姐妹</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
                           
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>房屋介绍：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>结构：<select name='houseframe' id='houseframe1' title='请选择您的房屋结构'>"
                            show_reg=show_reg&"<option value='1室' "
							If Instr(rs("houseframe"),"1室")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1室</option>"
                            show_reg=show_reg&"<option value='2室' "
							If Instr(rs("houseframe"),"2室")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2室</option>"
                            show_reg=show_reg&"<option value='3室' "
							If Instr(rs("houseframe"),"3室")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3室</option>"
                            show_reg=show_reg&"<option value='4室' "
							If Instr(rs("houseframe"),"4室")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4室</option>"
                            show_reg=show_reg&"<option value='别墅' "
							If Instr(rs("houseframe"),"别墅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">别墅</option>"
                            show_reg=show_reg&"<option value='其他' "
							If Instr(rs("houseframe"),"其他")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">其他</option>"
        show_reg=show_reg&"</select>&nbsp;<select name='houseframe' id='houseframe2'  title='请选择您的房屋结构'>"&_
                            "<option value='1厅' "
							If Instr(rs("houseframe"),"1厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1厅</option>"
                            show_reg=show_reg&"<option value='2厅' "
							If Instr(rs("houseframe"),"2厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2厅</option>"
                            show_reg=show_reg&"<option value='3厅' "
							If Instr(rs("houseframe"),"3厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3厅</option>"
                            show_reg=show_reg&"<option value='4厅' "
							If Instr(rs("houseframe"),"4厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4厅</option>"
                            show_reg=show_reg&"<option value='5厅' "
							If Instr(rs("houseframe"),"5厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">5厅</option>"
                            show_reg=show_reg&"<option value='6厅' "
							If Instr(rs("houseframe"),"6厅")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">6厅</option>"
        show_reg=show_reg&"</select>&nbsp;<select name='houseframe' id='houseframe3' title='请选择您的房屋结构'>"
                            show_reg=show_reg&"<option value='1卫' "
							If Instr(rs("houseframe"),"1卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1卫</option>"
                            show_reg=show_reg&"<option value='2卫' "
							If Instr(rs("houseframe"),"2卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2卫</option>"
                            show_reg=show_reg&"<option value='3卫' "
							If Instr(rs("houseframe"),"3卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3卫</option>"
                            show_reg=show_reg&"<option value='4卫' "
							If Instr(rs("houseframe"),"4卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4卫</option>"
                            show_reg=show_reg&"<option value='5卫' "
							If Instr(rs("houseframe"),"5卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">5卫</option>"
                            show_reg=show_reg&"<option value='6卫' "
							If Instr(rs("houseframe"),"6卫")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">6卫</option>"
                          show_reg=show_reg&"</select>"&_
						  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;面积："&_ 
                          "<select name='housespace' id='housespace' style='width:100px' title='请选择您的房屋面积'>"
                            show_reg=show_reg&"<option>请选择</option>"
                            show_reg=show_reg&"<option value='50-100' "
							If Instr(rs("housespace"),"50-100")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">50到100平方米</option>"
                            show_reg=show_reg&"<option value='100-150' "
							If Instr(rs("housespace"),"100-150")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">100到150平方米</option>"
                            show_reg=show_reg&"<option value='150-200' "
							If Instr(rs("housespace"),"150-200")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">150到200平方米</option>"
                            show_reg=show_reg&"<option value='above 200' "
							If Instr(rs("housespace"),"above 200")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">200平方米以上</option>"
                          show_reg=show_reg&"</select>"&_
						  "</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>是否有宠物：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value='有' name='pet' id='Pet1' onClick='ShowPet(1);' title='请选择您的宠物种类' "
	If Instr(rs("pet"),"没有")=0 Then show_reg=show_reg&"checked"
	show_reg=show_reg&"><label for='Pet1'>有</label>&nbsp;<INPUT name=pet type=radio value='没有' id='Pet2' onClick='ShowPet(0);' "
	If Instr(rs("pet"),"没有")>0 Then show_reg=show_reg&"checked"
	show_reg=show_reg&" ><label for='Pet2'>没有</label>&nbsp;&nbsp;&nbsp;<span width='21%' id='PetTypes'>"
	
	If Instr(rs("pet"),"没有")=0 Then
		show_reg=show_reg&"<select id='PetType' name='PetType'><option value=''>请选择</option><option value='狗' "
		If Instr(rs("PetType"),"狗")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">狗</option><option value='猫' "
		If Instr(rs("PetType"),"猫")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">猫</option><option value='鸟' "
		If Instr(rs("PetType"),"鸟")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">鸟</option><option value='其他' "
		If Instr(rs("PetType"),"其他")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">其他</option></select>"
	End If
	show_reg=show_reg&"</span></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                      
    
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>对外教的特殊要求：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>外教性别：<input id='asksex_0' type='radio' name='asksex' value='男' "
	If Instr(rs("asksex"),"男")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_0'>男</label>&nbsp;&nbsp;&nbsp;<input id='asksex_1' type='radio' name='asksex' value='女' "
	If Instr(rs("asksex"),"女")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_1'>女</label>&nbsp;&nbsp;&nbsp;<input id='asksex_2' type='radio' name='asksex' value='无所谓' "
	If Instr(rs("asksex"),"无所谓")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_2'>无所谓</label> &nbsp;&nbsp;&nbsp;</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>是否会说中文：<input id='issaychinese_0' type='radio' name='issaychinese' value='是' "
	If Instr(rs("issaychinese"),"是")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_0'>是</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_1' type='radio' name='issaychinese' value='否' "
	If Instr(rs("issaychinese"),"否")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_1'>否</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_2' type='radio' name='issaychinese' value='无所谓' "
	If Instr(rs("issaychinese"),"无所谓")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_2'>无所谓</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>家庭附加信息</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	                    
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>如何知道我们：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input id='howknowsite_0' type='checkbox' name='howknowsite' value='报纸杂志' "
	If Instr(rs("howknowsite"),"报纸杂志")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_0'>报纸杂志</label>&nbsp;&nbsp;<input id='howknowsite_1' type='checkbox' name='howknowsite' value='朋友介绍' "
	If Instr(rs("howknowsite"),"朋友介绍")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_1'>朋友介绍</label>&nbsp;&nbsp;<input id='howknowsite_2' type='checkbox' name='howknowsite' value='网站' "
	If Instr(rs("howknowsite"),"网站")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_2'>网站</label>&nbsp;&nbsp;<input id='howknowsite_3' type='checkbox' name='howknowsite' value='论坛' "
	If Instr(rs("howknowsite"),"论坛")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_3'>论坛</label>&nbsp;&nbsp;<input id='howknowsite_4' type='checkbox' name='howknowsite' value='搜索引擎' "
	If Instr(rs("howknowsite"),"搜索引擎")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_4'>搜索引擎</label>&nbsp;&nbsp;<input id='howknowsite_5' type='checkbox' name='howknowsite' value='其它' "
	If Instr(rs("howknowsite"),"其它")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_5'>其它</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>英语水平：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name='englishlevel' id='englishlevel_0' type='radio' value='不会' title='请选择您英语水平' "
	If Instr(rs("englishlevel"),"不会")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_0'>不会</label>&nbsp;&nbsp;<input name='englishlevel' id='englishlevel_1' type='radio' value='一般' checked title='请选择您英语水平' "
	If Instr(rs("englishlevel"),"一般")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_1'>一般</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_2' value='中等' title='请选择您英语水平' "
	If Instr(rs("englishlevel"),"中等")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_2'>中等</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_3' value='很好' title='请选择您英语水平' "
	If Instr(rs("englishlevel"),"很好")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_3'>很好</label>&nbsp;&nbsp;&nbsp;&nbsp;<strong>是否有私家车</strong>：<INPUT name=car id=car_0 type=radio value=有 title='请选择是否有私家车' "
	If Instr(rs("car"),"没有")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='car_0'>有</label>&nbsp;<INPUT name=car id=car_1 type=radio value=没有 title='请选择是否有私家车' "
	If Instr(rs("car"),"没有")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='car_1'>没有</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>卫生间的类型：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT name=toilet id=toilet_0 type=radio value=共用 title='请选择是否有独立的卫生间' "
	If Instr(rs("toilet"),"共用")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='toilet_0'>共用</label><INPUT type=radio value=房间独立 name=toilet id=toilet_1 title='请选择是否有独立的卫生间' "
	If Instr(rs("toilet"),"房间独立")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='toilet_1'>房间独立</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>有否电脑：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value=有 name=computer id=computer_0 title='请选择是否有电脑' "
	If Instr(rs("computer"),"没有")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='computer_0'>有</label>&nbsp;&nbsp;<INPUT name=computer id=computer_1 type=radio value=没有 title='请选择是否有电脑' "
	If Instr(rs("computer"),"没有")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='computer_1'>没有</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>有否宽带上网：</strong><INPUT type=radio value=有 name=internet id=internet_0 title='请选择是否有宽带可上网' "
	If Instr(rs("internet"),"没有")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='internet_0'>有</label>&nbsp;&nbsp;<INPUT name=internet id=internet_1 type=radio value=没有 title='请选择是否有宽带可上网' "
	If Instr(rs("internet"),"没有")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='internet_1'>没有</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>客房描述(或备注)：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><textarea name=remarkinfo cols=50 rows=3 id='remarkinfo' title='请输入您家客房介绍，是单人床还是双人床，有否衣柜，空调，写字台等'>"&rs("remarkinfo")&"</textarea></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	
	
''	show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
''	show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left></div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked><label for='passregtext'>同意</label>　<input class='input_radio' type=radio name=passregtext id=nopassregtext value='0'>不同意　&nbsp;<a href='HomestayBackctrl-InfoHome.asp?action=protocol' target='_blank'> → 查看注册条款</a></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf	
	
''	if oblog.setup(57,0)=1 then
''		show_reg=show_reg&"<tr> "& vbcrlf
''    	show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>验证码：</div></td>"& vbcrlf
''   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
''   		show_reg=show_reg&"</tr>"& vbcrlf
''	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 class='Form-td-left'></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' 保存修改资料 ' style='height:21px;font-size:13px;font-weight:bold'>　"& vbcrlf
	show_reg=show_reg&"<input type='button' name='historybackwl' value='返回上一步' onclick='javascript:history.go(-1);' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'><input type=hidden name=action value='chkreg'><input type=hidden name=FamilyTypeRequest value='"& FamilyTypeRequest &"'>"& vbcrlf
    show_reg=show_reg&"</div></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"</table>"& vbcrlf
	show_reg=show_reg&"</form><br />"& vbcrlf
	

	'show_reg=show_reg&"</div>"
    
	end if
	'set rs=nothing
	response.Write(show_reg)
	%>
	
	<SCRIPT language="javascript">
<!--
function checkssn(gotoURL) {
   var ssn=oblogform.username.value;
   var domain=oblogform.domain.value;
   var domainroot=oblogform.user_domainroot.value;
	   var open_url = gotoURL + "?username=" + ssn+"&domain="+domain+"&domainroot="+domainroot;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=322,height=200');
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i)!='-')) 
{
return 1;
}
}
return 0;//pass
}
function del_space(s)
{
	for(i=0;i<s.length;++i)
	{
	 if(s.charAt(i)!=" ")
		break;
	}

	for(j=s.length-1;j>=0;--j)
	{
	 if(s.charAt(j)!=" ")
		break;
	}

	return s.substring(i,++j);
}

function VerifySubmit()
{

	 city = del_space(document.all("city").value);
     if (city.length == 0)
     {
        alert("请选择所在城市!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("请输入Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("请输入您的真实姓名!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("请选择您的学习范围!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>


<script>
	function ShowPet(ID){
		if(ID=="1"){
		window.PetTypes.innerHTML="<select id=\"PetType\" name=\"PetType\"><option value=\"\" selected>请选择<\/option><option value=\"狗\" >狗<\/option><option value=\"猫\">猫<\/option><option value=\"鸟\">鸟<\/option><option value=\"其他\">其他<\/option><\/select>";
		//document.oblogform.CheckPet.value="1"
		}
		if(ID=="0"){
		window.PetTypes.innerHTML="";
		//document.oblogform.CheckPet.value="0"
		}
		}
</script>

	<%
end sub

function showfilepic(ext)
	ext=lcase(ext)
	if  ext="jpg" then
	response.Write("<img src=""images/filetype/jpg.gif"" class=""fileimg"" alt=""JPG文件"" />")
	elseif  ext="gif" then
	response.Write("<img src=""images/filetype/gif.gif"" class=""fileimg"" alt=""GIF文件"" />")
	elseif  ext="bmp" then
	response.Write("<img src=""images/filetype/bmp.gif"" class=""fileimg"" alt=""BMP文件"" />")
	elseif  ext="png" then
	response.Write("<img src=""images/filetype/png.gif"" class=""fileimg"" alt=""PNG文件"" />")
	elseif  ext="psd" then
	response.Write("<img src=""images/filetype/psd.gif"" class=""fileimg"" alt=""PSD文件"" />")
	elseif ext="rar" or ext="zip" or ext="arj" or ext="sit" then
	response.Write("<img src=""images/filetype/rar.gif"" class=""fileimg"" alt=""压缩文件"" />")
	elseif ext="xsl" then
	response.Write("<img src=""images/filetype/excel.gif"" class=""fileimg"" alt=""Excel文件"" />")
	elseif ext="doc" then
	response.Write("<img src=""images/filetype/word.gif"" class=""fileimg"" alt=""Word文件"" />")
	elseif ext="mp3" then
	response.Write("<img src=""images/filetype/mp3.gif"" class=""fileimg"" alt=""mp3文件"" />")
	elseif ext="rm" or ext="ram" then
	response.Write("<img src=""images/filetype/rm.gif"" class=""fileimg"" alt=""Real文件"" />")
	elseif ext="wmv" or ext="wma" or ext="mpg" or ext="avi" then
	response.Write("<img src=""images/filetype/media.gif"" class=""fileimg"" alt=""媒体文件"" />")
	else
	response.Write("<img src=""images/filetype/blank.gif"" class=""fileimg"" alt=""其他文件"" />")
	end if
end function

sub delfile()
	if id="" then
		oblog.adderrstr( "错误：请指定要删除的文件！")
		oblog.showusererrHome
		exit sub
	end if
	if instr(id,",")>0 then
		dim n,i
		id=FilterIDs(id)
		n=split(id,",")
		for i=0 to ubound(n)
			delonefile(n(i))
		next
	else
		delonefile(id)
	end if
	set rs=nothing
	oblog.showok "删除文件成功！",""
end sub

sub delonefile(id)
	id=clng(id)
	dim userid,filesize,filepath,fso,photofile,isphoto,imgsrc
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_upfile] where fileid=" & id
	else
		sql="select * from [oblog_upfile] where fileid=" & id&" and userid="&oblog.logined_uid
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		userid=rs("userid")
		filesize=clng(rs("file_size"))
		filepath=rs("file_path")
		photofile=rs("photofile")
		isphoto=rs("isphoto")
		rs.delete
		rs.update
		rs.close
		oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&userid)
		if filepath<>"" then 
			imgsrc=filepath
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			if instr("jpg,bmp,gif,png,pcx",right(imgsrc,3))>0 then
				imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
				imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
				if  fso.FileExists(Server.MapPath(imgsrc)) then
					fso.DeleteFile Server.MapPath(imgsrc)
				end if
			end if
			if fso.FileExists(Server.MapPath(filepath)) then 
				fso.DeleteFile Server.MapPath(filepath)
			end if
			set fso=nothing
		end if	
	else
		rs.close
	end if
end sub






sub sub_chkreg()
	if oblog.ChkPost()=false then
		oblog.adderrstr("系统不允许从外部提交！")
		oblog.showerr
		exit sub
	end if
	
''	if oblog.setup(57,0)=1 then
''		if not oblog.codepass then oblog.adderrstr("验证码错误，请刷新后重新输入！"):oblog.showerr
''	end if
	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	Dim birthday,job,tel,address
	Dim mobile,postnum,familynumber,familymember,houseframe,housespace,pet,PetType,asksex,issaychinese,howknowsite,englishlevel,car,fitment,toilet,computer,internet,remarkinfo,FamilyType
	
	
	
	if oblog.chkiplock() then
		oblog.adderrstr("对不起！你的IP已被锁定,不允许操作，请您发信息给管理员或工作人员！")
		oblog.showerr
		exit sub
	end if
''	regusername=oblog.filt_badstr(trim(request("username")))
''	if regusername="" then
''		response.Redirect "index.asp"
''	end if
''	regpassword=request("password")
''	re_regpassword=request("repassword")
	sex=request("sex")
	email=trim(request("email"))
'	'question=trim(request("question"))
'	'answer=trim(request("an"))
''	question="您尚未设置密码提示，如无法登录请与管理员联系，并为您找回密码：）"
''	answer="我家李慧颖"
	blogname=trim(request("blogname"))
	usertype=clng(request("usertype"))
	'nickname=trim(request("nickname"))
	user_domain=Lcase(trim(request("domain")))
	user_domainroot=trim(request("user_domainroot"))
	
	birthday=Trim(Request("selecty1"))&"-"&Trim(Request("selectm1"))&"-"&Trim(Request("selectd1"))
	job= Trim(Request("job"))
	tel= Trim(Request("tel"))
	address= Trim(Request("address"))
	
	mobile= Trim(Request("mobile"))
	postnum= Trim(Request("postnum"))
	familynumber= Trim(Request("familynumber"))
	familymember= Trim(Request("familymember"))
	houseframe= Trim(Request("houseframe"))
	housespace= Trim(Request("housespace"))
	pet= Trim(Request("pet"))
	PetType= Trim(Request("PetType"))
	asksex= Trim(Request("asksex"))
	issaychinese= Trim(Request("issaychinese"))
	howknowsite= Trim(Request("howknowsite"))
	englishlevel= Trim(Request("englishlevel"))
	car= Trim(Request("car"))
	fitment= Trim(Request("fitment"))
	toilet= Trim(Request("toilet"))
	computer= Trim(Request("computer"))
	internet= Trim(Request("internet"))
	remarkinfo= Trim(Request("remarkinfo"))
	
	FamilyType= Trim(Request("FamilyType"))
	
	Dim FamilyTypeRequest
	FamilyTypeRequest = request.Form("FamilyTypeRequest")
	
''	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("用户名不能为空，并且不能大于14小于4！")
''	if oblog.chk_regname(regusername) then oblog.adderrstr("用户名系统不允许注册！")
''	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("用户名中含有系统不允许的字符！")
''	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("用户名不允许全部为数字！")
''	if oblog.chkdomain(regusername)=false then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字及下划线！")
''	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
''		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
''		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
''		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
''		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
''		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字！")
''		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
''	end if
''	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("密码不能为空(不能大于14小于4)！")
''	if re_regpassword="" then oblog.adderrstr("重复密码不能为空！")
''	if regpassword<>re_regpassword then oblog.adderrstr("两次输入密码不同！")
''	if question="" or oblog.strLength(question)>150 then oblog.adderrstr("找回密码提示问题不能为空，也不能大于150！")
''	if answer="" or oblog.strLength(answer)>150 then oblog.adderrstr("找回密码问题答案不能为空，也不能大于150！")
	'if oblog.chk_regname(nickname) then oblog.adderrstr("此昵称系统不允许注册！")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("昵称不能不能大于50字符！")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("真实姓名不能为空(不能大于50字符)！")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("真实姓名中含有系统不允许的字符！")	
	''if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("用户名中含有非法字符！")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个昵称存在，请更改昵称！")		
	'end if
''	if user_domain<>"" then
''		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
''		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个域名存在，请更改域名！")
''	end if
	
	
	If IsDate(Cdate(birthday))=False Then oblog.adderrstr("您的生日格式尚未填好，请重新选择！")	

	
	
	
	if oblog.errstr<>"" then oblog.showerr:exit sub
	if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
''	regpassword=md5(regpassword)
	if not IsObject(conn) then link_database
	set rsreg=server.CreateObject("adodb.recordset")
	rsreg.open "select * from [oblog_user] where userid="& oblog.logined_uid &"",conn,1,3
	if not rsreg.eof then
''		rsreg.addnew
''		rsreg("username")=regusername
''		rsreg("password")=regpassword
''		if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
''			rsreg("user_domain")=user_domain
''			rsreg("user_domainroot")=user_domainroot
''		end if
''		rsreg("question")=question
''		rsreg("answer")=md5(answer)
		rsreg("useremail")=email
''		rsreg("user_level")=reguserlevel
		rsreg("user_isbest")=0
		rsreg("blogname")=blogname
		rsreg("user_classid")=usertype
		'rsreg("nickname")=nickname
		rsreg("province")=request("province")
		rsreg("city")=request("city")
''		rsreg("adddate")=ServerDate(now())
rsreg("lastModifyDate")=ServerDate(now())
		rsreg("lastloginip")=oblog.userip
		rsreg("lastlogintime")=ServerDate(now())
''		rsreg("user_dir")=oblog.setup(30,0)
''		rsreg("user_folder")=regusername
		
		rsreg("sex")= sex
		rsreg("birthday")= birthday
		rsreg("job")= job
		rsreg("tel")= tel
		rsreg("address")= address
		
		rsreg("mobile")= mobile
		rsreg("postnum")= postnum
		rsreg("familynumber")= familynumber
		rsreg("familymember")= familymember
		rsreg("houseframe")= houseframe
		rsreg("housespace")= housespace
		rsreg("pet")= pet
		rsreg("PetType")= PetType
		rsreg("asksex")= asksex
		rsreg("issaychinese")= issaychinese
		rsreg("howknowsite")= howknowsite
		rsreg("englishlevel")= englishlevel
		rsreg("car")= car
		rsreg("fitment")= fitment
		rsreg("toilet")= toilet
		rsreg("computer")= computer
		rsreg("internet")= internet
		rsreg("remarkinfo")= remarkinfo
		
		rsreg("FamilyType")= Request("FamilyType")
			
		rsreg.update
''		oblog.execute("update oblog_setup set user_count=user_count+1")
''		if is_unamedir=0 then
''			oblog.execute("update oblog_user set user_folder=userid where username='"&oblog.filt_badstr(regusername)&"'")
''		end if
''		show_reg="当前位置：<a href='index.asp'>首页</a>→完成注册<hr />"
''		oblog.CreateUserDir regusername,1
''		if oblog.setup(16,0)=1 then
''			show_reg=show_reg&"<ul><li><strong>恭喜！您已经成功注册为 </strong><font color=#038ad7>"& FamilyTypeRequest &"</font> ！</li></ul>"
''			show_reg=show_reg&"10秒后自动转到系统首页。"
''			show_reg=show_reg&"<script language=JavaScript>"
''			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",10000);"
''			show_reg=show_reg&"< /script> "
''		else
''			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
''			show_reg=show_reg&"<ul><li><strong>恭喜！您已经注册成功！</strong></li>"
''			show_reg=show_reg&"<li><a href='index.asp'>返回首页</a></li>"
''			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>点击进入管理后台选择您的喜欢的页面风格(5秒后自动进入管理后台)</a></li></ul>"
''			show_reg=show_reg&"<script language=JavaScript>"
''			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
''			show_reg=show_reg&"< /script>"
''		end if
	else
		oblog.adderrstr("系统中没有您的帐号存在，请再试一次，如果还有问题请联系管理员或工作人员！")
		oblog.showerr
		exit sub	
	end if
	rsreg.close
	set rsreg=nothing
	oblog.showok "保存修改成功！",""
	'ATFLAG
	'Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=now()	
''	Session("chk_regtime")=Now()
end sub
%>