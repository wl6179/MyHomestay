<!--#include file="user_topHome.asp"-->
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
				<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13146398085</span> &nbsp;&nbsp;&nbsp;
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
					
					
					
						<center><h2>�޸��ҵ�����Ӵ���ͥ����</h2></center>
						
			


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
	call sub_chkreg()'�����޸ļ�ͥ��������;
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
								��ͥ�û�����
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
			strGuide=strGuide & "����ǰ�����������ǣ�"
		case 1
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' ) order by fileid desc"
			end if
			strGuide=strGuide & "ͼƬ�ļ�"
		case 2
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit') order by fileid desc"
			end if
			strGuide=strGuide & "ѹ���ļ�"
		case 3
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='doc' or file_ext='xsl' or file_ext='txt' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='doc' or file_ext='xsl' or file_ext='txt') order by fileid desc"
			end if
			strGuide=strGuide & "�ĵ��ļ�"
		case 4
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where  file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf') order by fileid desc"
			end if
			strGuide=strGuide & "ý���ļ�"
		case 5
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where isphoto=1 order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and isphoto=1 order by fileid desc"
			end if
			strGuide=strGuide & "���"
		case else
	end select
	'strGuide=strGuide & "</td><td align='right'>"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (����0���ļ�)</h1><div id=""list"">"
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
	FamilyTypeRequest= "��ѽӴ���ͥ"
	Case "2"
	FamilyTypeRequest= "�շѽӴ���ͥ"
	Case "3"
	FamilyTypeRequest= "2008���˽Ӵ���ͥ"
	Case "1, 2"
	FamilyTypeRequest= "���/�շѽӴ���ͥ"
	Case "1, 3"
	FamilyTypeRequest= "��ѽӴ���ͥ/2008���˽Ӵ���ͥ"
	Case "1, 2, 3"
	FamilyTypeRequest= "���/�շ�/2008���˽Ӵ���ͥ"
	Case "2, 3"
	FamilyTypeRequest= "�շѽӴ���ͥ/2008���˽Ӵ���ͥ"
    End Select
	'FamilyType_n = Split(FamilyType, ",")
        'For i = 0 To UBound(FamilyType_n)
            'response.Write FamilyType_n(i)
        'Next
	'Response.Write(rs("FamilyType"))

	show_reg=show_reg&"<font color=#038ad7>����ǰ������� :</font></center>"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"��ǰϵͳ�ѹر�ע�ᡣ"
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
	  show_reg=show_reg&"������ѽӴ���ͥ "
	  
	  If instr(FamilyType,"2")>0 Then
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2' checked='checked'>"
	  Else
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2'>"
	  End If
	  show_reg=show_reg&"�����շѽӴ���ͥ "
	  
	  If instr(FamilyType,"3")>0 Then
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3' checked='checked'>"
	  Else
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3'>"
	  End If
	  show_reg=show_reg&"<font color='#038ad7'>����2008���˽Ӵ���ͥ</font>"
	  
	  
	  show_reg=show_reg&"<br />"
	  
	  
	  show_reg=show_reg&"<br />"
	

	
	
	
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ�ʺ���Ϣ</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ͥ��¼�ʺţ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=63% colspan=2>"&  rs("username")& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=20% height=25><div align=left >������</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
'    show_reg=show_reg&"<td width=63% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>�鿴�����ʺ��Ƿ�ռ��</a>"& vbcrlf 
'    show_reg=show_reg&"</td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
''    show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�������룺</div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=password type=password id=password size=15 maxlength=30><font color=#038ad7>�粻�޸ģ��뱣��Ϊ��</font></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
''	
''    show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��¼����ȷ�ϣ�</div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=repassword type=password id=repassword size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>������ʾ���⣺</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>������ʾ�𰸣�</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>���ڵ���(ʡ/��)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_city(rs("province"),rs("city")) &"<font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>Email��ַ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=35 maxlength=30 value='"& rs("useremail")& "'><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ʵ������</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30 value='"& rs("blogname") &"'><font color=#038ad7> *</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ĳƺ���<select name=sex>"
		If Cint(rs("sex"))=3 Then
		show_reg=show_reg&"<option value=3 selected='selected'>��ѡ��</option>"
		Else
		show_reg=show_reg&"<option value=3>��ѡ��</option>"
		End If
		If Cint(rs("sex"))=0 Then
		show_reg=show_reg&"<option value=0 selected='selected'>Ůʿ</option>"
		Else
		show_reg=show_reg&"<option value=0>Ůʿ</option>"
		End If
		If Cint(rs("sex"))=1 Then
		show_reg=show_reg&"<option value=1 selected='selected'>����</option>"
		Else
		show_reg=show_reg&"<option value=1>����</option>"
		End If
		
	show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ѡ��ѧϰ��Χ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�������ڣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"

	show_reg=show_reg& oblog.type_dateselect(rs("birthday"),1) &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>����ְҵ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_job2(rs("job")) &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>��ϵ��ʽ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�绰��<input name=tel value='"& rs("tel") &"' size='15' maxlength='50' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ֻ���<input name=mobile  value='"& rs("mobile") &"' size='15' maxlength='50' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ַ��<input name='address' value='"& rs("address") &"' size='30' maxlength='250' />&nbsp;&nbsp;&nbsp;�ʱࣺ<input name=postnum value='"& rs("postnum") &"' size='8' maxlength='6' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>��ͥ��Ա��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ������<input id='familynumber_0' type='radio' name='familynumber' value='1' "
	If Instr(rs("familynumber"),"1")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_0'>1</label>&nbsp;&nbsp;<input id='familynumber_1' type='radio' name='familynumber' value='2' "
	If Instr(rs("familynumber"),"2")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_1'>2</label>&nbsp;&nbsp;<input id='familynumber_2' type='radio' name='familynumber' value='3' "
	If Instr(rs("familynumber"),"3")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_2'>3</label>&nbsp;&nbsp;<input id='familynumber_3' type='radio' name='familynumber' value='4' "
	If Instr(rs("familynumber"),"4")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_3'>4</label>&nbsp;&nbsp;<input id='familynumber_4' type='radio' name='familynumber' value='����' "
	If Instr(rs("familynumber"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='familynumber_4'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ�ṹ��<input name='familymember' type='checkbox' id='Father' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Father'>����</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Mother' value='ĸ��' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"ĸ��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Mother'>ĸ��</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Husband' value='�ɷ�' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"�ɷ�")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Husband'>�ɷ�</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Wife' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Wife'>����</label><br>&nbsp;&nbsp;<font color=#ffffff>Homestay</font><input name='familymember' type='checkbox' id='Son' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Son'>����</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Daughter' value='Ů��' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"Ů��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Daughter'>Ů��</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Brother' value='�ֵ�' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"�ֵ�")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Brother'>�ֵ�</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Sister' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ' "
	If Instr(rs("familymember"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" ><label for='Sister'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
                           
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>���ݽ��ܣ�</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�ṹ��<select name='houseframe' id='houseframe1' title='��ѡ�����ķ��ݽṹ'>"
                            show_reg=show_reg&"<option value='1��' "
							If Instr(rs("houseframe"),"1��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1��</option>"
                            show_reg=show_reg&"<option value='2��' "
							If Instr(rs("houseframe"),"2��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2��</option>"
                            show_reg=show_reg&"<option value='3��' "
							If Instr(rs("houseframe"),"3��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3��</option>"
                            show_reg=show_reg&"<option value='4��' "
							If Instr(rs("houseframe"),"4��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4��</option>"
                            show_reg=show_reg&"<option value='����' "
							If Instr(rs("houseframe"),"����")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">����</option>"
                            show_reg=show_reg&"<option value='����' "
							If Instr(rs("houseframe"),"����")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">����</option>"
        show_reg=show_reg&"</select>&nbsp;<select name='houseframe' id='houseframe2'  title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='1��' "
							If Instr(rs("houseframe"),"1��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1��</option>"
                            show_reg=show_reg&"<option value='2��' "
							If Instr(rs("houseframe"),"2��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2��</option>"
                            show_reg=show_reg&"<option value='3��' "
							If Instr(rs("houseframe"),"3��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3��</option>"
                            show_reg=show_reg&"<option value='4��' "
							If Instr(rs("houseframe"),"4��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4��</option>"
                            show_reg=show_reg&"<option value='5��' "
							If Instr(rs("houseframe"),"5��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">5��</option>"
                            show_reg=show_reg&"<option value='6��' "
							If Instr(rs("houseframe"),"6��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">6��</option>"
        show_reg=show_reg&"</select>&nbsp;<select name='houseframe' id='houseframe3' title='��ѡ�����ķ��ݽṹ'>"
                            show_reg=show_reg&"<option value='1��' "
							If Instr(rs("houseframe"),"1��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">1��</option>"
                            show_reg=show_reg&"<option value='2��' "
							If Instr(rs("houseframe"),"2��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">2��</option>"
                            show_reg=show_reg&"<option value='3��' "
							If Instr(rs("houseframe"),"3��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">3��</option>"
                            show_reg=show_reg&"<option value='4��' "
							If Instr(rs("houseframe"),"4��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">4��</option>"
                            show_reg=show_reg&"<option value='5��' "
							If Instr(rs("houseframe"),"5��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">5��</option>"
                            show_reg=show_reg&"<option value='6��' "
							If Instr(rs("houseframe"),"6��")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">6��</option>"
                          show_reg=show_reg&"</select>"&_
						  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����"&_ 
                          "<select name='housespace' id='housespace' style='width:100px' title='��ѡ�����ķ������'>"
                            show_reg=show_reg&"<option>��ѡ��</option>"
                            show_reg=show_reg&"<option value='50-100' "
							If Instr(rs("housespace"),"50-100")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">50��100ƽ����</option>"
                            show_reg=show_reg&"<option value='100-150' "
							If Instr(rs("housespace"),"100-150")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">100��150ƽ����</option>"
                            show_reg=show_reg&"<option value='150-200' "
							If Instr(rs("housespace"),"150-200")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">150��200ƽ����</option>"
                            show_reg=show_reg&"<option value='above 200' "
							If Instr(rs("housespace"),"above 200")>0 Then show_reg=show_reg&"selected"
							show_reg=show_reg&">200ƽ��������</option>"
                          show_reg=show_reg&"</select>"&_
						  "</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>�Ƿ��г��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value='��' name='pet' id='Pet1' onClick='ShowPet(1);' title='��ѡ�����ĳ�������' "
	If Instr(rs("pet"),"û��")=0 Then show_reg=show_reg&"checked"
	show_reg=show_reg&"><label for='Pet1'>��</label>&nbsp;<INPUT name=pet type=radio value='û��' id='Pet2' onClick='ShowPet(0);' "
	If Instr(rs("pet"),"û��")>0 Then show_reg=show_reg&"checked"
	show_reg=show_reg&" ><label for='Pet2'>û��</label>&nbsp;&nbsp;&nbsp;<span width='21%' id='PetTypes'>"
	
	If Instr(rs("pet"),"û��")=0 Then
		show_reg=show_reg&"<select id='PetType' name='PetType'><option value=''>��ѡ��</option><option value='��' "
		If Instr(rs("PetType"),"��")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">��</option><option value='è' "
		If Instr(rs("PetType"),"è")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">è</option><option value='��' "
		If Instr(rs("PetType"),"��")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">��</option><option value='����' "
		If Instr(rs("PetType"),"����")>0 Then show_reg=show_reg&"selected"
		show_reg=show_reg&">����</option></select>"
	End If
	show_reg=show_reg&"</span></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                      
    
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>����̵�����Ҫ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>����Ա�<input id='asksex_0' type='radio' name='asksex' value='��' "
	If Instr(rs("asksex"),"��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_0'>��</label>&nbsp;&nbsp;&nbsp;<input id='asksex_1' type='radio' name='asksex' value='Ů' "
	If Instr(rs("asksex"),"Ů")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_1'>Ů</label>&nbsp;&nbsp;&nbsp;<input id='asksex_2' type='radio' name='asksex' value='����ν' "
	If Instr(rs("asksex"),"����ν")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='asksex_2'>����ν</label> &nbsp;&nbsp;&nbsp;</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�Ƿ��˵���ģ�<input id='issaychinese_0' type='radio' name='issaychinese' value='��' "
	If Instr(rs("issaychinese"),"��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_0'>��</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_1' type='radio' name='issaychinese' value='��' "
	If Instr(rs("issaychinese"),"��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_1'>��</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_2' type='radio' name='issaychinese' value='����ν' "
	If Instr(rs("issaychinese"),"����ν")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='issaychinese_2'>����ν</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	                    
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>���֪�����ǣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input id='howknowsite_0' type='checkbox' name='howknowsite' value='��ֽ��־' "
	If Instr(rs("howknowsite"),"��ֽ��־")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_0'>��ֽ��־</label>&nbsp;&nbsp;<input id='howknowsite_1' type='checkbox' name='howknowsite' value='���ѽ���' "
	If Instr(rs("howknowsite"),"���ѽ���")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_1'>���ѽ���</label>&nbsp;&nbsp;<input id='howknowsite_2' type='checkbox' name='howknowsite' value='��վ' "
	If Instr(rs("howknowsite"),"��վ")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_2'>��վ</label>&nbsp;&nbsp;<input id='howknowsite_3' type='checkbox' name='howknowsite' value='��̳' "
	If Instr(rs("howknowsite"),"��̳")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_3'>��̳</label>&nbsp;&nbsp;<input id='howknowsite_4' type='checkbox' name='howknowsite' value='��������' "
	If Instr(rs("howknowsite"),"��������")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_4'>��������</label>&nbsp;&nbsp;<input id='howknowsite_5' type='checkbox' name='howknowsite' value='����' "
	If Instr(rs("howknowsite"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&" /><label for='howknowsite_5'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>Ӣ��ˮƽ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name='englishlevel' id='englishlevel_0' type='radio' value='����' title='��ѡ����Ӣ��ˮƽ' "
	If Instr(rs("englishlevel"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_0'>����</label>&nbsp;&nbsp;<input name='englishlevel' id='englishlevel_1' type='radio' value='һ��' checked title='��ѡ����Ӣ��ˮƽ' "
	If Instr(rs("englishlevel"),"һ��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_1'>һ��</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_2' value='�е�' title='��ѡ����Ӣ��ˮƽ' "
	If Instr(rs("englishlevel"),"�е�")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_2'>�е�</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_3' value='�ܺ�' title='��ѡ����Ӣ��ˮƽ' "
	If Instr(rs("englishlevel"),"�ܺ�")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='englishlevel_3'>�ܺ�</label>&nbsp;&nbsp;&nbsp;&nbsp;<strong>�Ƿ���˽�ҳ�</strong>��<INPUT name=car id=car_0 type=radio value=�� title='��ѡ���Ƿ���˽�ҳ�' "
	If Instr(rs("car"),"û��")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='car_0'>��</label>&nbsp;<INPUT name=car id=car_1 type=radio value=û�� title='��ѡ���Ƿ���˽�ҳ�' "
	If Instr(rs("car"),"û��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='car_1'>û��</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>����������ͣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT name=toilet id=toilet_0 type=radio value=���� title='��ѡ���Ƿ��ж�����������' "
	If Instr(rs("toilet"),"����")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='toilet_0'>����</label><INPUT type=radio value=������� name=toilet id=toilet_1 title='��ѡ���Ƿ��ж�����������' "
	If Instr(rs("toilet"),"�������")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='toilet_1'>�������</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�з���ԣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value=�� name=computer id=computer_0 title='��ѡ���Ƿ��е���' "
	If Instr(rs("computer"),"û��")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='computer_0'>��</label>&nbsp;&nbsp;<INPUT name=computer id=computer_1 type=radio value=û�� title='��ѡ���Ƿ��е���' "
	If Instr(rs("computer"),"û��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='computer_1'>û��</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>�з���������</strong><INPUT type=radio value=�� name=internet id=internet_0 title='��ѡ���Ƿ��п��������' "
	If Instr(rs("internet"),"û��")=0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='internet_0'>��</label>&nbsp;&nbsp;<INPUT name=internet id=internet_1 type=radio value=û�� title='��ѡ���Ƿ��п��������' "
	If Instr(rs("internet"),"û��")>0 Then show_reg=show_reg&"checked='checked'"
	show_reg=show_reg&"><label for='internet_1'>û��</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�ͷ�����(��ע)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><textarea name=remarkinfo cols=50 rows=3 id='remarkinfo' title='���������ҿͷ����ܣ��ǵ��˴�����˫�˴����з��¹񣬿յ���д��̨��'>"&rs("remarkinfo")&"</textarea></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	
	
''	show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
''	show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left></div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked><label for='passregtext'>ͬ��</label>��<input class='input_radio' type=radio name=passregtext id=nopassregtext value='0'>��ͬ�⡡&nbsp;<a href='HomestayBackctrl-InfoHome.asp?action=protocol' target='_blank'> �� �鿴ע������</a></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf	
	
''	if oblog.setup(57,0)=1 then
''		show_reg=show_reg&"<tr> "& vbcrlf
''    	show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��֤�룺</div></td>"& vbcrlf
''   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
''   		show_reg=show_reg&"</tr>"& vbcrlf
''	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 class='Form-td-left'></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' �����޸����� ' style='height:21px;font-size:13px;font-weight:bold'>��"& vbcrlf
	show_reg=show_reg&"<input type='button' name='historybackwl' value='������һ��' onclick='javascript:history.go(-1);' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'><input type=hidden name=action value='chkreg'><input type=hidden name=FamilyTypeRequest value='"& FamilyTypeRequest &"'>"& vbcrlf
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
        alert("��ѡ�����ڳ���!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("������Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("������������ʵ����!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("��ѡ������ѧϰ��Χ!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>


<script>
	function ShowPet(ID){
		if(ID=="1"){
		window.PetTypes.innerHTML="<select id=\"PetType\" name=\"PetType\"><option value=\"\" selected>��ѡ��<\/option><option value=\"��\" >��<\/option><option value=\"è\">è<\/option><option value=\"��\">��<\/option><option value=\"����\">����<\/option><\/select>";
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
	response.Write("<img src=""images/filetype/jpg.gif"" class=""fileimg"" alt=""JPG�ļ�"" />")
	elseif  ext="gif" then
	response.Write("<img src=""images/filetype/gif.gif"" class=""fileimg"" alt=""GIF�ļ�"" />")
	elseif  ext="bmp" then
	response.Write("<img src=""images/filetype/bmp.gif"" class=""fileimg"" alt=""BMP�ļ�"" />")
	elseif  ext="png" then
	response.Write("<img src=""images/filetype/png.gif"" class=""fileimg"" alt=""PNG�ļ�"" />")
	elseif  ext="psd" then
	response.Write("<img src=""images/filetype/psd.gif"" class=""fileimg"" alt=""PSD�ļ�"" />")
	elseif ext="rar" or ext="zip" or ext="arj" or ext="sit" then
	response.Write("<img src=""images/filetype/rar.gif"" class=""fileimg"" alt=""ѹ���ļ�"" />")
	elseif ext="xsl" then
	response.Write("<img src=""images/filetype/excel.gif"" class=""fileimg"" alt=""Excel�ļ�"" />")
	elseif ext="doc" then
	response.Write("<img src=""images/filetype/word.gif"" class=""fileimg"" alt=""Word�ļ�"" />")
	elseif ext="mp3" then
	response.Write("<img src=""images/filetype/mp3.gif"" class=""fileimg"" alt=""mp3�ļ�"" />")
	elseif ext="rm" or ext="ram" then
	response.Write("<img src=""images/filetype/rm.gif"" class=""fileimg"" alt=""Real�ļ�"" />")
	elseif ext="wmv" or ext="wma" or ext="mpg" or ext="avi" then
	response.Write("<img src=""images/filetype/media.gif"" class=""fileimg"" alt=""ý���ļ�"" />")
	else
	response.Write("<img src=""images/filetype/blank.gif"" class=""fileimg"" alt=""�����ļ�"" />")
	end if
end function

sub delfile()
	if id="" then
		oblog.adderrstr( "������ָ��Ҫɾ�����ļ���")
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
	oblog.showok "ɾ���ļ��ɹ���",""
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
		oblog.adderrstr("ϵͳ��������ⲿ�ύ��")
		oblog.showerr
		exit sub
	end if
	
''	if oblog.setup(57,0)=1 then
''		if not oblog.codepass then oblog.adderrstr("��֤�������ˢ�º��������룡"):oblog.showerr
''	end if
	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	Dim birthday,job,tel,address
	Dim mobile,postnum,familynumber,familymember,houseframe,housespace,pet,PetType,asksex,issaychinese,howknowsite,englishlevel,car,fitment,toilet,computer,internet,remarkinfo,FamilyType
	
	
	
	if oblog.chkiplock() then
		oblog.adderrstr("�Բ������IP�ѱ�����,�������������������Ϣ������Ա������Ա��")
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
''	question="����δ����������ʾ�����޷���¼�������Ա��ϵ����Ϊ���һ����룺��"
''	answer="�Ҽ����ӱ"
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
	
''	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("�û�������Ϊ�գ����Ҳ��ܴ���14С��4��")
''	if oblog.chk_regname(regusername) then oblog.adderrstr("�û���ϵͳ������ע�ᣡ")
''	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("�û����к���ϵͳ��������ַ���")
''	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("�û���������ȫ��Ϊ���֣�")
''	if oblog.chkdomain(regusername)=false then oblog.adderrstr("�û������Ϲ淶��ֻ��ʹ��Сд��ĸ�����ּ��»��ߣ�")
''	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
''		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("��������Ϊ��(���ܴ���14���ַ�)��")
''		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("��������С��4���ַ���")
''		if oblog.chk_regname(user_domain) then oblog.adderrstr("������ϵͳ������ע�ᣡ")
''		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
''		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
''		if user_domainroot="" then oblog.adderrstr("����������Ϊ�գ�")
''	end if
''	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("���벻��Ϊ��(���ܴ���14С��4)��")
''	if re_regpassword="" then oblog.adderrstr("�ظ����벻��Ϊ�գ�")
''	if regpassword<>re_regpassword then oblog.adderrstr("�����������벻ͬ��")
''	if question="" or oblog.strLength(question)>150 then oblog.adderrstr("�һ�������ʾ���ⲻ��Ϊ�գ�Ҳ���ܴ���150��")
''	if answer="" or oblog.strLength(answer)>150 then oblog.adderrstr("�һ���������𰸲���Ϊ�գ�Ҳ���ܴ���150��")
	'if oblog.chk_regname(nickname) then oblog.adderrstr("���ǳ�ϵͳ������ע�ᣡ")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("�ǳ��к���ϵͳ��������ַ���")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("�ǳƲ��ܲ��ܴ���50�ַ���")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("��ʵ��������Ϊ��(���ܴ���50�ַ�)��")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("��ʵ�����к���ϵͳ��������ַ���")	
	''if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"��")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("�û����к��зǷ��ַ���")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������ǳƴ��ڣ�������ǳƣ�")		
	'end if
''	if user_domain<>"" then
''		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
''		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������������ڣ������������")
''	end if
	
	
	If IsDate(Cdate(birthday))=False Then oblog.adderrstr("�������ո�ʽ��δ��ã�������ѡ��")	

	
	
	
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
''		show_reg="��ǰλ�ã�<a href='index.asp'>��ҳ</a>�����ע��<hr />"
''		oblog.CreateUserDir regusername,1
''		if oblog.setup(16,0)=1 then
''			show_reg=show_reg&"<ul><li><strong>��ϲ�����Ѿ��ɹ�ע��Ϊ </strong><font color=#038ad7>"& FamilyTypeRequest &"</font> ��</li></ul>"
''			show_reg=show_reg&"10����Զ�ת��ϵͳ��ҳ��"
''			show_reg=show_reg&"<script language=JavaScript>"
''			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",10000);"
''			show_reg=show_reg&"< /script> "
''		else
''			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
''			show_reg=show_reg&"<ul><li><strong>��ϲ�����Ѿ�ע��ɹ���</strong></li>"
''			show_reg=show_reg&"<li><a href='index.asp'>������ҳ</a></li>"
''			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>�����������̨ѡ������ϲ����ҳ����(5����Զ���������̨)</a></li></ul>"
''			show_reg=show_reg&"<script language=JavaScript>"
''			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
''			show_reg=show_reg&"< /script>"
''		end if
	else
		oblog.adderrstr("ϵͳ��û�������ʺŴ��ڣ�������һ�Σ����������������ϵ����Ա������Ա��")
		oblog.showerr
		exit sub	
	end if
	rsreg.close
	set rsreg=nothing
	oblog.showok "�����޸ĳɹ���",""
	'ATFLAG
	'Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=now()	
''	Session("chk_regtime")=Now()
end sub
%>