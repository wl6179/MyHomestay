<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->

<%
'Ϊ��ͥͳ�Ʒ�����WL;

%>

<%
'if en_photo=0 then
'	oblog.adderrstr("�˹����ѱ�ϵͳ�رգ�")
'	oblog.showerr
'	Set oblog = Nothing
'end if
dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,strPlayerUrl

Dim id
id =Request("id")

strurl="HomestayFamilyMember.asp"
isbest=cint(request.QueryString("isbest"))
classid=request.QueryString("classid")
call sysshow_3()

''if isbest=1 then
''	mainsql=mainsql&" and isBestUser=1"
''	if strurl="HomestayFamilyMember.asp" then
''		strurl=strurl&"?isbest=1"
''	else
''		strurl=strurl&"&isbest=1"
''	end if
''	bstr1="���Ƽ��Ӵ���ͥ"
''end if

'If IsNumeric(classid) Then
'	classid=CLNG(classid)
'	If classid>0 Then 
'		mainsql= mainsql & " and sysClassid=" & classid & " "
'		if strurl="HomestayFamilyMember.asp" then
'			strurl=strurl&"?classid=" & classid
'		else
'			strurl=strurl&"&classid=" & classid
'		end if
'	End IF
'End If

strPlayerUrl= Replace(strurl,"HomestayFamilyMember.asp","photoplayer.asp")
		
call sub_showuserlist(mainsql,strurl)

show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom

sub sub_showuserlist(sql,strurl)
	dim topn,msql
	MaxPerPage=oblog.setup(79,0)
	strFileName=strurl'���ݺ���url��
''	if request("page")<>"" then
''    	currentPage=cint(request("page"))
''	else
''		currentPage=1
''	end if
	'''msql="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1 "&sql&" And isPower=0 order by fileid desc"
''	Dim sql001,sql002,sql003'��������WL;
''	sql001 = " And user_level<=6 And user_level_team=0 "'ɸѡ��ͥ�û����ļ���
''	'''sql002 = " And isUserPublicForView=1 And isCover=1 "'ɸѡ����ͥ�û���Ϊ����+������ļ���
''	sql003 = " And isValidate=1 "'ɸѡΪ��ͨ��վ����˵ļ�ͥ�û���
	'''msql="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto<>11111 "&sql&" And isPower<>11110  "& sql001 & sql002 & sql003 &" order by fileid desc"
''	msql="Select Top "&clng(oblog.setup(78,0))&" * from [oblog_user] Where userid<>0"& sql001 & sql002 & sql003 &" order by userid desc"

	Dim sql001,sql002,sql003'��������WL;
	sql001 = " And user_level<=6 And user_level_team=0 "'ɸѡΪ��ͥ�û���

	sql003 = " And isValidate=1 "'ɸѡΪ��ͨ��վ����˵ļ�ͥ�û���
	'��ע�빦��WL[2007-09-23];
	if isNumeric(id) then
		id = Cint(id)
	else
		Response.Write "<br>sorry,attack is denervate..."
		Response.end
	end if
	msql="Select * From [oblog_user] Where userid="& id & sql001 & sql002 & sql003 
	
	
	
	
	
	
	
	
	
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "�޷���ѯ�˽Ӵ���ͥ���������ݶ�ʧ�����߸ü�ͥ��ʱ��δͨ�����...<br>"
	else
		'Ϊ��ͥͳ�Ʒ�����WL;

		Session("tempGuestNow")=Session("tempGuestNow") & "1"'���浱ǰ�ÿ͵���ʱip��
		If Len(Session("tempGuestNow"))< 2 Or Len(Session("tempGuestNow"))> 2 Then oblog.execute("Update oblog_user Set user_siterefu_num=user_siterefu_num+1 Where userid="&id )
		'''response.Write Session("tempGuestNow")
		
''    	totalPut=rsmain.recordcount
''		'show_list=show_list & "������" & totalPut & " ����ͥ<br>"
''		if currentpage<1 then
''       		currentpage=1
''    	end if
''    	if (currentpage-1)*MaxPerPage>totalput then
''	   		if (totalPut mod MaxPerPage)=0 then
''	     		currentpage= totalPut \ MaxPerPage
''		  	else
''		      	currentpage= totalPut \ MaxPerPage + 1
''	   		end if
''
''    	end if
''	    if currentPage=1 then

        	getlist()
			
''        	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")
''   	 	else
''   	     	if (currentPage-1)*MaxPerPage<totalPut then
''         	   	rsmain.move  (currentPage-1)*MaxPerPage
''         		dim bookmark
''           		bookmark=rsmain.bookmark
''            	getlist()
''            	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")
''        	else
''	        	currentPage=1
''           		getlist()
''           		show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")
''	    	end if
''		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,bstr,n,fso
	dim title,userurl,imgsrc
''	Set fso = server.CreateObject("Scripting.FileSystemObject")
	show_list="<table width='100%' border='0'><tr><td>"
	show_list=show_list&"<center><h5><img src='/images/hi_myhomestay.gif' /> <font color=#E93D7D>��ͥ��ţ�MH00"& rsmain("userid") &"</font></h5></center>"
	'''If not rsmain.eof Then show_list=show_list& rsmain("blogname") &"</center>"
	'bstr=trim(Request.ServerVariables("query_string"))
	'if bstr<>"" then bstr="HomestayFamilyMember.asp?"&replace(replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="HomestayFamilyMember.asp?isbest=1"	
	show_list=show_list&bstr1&"</td><td align='right'> </td></tr>"
	show_list=show_list& "</table>"
	
	
	show_list=show_list&"<table width='100%'  align='center' cellpadding='0' cellspacing='1' style='position:relative;width:680px;font:13px;'>"& vbcrlf
	If not rsmain.eof Then
		show_list=show_list&"<tr height='22'><td>"& vbcrlf
		''''''---------------------------------
		
		
		dim str_usertype
	Dim show_reg
	
	'show_reg=show_reg&"<div  style='font-size:12px;' CLASS='Register'>"
	Dim FamilyTypeRequest,FamilyType,FamilyType_n
	FamilyType = rsmain("FamilyType")
	
	Select case rsmain("FamilyType")
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
	'Response.Write(rsmain("FamilyType"))


'������Ƭ���������գ�
	show_reg=show_reg&"<br />"
	If rsmain("upfile_isCoverPhoto_path")<>"" And Len(rsmain("upfile_isCoverPhoto_path"))>6 Then
		show_reg=show_reg&"<span onMouseOut=""this.style.backgroundColor=''""  onMouseOver=""this.style.backgroundColor='#87D2F1'""><a href='"& rsmain("upfile_isCoverPhoto_path") &"' target='_blank'>"
		show_reg=show_reg&"<img src='"& rsmain("upfile_isCoverPhoto_path") &"' style=""margin-left:53px;width:320px;height:240px;border:0px;padding:3px; border: 1px #cccccc solid; padding:5px; filter : progid:DXImageTransform.Microsoft.DropShadow(color=#eeeeee,offX=6,offY=5,positives=true);"" onerror=""this.src='/images/unknow_photo02.jpg';"" /></a></span>"
	Else
		show_reg=show_reg&"<img src='/images/unknow_photo02.jpg' style=""margin-left:53px; filter : progid:DXImageTransform.Microsoft.DropShadow(color=#eeeeee,offX=6,offY=5,positives=true);"" />"
	End If
	
	
	
	
	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"��ǰϵͳ�ѹر�ע�ᡣ"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",rsmain("user_classid"),0)
    str_usertype=str_usertype&"</select><font color=#038ad7> *</font>"
	show_reg=show_reg&"<form name=oblogform id=oblogform method=post action=HomestayBackctrl-InfoHome.asp onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    
	
	show_reg=show_reg&"<br /><font color=#038ad7>��ǰ������� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>"
	  If instr(FamilyType,"1")>0 Then
	  show_reg=show_reg&"<font color='#038ad7'>��</font>������ѽӴ���ͥ"
	  Else
	  show_reg=show_reg&""
	  End If	  
	  show_reg=show_reg&""
	  
	  If instr(FamilyType,"2")>0 Then
	  show_reg=show_reg&"<font color='#038ad7'>��</font>�����շѽӴ���ͥ"
	  Else
	  show_reg=show_reg&""
	  End If
	  show_reg=show_reg&""
	  
	  If instr(FamilyType,"3")>0 Then
	  show_reg=show_reg&"<font color='#038ad7'>��</font><font color='#038ad7'>����2008���˽Ӵ���ͥ</font>"
	  Else
	  show_reg=show_reg&""
	  End If
	  show_reg=show_reg&""
	  
	  
	  show_reg=show_reg&""
	  
	  
	  show_reg=show_reg&"<br />"
	

	
	
	
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ�ʺ���Ϣ</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ͥ����</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=63% colspan=2><strong>"&  Left(rsmain("blogname"),1) & vbcrlf 
	 
	 If Cint(rsmain("sex"))=1 Then show_reg=show_reg& "����"
	 If Cint(rsmain("sex"))=0 Then show_reg=show_reg& "Ůʿ"
	 If Cint(rsmain("sex"))=3 Then show_reg=show_reg& "����/Ůʿ"
    show_reg=show_reg&"</strong></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	

''	
''	show_reg=show_reg&"<tr> "& vbcrlf
''    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
''    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
''    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ע��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& rsmain("beizhu") &"</td>" 
    show_reg=show_reg&"</tr>"
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>���ڵ���(ʡ/��)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& rsmain("province")& " &nbsp;" &rsmain("city") &"</td>"
   
    show_reg=show_reg&"</tr>"

	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>��ѡ��ѧϰ��Χ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.requestID_selectName("oblog_userclass","classname","where id="& rsmain("user_classid")) &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�������ڣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"

	show_reg=show_reg& rsmain("birthday") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>����ְҵ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& rsmain("job") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>��ͥ��Ա��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ������"
	If Instr(rsmain("familynumber"),"1")>0 Then show_reg=show_reg&"1��"
	show_reg=show_reg&""
	If Instr(rsmain("familynumber"),"2")>0 Then show_reg=show_reg&"2��"
	show_reg=show_reg&""
	If Instr(rsmain("familynumber"),"3")>0 Then show_reg=show_reg&"3��"
	show_reg=show_reg&""
	If Instr(rsmain("familynumber"),"4")>0 Then show_reg=show_reg&"4��"
	show_reg=show_reg&""
	If Instr(rsmain("familynumber"),"����")>0 Then show_reg=show_reg&"���ࣨ����4����ͥ��Ա��"
	show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ�ṹ��"& rsmain("familymember") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
                           
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>���ݽ��ܣ�</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�ṹ��"& rsmain("houseframe")
                          show_reg=show_reg&""&_
						  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����"& rsmain("housespace") &_
						  "</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>�Ƿ��г��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"
	
	If Instr(rsmain("pet"),"û��")=0 Then show_reg=show_reg&"�г���"
	
	If Instr(rsmain("pet"),"û��")>0 Then show_reg=show_reg&"û��"
	
	If Instr(rsmain("pet"),"û��")=0 Then
		show_reg=show_reg& "&nbsp;&nbsp;&nbsp;("& rsmain("PetType") &")"
	End If
	show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                      
    
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'>����̵�����Ҫ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>����Ա�"& rsmain("asksex") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf              
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�Ƿ��˵���ģ�"& rsmain("issaychinese") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	                    

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>Ӣ��ˮƽ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& rsmain("englishlevel") 
	show_reg=show_reg&"&nbsp;&nbsp;&nbsp;&nbsp;<strong>�Ƿ���˽�ҳ�</strong>��"& rsmain("car") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>����������ͣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"
	show_reg=show_reg&""& rsmain("toilet") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�з���ԣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& rsmain("computer")
	show_reg=show_reg&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>�з���������</strong>"& rsmain("internet") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25 class='Form-td-left'><div align=left>�ͷ�����(��ע)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2 style='width:340px;'><br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&rsmain("remarkinfo")&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	

	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 class='Form-td-left'></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><!--<img src='/images/left-1-3.gif' style='padding-bottom:3px;' /> &nbsp;--><input type='button' name='closewl' value=' �� �� �� �� ' onclick='javascript:window.close();' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'> &nbsp;&nbsp;&nbsp;<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'> <input type=hidden name=action value='chkreg'><input type=hidden name=FamilyTypeRequest value='"& FamilyTypeRequest &"'>"& vbcrlf
    show_reg=show_reg&"</div></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"</table>"& vbcrlf
	show_reg=show_reg&"</form><br />"& vbcrlf
	

	'show_reg=show_reg&"</div>"
    
	end if
	'set rs=nothing
	'''response.Write(show_reg)
	
		''''''---------------------------------1.
		show_list=show_list&  show_reg  &"</td></tr>"& vbcrlf'ok;
		
	End If
	
	
	
	
	
	
		Dim rsPhoto,sqlPhoto,strTempPhoto
		Dim sqlTemp
		sqlTemp = " And user_level<=6 And user_level_team=0  And isUserPublicForView=1 And isCover=0 "'ѡ���Ƿ����������Ƭ��WL
		sqlPhoto="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto<>11111 "& sqlTemp &" And isPower<>11110 And oblog_upfile.userid="& id
		if not IsObject(conn) then link_database
		Set rsPhoto=Server.CreateObject("Adodb.RecordSet")
		'response.Write(msql)
		rsPhoto.Open sqlPhoto,Conn
		strTempPhoto=strTempPhoto  &"<br />"
			
		
		if rsPhoto.eof and rsPhoto.bof then
			strTempPhoto=strTempPhoto & ""
		else
			'''Dim totalPutPhoto
			'''totalPutPhoto= rsPhoto.recordcount
			'show_list=show_list & "������" & totalPutPhoto & " ����ͥ<br>"
			
				'''getlistPhoto()
				strTempPhoto ="�ҵļ�ͥ��Ἧ:<br /><hr style='color:#DE728A' /><br /><br />"
				Dim nn
				nn=0
				Do while not rsPhoto.Eof
					strTempPhoto= strTempPhoto &"<span onMouseOut=""this.style.backgroundColor=''""  onMouseOver=""this.style.backgroundColor='#BFE0FB'""><a href='"& rsPhoto("file_path") &"' target='_blank'>"
					strTempPhoto= strTempPhoto &"<img src='"& rsPhoto("file_path") &"' style=""margin-left:53px;width:130px;height:100px;border:0px;padding:3px; border: 1px #cccccc solid; padding:5px;  filter:progid:DXImageTransform.Microsoft.DropShadow(color=#eeeeee,offX=3,offY=4,positives=true);"" onerror=""this.src='/images/unknow_photo02.jpg';"" /></a></span>"
				nn = nn+1
				If nn = 3 Then strTempPhoto= strTempPhoto &"<br /><br />"
				If nn = 3 Then nn = 0
				rsPhoto.Movenext
				Loop
				
		end if
		
		rsPhoto.Close
		set rsPhoto=Nothing
				
		strTempPhoto= strTempPhoto  &"<br />"
		
		
		
		
		

		
		
	
	''''''---------------------------------2.
	show_list=show_list &"</table> <div style='position:relative;float:left;'>"& strTempPhoto &"</div>"
	'set fso=nothing
	
	
	
		
	
		
		
end sub

'��ȡϵͳ����
Function GetSysClasses()
	Dim rst,sReturn
	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value=0>���з���</option>" & VBCRLF & sReturn
		sReturn="<form name=photoform method=get><tr><td clospan=2>ѡ�������ࣺ<select name=classid onchange=""this.form.submit()"">" & VBCRLF & sReturn & "</select></tr></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>
