<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%

dim UserSearch

UserSearch=trim(request("UserSearch"))

if UserSearch="" then
	UserSearch=0
else
	UserSearch=Clng(UserSearch)
end if

'if en_photo=0 then
'	oblog.adderrstr("�˹����ѱ�ϵͳ�رգ�")
'	oblog.showerr
'	Set oblog = Nothing
'end if
dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,strPlayerUrl
strurl="HomestayFamilyChina.asp"
isbest=cint(request.QueryString("isbest"))
classid=request.QueryString("classid")
call sysshow_3()

if request("province")<>"" or request("city")<>"" or request("usertype")<>"" or request("houseframe")<>"" then
	if strurl="HomestayFamilyChina.asp" then
		strurl=strurl&"?province="& request("province") &"&city="& request("city") &"&usertype="& request("usertype") &"&houseframe="& request("houseframe") 
	else
		strurl=strurl&"&province=&city=&usertype=0&houseframe=&houseframe=2%CC%FC&houseframe="
	end if
end if

if isbest=1 then
	mainsql=mainsql&" and isBestUser=1"
	if strurl="HomestayFamilyChina.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="���Ƽ��Ӵ���ͥ"
end if




'If IsNumeric(classid) Then
'	classid=CLNG(classid)
'	If classid>0 Then 
'		mainsql= mainsql & " and sysClassid=" & classid & " "
'		if strurl="HomestayFamilyChina.asp" then
'			strurl=strurl&"?classid=" & classid
'		else
'			strurl=strurl&"&classid=" & classid
'		end if
'	End IF
'End If

strPlayerUrl= Replace(strurl,"HomestayFamilyChina.asp","photoplayer.asp")
		
call sub_showuserlist(mainsql,strurl)

show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom

sub sub_showuserlist(sql,strurl)
	dim topn,msql
	'MaxPerPage=oblog.setup(79,0)
	MaxPerPage=18
	strFileName=strurl'���ݺ���url��
	if request("page")<>"" then
    	currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	
	Dim strSQL_Search,userid
	If isNumeric(Replace( Trim(Request("SearchUserId")),"MH00","")) Then
		userid= Cint(Replace( Trim(Request("SearchUserId")),"MH00",""))
	Else
		userid= 0
	End If
	
	Select Case UserSearch
		Case 0
		strSQL_Search= ""
		If Request("province")="" And Request("city")="" And Request("usertype")="" And Request("houseframe")="" Then
			strSQL_Search= strSQL_Search & ""'û�п��ٲ�ѯʱ��
		Else
		'��ע�빦��WL[2007-09-23];�����ַ�������������
		Dim Request_province,Request_city,Request_houseframe
		Request_province = replace(Request("province"),"'","''")
		Request_city = replace(Request("city"),"'","''")
		Request_houseframe = replace(Request("houseframe"),"'","''")
		
		Request_province = replace(Request("province"),";","��")
		Request_city = replace(Request("city"),";","��")
		Request_houseframe = replace(Request("houseframe"),";","��")
		
		Request_province = replace(Request("province"),"(","��")
		Request_city = replace(Request("city"),"(","��")
		Request_houseframe = replace(Request("houseframe"),"(","��")
		
		Request_province = replace(Request("province"),")","��")
		Request_city = replace(Request("city"),")","��")
		Request_houseframe = replace(Request("houseframe"),")","��")
		
		
			If Request_province<>"" Then strSQL_Search= strSQL_Search &" And province='"& Request_province &"' "
			If Request_city<>"" And Request_city<>"ѡ��(*)" Then strSQL_Search= strSQL_Search &" And city='"& Request_city &"' "
			If Request("usertype")<>"" And Request("usertype")<>"0" Then strSQL_Search= strSQL_Search &" And user_classid="& Cint(Request("usertype")) &" "
			
			If Request_houseframe<>"" And Trim(Request_houseframe)<>", ," Then'���⴦��׼�ԣ�WL;
			Dim TempSql
			TempSql = split(Trim(Request_houseframe),",")'����TempSql(0);TempSql(1);TempSql(2)
				strSQL_Search= strSQL_Search &" And ( houseframe like '%"& TempSql(0) &"%' And houseframe like '%"& TempSql(1) &"%'  And houseframe like '%"& TempSql(2) &"%' ) "
			End If
		End If
		
		Case 10
		strSQL_Search= " And userid="& userid
		
	End Select
	'''RESPONSE.Write(strSQL_Search)
	
	'''msql="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1 "&sql&" And isPower=0 order by fileid desc"
	Dim sql001,sql002,sql003'��������WL;
	sql001 = " And user_level<=6 And user_level_team=0 "'ɸѡ��ͥ�û����ļ���
	'''sql002 = " And isUserPublicForView=1 And isCover=1 "'ɸѡ����ͥ�û���Ϊ����+������ļ���
	sql003 = " And isValidate=1 "'ɸѡΪ��ͨ��վ����˵ļ�ͥ�û���
	'''msql="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto<>11111 "&sql&" And isPower<>11110  "& sql001 & sql002 & sql003 &" order by fileid desc"
	'msql="Select Top "&clng(oblog.setup(78,0))&" * from [oblog_user] Where userid<>0"& sql001 & sql002 & sql003 & strSQL_Search &" order by isStop asc,userid desc"
	
	

	msql="Select Top 500 * From [oblog_user] Where userid<>0 And isStop=0"& sql001 & sql002 & sql003 & strSQL_Search &" Order By isCoverPhoto Desc,isStop Asc,userid Desc"
	
	
	
	
	
	
	
	
	
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		
		If Trim(Request("SearchUserId"))="" Then'��û��ֱ�Ӳ�ѯ��ŵ�ʱ�򣬶���ʾ��WL;			
			If Request("province")="" And Request("city")="" And Request("usertype")="" And Request("houseframe")="" Then
				show_list=show_list&"</table><hr>"
			Else
				show_list=show_list&"&nbsp;��ǰ�Ӵ���ͥ��&nbsp;&nbsp;"
				'''response.Write Request("province")
				show_list=show_list&"<font color=#038ad7>"
				If Request("province")<>"" Or Request("city")<>"" Then show_list=show_list& "���ڵ���:&nbsp;<font color=#F7498C>��</font>"& Request("province")
				If Request("city")<>"" Then show_list=show_list& "&nbsp;"& Request("city")
				
					If Request("usertype")<>"" And Request("usertype")<>"0" And ( Request("province")<>"" Or Request("city")<>"") Then show_list=show_list& "&nbsp;---"
				If Request("usertype")<>"" And Request("usertype")<>"0" Then show_list=show_list& "&nbsp;����:&nbsp;<font color=#F7498C>��</font>"& oblog.requestID_selectName("oblog_userclass","classname","where id="& Cint(Request("usertype")))
				
					If Request("houseframe")<>"" And Trim(Request("houseframe"))<>", ," And ((Request("usertype")<>"" And Request("usertype")<>"0") Or (Request("province")<>"" Or Request("city")<>"")) Then show_list=show_list& "&nbsp;---"
					
				If Request("houseframe")<>"" And Trim(Request("houseframe"))<>", ," Then show_list=show_list& "&nbsp;���ݽṹ:&nbsp;<font color=#F7498C>��</font>"& Replace(Request("houseframe"),",","&nbsp;")
				show_list=show_list&"</font>"
				show_list=show_list&"</table><hr>"
			End If
		End If
		
		
		show_list=show_list & "<font color=#BE0000>û���ҵ��Ӵ���ͥ</font>"
		
		If Trim(Request("SearchUserId"))<>"" Then
			show_list=show_list & "&nbsp;""<font color=#BE0000>"& Trim(Request("SearchUserId")) &"</font>"""
		End If
		
		
		
		show_list=show_list & "<br>"
	else
	
		If Trim(Request("SearchUserId"))<>"" Then
			show_list=show_list & "&nbsp;�Ѳ�ѯ�� �� <font color=#F7498C>��</font>�Ӵ���ͥ&nbsp;&nbsp;""<font color=#BE0000>"& Trim(Request("SearchUserId")) &"</font>""<br>"
		End If
		
    	totalPut=rsmain.recordcount
		'show_list=show_list & "������" & totalPut & " ����ͥ<br>"
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
        	getlist()
        	show_list=show_list&"<center>"&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")&"</center>"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&"<center>"&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")&"</center>"
        	else
	        	currentPage=1
           		getlist()
           		show_list=show_list&"<center>"&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"���Ӵ���ͥ")&"</center>"
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,bstr,n,fso
	dim title,userurl,imgsrc
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	show_list=show_list & "<table width='100%' border='0' style='font-size:13px;'><tr><td>"
	show_list=show_list&"��ǰλ�ã�<a href='index.asp'>��ҳ</a> �� <a href='HomestayFamilyChina.asp'>�й��Ӵ���ͥ</a> (��" & totalPut & " ���Ӵ���ͥ)"
	'bstr=trim(Request.ServerVariables("query_string"))
	'if bstr<>"" then bstr="HomestayFamilyChina.asp?"&replace(replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="HomestayFamilyChina.asp?isbest=1"	
	show_list=show_list&bstr1&"</td><td align='right'><a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')""></a></td></tr>"
	show_list=show_list&  ""
	
	If Request("province")="" And Request("city")="" And Request("usertype")="" And Request("houseframe")="" Then
		show_list=show_list&"</table><hr>"
	Else
		show_list=show_list&"&nbsp;��ǰ�Ӵ���ͥ��ѯ&nbsp;&nbsp;"
		'''response.Write Request("province")
		show_list=show_list&"<font color=#038ad7>"
		If Request("province")<>"" Or Request("city")<>"" Then show_list=show_list& "���ڵ���:&nbsp;<font color=#F7498C>��</font>"& Request("province")
		If Request("city")<>"" Then show_list=show_list& "&nbsp;"& Request("city")
		
			If Request("usertype")<>"" And Request("usertype")<>"0" And ( Request("province")<>"" Or Request("city")<>"") Then show_list=show_list& "&nbsp;---"
		If Request("usertype")<>"" And Request("usertype")<>"0" Then show_list=show_list& "&nbsp;����:&nbsp;<font color=#F7498C>��</font>"& oblog.requestID_selectName("oblog_userclass","classname","where id="& Cint(Request("usertype")))
		
			If Request("houseframe")<>"" And Trim(Request("houseframe"))<>", ," And ((Request("usertype")<>"" And Request("usertype")<>"0") Or (Request("province")<>"" Or Request("city")<>"")) Then show_list=show_list& "&nbsp;---"
			
		If Request("houseframe")<>"" And Trim(Request("houseframe"))<>", ," Then show_list=show_list& "&nbsp;���ݽṹ:&nbsp;<font color=#F7498C>��</font>"& Replace(Request("houseframe"),",","&nbsp;")
		show_list=show_list&"</font>"
		show_list=show_list&"</table><hr>"
	End If
	
	
	show_list=show_list&"<table width='100%'  align='center' cellpadding='0' cellspacing='1' style='position:relative;width:680px;font:13px;'>"& vbcrlf
	Dim titleTemp'WL�¼ӣ�
	do while not rsmain.eof
	

		
		show_list=show_list&"<tr height='22'>"& vbcrlf
		for n=1 to 3
			if rsmain.eof then
				show_list=show_list&"<td width=''></td>"& vbcrlf
			else
			
			titleTemp="======== ��ͥ��Ϣ ========" & vbcrlf & "�Ա�"		
				titleTemp=titleTemp& vbcrlf & "���ݽ��ܣ�"& rsmain("houseframe")
				titleTemp=titleTemp& vbcrlf & "�����"& rsmain("housespace")
				titleTemp=titleTemp& vbcrlf & "�з���"
					If Instr(rsmain("pet"),"û��")=0 Then titleTemp=titleTemp&"�г���"
			
					If Instr(rsmain("pet"),"û��")>0 Then titleTemp=titleTemp&"û��"
					
					If Instr(rsmain("pet"),"û��")=0 Then
						titleTemp=titleTemp& "&nbsp;&nbsp;&nbsp;("& rsmain("PetType") &")"
					End If
				titleTemp=titleTemp& vbcrlf & "�з���ԣ�"& rsmain("computer")
				titleTemp=titleTemp& vbcrlf & "���������"& rsmain("internet")		
			titleTemp=titleTemp& vbcrlf & "��ʼ���ڣ�" & rsmain("adddate")
			
			'''title="ͼƬ˵��:"&oblog.filt_html(rsmain("userid"))
			userurl="<a href='more.asp?id="& rsmain("userid") &"' target='_blank'>"
			If rsmain("upfile_isCoverPhoto_path")<>"" And Len(rsmain("upfile_isCoverPhoto_path"))>6 Then
			imgsrc=rsmain("upfile_isCoverPhoto_path")'ͼƬ��ַWL��
			Else
			imgsrc="/images/unknow_photo02.jpg"'ͼƬ��ַWL��
			End If			
			imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
			imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),""&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
			'''imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
			if  not fso.FileExists(Server.MapPath(imgsrc)) then
				imgsrc= imgsrc'ͼƬ��ַWL;
			end if
		'''RESPONSE.Write(imgsrc)
			if rsmain(5)<>"" then  
				userurl=userurl&oblog.filt_html(rsmain(5))&"</a>"
			else
				userurl=userurl&oblog.filt_html(rsmain(4))&"</a>"
			end if
			show_list=show_list& "<td align='left' style=''>"
				show_list=show_list& "<span class='photoLeft' style='float:left;padding:3px;' onMouseOut=""this.style.backgroundColor=''""  onMouseOver=""this.style.backgroundColor='#87D2F1'"">"
				
				If imgsrc = rsmain("upfile_isCoverPhoto_path") Then show_list=show_list& "<a href='"& rsmain("upfile_isCoverPhoto_path") &"' title='"&title&"' target='_blank'>"
					show_list=show_list& "<img style=""border: 1px #cccccc solid; padding:5px; filter : progid:DXImageTransform.Microsoft.DropShadow(color=#eeeeee,offX=3,offY=3,positives=true);"" alt="""& titleTemp &""" src='"&imgsrc&"' height='100' width='130' border='0' onerror=""this.src='/images/unknow_photo02.jpg';"" />"
				If imgsrc = rsmain("upfile_isCoverPhoto_path") Then show_list=show_list& "</a>"
				
				show_list=show_list& "<br />"
				show_list=show_list& "&nbsp;&nbsp;&nbsp;����:<span style='padding:3px;color:#333333;font-weight:bold;'>"&rsmain("province")&" "&rsmain("city")&"</span>"
				show_list=show_list& "</span><br />"
				
				show_list=show_list& "<span class='photoRight' style='float:left111;clear:both;'>"
					show_list=show_list& "<ul style='list-style-type:none;'>"
					
					show_list=show_list& "<li style='padding-left:0px;' title='��ͥ��ţ�MH00"&rsmain("userid")&"' style='cursor:hand;'><font color=#E93D7D>��ͥ��ţ�MH00"& rsmain("userid")
					show_list=show_list& "</font></li>"
					
					show_list=show_list& "<li style='padding-left:0px;' title='��ͥ����"& Left(rsmain("blogname"),1)  	 
	 If Cint(rsmain("sex"))=1 Then show_list=show_list& "����"
	 If Cint(rsmain("sex"))=0 Then show_list=show_list& "Ůʿ"
	 If Cint(rsmain("sex"))=3 Then show_list=show_list& "����/Ůʿ" 
	 				show_list=show_list& "' style='cursor:hand;'><font color=#E93D7D>��ͥ����"
					show_list=show_list& "</font><strong>"&  Left(rsmain("blogname"),1) & vbcrlf 	 
	 If Cint(rsmain("sex"))=1 Then show_list=show_list& "����"
	 If Cint(rsmain("sex"))=0 Then show_list=show_list& "Ůʿ"
	 If Cint(rsmain("sex"))=3 Then show_list=show_list& "����/Ůʿ"
    show_list=show_list&"</strong></li>"
					
					show_list=show_list& "<li style='padding-left:0px;' title='��ע��"& rsmain("beizhu") &"' style='cursor:hand;'>��ע��"& rsmain("beizhu")
					show_list=show_list& "</li>"
					show_list=show_list& "<li style='padding-left:0px;' title='���ݽ��ܣ�"&rsmain("houseframe")&"' style='cursor:hand;'>���ݽ��ܣ�"& rsmain("houseframe")
					show_list=show_list& "</li>"
					show_list=show_list& "<li style='padding-left:0px;'>"
					If rsmain("isStop")=1 Then
						show_list=show_list& "�Ӵ�״̬��<strong>�ѽӴ�</strong>��"
					Else
						show_list=show_list& "���������"& rsmain("internet")
					End If
					show_list=show_list& "</li>"
					show_list=show_list& "<li style='padding-left:0px;'>��ϸ��Ϣ��<img src='/images/left-1-3.gif' /> <a href='HomestayFamilyMember.asp?id="&rsmain("userid")&"' title='�鿴��ͥ��ϸ����'><font color='#038ad7'>�������</a>"
					show_list=show_list& "</li>"
					
					show_list=show_list& "</ul>"
				show_list=show_list& "</span>"
				
				show_list=show_list& "<span class='photoClear' style='clear:both;'>&nbsp;</span>"& vbcrlf
			show_list=show_list& "</td>"& vbcrlf
			i=i+1
			if not rsmain.eof then rsmain.movenext
			end if
		next
		show_list=show_list&"</tr>"& vbcrlf
		if i>=MaxPerPage then exit do	
	loop		
	show_list=show_list&"</table>"
	set fso=nothing
end sub

'��ȡϵͳ����
Function GetSysClasses123()
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

''��ȡ���ڵ���(ʡ/��)����ѧϰ�������ࡪ�����ݽṹ������ͥ��Ų�ѯ��
'Function GetSysClassesSearchHome()
'	Dim strType_city
'	Dim str_usertype
'	Dim rst,sReturn
'	
'	Dim Request_province,Request_city
'	Request_province= Request("province")
'	Request_city= Request("city")
'	
'	Dim Request_usertype
'	Request_usertype= Cint(Request("usertype"))
'	
'	Dim Request_houseframe
'	Request_houseframe= Trim(Request("houseframe"))
'	
'	strType_city =strType_city & "<Div id='line01' style='margin:0px;padding-top:3px;'>&nbsp;���ڵ���&nbsp;" &oblog.type_citySelectSubmit(Request_province,Request_city )'����ѡ�������б�
'	
'	
'	str_usertype="<select name=usertype id=usertype onchange='submit();'>"
'	str_usertype=str_usertype&oblog.show_class("user",Request_usertype,0)
'    str_usertype=str_usertype&"</select>"	
'	str_usertype = "&nbsp;�����ֲ�ѯ&nbsp;" &str_usertype&"</Div>"'����ѡ�������б�
'	
''	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
''	If rst.Eof Then
''		sReturn=""
''	Else
''		Do While Not rst.Eof
''			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
''			rst.Movenext
''		Loop
''		sReturn = "<option value=0>���з���</option>" & VBCRLF & sReturn
'		sReturn="<form name=selectsubmitform action='HomestayFamilyChina.asp' method=get><tr><td clospan=2>" & strType_city & str_usertype & "<Div id='line02' style='margin:0px;padding-top:6px;'>&nbsp;���ݽṹ&nbsp;<select name='houseframe' id='houseframe1' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
'                            "<option value='' selected>��</option>"&_
'							"<option value='1��'"
'							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">1��</option>"&_
'                            "<option value='2��'"
'							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">2��</option>"&_
'                            "<option value='3��'"
'							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">3��</option>"&_
'                            "<option value='4��'"
'							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">4��</option>"&_
'                            "<option value='5��'"
'							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">5��</option>"&_
'                            "<option value='6��'"
'							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">6��</option>"&_
'                          "</select>&nbsp;<select name='houseframe' id='houseframe2' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
'						  "<option value='' selected>��</option>"&_
'                            "<option value='1��'"
'							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">1��</option>"&_
'                            "<option value='2��'"
'							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">2��</option>"&_
'                            "<option value='3��'"
'							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">3��</option>"&_
'                            "<option value='4��'"
'							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">4��</option>"&_
'                            "<option value='5��'"
'							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">5��</option>"&_
'                            "<option value='6��'"
'							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">6��</option>"&_
'                          "</select>&nbsp;<select name='houseframe' id='houseframe3' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
'						  "<option value='' selected>��</option>"&_
'                            "<option value='1��'"
'							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">1��</option>"&_
'                            "<option value='2��'"
'							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">2��</option>"&_
'                            "<option value='3��'"
'							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">3��</option>"&_
'                            "<option value='4��'"
'							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">4��</option>"&_
'                            "<option value='5��'"
'							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">5��</option>"&_
'                            "<option value='6��'"
'							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
'							sReturn=sReturn & ">6��</option>"&_
'                          "</select></Div> </tr></form>"
''	End If
''	rst.Close
''	Set rst=Nothing
'	GetSysClassesSearchHome = sReturn
'End Function

%>
