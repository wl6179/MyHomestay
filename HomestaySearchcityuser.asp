<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim mainsql,usertype,strurl,rsmain,province,city,bstr1,isbest,show_list,ustr
strurl="HomestaySearchcityuser.asp"
usertype=cint(request.QueryString("usertype"))
isbest=cint(request.QueryString("isbest"))
province=oblog.filt_badstr(request("province"))
city=oblog.filt_badstr(request("city"))
call sysshow()

if usertype>0 then
	set rsmain=oblog.execute("select id from oblog_userclass where parentpath like '"&usertype&",%' OR parentpath like '%,"&usertype&"' OR parentpath like '%,"&usertype&",%'")
	while not rsmain.eof
		ustr=ustr&","&rsmain(0)
		rsmain.movenext
	wend
	ustr=usertype&ustr
	mainsql=" and oblog_user.user_classid in ("&ustr&")"
	strurl="HomestaySearchcityuser.asp?usertype="&usertype
	'mainsql="and user_classid="&usertype
else
	mainsql=""
end if
if province<>"" then
	strurl=strurl&"?province="&province
	mainsql=mainsql&" and province='"&province&"'"
end if
if city<>"" then
	strurl=strurl&"&city="&city
	mainsql=mainsql&" and city='"&city&"'"
end if
if isbest=1 then
	mainsql=mainsql&" and user_isbest=1"
	if strurl="HomestaySearchcityuser.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="���Ƽ��û�"
end if

call sub_showuserlist(mainsql,strurl)
show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql
	MaxPerPage=oblog.setup(45,0)
	strFileName=strurl
	if request("page")<>"" then
    	currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	
	
	Dim sql001,sql002,sql003'��������WL;
	sql001 = " And user_level<=6 And user_level_team=0 "'ɸѡΪ��ͥ�û���

	sql003 = " And isValidate=1 "'ɸѡΪ��ͨ��վ����˵ļ�ͥ�û���
	
	msql="select username,blogname,sex,useremail,qq,msn,log_count,homepage,adddate,userid,province,city,user_classid,houseframe,housespace,pet,PetType,computer,internet from [oblog_user] where lockuser=0 "& sql & sql001 & sql002 & sql003 &" order by userid desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "û���ҵ���Ҫ��ѯ�ļ�ͥ......<br>"
	else
    	totalPut=rsmain.recordcount
		'show_list=show_list & "������" & totalPut & " λ�û�<br>"
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
        	show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"����ͥ") &"</center>"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"����ͥ") &"</center>"
        	else
	        	currentPage=1
           		getlist()
           		show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"����ͥ") &"</center>"
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim rstmp,i,bstr
	dim title,userurl
	set rstmp=conn.execute("select classname from oblog_userclass where id="&usertype)
	show_list="<table width='100%' border='0' style='font-size:13px;'><tr><td>"
	if not rstmp.eof then
		show_list=show_list&"��ǰλ�ã�<a href='index.asp'>��ҳ</a>����ͥ���("&rstmp(0)&")"
	end if
	if usertype=0 then
		show_list=show_list&"��ǰλ�ã�<a href='index.asp'>��ҳ</a>���Ӵ���ͥ&nbsp;(���ҵ�" & totalPut & "��ͥ)"
	end if
	bstr=trim(Request.ServerVariables("query_string"))
	if bstr<>"" then bstr="HomestaySearchcityuser.asp?"&replace(replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="HomestaySearchcityuser.asp?isbest=1"
	show_list=show_list&bstr1&"</td><td align='right'><a href='"&bstr&"'><font color=#038ad7>"& province &""& city &"���Ƽ��Ӵ���ͥ</font></a></td></tr></table><hr>"
		set rstmp=nothing
			show_list=show_list&"<table style='font-size:13px;' width='100%'  align='center' cellpadding='0' cellspacing='1'>"& vbcrlf
			show_list=show_list&"<tr height='25' ><td width='100' ><strong>��ͥ���</strong></td>"& vbcrlf
			show_list=show_list&"<td> <strong>�����ƺ�</strong></td>"& vbcrlf
			show_list=show_list&"<td width='100'> <strong>���ڵ���(ʡ/��)</strong></td>"& vbcrlf
			show_list=show_list&"<td width='130' style='padding-left:13px;'> <strong>ѧϰ��������</strong></td>"& vbcrlf
			show_list=show_list&"</tr>"& vbcrlf
     do while not rsmain.eof	
			title="======== ��ͥ��Ϣ ========" & vbcrlf & "�Ա�"
		
		
			title=title& vbcrlf & "���ݽ��ܣ�"& rsmain("houseframe")
			title=title& vbcrlf & "�����"& rsmain("housespace")
			title=title& vbcrlf & "�з���"
				If Instr(rsmain("pet"),"û��")=0 Then title=title&"�г���"
		
				If Instr(rsmain("pet"),"û��")>0 Then title=title&"û��"
				
				If Instr(rsmain("pet"),"û��")=0 Then
					title=title& "&nbsp;&nbsp;&nbsp;("& rsmain("PetType") &")"
				End If
			title=title& vbcrlf & "�з���ԣ�"& rsmain("computer")
			title=title& vbcrlf & "���п����"& rsmain("internet")
		
		title=title& vbcrlf & "��ʼ���ڣ�" & rsmain(8)
		show_list=show_list&"<tr><td height='30'><a href='HomestayFamilyMember.asp?id=" &rsmain("userid")&"' title='����鿴��ϸ����"& vbcrlf&title&"'><font color=#E93D7D>MH00"& rsmain("userid") &"</font></a></td>"& vbcrlf
		show_list=show_list&"<td height='30'>"& Left(rsmain("blogname"),1) 
			If Cint(rsmain("sex"))=1 Then show_list=show_list& "����"
			 If Cint(rsmain("sex"))=0 Then show_list=show_list& "Ůʿ"
			 If Cint(rsmain("sex"))=3 Then show_list=show_list& "����/Ůʿ"
		show_list=show_list&"</td>"& vbcrlf
		show_list=show_list&"<td height='30'>"&rsmain(10) &"&nbsp;&nbsp;"& rsmain(11)&"</td>"& vbcrlf
		show_list=show_list&"<td height='30' style='padding-left:13px;'>"& oblog.requestID_selectName("oblog_userclass","classname","where id="& rsmain("user_classid")) &"</td>"& vbcrlf
		show_list=show_list&"</tr>"& vbcrlf
		rsmain.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop		
	show_list=show_list&"</table>"
	

end sub
%>
