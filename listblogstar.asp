<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim mainsql,usertype,strurl,rsmain,show_list
strurl="listblogstar.asp"
call sysshow()

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
	msql="select oblog_blogstar.*,oblog_user.*,oblog_user.blogname As blogname from [oblog_blogstar],oblog_user where oblog_blogstar.userid= oblog_user.userid And ispass=1 order by addtime desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "������0�����Ǽ�֮ͥ��<br>"
	else
    	totalPut=rsmain.recordcount
		'show_list=show_list & "������" & totalPut & " �����Ǽ�ͥ<br>"
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
        	show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"�����Ǽ�ͥ") &"</center>" 
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"�����Ǽ�ͥ") &"</center>"
        	else
	        	currentPage=1
           		getlist()
           		show_list=show_list &"<center>"& oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"�����Ǽ�ͥ") &"</center>"
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim rstmp,i
	show_list="<div style='font-size:13px;'>��ǰ���ã�<a href='index.asp'>��ҳ</a>���������Ǽ�ͥ(����" & totalPut & "��) �����˹�����ʱ������...</div><hr />"
	show_list=show_list&"<table style='font-size:13px;' width='100%'  align='center' cellpadding='0' cellspacing='1'>"& vbcrlf
	show_list=show_list&"<tr height='25' ><td width='170' align='center'><strong>������Ƭ</strong></td>"& vbcrlf
	show_list=show_list&"<td align='left' width=80  style='margin:5px;'> <strong>�����ƺ�</strong></td>"& vbcrlf
	show_list=show_list&"<td align='center'> <strong>���</strong></td>"& vbcrlf
	show_list=show_list&"</tr>"& vbcrlf
    do while not rsmain.eof	
		'''show_list=show_list&"<tr><td align='center'> <a href='"&rsmain("userurl")&"' target='_blank'><img style='width:130px;height:100px;' src="""&rsmain("picurl")&"""  hspace=""3"" border=""0"" vspace=""3"" alt='"&oblog.filt_html(rsmain("blogname"))&"' /></a></td>"& vbcrlf
		
		show_list=show_list&"<tr><td align='center'> <a href='HomestayFamilyMember.asp?id="& rsmain("userid") &"' target='_blank'><img style='width:130px;height:100px;' src="""&rsmain("picurl")&"""  hspace=""3"" border=""0"" vspace=""3"" alt='"& Left(rsmain("blogname"),1)
			If Cint(rsmain("sex"))=1 Then show_list=show_list& "����"
			 If Cint(rsmain("sex"))=0 Then show_list=show_list& "Ůʿ"
			 If Cint(rsmain("sex"))=3 Then show_list=show_list& "����/Ůʿ"
		
		show_list=show_list& "' /></a></td>"& vbcrlf
		
		'''show_list=show_list&"<td align='left'><a href='"&rsmain("userurl")&"' target='_blank'>"& Left(rsmain("blogname"),1)
		show_list=show_list&"<td align='left'><a href='HomestayFamilyMember.asp?id="& rsmain("userid") &"' target='_blank'>"& Left(rsmain("blogname"),1)
			If Cint(rsmain("sex"))=1 Then show_list=show_list& "����"
			 If Cint(rsmain("sex"))=0 Then show_list=show_list& "Ůʿ"
			 If Cint(rsmain("sex"))=3 Then show_list=show_list& "����/Ůʿ"
		
		show_list=show_list& "</a></td>"& vbcrlf
		show_list=show_list&"<td align='left'>"&oblog.filt_html(rsmain("info"))&rsmain("remarkinfo")&"</td>"& vbcrlf
		show_list=show_list&"</tr>"& vbcrlf
		rsmain.movenext
		i=i+1
		if i>=MaxPerPage then exit do	
	loop		
	show_list=show_list&"</table>"
end sub
%>
