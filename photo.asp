<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
if en_photo=0 then
	oblog.adderrstr("此功能已被系统关闭！")
	oblog.showerr
	Set oblog = Nothing
end if
dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,strPlayerUrl
strurl="photo.asp"
isbest=cint(request.QueryString("isbest"))
classid=request.QueryString("classid")
call sysshow()

if isbest=1 then
	mainsql=mainsql&" and user_isbest=1"
	if strurl="photo.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="→推荐图片"
end if

If IsNumeric(classid) Then
	classid=CLNG(classid)
	If classid>0 Then 
		mainsql= mainsql & " and sysClassid=" & classid & " "
		if strurl="photo.asp" then
			strurl=strurl&"?classid=" & classid
		else
			strurl=strurl&"?classid=" & classid
		end if
	End IF
End If

strPlayerUrl= Replace(strurl,"photo.asp","photoplayer.asp")
		
call sub_showuserlist(mainsql,strurl)
show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom
sub sub_showuserlist(sql,strurl)
	dim topn,msql
	MaxPerPage=oblog.setup(79,0)
	strFileName=strurl
	if request("page")<>"" then
    	currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	msql="select top "&clng(oblog.setup(78,0))&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1 "&sql&" And isPower=0 order by fileid desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "共调用0个图片<br>"
	else
    	totalPut=rsmain.recordcount
		'show_list=show_list & "共调用" & totalPut & " 个图片<br>"
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
        	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
        	else
	        	currentPage=1
           		getlist()
           		show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
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
	show_list="<table width='100%' border='0'><tr><td>"
	show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→相册(共调用" & totalPut & " 个图片)"
	'bstr=trim(Request.ServerVariables("query_string"))
	'if bstr<>"" then bstr="photo.asp?"&replace(replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="photo.asp?isbest=1"	
	show_list=show_list&bstr1&"</td><td align='right'><a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')"">启用自动播放</a></td></tr>"
	show_list=show_list& GetSysClasses & "</table><hr>"
	
	show_list=show_list&"<table width='100%'  align='center' cellpadding='0' cellspacing='1'>"& vbcrlf
	do while not rsmain.eof
		show_list=show_list&"<tr height='22'>"& vbcrlf
		for n=1 to 3
			if rsmain.eof then
			show_list=show_list&"<td width='25%'></td>"& vbcrlf
			else
			title="图片说明:"&oblog.filt_html(rsmain(1))
			userurl="<a href='more.asp?id="& rsmain("logid") &"' target='_blank'>"
			imgsrc=rsmain(0)
			imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
			imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
			if  not fso.FileExists(Server.MapPath(imgsrc)) then
				imgsrc=rsmain(0)
			end if
			if rsmain(5)<>"" then  
				userurl=userurl&oblog.filt_html(rsmain(5))&"</a>"
			else
				userurl=userurl&oblog.filt_html(rsmain(4))&"</a>"
			end if
			show_list=show_list&"<td align='center'> <a href='"&rsmain(0)&"' title='"&title&"' target='_blank'><img src='"&imgsrc&"' height='100' width='130' border='0' /></a><br />来自:"&userurl&"</td>"& vbcrlf
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

'获取系统分类
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
		sReturn = "<option value=0>所有分类</option>" & VBCRLF & sReturn
		sReturn="<form name=photoform method=get><tr><td clospan=2>选择相册分类：<select name=classid onchange=""this.form.submit()"">" & VBCRLF & sReturn & "</select></tr></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>
