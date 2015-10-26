<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
Dim show_listTemp'WL;

dim mainsql,classid,strurl,rsmain,show_list,ustr
dim keyword,selecttype,isbest,bstr1,userid
classid=request.QueryString("classid")
If classid="" then
	classid =0
Else
	classid = Cint(classid)
End If
'''Response.Write classid
keyword=EncodeJP(oblog.filt_badstr(request("keyword")))
selecttype=oblog.filt_badstr(request("selecttype"))
isbest=clng(request.QueryString("isbest"))
userid=clng(request.QueryString("userid"))
call sysshow()
strurl="HomestayNav.asp"
if classid<>0 then
	set rsmain=oblog.execute("select id from oblog_logclass where parentpath like '"&classid&",%' OR parentpath like '%,"&classid&"' OR parentpath like '%,"&classid&",%'")
	while not rsmain.eof
		ustr=ustr&","&rsmain(0)
		rsmain.movenext
	wend
	ustr=classid&ustr
	mainsql=" and oblog_log.classid in ("&ustr&")"
	'mainsql=" and oblog_log.classid="&classid
	strurl="HomestayNav.asp?classid="&classid
end if
if keyword<>"" then
	select case selecttype
	case "topic"
		mainsql=" and topic like '%"&keyword&"%'"
		strurl="HomestayNav.asp?keyword="&keyword&"&selecttype="&selecttype
	case "logtext"
		if oblog.setup(77,0)=1 then
		mainsql=" and logtext like '%"&keyword&"%'"
		strurl="HomestayNav.asp?keyword="&keyword&"&selecttype="&selecttype
		else
		oblog.adderrstr("当前系统已经关闭资讯内容搜索。")
		oblog.showerr
		end if
	case "id"
		mainsql=" and author like '%"&keyword&"%'"
		strurl="HomestayNav.asp?keyword="&keyword&"&selecttype="&selecttype
	end select
end if
if isbest=1 then
	mainsql=mainsql&" and isbest=1"
	if strurl="HomestayNav.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="→<font color=#F7498C>推荐资讯</font>"
end if
if userid>0 then
	mainsql=mainsql&" and userid="&userid
	if strurl="HomestayNav.asp" then
		strurl=strurl&"?userid="&userid
	else
		strurl=strurl&"&userid="&userid
	end if
end if
call sub_showlist(mainsql,strurl)
show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom

sub sub_showlist(sql,strurl)
	dim topn
	dim msql
	MaxPerPage=oblog.setup(8,0)
	strFileName=strurl
	if request("page")<>"" then
    	currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	topn=oblog.setup(14,0)
	'''msql="select top "&topn&" oblog_log.topic,oblog_log.author,oblog_log.addtime,oblog_log.commentnum,oblog_log.logid,oblog_logclass.classname,oblog_logclass.id,userid,logfile from oblog_log,oblog_logclass where oblog_log.classid=oblog_logclass.id and ishide=0 and passcheck=1 and isdraft=0 and blog_password=0"&sql
	msql="select top "&topn&" oblog_log.topic,oblog_log.author,oblog_log.addtime,oblog_log.commentnum,oblog_log.logid,oblog_logclass.classname,oblog_logclass.id,userid,logfile from oblog_log,oblog_logclass where oblog_log.classid=oblog_logclass.classid and ishide=0 and passcheck=1 and isdraft=0 and blog_password=0"&sql
	
	msql=msql&" order by oblog_log.logid desc"
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'''response.Write(msql)
	if not IsObject(conn) then link_database
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "&nbsp;&nbsp;当前共有0篇资讯<br>"
	else
    	totalPut=rsmain.recordcount
		
		show_listTemp= show_listTemp & "&nbsp;&nbsp;共有" & totalPut & " 篇资讯"
		'WL;
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
			show_list=show_list&"<table style='font-size:13px;' align='center'><tr><td align='center'>"
        	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"篇资讯")
			show_list=show_list&"</td></tr></table >"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
				show_list=show_list&"<table style='font-size:13px;' align='center'><tr><td align='center'>"
            	show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"篇资讯")
				show_list=show_list&"</td></tr></table >"
        	else
	        	currentPage=1
           		getlist()
				show_list=show_list&"<table style='font-size:13px;' align='center'><tr><td align='center'>"
           		show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"篇资讯")
				show_list=show_list&"</td></tr></table >"
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,strtopic,userurl,bstr
	'show_list="<table width='100%' border='0'><tr><td>"
		show_list="<TABLE style='font-size:13px;' width='440' border=0 cellPadding=0  cellSpacing=0>"
		  
		  show_list=show_list&"<TR>"
			show_list=show_list&"<TD vAlign=top>"
			show_list=show_list&"<table style='font-size:13px;' width='100%' border='0' cellpadding='0' cellspacing='0'>"
			  show_list=show_list&"<tr>"
				show_list=show_list&"<td height='28' background='images/blog/bloglist_r_1.gif123'>"
				
				show_list=show_list&"<table style='font-size:13px;' width='100%' border='0' cellpadding='0' cellspacing='0'>"
				  show_list=show_list&"<tr>"
					show_list=show_list&"<td width='21'>&nbsp;</td>"
					show_list=show_list&"<td valign='bottom' id='bloglisttop'>"
					
					
	if keyword="" then
		if classid<>0 then
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a> &gt;&gt; <a href='"
				Dim strurlTemp
				If Instr(strurl,"&isbest=1")>0 Then
					strurlTemp=Replace(strurl,"&isbest=1","")
				Else
					strurlTemp= strurl
				End If
				''strurlTemp =strurl
				'''response.Write strurlTemp
			show_list=show_list& strurlTemp &"'>"&rsmain(5)&"</a>&nbsp;"
		else
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a> &gt;&gt; 所有资讯(所有栏目)&nbsp;"	
		end if
	else
		select case selecttype
		case "topic"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索资讯标题关键字“"&keyword&"”&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		case "logtext"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索资讯内容关键字“"&keyword&"”&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		case "id"
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索用户名称关键字“"&keyword&"”&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		case else
			show_list=show_list&"当前位置：<a href='index.asp'>首页</a>→搜索关键字“"&keyword&"”&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		end select
	end if
	
	show_list= show_list & show_listTemp'WL;
	
	bstr=trim(Request.ServerVariables("query_string"))
	if bstr<>"" then bstr="HomestayNav.asp?"&replace(replace(bstr,"&isbest=1",""),"isbest=1","")&"&isbest=1" else bstr="HomestayNav.asp?isbest=1"
	show_list=show_list&bstr1&"&nbsp;&nbsp;&nbsp;<a href='"&bstr&"'>非常推荐</a>"
		show_list=show_list&"</td>"
		
            show_list=show_list&"<td align='right'>&nbsp;</td>"
          show_list=show_list&"</tr>"
        show_list=show_list&"</table>"
		
		show_list=show_list&"</td>"
      show_list=show_list&"</tr>"
      show_list=show_list&"<tr>"
        show_list=show_list&"<td background='images/blog/bloglist_bg_right.gif123' id='bloglist_tdhcenter' style='padding:10px'>"
		show_list=show_list&"<TABLE style='font-size:13px;' width='100%' border='0' cellpadding='0' cellspacing='0' id='bloglist_tdhcenter'>"
          show_list=show_list&"<TBODY>"
	
	'show_list=show_list&bstr1&"</td><td align='right'><a href='"&bstr&"'>查看精华资讯</a></td></tr></table><hr>"
	'show_list=show_list&"<table width='100%' border='0'><tr>"
	'show_list=show_list&"<td><strong>资讯标题</strong></td><td width='100' align='center'><strong>作者</strong></td><td width='60' align='center'><strong>日期</strong></td><td width='40' align='center'><strong>评论</strong></td></tr>"
	show_list=show_list&"<TR>"
	show_list=show_list&"<TD></TD>"
	show_list=show_list&""
	show_list=show_list&""
	show_list=show_list&""
	show_list=show_list&"</TR>"
		do while not rsmain.eof
		strtopic=oblog.filt_html(rsmain(0))
		if oblog.strLength(strtopic)>50 then
			strtopic=oblog.InterceptStr(strtopic,47)&"..."
		end if
				show_list=show_list&"<TR>"
				show_list=show_list&"<TD style='height:28px;'><span style='width:30px;background:url(/images/left-1-3.gif) no-repeat 12px 5px; padding-left:13px; margin-right:5px;'>&nbsp;</span> <A title='"&oblog.filt_html(rsmain(0))&"' href='"&rsmain(8)&"' target='_blank'>"&strtopic&"</A></TD>"
				show_list=show_list&""
				show_list=show_list&""
				show_list=show_list&""
				show_list=show_list&"</TR>"
			
'			show_list=show_list&"<tr><td><a href='"&rsmain(8)&"' title='"&oblog.filt_html(rsmain(0))&"' target=_blank>"&strtopic&"</a></td>"
'			show_list=show_list&"<td align='center'><a href='go.asp?userid="&rsmain(7)&"' target=_blank>"&oblog.filt_html(rsmain(1))&"</a></td>"
'			show_list=show_list&"<td align='center'>"&mid(formatdatetime(rsmain(2),2),6)&"</td>"
'			show_list=show_list&"<td align='center'>"&rsmain(3)&"</td></tr>"
		rsmain.movenext
		i=i+1
		if i>=MaxPerPage then exit do
		loop
	'show_list=show_list&"</table>"
	          show_list=show_list&"</TBODY>"
        show_list=show_list&"</TABLE>"
          show_list=show_list&"<table style='font-size:13px;' width='100%' border='0' cellpadding='5'>"
            show_list=show_list&"<tr>"
              show_list=show_list&"<td height='0' align='center' bgcolor='#'></td>"
            show_list=show_list&"</tr>"

          show_list=show_list&"</table>"
          show_list=show_list&"</td>"
      show_list=show_list&"</tr>"
      show_list=show_list&"<tr>"
        show_list=show_list&"<td><!--<img src='/images/3bloglist_center_1.gif' height='15'>--></td>"
      show_list=show_list&"</tr>"
    show_list=show_list&"</table>"
      show_list=show_list&"</TD>"
  show_list=show_list&"</TR>"
  
show_list=show_list&"</TABLE>"


end sub
%>