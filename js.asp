<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.autoupdate=false
oblog.start
dim js_blogurl,n
js_blogurl=trim(oblog.setup(3,0))
n=cint(request("n"))
if n=0 then n=1
select case cint(request("j"))
	case 1
	call tongji()
	case 2
	call topuser()
	case 3
	call adduser()
	case 4
	call listclass()
	case 5
	call showusertype()
	case 6
	call listbestblog()
	case 7
	call showlogin()
	case 8
	call showplace()
	case 88'WL显示推荐外教的标签；
	call showplace2()'WL显示推荐外教的标签；
	case 9
	call showphoto()
	case 10
	call showblogstars()
	
	case 11
	call adduser11()
	
	case 12
	call showlog12()'显示外教到京【时间表】；
	
	case 0
	call showlog()
end select
            
sub tongji()
	dim rs,logcount,commentcount,messagecount,usercount
	dim today_log,yesterday_log
	set rs=oblog.execute("select log_count,comment_count,message_count,user_count from oblog_setup")
	logcount=rs(0)
	commentcount=rs(1)
	messagecount=rs(2)
	usercount=rs(3)
	if is_sqldata then
		set rs=oblog.execute("select count(logid) from oblog_log where datediff(d,truetime,GetDate())=0")
	else
		set rs=oblog.execute("select count(logid) from oblog_log where datediff('d',truetime,now())=0")
	end if
	today_log=rs(0)
	if is_sqldata=1 then
		set rs=oblog.execute("select count(logid) from oblog_log where datediff(d,truetime,GetDate())=1")
	else	
		set rs=oblog.execute("select count(logid) from oblog_log where datediff('d',truetime,now())=1")
	end if
	yesterday_log=rs(0)
	%> 
 document.write('◎- 用户总数 <font color=green><%=usercount%></font><br> ◎- 文章总数 <font color=green><%=logcount%></font><br> ◎- 评论总数 <font color=green><%=commentcount%></font><br> ◎- 留言总数 <font color=green><%=messagecount%></font>');
 document.write('<br> ◎- 今天文章 <font color=red><%=Today_log%></font><br> ◎- 昨天文章 <font color=green><%=yesterday_log%></font></font>')
<%
	set rs=nothing
end sub

sub topuser()
	dim i,blogname,rs,userurl,order,ordersql
	order=clng(request("order"))
	i=0
	if order=1 then
		ordersql="user_siterefu_num"
	else
		ordersql="log_count"
	end if
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user] order by "&ordersql&" desc")	
	do while Not RS.Eof and n>i
	if trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.setup(12,0)=1 then
		userurl="http://"&rs(4)&"."&trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub


sub adduser()
	dim i,blogname,rs,userurl
	i=0
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user]  order by userid desc")	
	do while Not RS.Eof and n>i
	if trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.setup(12,0)=1 then
		userurl="http://"&rs(4)&"."&trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub

sub adduser11()'最新注册的中国家庭；
	dim i,blogname,rs,userurl
	i=0
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot,sex,province,city from [oblog_user] where user_level_team=0 and user_level=6  order by userid desc")	
	do while Not RS.Eof and n>i
		if trim(rs(2))<>"" then
			blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
		else
			blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
		end if
		if oblog.setup(12,0)=1 then
			userurl="http://"&rs(4)&"."&trim(rs(5))
		else
			'userurl=js_blogurl&"go.asp?userid="&rs(3)
			userurl=js_blogurl&"HomestayFamilyMember.asp?id="& rs(3) 'WL;
		end if
		response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"& Left(rs(2),1) &"先生/女士的家庭资料>');"
		'''response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
		response.write "document.write('&nbsp;"& Left(blogname,1) &"');" & vbcrlf
		 If Cint(rs("sex"))=1 Then response.write "document.write('先生');" & vbcrlf
		 If Cint(rs("sex"))=0 Then response.write "document.write('女士');" & vbcrlf
		 If Cint(rs("sex"))=3 Then response.write "document.write('先生/女士');"  & vbcrlf
		response.write "document.write('</a>');" & vbcrlf
		response.write "document.write('&nbsp;&nbsp;&nbsp;("& rs("province") &" &nbsp;"& rs("city") &")');"& vbcrlf
		response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub


sub listbestblog()
	dim i,blogname,rs,userurl
	i=0
	set rs=oblog.execute("select top "&n&" username,log_count,blogname,userid,user_domain,user_domainroot from [oblog_user] where user_isbest=1 order by log_count desc")	
	do while Not RS.Eof and n>i
	if trim(rs(2))<>"" then
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	else
		blogname=oblog.filt_html(Replace(Replace(Replace(Replace(rs(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	end if
	if oblog.setup(12,0)=1 then
		userurl="http://"&rs(4)&"."&trim(rs(5))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(3)
	end if
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&userurl&" target=_blank title=查看"&rs(0)&"的blog页>');"
    response.write "document.write('"&blogname&"("&rs(1)&")</a>');"
	response.write "document.write('</span><br>');"
	rs.MoveNext
	i=i+1
	Loop
	set rs=nothing
end sub

sub showlogin()
	response.Write("function chkdiv(divid){var chkid=document.getElementById(divid);if(chkid != null){return true; }else {return false; }}"&VbCrLf)
	response.write "document.write('<div id=""ob_login""></div><script src="&js_blogurl&"login.asp?action=showjs&injs=1></script>');"
end sub

sub showplace()
	response.write "document.write('"&Replace(Replace(Replace(Replace(oblog.setup(18,0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"   
end sub

sub showplace2()'WL显示推荐外教的标签；
	response.write "document.write('"&Replace(Replace(Replace(Replace(oblog.setup(85,0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"   
end sub

sub showusertype()
	dim rs
	set rs=oblog.execute("select id,classname from [oblog_userclass] order by RootID,OrderID")
	do while Not RS.Eof
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&js_blogurl&"listblogger.asp?usertype="& rs(0) &" target=_blank title="&rs(1)&"的用户列表>');"
    response.write "document.write('"&rs(1)&"</a>');"
	response.write "document.write('</span><br>');"
	rs.MoveNext
	Loop
	set rs=nothing
end sub

sub listclass()
	dim rs
	set rs=oblog.execute("select id,classname from [oblog_logclass] order by RootID,OrderID")
	do while Not RS.Eof
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&js_blogurl&"HomestayNav.asp?classid="& rs(0) &" target=_blank title="&rs(1)&"的文章列表>');"
    response.write "document.write('"&rs(1)&"</a>');"
	response.write "document.write('</span><br>');"
	rs.MoveNext
	Loop
	set rs=nothing
end sub

sub showlog()
	dim rs,sql,ars,i
	dim orders,topic,isbest
	dim postname,classid,posttime,userid
	dim usersql,isbestsql,userurl,sdatesql
	if request("user")<>"" then 
   		userid=clng(request("user"))
	else
		userid=0
	end if
	if trim(request("orders"))=1 then
		orders="iis"
	elseif trim(request("orders"))=2 then
		orders="logid"
	elseif trim(request("orders"))=3 then
		orders="commentnum"
	else
		response.Write("错误的参数")
		response.End()
	end if
	if trim(request("classid"))="all" then
            classid=""
	else
		if isnumeric(request("classid")) then
			classid=" and classid="&cint(trim(request("classid")))&""
		else
			response.Write("错误的参数")
			response.End()
		end if
    end if
	if userid>0 then
		usersql=" and oblog_log.userid="&userid
	else
		usersql=""
	end if
	if not isnumeric(request("sdate")) then
		response.Write("错误的参数")
		response.End()	
	end if
	if not isnumeric(request("n")) then
		response.Write("错误的参数")
		response.End()
	elseif cint(request("n"))>100 then
		response.Write("不能调用大于100条数据")
		response.End()
	end if
	
	if cint(request("action"))=2 then
		isbestsql=" and isbest=1"
	else
		isbestsql=""
	end if
	if is_sqldata=1 then
		sdatesql=" and datediff(d,truetime,getdate())<"&cint(request("sdate"))
	else
		sdatesql=" and datediff('d',truetime,now())<"&cint(request("sdate"))
	end if
	set rs=oblog.execute("select top "&n&" author,topic,logid,classid,subjectid,truetime,iis,commentnum,logfile,oblog_log.userid,user_domain,user_domainroot from oblog_log,oblog_user where passcheck=1 and isdraft=0 "&sdatesql&isbestsql&classid&usersql&" and oblog_user.userid=oblog_log.userid ORDER BY "&orders&" desc")
	i=0
	do while Not RS.Eof and i<cint(request("n"))
    postname=trim(rs(0))
    POSTTIME=rs(5)
	topic=oblog.filt_html(Replace(Replace(Replace(Replace(rs(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	if oblog.setup(12,0)=1 then
		userurl="http://"&rs(10)&"."&trim(rs(11))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(9)
	end if
	if oblog.strLength(topic)>Cint(request("tlen")) then
        topic=oblog.InterceptStr(topic,request("tlen")+3)&"..."
    end if
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt"">');"
	if request("classname")=1 then
	set ars=oblog.execute("select classname from oblog_logclass where id="&rs(3))
		if not ars.eof then
			response.write "document.write('<a href="&js_blogurl&"HomestayNav.asp?classid="&rs(3)&" target=_blank>〖"&oblog.filt_html(ars(0))&"〗</a>');"
		end if
	end if

	if request("subjectname")=1 then
	set ars=oblog.execute("select subjectname from oblog_subject where subjectid="&rs(4))
		if not ars.eof then
			response.write "document.write('<a href="&js_blogurl&"blog.asp?name="&rs(0)&"&subjectid="&rs(4)&" target=_blank>["&oblog.filt_html(ars(0))&"]</a>');"
		end if
	end if
    response.write "document.write('<a href="&js_blogurl&"go.asp?logid="&rs(2)&" title="&topic&" target=_blank>');"
    response.write "document.write('"&topic&"');"
	response.write "document.write('</a>');"

	select case cint(request("info"))
	case 0
	case 1
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,0)&"</font>)');"
	case 2
	response.write "document.write('(<font color=green>"&POSTTIME&"</font>)');"
	case 3
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>)');"
	case 4
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&rs(6)&"</font>)');"
	case 5
	response.write "document.write('(<font color=green>"&rs(6)&"</font>)');"
	case 6
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
	case 7
	response.write "document.write('(<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
	case 8
	response.write "document.write('(<font color=green>"&rs(7)&"</font>)');"
	case else

	end select

	response.write "document.write('</span><br>');"
	RS.MoveNext
	i=i+1
	Loop
	rs.close
    set ars=nothing
	set rs=nothing
end sub


sub showlog12()'显示外教到京【时间表】WL;
	dim rs,sql,ars,i
	dim orders,topic,isbest
	dim postname,classid,posttime,userid
	dim usersql,isbestsql,userurl,sdatesql
	
	Dim isTcomehere,TcomehereDate,TName,TcomehereHowLongTime,TcomehereWhither,TNationality
	if request("user")<>"" then 
   		userid=clng(request("user"))
	else
		userid=0
	end if
	if trim(request("orders"))=1 then
		orders="iis"
	elseif trim(request("orders"))=2 then
		orders="logid"
	elseif trim(request("orders"))=3 then
		orders="commentnum"
	else
		response.Write("错误的参数")
		response.End()
	end if
	if trim(request("classid"))="all" then
            classid=""
	else
		if isnumeric(request("classid")) then
			classid=" and classid="&cint(trim(request("classid")))&""
		else
			response.Write("错误的参数")
			response.End()
		end if
    end if
	if userid>0 then
		usersql=" and oblog_log.userid="&userid
	else
		usersql=""
	end if
	if not isnumeric(request("sdate")) then
		response.Write("错误的参数")
		response.End()	
	end if
	if not isnumeric(request("n")) then
		response.Write("错误的参数")
		response.End()
	elseif cint(request("n"))>100 then
		response.Write("不能调用大于100条数据")
		response.End()
	end if
	
	if cint(request("action"))=2 then
		isbestsql=" and isbest=1"
	else
		isbestsql=""
	end if
	if is_sqldata=1 then
		sdatesql=" and datediff(d,truetime,getdate())<"&cint(request("sdate"))
	else
		sdatesql=" and datediff('d',truetime,now())<"&cint(request("sdate"))
	end if
	set rs=oblog.execute("Select Top "&n&" author,topic,logid,classid,subjectid,truetime,iis,commentnum,logfile,oblog_log.userid,user_domain,user_domainroot,isTcomehere,TcomehereDate,TName,TcomehereHowLongTime,TcomehereWhither,TNationality  From oblog_log,oblog_user Where isTcomehere=1 AND passcheck=1 and isdraft=0 "&sdatesql&isbestsql&classid&usersql&" and oblog_user.userid=oblog_log.userid ORDER BY "&orders&" desc")
	
	
		
	i=0
	do while Not RS.Eof and i<cint(request("n"))
	
	isTcomehere=rs("isTcomehere")'是否记录进 ‘外教到来时间表’？
	TcomehereDate=rs("TcomehereDate")'外教到来时间表
	TName=rs("TName")'外教的姓名
	TcomehereHowLongTime=rs("TcomehereHowLongTime")'外教到来打算住几个月
	TcomehereWhither=rs("TcomehereWhither")'外教目的地（如海淀区、sohu现代城附近）
	TNationality=rs("TNationality")'外教的国籍
	
	
    postname=trim(rs(0))
    POSTTIME=rs(5)
	topic=oblog.filt_html(Replace(Replace(Replace(Replace(rs(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),""))
	if oblog.setup(12,0)=1 then
		userurl="http://"&rs(10)&"."&trim(rs(11))
	else
		userurl=js_blogurl&"go.asp?userid="&rs(9)
	end if
	if oblog.strLength(topic)>Cint(request("tlen")) then
        topic=oblog.InterceptStr(topic,request("tlen")+3)&"..."
    end if
	response.write "document.write('<FONT color=#800000 style=font-family:webdings>4</FONT><span style=""font-size:9pt;line-height: 15pt"">');"
	if request("classname")=1 then
	set ars=oblog.execute("select classname from oblog_logclass where id="&rs(3))
		if not ars.eof then
			response.write "document.write('<a href="&js_blogurl&"HomestayNav.asp?classid="&rs(3)&" target=_blank>〖"&oblog.filt_html(ars(0))&"〗</a>');"
		end if
	end if

	if request("subjectname")=1 then
	set ars=oblog.execute("select subjectname from oblog_subject where subjectid="&rs(4))
		if not ars.eof then
			response.write "document.write('<a href="&js_blogurl&"blog.asp?name="&rs(0)&"&subjectid="&rs(4)&" target=_blank>["&oblog.filt_html(ars(0))&"]</a>');"
		end if
	end if
    response.write "document.write('<a href="&js_blogurl&"go.asp?logid="&rs(2)&" title="&topic&" target=_blank>');"
    response.write "document.write('["& TcomehereDate &"]&nbsp;<strong>"& TName &"</strong>&nbsp;"& TNationality &"&nbsp;,"& TcomehereHowLongTime &"&nbsp;,-->"& TcomehereWhither &"');"
	response.write "document.write('</a><br />');"

	select case cint(request("info"))
	case 0
	case 1
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,0)&"</font>)');"
	case 2
	response.write "document.write('(<font color=green>"&POSTTIME&"</font>)');"
	case 3
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>)');"
	case 4
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&rs(6)&"</font>)');"
	case 5
	response.write "document.write('(<font color=green>"&rs(6)&"</font>)');"
	case 6
	response.write "document.write('(<a href="&userurl&" target=_blank>"&postname&"</a>,<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
	case 7
	response.write "document.write('(<font color=green>"&formatdatetime(POSTTIME,1)&"</font>)');"
	case 8
	response.write "document.write('(<font color=green>"&rs(7)&"</font>)');"
	case else

	end select

	response.write "document.write('</span><br>');"
	RS.MoveNext
	i=i+1
	Loop
	rs.close
    set ars=nothing
	set rs=nothing
end sub


sub showphoto()
	dim rs,n,i,w,h,show_newphoto
	n=clng(request("n"))
	i=clng(request("i"))
	w=clng(request("w"))
	h=clng(request("h"))
	if i=1 then i="<br />" else i=""
	set rs=oblog.execute("select top "&clng(n)&" file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,user_folder from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1  order by fileid desc")
	while not rs.eof
		show_newphoto=show_newphoto&"<a href='"&js_blogurl&rs("user_dir")&"/"&rs("user_folder")&"/MyHomestay."&f_ext&"?uid="&rs("userid")&"&do=album' target='_blank'><img src="""&js_blogurl&rs(0)&""" width="""&w&""" height="""&h&""" hspace=""6"" border=""0"" vspace=""6"" alt='"&oblog.filt_html(rs(1))&"' /></a>"&i
		rs.movenext
	wend
	response.write "document.write('"&Replace(Replace(Replace(Replace(show_newphoto,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"
	set rs=nothing
end sub

sub showblogstars()
	dim rs,n,i,w,h,show_blogstars
	n=clng(request("n"))
	i=clng(request("i"))
	w=clng(request("w"))
	h=clng(request("h"))	
	show_blogstars=show_blogstar2(n,i,w,h)
	response.write "document.write('"&Replace(Replace(Replace(Replace(show_blogstars,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"');"
	set rs=nothing
end sub

Public Function show_blogstar2(iNumber,iPerline,iWidth,iHeight)
	Dim rs,iCount,sLine
	If Not IsNumeric(iNumber) Then 
		iNumber=1
	Else
		iNumber=CLng(iNumber)
	End If
	'iWidth=160
	'iHeight=160
	If iNumber=0 Then iNumber=1
	set rs=oblog.execute("select top " & iNumber & " * from oblog_blogstar where ispass=1 order by id desc")	
	if not rs.eof then
		show_blogstar2="<table style=""table-layout: fixed"" width=" & (iWidth +2) * iPerline &" border=0><tr>"
		If iNumber=1 Then
			sLine = "<td nowrap  valign=top style=""width:" & (iWidth+2)& "px;white-space:nowrap""><a href='"&rs("userurl")&"' target='_blank'><img src="""&rs("picurl")&"""  hspace=""3"" border=""0"" vspace=""3"" alt='"&oblog.filt_html(rs("blogname"))&"' onload=""javascript:if(this.width>" & iWidth &") this.style.width=" & iWidth &";"" /></a><BR/>"
			sLine = sLine &"用户："&"<a href='"&rs("userurl")&"' target='_blank'>"&oblog.filt_html(rs("blogname"))&"</a><BR/>"
			sLine = sLine &"简介："&oblog.filt_html(rs("info"))&"</td>" 
			show_blogstar2 = show_blogstar2 & sLine & "</tr>" & VBCRLF
		'多图片时强制大小统一
		Else		
			iCount=1			
			Do While Not rs.Eof				
				sLine = "<td nowrap  valign=top style=""width:" & (iWidth+2)& "px;white-space:nowrap""><a href='"&rs("userurl")&"' target='_blank'><img src="""&rs("picurl")&"""  hspace=""3"" border=""0"" vspace=""3"" alt='"&oblog.filt_html(rs("blogname"))&"' width=" & iWidth &" height=" & iHeight &" /></a><BR/>"
				sLine = sLine &"用户："&"<a href='"&rs("userurl")&"' target='_blank'>"&oblog.filt_html(rs("blogname"))&"</a><BR/>"
				sLine = sLine &"简介："&oblog.filt_html(rs("info"))&"</td>" & VBCRLF
				show_blogstar2 = show_blogstar2 & sLine
				If iCount Mod iPerline=0 Then show_blogstar2 = show_blogstar2 & "</tr>"
				iCount = iCount+1
				rs.MoveNext
			Loop
			If Right(show_blogstar2,5)<>"</tr>" Then show_blogstar2 = show_blogstar2 & "</tr>"
		End If
		show_blogstar2 = show_blogstar2 & "</table>"
	Else
		show_blogstar2=" "
	End If
	rs.Close
	set rs=nothing
End Function
%>