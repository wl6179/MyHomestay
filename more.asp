<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim rs,logid,rstmp,password,logfile,uid,blogpw,action,encommment,commenttopic,log_month,logtopic,user_folder,user_dir
dim MaxPerPage,show,blog
logid=clng(request("id"))
password=trim(request("password"))
action=trim(request("action"))
currentpage=clng(request("page"))
MaxPerPage=10
if logid=0 then
	oblog.adderrstr("文章参数错误")
	oblog.showerr
end if
set rs=oblog.execute("select * from oblog_log where logid="&logid)
if rs.eof then
	set rs=nothing
	oblog.adderrstr("无此文章")
	oblog.showerr
end if
logfile=rs("logfile")
uid=rs("userid")
logtopic=rs("topic")
if rs("ishide")=0 and (rs("ispassword")="" or isnull(rs("ispassword"))) and action="" then
	set rs=nothing
	response.Redirect(logfile)
end if
set rstmp=oblog.execute("select * from oblog_user where userid="&uid)
user_dir=rstmp("user_dir")
user_folder=rstmp("user_folder")
if rs("blog_password")=1 then
	blogpw=Request.Cookies(cookies_name)("blogpw")
	set rstmp=oblog.execute("select * from oblog_user where userid="&uid)
	'user_dir=rsttmp("user_dir")
	'user_folder=rsttmp("user_folder")
	if (rstmp("blog_password")<>"" or isnull(rstmp("blog_password"))=false) and blogpw<>rstmp("blog_password") then
		set rs=nothing
		set rstmp=nothing
		response.Redirect("chkblogpassword.asp?userid="&uid&"&fromurl="&replace(oblog.GetUrl,"&","$"))
	end if
end if
if rs("ishide")=1 then
	if not oblog.checkuserlogined() then
		oblog.adderrstr("需要登陆才可以查看隐藏文章!")
		oblog.showerr
	else
		set rstmp=oblog.execute("select id from oblog_friend where userid="&rs("userid")&" and friendid="&oblog.logined_uid)
		if rstmp.eof and oblog.logined_uid<>rs("userid") then
			set rs=nothing
			set rstmp=nothing
			oblog.adderrstr("您无权限查看此文章，请联系blog主人!")
			oblog.showerr
		else
			set rstmp=nothing
		end if
	end if
end if

if rs("ispassword")<>"" then
	if password="" and request.Cookies(cookies_name)("logpw_"&logid)="" and action="" then
		set rs=nothing
		response.Redirect(logfile)
	end if
	if password<>"" then password=md5(password)
	if password<>rs("ispassword") and  request.Cookies(cookies_name)("logpw_"&logid)<>rs("ispassword") then
		set rs=nothing
		oblog.adderrstr("文章访问密码错误,请重新输入!")
		oblog.showerr
	else
		if password<>"" then Response.Cookies(cookies_name)("logpw_"&logid)=password
		set rs=nothing
	end if
end if 

call main()

sub main()
	select case action
		case "comment"
		call showcomment
		case else
			'Response.Redirect blogdir & user_dir & "/" & user_folder & "/cmd.asp?uid=" & uid & "&do=show&id=" &logid
		call showlog()
	end select 
end sub

sub showlog()
	set blog=new class_blog
	blog.userid=uid
	blog.showpwlog=true
	blog.update_log logid,0 
	show=blog.filt_pwblog(blog.m_log,logtopic)
	show=replace(show,"savecomment.asp?logid=","savecomment.asp?t=1&logid=")
	response.Write(show)
	set blog=nothing
end sub

sub showcomment()
	set blog=new class_blog
	encommment=rs("isencomment")
	commenttopic="Re:"&oblog.filt_html(logtopic)
	if int(month(rs("addtime")))<10 then
		log_month=year(rs("addtime"))&"0"&month(rs("addtime"))
	else
		log_month=year(rs("addtime"))&month(rs("addtime"))
	end if
	blog.userid=uid
	blog.showpwlog=true
	blog.showcmt logid
	show=blog.filt_pwblog(blog.m_commentsmore,rs("topic")&"--所有评论")
	response.Write(show)
	set blog=nothing	
end sub
%>