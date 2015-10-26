<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim rs,password,uid,action,show,blog,logid,subjectid,userid,user_path
uid=clng(request("userid"))
logid=clng(request("logid"))
subjectid=clng(request("subjectid"))
password=Request.Cookies(cookies_name)("blogpw")
action=request("action")
if uid=0 then
	oblog.adderrstr("用户参数错误")
	oblog.showerr
end if
set rs=oblog.execute("select * from oblog_user where userid="&uid)
if rs.eof then
	set rs=nothing
	oblog.adderrstr("无此用户")
	oblog.showerr
end if
if (rs("blog_password")<>"" or isnull(rs("blog_password"))=false) and password<>rs("blog_password") then
	uid=rs("userid")
	set rs=nothing
	response.Redirect("chkblogpassword.asp?userid="&uid&"&fromurl="&replace(oblog.GetUrl,"&","$"))
end if
set blog=new class_blog
blog.showpwblog=true
blog.userid=rs("userid")
userid=rs("userid")
user_path=blogdir & rs("user_dir") & "/" & rs("user_folder") & "/"
select case action
	case "blog"
	call showblog
	case "log"
	call showlog
	case "message"
	call showmessage
	case "album"
	call showalbum
end select
response.Write(show)
set rs=nothing
set blog=nothing

sub showblog()
	blog.update_index(0) 
	show=blog.filt_pwblog(blog.m_index,rs("blogname"))
	show=repl_c(show)
end sub

sub showlog()
	blog.update_log logid,0 
	show=blog.filt_pwblog(blog.m_log,rs("blogname"))
	show=repl_c(show)
end sub

sub showmessage()
	dim sdate,edate
	blog.update_message 0
	show=blog.filt_pwblog(blog.m_message,rs("blogname"))
	show = show & "<script src=""" & blogdir & "commentedit.asp""></script>" & vbCrlf
	show=repl_c(show)
end sub

sub showalbum
	dim sdate,edate
	sdate=trim(request("sdate"))
	edate=trim(request("edate"))
	if isdate(sdate) and isdate(edate) then
		blog.update_album 0,sdate,edate,""
	else
		blog.update_album 0,0,0,""
	end if
	show=blog.filt_pwblog(blog.m_album,rs("blogname"))
	show=repl_c(show)
end sub

function repl_c(show)
	if f_ext="htm" or f_ext="html" then
		show=replace(show,"$show_calendar$","<div id=""calendar""></div><script src='"&blogdir&blog.user_path&"/calendar/"&blog.newcalendar(blog.user_path&"/calendar")&".htm'></script>")
	else
		show=replace(show,"$show_calendar$","<div id=""calendar"">"&oblog.readfile(blog.user_path&"\calendar",blog.newcalendar(blog.user_path&"/calendar")&".htm")&"</div>")
	end if
	repl_c=show
end function
%>