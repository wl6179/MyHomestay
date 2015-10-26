<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
if oblog.chkiplock() then response.write("你的ip已经被系统锁定!"):response.End()
if oblog.ChkPost()=false then response.write("不允许从外部提交"):response.End()
oblog.chk_commenttime
dim userid,rs,username,password,blog,isguest,message,mainuserid,messagetopic,homepage
userid=clng(request("userid"))
username=trim(request.form("username"))
message=trim(request.form("edit"))
messagetopic=trim(request.Form("commenttopic"))
homepage=oblog.InterceptStr(request.Form("homepage"),250)
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("用户名不能为空且不能大于20个字符！")
if oblog.chk_badword(username)>0 then oblog.adderrstr("户名中含有系统不允许的字符！")
if message="" or oblog.strLength(message)>oblog.setup(76,0) then oblog.adderrstr("留言内容不能为空且不能大于"&oblog.setup(76,0)&"英文字符)！")
if oblog.chk_badword(message)>0 then oblog.adderrstr("留言内容中含有系统不允许的字符！")
if messagetopic="" or oblog.strLength(messagetopic)>200 then oblog.adderrstr("留言标题不能为空且不能大于200英文字符)！")
if oblog.chk_badword(messagetopic)>0 then oblog.adderrstr("留言标题中含有系统不允许的字符！")
if oblog.chk_badword(homepage)>0 then oblog.adderrstr("主页连接中含有系统不允许的字符！")
if oblog.errstr<>"" then oblog.showerr:response.End()
isguest=1
password=trim(request.form("password"))
if oblog.setup(11,0)=0 or password<>"" then 
	if oblog.checkuserlogined()=false then
		password=md5(password)
		if is_ot_user=1 then
			oblog.ot_chklogin username,password,0
		else
			oblog.ob_chklogin username,password,0
		end if
	end if
	isguest=0
end if
if oblog.setup(11,0)=0 and not oblog.checkuserlogined() then
	oblog.adderrstr("需要登陆后才能发布留言")
	oblog.showerr
else
	if oblog.setup(59,0)=1 and request("CodeStr")<>"" then
		if not oblog.codepass then oblog.adderrstr("验证码错误，请返回刷新后重新输入！"):oblog.showerr
	end if
	set blog=new class_blog
	if not IsObject(conn) then link_database
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from oblog_message",conn,2,2
	rs.addnew
	rs("userid")=userid
	rs("message_user")=EncodeJP(oblog.filt_badword(username))
	rs("messagetopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(messagetopic),250))
	if true_domain=1 then
		message=Replace(message, Chr(10), "<br />")
	end if
	rs("message")=EncodeJP(oblog.filtpath(oblog.filt_badword(message)))
	rs("homepage")=homepage
	rs("addtime")=ServerDate(now())
	rs("addip")=oblog.userip
	rs("isguest")=isguest
	rs.update
	rs.close
	set rs=nothing
	oblog.execute("update oblog_user set message_count=message_count+1 where userid="&userid)
	oblog.execute("update oblog_setup set message_count=message_count+1")
	Session("chk_commenttime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=ServerDate(now())
	blog.userid=userid
	blog.update_message 3
	blog.update_newmessage userid
	Call blog.CreateFunctionPage
	response.Redirect(blog.gourl)	
	set blog=nothing
end if
%>