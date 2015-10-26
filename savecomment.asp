<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
if oblog.chkiplock() then response.write("你的ip已经被系统锁定!"):response.End()
if oblog.ChkPost()=false then response.write("不允许从外部提交!"):response.End()
oblog.chk_commenttime
dim logid,rs,username,password,blog,isguest,comment,mainuserid,commenttopic
logid=clng(request("logid"))
username=oblog.filt_badstr(trim(request.form("username")))
comment=trim(request.form("edit"))
commenttopic=trim(request.Form("commenttopic"))
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("昵称不能为空且不能大于20个字符！")
if oblog.chk_badword(username)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
if comment="" or oblog.strLength(comment)>oblog.setup(76,0) then oblog.adderrstr("回复内容不能为空且不能大于"&oblog.setup(76,0)&"英文字符！")
if oblog.chk_badword(comment)>0 then oblog.adderrstr("回复内容中含有系统不允许的字符！")
if commenttopic="" or oblog.strLength(commenttopic)>200 then oblog.adderrstr("回复标题不能为空且不能大于200英文字符！")
if oblog.chk_badword(commenttopic)>0 then oblog.adderrstr("回复标题中含有系统不允许的字符！")
if oblog.chk_badword(request.Form("homepage"))>0 then oblog.adderrstr("主页地址中含有系统不允许的字符！")
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
if oblog.setup(11,0)=0 and oblog.checkuserlogined()=false then
	oblog.adderrstr("需要登陆后才能发布评论")
	oblog.showerr
else
	if oblog.setup(59,0)=1 and request("CodeStr")<>"" then
		if not oblog.codepass then oblog.adderrstr("验证码错误，请返回刷新后重新输入！"):oblog.showerr
	end if
	set blog=new class_blog
	'增加对加密文章的处理，防止通过URL连接或软件方式对文章进行回复
	set rs=oblog.execute("select userid,ispassword,ishide from oblog_log where logid="&logid)	
	if rs.eof then response.Write("参数错误"):set rs=nothing:response.End()
	If request.Cookies(cookies_name)("logpw_"&logid)<>rs("ispassword")  Then
		response.Write("错误的操作!")
		Set rs=nothing
		Response.End()
	End If
	mainuserid=rs(0)
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select top 1 * from oblog_comment",conn,2,2
	rs.addnew
	rs("mainid")=logid
	rs("userid")=mainuserid
	rs("comment_user")=EncodeJP(username)
	rs("commenttopic")=EncodeJP(oblog.InterceptStr(oblog.filt_badword(commenttopic),250))
	if true_domain=1 then
		comment=Replace(comment, Chr(10), "<br />")'wl可能出问题的地方；
	end if
	rs("comment")=EncodeJP(oblog.filtpath(oblog.filt_badword(comment)))
	rs("homepage")=oblog.InterceptStr(request.Form("homepage"),250)
	rs("addtime")=ServerDate(now())
	rs("addip")=oblog.userip
	rs("isguest")=isguest
	rs.update
	rs.close
	set rs=nothing
	oblog.execute("update oblog_log set commentnum=commentnum+1 where logid="&logid)
	oblog.execute("update oblog_user set comment_count=comment_count+1 where userid="&mainuserid)
	oblog.execute("update oblog_setup set comment_count=comment_count+1")
	Session("chk_commenttime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=ServerDate(now())
	blog.userid=mainuserid
	blog.update_comment(mainuserid)
	Server.ScriptTimeOut=99999	
	blog.update_log logid,3
	blog.update_comment mainuserid
	Call blog.CreateFunctionPage
	if request("t")="1" then
		response.Redirect("more.asp?id="&logid)
	else
		response.Redirect(blog.gourl)
response.Write "<br>"&comment&"111  "
	end if
	set blog=nothing
end if
%>