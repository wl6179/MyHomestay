<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
if oblog.chkiplock() then response.write("���ip�Ѿ���ϵͳ����!"):response.End()
if oblog.ChkPost()=false then response.write("��������ⲿ�ύ!"):response.End()
oblog.chk_commenttime
dim logid,rs,username,password,blog,isguest,comment,mainuserid,commenttopic
logid=clng(request("logid"))
username=oblog.filt_badstr(trim(request.form("username")))
comment=trim(request.form("edit"))
commenttopic=trim(request.Form("commenttopic"))
if username="" or oblog.strLength(username)>20 then oblog.adderrstr("�ǳƲ���Ϊ���Ҳ��ܴ���20���ַ���")
if oblog.chk_badword(username)>0 then oblog.adderrstr("�ǳ��к���ϵͳ��������ַ���")
if comment="" or oblog.strLength(comment)>oblog.setup(76,0) then oblog.adderrstr("�ظ����ݲ���Ϊ���Ҳ��ܴ���"&oblog.setup(76,0)&"Ӣ���ַ���")
if oblog.chk_badword(comment)>0 then oblog.adderrstr("�ظ������к���ϵͳ��������ַ���")
if commenttopic="" or oblog.strLength(commenttopic)>200 then oblog.adderrstr("�ظ����ⲻ��Ϊ���Ҳ��ܴ���200Ӣ���ַ���")
if oblog.chk_badword(commenttopic)>0 then oblog.adderrstr("�ظ������к���ϵͳ��������ַ���")
if oblog.chk_badword(request.Form("homepage"))>0 then oblog.adderrstr("��ҳ��ַ�к���ϵͳ��������ַ���")
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
	oblog.adderrstr("��Ҫ��½����ܷ�������")
	oblog.showerr
else
	if oblog.setup(59,0)=1 and request("CodeStr")<>"" then
		if not oblog.codepass then oblog.adderrstr("��֤������뷵��ˢ�º��������룡"):oblog.showerr
	end if
	set blog=new class_blog
	'���ӶԼ������µĴ�����ֹͨ��URL���ӻ������ʽ�����½��лظ�
	set rs=oblog.execute("select userid,ispassword,ishide from oblog_log where logid="&logid)	
	if rs.eof then response.Write("��������"):set rs=nothing:response.End()
	If request.Cookies(cookies_name)("logpw_"&logid)<>rs("ispassword")  Then
		response.Write("����Ĳ���!")
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
		comment=Replace(comment, Chr(10), "<br />")'wl���ܳ�����ĵط���
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