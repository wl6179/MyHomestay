<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/class_blog.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim rs,uname,domain,sql,hideurl,subjectid,reurl,uid
uname=oblog.filt_badstr(request("name"))
subjectid=clng(request("subjectid"))
domain=trim(request("domain"))
if uname="" and domain="" then
	oblog.adderrstr("参数错误")
	oblog.showerr
end if
if domain<>"" then
	dim domain1,domain2
	domain=lcase(domain)
	domain=replace(domain,"http://","")
	domain=replace(domain,"/","")
	domain1=oblog.filt_badstr(replace(left(domain,instr(domain,".")),".",""))
	if trim(domain1)="" then response.Redirect("index.asp")
	domain2=oblog.filt_badstr(right(domain,len(domain)-instr(domain,".")))
	sql="select user_dir,userid,hideurl,user_folder from oblog_user where user_domain='"&domain1&"' and user_domainroot='"&domain2&"'"
end if
if uname<>"" then
	if subjectid>0 then
		set rs=oblog.execute("select userid from oblog_subject where subjectid="&subjectid)
		if not rs.eof then
			sql="select user_dir,userid,hideurl,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where userid="&rs(0)
		end if
	else
		sql="select user_dir,userid,hideurl,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where username='"&uname&"'"
	end if
end if
set rs=oblog.execute(sql)
if not rs.eof then
	hideurl=rs(2)
	if true_domain=1 then
		if rs("custom_domain")="" or isnull(rs("custom_domain")) then
			reurl="http://"&rs("user_domain")&"."&rs("user_domainroot")
		else
			reurl="http://"&rs("custom_domain")
		end if
	else
		reurl=blogdir&rs(0)&"/"&rs(3)
	end if
	if subjectid>0 then
		reurl=reurl&"/MyHomestay."&f_ext&"?uid="&rs(1)&"&do=blogs&id="&subjectid
	else
		if oblog.setup(81,0)=1 then
			reurl=reurl&"/index."&f_ext
		else
			reurl=reurl
		end if
	end if
	set rs=nothing
	if domain<>"" and hideurl=1 then
		response.Redirect(reurl)
	else
		response.Write("<script language=JavaScript>top.location='"&reurl&"';</script>")
		'response.Redirect(reurl)
	end if
else
	set rs=nothing
	oblog.adderrstr("错误：无此blog用户!")
	oblog.showerr
end if
%>