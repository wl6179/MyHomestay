<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%

dim action,id,rs,n,oblog
dim refreshLimitTime,timeStamp,fv
action=request.QueryString("action")
'关闭统计功能
If EN_COUNT=0 Then
	Select Case action
		Case "site"
			Response.Write "site_count.innerHTML=""-"";" 
		Case "log"
			Response.Write "document.write('-');"
		Case "code"
			Call comment_code
		Case "logtb"
			Response.Write "ob_logreaded.innerHTML=""-"";" 
			Response.Write "ob_tbnum.innerHTML=""-"";" 
		Case "logtb31"
		Case "logs"
			'暂不处理
	End Select
	Response.End
End If
	
set oblog=new class_sys
oblog.autoupdate=false
oblog.start

Select Case action
	Case "site"
		Call site_count
	Case "log"
		Call log_count
	Case "code" '兼容3.0
		Call comment_code
	Case "code31"
		Call comment_code31
	Case "logtb"	'兼容3.0版本的统计
		Call logtb_count("3.0")
	Case "logtb31"	'3.1版本的文章统计，增加(）输出
		Call logtb_count("3.1")
	Case "logs"
		Call logs_count
end Select

sub site_count
	id=clng(request.QueryString("id"))
	refreshLimitTime  =  oblog.setup(55,0)
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if	
	if request.cookies(cookies_name)("lastvisit_fresh_site"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_site"&id)=time()
		fv=true
	end if
	timeStamp=time()
	if not IsObject(conn) then link_database
	set rs=server.createobject("adodb.recordset")
	rs.open "Select user_siterefu_num from oblog_user where userid="&id,conn,1,3
	n=rs(0)+1
	if (datediff("s",request.cookies(cookies_name)("lastvisit_fresh_site"&id),timeStamp)>refreshLimitTime) or fv=true then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_site"&id)=timeStamp
	end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	'Response.Write "document.write('"&n&"');"
	Response.Write oblog.htm2js_div(n,"site_count")
end sub

sub log_count
	id=clng(request.QueryString("id"))
	refreshLimitTime  =  oblog.setup(55,0)
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if	
	if request.cookies(cookies_name)("lastvisit_fresh_log"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_log"&id)=time()
		fv=true
	end if
	timeStamp=time()
	if not IsObject(conn) then link_database
	set rs=server.createobject("adodb.recordset")
	rs.open "Select iis from oblog_log where logid="&id,conn,1,3
	n=rs(0)+1
	if (datediff("s",request.cookies(cookies_name)("lastvisit_fresh_log"&id),timeStamp)>refreshLimitTime)  or fv=true then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_log"&id)=timeStamp
	end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	Response.Write "document.write('"&n&"');"
end sub

sub logtb_count(ver)
	id=clng(request.QueryString("id"))
	dim tbn
	refreshLimitTime  =  oblog.setup(55,0)
	if refreshLimitTime="" or isnull(refreshLimitTime) then
		refreshLimitTime=0
	end if	
	if request.cookies(cookies_name)("lastvisit_fresh_log"&id)="" then
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_log"&id)=time()
		fv=true
	end if
	timeStamp=time()
	if not IsObject(conn) then link_database
	set rs=server.createobject("adodb.recordset")
	rs.open "Select iis,trackbacknum from oblog_log where logid="&id,conn,1,3
	n=rs(0)+1
	tbn=rs(1)
	if (datediff("s",request.cookies(cookies_name)("lastvisit_fresh_log"&id),timeStamp)>refreshLimitTime)  or fv=true then
		rs(0)=n
		rs.update
		if cookies_domain<>"" then Response.Cookies(cookies_name).Domain=cookies_domain
		Response.cookies(cookies_name)("lastvisit_fresh_log"&id)=timeStamp
	end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	if ver="3.0" then
		Response.Write "document.getElementById('ob_logreaded').innerHTML="""&n&""";" 
		Response.Write "document.getElementById('ob_tbnum').innerHTML="""&tbn&""";" 
	else
		Response.Write oblog.htm2js_div("("&n&")","ob_logreaded")
		Response.Write oblog.htm2js_div("("&tbn&")","ob_tbnum")
	end if
end sub

sub logs_count
	dim i,strid
	id=oblog.filt_badstr(trim(request.QueryString("id")))
	if id="" then exit sub
	id=split(id,"$")
	for i=0 to Ubound(id)
		if id(i)<>"" then
			if strid="" then
				strid=clng(id(i))
			else
				strid=strid&","&clng(id(i))
			end if
		end if
	next
	set rs=oblog.execute("Select logid,iis,commentnum,trackbacknum from oblog_log where logid in ("&strid&")")
	while not rs.eof
		Response.Write oblog.htm2js_div("("&rs(1)&")","ob_logr"&rs(0))
		Response.Write oblog.htm2js_div("("&rs(2)&")","ob_logc"&rs(0))
		Response.Write oblog.htm2js_div("("&rs(3)&")","ob_logt"&rs(0))
		rs.movenext
	wend
	set rs=nothing
	set conn=nothing
end sub

sub comment_code
	if oblog.setup(59,0)=1 then
		Response.Write(oblog.htm2js("验证码：<input name=""CodeStr"" type=""text"" size=""6"" maxlength=""4"" />"&oblog.getcode&" "))
	end if
end sub

sub comment_code31
	if oblog.setup(59,0)=1 then
		Response.Write(oblog.htm2js_div("验证码：<input name=""CodeStr"" type=""text"" size=""6"" maxlength=""4"" />"&oblog.getcode&" ","ob_code"))
	end if
end sub
%>