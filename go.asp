<!--#include file="conn.asp"-->
<%
dim logid,commentid,userid,messageid,action,user,sThisMonth,sMonth
dim rs,uf,sql
logid=Trim(Request("logid"))
commentid=Trim(Request("commentid"))
messageid=Trim(Request("messageid"))
userid=Trim(Request("userid"))
user=Trim(Request("user"))
action=Trim(Request("action"))
	
If logid<>"" And IsNumeric(logid) Then 
	logid=CLNG(logid)
Else
	logid=""
End If
If commentid<>"" And IsNumeric(commentid) Then 
	commentid=CLNG(commentid)
Else
	commentid=""
End If
If messageid<>"" And IsNumeric(messageid) Then 
	messageid=CLNG(messageid)
Else
	messageid=""
End If
If userid<>"" And IsNumeric(userid) Then 
	userid=CLNG(userid)
Else
	userid=""
End If
if not IsObject(conn) then link_database
if logid<>"" then
	sql="select logfile from oblog_log where logid="&logid
	set rs=conn.execute(sql)
	if not rs.eof then
		dim logf,c
		if c<>"" then c="#"&c
		logf=rs(0)&c
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		response.Redirect(logf)
	else
		set rs=nothing
		response.Write("无此文章")
	end if
elseif userid<>"" then
	set rs=conn.execute("select user_dir,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where userid="&clng(userid))
	if not rs.eof then
		if true_domain=1 then
			if rs("custom_domain")="" or isnull(rs("custom_domain")) then
				uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/index."&f_ext
			else
				uf="http://"&rs("custom_domain")&"/index."&f_ext
			end if
		else
			uf=rs("user_dir")&"/"&rs("user_folder")&"/index."&f_ext
		end if
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		response.Redirect(uf)
	else
		set rs=nothing
		response.Write("无此用户")
	end if
elseif messageid<>"" then
	set rs=conn.execute("select user_dir,oblog_user.userid,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_message,oblog_user where oblog_message.userid=oblog_user.userid and messageid="&clng(messageid))
	if not rs.eof then
		if true_domain=1 then
			if rs("custom_domain")="" or isnull(rs("custom_domain")) then
				uf="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/message."&f_ext
			else
				uf="http://"&rs("custom_domain")&"/message."&f_ext
			end if
		else
			uf=rs("user_dir")&"/"&rs("user_folder")&"/message."&f_ext
		end if
		set rs=nothing
		if isobject(conn) then conn.close:set conn=nothing
		response.Redirect(uf)
	else
		set rs=nothing
		response.Write("无此留言")
	end if
'按用户名访问	
elseif user<>"" Then
	user=Replace(user,"'","")
	user=Replace(user,"%","")
	user=Replace(user," ","")
	user=Replace(user,"--","")
	If user<>"" Then 
		set rs=conn.execute("Select user_dir,user_folder From oblog_user where username='" & user & "'")
		if not rs.eof then
			uf=rs(0)&"/"&rs(1)&"/index."&f_ext
			set rs=nothing
			if isobject(conn) then conn.close:set conn=nothing
			response.Redirect(uf)
		else
			set rs=nothing
			response.Write("无此用户")
		end if
	Else
		response.Write("无此用户")
	End If
end if
if isobject(conn) then conn.close:set conn=nothing
%>