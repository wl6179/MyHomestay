<!--#include file="conn.asp"-->
<%
dim rs,skinid,userid
dim show,skinshowlog,i
skinid=clng(request("id"))
userid=clng(request("userid"))
if not IsObject(conn) then link_database
if skinid>0 then
	set rs=conn.execute("select skinmain,skinshowlog from oblog_userskin where id="&skinid)
	if rs.eof then
		response.write "ģ�治���ڣ�"
	else
		skinshowlog=rs(1)
		for i=1 to 6
			skinshowlog=skinshowlog+rs(1)
		next
		show=replace(rs(0),"$show_log$",skinshowlog)
		response.write show
	end if
elseif userid>0 then
	set rs=conn.execute("select bak_skin1,bak_skin2 from oblog_user where userid="&userid)
	if rs.eof then
		response.write "�û������ڣ�"
	else
		if rs(0)="" or rs(1)="" or isnull(rs(0)) or isnull(rs(1)) then
			response.Write("��ǰû�б���ģ��")
		else 
			skinshowlog=rs(1)
			for i=1 to 6
				skinshowlog=skinshowlog+rs(1)
			next
			show=replace(rs(0),"$show_log$",skinshowlog)
			response.write show
		end if
	end if
end if
set rs=nothing
if isobject(conn) then conn.close:set conn=nothing
%>