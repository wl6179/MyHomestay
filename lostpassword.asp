<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
if is_ot_user then
	if not IsObject(conn) then link_database
	response.Redirect(ot_lostpasswordurl)
	response.End()
end if
dim action,show_getpassword
action=cint(request("action"))
call sysshow()
show_getpassword="��ǰλ�ã�<a href='index.asp'>��ҳ</a>���һ�����<hr noshade>"
select case action
	case 1
	call sub_getpassword_1()
	case 2
	call sub_getpassword_2()
	case 3
	call sub_getpassword_3()
	case else
	call sub_getpassword_0()
end select


show=replace(show,"$show_list$",show_getpassword)
response.Write show&oblog.site_bottom

dim pass_username,daan

sub sub_getpassword_0
	show_getpassword=show_getpassword&"<form name='form1' method='post' action=''>"
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0 borderColor=#111111 style='BORDER-COLLAPSE: collapse'>"
	show_getpassword=show_getpassword&"<tr><td height='18' class='bian'><strong>�һ������һ��:</strong></td> </tr><tr>"
	show_getpassword=show_getpassword&"<td height='200'><div align='center'>�������û���:"
	show_getpassword=show_getpassword&"<input name='uid' type='text' id='uid' size='23' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br><input name='Submit' type='submit' id='Submit' value='��һ��'>"
	show_getpassword=show_getpassword&"<input name='action'  id='action' type='hidden' value='1'>"
	show_getpassword=show_getpassword&"</div></td></tr></table></form>"
end sub

sub sub_getpassword_1
	dim rs
	pass_username=oblog.filt_badstr(trim(request("uid")))
	if pass_username="" then oblog.adderrstr("�û�������Ϊ��(���ܴ���14С��4)��")
	set rs=oblog.execute ("select * from [oblog_user] where username='"&pass_username&"'")
	if rs.eof or err then oblog.adderrstr("���û��������ڣ�")
	if oblog.errstr<>"" then oblog.showerr:exit sub
	show_getpassword=show_getpassword&"<form name='form1' method='post' action=''>"
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0 borderColor=#111111 style='BORDER-COLLAPSE: collapse'>"
	show_getpassword=show_getpassword&"<tr><td height='18' ><strong>�һ�����ڶ���:</strong></td> "
	show_getpassword=show_getpassword&"</tr><tr> <td height='200'><div align='center'>�������û���:"
	show_getpassword=show_getpassword&"<input name='uid' type='text' id='uid' value='"&rs("username")&"' size='30' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br>������ʾ����:"
	show_getpassword=show_getpassword&"<input name='tishi' type='text' id='tishi' value='"&oblog.filt_html(rs("Question"))&"' size='30' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br>���������:"
	show_getpassword=show_getpassword&"<input name='daan' type='text' id='daan' size='30' maxlength='26'>"
	show_getpassword=show_getpassword&"<br><br><input name='Submit' type='submit' id='Submit' value='��һ��'>"
	show_getpassword=show_getpassword&"<input name='action'  id='action' type='hidden' value='��'>"
	show_getpassword=show_getpassword&"</div></td></tr></table></form>"
	rs.close
	set rs=nothing
end sub

sub sub_getpassword_2
	dim tishi,rs
	pass_username=oblog.filt_badstr(trim(request("uid")))
	daan=md5(trim(request("daan")))
 	if daan="" then oblog.adderrstr("�Բ���������ʾ����𰸲���Ϊ�գ�")
	set rs=oblog.execute("select * from [oblog_user] where username='"&pass_username&"' and answer='"&daan&"'")
	if rs.eof or err then oblog.adderrstr("������ʾ����ش���󣡣�")
	if oblog.errstr<>"" then oblog.showerr:exit sub
	show_getpassword=show_getpassword&"<TABLE width='400' border=0 cellPadding=0 cellSpacing=0  align='center' style='BORDER-COLLAPSE: collapse'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td height='18' class='bian'><strong>���,�������趨����:</strong></td></tr><tr>"& vbcrlf
	show_getpassword=show_getpassword&"<td><table width='100%' border='0' cellpadding='0' cellspacing='0'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td><FORM action='lostpassword.asp' method='post' name='regform' >"& vbcrlf
	show_getpassword=show_getpassword&"<br><br><table width='60%' border='0' align='center' cellpadding='0' cellspacing='0'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr><td><table  border='0' cellspacing='0' cellpadding='5'>"& vbcrlf
	show_getpassword=show_getpassword&"<tr> <td width='37%'><div align='right'>"& vbcrlf
	show_getpassword=show_getpassword&"<p>�����룺</p></div></td><td colspan='2'><input name='new_pass' type='password' id='new_pass'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr><tr><td><div align='right'>��֤���룺</div></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td colspan='2'><input name='new_pass2' type='password' id='new_pass2'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr><tr><td><div align='right'> </div></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td width='17%'><input type='submit' name='Submit' value='ȷ��'></td>"& vbcrlf
	show_getpassword=show_getpassword&"<td width='46%'><input type='reset' name='Submit2' value='ȡ��'></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr></table><input name='action'  id='action' type='hidden' value='3'><input name='uid'  id='uid' type='hidden' value='"&pass_username&"'><input name='daan'  id='daan' type='hidden' value='"&daan&"'></td></tr></table></form><br><div align='center'> </div></td>"& vbcrlf
	show_getpassword=show_getpassword&"</tr></table></td></tr></table>"& vbcrlf
	rs.close
	set rs=nothing
end sub

sub sub_getpassword_3()
	dim password,repassword
	pass_username=oblog.filt_badstr(trim(request("uid")))
	daan=oblog.filt_badstr(trim(request("daan")))
	password=trim(request("new_pass"))
	repassword=trim(request("new_pass2"))
	if password="" or oblog.strLength(password)>14 or oblog.strLength(password)<4 then oblog.adderrstr("���벻��Ϊ��(���ܴ���14С��4)��")
	if repassword="" then oblog.adderrstr("�ظ����벻��Ϊ�գ�")
	if password<>repassword then oblog.adderrstr("�����������벻ͬ��")
	if oblog.errstr<>"" then oblog.showerr:exit sub
	password=md5(password)
	oblog.execute("update oblog_user set [password]='"&password&"' where username='"&pass_username&"' and answer='"&daan&"'" )
	oblog.savecookie pass_username,password,0,""
	show_getpassword="��ǰλ�ã�<a href='index.asp'>��ҳ</a>���޸�����ɹ�<hr noshade>"
	show_getpassword=show_getpassword&"<strong>���������Ѿ��޸ĳɹ���</strong><br>"
	show_getpassword=show_getpassword&"<a href='index.asp'>���������ҳ��</a><br>"
	show_getpassword=show_getpassword&"5����Զ�������ҳ��"
	show_getpassword=show_getpassword&"<script language=JavaScript>"
	show_getpassword=show_getpassword&"setTimeout(""window.location='index.asp'"",5000);"
	show_getpassword=show_getpassword&"</script>"
end sub

%>