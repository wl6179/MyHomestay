<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
If Request("action")="protocol" Then
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>��ע������<hr />"&oblog.setup(47,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������	
End If
dim action,show_reg,chkregtime
chkregtime=300 '�ظ�ע���ʱ����,��λ����
if is_ot_user then
	if not IsObject(conn) then link_database
	response.Redirect(ot_regurl)
	response.End()
end if
action=trim(request.Form("action"))
call sysshow()  '���Ȼ��ģ��~...�����£������ж��Ǵ���ע�� ���� ��ʾ��ע����ˣ�
select case action
	case "chkreg"
	call sub_chkreg()'����ע��;
	case else 
	call sub_showreg()'��ʾ��ע�����;
end select
show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'����ǩ��$show_list$��Ϊ ���е����� show_reg �� + �ײ���Ȩ��Ϣ���룻
response.Write show'ȫ�������
sub sub_showreg()
	dim rs
	dim str_usertype
	
	'show_reg=show_reg&"<div  style='font-size:12px;' CLASS='Register'>"

         

	show_reg=show_reg&"��ǰλ�ã�<a href='index.asp'>��ҳ</a>��ע�����û�<hr />"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"��ǰϵͳ�ѹر�ע�ᡣ"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",0,0)
    str_usertype=str_usertype&"</select><font color=#ff0000> *</font>"
	show_reg=show_reg&"<form name=oblogform method=post action=reg.asp onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>��¼�û�����</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=53% colspan=2> <input name=username type=text id=username size=15 maxlength=30><font color=#ff0000 > *</font>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=30% height=25><div align=right >������</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td width=53% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>�鿴�û����������Ƿ�ռ��</a>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr >"& vbcrlf 
    show_reg=show_reg&"<td width=30% height=25><div align=right>�����¼���룺</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=password type=password id=password size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>��¼����ȷ�ϣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=repassword type=password id=repassword size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>������ʾ���⣺</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>������ʾ�𰸣�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>����(ʡ/��)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&oblog.type_city("","")&"<font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>Email��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>��ʵ������</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30><font color=#ff0000> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>Blog���</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=30% height=25><div align=right>ע�����</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked>ͬ�⡡<input class='input_radio' type=radio name=passregtext id=passregtext value='0'>��ͬ�⡡<a href='reg.asp?action=protocol' target='_blank'>���鿴ע������</a></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf	
	
	if oblog.setup(57,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
    	show_reg=show_reg&"<td width=30% height=25><div align=right>��֤�룺</div></td>"& vbcrlf
   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
   		show_reg=show_reg&"</tr>"& vbcrlf
	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 ></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' �ύ '>��������"& vbcrlf
	show_reg=show_reg&"<input type=reset name=Submit2 value=' ���� '><input type=hidden name=action value='chkreg'>"& vbcrlf
    show_reg=show_reg&"</div></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"</table>"& vbcrlf
	show_reg=show_reg&"</form>"& vbcrlf
	

	'show_reg=show_reg&"</div>"
    
	end if
	set rs=nothing
end sub

sub sub_chkreg()
	if oblog.ChkPost()=false then
		oblog.adderrstr("ϵͳ��������ⲿ�ύ��")
		oblog.showerr
		exit sub
	end if
	chk_regtime()
	if oblog.setup(57,0)=1 then
		if not oblog.codepass then oblog.adderrstr("��֤�������ˢ�º��������룡"):oblog.showerr
	end if
	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	if oblog.chkiplock() then
		oblog.adderrstr("�Բ������IP�ѱ�����,������ע�ᣡ")
		oblog.showerr
		exit sub
	end if
	regusername=oblog.filt_badstr(trim(request("username")))
	if regusername="" then
		response.Redirect "index.asp"
	end if
	regpassword=request("password")
	re_regpassword=request("repassword")
	sex=request("sex")
	email=trim(request("email"))
	question=trim(request("question"))
	answer=trim(request("an"))
	blogname=trim(request("blogname"))
	usertype=clng(request("usertype"))
	'nickname=trim(request("nickname"))
	user_domain=Lcase(trim(request("domain")))
	user_domainroot=trim(request("user_domainroot"))
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("�û�������Ϊ��(���ܴ���14С��4)��")
	if oblog.chk_regname(regusername) then oblog.adderrstr("�û���ϵͳ������ע�ᣡ")
	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("�û����к���ϵͳ��������ַ���")
	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("�û���������ȫ��Ϊ���֣�")
	if oblog.chkdomain(regusername)=false then oblog.adderrstr("�û������Ϲ淶��ֻ��ʹ��Сд��ĸ�����ּ��»��ߣ�")
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("��������Ϊ��(���ܴ���14���ַ�)��")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("��������С��4���ַ���")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("������ϵͳ������ע�ᣡ")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("�����к���ϵͳ��������ַ���")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("�������Ϲ淶��ֻ��ʹ��Сд��ĸ�����֣�")
		if user_domainroot="" then oblog.adderrstr("����������Ϊ�գ�")
	end if
	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("���벻��Ϊ��(���ܴ���14С��4)��")
	if re_regpassword="" then oblog.adderrstr("�ظ����벻��Ϊ�գ�")
	if regpassword<>re_regpassword then oblog.adderrstr("�����������벻ͬ��")
	if question="" or oblog.strLength(question)>50 then oblog.adderrstr("�һ�������ʾ���ⲻ��Ϊ��(���ܴ���50)��")
	if answer="" or oblog.strLength(answer)>50 then oblog.adderrstr("�һ���������𰸲���Ϊ��(���ܴ���50)��")
	'if oblog.chk_regname(nickname) then oblog.adderrstr("���ǳ�ϵͳ������ע�ᣡ")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("�ǳ��к���ϵͳ��������ַ���")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("�ǳƲ��ܲ��ܴ���50�ַ���")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("��ʵ��������Ϊ��(���ܴ���50�ַ�)��")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("��ʵ�����к���ϵͳ��������ַ���")	
	if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"��")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("�û����к��зǷ��ַ���")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������ǳƴ��ڣ�������ǳƣ�")		
	'end if
	if user_domain<>"" then
		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("ϵͳ���Ѿ�������������ڣ������������")
	end if
	if oblog.errstr<>"" then oblog.showerr:exit sub
	if oblog.setup(16,0)=1 then	reguserlevel=6 else reguserlevel=7
	regpassword=md5(regpassword)
	if not IsObject(conn) then link_database
	set rsreg=server.CreateObject("adodb.recordset")
	rsreg.open "select * from [oblog_user] where username='"& oblog.filt_badstr(regusername) &"'",conn,1,3
	if rsreg.eof then
		rsreg.addnew
		rsreg("username")=regusername
		rsreg("password")=regpassword
		if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
			rsreg("user_domain")=user_domain
			rsreg("user_domainroot")=user_domainroot
		end if
		rsreg("question")=question
		rsreg("answer")=md5(answer)
		rsreg("useremail")=email
		rsreg("user_level")=reguserlevel
		rsreg("user_isbest")=0
		rsreg("blogname")=blogname
		rsreg("user_classid")=usertype
		'rsreg("nickname")=nickname
		rsreg("province")=request("province")
		rsreg("city")=request("city")
		rsreg("adddate")=ServerDate(now())
		rsreg("lastloginip")=oblog.userip
		rsreg("lastlogintime")=ServerDate(now())
		rsreg("user_dir")=oblog.setup(30,0)
		rsreg("user_folder")=regusername		
		rsreg.update
		oblog.execute("update oblog_setup set user_count=user_count+1")
		if is_unamedir=0 then
			oblog.execute("update oblog_user set user_folder=userid where username='"&oblog.filt_badstr(regusername)&"'")
		end if
		show_reg="��ǰλ�ã�<a href='index.asp'>��ҳ</a>�����ע��<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>ע��ɹ�������ǰϵͳ����Ϊ��Ҫͨ����ˣ�����ʱû�й���Ȩ�ޣ�</strong></li></ul>"
			show_reg=show_reg&"5����Զ�ת��ϵͳ��ҳ��"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",5000);"
			show_reg=show_reg&"</script>"
		else
			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
			show_reg=show_reg&"<ul><li><strong>��ϲ�����Ѿ�ע��ɹ���</strong></li>"
			show_reg=show_reg&"<li><a href='index.asp'>������ҳ</a></li>"
			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>�����������̨ѡ������ϲ����ҳ����(5����Զ���������̨)</a></li></ul>"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
			show_reg=show_reg&"</script>"
		end if
	else
		oblog.adderrstr("ϵͳ���Ѿ�������û������ڣ�������û�����")
		oblog.showerr
		exit sub	
	end if
	rsreg.close
	set rsreg=nothing
	'ATFLAG
	'Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))=now()	
	Session("chk_regtime")=Now()
end sub

sub chk_regtime()
	dim lasttime
	'ATFLAG
	'lasttime = Session("chk_regtime"&replace(Request.ServerVariables("REMOTE_ADDR"),".",""))
	lasttime = Session("chk_regtime")
	If IsDate(lasttime) Then 
		If DateDiff("s",lasttime,Now()) < chkregtime then
			Response.Write("<script language=javascript>alert('"&chkregtime&"�������ظ�ע�ᡣ');window.history.back(-1);</script>")
			Response.End
		end if
	end if
end sub
%>
<SCRIPT language="javascript">
<!--
function checkssn(gotoURL) {
   var ssn=oblogform.username.value;
   var domain=oblogform.domain.value;
   var domainroot=oblogform.user_domainroot.value;
	   var open_url = gotoURL + "?username=" + ssn+"&domain="+domain+"&domainroot="+domainroot;
	   window.open(open_url,'','status=0,directories=0,resizable=0,toolbar=0,location=0,scrollbars=0,width=322,height=200');
}
function checkerr(string)
{
var i=0;
for (i=0; i<string.length; i++)
{
if((string.charAt(i) < '0' || string.charAt(i) > '9') && (string.charAt(i) < 'a' || string.charAt(i) > 'z')&& (string.charAt(i)!='-')) 
{
return 1;
}
}
return 0;//pass
}
function del_space(s)
{
	for(i=0;i<s.length;++i)
	{
	 if(s.charAt(i)!=" ")
		break;
	}

	for(j=s.length-1;j>=0;--j)
	{
	 if(s.charAt(j)!=" ")
		break;
	}

	return s.substring(i,++j);
}

function VerifySubmit()
{
	 if (document.oblogform.passregtext[1].checked )
     {
        alert("����ͬ��ע���������ע��!");
	return false;
   }	 
   
 
		 
	uid = del_space(document.oblogform.username.value);
     if (uid.length == 0)
     {
        alert("�������û���!");
	return false;
     }
	 
	pwd = del_space(document.oblogform.password.value);
     if (pwd.length == 0)
     {
        alert("����������!");
	return false;
     }
	 
	 pwd = del_space(document.oblogform.password.value);
     if (pwd.length < 6)
     {
        alert("����Ҫ����6���ַ���");
	return false;
     }

	pwd2 = del_space(document.oblogform.repassword.value);
     if (pwd2!=pwd)
     {
        alert("������������벻һ�£�");
		return false;
     }
	 
	 tishi = del_space(document.oblogform.question.value);
     if (tishi.length == 0)
     {
        alert("������������ʾ����");
        return false;
     }
	 
	 tsda = del_space(document.oblogform.an.value);
     if (tsda.length == 0)
     {
        alert("������������ʾ����𰸣�");
        return false;
     }
	 city = del_space(document.all("city").value);
     if (city.length == 0)
     {
        alert("��ѡ�����ڳ���!");
	return false;
     }		
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("������Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email��ַ��Ч!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("������������ʵ����!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("��ѡ���������!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>