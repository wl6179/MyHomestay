<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
If Request("action")="protocol" Then
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>���Ӵ���ͥע������<hr />"&oblog.setup(47,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������	
End If
dim action,show_reg,chkregtime
chkregtime=30 '�ظ�ע���ʱ����,��λ����
if is_ot_user then
	if not IsObject(conn) then link_database
	response.Redirect(ot_regurl)
	response.End()
end if
action=trim(request.Form("action"))
call sysshow_3()  '�����ݿ����Ȼ��03ģ��~...�����£������ж��Ǵ���ע�� ���� ��ʾ��ע����ˣ�
select case action
	case "chkreg"
	call sub_chkreg()'����ע��;
	case "sub_showreg"
	call sub_showreg()';
	case else 
	call sub_showregSelect()'��ʾ��ע��ѡ�����ͱ���;
end select
show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'����ǩ��$show_list$��Ϊ ���е����� show_reg �� + �ײ���Ȩ��Ϣ���룻
response.Write show'ȫ�������

sub sub_showregSelect()

	show_reg=show_reg&"��ǰλ�ã�<a href='index.asp'>��ҳ</a> �� ����ѡ��Ӵ���ͥ����<br /><br /><font color=#038ad7>����ѡ��Ӵ���ͥ����</font><br /><br />"
	show_reg=show_reg&"<form name='FamilyForm' method='post' action='RegisterHome.asp' onSubmit='return Check();'>"

	show_reg=show_reg&"<script language='JavaScript'>"
		show_reg=show_reg&"function buttontrue(){"
		show_reg=show_reg&"document.FamilyForm.SubmitInfo.disabled=false;"
		show_reg=show_reg&"}	"	
	show_reg=show_reg&"</script>"
	
	show_reg=show_reg&"<script language='JavaScript'>"
		show_reg=show_reg&"function Check(){"
		show_reg=show_reg&"if(document.FamilyForm.FamilyType.value==''){"
		show_reg=show_reg&"alert('��ѡ��������ĽӴ���ͥ����');"
		show_reg=show_reg&"FamilyForm.FamilyType.focus();"
		show_reg=show_reg&"return false;"
		show_reg=show_reg&"}"
		show_reg=show_reg&"}"
	show_reg=show_reg&"</script>"

	
	  show_reg=show_reg&"<center><input name='FamilyType' type='checkbox' value='1' onClick='buttontrue();'>"
	  show_reg=show_reg&"������ѽӴ���ͥ "
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2' onClick='buttontrue();'>"
	  show_reg=show_reg&"�����շѽӴ���ͥ "
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3' onClick='buttontrue();'>"
	  show_reg=show_reg&"<font color='#038ad7'>����2008���˽Ӵ���ͥ</font>"
	  
	  show_reg=show_reg&"<input name='action' type='hidden' value='sub_showreg' />"
	  show_reg=show_reg&"<br /><br /><input type='submit' disabled id='SubmitInfo' name='SubmitInfo' value=' ��һ��(Next) &gt;&gt;' style='height:28px;width:150px;' ></center>"
	  
	  
	  show_reg=show_reg&"</form><br /><br /><br />"
	  
	  
	  show_reg=show_reg& "<div align='center'>"
      show_reg=show_reg& "<iframe frameborder='1' height='530' id='Terms' marginheight='0' marginwidth='0' name='Terms' scrolling='yes' src='ad/item.htm' width='500'></iframe>"
      show_reg=show_reg& "</div>"


end sub

sub sub_showreg()
	dim rs
	dim str_usertype
	
	'show_reg=show_reg&"<div  style='font-size:12px;' CLASS='Register'>"
	Dim FamilyTypeRequest,FamilyType
	
	Select case Request.Form("FamilyType")
	Case "1"
	FamilyTypeRequest= "��ѽӴ���ͥ"
	Case "2"
	FamilyTypeRequest= "�շѽӴ���ͥ"
	Case "3"
	FamilyTypeRequest= "2008���˽Ӵ���ͥ"
	Case "1, 2"
	FamilyTypeRequest= "���/�շѽӴ���ͥ"
	Case "1, 3"
	FamilyTypeRequest= "��ѽӴ���ͥ/2008���˽Ӵ���ͥ"
	Case "1, 2, 3"
	FamilyTypeRequest= "���/�շ�/2008���˽Ӵ���ͥ"
	Case "2, 3"
	FamilyTypeRequest= "�շѽӴ���ͥ/2008���˽Ӵ���ͥ"
    End Select
	'response.Write("("&FamilyTypeRequest&")") 

	show_reg=show_reg&"��ǰλ�ã�<a href='index.asp'>��ҳ</a> �� ����(���/�շ�/2008����)�Ӵ���ͥ<br /><br /><center><h3 style='height:38px;color:#00A1E4;background:url(/images/HomestayHoltel_internationalization.gif) no-repeat 30px -1px; height:130px; padding-top:0px; line-height:88px; padding-left:200px; font-size:15px;'>&nbsp;����Ӵ���ͥ</h3><font color=#038ad7>����ǰ������� "& FamilyTypeRequest &"</font></center>"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"��ǰϵͳ�ѹر�ע�ᡣ"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",0,0)
    str_usertype=str_usertype&"</select><font color=#038ad7> *</font>"
	show_reg=show_reg&"<form name='oblogform' id='oblogform' method='post' action='RegisterHome.asp' onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>��ͥ�ʺ���Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>��ͥ��¼�ʺţ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=63% colspan=2> <input name=username type=text id=username size=15 maxlength=30><font color=#038ad7 > *</font>&nbsp;(��:homestay2008)&nbsp;<a href=""javascript:checkssn('checkssn.asp')"";>���鿴���ʺ��Ƿ�ռ��</a>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=20% height=25><div align=left >������</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
'    show_reg=show_reg&"<td width=63% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>�鿴�����ʺ��Ƿ�ռ��</a>"& vbcrlf 
'    show_reg=show_reg&"</td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>�������룺</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=password type=password id=password size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>��¼����ȷ�ϣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=repassword type=password id=repassword size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>������ʾ���⣺</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>������ʾ�𰸣�</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>���ڵ���(ʡ/��)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&oblog.type_city("","")&"<font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>Email��ַ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=35 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>��ʵ������</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30><font color=#038ad7> *</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ĳƺ���<select name=sex><option value=3>��ѡ��</option><option value=0>Ůʿ</option><option value=1>����</option></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>��ѡ��ѧϰ��Χ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>�������ڣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><select name='y' size=1>"

	show_reg=show_reg& oblog.y()&"</select>��<select name='m' size=1>"& oblog.m()&"</select>��<select name='d' size=1>"& oblog.d()&"</select>��</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>����ְҵ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_job2("") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>��ϵ��ʽ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�绰��<input name=tel value='' size='15' maxlength='50' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ֻ���<input name=mobile  value='' size='15' maxlength='50' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ַ��<input name='address' value='' size='30' maxlength='250' />&nbsp;&nbsp;&nbsp;�ʱࣺ<input name=postnum value='' size='8' maxlength='6' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>��ͥ��Ա��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ������<input id='familynumber_0' type='radio' name='familynumber' value='1' /><label for='familynumber_0'>1</label>&nbsp;&nbsp;<input id='familynumber_1' type='radio' name='familynumber' value='2' /><label for='familynumber_1'>2</label>&nbsp;&nbsp;<input id='familynumber_2' type='radio' name='familynumber' value='3' checked='checked' /><label for='familynumber_2'>3</label>&nbsp;&nbsp;<input id='familynumber_3' type='radio' name='familynumber' value='4' /><label for='familynumber_3'>4</label>&nbsp;&nbsp;<input id='familynumber_4' type='radio' name='familynumber' value='����' /><label for='familynumber_4'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>��ͥ�ṹ��<input name='familymember' type='checkbox' id='Father' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Father'>����</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Mother' value='ĸ��' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Mother'>ĸ��</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Husband' value='�ɷ�' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Husband'>�ɷ�</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Wife' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Wife'>����</label><br>&nbsp;&nbsp;<font color=#ffffff>Homestay</font><input name='familymember' type='checkbox' id='Son' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Son'>����</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Daughter' value='Ů��' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Daughter'>Ů��</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Brother' value='�ֵ�' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Brother'>�ֵ�</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Sister' value='����' title='��ѡ�����ļ�ͥ�ṹ���ɶ�ѡ'><label for='Sister'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
                           
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>���ݽ��ܣ�</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�ṹ��<select name='houseframe' id='houseframe1' title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='1��' selected>1��</option>"&_
                            "<option value='2��'>2��</option>"&_
                            "<option value='3��'>3��</option>"&_
                            "<option value='4��'>4��</option>"&_
                            "<option value='����'>����</option>"&_
                            "<option value='����'>����</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe2'  title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='1��' selected>1��</option>"&_
                            "<option value='2��'>2��</option>"&_
                            "<option value='3��'>3��</option>"&_
                            "<option value='4��'>4��</option>"&_
                            "<option value='5��'>5��</option>"&_
                            "<option value='6��'>6��</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe3' title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='1��' selected>1��</option>"&_
                            "<option value='2��'>2��</option>"&_
                            "<option value='3��'>3��</option>"&_
                            "<option value='4��'>4��</option>"&_
                            "<option value='5��'>5��</option>"&_
                            "<option value='6��'>6��</option>"&_
                          "</select>"&_
						  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�����"&_ 
                          "<select name='housespace' id='housespace' style='width:100px' title='��ѡ�����ķ������'>"&_
                            "<option selected>��ѡ��</option>"&_
                            "<option value='50-100'>50��100ƽ����</option>"&_
                            "<option value='100-150'>100��150ƽ����</option>"&_
                            "<option value='150-200'>150��200ƽ����</option>"&_
                            "<option value='above 200'>200ƽ��������</option>"&_
                          "</select>"&_
						  "</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>�Ƿ��г��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value='��' name='pet' id='Pet1' onClick='ShowPet(1);' title='��ѡ�����ĳ�������'><label for='Pet1'>��</label>&nbsp;<INPUT name=pet type=radio value='û��' checked id='Pet2' onClick='ShowPet(0);'><label for='Pet2'>û��</label>&nbsp;&nbsp;&nbsp;<span width='21%' id='PetTypes'></span></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                      
    
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>����̵�����Ҫ��</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>����Ա�<input id='asksex_0' type='radio' name='asksex' value='��' /><label for='asksex_0'>��</label>&nbsp;&nbsp;&nbsp;<input id='asksex_1' type='radio' name='asksex' value='Ů' /><label for='asksex_1'>Ů</label>&nbsp;&nbsp;&nbsp;<input id='asksex_2' type='radio' name='asksex' value='����ν' checked='checked' /><label for='asksex_2'>����ν</label> &nbsp;&nbsp;&nbsp;</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>�Ƿ��˵���ģ�<input id='issaychinese_0' type='radio' name='issaychinese' value='��' /><label for='issaychinese_0'>��</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_1' type='radio' name='issaychinese' value='��' /><label for='issaychinese_1'>��</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_2' type='radio' name='issaychinese' value='����ν' checked='checked' /><label for='issaychinese_2'>����ν</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>��ͥ������Ϣ</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	                    
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>���֪�����ǣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input id='howknowsite_0' type='checkbox' name='howknowsite' value='��ֽ��־' /><label for='howknowsite_0'>��ֽ��־</label>&nbsp;&nbsp;<input id='howknowsite_1' type='checkbox' name='howknowsite' value='���ѽ���' /><label for='howknowsite_1'>���ѽ���</label>&nbsp;&nbsp;<input id='howknowsite_2' type='checkbox' name='howknowsite' value='��վ' /><label for='howknowsite_2'>��վ</label>&nbsp;&nbsp;<input id='howknowsite_3' type='checkbox' name='howknowsite' value='��̳' /><label for='howknowsite_3'>��̳</label>&nbsp;&nbsp;<input id='howknowsite_4' type='checkbox' name='howknowsite' value='��������' /><label for='howknowsite_4'>��������</label>&nbsp;&nbsp;<input id='howknowsite_5' type='checkbox' name='howknowsite' value='����' /><label for='howknowsite_5'>����</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>Ӣ��ˮƽ��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name='englishlevel' id='englishlevel_0' type='radio' value='����' title='��ѡ����Ӣ��ˮƽ'><label for='englishlevel_0'>����</label>&nbsp;&nbsp;<input name='englishlevel' id='englishlevel_1' type='radio' value='һ��' checked title='��ѡ����Ӣ��ˮƽ'><label for='englishlevel_1'>һ��</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_2' value='�е�' title='��ѡ����Ӣ��ˮƽ'><label for='englishlevel_2'>�е�</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_3' value='�ܺ�' title='��ѡ����Ӣ��ˮƽ'><label for='englishlevel_3'>�ܺ�</label>&nbsp;&nbsp;&nbsp;&nbsp;<strong>�Ƿ���˽�ҳ�</strong>��<INPUT name=car id=car_0 type=radio value=�� title='��ѡ���Ƿ���˽�ҳ�'><label for='car_0'>��</label>&nbsp;<INPUT name=car id=car_1 type=radio value=û�� checked title='��ѡ���Ƿ���˽�ҳ�'><label for='car_1'>û��</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>����������ͣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT name=toilet id=toilet_0 type=radio value=���� checked title='��ѡ���Ƿ��ж�����������'><label for='toilet_0'>����</label><INPUT type=radio value=������� name=toilet id=toilet_1 title='��ѡ���Ƿ��ж�����������'><label for='toilet_1'>�������</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>�з���ԣ�</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value=�� name=computer id=computer_0 title='��ѡ���Ƿ��е���'><label for='computer_0'>��</label>&nbsp;&nbsp;<INPUT name=computer id=computer_1 type=radio value=û�� checked title='��ѡ���Ƿ��е���'><label for='computer_1'>û��</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>�з���������</strong><INPUT type=radio value=�� name=internet id=internet_0 title='��ѡ���Ƿ��п��������'><label for='internet_0'>��</label>&nbsp;&nbsp;<INPUT name=internet id=internet_1 type=radio value=û�� checked title='��ѡ���Ƿ��п��������'><label for='internet_1'>û��</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>�ͷ�����(��ע)��</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><textarea name=remarkinfo cols=50 rows=3 id='remarkinfo' title='���������ҿͷ����ܣ��ǵ��˴�����˫�˴����з��¹񣬿յ���д��̨��'></textarea></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left></div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked><label for='passregtext'>ͬ��</label>��<input class='input_radio' type=radio name=passregtext id=nopassregtext value='0'>��ͬ�⡡&nbsp;<a href='RegisterHome.asp?action=protocol' target='_blank'> �� �鿴ע������</a></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf	
	
	if oblog.setup(57,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
    	show_reg=show_reg&"<td width=20% height=25><div align=left>��֤�룺</div></td>"& vbcrlf
   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
   		show_reg=show_reg&"</tr>"& vbcrlf
	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 ></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' �ύ����Ӵ���ͥ ' style='height:21px;font-size:13px;font-weight:bold'>��"& vbcrlf
	show_reg=show_reg&"<input type='button' name='historybackwl' value='������һ��' onclick='javascript:history.go(-1);' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'><input type=hidden name=action value='chkreg'><input type=hidden name=FamilyType value='"& Request.Form("FamilyType") &"'><input type=hidden name=FamilyTypeRequest value='"& FamilyTypeRequest &"'>"& vbcrlf
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
	Dim birthday,job,tel,address
	Dim mobile,postnum,familynumber,familymember,houseframe,housespace,pet,PetType,asksex,issaychinese,howknowsite,englishlevel,car,fitment,toilet,computer,internet,remarkinfo
	
	
	
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
'	'question=trim(request("question"))
'	'answer=trim(request("an"))
	question="����δ����������ʾ�����޷���¼�������Ա��ϵ����Ϊ���һ����룺��"
	answer="�Ҽ����ӱ"
	blogname=trim(request("blogname"))
	usertype=clng(request("usertype"))
	'nickname=trim(request("nickname"))
	user_domain=Lcase(trim(request("domain")))
	user_domainroot=trim(request("user_domainroot"))
	
	birthday=Trim(Request("y"))&"-"&Trim(Request("m"))&"-"&Trim(Request("d"))
	job= Trim(Request("job"))
	tel= Trim(Request("tel"))
	address= Trim(Request("address"))
	
	mobile= Trim(Request("mobile"))
	postnum= Trim(Request("postnum"))
	familynumber= Trim(Request("familynumber"))
	familymember= Trim(Request("familymember"))
	houseframe= Trim(Request("houseframe"))
	housespace= Trim(Request("housespace"))
	pet= Trim(Request("pet"))
	PetType= Trim(Request("PetType"))
	asksex= Trim(Request("asksex"))
	issaychinese= Trim(Request("issaychinese"))
	howknowsite= Trim(Request("howknowsite"))
	englishlevel= Trim(Request("englishlevel"))
	car= Trim(Request("car"))
	fitment= Trim(Request("fitment"))
	toilet= Trim(Request("toilet"))
	computer= Trim(Request("computer"))
	internet= Trim(Request("internet"))
	remarkinfo= Trim(Request("remarkinfo"))
	
	Dim FamilyTypeRequest
	FamilyTypeRequest = request("FamilyTypeRequest")
	
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("�û�������Ϊ�գ����Ҳ��ܴ���14С��4��")
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
	if question="" or oblog.strLength(question)>150 then oblog.adderrstr("�һ�������ʾ���ⲻ��Ϊ�գ�Ҳ���ܴ���150��")
	if answer="" or oblog.strLength(answer)>150 then oblog.adderrstr("�һ���������𰸲���Ϊ�գ�Ҳ���ܴ���150��")
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
	
	
	If IsDate(birthday)=False Then oblog.adderrstr("�������ո�ʽ��δ��ã�������ѡ��")	

	
	
	
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
		
		rsreg("sex")= sex
		rsreg("birthday")= birthday
		rsreg("job")= job
		rsreg("tel")= tel
		rsreg("address")= address
		
		rsreg("mobile")= mobile
		rsreg("postnum")= postnum
		rsreg("familynumber")= familynumber
		rsreg("familymember")= familymember
		rsreg("houseframe")= houseframe
		rsreg("housespace")= housespace
		rsreg("pet")= pet
		rsreg("PetType")= PetType
		rsreg("asksex")= asksex
		rsreg("issaychinese")= issaychinese
		rsreg("howknowsite")= howknowsite
		rsreg("englishlevel")= englishlevel
		rsreg("car")= car
		rsreg("fitment")= fitment
		rsreg("toilet")= toilet
		rsreg("computer")= computer
		rsreg("internet")= internet
		rsreg("remarkinfo")= remarkinfo
		
		rsreg("FamilyType")= Request("FamilyType")
			
		rsreg.update
		oblog.execute("update oblog_setup set user_count=user_count+1")
		if is_unamedir=0 then
			oblog.execute("update oblog_user set user_folder=userid where username='"&oblog.filt_badstr(regusername)&"'")
		end if
'wl20070816,����ע��ɹ�����Ϣ----------------------------------	Begin	
		dim sqlt,rst
		sqlt="select * from oblog_pm"
		set rst=server.createobject("adodb.recordset")
		rst.open sqlt,conn,1,3
		rst.addnew
		rst("incept")=oblog.Interceptstr(regusername,50)'wl---regusername�û���;
		rst("topic")=oblog.Interceptstr(oblog.filt_badword("��ϲ��ע��ɹ�...��ӭ�����ҵ�ס������"),100)
		rst("content")=oblog.setup(86,0)
		rst("sender")="�ҵ�ס����"
		rst.update
		rst.close
		set rst= nothing
'wl20070816,����ע��ɹ�����Ϣ----------------------------------	End		
		show_reg="��ǰλ�ã�<a href='index.asp'>��ҳ</a>�����ע��<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>��ϲ�����Ѿ��ɹ�ע��Ϊ </strong><font color=#038ad7>"& FamilyTypeRequest &"</font> ��</li></ul>"
			show_reg=show_reg&"10����Զ�ת��ϵͳ��ҳ��"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",10000);"
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
	 
//	 tishi = del_space(document.oblogform.question.value);
//     if (tishi.length == 0)
//     {
//        alert("������������ʾ����");
//        return false;
//     }
//	 
//	 tsda = del_space(document.oblogform.an.value);
//     if (tsda.length == 0)
//     {
//        alert("������������ʾ����𰸣�");
//        return false;
//     }
////	WL���ظ��룡 city = del_space(document.all("city").value);
////     if (city.length == 0)
////     {
////        alert("��ѡ�����ڳ���!");
////	return false;
////     }		
	 
	 
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
        alert("��ѡ������ѧϰ��Χ!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>


<script>
	function ShowPet(ID){
		if(ID=="1"){
		window.PetTypes.innerHTML="<select id=\"PetType\" name=\"PetType\"><option value=\"\" selected>��ѡ��<\/option><option value=\"��\" >��<\/option><option value=\"è\">è<\/option><option value=\"��\">��<\/option><option value=\"����\">����<\/option><\/select>";
		//document.oblogform.CheckPet.value="1"
		}
		if(ID=="0"){
		window.PetTypes.innerHTML="";
		//document.oblogform.CheckPet.value="0"
		}
		}
</script>