<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
Select Case Request("action")
	Case "about_us"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>����������<hr />"&oblog.setup(87,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case "contact_us"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>����ϵ����<hr />"&oblog.setup(88,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case "cooperate_info"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>����������<hr />"&oblog.setup(89,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case "help"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>����������<hr />"&oblog.setup(90,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case "especially"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>���ر���ʾ<hr />"&oblog.setup(91,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case "recommend"
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>���Ƽ��������<hr />"&oblog.setup(92,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
	
	Case else 
	call sysshow()  '���Ȼ��ģ��~...���滻$show_list$Ϊע��������Ϣ��
	show=replace(show,"$show_list$","��ǰλ�ã�<a href='index.asp'>��ҳ</a>����ӭ�����ҵ�ס����<hr />"&oblog.setup(86,0))
	response.Write show&oblog.site_bottom
	Response.End '���ҽ������������
End Select

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

	


end sub

sub sub_showreg()
	
	
end sub

sub sub_chkreg()


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