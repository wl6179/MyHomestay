<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->



<style type=text/css>
/* Begin��wl ���Ŀ��:ʵ�ָ�����д���������ʽ */
.input_text {
font-family:Tahoma, Arial, ����;
font-size: 8pt;
line-height:15px;
COLOR: #000000;
border-top-style: none;
border-right-style: none;

border-left-style: none;
border-bottom-color: #000000;
border-top-width: 0px;
border-right-width: 0px;
border-bottom-width: 1px;
border-left-width: 0px;

background:#BFE0FB;

}

#a_active {
color:#2f2f2f;
}

.img_style {
height:13px;
width:15px;
cursor:hand;
}
/* End��  wl */
</style>

<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
		&#8226; <a href="case_base_info.asp">վ������</a>
		&#8226; <a href="case_base_info.asp?action=placard">�޸Ĺ���</a>
		&#8226; <a href="case_base_info.asp?action=links">��������</a>
		&#8226; <a href="case_base_info.asp?action=blogpassword">��վ����</a>
		&#8226; <a href="case_base_info.asp?action=blogstar">�û�����</a>
		&#8226; <a href="case_base_info.asp?action=baseinfo">��������</a>
		&#8226; <a href="case_base_info.asp?action=userpassword">�޸��ҵ�����</a>
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			���ﵵ��-������Ϣ
	</div>
	
    <div class="content_body">
<%
dim action
action=request("action")
select case action
	case "savebaseinfo"
	call savebaseinfo
	case "baseinfo"
	call baseinfo
	case else
	'call sitesetup	
	call baseinfo
end select

%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>  
  
  <div id="bottom"><%=oblog.user_copyright%></div>

</body>
</html>
<%



sub baseinfo
dim rs
set rs=oblog.execute("select * from case_base_info where user_id="&oblog.logined_uid)

'Begin: �ų��޽����������û�����ǰע����޸ü�¼�����û��½�������
If rs.bof and rs.EOF Then
	oblog.adderrstr("����Ҫ�½��Լ��Ľ���������")
	
	rs.Close
	set rs = nothing
End If
	
'�ְ��˳������� ������ֱ�ӽ��봴���½�������������Ϣ������
if oblog.errstr<>"" then oblog.showusererr:exit sub
'End :

%>
<style>
#list ul{ width: 90%;margin-bottom: 0;}
#list ul li.t1 { width: 180px;text-align: right;} 
#list ul li.t2 { width: 400px;text-align: left;} 
</style>

<%
'��ѯ�û�ע���ҽ���Ƿ����������ﲡ��
Dim rsEMR, strEMR

'��ѯ��Ӳ�����������δ�����������û�ע���ҽ�����ٸ���ҽ�����Ҳ������ڲ������Բ�����
strEMR = "Select f.id, f.classname, f.FilePath, f.ordernum, f.locked From [91health]..[oblog_user] u91 " & _
				"Left Join [54doctor]..[oblog_user] u54 On (u91.Reg_Site_ID=u54.userid) " & _
				
				"Left Join [EMR]..[EMR_DB] d On (u54.EMR_ID=d.id) " & _
				 
				"Right Join [EMR]..[EMR_Form] f On (f.EMR_ID = d.id ) " & _
				
				"Where d.locked=0 and f.locked=0 and u91.userid =  " & oblog.logined_uid & _
					"Order By f.ordernum"

set rsEMR=oblog.execute(strEMR)

%>

<h1><a href="case_base_info.asp" id="a_active">������Ϣ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

<%
'ѭ���г����ﲡ��������
Do While Not rsEMR.EOF
%>
	<a href=<%=rsEMR("FilePath")%>><% =rsEMR("classname") %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<!--
	<a href="BingShi.asp" id="a_active">���߲�ʷ</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="JianCha_TiGe.asp">����������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="JianCha_FuZhu.asp">ʵ���Ҽ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="ZhuanKe_YueJing.asp">�¾����</a>
-->
<%
	rsEMR.Movenext
Loop


	rsEMR.Close
	set rsEMR = Nothing
%>


</h1>
<div id="list">
<form name="oblogform" action="case_base_info.asp" method="post">
    <ul>
	<li class="t1">��ʵ������</li>
	<li class="t2"><input name="realname" type="text" class="input_text" value="<%=oblog.filt_html(rs("real_name"))%>" size="20" maxlength="20" />  <font color=#ff0000> *</font>    </li>
    </ul>
    <ul>
	<li class="t1">�Ա�</li>
	<li class="t2"><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then response.write "checked"%> />
        �� &nbsp;&nbsp; <input type="radio" value="0" name="sex" <%if rs("Sex")=0 then response.write "checked"%> />Ů <font color=#ff0000> *</font></li>
    </ul>

	<ul>
	<li class="t1">�������ڣ�</li>
	<li class="t2">
		<input value="<%=year(rs("birthday"))%>" class="input_text" name="y" size="4" maxlength="4" />��
		<input value="<%=month(rs("birthday"))%>" class="input_text" name="m" size="2" maxlength="2" />��
		<input value="<%=day(rs("birthday"))%>" class="input_text" name="d" size="2" maxlength="2" />�� <font color=#ff0000> *</font>
	  </li>
    </ul>
	
	<ul>
	<li class="t1">���壺</li>
	<li class="t2"><input name="nation"   type="text" class="input_text" value="<%=oblog.filt_html(rs("nation"))%>" size="20" maxlength="20" /> <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">������</li>
	<li class="t2">
		<input type="radio" value="0" name="marriage" <%if rs("marriage")=0 then response.write "checked"%> />δ &nbsp;&nbsp; 
		<input type="radio" value="1" name="marriage" <%if rs("marriage")=1 then response.write "checked"%> />�� &nbsp;&nbsp;
		<input type="radio" value="2" name="marriage" <%if rs("marriage")=2 then response.write "checked"%> />�� &nbsp;&nbsp; 
		<input type="radio" value="3" name="marriage" <%if rs("marriage")=3 then response.write "checked"%> />ɥ &nbsp;&nbsp;  <font color=#ff0000> *</font></li>
    </ul>
	
	 <ul>
	<li class="t1">ְҵ��</li>
	<li class="t2"><input name="vocation"   type="text" class="input_text" value="<%=oblog.filt_html(rs("vocation"))%>" size="30" maxlength="30" /> <font color=#ff0000> *</font> </li>
    </ul>
	
	<ul>
	<li class="t1">�����أ�</li>
	<li class="t2"> <%=oblog.type_city(rs("birth_province"),rs("birth_city"))%>  <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">������λ����ַ��</li>
	<li class="t2"><input name="AddressOffice" value="<%=oblog.filt_html(rs("Address_Office"))%>" class="input_text" size="40" maxlength="250" /> <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">���ڵ�ַ/��סַ��</li>
	<li class="t2"><input name="AddressHome" class="input_text" value="<%=oblog.filt_html(rs("Address_Home"))%>" size="40" maxlength="250" /> <font color=#ff0000> *</font></li>
    </ul>
	
    <ul>
	<li class="t1">Email��</li>
	<li class="t2"><input name="Email" class="input_text" value="<%=oblog.filt_html(rs("Email"))%>" size="30" maxlength="50" /> <font color=#ff0000> *</font> 
      </li>
    </ul>

    <ul>
	<li class="t1">MSN��</li>
	<li class="t2"><input name="msn" class="input_text" value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></li>
    </ul>
	
	<ul>
	<li class="t1">�̶��绰��</td>
	<li class="t2"><input name="Telephone" class="input_text" value="<%=oblog.filt_html(rs("Telephone"))%>" size="30" maxlength="50" /> <font color=#ff0000> </font></li>
    </ul>
	
	<ul>
	<li class="t1">�ƶ��绰��</td>
	<li class="t2"><input name="Mobile" class="input_text" value="<%=oblog.filt_html(rs("Mobile"))%>" size="30" maxlength="50" /> <font color=#ff0000> </font></li>
    </ul>
	
	<ul>
	<li class="t1">��ע��</td>
	<li class="t2"><input name="Remark" class="input_text" value="<%=oblog.filt_html(rs("Remark"))%>" size="30" maxlength="250" /></li>
    </ul>

	  <ul>
	  <li class="t1"></li>
	  <li class="t2"><input name="action" type="hidden" value="savebaseinfo" /> 
        <input type=submit value=" ���� " /></li>
		</ul>
  </form>
</div>
<%
set rs=nothing
end sub

sub savebaseinfo
	dim rs,email,birthday
	Dim realname, sex, birthprovince, birthcity, nation, Marriage, vocation, msn, telephone, Mobile, addressOffice, addressHome, Remark, Checked
	
	email=trim(request("email"))
	
	'����ע���û�ֻ��д��ݣ�������д�·ݡ��շ�
	Dim m,d
	m = trim(request("m"))
	d = trim(request("d"))
	
	If m = "" Then m = "1"
	If d = "" Then d = "1"
	
	birthday=trim(request("y"))&"-"& m &"-"& d
	
	'birthday=trim(request("y"))&"-"&trim(request("m"))&"-"&trim(request("d"))
	

	if birthday = "--" then
		birthday=""
	else
		if not isdate(birthday) then oblog.adderrstr("�������ڸ�ʽ����")
		if clng(trim(request("y")))>2060 then oblog.adderrstr("������ݹ���")
		if clng(trim(request("y")))<1900 then oblog.adderrstr("������ݹ�С��")
	end if
	
	'������
	if trim(request("realname"))="" then oblog.adderrstr("������ʵ��������Ϊ�գ�")
	if birthday="" then oblog.adderrstr("�����������ڲ���Ϊ�գ�")
	if request("province")="" Or request("city")="" then oblog.adderrstr("���ĳ����ز���Ϊ�գ�")
	if trim(request("vocation"))="" then oblog.adderrstr("����ְҵ����Ϊ�գ�")
	if trim(request("AddressOffice"))="" then oblog.adderrstr("���Ĺ�����λ����ַ����Ϊ�գ�")
	if trim(request("AddressHome"))="" then oblog.adderrstr("���Ļ��ڵ�ַ/��סַ����Ϊ�գ�")
	'if trim(request("telephone"))="" Or trim(request("Mobile"))="" then oblog.adderrstr("������ϵ�绰����Ϊ�գ�")
	
	if email="" then oblog.adderrstr("�����ʼ���ַ����Ϊ�գ�")
	if not oblog.IsValidEmail(email) then oblog.adderrstr("�����ʼ���ַ��ʽ����")
	
	if oblog.errstr<>"" then oblog.showusererr:exit sub
	
	realname = oblog.filt_astr(request("realname"),20)
	sex = clng(request("sex"))
	birthprovince = request("province")
	birthcity = request("city")
	nation = oblog.filt_astr(request("nation"),50)
	Marriage = clng(request("Marriage"))
	vocation = oblog.filt_astr(request("vocation"),250)
	email = oblog.filt_astr(email,50)
	msn = oblog.filt_astr(request("msn"),50)
	telephone = oblog.filt_astr(request("telephone"),50)
	Mobile = oblog.filt_astr(request("Mobile"),50)
	addressOffice = oblog.filt_astr(request("addressOffice"),250)
	addressHome = oblog.filt_astr(request("addressHome"),250)
	Remark = oblog.filt_astr(request("Remark"),250)
	Checked = 0
	
	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from case_base_info where user_id="&oblog.logined_uid,conn,1,3
	if not rs.eof then
		rs("real_name")= realname
		rs("sex")= sex
		rs("birth_province")= birthprovince
		rs("birth_city")= birthcity
		if birthday<>"" then 
			rs("birthday")=birthday		'�����������
			
			'�����Ľ�����Ӥ��Ҫ���⴦����ȷ���·ݺ�����
			rs("age") = CStr(DateDiff("yyyy", birthday, Now)) & "��"
		End If
		
		rs("nation")= nation
		rs("Marriage")= Marriage
		rs("vocation")= vocation
		
		rs("email")= email
		rs("msn")= msn
		rs("telephone")= telephone
		rs("Mobile")= Mobile
		
		rs("address_office")= addressOffice
		rs("address_home")= addressHome
		
		rs("Remark")= Remark
		
		rs("Checked")=Checked
		
		rs.update
		rs.close
	end if
	set rs=nothing
	oblog.showok "�������ϳɹ�!",""
end sub

%>