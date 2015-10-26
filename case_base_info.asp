<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->



<style type=text/css>
/* Begin：wl 添加目的:实现各种填写表单界面的样式 */
.input_text {
font-family:Tahoma, Arial, 宋体;
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
/* End：  wl */
</style>

<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
		&#8226; <a href="case_base_info.asp">站点设置</a>
		&#8226; <a href="case_base_info.asp?action=placard">修改公告</a>
		&#8226; <a href="case_base_info.asp?action=links">友情连接</a>
		&#8226; <a href="case_base_info.asp?action=blogpassword">整站加密</a>
		&#8226; <a href="case_base_info.asp?action=blogstar">用户明星</a>
		&#8226; <a href="case_base_info.asp?action=baseinfo">个人资料</a>
		&#8226; <a href="case_base_info.asp?action=userpassword">修改我的密码</a>
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			随诊档案-基本信息
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

'Begin: 排除无健康档案的用户，以前注册的无该记录，可用户新建来创建
If rs.bof and rs.EOF Then
	oblog.adderrstr("你需要新建自己的健康档案！")
	
	rs.Close
	set rs = nothing
End If
	
'现按退出来处理 ，后期直接进入创建新健康档案基本信息的流程
if oblog.errstr<>"" then oblog.showusererr:exit sub
'End :

%>
<style>
#list ul{ width: 90%;margin-bottom: 0;}
#list ul li.t1 { width: 180px;text-align: right;} 
#list ul li.t2 { width: 400px;text-align: left;} 
</style>

<%
'查询用户注册的医生是否设置了随诊病历
Dim rsEMR, strEMR

'查询添加病历、病历表单未锁定，根据用户注册的医生，再根据医生查找病历，在查找所以病历表单
strEMR = "Select f.id, f.classname, f.FilePath, f.ordernum, f.locked From [91health]..[oblog_user] u91 " & _
				"Left Join [54doctor]..[oblog_user] u54 On (u91.Reg_Site_ID=u54.userid) " & _
				
				"Left Join [EMR]..[EMR_DB] d On (u54.EMR_ID=d.id) " & _
				 
				"Right Join [EMR]..[EMR_Form] f On (f.EMR_ID = d.id ) " & _
				
				"Where d.locked=0 and f.locked=0 and u91.userid =  " & oblog.logined_uid & _
					"Order By f.ordernum"

set rsEMR=oblog.execute(strEMR)

%>

<h1><a href="case_base_info.asp" id="a_active">基本信息</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

<%
'循环列出随诊病历的链接
Do While Not rsEMR.EOF
%>
	<a href=<%=rsEMR("FilePath")%>><% =rsEMR("classname") %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<!--
	<a href="BingShi.asp" id="a_active">患者病史</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="JianCha_TiGe.asp">各项体检情况</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="JianCha_FuZhu.asp">实验室检查</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="ZhuanKe_YueJing.asp">月经情况</a>
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
	<li class="t1">真实姓名：</li>
	<li class="t2"><input name="realname" type="text" class="input_text" value="<%=oblog.filt_html(rs("real_name"))%>" size="20" maxlength="20" />  <font color=#ff0000> *</font>    </li>
    </ul>
    <ul>
	<li class="t1">性别：</li>
	<li class="t2"><input type="radio" value="1" name="sex" <%if rs("Sex")=1 then response.write "checked"%> />
        男 &nbsp;&nbsp; <input type="radio" value="0" name="sex" <%if rs("Sex")=0 then response.write "checked"%> />女 <font color=#ff0000> *</font></li>
    </ul>

	<ul>
	<li class="t1">出生日期：</li>
	<li class="t2">
		<input value="<%=year(rs("birthday"))%>" class="input_text" name="y" size="4" maxlength="4" />年
		<input value="<%=month(rs("birthday"))%>" class="input_text" name="m" size="2" maxlength="2" />月
		<input value="<%=day(rs("birthday"))%>" class="input_text" name="d" size="2" maxlength="2" />日 <font color=#ff0000> *</font>
	  </li>
    </ul>
	
	<ul>
	<li class="t1">民族：</li>
	<li class="t2"><input name="nation"   type="text" class="input_text" value="<%=oblog.filt_html(rs("nation"))%>" size="20" maxlength="20" /> <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">婚姻：</li>
	<li class="t2">
		<input type="radio" value="0" name="marriage" <%if rs("marriage")=0 then response.write "checked"%> />未 &nbsp;&nbsp; 
		<input type="radio" value="1" name="marriage" <%if rs("marriage")=1 then response.write "checked"%> />已 &nbsp;&nbsp;
		<input type="radio" value="2" name="marriage" <%if rs("marriage")=2 then response.write "checked"%> />离 &nbsp;&nbsp; 
		<input type="radio" value="3" name="marriage" <%if rs("marriage")=3 then response.write "checked"%> />丧 &nbsp;&nbsp;  <font color=#ff0000> *</font></li>
    </ul>
	
	 <ul>
	<li class="t1">职业：</li>
	<li class="t2"><input name="vocation"   type="text" class="input_text" value="<%=oblog.filt_html(rs("vocation"))%>" size="30" maxlength="30" /> <font color=#ff0000> *</font> </li>
    </ul>
	
	<ul>
	<li class="t1">出生地：</li>
	<li class="t2"> <%=oblog.type_city(rs("birth_province"),rs("birth_city"))%>  <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">工作单位及地址：</li>
	<li class="t2"><input name="AddressOffice" value="<%=oblog.filt_html(rs("Address_Office"))%>" class="input_text" size="40" maxlength="250" /> <font color=#ff0000> *</font></li>
    </ul>
	
	<ul>
	<li class="t1">户口地址/现住址：</li>
	<li class="t2"><input name="AddressHome" class="input_text" value="<%=oblog.filt_html(rs("Address_Home"))%>" size="40" maxlength="250" /> <font color=#ff0000> *</font></li>
    </ul>
	
    <ul>
	<li class="t1">Email：</li>
	<li class="t2"><input name="Email" class="input_text" value="<%=oblog.filt_html(rs("Email"))%>" size="30" maxlength="50" /> <font color=#ff0000> *</font> 
      </li>
    </ul>

    <ul>
	<li class="t1">MSN：</li>
	<li class="t2"><input name="msn" class="input_text" value="<%=oblog.filt_html(rs("Msn"))%>" size="30" maxlength="50" /></li>
    </ul>
	
	<ul>
	<li class="t1">固定电话：</td>
	<li class="t2"><input name="Telephone" class="input_text" value="<%=oblog.filt_html(rs("Telephone"))%>" size="30" maxlength="50" /> <font color=#ff0000> </font></li>
    </ul>
	
	<ul>
	<li class="t1">移动电话：</td>
	<li class="t2"><input name="Mobile" class="input_text" value="<%=oblog.filt_html(rs("Mobile"))%>" size="30" maxlength="50" /> <font color=#ff0000> </font></li>
    </ul>
	
	<ul>
	<li class="t1">备注：</td>
	<li class="t2"><input name="Remark" class="input_text" value="<%=oblog.filt_html(rs("Remark"))%>" size="30" maxlength="250" /></li>
    </ul>

	  <ul>
	  <li class="t1"></li>
	  <li class="t2"><input name="action" type="hidden" value="savebaseinfo" /> 
        <input type=submit value=" 更新 " /></li>
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
	
	'运行注册用户只填写年份，不必填写月份、日份
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
		if not isdate(birthday) then oblog.adderrstr("生日日期格式错误！")
		if clng(trim(request("y")))>2060 then oblog.adderrstr("生日年份过大！")
		if clng(trim(request("y")))<1900 then oblog.adderrstr("生日年份过小！")
	end if
	
	'错误检查
	if trim(request("realname"))="" then oblog.adderrstr("您的真实姓名不能为空！")
	if birthday="" then oblog.adderrstr("您的生日日期不能为空！")
	if request("province")="" Or request("city")="" then oblog.adderrstr("您的出生地不能为空！")
	if trim(request("vocation"))="" then oblog.adderrstr("您的职业不能为空！")
	if trim(request("AddressOffice"))="" then oblog.adderrstr("您的工作单位及地址不能为空！")
	if trim(request("AddressHome"))="" then oblog.adderrstr("您的户口地址/现住址不能为空！")
	'if trim(request("telephone"))="" Or trim(request("Mobile"))="" then oblog.adderrstr("您的联系电话不能为空！")
	
	if email="" then oblog.adderrstr("电子邮件地址不能为空！")
	if not oblog.IsValidEmail(email) then oblog.adderrstr("电子邮件地址格式错误！")
	
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
			rs("birthday")=birthday		'保存出生日期
			
			'后续改进：对婴儿要特殊处理，精确到月份和天数
			rs("age") = CStr(DateDiff("yyyy", birthday, Now)) & "岁"
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
	oblog.showok "保存资料成功!",""
end sub

%>